// Sentry must be first — before any other import
require("./instrument.js");
const Sentry = require("@sentry/node");

const express = require("express");
const session = require("express-session");
const msal = require("@azure/msal-node");
const path = require("path");
const https = require("https");
const fs = require("fs");
const { Pool } = require("pg");
require("dotenv").config();
const { sendChatMessage } = require("./lib/chat");
const chatUsage = require("./lib/chat-usage");
const { buildCache, tryCache } = require("./lib/chat-cache");

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

// ── Azure AD SSO config ──
const REDIRECT_URI = process.env.REDIRECT_URI || "http://localhost:3000/auth/callback";
const LOCAL_DEV = !process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET;
let msalClient = null;
if (!LOCAL_DEV) {
  const msalConfig = {
    auth: {
      clientId: process.env.AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
  };
  msalClient = new msal.ConfidentialClientApplication(msalConfig);
} else {
  console.log("[Auth] Azure AD creds missing — running in local dev mode (no SSO)");
}

// Trust Railway's reverse proxy so secure cookies work behind HTTPS
app.set("trust proxy", 1);

// Session
if (!process.env.SESSION_SECRET && !LOCAL_DEV) {
  console.error("[FATAL] SESSION_SECRET environment variable is required");
  process.exit(1);
}
app.use(session({
  secret: process.env.SESSION_SECRET || "local-dev-secret",
  resave: false,
  saveUninitialized: false,
  cookie: { secure: process.env.NODE_ENV === "production", maxAge: 8 * 60 * 60 * 1000 },
}));

// ── Azure AD SSO routes ──
app.get("/auth/login", async (req, res) => {
  try {
    const authUrl = await msalClient.getAuthCodeUrl({
      scopes: ["user.read"],
      redirectUri: REDIRECT_URI,
    });
    res.redirect(authUrl);
  } catch (err) {
    console.error("[Auth] Login redirect error:", err.message);
    res.status(500).send("Authentication unavailable — " + err.message);
  }
});

app.get("/auth/callback", async (req, res) => {
  if (req.session && req.session.user) return res.redirect("/");
  if (!req.query.code) return res.redirect("/auth/login");
  try {
    const tokenResponse = await msalClient.acquireTokenByCode({
      code: req.query.code,
      scopes: ["user.read"],
      redirectUri: REDIRECT_URI,
    });
    req.session.user = {
      name: tokenResponse.account.name,
      email: tokenResponse.account.username,
    };
    Sentry.setUser({ email: req.session.user.email, username: req.session.user.name });
    req.session.save(() => res.redirect("/"));
  } catch (err) {
    console.error("[Auth] Callback error:", err.message);
    Sentry.withScope((scope) => { scope.setTag("where", "auth.callback"); Sentry.captureException(err); });
    // Show the error instead of redirecting to prevent infinite loops
    res.status(500).send("Authentication failed — " + err.message);
  }
});

// Debug route to check config (remove after SSO is working)
app.get("/auth/debug", (req, res) => {
  res.json({
    hasClientId: !!process.env.AZURE_CLIENT_ID,
    hasTenantId: !!process.env.AZURE_TENANT_ID,
    hasClientSecret: !!process.env.AZURE_CLIENT_SECRET,
    redirectUri: REDIRECT_URI,
    nodeEnv: process.env.NODE_ENV,
    trustProxy: app.get("trust proxy"),
    sessionUser: req.session ? req.session.user : null,
    protocol: req.protocol,
    secure: req.secure,
  });
});

app.get("/auth/logout", (req, res) => {
  req.session.destroy();
  const postLogoutUri = REDIRECT_URI.replace("/auth/callback", "/auth/signed-out");
  res.redirect(`https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(postLogoutUri)}`);
});

// Login page and signed-out page (served before auth guard)
app.get("/auth/login-page", (req, res) => {
  if (req.session && req.session.user) return res.redirect("/");
  res.sendFile(path.join(__dirname, "public", "login.html"));
});
app.get("/auth/signed-out", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "signed-out.html"));
});

// Serve favicon before auth guard
app.get("/favicon.svg", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "favicon.svg"));
});
app.get("/favicon.ico", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "favicon.svg"), {
    headers: { "Content-Type": "image/svg+xml" },
  });
});

// Auth guard — block everything except /auth/* and /api/health
app.use((req, res, next) => {
  if (LOCAL_DEV) { req.session.user = req.session.user || { name: "Local Dev", email: "dev@localhost" }; return next(); }
  if (req.path.startsWith("/auth/")) return next();
  if (req.path === "/api/health") return next();
  if (req.session && req.session.user) {
    Sentry.setUser({ email: req.session.user.email, username: req.session.user.name });
    return next();
  }
  res.redirect("/auth/login-page");
});

// ── Postgres Cache (Railway provides DATABASE_URL) ──
const pool = process.env.DATABASE_URL
  ? new Pool({ connectionString: process.env.DATABASE_URL, ssl: process.env.DATABASE_URL.includes("localhost") ? false : { rejectUnauthorized: false } })
  : null;

async function initDB() {
  if (!pool) { console.log("[DB] No DATABASE_URL — running without Postgres cache"); return; }
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS api_cache (
      report_key TEXT PRIMARY KEY,
      data JSONB NOT NULL,
      fetched_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS wo_writes (
      id SERIAL PRIMARY KEY,
      wo_ref TEXT NOT NULL,
      wo_id TEXT,
      field_name TEXT NOT NULL,
      old_value TEXT,
      new_value TEXT,
      user_email TEXT,
      mode TEXT NOT NULL,
      status TEXT NOT NULL,
      error_message TEXT,
      attempted_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_wo_writes_ref ON wo_writes(wo_ref)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_wo_writes_attempted_at ON wo_writes(attempted_at DESC)`);
    console.log("[DB] Cache + audit tables ready");
  } catch (e) { console.error("[DB] Init error:", e.message); }
}

async function loadFromDB() {
  if (!pool) return;
  try {
    const woRes = await pool.query("SELECT data, fetched_at FROM api_cache WHERE report_key = 'workorders'");
    if (woRes.rows.length) {
      cachedWO = woRes.rows[0].data;
      cacheWOTimestamp = new Date(woRes.rows[0].fetched_at).getTime();
      console.log("[DB] Loaded " + cachedWO.length + " WOs from Postgres (cached " + Math.round((Date.now() - cacheWOTimestamp) / 60000) + "m ago)");
    }
    const reqRes = await pool.query("SELECT data, fetched_at FROM api_cache WHERE report_key = 'requests'");
    if (reqRes.rows.length) {
      cachedReq = reqRes.rows[0].data;
      cacheReqTimestamp = new Date(reqRes.rows[0].fetched_at).getTime();
      console.log("[DB] Loaded " + cachedReq.length + " Requests from Postgres (cached " + Math.round((Date.now() - cacheReqTimestamp) / 60000) + "m ago)");
    }
  } catch (e) { console.error("[DB] Load error:", e.message); }
}

async function saveToDB(key, data) {
  if (!pool) return;
  try {
    await pool.query(
      `INSERT INTO api_cache (report_key, data, fetched_at) VALUES ($1, $2, NOW())
       ON CONFLICT (report_key) DO UPDATE SET data = $2, fetched_at = NOW()`,
      [key, JSON.stringify(data)]
    );
    console.log("[DB] Saved " + data.length + " records to Postgres (" + key + ")");
  } catch (e) { console.error("[DB] Save error:", e.message); }
}

// ── Axxerion API config ──
const AX_URL = "https://ipg.axxerion.us/webservices/ipg/rest/functions/completereportresult";
const AX_AUTH = "Basic " + Buffer.from((process.env.AX_USER || "iapiuser") + ":" + (process.env.AX_PASS || "")).toString("base64");
const AX_REPORT_WO = "IPG-REP-085";
const AX_REPORT_WO_INCR = "IPG-REP-101"; // incremental: updateTime >= -2 days
const AX_REPORT_REQ = "IPG-REP-110";
const REFRESH_INTERVAL_WO_FULL = 4 * 60 * 60 * 1000; // 4 hours for full WO pull (11K+ records)
const REFRESH_INTERVAL_WO_INCR = 10 * 60 * 1000; // 10 minutes for incremental WO pull
const REFRESH_INTERVAL_REQ = 10 * 60 * 1000; // 10 minutes for requests (small dataset)
const RETRY_DELAY = 60 * 1000; // 1 minute retry when a fetch was skipped

let cachedWO = null;
let cacheWOTimestamp = 0;
let fetchWOInProgress = false;
let fetchWOIncrInProgress = false;
let lastFetchStartTime = 0;
let lastIncrUpdate = 0; // track last incremental merge time

let cachedReq = null;
let cacheReqTimestamp = 0;
let fetchReqInProgress = false;

let authFailed = false; // stop all fetches if API returns unauthorized
let refreshTimer = null;

function fetchReport(reference, label, onSuccess) {
  const start = Date.now();
  console.log(`[Cache] Fetching ${label} (${reference})...`);

  const body = JSON.stringify({ reference });
  const opts = {
    hostname: "ipg.axxerion.us",
    path: "/webservices/ipg/rest/functions/completereportresult",
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: AX_AUTH,
      "Content-Length": Buffer.byteLength(body),
    },
  };

  const req = https.request(opts, (res) => {
    let raw = "";
    res.on("data", (chunk) => (raw += chunk));
    res.on("end", () => {
      const elapsed = ((Date.now() - start) / 1000).toFixed(1);

      // ── Unauthorized: stop all future fetches immediately ──
      if (res.statusCode === 401 || raw.trim().toLowerCase() === "unauthorized") {
        console.error(`[Cache] ⛔ ${label}: UNAUTHORIZED (HTTP ${res.statusCode}) — halting all API fetches. Update credentials and restart.`);
        Sentry.withScope((scope) => {
          scope.setTag("where", "axxerion.fetchReport");
          scope.setTag("severity", "critical");
          scope.setExtra("label", label);
          scope.setExtra("reference", reference);
          scope.setExtra("statusCode", res.statusCode);
          Sentry.captureMessage(`Axxerion API unauthorized — credentials rejected (${label})`, "error");
        });
        authFailed = true;
        if (refreshTimer) { clearInterval(refreshTimer); refreshTimer = null; }
        onSuccess(null);
        return;
      }

      try {
        const json = JSON.parse(raw);
        if (json.data && json.data.length) {
          console.log(
            `[Cache] ${label}: ${json.data.length} records (${(raw.length / 1024 / 1024).toFixed(1)}MB) in ${elapsed}s`
          );
          onSuccess(json.data);
        } else {
          console.warn(`[Cache] ${label} empty in ${elapsed}s:`, json.errorMessage || "no data");
          onSuccess(null);
        }
      } catch (e) {
        console.error(`[Cache] ${label} parse error in ${elapsed}s:`, e.message);
        Sentry.withScope((scope) => {
          scope.setTag("where", "axxerion.fetchReport.parse");
          scope.setExtra("label", label);
          scope.setExtra("reference", reference);
          Sentry.captureException(e);
        });
        onSuccess(null);
      }
    });
  });

  req.setTimeout(20 * 60 * 1000, () => {
    console.error(`[Cache] ${label} timed out after 20 minutes`);
    req.destroy();
    onSuccess(null);
  });

  req.on("error", (e) => {
    console.error(`[Cache] ${label} fetch failed:`, e.message);
    onSuccess(null);
  });

  req.write(body);
  req.end();
}

function fetchWorkOrders() {
  if (authFailed) { console.log("[Cache] Skipping WO fetch — credentials invalid."); return; }
  if (fetchWOInProgress) {
    console.log("[Cache] Skipping Work Orders — previous fetch still in progress");
    setTimeout(fetchWorkOrders, RETRY_DELAY);
    return;
  }
  fetchWOInProgress = true;
  lastFetchStartTime = Date.now();
  fetchReport(AX_REPORT_WO, "Work Orders", (data) => {
    fetchWOInProgress = false;
    if (data) { cachedWO = data; cacheWOTimestamp = Date.now(); saveToDB("workorders", data); buildCache(() => ({ workOrders: cachedWO || [], requests: cachedReq || [] })); }
  });
}

function fetchWorkOrdersIncremental() {
  if (authFailed) { console.log("[Cache] Skipping incremental WO fetch — credentials invalid."); return; }
  if (!cachedWO) { console.log("[Cache] Skipping incremental — no full dataset yet, waiting for full fetch"); return; }
  if (fetchWOIncrInProgress) {
    console.log("[Cache] Skipping incremental WOs — previous fetch still in progress");
    setTimeout(fetchWorkOrdersIncremental, RETRY_DELAY);
    return;
  }
  fetchWOIncrInProgress = true;
  fetchReport(AX_REPORT_WO_INCR, "Work Orders (incremental)", (data) => {
    fetchWOIncrInProgress = false;
    if (data && data.length) {
      // Build a map of existing WOs by ID for fast lookup
      const woMap = new Map();
      cachedWO.forEach((wo, idx) => woMap.set(wo.ID, idx));

      let updated = 0, added = 0;
      data.forEach((wo) => {
        const existingIdx = woMap.get(wo.ID);
        if (existingIdx !== undefined) {
          cachedWO[existingIdx] = wo; // replace with updated record
          updated++;
        } else {
          cachedWO.push(wo); // new WO not in full dataset yet
          woMap.set(wo.ID, cachedWO.length - 1);
          added++;
        }
      });

      lastIncrUpdate = Date.now();
      cacheWOTimestamp = Date.now(); // mark data as fresh
      console.log(`[Cache] Incremental merge: ${data.length} records — ${updated} updated, ${added} new (total: ${cachedWO.length})`);
      saveToDB("workorders", cachedWO);
      buildCache(() => ({ workOrders: cachedWO || [], requests: cachedReq || [] }));
    }
  });
}

function fetchRequests() {
  if (authFailed) { console.log("[Cache] Skipping Request fetch — credentials invalid."); return; }
  if (fetchReqInProgress) {
    console.log("[Cache] Skipping Requests — previous fetch still in progress");
    setTimeout(fetchRequests, RETRY_DELAY);
    return;
  }
  fetchReqInProgress = true;
  fetchReport(AX_REPORT_REQ, "Requests", (data) => {
    fetchReqInProgress = false;
    if (data) { cachedReq = data; cacheReqTimestamp = Date.now(); saveToDB("requests", data); buildCache(() => ({ workOrders: cachedWO || [], requests: cachedReq || [] })); }
  });
}

// ── API endpoints for browser ──
app.get("/api/workorders", (req, res) => {
  if (cachedWO) {
    const ageMin = Math.round((Date.now() - cacheWOTimestamp) / 60000);
    res.json({ data: cachedWO, cached: true, age: ageMin, count: cachedWO.length, refreshing: fetchWOInProgress });
  } else {
    res.json({ data: [], cached: false, message: "Cache loading, try again shortly", refreshing: fetchWOInProgress });
  }
});

app.get("/api/requests", (req, res) => {
  if (cachedReq) {
    const ageMin = Math.round((Date.now() - cacheReqTimestamp) / 60000);
    res.json({ data: cachedReq, cached: true, age: ageMin, count: cachedReq.length, refreshing: fetchReqInProgress });
  } else {
    res.json({ data: [], cached: false, message: "Cache loading, try again shortly", refreshing: fetchReqInProgress });
  }
});

app.get("/api/status", (req, res) => {
  const nextWOFullMs = lastFetchStartTime ? (lastFetchStartTime + REFRESH_INTERVAL_WO_FULL) - Date.now() : 0;
  const nextWOFullMin = Math.max(0, Math.round(nextWOFullMs / 60000));
  res.json({
    workorders: { cached: !!cachedWO, count: cachedWO ? cachedWO.length : 0, ageMinutes: cachedWO ? Math.round((Date.now() - cacheWOTimestamp) / 60000) : null, fetchInProgress: fetchWOInProgress || fetchWOIncrInProgress, refreshIntervalHrs: 4, incrementalIntervalMin: 10, lastIncrUpdateMin: lastIncrUpdate ? Math.round((Date.now() - lastIncrUpdate) / 60000) : null },
    requests: { cached: !!cachedReq, count: cachedReq ? cachedReq.length : 0, ageMinutes: cachedReq ? Math.round((Date.now() - cacheReqTimestamp) / 60000) : null, fetchInProgress: fetchReqInProgress, refreshIntervalMin: 10 },
    nextWORefreshMin: nextWOFullMin,
    authFailed: authFailed,
    dbConnected: !!pool,
  });
});

// ── Manual refresh trigger ──
app.post("/api/refresh", (req, res) => {
  const full = req.body && req.body.full;
  if (full) {
    fetchWorkOrders(); // full 11K+ pull
  } else {
    fetchWorkOrdersIncremental(); // quick incremental
  }
  fetchRequests();
  res.json({ ok: true, message: full ? "Full refresh triggered" : "Incremental refresh triggered" });
});

// ── Axxerion Write Proxy ──
// Feature-flagged. When AXXERION_WRITE_ENABLED !== 'true', runs in dry-run mode:
// audits the attempt + updates in-memory cache, but never calls Axxerion.
// Endpoint pattern (from API Manual V2.1):
//   PUT /webservices/ipg/rest/functions/update/WorkOrder/{id}
//   Body: { "<internalFieldCode>": "<value>", ... }
// Date format per manual: dd-MM-yy HH:mm. Field codes are case sensitive.
// Workflow status transitions go through executefunction/{objectName}/{id}/{functionName} —
// requires explicit workflow perms (separate ask to San).
const AX_WRITE_ENABLED = process.env.AXXERION_WRITE_ENABLED === "true";
const AX_WRITE_URL = process.env.AXXERION_WRITE_URL ||
  "https://ipg.axxerion.us/webservices/ipg/rest/functions/update/WorkOrder";

// Display label → internal field code (from IPG-REP-085 schema CSV pulled 2026-04-29).
// Adding to this map alone is NOT enough — also add to AX_WRITABLE_FIELDS below.
const AX_FIELD_CODES = {
  "Scheduled from":   "scheduledStartTime",
  "Scheduled until":  "scheduledEndTime",
  "Actual start date": "actualStartDate",
  "Actual end date":   "actualEndDate",
};

// Whitelist of fields Ops can write to.
// Status changes intentionally excluded — workflow-controlled, separate permission flow.
const AX_WRITABLE_FIELDS = new Set(Object.keys(AX_FIELD_CODES));

// Subset of writable fields that are datetimes — these get dd-MM-yy HH:mm formatting.
const AX_DATETIME_FIELDS = new Set([
  "Scheduled from",
  "Scheduled until",
  "Actual start date",
  "Actual end date",
]);

function formatAxDateTime(value) {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value); // unparseable — pass through unchanged
  const pad = (n) => String(n).padStart(2, "0");
  return pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + String(d.getFullYear()).slice(-2)
    + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
}

async function logWrite(entry) {
  if (!pool) return;
  try {
    await pool.query(
      `INSERT INTO wo_writes (wo_ref, wo_id, field_name, old_value, new_value, user_email, mode, status, error_message)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
      [entry.ref, entry.id, entry.field, entry.oldValue, entry.newValue, entry.user, entry.mode, entry.status, entry.error]
    );
  } catch (e) { console.error("[WriteLog] Insert failed:", e.message); }
}

function findCachedWO(ref) {
  if (!cachedWO) return null;
  return cachedWO.find((r) => r.Reference === ref) || null;
}

function applyCachedUpdate(ref, field, value) {
  if (!cachedWO) return;
  const wo = cachedWO.find((r) => r.Reference === ref);
  if (wo) wo[field] = value;
}

async function callAxxerionWrite({ id, updates }) {
  if (!id) throw new Error("WorkOrder id required for live write");

  // Map display labels → internal field codes; format datetimes per API manual.
  const body = {};
  for (const [label, value] of Object.entries(updates)) {
    const code = AX_FIELD_CODES[label];
    if (!code) throw new Error("No internal field code mapped for: " + label);
    body[code] = AX_DATETIME_FIELDS.has(label) ? formatAxDateTime(value) : value;
  }

  const url = AX_WRITE_URL.replace(/\/$/, "") + "/" + encodeURIComponent(id);
  const resp = await fetch(url, {
    method: "PUT",
    headers: { Authorization: AX_AUTH, "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(30000),
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error("Axxerion " + resp.status + ": " + text.slice(0, 300));
  }
  return resp.json().catch(() => ({}));
}

app.post("/api/axxerion/update-wo", async (req, res) => {
  const { ref, id, updates, user } = req.body || {};
  if (!ref || !updates || typeof updates !== "object") {
    return res.status(400).json({ ok: false, error: "ref and updates object required" });
  }

  const fields = Object.keys(updates);
  const invalid = fields.filter((f) => !AX_WRITABLE_FIELDS.has(f));
  if (invalid.length) {
    return res.status(400).json({ ok: false, error: "Field(s) not writable: " + invalid.join(", ") });
  }
  if (!fields.length) {
    return res.status(400).json({ ok: false, error: "No fields to update" });
  }

  const cached = findCachedWO(ref);
  const woId = id || (cached && cached.ID) || null;
  const userEmail = user || (req.session && req.session.user && req.session.user.email) || "unknown";
  const mode = AX_WRITE_ENABLED ? "live" : "dryrun";
  const results = [];

  const oldValues = {};
  for (const field of fields) oldValues[field] = cached ? (cached[field] || "") : "";

  if (mode === "dryrun") {
    for (const field of fields) {
      await logWrite({
        ref, id: woId, field,
        oldValue: oldValues[field],
        newValue: String(updates[field] || ""),
        user: userEmail, mode: "dryrun", status: "success", error: null,
      });
      applyCachedUpdate(ref, field, updates[field]);
      results.push({ field, status: "dryrun", oldValue: oldValues[field], newValue: updates[field] });
    }
    return res.json({ ok: true, mode: "dryrun", message: "Dry-run: cache updated locally, Axxerion NOT called", results });
  }

  try {
    await callAxxerionWrite({ id: woId, ref, updates });
    for (const field of fields) {
      await logWrite({
        ref, id: woId, field,
        oldValue: oldValues[field],
        newValue: String(updates[field] || ""),
        user: userEmail, mode: "live", status: "success", error: null,
      });
      applyCachedUpdate(ref, field, updates[field]);
      results.push({ field, status: "success", oldValue: oldValues[field], newValue: updates[field] });
    }
    setTimeout(fetchWorkOrdersIncremental, 2000);
    res.json({ ok: true, mode: "live", message: "Axxerion updated", results });
  } catch (err) {
    for (const field of fields) {
      await logWrite({
        ref, id: woId, field,
        oldValue: oldValues[field],
        newValue: String(updates[field] || ""),
        user: userEmail, mode: "live", status: "error",
        error: err.message || String(err),
      });
    }
    res.status(502).json({ ok: false, mode: "live", error: err.message || String(err) });
  }
});

app.get("/api/axxerion/write-history", async (req, res) => {
  const ref = req.query.ref;
  if (!pool) return res.json({ history: [] });
  try {
    const sql = ref
      ? `SELECT * FROM wo_writes WHERE wo_ref = $1 ORDER BY attempted_at DESC LIMIT 100`
      : `SELECT * FROM wo_writes ORDER BY attempted_at DESC LIMIT 100`;
    const result = ref ? await pool.query(sql, [ref]) : await pool.query(sql);
    res.json({ history: result.rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/axxerion/write-status", (req, res) => {
  res.json({
    enabled: AX_WRITE_ENABLED,
    mode: AX_WRITE_ENABLED ? "live" : "dryrun",
    writableFields: Array.from(AX_WRITABLE_FIELDS),
    fieldCodes: AX_FIELD_CODES,
    endpointConfigured: !!AX_WRITE_URL,
  });
});

// ── Ops Queue Persistence ──
const OPS_FILE = path.join(__dirname, "data", "ops.json");

function loadOps() {
  try {
    if (fs.existsSync(OPS_FILE)) return JSON.parse(fs.readFileSync(OPS_FILE, "utf8"));
  } catch (e) { console.error("[Ops] Error loading ops.json:", e.message); }
  return { logs: {}, appointments: {}, emails: {}, vendors: {}, notes: {} };
}

function saveOps(data) {
  try {
    const dir = path.dirname(OPS_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(OPS_FILE, JSON.stringify(data, null, 2), "utf8");
  } catch (e) { console.error("[Ops] Error saving ops.json:", e.message); }
}

// Get all ops data
app.get("/api/ops", (req, res) => {
  res.json(loadOps());
});

// Log an action (call, email, note) for a WO
app.post("/api/ops/log", (req, res) => {
  const { ref, action, note, user } = req.body;
  if (!ref || !action) return res.status(400).json({ error: "ref and action required" });
  const ops = loadOps();
  if (!ops.logs[ref]) ops.logs[ref] = [];
  ops.logs[ref].unshift({ date: new Date().toISOString(), action, note: note || "", user: user || "ops" });
  saveOps(ops);
  res.json({ ok: true, logs: ops.logs[ref] });
});

// Set/update appointment date for a WO
app.post("/api/ops/appointment", (req, res) => {
  const { ref, date, confirmed, time } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  ops.appointments[ref] = { date: date || null, confirmed: !!confirmed, time: time || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true, appointment: ops.appointments[ref] });
});

// Track email sent to vendor
app.post("/api/ops/email", (req, res) => {
  const { ref, to, type, subject } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  if (!ops.emails[ref]) ops.emails[ref] = [];
  ops.emails[ref].unshift({ sentAt: new Date().toISOString(), to: to || "", type: type || "invoice", subject: subject || "" });
  saveOps(ops);
  res.json({ ok: true, emails: ops.emails[ref] });
});

// Store/update vendor contact info
app.post("/api/ops/vendor", (req, res) => {
  const { name, email, phone, contact } = req.body;
  if (!name) return res.status(400).json({ error: "vendor name required" });
  const ops = loadOps();
  ops.vendors[name] = { email: email || "", phone: phone || "", contact: contact || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true, vendor: ops.vendors[name] });
});

// Get vendor contact info
app.get("/api/ops/vendors", (req, res) => {
  const ops = loadOps();
  res.json(ops.vendors || {});
});

// Save/update notes for a WO
app.post("/api/ops/note", (req, res) => {
  const { ref, note } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  ops.notes[ref] = { text: note || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true });
});

// Delete a specific log entry or clear queue action
app.post("/api/ops/dismiss", (req, res) => {
  const { ref, queue } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  if (!ops.dismissed) ops.dismissed = {};
  if (!ops.dismissed[ref]) ops.dismissed[ref] = {};
  ops.dismissed[ref][queue] = new Date().toISOString();
  saveOps(ops);
  res.json({ ok: true });
});

// ── Snow Contract Persistence ──
const SNOW_FILE = path.join(__dirname, "data", "snow-contracts.json");

function loadSnowContracts() {
  try {
    if (fs.existsSync(SNOW_FILE)) return JSON.parse(fs.readFileSync(SNOW_FILE, "utf8"));
  } catch (e) { console.error("[Snow] Error loading snow-contracts.json:", e.message); }
  return {};
}

function saveSnowContracts(data) {
  try {
    const dir = path.dirname(SNOW_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(SNOW_FILE, JSON.stringify(data, null, 2), "utf8");
  } catch (e) { console.error("[Snow] Error saving snow-contracts.json:", e.message); }
}

// Get all snow contracts
app.get("/api/snow/contracts", (req, res) => {
  res.json(loadSnowContracts());
});

// Save/update a snow contract for a property
app.post("/api/snow/contract", (req, res) => {
  const { property, vendor, tiers, saltOnly, seasonStart, seasonEnd, accrualBuffer, notes } = req.body;
  if (!property) return res.status(400).json({ error: "property required" });
  const contracts = loadSnowContracts();
  contracts[property] = {
    vendor: vendor || "",
    tiers: tiers || [],
    saltOnly: saltOnly || 0,
    seasonStart: seasonStart || "11-01",
    seasonEnd: seasonEnd || "04-15",
    accrualBuffer: accrualBuffer || 1.20,
    notes: notes || "",
    updatedAt: new Date().toISOString()
  };
  saveSnowContracts(contracts);
  res.json({ ok: true, contract: contracts[property] });
});

// Delete a snow contract
app.post("/api/snow/contract/delete", (req, res) => {
  const { property } = req.body;
  if (!property) return res.status(400).json({ error: "property required" });
  const contracts = loadSnowContracts();
  delete contracts[property];
  saveSnowContracts(contracts);
  res.json({ ok: true });
});

// Health check (bypasses auth guard above)
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// Current user info for frontend (extended with admin flag + budget)
app.get("/api/me", async (req, res) => {
  const u = req.session.user;
  if (!u) return res.json(null);
  try {
    const budget = await chatUsage.checkBudget(pool, u.email, u.name);
    res.json({ ...u, isAdmin: budget.isAdmin, mtdCost: budget.mtdCost, monthlyBudget: budget.monthlyBudget, percentUsed: budget.percentUsed });
  } catch (e) {
    res.json({ ...u, isAdmin: chatUsage.isAdmin(u.email) });
  }
});

// ── Chat API ──
app.post("/api/chat", async (req, res) => {
  const u = req.session.user;
  if (!u) return res.status(401).json({ error: "Unauthorized" });
  const { messages } = req.body || {};
  if (!Array.isArray(messages) || !messages.length) return res.status(400).json({ error: "messages required" });

  // Check FAQ cache first — only for single-turn questions (first message or last message is standalone)
  const lastMsg = messages[messages.length - 1];
  if (lastMsg && lastMsg.role === "user" && messages.filter(m => m.role === "user").length === 1) {
    const cached = tryCache(lastMsg.content);
    if (cached.hit) {
      console.log(`[Chat] Cache hit: "${cached.cacheKey}" for "${lastMsg.content.slice(0, 60)}"`);
      try {
        const budget = await chatUsage.checkBudget(pool, u.email, u.name);
        return res.json({ reply: cached.reply, budget, toolCalls: [{ name: "cache", input: {}, summary: "instant · no API cost" }] });
      } catch (e) {}
    }
  }

  try {
    const result = await sendChatMessage({
      user: u,
      messages,
      pool,
      getData: () => ({ workOrders: cachedWO || [], requests: cachedReq || [] }),
    });
    res.json(result);
  } catch (e) {
    console.error("[/api/chat] error:", e.message);
    Sentry.withScope((scope) => {
      scope.setTag("where", "api.chat");
      scope.setUser({ email: u.email, username: u.name });
      scope.setExtra("messageCount", messages.length);
      Sentry.captureException(e);
    });
    res.status(500).json({ error: "Something went wrong with the AI assistant. Please reach out to IT support." });
  }
});

app.get("/api/chat/budget", async (req, res) => {
  const u = req.session.user;
  if (!u) return res.status(401).json({ error: "Unauthorized" });
  try {
    const b = await chatUsage.checkBudget(pool, u.email, u.name);
    res.json(b);
  } catch (e) { res.status(500).json({ error: "Failed to fetch budget" }); }
});

// ── Admin-only endpoints ──
function requireAdmin(req, res, next) {
  const u = req.session.user;
  if (!u) return res.status(401).json({ error: "Unauthorized" });
  chatUsage.checkBudget(pool, u.email, u.name)
    .then(b => { if (b.isAdmin) next(); else res.status(403).json({ error: "Admin access required" }); })
    .catch(() => res.status(500).json({ error: "Admin check failed" }));
}

app.get("/api/admin/usage", requireAdmin, async (req, res) => {
  try {
    const stats = await chatUsage.getUsageStats(pool);
    res.json(stats);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/api/admin/budget", requireAdmin, async (req, res) => {
  const { email, budget, name } = req.body || {};
  if (!email || budget == null) return res.status(400).json({ error: "email and budget required" });
  try {
    await chatUsage.upsertUserBudget(pool, email.toLowerCase().trim(), name || null, parseFloat(budget), req.session.user.email);
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post("/api/admin/add-admin", requireAdmin, async (req, res) => {
  const { email } = req.body || {};
  if (!email) return res.status(400).json({ error: "email required" });
  try {
    await chatUsage.addAdmin(pool, email.toLowerCase().trim(), req.session.user.email);
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post("/api/admin/remove-admin", requireAdmin, async (req, res) => {
  const { email } = req.body || {};
  if (!email) return res.status(400).json({ error: "email required" });
  try {
    await chatUsage.removeAdmin(pool, email.toLowerCase().trim(), req.session.user.email);
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Admin page (guarded)
app.get("/admin", async (req, res) => {
  const u = req.session.user;
  if (!u) return res.redirect("/auth/login-page");
  try {
    const b = await chatUsage.checkBudget(pool, u.email, u.name);
    if (!b.isAdmin) return res.status(403).send("Admin access required");
  } catch (e) { return res.status(500).send("Admin check failed"); }
  res.sendFile(path.join(__dirname, "public", "admin.html"));
});

app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Sentry error handler must be registered AFTER all controllers and BEFORE any other error middleware
Sentry.setupExpressErrorHandler(app);

app.listen(PORT, async () => {
  console.log(`Server running on port ${PORT}`);
  // 1. Init DB and load cached data (instant LIVE on deploy)
  await initDB();
  await chatUsage.initChatTables(pool);
  await loadFromDB();
  if (cachedWO) {
    console.log("[Startup] Dashboard ready with " + cachedWO.length + " WOs from DB cache");
    buildCache(() => ({ workOrders: cachedWO || [], requests: cachedReq || [] }));
  }
  // 2. Fetch fresh data — full WO pull on startup + every 4h, incremental every 10m, requests every 10m
  fetchWorkOrders();
  fetchRequests();
  setInterval(fetchWorkOrders, REFRESH_INTERVAL_WO_FULL);
  setInterval(fetchWorkOrdersIncremental, REFRESH_INTERVAL_WO_INCR);
  setInterval(fetchRequests, REFRESH_INTERVAL_REQ);
  console.log("[Startup] Timers: Full WOs every 4h, Incremental WOs every 10m, Requests every 10m");
});
