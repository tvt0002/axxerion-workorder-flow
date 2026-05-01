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
const audit = require("./lib/audit");
const calldirJob = require("./lib/calllist-job");

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
    audit.logEvent({
      actor: req.session.user.email,
      actorType: "user",
      action: "user.login",
      entityType: "user",
      entityId: req.session.user.email,
      metadata: { name: req.session.user.name },
      sessionId: req.sessionID,
    });
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
    audit.init(pool);
    await audit.ensureTable();
    console.log("[DB] Cache + audit_log + wo_writes tables ready");
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
      // Capture diffs before merging — fire audit events asynchronously after.
      const diffs = [];
      const newWOs = [];
      data.forEach((wo) => {
        const existingIdx = woMap.get(wo.ID);
        if (existingIdx !== undefined) {
          const prior = cachedWO[existingIdx];
          diffs.push({ prior, next: wo });
          cachedWO[existingIdx] = wo;
          updated++;
        } else {
          newWOs.push(wo);
          cachedWO.push(wo);
          woMap.set(wo.ID, cachedWO.length - 1);
          added++;
        }
      });
      // Fire-and-forget audit writes — never block the merge.
      (async () => {
        try {
          for (const d of diffs) await audit.diffAndLogWO(d.prior, d.next);
          for (const wo of newWOs) await audit.logWOCreated(wo);
        } catch (e) { console.error("[Audit] diff log error:", e.message); }
      })();

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
//
// "Start"/"End" map to startDate/endDate — the fields vendors actually populate when
// they click [Start Work] / [End Work] in the Axxerion UI (8,434 of 12,151 WOs have
// `endDate` populated vs only 1,936 with `actualEndDate`, per the 2026-04-30 discovery).
// "Actual start date"/"Actual end date" mappings are kept for back-compat but rarely used.
const AX_FIELD_CODES = {
  "Scheduled from":    "scheduledStartTime",
  "Scheduled until":   "scheduledEndTime",
  "Actual start date": "actualStartDate",
  "Actual end date":   "actualEndDate",
  "Start":             "startDate",
  "End":               "endDate",
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
  "Start",
  "End",
]);

// Axxerion stores datetimes as UTC (per San 2026-04-29). Use UTC accessors so a
// caller-provided wall-clock (or ISO with offset) lands at the intended UTC instant.
function formatAxDateTime(value) {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  const pad = (n) => String(n).padStart(2, "0");
  return pad(d.getUTCDate()) + "-" + pad(d.getUTCMonth() + 1) + "-" + String(d.getUTCFullYear()).slice(-2)
    + " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes());
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
    transitions: AX_WORKFLOW_TRANSITIONS,
    endpointConfigured: !!AX_WRITE_URL,
  });
});

// ── Workflow Status Transitions ──
// PUT /webservices/ipg/rest/functions/executefunction/WorkOrder/{id}/{functionName}
// Function names confirmed by San on 2026-04-29 (workflow perms still pending).
// Casing matters: startWork is camelCase, others lowercase.
const AX_EXECUTE_URL = process.env.AXXERION_EXECUTE_URL ||
  "https://ipg.axxerion.us/webservices/ipg/rest/functions/executefunction/WorkOrder";

const AX_WORKFLOW_TRANSITIONS = [
  { from: "Assigned",            to: "Accepted",             fn: "confirmworker" },
  { from: "Assigned",            to: "Need Reassignment",    fn: "reassign" },
  { from: "Accepted",            to: "Work In Progress",     fn: "startWork" },
  { from: "Work In Progress",    to: "Work Finished",        fn: "stopwork" },
  { from: "Prepare Financials",  to: "Financials Submitted", fn: "submitfinancials" },
];

function findTransition(fromStatus, toStatus) {
  const f = String(fromStatus || "").trim();
  const t = String(toStatus || "").trim();
  return AX_WORKFLOW_TRANSITIONS.find((x) =>
    x.from.toLowerCase() === f.toLowerCase() && x.to.toLowerCase() === t.toLowerCase()
  );
}

async function callAxxerionTransition({ id, functionName }) {
  if (!id) throw new Error("WorkOrder id required for transition");
  if (!functionName) throw new Error("functionName required");
  const url = AX_EXECUTE_URL.replace(/\/$/, "") + "/" + encodeURIComponent(id) + "/" + encodeURIComponent(functionName);
  const resp = await fetch(url, {
    method: "PUT",
    headers: { Authorization: AX_AUTH, "Content-Type": "application/json" },
    body: "{}",
    signal: AbortSignal.timeout(30000),
  });
  const text = await resp.text().catch(() => "");
  let json = null; try { json = JSON.parse(text); } catch {}
  if (!resp.ok) throw new Error("Axxerion " + resp.status + ": " + text.slice(0, 300));
  // Axxerion returns 200 with errorMessage when permission denied / workflow blocked.
  if (json && json.errorMessage) throw new Error(json.errorMessage + (json.warningMessage ? " — " + json.warningMessage : ""));
  return json || {};
}

app.post("/api/axxerion/transition-status", async (req, res) => {
  const { ref, id, toStatus, user } = req.body || {};
  if (!ref || !toStatus) return res.status(400).json({ ok: false, error: "ref and toStatus required" });

  const cached = findCachedWO(ref);
  const woId = id || (cached && cached.ID) || null;
  const fromStatus = cached ? (cached.Status || "") : "";
  const transition = findTransition(fromStatus, toStatus);
  if (!transition) {
    return res.status(400).json({
      ok: false,
      error: "No workflow function for transition: " + fromStatus + " → " + toStatus,
      validTransitions: AX_WORKFLOW_TRANSITIONS.filter((x) => x.from.toLowerCase() === String(fromStatus).toLowerCase()),
    });
  }

  const userEmail = user || (req.session && req.session.user && req.session.user.email) || "unknown";
  const mode = AX_WRITE_ENABLED ? "live" : "dryrun";

  if (mode === "dryrun") {
    await logWrite({
      ref, id: woId, field: "status_transition",
      oldValue: fromStatus, newValue: toStatus,
      user: userEmail, mode: "dryrun", status: "success", error: null,
    });
    applyCachedUpdate(ref, "Status", toStatus);
    return res.json({ ok: true, mode: "dryrun", transition, message: "Dry-run: cache updated, Axxerion NOT called" });
  }

  try {
    await callAxxerionTransition({ id: woId, functionName: transition.fn });
    await logWrite({
      ref, id: woId, field: "status_transition",
      oldValue: fromStatus, newValue: toStatus,
      user: userEmail, mode: "live", status: "success", error: null,
    });
    applyCachedUpdate(ref, "Status", toStatus);
    setTimeout(fetchWorkOrdersIncremental, 2000);
    res.json({ ok: true, mode: "live", transition, message: "Status transitioned in Axxerion" });
  } catch (err) {
    await logWrite({
      ref, id: woId, field: "status_transition",
      oldValue: fromStatus, newValue: toStatus,
      user: userEmail, mode: "live", status: "error", error: err.message || String(err),
    });
    res.status(502).json({ ok: false, mode: "live", error: err.message || String(err) });
  }
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

// ── Vendor Profile Aggregation ──
// Aggregates cachedWO into per-vendor stats: WO count, spend, NTE, open count,
// top properties + services, last WO, plus merged contact info from OPS_DATA + Axxerion.
const TERMINAL_VENDOR_STATUSES = new Set(["Closed", "Cancelled", "Invoiced", "Completed", "Financials Submitted", "No Invoice", "Warranty", "Internal"]);

function buildVendorProfiles() {
  if (!cachedWO || !cachedWO.length) return [];
  const ops = loadOps();
  const opsVendors = ops.vendors || {};
  const map = new Map();

  for (const r of cachedWO) {
    const name = (r["Vendor"] || "").trim();
    if (!name) continue;
    let v = map.get(name);
    if (!v) {
      v = {
        name,
        woCount: 0,
        openCount: 0,
        totalSpend: 0,
        totalNTE: 0,
        properties: new Map(),
        services: new Map(),
        statuses: new Map(),
        lastWODate: null,
        contact: { phone: "", mobile: "", email: "" },
        recentRefs: [],
      };
      map.set(name, v);
    }
    v.woCount++;
    const status = String(r["Status"] || "");
    if (!TERMINAL_VENDOR_STATUSES.has(status)) v.openCount++;
    const spend = parseFloat(r["Total"] || "0") || 0;
    const nte = parseFloat(r["Not To Exceed"] || "0") || 0;
    v.totalSpend += spend;
    v.totalNTE += nte;
    const prop = r["Property"] || "(unknown)";
    v.properties.set(prop, (v.properties.get(prop) || 0) + 1);
    const svc = r["Problem Type"] || "(unknown)";
    v.services.set(svc, (v.services.get(svc) || 0) + 1);
    v.statuses.set(status || "(blank)", (v.statuses.get(status || "(blank)") || 0) + 1);
    const created = r["Created"] || "";
    if (created && (!v.lastWODate || new Date(created) > new Date(v.lastWODate))) v.lastWODate = created;
    if (!v.contact.phone && r["Phone"]) v.contact.phone = r["Phone"];
    if (!v.contact.mobile && r["Mobile"]) v.contact.mobile = r["Mobile"];
    if (!v.contact.email && r["Email"]) v.contact.email = r["Email"];
  }

  // Merge OPS_DATA vendor contact (manual entries take precedence over Axxerion).
  for (const [name, v] of map) {
    const manual = opsVendors[name] || {};
    if (manual.phone) v.contact.phone = manual.phone;
    if (manual.email) v.contact.email = manual.email;
    if (manual.contact) v.contact.contactName = manual.contact;
  }

  // Convert maps to sorted arrays + serialize.
  return Array.from(map.values()).map((v) => ({
    name: v.name,
    woCount: v.woCount,
    openCount: v.openCount,
    totalSpend: Math.round(v.totalSpend * 100) / 100,
    totalNTE: Math.round(v.totalNTE * 100) / 100,
    propertyCount: v.properties.size,
    serviceCount: v.services.size,
    topProperties: Array.from(v.properties.entries()).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, count]) => ({ name, count })),
    topServices: Array.from(v.services.entries()).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([name, count]) => ({ name, count })),
    statuses: Array.from(v.statuses.entries()).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count })),
    lastWODate: v.lastWODate,
    contact: v.contact,
  })).sort((a, b) => b.woCount - a.woCount);
}

app.get("/api/vendors/profiles", (req, res) => {
  res.json({ vendors: buildVendorProfiles(), generatedAt: new Date().toISOString() });
});

app.get("/api/vendors/profile/:name", (req, res) => {
  const target = decodeURIComponent(req.params.name);
  const profile = buildVendorProfiles().find((v) => v.name.toLowerCase() === target.toLowerCase());
  if (!profile) return res.status(404).json({ error: "Vendor not found", name: target });
  // Include the full WO list for this vendor (most recent first).
  const wos = (cachedWO || [])
    .filter((r) => (r["Vendor"] || "").trim().toLowerCase() === target.toLowerCase())
    .map((r) => ({
      ref: r["Reference"] || "",
      property: r["Property"] || "",
      status: r["Status"] || "",
      priority: r["Priority"] || "",
      service: r["Problem Type"] || "",
      created: r["Created"] || "",
      total: parseFloat(r["Total"] || "0") || 0,
      nte: parseFloat(r["Not To Exceed"] || "0") || 0,
      bookmark: r["Bookmark"] || "",
    }))
    .sort((a, b) => new Date(b.created) - new Date(a.created));
  res.json({ profile, workorders: wos });
});

// Log an action (call, email, note) for a WO
app.post("/api/ops/log", (req, res) => {
  const { ref, action, note, user } = req.body;
  if (!ref || !action) return res.status(400).json({ error: "ref and action required" });
  const ops = loadOps();
  if (!ops.logs[ref]) ops.logs[ref] = [];
  const entry = { date: new Date().toISOString(), action, note: note || "", user: user || "ops" };
  ops.logs[ref].unshift(entry);
  saveOps(ops);
  const a = audit.actorFromReq(req);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: "call.logged",
    entityType: "wo",
    entityId: ref,
    newValue: { action, note: note || "" },
  });
  res.json({ ok: true, logs: ops.logs[ref] });
});

// Set/update appointment date for a WO
app.post("/api/ops/appointment", (req, res) => {
  const { ref, date, confirmed, time } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  const old = ops.appointments[ref] || null;
  const next = { date: date || null, confirmed: !!confirmed, time: time || "", updatedAt: new Date().toISOString() };
  ops.appointments[ref] = next;
  saveOps(ops);
  const a = audit.actorFromReq(req);
  // Distinguish "newly confirmed" from generic schedule update
  const isConfirm = !!confirmed && (!old || !old.confirmed);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: isConfirm ? "appt.confirmed" : "appt.set",
    entityType: "wo",
    entityId: ref,
    oldValue: old,
    newValue: next,
  });
  res.json({ ok: true, appointment: ops.appointments[ref] });
});

// Track email sent to vendor
app.post("/api/ops/email", (req, res) => {
  const { ref, to, type, subject } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  if (!ops.emails[ref]) ops.emails[ref] = [];
  const entry = { sentAt: new Date().toISOString(), to: to || "", type: type || "invoice", subject: subject || "" };
  ops.emails[ref].unshift(entry);
  saveOps(ops);
  const a = audit.actorFromReq(req);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: "email.sent",
    entityType: "wo",
    entityId: ref,
    newValue: entry,
  });
  res.json({ ok: true, emails: ops.emails[ref] });
});

// Store/update vendor contact info
app.post("/api/ops/vendor", (req, res) => {
  const { name, email, phone, contact } = req.body;
  if (!name) return res.status(400).json({ error: "vendor name required" });
  const ops = loadOps();
  const old = ops.vendors[name] || null;
  const next = { email: email || "", phone: phone || "", contact: contact || "", updatedAt: new Date().toISOString() };
  ops.vendors[name] = next;
  saveOps(ops);
  const a = audit.actorFromReq(req);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: "vendor.contact_updated",
    entityType: "vendor",
    entityId: name,
    oldValue: old,
    newValue: next,
  });
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
  const old = ops.notes[ref] || null;
  ops.notes[ref] = { text: note || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  const a = audit.actorFromReq(req);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: "note.added",
    entityType: "wo",
    entityId: ref,
    oldValue: old ? { text: old.text } : null,
    newValue: { text: note || "" },
  });
  res.json({ ok: true });
});

// ── Audit log read endpoints ──
// Per-entity timeline (drives per-WO history drawer)
app.get("/api/audit/entity/:type/:id", async (req, res) => {
  try {
    const rows = await audit.getEntityHistory(req.params.type, req.params.id, parseInt(req.query.limit, 10) || 200);
    res.json({ rows });
  } catch (e) {
    console.error("[Audit] entity read error:", e.message);
    res.status(500).json({ error: "Failed to load history" });
  }
});

// Login frequency stats per user (drives Logins tab)
app.get("/api/audit/login-stats", async (req, res) => {
  try {
    const stats = await audit.getLoginStats();
    res.json({ rows: stats });
  } catch (e) {
    console.error("[Audit] login-stats error:", e.message);
    res.status(500).json({ error: "Failed to load login stats" });
  }
});

// Per-user activity stats (drives Activity tab)
app.get("/api/audit/activity-stats", async (req, res) => {
  try {
    const days = Math.max(1, Math.min(365, parseInt(req.query.days, 10) || 7));
    const stats = await audit.getActivityStats(days);
    res.json({ rows: stats, days });
  } catch (e) {
    console.error("[Audit] activity-stats error:", e.message);
    res.status(500).json({ error: "Failed to load activity stats" });
  }
});

// System feed with optional filters (drives /audit page + my-day view)
app.get("/api/audit", async (req, res) => {
  try {
    const filters = {};
    let actor = req.query.actor;
    if (actor === "me" && req.session && req.session.user) actor = req.session.user.email;
    if (actor) filters.actor = actor;
    if (req.query.actorType) filters.actorType = req.query.actorType;
    if (req.query.action) filters.action = req.query.action;
    if (req.query.entityType) filters.entityType = req.query.entityType;
    if (req.query.entityId) filters.entityId = req.query.entityId;
    if (req.query.since) filters.since = req.query.since;
    if (req.query.until) filters.until = req.query.until;
    filters.limit = parseInt(req.query.limit, 10) || 100;
    filters.offset = parseInt(req.query.offset, 10) || 0;
    const result = await audit.getAuditFeed(filters);
    res.json({ ...result, currentUser: req.session && req.session.user ? req.session.user.email : null });
  } catch (e) {
    console.error("[Audit] feed read error:", e.message);
    res.status(500).json({ error: "Failed to load audit feed" });
  }
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
  const a = audit.actorFromReq(req);
  audit.logEvent({
    actor: a.actor, actorType: a.actorType, sessionId: a.sessionId,
    action: "wo.dismissed",
    entityType: "wo",
    entityId: ref,
    metadata: { queue: queue || null },
  });
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
    audit.logEvent({
      actor: req.session.user.email, actorType: "user", sessionId: req.sessionID,
      action: "budget.set",
      entityType: "user",
      entityId: email.toLowerCase().trim(),
      newValue: { budget: parseFloat(budget) },
      metadata: { name: name || null },
    });
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post("/api/admin/add-admin", requireAdmin, async (req, res) => {
  const { email } = req.body || {};
  if (!email) return res.status(400).json({ error: "email required" });
  try {
    await chatUsage.addAdmin(pool, email.toLowerCase().trim(), req.session.user.email);
    audit.logEvent({
      actor: req.session.user.email, actorType: "user", sessionId: req.sessionID,
      action: "user.role_changed",
      entityType: "user",
      entityId: email.toLowerCase().trim(),
      newValue: { role: "admin" },
      metadata: { granted_by: req.session.user.email },
    });
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.post("/api/admin/remove-admin", requireAdmin, async (req, res) => {
  const { email } = req.body || {};
  if (!email) return res.status(400).json({ error: "email required" });
  try {
    await chatUsage.removeAdmin(pool, email.toLowerCase().trim(), req.session.user.email);
    audit.logEvent({
      actor: req.session.user.email, actorType: "user", sessionId: req.sessionID,
      action: "user.role_changed",
      entityType: "user",
      entityId: email.toLowerCase().trim(),
      oldValue: { role: "admin" },
      newValue: { role: "user" },
      metadata: { revoked_by: req.session.user.email },
    });
    res.json({ ok: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Manual trigger for the call directory sync (admin-only).
// Body: { dryRun?: boolean, force?: boolean }
app.post("/api/admin/sync-calldir", requireAdmin, async (req, res) => {
  const { dryRun = false, force = false } = req.body || {};
  try {
    const result = await calldirJob.runSync({ pool, dryRun, force });
    audit.logEvent({
      actor: req.session.user.email, actorType: "user", sessionId: req.sessionID,
      action: "calldir.sync_triggered",
      entityType: "calldir_sync",
      entityId: String(result.run_id || 'manual'),
      metadata: { dryRun, force, summary: result.summary, action: result.action },
    });
    res.json(result);
  } catch (e) { res.status(500).json({ error: e.message }); }
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

  // 3. Daily call-directory sync at 07:00 America/New_York.
  await calldirJob.ensureStateTable(pool);
  let calldirLastRunDate = null;
  setInterval(() => {
    const now = new Date();
    const etTime = now.toLocaleString("en-US", {
      timeZone: "America/New_York", hour: "2-digit", minute: "2-digit", hour12: false,
    });
    const etDate = now.toLocaleDateString("en-US", { timeZone: "America/New_York" });
    if (etTime === "07:00" && etDate !== calldirLastRunDate) {
      calldirLastRunDate = etDate;
      console.log("[calldir-sync] daily 07:00 ET trigger");
      calldirJob.runSync({ pool }).catch(err => {
        console.error("[calldir-sync] daily run failed:", err.message);
        Sentry.captureException(err, { tags: { job: "calldir-sync" } });
      });
    }
  }, 60_000);
  console.log("[Startup] Call directory sync scheduled daily at 07:00 ET");
});
