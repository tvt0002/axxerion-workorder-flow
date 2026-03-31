const express = require("express");
const path = require("path");
const https = require("https");
const fs = require("fs");
const { Pool } = require("pg");

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

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
    console.log("[DB] Cache table ready");
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
const AX_REPORT_REQ = "IPG-REP-087";
const REFRESH_INTERVAL_WO = 4 * 60 * 60 * 1000; // 4 hours for full WO pull (11K+ records)
const REFRESH_INTERVAL_REQ = 10 * 60 * 1000; // 10 minutes for requests (small dataset)
const RETRY_DELAY = 60 * 1000; // 1 minute retry when a fetch was skipped

let cachedWO = null;
let cacheWOTimestamp = 0;
let fetchWOInProgress = false;
let lastFetchStartTime = 0;

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
    if (data) { cachedWO = data; cacheWOTimestamp = Date.now(); saveToDB("workorders", data); }
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
    if (data) { cachedReq = data; cacheReqTimestamp = Date.now(); saveToDB("requests", data); }
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
  const nextWORefreshMs = lastFetchStartTime ? (lastFetchStartTime + REFRESH_INTERVAL_WO) - Date.now() : 0;
  const nextWORefreshMin = Math.max(0, Math.round(nextWORefreshMs / 60000));
  res.json({
    workorders: { cached: !!cachedWO, count: cachedWO ? cachedWO.length : 0, ageMinutes: cachedWO ? Math.round((Date.now() - cacheWOTimestamp) / 60000) : null, fetchInProgress: fetchWOInProgress, refreshIntervalHrs: 4 },
    requests: { cached: !!cachedReq, count: cachedReq ? cachedReq.length : 0, ageMinutes: cachedReq ? Math.round((Date.now() - cacheReqTimestamp) / 60000) : null, fetchInProgress: fetchReqInProgress, refreshIntervalMin: 10 },
    nextWORefreshMin: nextWORefreshMin,
    authFailed: authFailed,
    dbConnected: !!pool,
  });
});

// ── Manual refresh trigger ──
app.post("/api/refresh", (req, res) => {
  fetchWorkOrders();
  fetchRequests();
  res.json({ ok: true, message: "Refresh triggered" });
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

app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, async () => {
  console.log(`Server running on port ${PORT}`);
  // 1. Init DB and load cached data (instant LIVE on deploy)
  await initDB();
  await loadFromDB();
  if (cachedWO) console.log("[Startup] Dashboard ready with " + cachedWO.length + " WOs from DB cache");
  // 2. Fetch fresh data — Requests every 10 min, full WO pull every 4 hours
  fetchWorkOrders();
  fetchRequests();
  setInterval(fetchWorkOrders, REFRESH_INTERVAL_WO);
  setInterval(fetchRequests, REFRESH_INTERVAL_REQ);
  console.log("[Startup] Timers: WOs every 4h, Requests every 10m");
});
