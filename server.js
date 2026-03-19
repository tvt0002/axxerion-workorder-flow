require("dotenv").config();
const express = require("express");
const path = require("path");
const https = require("https");
const fs = require("fs");

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

// ── Axxerion API config ──
const AX_USER = process.env.AX_USER;
const AX_PASS = process.env.AX_PASS;
if (!AX_USER || !AX_PASS) {
  console.error("Missing AX_USER or AX_PASS environment variables");
  process.exit(1);
}
const AX_AUTH = "Basic " + Buffer.from(`${AX_USER}:${AX_PASS}`).toString("base64");
const AX_REPORT_WO = "IPG-REP-085";
const AX_REPORT_REQ = "IPG-REP-087";
const REFRESH_INTERVAL = 10 * 60 * 1000; // 10 minutes
const RETRY_DELAY = 60 * 1000; // 1 minute retry when a fetch was skipped

let cachedWO = null;
let cacheWOTimestamp = 0;
let fetchWOInProgress = false;
let lastFetchStartTime = 0;

let cachedReq = null;
let cacheReqTimestamp = 0;
let fetchReqInProgress = false;

// Auth error backoff — stop retrying if credentials are rejected
let authDisabled = false;
let authErrorCount = 0;
const MAX_AUTH_ERRORS = 3;

function fetchReport(reference, label, onSuccess) {
  if (authDisabled) {
    console.warn(`[Cache] Skipping ${label} — auth disabled after ${authErrorCount} consecutive failures. Restart to retry.`);
    onSuccess(null);
    return;
  }

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

      // Stop immediately on auth errors (401/403)
      if (res.statusCode === 401 || res.statusCode === 403) {
        authErrorCount++;
        console.error(`[Cache] ${label} auth error (${res.statusCode}) — attempt ${authErrorCount}/${MAX_AUTH_ERRORS}`);
        if (authErrorCount >= MAX_AUTH_ERRORS) {
          authDisabled = true;
          console.error(`[Cache] AUTH DISABLED — ${MAX_AUTH_ERRORS} consecutive auth failures. Check AX_USER/AX_PASS. Restart to retry.`);
        }
        onSuccess(null);
        return;
      }

      // Reset auth error count on any successful response
      authErrorCount = 0;

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

  req.setTimeout(5 * 60 * 1000, () => {
    console.error(`[Cache] ${label} timed out after 5 minutes`);
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

function fetchFromAxxerion() {
  let skipped = false;

  if (!fetchWOInProgress) {
    fetchWOInProgress = true;
    lastFetchStartTime = Date.now();
    fetchReport(AX_REPORT_WO, "Work Orders", (data) => {
      fetchWOInProgress = false;
      if (data) { cachedWO = data; cacheWOTimestamp = Date.now(); }
    });
  } else {
    console.log("[Cache] Skipping Work Orders — previous fetch still in progress");
    skipped = true;
  }

  if (!fetchReqInProgress) {
    fetchReqInProgress = true;
    fetchReport(AX_REPORT_REQ, "Requests", (data) => {
      fetchReqInProgress = false;
      if (data) { cachedReq = data; cacheReqTimestamp = Date.now(); }
    });
  } else {
    console.log("[Cache] Skipping Requests — previous fetch still in progress");
    skipped = true;
  }

  if (skipped) {
    console.log("[Cache] Will retry skipped reports in 1 minute");
    setTimeout(fetchFromAxxerion, RETRY_DELAY);
  }
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
  const nextRefreshMs = lastFetchStartTime ? (lastFetchStartTime + REFRESH_INTERVAL) - Date.now() : 0;
  const nextRefreshMin = Math.max(0, Math.round(nextRefreshMs / 60000));
  res.json({
    workorders: { cached: !!cachedWO, count: cachedWO ? cachedWO.length : 0, ageMinutes: cachedWO ? Math.round((Date.now() - cacheWOTimestamp) / 60000) : null, fetchInProgress: fetchWOInProgress },
    requests: { cached: !!cachedReq, count: cachedReq ? cachedReq.length : 0, ageMinutes: cachedReq ? Math.round((Date.now() - cacheReqTimestamp) / 60000) : null, fetchInProgress: fetchReqInProgress },
    nextRefreshMin: nextRefreshMin,
    authDisabled: authDisabled,
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

app.get("/api/ops", (req, res) => { res.json(loadOps()); });

app.post("/api/ops/log", (req, res) => {
  const { ref, action, note, user } = req.body;
  if (!ref || !action) return res.status(400).json({ error: "ref and action required" });
  const ops = loadOps();
  if (!ops.logs[ref]) ops.logs[ref] = [];
  ops.logs[ref].unshift({ date: new Date().toISOString(), action, note: note || "", user: user || "ops" });
  saveOps(ops);
  res.json({ ok: true, logs: ops.logs[ref] });
});

app.post("/api/ops/appointment", (req, res) => {
  const { ref, date, confirmed, time } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  ops.appointments[ref] = { date: date || null, confirmed: !!confirmed, time: time || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true, appointment: ops.appointments[ref] });
});

app.post("/api/ops/email", (req, res) => {
  const { ref, to, type, subject } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  if (!ops.emails[ref]) ops.emails[ref] = [];
  ops.emails[ref].unshift({ sentAt: new Date().toISOString(), to: to || "", type: type || "invoice", subject: subject || "" });
  saveOps(ops);
  res.json({ ok: true, emails: ops.emails[ref] });
});

app.post("/api/ops/vendor", (req, res) => {
  const { name, email, phone, contact } = req.body;
  if (!name) return res.status(400).json({ error: "vendor name required" });
  const ops = loadOps();
  ops.vendors[name] = { email: email || "", phone: phone || "", contact: contact || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true, vendor: ops.vendors[name] });
});

app.get("/api/ops/vendors", (req, res) => { res.json(loadOps().vendors || {}); });

app.post("/api/ops/note", (req, res) => {
  const { ref, note } = req.body;
  if (!ref) return res.status(400).json({ error: "ref required" });
  const ops = loadOps();
  ops.notes[ref] = { text: note || "", updatedAt: new Date().toISOString() };
  saveOps(ops);
  res.json({ ok: true });
});

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

app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  // Fetch immediately on startup, then every 30 min
  fetchFromAxxerion();
  setInterval(fetchFromAxxerion, REFRESH_INTERVAL);
});
