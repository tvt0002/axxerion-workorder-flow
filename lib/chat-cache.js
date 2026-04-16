// lib/chat-cache.js — Pre-computed FAQ cache for instant chat responses
// Skips Claude API entirely for common questions, saving cost and latency.

const CLOSED_STATUSES = new Set(["Closed", "Invoiced", "Cancelled", "Completed"]);
function isOpen(status) { return !CLOSED_STATUSES.has(status); }
function parseNum(x) {
  if (x == null) return 0;
  if (typeof x === "number") return x;
  const n = parseFloat(String(x).replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}
function fmtCurrency(n) {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ── Cached answers (rebuilt on data refresh) ──
let cache = {};
let cacheBuiltAt = 0;

function buildCache(getData) {
  const { workOrders, requests } = getData();
  if (!workOrders || !workOrders.length) return;
  const now = Date.now();

  const openWOs = workOrders.filter(r => isOpen(r["Status"]));

  // 1. Open critical WOs
  const criticalOpen = openWOs.filter(r => r["Priority"] === "Critical");
  let criticalText = `**${criticalOpen.length} open Critical work orders** right now.`;
  if (criticalOpen.length > 0) {
    const top = criticalOpen.slice(0, 8);
    criticalText += "\n\n| Reference | Property | Status | Vendor | Subject |\n|---|---|---|---|---|";
    top.forEach(r => {
      criticalText += `\n| ${r["Reference"]} | ${r["Property"] || "—"} | ${r["Status"]} | ${r["Vendor"] || "—"} | ${(r["Subject"] || "").slice(0, 50)} |`;
    });
    if (criticalOpen.length > 8) criticalText += `\n\n...and ${criticalOpen.length - 8} more. Ask me to filter by region, property, or FM for a narrower list.`;
  }
  cache["open_critical"] = criticalText;

  // 2. Overall open WO stats
  const byPriority = {};
  const byStatus = {};
  const byRegion = {};
  let totalNTE = 0, totalSpend = 0, nteExceeded = 0;
  openWOs.forEach(r => {
    byPriority[r["Priority"]] = (byPriority[r["Priority"]] || 0) + 1;
    byStatus[r["Status"]] = (byStatus[r["Status"]] || 0) + 1;
    if (r["Region"]) byRegion[r["Region"]] = (byRegion[r["Region"]] || 0) + 1;
    totalNTE += parseNum(r["Not To Exceed"]);
    totalSpend += parseNum(r["Total"]);
    if (r["NTE Exceeded"] === "Yes") nteExceeded++;
  });

  let statsText = `**${openWOs.length} open work orders** across ${Object.keys(byRegion).length} regions.\n\n`;
  statsText += "**By Priority:**\n";
  ["Critical", "High", "Medium", "Low", "Very Low"].forEach(p => {
    if (byPriority[p]) statsText += `- ${p}: ${byPriority[p]}\n`;
  });
  statsText += `\nTotal NTE: ${fmtCurrency(totalNTE)} · Total Spend: ${fmtCurrency(totalSpend)} · ${nteExceeded} over budget`;
  cache["open_stats"] = statsText;

  // 3. FM workload
  const fmData = {};
  openWOs.forEach(r => {
    const fm = r["Facility Manager"] || "(Unassigned)";
    if (!fmData[fm]) fmData[fm] = { total: 0, critical: 0, high: 0, nte: 0, spend: 0 };
    fmData[fm].total++;
    if (r["Priority"] === "Critical") fmData[fm].critical++;
    if (r["Priority"] === "High") fmData[fm].high++;
    fmData[fm].nte += parseNum(r["Not To Exceed"]);
    fmData[fm].spend += parseNum(r["Total"]);
  });
  let fmText = "**FM Workload (open WOs):**\n\n| Facility Manager | Open WOs | Critical | High | Total NTE | Total Spend |\n|---|---|---|---|---|---|";
  Object.entries(fmData).sort((a, b) => b[1].total - a[1].total).forEach(([fm, d]) => {
    fmText += `\n| ${fm} | ${d.total} | ${d.critical} | ${d.high} | ${fmtCurrency(d.nte)} | ${fmtCurrency(d.spend)} |`;
  });
  cache["fm_workload"] = fmText;

  // 4. Top vendors by spend (all time)
  const vendorSpend = {};
  workOrders.forEach(r => {
    const v = r["Vendor"];
    if (!v) return;
    if (!vendorSpend[v]) vendorSpend[v] = { count: 0, spend: 0, nte: 0 };
    vendorSpend[v].count++;
    vendorSpend[v].spend += parseNum(r["Total"]);
    vendorSpend[v].nte += parseNum(r["Not To Exceed"]);
  });
  const topVendors = Object.entries(vendorSpend).sort((a, b) => b[1].spend - a[1].spend).slice(0, 10);
  let vendorText = "**Top 10 vendors by total spend:**\n\n| Vendor | WO Count | Total Spend | Total NTE |\n|---|---|---|---|";
  topVendors.forEach(([v, d]) => {
    vendorText += `\n| ${v} | ${d.count} | ${fmtCurrency(d.spend)} | ${fmtCurrency(d.nte)} |`;
  });
  cache["top_vendors"] = vendorText;

  // 5. NTE exceeded
  const exceeded = openWOs.filter(r => r["NTE Exceeded"] === "Yes");
  let nteText = `**${exceeded.length} open WOs over budget (NTE exceeded).**`;
  if (exceeded.length > 0) {
    const top = exceeded.sort((a, b) => (parseNum(b["Total"]) - parseNum(b["Not To Exceed"])) - (parseNum(a["Total"]) - parseNum(a["Not To Exceed"]))).slice(0, 8);
    nteText += "\n\n| Reference | Property | Vendor | NTE | Actual | Over By |\n|---|---|---|---|---|---|";
    top.forEach(r => {
      const nte = parseNum(r["Not To Exceed"]);
      const actual = parseNum(r["Total"]);
      nteText += `\n| ${r["Reference"]} | ${r["Property"] || "—"} | ${r["Vendor"] || "—"} | ${fmtCurrency(nte)} | ${fmtCurrency(actual)} | ${fmtCurrency(actual - nte)} |`;
    });
    if (exceeded.length > 8) nteText += `\n\n...and ${exceeded.length - 8} more.`;
  }
  cache["nte_exceeded"] = nteText;

  // 6. Open requests summary
  if (requests && requests.length) {
    const openReqs = requests.filter(r => r["Status"] !== "Closed" && r["Status"] !== "Cancelled");
    const reqByStatus = {};
    openReqs.forEach(r => { reqByStatus[r["Status"]] = (reqByStatus[r["Status"]] || 0) + 1; });
    let reqText = `**${openReqs.length} open maintenance requests.**\n\n| Status | Count |\n|---|---|`;
    Object.entries(reqByStatus).sort((a, b) => b[1] - a[1]).forEach(([s, c]) => {
      reqText += `\n| ${s} | ${c} |`;
    });
    cache["open_requests"] = reqText;
  }

  // 7. By-region breakdown
  let regionText = `**Open WOs by region:**\n\n| Region | Count |\n|---|---|`;
  Object.entries(byRegion).sort((a, b) => b[1] - a[1]).forEach(([r, c]) => {
    regionText += `\n| ${r} | ${c} |`;
  });
  cache["region_breakdown"] = regionText;

  cacheBuiltAt = now;
  console.log("[ChatCache] Built " + Object.keys(cache).length + " cached answers");
}

// ── Pattern matching: map user message → cache key ──
const PATTERNS = [
  { key: "open_critical", patterns: [
    /how many.*(?:open|active).*critical/i,
    /(?:open|active).*critical.*(?:wo|work order)/i,
    /critical.*(?:open|active)/i,
    /^critical\s*(?:wo|work order)?s?\s*$/i,
  ]},
  { key: "open_stats", patterns: [
    /(?:overall|total|all).*(?:open|active).*(?:stat|count|summary|overview)/i,
    /how many.*(?:open|active).*(?:wo|work order)/i,
    /(?:open|active).*(?:wo|work order).*(?:count|total|stat)/i,
    /^(?:open|active)\s*(?:wo|work order)?s?\s*$/i,
    /how are we doing/i,
    /give me (?:a |the )?(?:overview|summary|snapshot)/i,
    /^dashboard\s*(?:stats|summary)?$/i,
  ]},
  { key: "fm_workload", patterns: [
    /(?:fm|facility manager).*workload/i,
    /workload.*(?:fm|facility manager)/i,
    /(?:who|which fm).*(?:busiest|most work)/i,
    /show me.*workload/i,
    /^workload$/i,
  ]},
  { key: "top_vendors", patterns: [
    /(?:top|most).*vendor.*(?:spend|spent|cost|expensive)/i,
    /vendor.*(?:most|highest).*(?:spend|spent|cost)/i,
    /(?:which|what) vendor.*(?:spent|spend) the most/i,
    /biggest vendor/i,
  ]},
  { key: "nte_exceeded", patterns: [
    /(?:nte|budget).*(?:exceeded|over|blown)/i,
    /(?:exceeded|over).*(?:nte|budget)/i,
    /over.*budget/i,
    /(?:wo|work order).*(?:exceeded|over).*nte/i,
  ]},
  { key: "open_requests", patterns: [
    /(?:open|active|pending).*(?:request|wr|maintenance request)/i,
    /(?:request|wr).*(?:count|total|stat|summary)/i,
    /how many.*request/i,
  ]},
  { key: "region_breakdown", patterns: [
    /(?:wo|work order).*(?:by|per|each) region/i,
    /region.*breakdown/i,
    /breakdown.*region/i,
    /^by region$/i,
  ]},
];

/**
 * Try to match a user message to a cached answer.
 * Returns { hit: true, reply: "...", cacheKey: "..." } or { hit: false }
 */
function tryCache(userMessage) {
  if (!cacheBuiltAt || !userMessage) return { hit: false };
  const msg = userMessage.trim();
  for (const { key, patterns } of PATTERNS) {
    for (const rx of patterns) {
      if (rx.test(msg) && cache[key]) {
        return { hit: true, reply: cache[key], cacheKey: key };
      }
    }
  }
  return { hit: false };
}

module.exports = { buildCache, tryCache };
