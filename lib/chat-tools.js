// lib/chat-tools.js — Claude tool definitions + handlers for Axxerion FM data
// Data source: in-memory cachedWO / cachedReq arrays passed in via getData()

const { STATUS_GUIDE, getStatusInfo, listStatusesByParty } = require("./status-guide");

const MAX_RESULTS = 25;

const TOOL_DEFINITIONS = [
  {
    name: "search_work_orders",
    description: "Search work orders with flexible filters. Use this for questions like 'show me open WOs in the Northwest', 'find WOs for ABC Plumbing', 'critical WOs assigned to Tim MacVittie'. Returns matching WOs (capped at 25). For broader stats use get_stats instead.",
    input_schema: {
      type: "object",
      properties: {
        property: { type: "string", description: "Partial property name match (e.g. 'Seattle', 'L050')" },
        vendor: { type: "string", description: "Partial vendor name match" },
        status: { type: "string", description: "Status filter (e.g. 'Assigned', 'Work finished', 'Closed'). Use 'open' to mean anything not Closed/Invoiced/Cancelled/Completed." },
        priority: { type: "string", description: "Priority: Critical, High, Medium, Low, Very Low" },
        region: { type: "string", description: "Region filter (e.g. 'PNW', 'Northeast', 'California 1')" },
        dm: { type: "string", description: "Area Manager / District Manager name (partial match)" },
        facility_manager: { type: "string", description: "Facility Manager name: 'Tim MacVittie' or 'Felipe Lopez'" },
        problem_type: { type: "string", description: "Problem type / service type (e.g. 'Elevator', 'Plumbing', 'HVAC')" },
        nte_exceeded: { type: "boolean", description: "Filter to WOs where spend exceeded NTE" },
        created_since_days: { type: "number", description: "Only WOs created within the last N days" },
        limit: { type: "number", description: "Max results to return (default 25)" },
      },
    },
  },
  {
    name: "get_work_order_details",
    description: "Get full details for a single work order by its Reference number (e.g. 'WRK-261615'). Use when the user asks about a specific WO.",
    input_schema: {
      type: "object",
      properties: {
        reference: { type: "string", description: "Work order reference number" },
      },
      required: ["reference"],
    },
  },
  {
    name: "search_requests",
    description: "Search maintenance requests (pre-WO tickets). Filters: property, requestor, status, priority.",
    input_schema: {
      type: "object",
      properties: {
        property: { type: "string" },
        requestor: { type: "string" },
        status: { type: "string", description: "e.g. 'FM Review', 'AM/DM Review', 'Draft', 'Closed'" },
        priority: { type: "string" },
        limit: { type: "number" },
      },
    },
  },
  {
    name: "get_vendor_summary",
    description: "Get a summary for a vendor: total WO count, total NTE, total spend, contact info, and the N most recent WOs. Use when the user asks 'tell me about vendor X' or wants vendor stats.",
    input_schema: {
      type: "object",
      properties: {
        vendor: { type: "string", description: "Vendor name (partial match OK)" },
        recent_count: { type: "number", description: "How many recent WOs to include (default 10)" },
      },
      required: ["vendor"],
    },
  },
  {
    name: "get_property_summary",
    description: "Get a summary for a property: WO counts by status and priority, total NTE, total spend, FM/DM, recent WOs and requests.",
    input_schema: {
      type: "object",
      properties: {
        property: { type: "string", description: "Property name (partial match OK)" },
        recent_count: { type: "number", description: "How many recent items to include (default 10)" },
      },
      required: ["property"],
    },
  },
  {
    name: "get_stats",
    description: "Get overall dashboard-style statistics: total WOs, breakdown by status, priority, region, facility manager, NTE exceeded count, total spend. Use for high-level 'how are we doing' questions.",
    input_schema: {
      type: "object",
      properties: {
        open_only: { type: "boolean", description: "Only count open WOs (not Closed/Invoiced/Cancelled/Completed). Default true." },
        facility_manager: { type: "string", description: "Scope stats to a single FM" },
        region: { type: "string", description: "Scope stats to a single region" },
      },
    },
  },
  {
    name: "get_fm_workload",
    description: "Return workload breakdown per Facility Manager: count of open WOs, by priority, total NTE, total spend.",
    input_schema: {
      type: "object",
      properties: {
        open_only: { type: "boolean", description: "Default true" },
      },
    },
  },
  {
    name: "get_status_info",
    description: "Look up the meaning, next action, and responsible party for a specific WO or WR status. Use this when a user asks what a status means, who handles it, or what to do about it.",
    input_schema: {
      type: "object",
      properties: {
        status: { type: "string", description: "The status name (e.g. 'Assigned', 'Work Finished', 'FM Review', 'Change Order Submitted'). Case-insensitive." },
      },
      required: ["status"],
    },
  },
  {
    name: "list_statuses_by_party",
    description: "Given a responsible party (e.g. 'Overseas Team', 'Facility Manager', 'District Manager', 'Area Manager', 'Axxerion Support', 'Vendor', 'DM', 'FM'), return every WR/WO status that party is responsible for. Use this whenever the user asks what a team is responsible for, handles, or owns. Also use it when counting live WOs in those statuses (combine with search_work_orders).",
    input_schema: {
      type: "object",
      properties: {
        party: { type: "string", description: "Partial match against the responsible party string in the status guide. E.g. 'overseas' matches 'Overseas Team'." },
      },
      required: ["party"],
    },
  },
];

// ── Helpers ──
const CLOSED_STATUSES = new Set(["Closed", "Invoiced", "Cancelled", "Completed"]);

function isOpen(status) { return !CLOSED_STATUSES.has(status); }

function iLike(haystack, needle) {
  if (!haystack || !needle) return !needle;
  return String(haystack).toLowerCase().includes(String(needle).toLowerCase());
}

function parseNum(x) {
  if (x == null) return 0;
  if (typeof x === "number") return x;
  const n = parseFloat(String(x).replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}

function parseDate(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Convert raw API WO to a lean summary object for JSON return
function summarizeWO(r) {
  return {
    reference: r["Reference"],
    property: r["Property"],
    status: r["Status"],
    priority: r["Priority"],
    problem_type: r["Problem Type"],
    subject: r["Subject"],
    vendor: r["Vendor"],
    region: r["Region"],
    area_manager: r["Area Manager"],
    facility_manager: r["Facility Manager"],
    nte: parseNum(r["Not To Exceed"]),
    spend: parseNum(r["Total"]),
    nte_exceeded: r["NTE Exceeded"] === "Yes",
    created: r["Created"],
    scheduled_from: r["Scheduled from"],
    actual_end: r["Actual end date"],
    closed: r["Closed"],
    vendor_phone: r["Phone"],
    vendor_email: r["Email"],
    invoice_number: r["Vendor Invoice #"],
    bookmark: r["Bookmark"],
  };
}

function summarizeReq(r) {
  return {
    reference: r["Reference"],
    property: r["Concerns"],
    status: r["Status"],
    priority: r["Priority"],
    subject: r["Subject"],
    requestor: r["Requestor"],
    request_date: r["Request date"],
    facility_manager: r["Facility Manager"],
    bookmark: r["Bookmark"],
  };
}

// ── Tool handlers ──
function toolSearchWorkOrders(input, getData) {
  const { workOrders } = getData();
  if (!workOrders || !workOrders.length) return { error: "Work order data not loaded yet" };
  const limit = Math.min(input.limit || MAX_RESULTS, MAX_RESULTS);
  const sinceMs = input.created_since_days ? Date.now() - input.created_since_days * 86400000 : null;

  const filtered = workOrders.filter(r => {
    if (input.property && !iLike(r["Property"], input.property)) return false;
    if (input.vendor && !iLike(r["Vendor"], input.vendor)) return false;
    if (input.status) {
      if (input.status.toLowerCase() === "open") { if (!isOpen(r["Status"])) return false; }
      else if (!iLike(r["Status"], input.status)) return false;
    }
    if (input.priority && !iLike(r["Priority"], input.priority)) return false;
    if (input.region && !iLike(r["Region"], input.region)) return false;
    if (input.dm && !iLike(r["Area Manager"], input.dm)) return false;
    if (input.facility_manager && !iLike(r["Facility Manager"], input.facility_manager)) return false;
    if (input.problem_type && !iLike(r["Problem Type"], input.problem_type)) return false;
    if (input.nte_exceeded === true && r["NTE Exceeded"] !== "Yes") return false;
    if (sinceMs) {
      const d = parseDate(r["Created"]);
      if (!d || d.getTime() < sinceMs) return false;
    }
    return true;
  });

  return {
    total_matches: filtered.length,
    showing: Math.min(filtered.length, limit),
    results: filtered.slice(0, limit).map(summarizeWO),
  };
}

function toolGetWODetails(input, getData) {
  const { workOrders } = getData();
  if (!workOrders || !workOrders.length) return { error: "Work order data not loaded yet" };
  const ref = (input.reference || "").toLowerCase().trim();
  if (!ref) return { error: "reference required" };
  const wo = workOrders.find(r => String(r["Reference"]||"").toLowerCase() === ref);
  if (!wo) return { error: "Work order not found: " + input.reference };
  // Return fuller detail for single WO
  return {
    ...summarizeWO(wo),
    assignee: wo["Assignee"],
    executor: wo["Executor"],
    address: wo["Address"],
    location: wo["Location"],
    floor: wo["Floor"],
    maintenance_schedule: wo["Maintenance schedule"],
    scheduled_until: wo["Scheduled until"],
    actual_start: wo["Actual start date"],
    secondary_contact: wo["Secondary Contact"],
    vendor_mobile: wo["Mobile"],
    vendor_estimated_cost: parseNum(wo["Vendor Estimated Cost"]),
    linked_request: wo["Request"],
  };
}

function toolSearchRequests(input, getData) {
  const { requests } = getData();
  if (!requests || !requests.length) return { error: "Request data not loaded yet" };
  const limit = Math.min(input.limit || MAX_RESULTS, MAX_RESULTS);
  const filtered = requests.filter(r => {
    if (input.property && !iLike(r["Concerns"], input.property)) return false;
    if (input.requestor && !iLike(r["Requestor"], input.requestor)) return false;
    if (input.status && !iLike(r["Status"], input.status)) return false;
    if (input.priority && !iLike(r["Priority"], input.priority)) return false;
    return true;
  });
  return {
    total_matches: filtered.length,
    showing: Math.min(filtered.length, limit),
    results: filtered.slice(0, limit).map(summarizeReq),
  };
}

function toolGetVendorSummary(input, getData) {
  const { workOrders } = getData();
  if (!workOrders) return { error: "Data not loaded" };
  if (!input.vendor) return { error: "vendor required" };
  const matching = workOrders.filter(r => iLike(r["Vendor"], input.vendor));
  if (!matching.length) return { vendor: input.vendor, total_wos: 0, message: "No work orders found for this vendor" };

  // Collect contact info from the most recent WO that has it
  let phone = "", mobile = "", email = "";
  const sorted = [...matching].sort((a,b) => {
    const da = parseDate(a["Created"]), db = parseDate(b["Created"]);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });
  for (const r of sorted) {
    if (!phone && r["Phone"]) phone = r["Phone"];
    if (!mobile && r["Mobile"]) mobile = r["Mobile"];
    if (!email && r["Email"]) email = r["Email"];
    if (phone && mobile && email) break;
  }

  const totalNTE = matching.reduce((s,r) => s + parseNum(r["Not To Exceed"]), 0);
  const totalSpend = matching.reduce((s,r) => s + parseNum(r["Total"]), 0);
  const openCount = matching.filter(r => isOpen(r["Status"])).length;
  const nteExceeded = matching.filter(r => r["NTE Exceeded"] === "Yes").length;
  const properties = [...new Set(matching.map(r => r["Property"]).filter(Boolean))];

  const recent = sorted.slice(0, input.recent_count || 10).map(summarizeWO);

  // Use the actual vendor name from the first match (more accurate than user input)
  const actualVendor = matching[0]["Vendor"];

  return {
    vendor: actualVendor,
    total_wos: matching.length,
    open_wos: openCount,
    nte_exceeded_count: nteExceeded,
    total_nte: Math.round(totalNTE * 100) / 100,
    total_spend: Math.round(totalSpend * 100) / 100,
    properties_served: properties.length,
    contact: { phone, mobile, email },
    recent_work_orders: recent,
  };
}

function toolGetPropertySummary(input, getData) {
  const { workOrders, requests } = getData();
  if (!workOrders) return { error: "Data not loaded" };
  if (!input.property) return { error: "property required" };

  const woMatching = workOrders.filter(r => iLike(r["Property"], input.property));
  const reqMatching = (requests || []).filter(r => iLike(r["Concerns"], input.property));
  if (!woMatching.length && !reqMatching.length) return { property: input.property, message: "No WOs or requests found for this property" };

  const statusBreakdown = {};
  const priorityBreakdown = {};
  woMatching.forEach(r => {
    statusBreakdown[r["Status"]] = (statusBreakdown[r["Status"]] || 0) + 1;
    priorityBreakdown[r["Priority"]] = (priorityBreakdown[r["Priority"]] || 0) + 1;
  });
  const openWOs = woMatching.filter(r => isOpen(r["Status"]));
  const totalNTE = woMatching.reduce((s,r) => s + parseNum(r["Not To Exceed"]), 0);
  const totalSpend = woMatching.reduce((s,r) => s + parseNum(r["Total"]), 0);

  const actualProperty = woMatching[0] ? woMatching[0]["Property"] : reqMatching[0]["Concerns"];
  const sample = woMatching[0] || {};

  const sortByCreated = (a,b) => {
    const da = parseDate(a["Created"] || a["Request date"]);
    const db = parseDate(b["Created"] || b["Request date"]);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  };
  const recentWOs = [...woMatching].sort(sortByCreated).slice(0, input.recent_count || 10).map(summarizeWO);
  const recentReqs = [...reqMatching].sort(sortByCreated).slice(0, input.recent_count || 10).map(summarizeReq);

  return {
    property: actualProperty,
    region: sample["Region"],
    area_manager: sample["Area Manager"],
    facility_manager: sample["Facility Manager"],
    total_work_orders: woMatching.length,
    open_work_orders: openWOs.length,
    total_requests: reqMatching.length,
    total_nte: Math.round(totalNTE * 100) / 100,
    total_spend: Math.round(totalSpend * 100) / 100,
    status_breakdown: statusBreakdown,
    priority_breakdown: priorityBreakdown,
    recent_work_orders: recentWOs,
    recent_requests: recentReqs,
  };
}

function toolGetStats(input, getData) {
  const { workOrders } = getData();
  if (!workOrders || !workOrders.length) return { error: "Data not loaded" };
  const openOnly = input.open_only !== false; // default true

  let scoped = workOrders;
  if (input.facility_manager) scoped = scoped.filter(r => iLike(r["Facility Manager"], input.facility_manager));
  if (input.region) scoped = scoped.filter(r => iLike(r["Region"], input.region));
  if (openOnly) scoped = scoped.filter(r => isOpen(r["Status"]));

  const byStatus = {}, byPriority = {}, byRegion = {}, byFM = {};
  let totalNTE = 0, totalSpend = 0, nteExceeded = 0;
  scoped.forEach(r => {
    byStatus[r["Status"]] = (byStatus[r["Status"]] || 0) + 1;
    byPriority[r["Priority"]] = (byPriority[r["Priority"]] || 0) + 1;
    if (r["Region"]) byRegion[r["Region"]] = (byRegion[r["Region"]] || 0) + 1;
    if (r["Facility Manager"]) byFM[r["Facility Manager"]] = (byFM[r["Facility Manager"]] || 0) + 1;
    totalNTE += parseNum(r["Not To Exceed"]);
    totalSpend += parseNum(r["Total"]);
    if (r["NTE Exceeded"] === "Yes") nteExceeded++;
  });

  return {
    scope: { open_only: openOnly, facility_manager: input.facility_manager || "all", region: input.region || "all" },
    total_count: scoped.length,
    total_nte: Math.round(totalNTE * 100) / 100,
    total_spend: Math.round(totalSpend * 100) / 100,
    nte_exceeded_count: nteExceeded,
    by_status: byStatus,
    by_priority: byPriority,
    by_region: byRegion,
    by_facility_manager: byFM,
  };
}

function toolGetFMWorkload(input, getData) {
  const { workOrders } = getData();
  if (!workOrders || !workOrders.length) return { error: "Data not loaded" };
  const openOnly = input.open_only !== false;
  const scoped = openOnly ? workOrders.filter(r => isOpen(r["Status"])) : workOrders;

  const fms = {};
  scoped.forEach(r => {
    const fm = r["Facility Manager"] || "(Unassigned)";
    if (!fms[fm]) fms[fm] = { total: 0, by_priority: {}, total_nte: 0, total_spend: 0 };
    fms[fm].total++;
    fms[fm].by_priority[r["Priority"]] = (fms[fm].by_priority[r["Priority"]] || 0) + 1;
    fms[fm].total_nte += parseNum(r["Not To Exceed"]);
    fms[fm].total_spend += parseNum(r["Total"]);
  });
  Object.values(fms).forEach(v => {
    v.total_nte = Math.round(v.total_nte * 100) / 100;
    v.total_spend = Math.round(v.total_spend * 100) / 100;
  });
  return { scope: { open_only: openOnly }, workload: fms };
}

function toolGetStatusInfo(input) {
  const info = getStatusInfo(input.status);
  if (!info) return { error: "Status not found in guide: " + input.status + ". Known statuses include: " + [...STATUS_GUIDE.wr.map(s=>s.status), ...STATUS_GUIDE.wo.map(s=>s.status)].join(", ") };
  return info;
}

function toolListStatusesByParty(input) {
  const matches = listStatusesByParty(input.party);
  if (!matches.length) return { error: "No statuses found for party matching: " + input.party, hint: "Try 'Overseas Team', 'Facility Manager', 'District Manager', 'Area Manager', 'Axxerion Support'." };
  return {
    party_query: input.party,
    count: matches.length,
    statuses: matches.map(s => ({ status: s.status, type: s.type, party: s.party, meaning: s.meaning, action: s.action, phase: s.phase || null })),
  };
}

function executeTool(name, input, getData) {
  try {
    switch (name) {
      case "search_work_orders": return toolSearchWorkOrders(input || {}, getData);
      case "get_work_order_details": return toolGetWODetails(input || {}, getData);
      case "search_requests": return toolSearchRequests(input || {}, getData);
      case "get_vendor_summary": return toolGetVendorSummary(input || {}, getData);
      case "get_property_summary": return toolGetPropertySummary(input || {}, getData);
      case "get_stats": return toolGetStats(input || {}, getData);
      case "get_fm_workload": return toolGetFMWorkload(input || {}, getData);
      case "get_status_info": return toolGetStatusInfo(input || {});
      case "list_statuses_by_party": return toolListStatusesByParty(input || {});
      default: return { error: "Unknown tool: " + name };
    }
  } catch (e) {
    console.error("[Chat] Tool error (" + name + "):", e.message);
    return { error: "Tool execution failed: " + e.message };
  }
}

module.exports = { TOOL_DEFINITIONS, executeTool };
