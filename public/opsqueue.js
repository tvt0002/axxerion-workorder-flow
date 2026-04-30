/* ── Ops Queue Logic ── */
var OPS_DATA = { logs: {}, appointments: {}, emails: {}, vendors: {}, notes: {}, dismissed: {} };
var OPS_QUEUE = 'q1';
var OPS_LOADED = false;
var OQ_SORT = { col: null, dir: 1 }; // current sort state

/* ── Store directory lookup (from /call-directory/data/stores.json) ── */
var STORE_LOOKUP = null;
function loadStoreDirectory() {
  fetch('/call-directory/data/stores.json')
    .then(function(r){ return r.json(); })
    .then(function(stores){
      var map = {};
      stores.forEach(function(s){
        if (s.name) map[s.name.toLowerCase().trim()] = s;
        if (s.code) map[s.code.toLowerCase().trim()] = s;
      });
      STORE_LOOKUP = map;
      if (OPS_LOADED && typeof renderQueue === 'function') renderQueue();
    })
    .catch(function(e){ console.error('[Store] Load error:', e); STORE_LOOKUP = {}; });
}
function getStoreInfo(property) {
  if (!STORE_LOOKUP || !property) return null;
  var p = String(property).toLowerCase().trim();
  if (STORE_LOOKUP[p]) return STORE_LOOKUP[p];
  // Partial match: property contains store name (e.g. "Titusville FL" matches "titusville")
  for (var k in STORE_LOOKUP) {
    if (k.length < 4) continue;
    if (p.indexOf(k) !== -1) return STORE_LOOKUP[k];
  }
  return null;
}

function oqSortVal(item, col) {
  var v = item[col];
  if (v === undefined || v === null || v === '') return null;
  if (typeof v === 'number') return v;
  // Try parsing as number
  var n = parseFloat(String(v).replace(/[$,]/g, ''));
  if (!isNaN(n)) return n;
  return String(v).toLowerCase();
}

function oqSortItems(items, col, dir) {
  if (!col) return items;
  return items.slice().sort(function(a, b) {
    var va = oqSortVal(a, col), vb = oqSortVal(b, col);
    if (va === null && vb === null) return 0;
    if (va === null) return 1;
    if (vb === null) return -1;
    if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
    return va < vb ? -dir : va > vb ? dir : 0;
  });
}

function oqSortHeader(label, key, isRight) {
  var arrow = OQ_SORT.col === key ? (OQ_SORT.dir === 1 ? ' ▲' : ' ▼') : '';
  var cls = isRight ? ' class="r"' : '';
  return '<th' + cls + ' style="cursor:pointer;user-select:none" data-sortkey="' + key + '">' + label + arrow + '</th>';
}

function bindOqSort() {
  var head = document.getElementById('oqHead');
  if (!head) return;
  var resetBtn = document.getElementById('oqSortReset');
  if (resetBtn) resetBtn.style.display = OQ_SORT.col ? '' : 'none';
  head.querySelectorAll('th[data-sortkey]').forEach(function(th) {
    th.addEventListener('click', function() {
      var key = th.getAttribute('data-sortkey');
      if (OQ_SORT.col === key) { OQ_SORT.dir *= -1; }
      else { OQ_SORT.col = key; OQ_SORT.dir = 1; }
      renderQueue();
    });
  });
}

// Queue status definitions (aligned with Axxerion WR-WO Statuses guide)
function isInfoNeeded(val) {
  if (!val) return false;
  var v = val.toUpperCase();
  return v === 'INFO NEEDED' || v.indexOf('INFO NEEDED') === 0 || v === 'NFO NEEDED' || v.indexOf('NFO NEEDED') === 0;
}
var Q2_STATUSES = new Set(['Assigned', 'Accepted']);
var Q4_STATUSES = new Set(['Work finished', 'Work Finished']);
var TERMINAL_STATUSES = new Set(['Closed', 'Cancelled', 'Invoiced', 'Completed']);

function loadOpsData(cb) {
  fetch('/api/ops').then(function(r) { return r.json(); }).then(function(d) {
    OPS_DATA = d;
    OPS_LOADED = true;
    if (cb) cb();
  }).catch(function(e) { console.error('[Ops] Load error:', e); if (cb) cb(); });
}

function opsApi(endpoint, body, cb) {
  fetch('/api/ops/' + endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }).then(function(r) { return r.json(); }).then(function(d) {
    loadOpsData(function() { renderQueue(); });
    if (cb) cb(d);
  }).catch(function(e) { console.error('[Ops] API error:', e); });
}

function hoursAgo(dateStr) {
  if (!dateStr) return 9999;
  // Parse "M/D/YY h:mm AM/PM" format from Axxerion
  var d = new Date(dateStr);
  if (isNaN(d.getTime())) return 9999;
  return Math.floor((Date.now() - d.getTime()) / 3600000);
}

function daysAgo(dateStr) {
  return Math.floor(hoursAgo(dateStr) / 24);
}

function todayStr() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function fmtDate(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function fmtShortDate(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

// ── Queue Data Builders ──

function getQ1Data() {
  // Requests needing more info
  if (!ALLDATA || !ALLDATA.length) return [];
  return ALLDATA.filter(function(r) {
    if (r[15] !== 'Request') return false;
    if (TERMINAL_STATUSES.has(r[1])) return false;
    if (!isInfoNeeded(r[2]) && !isInfoNeeded(r[1])) return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q1) return false;
    return true;
  }).map(function(r) {
    var ref = r[10] || '';
    var logs = OPS_DATA.logs[ref] || [];
    var note = OPS_DATA.notes[ref] || {};
    var store = getStoreInfo(r[0]);
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', subject: r[3], created: r[12], requestor: r[16] || r[8], bookmark: r[14], logs: logs, note: note.text || '', age: daysAgo(r[12]), lastAction: logs.length ? logs[0] : null, storeCode: store ? (store.code || '') : '', storePhone: store ? (store.store_office_phone || '') : '', storeCellPhone: store ? (store.store_cell_phone || '') : '' };
  }).sort(function(a, b) { return a.age - b.age; }); // newest first
}

function getQ2Data() {
  // WOs assigned 24+ hours, not yet scheduled, no execution dates yet
  if (!ALLDATA || !ALLDATA.length) return [];
  return ALLDATA.filter(function(r) {
    if (r[15] === 'Request') return false;
    if (!Q2_STATUSES.has(r[1])) return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q2) return false;
    // Out of Q2 if ANY scheduling/execution date has been touched. Vendors don't always
    // populate fields in order — checking only Scheduled from leaks the ones that filled
    // Scheduled until, Actual start, or Actual end first.
    if (r[19] || r[20] || r[21] || r[22]) return false;
    if (OPS_DATA.appointments[ref] && OPS_DATA.appointments[ref].date) return false;
    var hrs = hoursAgo(r[12]);
    return hrs >= 24;
  }).map(function(r) {
    var ref = r[10] || '';
    var logs = OPS_DATA.logs[ref] || [];
    var vendor = r[6] || '';
    var vendorInfo = OPS_DATA.vendors[vendor] || {};
    var appt = OPS_DATA.appointments[ref] || {};
    var vPhone = r[27] || r[28] || vendorInfo.phone || '';
    var vEmail = r[29] || vendorInfo.email || '';
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', service: r[3], vendor: vendor, executor: r[18] || '', vendorPhone: vPhone, vendorEmail: vEmail, created: r[12], bookmark: r[14], logs: logs, hours: hoursAgo(r[12]), callCount: logs.filter(function(l) { return l.action === 'call'; }).length, lastAction: logs.length ? logs[0] : null, schedFrom: r[19] || '', schedUntil: r[20] || '', actualStart: r[21] || '', actualEnd: r[22] || '', apptDate: appt.date || '', apptTime: appt.time || '' };
  }).sort(function(a, b) { return a.hours - b.hours; }); // newest first
}

// Statuses that mean Q3 should skip the WO regardless of date fields.
// Covers terminal states (Closed/Cancelled/Invoiced/Completed) plus statuses that
// belong to Q4 or downstream (Work finished, Financials Submitted, etc.) — guards
// against dirty data where actualEnd was never set on retired WOs.
var Q3_EXCLUDED_STATUSES = new Set([
  'Closed', 'Cancelled', 'Invoiced', 'Completed',
  'Work finished', 'Work Finished',
  'Financials Submitted', 'No Invoice/Warranty/Internal',
]);

function getQ3Data() {
  // Active-work tracker: today's appointments + any WIP work the vendor started
  // but hasn't reported finished yet. Drops to Q4 the moment Actual end date populates.
  if (!ALLDATA || !ALLDATA.length) return [];
  var today = todayStr();
  return ALLDATA.filter(function(r) {
    if (r[15] === 'Request') return false;
    if (Q3_EXCLUDED_STATUSES.has(r[1])) return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q3) return false;
    // Once Actual end date is populated the WO is done — Q4 picks it up.
    if (r[22]) return false;
    // WIP: vendor has started work but not finished — keep in Q3 until verified done.
    if (r[21]) return true;
    // Today's scheduled appointment (Axxerion Scheduled from)
    if (r[19]) {
      var sf = new Date(r[19]);
      if (!isNaN(sf.getTime())) {
        var sfDate = sf.getFullYear() + '-' + String(sf.getMonth() + 1).padStart(2, '0') + '-' + String(sf.getDate()).padStart(2, '0');
        if (sfDate === today) return true;
      }
    }
    // Fall back to manual ops appointment
    var appt = OPS_DATA.appointments[ref];
    if (appt && appt.date === today) return true;
    return false;
  }).map(function(r) {
    var ref = r[10] || '';
    var appt = OPS_DATA.appointments[ref] || {};
    var logs = OPS_DATA.logs[ref] || [];
    // Derive time from Axxerion Scheduled from if available
    var time = appt.time || '';
    if (!time && r[19]) {
      var sf = new Date(r[19]);
      if (!isNaN(sf.getTime())) time = sf.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
    }
    var vendorInfo = OPS_DATA.vendors[r[6] || ''] || {};
    var vEmail = r[29] || vendorInfo.email || '';
    var isWip = !!r[21];
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', service: r[3], vendor: r[6] || '', executor: r[18] || '', time: time, confirmed: appt.confirmed, bookmark: r[14], logs: logs, lastAction: logs.length ? logs[0] : null, vendorEmail: vEmail, schedFrom: r[19] || '', actualStart: r[21] || '', actualEnd: r[22] || '', wip: isWip };
  }).sort(function(a, b) {
    // WIP first (active jobs needing verification), then today's appointments.
    // WIP sub-sort: most recently started first. Today sub-sort: chronological.
    if (a.wip !== b.wip) return a.wip ? -1 : 1;
    if (a.wip && b.wip) {
      var ta = new Date(a.actualStart).getTime() || 0;
      var tb = new Date(b.actualStart).getTime() || 0;
      return tb - ta; // newest started first
    }
    return (a.time || 'ZZ').localeCompare(b.time || 'ZZ');
  });
}

// Status that means invoice already settled, vendor already submitted, or WO retired —
// never surface in Q4. Q4 is a "missing invoice — chase the vendor" list, so any status
// that already has the invoice handled (or explicitly says no invoice is coming) is out.
var Q4_TERMINAL_STATUSES = new Set([
  'Closed', 'Cancelled', 'Invoiced', 'Completed',
  'Financials Submitted',           // vendor already submitted invoice
  'No Invoice/Warranty/Internal',   // no invoice expected
]);

function getQ4Data() {
  // WOs with work done, awaiting invoice. Match if EITHER:
  //   - status is Work finished, OR
  //   - Actual end date is populated (vendor finished but didn't move status)
  // Filter out terminal statuses so Closed/Invoiced WOs with end dates don't leak in.
  if (!ALLDATA || !ALLDATA.length) return [];
  return ALLDATA.filter(function(r) {
    if (r[15] !== 'Work Order') return false;
    var ref = r[10] || '';
    if (!ref) return false;
    if (Q4_TERMINAL_STATUSES.has(r[1])) return false;
    var hasActualEnd = !!r[22];
    if (!Q4_STATUSES.has(r[1]) && !hasActualEnd) return false;
    // Skip rows missing key data (property, vendor, or service type)
    if (!r[0] && !r[6] && !r[3]) return false;
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q4) return false;
    return true;
  }).map(function(r) {
    var ref = r[10] || '';
    var emails = OPS_DATA.emails[ref] || [];
    var logs = OPS_DATA.logs[ref] || [];
    var vendor = r[6] || '';
    var vendorInfo = OPS_DATA.vendors[vendor] || {};
    var finishedDate = r[22] || r[23] || r[12]; // Actual end, Closed, or Created as fallback
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', service: r[3], vendor: vendor, vendorEmail: vendorInfo.email || '', vendorEstCost: r[31] || 0, bookmark: r[14], logs: logs, emails: emails, emailCount: emails.length, lastEmail: emails.length ? emails[0] : null, lastAction: logs.length ? logs[0] : null, finishedDate: finishedDate, daysSinceFinished: daysAgo(finishedDate) };
  }).sort(function(a, b) {
    // Most recently finished first — fresh closeouts are easier to chase
    // (vendor still has the job in mind). Old ones drift to bottom but
    // remain accessible via column sort.
    return a.daysSinceFinished - b.daysSinceFinished;
  });
}

function getQ5Data() {
  // All sent emails from OPS_DATA.emails, joined with WO data
  var results = [];
  var refs = Object.keys(OPS_DATA.emails || {});
  refs.forEach(function(ref) {
    var emails = OPS_DATA.emails[ref] || [];
    var wo = ALLDATA ? ALLDATA.find(function(r) { return r[10] === ref; }) : null;
    emails.forEach(function(em) {
      results.push({
        ref: ref,
        property: wo ? wo[0] : '',
        vendor: wo ? (wo[6] || '') : '',
        service: wo ? (wo[3] || '') : '',
        status: wo ? (wo[1] || '') : '',
        bookmark: wo ? (wo[14] || '') : '',
        to: em.to || '',
        subject: em.subject || '',
        type: em.type || 'invoice',
        sentAt: em.sentAt || '',
        sentAtFmt: fmtDate(em.sentAt)
      });
    });
  });
  // Sort newest first
  results.sort(function(a, b) {
    return new Date(b.sentAt) - new Date(a.sentAt);
  });
  return results;
}

// Priority color helper
function priColor(p) {
  if (!p) return 'color:var(--muted)';
  var l = p.toLowerCase();
  if (l === 'critical') return 'color:var(--red);font-weight:600';
  if (l === 'high') return 'color:var(--orange);font-weight:600';
  if (l === 'medium') return 'color:var(--yellow)';
  if (l === 'low') return 'color:var(--green)';
  return 'color:var(--muted)';
}

// ── Queue Rendering ──

function switchQueue(q, btn) {
  OPS_QUEUE = q;
  OQ_SORT = { col: null, dir: 1 };
  localStorage.setItem('ax_active_queue', q);
  window.location.hash = 'opsqueue/' + q;
  document.querySelectorAll('.oq-tab').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  document.querySelectorAll('.oq-kpi').forEach(function(k) {
    if (k.dataset.queue === q) {
      k.style.opacity = '1';
      k.style.transform = 'scale(1.03)';
      k.style.boxShadow = '0 0 0 2px rgba(var(--accent-rgb),.5)';
      k.style.transition = 'all .2s ease';
    } else {
      k.style.opacity = '0.4';
      k.style.transform = 'scale(1)';
      k.style.boxShadow = '';
      k.style.transition = 'all .2s ease';
    }
  });
  renderQueue();
}

function renderQueueKPIs() {
  var q1 = getQ1Data(), q2 = getQ2Data(), q3 = getQ3Data(), q4 = getQ4Data(), q5 = getQ5Data();
  document.getElementById('oqK1').textContent = q1.length;
  document.getElementById('oqK2').textContent = q2.length;
  document.getElementById('oqK3').textContent = q3.length;
  document.getElementById('oqK4').textContent = q4.length;
  document.getElementById('oqK5').textContent = q5.length;
  // Update tab counts
  document.querySelectorAll('.oq-tab').forEach(function(t) {
    var q = t.getAttribute('data-queue');
    var c = q === 'q1' ? q1.length : q === 'q2' ? q2.length : q === 'q3' ? q3.length : q === 'q4' ? q4.length : q5.length;
    var lbl = t.textContent.replace(/\s*\(\d+\)$/, '');
    t.textContent = lbl + ' (' + c + ')';
  });
}

function renderQueue() {
  renderQueueKPIs();
  var head = document.getElementById('oqHead');
  var body = document.getElementById('oqBody');
  var empty = document.getElementById('oqEmpty');
  var items, headHTML, rowsHTML;

  if (OPS_QUEUE === 'q1') {
    items = getQ1Data();
    headHTML = '<tr>' + oqSortHeader('Age','age') + oqSortHeader('Reference','ref') + oqSortHeader('Property','property') + oqSortHeader('Store #','storeCode') + oqSortHeader('Store Phone','storePhone') + oqSortHeader('Status','status') + oqSortHeader('Priority','priority') + oqSortHeader('Subject','subject') + oqSortHeader('Requestor','requestor') + '<th>Last Action</th><th>Actions</th></tr>';
    items = oqSortItems(items, OQ_SORT.col, OQ_SORT.dir);
    rowsHTML = items.map(function(d) {
      var ageClass = d.age >= 3 ? 'color:var(--red)' : d.age >= 1 ? 'color:var(--orange)' : 'color:var(--green)';
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastAct = d.lastAction ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastAction.date) + ' · ' + d.lastAction.action + '</span>' : '<span style="font-size:10px;color:var(--red)">No action yet</span>';
      var codeCell = d.storeCode ? '<span style="font-family:\'DM Mono\',monospace;font-size:11px;color:var(--muted)">' + d.storeCode + '</span>' : '<span style="color:var(--muted);font-size:10px">—</span>';
      var phoneCell;
      if (d.storePhone) {
        var tel = d.storePhone.replace(/\D/g,'');
        var cellTitle = d.storeCellPhone ? ' title="Cell: ' + d.storeCellPhone + '"' : '';
        phoneCell = '<a href="tel:' + tel + '"' + cellTitle + ' style="color:var(--accent);font-family:\'DM Mono\',monospace;font-size:11px;text-decoration:none">' + d.storePhone + '</a>';
      } else {
        phoneCell = '<span style="color:var(--muted);font-size:10px">—</span>';
      }
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + ageClass + '">' + d.age + 'd</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + codeCell + '</td>'
        + '<td>' + phoneCell + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.subject + '</td>'
        + '<td>' + d.requestor + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q1\')">Log Action</button> <button class="oq-btn" title="History" onclick="openWOHistory(\'' + d.ref + '\',{property:\'' + (d.property||'').replace(/'/g,"\\'") + '\',bookmark:\'' + (d.bookmark||'').replace(/'/g,"\\'") + '\'})">🕐</button> <button class="oq-btn oq-btn-dim" onclick="dismissOq(\'' + d.ref + '\',\'q1\')">Dismiss</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q2') {
    items = getQ2Data();
    headHTML = '<tr>' + oqSortHeader('Hours','hours') + oqSortHeader('Reference','ref') + oqSortHeader('Property','property') + oqSortHeader('Status','status') + oqSortHeader('Priority','priority') + oqSortHeader('Service Type','service') + oqSortHeader('Vendor','vendor') + oqSortHeader('Phone','vendorPhone') + oqSortHeader('Email','vendorEmail') + oqSortHeader('Sched Start','schedFrom') + oqSortHeader('Sched End','schedUntil') + oqSortHeader('Actual Start','actualStart') + oqSortHeader('Actual End','actualEnd') + oqSortHeader('Appointment','apptDate') + oqSortHeader('Calls','callCount') + '<th>Last Action</th><th>Actions</th></tr>';
    items = oqSortItems(items, OQ_SORT.col, OQ_SORT.dir);
    rowsHTML = items.map(function(d) {
      var hrsClass = d.hours >= 72 ? 'color:var(--red)' : d.hours >= 48 ? 'color:var(--orange)' : 'color:var(--yellow)';
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var phone = d.vendorPhone ? '<a href="tel:' + d.vendorPhone + '" style="color:var(--accent)">' + d.vendorPhone + '</a>' : '<span style="color:var(--red);font-size:10px">No phone</span>';
      var lastAct = d.lastAction ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastAction.date) + ' · ' + d.lastAction.action + '</span>' : '<span style="font-size:10px;color:var(--red)">No calls</span>';
      var dateSty = 'font-family:\'DM Mono\',monospace;font-size:10px;white-space:nowrap';
      var apptTxt = d.apptDate ? d.apptDate + (d.apptTime ? ' ' + d.apptTime : '') : '';
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + hrsClass + '">' + d.hours + 'h</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + phone + '</td>'
        + '<td>' + (d.vendorEmail ? '<a href="mailto:' + d.vendorEmail + '" style="color:var(--accent);font-size:11px">' + d.vendorEmail + '</a>' : '<span style="color:var(--muted);font-size:10px">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.schedFrom || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.schedUntil || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualStart || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualEnd || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (apptTxt || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;text-align:center">' + d.callCount + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q2\')">Log Call</button> <button class="oq-btn oq-btn-g" onclick="openOqSchedule(\'' + d.ref + '\')">Schedule</button> <button class="oq-btn" title="History" onclick="openWOHistory(\'' + d.ref + '\',{property:\'' + (d.property||'').replace(/'/g,"\\'") + '\',bookmark:\'' + (d.bookmark||'').replace(/'/g,"\\'") + '\'})">🕐</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q3') {
    items = getQ3Data();
    headHTML = '<tr>' + oqSortHeader('Time','time') + oqSortHeader('Reference','ref') + oqSortHeader('Property','property') + oqSortHeader('Status','status') + oqSortHeader('Priority','priority') + oqSortHeader('Vendor','vendor') + oqSortHeader('Vendor Email','vendorEmail') + oqSortHeader('Service Type','service') + oqSortHeader('Sched Start','schedFrom') + oqSortHeader('Actual Start','actualStart') + oqSortHeader('Actual End','actualEnd') + oqSortHeader('Confirmed','confirmed') + '<th>Last Action</th><th>Actions</th></tr>';
    items = oqSortItems(items, OQ_SORT.col, OQ_SORT.dir);
    rowsHTML = items.map(function(d) {
      var confBadge = d.wip
        ? '<span style="color:var(--accent);font-weight:600">In Progress</span>'
        : (d.confirmed ? '<span style="color:var(--green);font-weight:600">Confirmed</span>' : '<span style="color:var(--orange);font-weight:600">Pending</span>');
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastAct = d.lastAction ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastAction.date) + ' · ' + d.lastAction.action + '</span>' : '<span style="font-size:10px;color:var(--muted)">—</span>';
      var dateSty = 'font-family:\'DM Mono\',monospace;font-size:11px';
      var muted = '<span style="color:var(--muted)">—</span>';
      var timeCell = d.wip ? 'WIP' : (d.time || 'TBD');
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:12px;font-weight:600">' + timeCell + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td>' + (d.vendorEmail ? '<a href="mailto:' + d.vendorEmail + '" style="color:var(--accent);font-size:11px">' + d.vendorEmail + '</a>' : '<span style="color:var(--muted);font-size:10px">—</span>') + '</td>'
        + '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td style="' + dateSty + '">' + (d.schedFrom || muted) + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualStart || muted) + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualEnd || muted) + '</td>'
        + '<td>' + confBadge + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn oq-btn-g" onclick="confirmAppt(\'' + d.ref + '\')">Confirm</button> <button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q3\')">Log</button> <button class="oq-btn oq-btn-r" onclick="openOqAction(\'' + d.ref + '\',\'q3\',\'noshow\')">No Show</button> <button class="oq-btn" title="History" onclick="openWOHistory(\'' + d.ref + '\',{property:\'' + (d.property||'').replace(/'/g,"\\'") + '\',bookmark:\'' + (d.bookmark||'').replace(/'/g,"\\'") + '\'})">🕐</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q4') {
    items = getQ4Data();
    headHTML = '<tr>' + oqSortHeader('Days Since End','daysSinceFinished') + oqSortHeader('Date Finished','finishedDate') + oqSortHeader('Reference','ref') + oqSortHeader('Property','property') + oqSortHeader('Vendor','vendor') + oqSortHeader('Service Type','service') + oqSortHeader('Status','status') + oqSortHeader('Priority','priority') + oqSortHeader('Vendor Est Cost','vendorEstCost',true) + oqSortHeader('Emails Sent','emailCount') + '<th>Last Email</th><th>Actions</th></tr>';
    items = oqSortItems(items, OQ_SORT.col, OQ_SORT.dir);
    rowsHTML = items.map(function(d) {
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastEmail = d.lastEmail ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastEmail.sentAt) + '</span>' : '<span style="font-size:10px;color:var(--red)">Never sent</span>';
      var emailBtn = '<button class="oq-btn oq-btn-g" onclick="sendInvoiceEmail(\'' + d.ref + '\',\'' + d.vendor.replace(/'/g, "\\'") + '\')">Draft Email</button>';
      var daysClass = d.daysSinceFinished >= 14 ? 'color:var(--red)' : d.daysSinceFinished >= 7 ? 'color:var(--orange)' : 'color:var(--green)';
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + daysClass + '">' + d.daysSinceFinished + 'd</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:10px;white-space:nowrap">' + (d.finishedDate ? fmtShortDate(d.finishedDate) : '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td class="r" style="font-family:\'DM Mono\',monospace;font-size:11px">' + (d.vendorEstCost ? '$' + Number(d.vendorEstCost).toLocaleString() : '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;text-align:center">' + d.emailCount + '</td>'
        + '<td>' + lastEmail + '</td>'
        + '<td>' + emailBtn + ' <button class="oq-btn" title="History" onclick="openWOHistory(\'' + d.ref + '\',{property:\'' + (d.property||'').replace(/'/g,"\\'") + '\',bookmark:\'' + (d.bookmark||'').replace(/'/g,"\\'") + '\'})">🕐</button> <button class="oq-btn oq-btn-dim" onclick="dismissOq(\'' + d.ref + '\',\'q4\')">Dismiss</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q5') {
    items = getQ5Data();
    headHTML = '<tr>' + oqSortHeader('Sent','sentAtFmt') + oqSortHeader('Reference','ref') + oqSortHeader('Property','property') + oqSortHeader('Vendor','vendor') + oqSortHeader('Service Type','service') + oqSortHeader('To','to') + oqSortHeader('Subject','subject') + '<th>Actions</th></tr>';
    items = oqSortItems(items, OQ_SORT.col, OQ_SORT.dir);
    rowsHTML = items.map(function(d) {
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:10px;white-space:nowrap">' + d.sentAtFmt + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + d.to + '</td>'
        + '<td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px">' + d.subject + '</td>'
        + '<td><button class="oq-btn oq-btn-g" onclick="sendInvoiceEmail(\'' + d.ref + '\',\'' + d.vendor.replace(/'/g, "\\'") + '\')">Resend</button> <button class="oq-btn" title="History" onclick="openWOHistory(\'' + d.ref + '\',{property:\'' + (d.property||'').replace(/'/g,"\\'") + '\',bookmark:\'' + (d.bookmark||'').replace(/'/g,"\\'") + '\'})">🕐</button></td>'
        + '</tr>';
    }).join('');
  }

  head.innerHTML = headHTML;
  body.innerHTML = rowsHTML;
  empty.style.display = items.length ? 'none' : 'block';
  document.getElementById('oqTable').style.display = items.length ? '' : 'none';
  bindOqSort();
}

// ── Actions ──

function openOqModal(html) {
  document.getElementById('oqModalContent').innerHTML = html;
  document.getElementById('oqModal').style.display = 'flex';
}
function closeOqModal() {
  document.getElementById('oqModal').style.display = 'none';
}

function openOqAction(ref, queue, preset) {
  var logs = OPS_DATA.logs[ref] || [];
  var note = (OPS_DATA.notes[ref] || {}).text || '';
  var actions = queue === 'q1' ? ['outreach', 'info_received', 'escalated', 'note']
    : queue === 'q2' ? ['call', 'voicemail', 'scheduled', 'no_answer', 'note']
    : queue === 'q3' ? ['vendor_confirmed', 'property_notified', 'noshow', 'rescheduled', 'note']
    : ['email_sent', 'invoice_received', 'follow_up', 'note'];
  if (preset && actions.indexOf(preset) === -1) actions.unshift(preset);

  var logHTML = logs.slice(0, 10).map(function(l) {
    return '<div style="font-size:11px;padding:6px 0;border-bottom:1px solid var(--border)"><span style="color:var(--accent);font-family:\'DM Mono\',monospace;font-size:9px">' + fmtDate(l.date) + '</span> <span style="font-weight:600">' + l.action + '</span>' + (l.note ? ' — ' + l.note : '') + '</div>';
  }).join('');

  var html = '<div class="psl" style="margin-bottom:12px">LOG ACTION — ' + ref + '</div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">ACTION</label>'
    + '<select id="oqActType" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%">'
    + actions.map(function(a) { return '<option value="' + a + '"' + (a === preset ? ' selected' : '') + '>' + a.replace(/_/g, ' ') + '</option>'; }).join('')
    + '</select></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">NOTE</label>'
    + '<textarea id="oqActNote" rows="3" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:8px 10px;width:100%;resize:vertical;" placeholder="Optional note..."></textarea></div>'
    + '<button class="oq-btn oq-btn-g" style="width:100%;padding:10px" onclick="submitOqAction(\'' + ref + '\')">Save</button>'
    + (logHTML ? '<div style="margin-top:16px"><div class="psl">Recent Activity</div>' + logHTML + '</div>' : '');

  openOqModal(html);
}

function submitOqAction(ref) {
  var action = document.getElementById('oqActType').value;
  var note = document.getElementById('oqActNote').value;
  opsApi('log', { ref: ref, action: action, note: note });
  closeOqModal();
}

function openOqSchedule(ref) {
  var appt = OPS_DATA.appointments[ref] || {};
  var html = '<div class="psl" style="margin-bottom:12px">SCHEDULE APPOINTMENT — ' + ref + '</div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">DATE</label>'
    + '<input type="date" id="oqApptDate" value="' + (appt.date || '') + '" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">TIME (optional)</label>'
    + '<input type="text" id="oqApptTime" value="' + (appt.time || '') + '" placeholder="e.g. 10:00 AM" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<button class="oq-btn oq-btn-g" style="width:100%;padding:10px" onclick="submitOqSchedule(\'' + ref + '\')">Save Appointment</button>';
  openOqModal(html);
}

function submitOqSchedule(ref) {
  var date = document.getElementById('oqApptDate').value;
  var time = document.getElementById('oqApptTime').value;
  if (!date) { alert('Please select a date'); return; }
  opsApi('appointment', { ref: ref, date: date, time: time, confirmed: false });
  opsApi('log', { ref: ref, action: 'scheduled', note: 'Appointment set for ' + date + (time ? ' at ' + time : '') });
  closeOqModal();
}

function confirmAppt(ref) {
  var appt = OPS_DATA.appointments[ref] || {};
  opsApi('appointment', { ref: ref, date: appt.date, time: appt.time, confirmed: true });
  opsApi('log', { ref: ref, action: 'vendor_confirmed', note: 'Day-of confirmation' });
}

function dismissOq(ref, queue) {
  opsApi('dismiss', { ref: ref, queue: queue });
}

function openVendorEdit(vendorName) {
  var info = OPS_DATA.vendors[vendorName] || {};
  var html = '<div class="psl" style="margin-bottom:12px">VENDOR CONTACT — ' + vendorName + '</div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">EMAIL</label>'
    + '<input type="email" id="oqVEmail" value="' + (info.email || '') + '" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">PHONE</label>'
    + '<input type="text" id="oqVPhone" value="' + (info.phone || '') + '" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">CONTACT NAME</label>'
    + '<input type="text" id="oqVContact" value="' + (info.contact || '') + '" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<button class="oq-btn oq-btn-g" style="width:100%;padding:10px" onclick="submitVendor(\'' + vendorName.replace(/'/g, "\\'") + '\')">Save</button>';
  openOqModal(html);
}

function submitVendor(name) {
  var email = document.getElementById('oqVEmail').value;
  var phone = document.getElementById('oqVPhone').value;
  var contact = document.getElementById('oqVContact').value;
  opsApi('vendor', { name: name, email: email, phone: phone, contact: contact });
  closeOqModal();
}

function sendInvoiceEmail(ref, vendorName) {
  var info = OPS_DATA.vendors[vendorName] || {};
  var wo = ALLDATA.find(function(r) { return r[10] === ref; });
  var bookmark = wo ? wo[14] : '';
  var property = wo ? wo[0] : '';
  var service = wo ? wo[3] : '';
  var contactName = info.contact || vendorName;

  var defaultSubject = 'Action Required: Upload Invoice & Paperwork — WO ' + ref + ' (' + property + ')';
  var defaultBody = 'Hi ' + contactName + ',\n\n'
    + 'This is a follow-up regarding the completed work order below:\n\n'
    + '  Work Order: ' + ref + '\n'
    + '  Property: ' + property + '\n'
    + (service ? '  Service: ' + service + '\n' : '')
    + '\nPlease upload your invoice and any supporting paperwork (lien waivers, completion photos, etc.) directly into Axxerion using the link below:\n\n'
    + (bookmark ? bookmark + '\n\n' : '[Axxerion link not available — please contact us for upload instructions]\n\n')
    + 'If you are unable to access the link, you may reply to this email with your documents attached and we will upload them on your behalf.\n\n'
    + 'Please submit within 7 business days to avoid payment delays.\n\n'
    + 'Thank you,\nSecure Space Operations';

  var html = '<div class="psl" style="margin-bottom:12px">DRAFT EMAIL — ' + ref + '</div>'
    + (bookmark ? '<div style="margin-bottom:12px;padding:10px 12px;background:rgba(var(--accent-rgb),.08);border:1px solid rgba(var(--accent-rgb),.2);border-radius:6px;font-size:11px"><span style="color:var(--accent);font-weight:600">Axxerion Upload Link:</span> <a href="' + bookmark + '" target="_blank" style="color:var(--accent);word-break:break-all;margin-left:6px">' + bookmark + '</a></div>' : '<div style="margin-bottom:12px;padding:10px 12px;background:rgba(var(--orange-rgb),.08);border:1px solid rgba(var(--orange-rgb),.2);border-radius:6px;font-size:11px;color:var(--orange)">No Axxerion link available — bookmark not set on this WO</div>')
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">TO' + (!info.email ? ' <span style="color:var(--orange)">(no vendor email on file)</span>' : '') + '</label>'
    + '<input type="email" id="oqEmailTo" value="' + (info.email || '').replace(/"/g, '&quot;') + '" placeholder="vendor@example.com" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">SUBJECT</label>'
    + '<input type="text" id="oqEmailSubject" value="' + defaultSubject.replace(/"/g, '&quot;') + '" style="font-family:\'DM Mono\',monospace;font-size:12px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:6px 10px;width:100%"></div>'
    + '<div style="margin-bottom:12px"><label style="font-family:\'DM Mono\',monospace;font-size:10px;color:var(--muted);display:block;margin-bottom:4px">BODY</label>'
    + '<textarea id="oqEmailBody" rows="14" style="font-family:\'DM Mono\',monospace;font-size:11px;background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:8px 10px;width:100%;resize:vertical;line-height:1.6">' + defaultBody.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</textarea></div>'
    + '<div style="display:flex;gap:8px">'
    + '<button class="oq-btn oq-btn-g" style="flex:1;padding:10px" onclick="confirmSendEmail(\'' + ref + '\',\'' + vendorName.replace(/'/g, "\\'") + '\')">Open in Email Client</button>'
    + '<button class="oq-btn" style="flex:1;padding:10px" onclick="logEmailOnly(\'' + ref + '\',\'' + vendorName.replace(/'/g, "\\'") + '\')">Log as Sent</button>'
    + '</div>'
    + '<div style="margin-top:6px;font-size:10px;color:var(--muted);text-align:center">"Open in Email Client" launches your default email app · "Log as Sent" just records it</div>';
  openOqModal(html);
}

function confirmSendEmail(ref, vendorName) {
  var to = document.getElementById('oqEmailTo').value;
  if (!to) { alert('Please enter a recipient email address'); return; }
  var subject = document.getElementById('oqEmailSubject').value;
  var body = document.getElementById('oqEmailBody').value;
  // Save vendor email if not already stored
  var info = OPS_DATA.vendors[vendorName] || {};
  if (!info.email && to) {
    opsApi('vendor', { name: vendorName, email: to, phone: info.phone || '', contact: info.contact || '' });
  }
  // Open mailto link
  var mailto = 'mailto:' + encodeURIComponent(to)
    + '?subject=' + encodeURIComponent(subject)
    + '&body=' + encodeURIComponent(body);
  window.open(mailto, '_blank');
  // Log it
  opsApi('email', { ref: ref, to: to, type: 'invoice', subject: subject });
  opsApi('log', { ref: ref, action: 'email_sent', note: 'Invoice/paperwork request sent to ' + to });
  closeOqModal();
}

function logEmailOnly(ref, vendorName) {
  var to = document.getElementById('oqEmailTo').value;
  if (!to) { alert('Please enter a recipient email address'); return; }
  var subject = document.getElementById('oqEmailSubject').value;
  // Save vendor email if not already stored
  var info = OPS_DATA.vendors[vendorName] || {};
  if (!info.email && to) {
    opsApi('vendor', { name: vendorName, email: to, phone: info.phone || '', contact: info.contact || '' });
  }
  opsApi('email', { ref: ref, to: to, type: 'invoice', subject: subject });
  opsApi('log', { ref: ref, action: 'email_sent', note: 'Invoice/paperwork request sent to ' + to + ' (logged manually)' });
  closeOqModal();
}

// ── Init ──
function initOpsQueue(hashQueue) {
  var savedQ = hashQueue || localStorage.getItem('ax_active_queue') || 'q1';
  loadStoreDirectory();
  loadOpsData(function() {
    switchQueue(savedQ, document.querySelector('.oq-tab[data-queue="' + savedQ + '"]'));
  });
  // KPI click handlers
  document.querySelectorAll('.oq-kpi').forEach(function(k) {
    k.addEventListener('click', function() {
      var q = k.getAttribute('data-queue');
      var tab = document.querySelector('.oq-tab[data-queue="' + q + '"]');
      switchQueue(q, tab);
    });
  });
}
