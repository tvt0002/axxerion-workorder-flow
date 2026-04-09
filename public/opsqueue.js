/* ── Ops Queue Logic ── */
var OPS_DATA = { logs: {}, appointments: {}, emails: {}, vendors: {}, notes: {}, dismissed: {} };
var OPS_QUEUE = 'q1';
var OPS_LOADED = false;

// Queue status definitions (aligned with Axxerion WR-WO Statuses guide)
var Q1_PRIORITIES = new Set(['Info Needed']);
var Q1_STATUSES = new Set(['Info Needed', 'NFO NEEDED', 'Nfo Needed', 'Info needed']);
var Q2_STATUSES = new Set(['Assigned', 'Accepted']);
var Q4_STATUSES = new Set(['Work finished', 'Work Finished']);

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
    if (!Q1_PRIORITIES.has(r[2]) && !Q1_STATUSES.has(r[1])) return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q1) return false;
    return true;
  }).map(function(r) {
    var ref = r[10] || '';
    var logs = OPS_DATA.logs[ref] || [];
    var note = OPS_DATA.notes[ref] || {};
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', subject: r[3], created: r[12], requestor: r[16] || r[8], bookmark: r[14], logs: logs, note: note.text || '', age: daysAgo(r[12]), lastAction: logs.length ? logs[0] : null };
  }).sort(function(a, b) { return b.age - a.age; });
}

function getQ2Data() {
  // WOs assigned 24+ hours, not yet scheduled
  if (!ALLDATA || !ALLDATA.length) return [];
  return ALLDATA.filter(function(r) {
    if (r[15] === 'Request') return false;
    if (!Q2_STATUSES.has(r[1])) return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q2) return false;
    // Has appointment already? Check Axxerion Scheduled from first, then manual ops tracking
    if (r[19]) return false;
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
  }).sort(function(a, b) { return b.hours - a.hours; });
}

function getQ3Data() {
  // WOs with appointment today — check both Axxerion Scheduled from and manual ops tracking
  if (!ALLDATA || !ALLDATA.length) return [];
  var today = todayStr();
  return ALLDATA.filter(function(r) {
    if (r[15] === 'Request') return false;
    var ref = r[10] || '';
    if (OPS_DATA.dismissed && OPS_DATA.dismissed[ref] && OPS_DATA.dismissed[ref].q3) return false;
    // Check Axxerion Scheduled from date
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
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', service: r[3], vendor: r[6] || '', executor: r[18] || '', time: time, confirmed: appt.confirmed, bookmark: r[14], logs: logs, lastAction: logs.length ? logs[0] : null };
  }).sort(function(a, b) { return (a.time || 'ZZ').localeCompare(b.time || 'ZZ'); });
}

function getQ4Data() {
  // WOs with work done, awaiting invoice
  if (!ALLDATA || !ALLDATA.length) return [];
  return ALLDATA.filter(function(r) {
    if (r[15] !== 'Work Order') return false;
    if (!Q4_STATUSES.has(r[1])) return false;
    var ref = r[10] || '';
    if (!ref) return false;
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
    return { row: r, ref: ref, property: r[0], status: r[1], priority: r[2] || '', service: r[3], vendor: vendor, vendorEmail: vendorInfo.email || '', bookmark: r[14], logs: logs, emails: emails, emailCount: emails.length, lastEmail: emails.length ? emails[0] : null, lastAction: logs.length ? logs[0] : null, finishedDate: finishedDate, daysSinceFinished: daysAgo(finishedDate) };
  }).sort(function(a, b) {
    // Sort by: no emails first, then oldest finished date
    if (a.emailCount === 0 && b.emailCount > 0) return -1;
    if (b.emailCount === 0 && a.emailCount > 0) return 1;
    return b.daysSinceFinished - a.daysSinceFinished;
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
  localStorage.setItem('ax_active_queue', q);
  document.querySelectorAll('.oq-tab').forEach(function(b) { b.classList.remove('active'); });
  document.querySelectorAll('.oq-kpi').forEach(function(k) { k.style.boxShadow = ''; });
  if (btn) btn.classList.add('active');
  var kpi = document.querySelector('.oq-kpi[data-queue="' + q + '"]');
  if (kpi) kpi.style.boxShadow = '0 0 0 2px rgba(var(--accent-rgb),.4)';
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
    headHTML = '<tr><th>Age</th><th>Reference</th><th>Property</th><th>Status</th><th>Priority</th><th>Subject</th><th>Requestor</th><th>Last Action</th><th>Actions</th></tr>';
    rowsHTML = items.map(function(d) {
      var ageClass = d.age >= 3 ? 'color:var(--red)' : d.age >= 1 ? 'color:var(--orange)' : 'color:var(--green)';
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastAct = d.lastAction ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastAction.date) + ' · ' + d.lastAction.action + '</span>' : '<span style="font-size:10px;color:var(--red)">No action yet</span>';
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + ageClass + '">' + d.age + 'd</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.subject + '</td>'
        + '<td>' + d.requestor + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q1\')">Log Action</button> <button class="oq-btn oq-btn-dim" onclick="dismissOq(\'' + d.ref + '\',\'q1\')">Dismiss</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q2') {
    items = getQ2Data();
    headHTML = '<tr><th>Hours</th><th>Reference</th><th>Property</th><th>Status</th><th>Priority</th><th>Service Type</th><th>Vendor</th><th>Phone</th><th>Sched Start</th><th>Sched End</th><th>Actual Start</th><th>Actual End</th><th>Appointment</th><th>Calls</th><th>Last Action</th><th>Actions</th></tr>';
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
        + '<td style="' + dateSty + '">' + (d.schedFrom || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.schedUntil || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualStart || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (d.actualEnd || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="' + dateSty + '">' + (apptTxt || '<span style="color:var(--muted)">—</span>') + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;text-align:center">' + d.callCount + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q2\')">Log Call</button> <button class="oq-btn oq-btn-g" onclick="openOqSchedule(\'' + d.ref + '\')">Schedule</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q3') {
    items = getQ3Data();
    headHTML = '<tr><th>Time</th><th>Reference</th><th>Property</th><th>Status</th><th>Priority</th><th>Vendor</th><th>Service Type</th><th>Confirmed</th><th>Last Action</th><th>Actions</th></tr>';
    rowsHTML = items.map(function(d) {
      var confBadge = d.confirmed ? '<span style="color:var(--green);font-weight:600">Confirmed</span>' : '<span style="color:var(--orange);font-weight:600">Pending</span>';
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastAct = d.lastAction ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastAction.date) + ' · ' + d.lastAction.action + '</span>' : '<span style="font-size:10px;color:var(--muted)">—</span>';
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:12px;font-weight:600">' + (d.time || 'TBD') + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td>' + confBadge + '</td>'
        + '<td>' + lastAct + '</td>'
        + '<td><button class="oq-btn oq-btn-g" onclick="confirmAppt(\'' + d.ref + '\')">Confirm</button> <button class="oq-btn" onclick="openOqAction(\'' + d.ref + '\',\'q3\')">Log</button> <button class="oq-btn oq-btn-r" onclick="openOqAction(\'' + d.ref + '\',\'q3\',\'noshow\')">No Show</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q4') {
    items = getQ4Data();
    headHTML = '<tr><th>Days</th><th>Reference</th><th>Property</th><th>Vendor</th><th>Service Type</th><th>Status</th><th>Priority</th><th>Emails Sent</th><th>Last Email</th><th>Actions</th></tr>';
    rowsHTML = items.map(function(d) {
      var refLink = d.bookmark ? '<a href="' + d.bookmark + '" target="_blank" style="color:var(--accent)">' + d.ref + '</a>' : d.ref;
      var lastEmail = d.lastEmail ? '<span style="font-size:10px;color:var(--muted)">' + fmtDate(d.lastEmail.sentAt) + '</span>' : '<span style="font-size:10px;color:var(--red)">Never sent</span>';
      var emailBtn = '<button class="oq-btn oq-btn-g" onclick="sendInvoiceEmail(\'' + d.ref + '\',\'' + d.vendor.replace(/'/g, "\\'") + '\')">Draft Email</button>';
      var daysClass = d.daysSinceFinished >= 14 ? 'color:var(--red)' : d.daysSinceFinished >= 7 ? 'color:var(--orange)' : 'color:var(--green)';
      return '<tr>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + daysClass + '">' + d.daysSinceFinished + 'd</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px">' + refLink + '</td>'
        + '<td>' + d.property + '</td>'
        + '<td>' + d.vendor + '</td>'
        + '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + d.service + '</td>'
        + '<td>' + (typeof statusTooltipHTML==='function'?statusTooltipHTML(d.status):d.status) + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;font-size:11px;' + priColor(d.priority) + '">' + d.priority + '</td>'
        + '<td style="font-family:\'DM Mono\',monospace;text-align:center">' + d.emailCount + '</td>'
        + '<td>' + lastEmail + '</td>'
        + '<td>' + emailBtn + ' <button class="oq-btn oq-btn-dim" onclick="dismissOq(\'' + d.ref + '\',\'q4\')">Dismiss</button></td>'
        + '</tr>';
    }).join('');
  }

  else if (OPS_QUEUE === 'q5') {
    items = getQ5Data();
    headHTML = '<tr><th>Sent</th><th>Reference</th><th>Property</th><th>Vendor</th><th>Service Type</th><th>To</th><th>Subject</th><th>Actions</th></tr>';
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
        + '<td><button class="oq-btn oq-btn-g" onclick="sendInvoiceEmail(\'' + d.ref + '\',\'' + d.vendor.replace(/'/g, "\\'") + '\')">Resend</button></td>'
        + '</tr>';
    }).join('');
  }

  head.innerHTML = headHTML;
  body.innerHTML = rowsHTML;
  empty.style.display = items.length ? 'none' : 'block';
  document.getElementById('oqTable').style.display = items.length ? '' : 'none';
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
function initOpsQueue() {
  var savedQ = localStorage.getItem('ax_active_queue') || 'q1';
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
