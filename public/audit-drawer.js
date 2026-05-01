/* Audit history side drawer.
 * Usage: window.openAuditDrawer({ entityType, entityId, title })
 * Self-contained — appends its own DOM on first call. Loads /api/audit/entity/:type/:id.
 */
(function(){
  if (window.__auditDrawerInit) return;
  window.__auditDrawerInit = true;

  const ACTION_LABELS = {
    "call.logged":                 "Call logged",
    "appt.set":                    "Appointment set",
    "appt.confirmed":              "Appointment confirmed",
    "email.sent":                  "Email sent",
    "vendor.contact_updated":      "Vendor contact updated",
    "note.added":                  "Note added",
    "wo.dismissed":                "Dismissed from queue",
    "wo.status_changed":           "Status changed",
    "wo.vendor_changed":           "Vendor changed",
    "wo.scheduled_from_changed":   "Scheduled Start changed",
    "wo.scheduled_until_changed":  "Scheduled End changed",
    "wo.actual_start_changed":     "Actual Start changed",
    "wo.actual_end_changed":       "Actual End changed",
    "wo.closed_changed":           "Closed flag changed",
    "wo.vendor_invoice_changed":   "Vendor invoice # changed",
    "wo.priority_changed":         "Priority changed",
    "wo.executor_changed":         "Executor changed",
    "wo.created":                  "Work order created",
    "user.login":                  "Logged in",
    "user.role_changed":           "Role changed",
    "budget.set":                  "Budget set",
  };

  function el(tag, attrs, ...children) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === "style") e.setAttribute("style", attrs[k]);
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k.startsWith("on")) e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    }
    children.forEach(c => { if (c == null) return; e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }

  function fmtTime(iso) {
    try {
      const d = new Date(iso);
      return d.toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" });
    } catch (e) { return iso; }
  }

  function actorBadge(row) {
    if (row.actor_type === "axxerion") {
      return '<span style="display:inline-block;background:rgba(var(--accent-rgb),.15);color:var(--accent);font-size:10px;padding:2px 6px;border-radius:10px;font-weight:600">AXXERION</span>';
    }
    if (row.actor === "system") {
      return '<span style="display:inline-block;background:rgba(var(--muted-rgb,128,128,128),.2);color:var(--muted);font-size:10px;padding:2px 6px;border-radius:10px">SYSTEM</span>';
    }
    const short = String(row.actor || "").split("@")[0] || "user";
    return '<span style="display:inline-block;background:rgba(var(--green-rgb),.15);color:var(--green);font-size:10px;padding:2px 6px;border-radius:10px;font-weight:600">' + short.toUpperCase() + '</span>';
  }

  function summarizeChange(row) {
    const ov = row.old_value;
    const nv = row.new_value;
    // Status / vendor / date scalar diffs: "old → new"
    if (ov != null && typeof ov !== "object" && nv != null && typeof nv !== "object") {
      return '<span style="color:var(--muted)">' + (ov || "—") + '</span> → <strong>' + (nv || "—") + '</strong>';
    }
    // Notes
    if (row.action === "note.added" && nv && nv.text) return '"' + String(nv.text).slice(0, 200) + '"';
    // Calls
    if (row.action === "call.logged" && nv) {
      const a = nv.action || "call";
      const note = nv.note ? ' — ' + String(nv.note).slice(0, 200) : '';
      return '<strong>' + a + '</strong>' + note;
    }
    // Appointments
    if ((row.action === "appt.set" || row.action === "appt.confirmed") && nv) {
      return (nv.date || "TBD") + (nv.time ? " " + nv.time : "") + (nv.confirmed ? " ✓" : "");
    }
    // Emails
    if (row.action === "email.sent" && nv) {
      return 'to <strong>' + (nv.to || "?") + '</strong>: ' + (nv.subject || nv.type || "");
    }
    // Vendor contact
    if (row.action === "vendor.contact_updated" && nv) {
      const parts = [];
      if (nv.email) parts.push("email: " + nv.email);
      if (nv.phone) parts.push("phone: " + nv.phone);
      return parts.join(" · ");
    }
    // Dismissals
    if (row.action === "wo.dismissed") {
      return row.metadata && row.metadata.queue ? "from " + row.metadata.queue.toUpperCase() : "";
    }
    // WO created
    if (row.action === "wo.created" && nv) {
      return (nv.status || "") + " · " + (nv.vendor || "no vendor") + " · " + (nv.property || "");
    }
    // Role / budget
    if (row.action === "user.role_changed" && nv) return "role → " + (nv.role || "?");
    if (row.action === "budget.set" && nv) return "$" + nv.budget;
    return "";
  }

  function renderRow(row) {
    const label = ACTION_LABELS[row.action] || row.action;
    const summary = summarizeChange(row);
    const actor = actorBadge(row);
    return el("div", {
      style: "padding:14px 16px;border-bottom:1px solid var(--line);display:flex;flex-direction:column;gap:6px"
    },
      el("div", {
        style: "display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:11px;color:var(--muted);font-family:'DM Mono',monospace"
      },
        el("span", { html: actor + ' &middot; ' + fmtTime(row.occurred_at) }),
        el("span", { style: "opacity:.5" }, "#" + row.id)
      ),
      el("div", { style: "font-size:13px;font-weight:600" }, label),
      summary ? el("div", { style: "font-size:12px;line-height:1.5;color:var(--text)", html: summary }) : null
    );
  }

  let drawerEl = null;
  let backdropEl = null;
  let titleEl = null;
  let bodyEl = null;

  function build() {
    backdropEl = el("div", {
      style: "position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:1000;display:none;backdrop-filter:blur(2px)",
      onclick: close
    });
    drawerEl = el("div", {
      style: "position:fixed;top:0;right:0;bottom:0;width:480px;max-width:100vw;background:var(--card);border-left:1px solid var(--line);box-shadow:-8px 0 32px rgba(0,0,0,.3);z-index:1001;display:none;flex-direction:column;font-family:Inter,sans-serif"
    });
    const header = el("div", {
      style: "padding:18px 20px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;flex-shrink:0"
    });
    titleEl = el("div", { style: "display:flex;flex-direction:column;gap:2px" });
    const closeBtn = el("button", {
      style: "background:none;border:none;color:var(--muted);font-size:24px;cursor:pointer;padding:0;width:32px;height:32px",
      onclick: close
    }, "×");
    header.appendChild(titleEl);
    header.appendChild(closeBtn);
    bodyEl = el("div", { style: "flex:1;overflow-y:auto" });
    drawerEl.appendChild(header);
    drawerEl.appendChild(bodyEl);
    document.body.appendChild(backdropEl);
    document.body.appendChild(drawerEl);
    document.addEventListener("keydown", e => { if (e.key === "Escape" && drawerEl.style.display === "flex") close(); });
  }

  function close() {
    if (!drawerEl) return;
    drawerEl.style.display = "none";
    backdropEl.style.display = "none";
  }

  async function open({ entityType, entityId, title, subtitle, bookmark }) {
    if (!drawerEl) build();
    titleEl.innerHTML = '';
    titleEl.appendChild(el("div", { style: "font-size:16px;font-weight:700;font-family:Syne,sans-serif" }, title || entityId));
    if (subtitle) titleEl.appendChild(el("div", { style: "font-size:11px;color:var(--muted);font-family:'DM Mono',monospace" }, subtitle));
    if (bookmark) {
      titleEl.appendChild(el("a", { href: bookmark, target: "_blank", style: "font-size:11px;color:var(--accent);text-decoration:none;margin-top:2px" }, "Open in Axxerion ↗"));
    }
    bodyEl.innerHTML = '<div style="padding:32px;text-align:center;color:var(--muted)">Loading history...</div>';
    drawerEl.style.display = "flex";
    backdropEl.style.display = "block";
    try {
      const res = await fetch("/api/audit/entity/" + encodeURIComponent(entityType) + "/" + encodeURIComponent(entityId) + "?limit=300");
      const data = await res.json();
      bodyEl.innerHTML = '';
      if (!data.rows || !data.rows.length) {
        bodyEl.appendChild(el("div", { style: "padding:32px;text-align:center;color:var(--muted);font-size:13px" }, "No activity yet for this " + entityType + "."));
        return;
      }
      data.rows.forEach(r => bodyEl.appendChild(renderRow(r)));
    } catch (e) {
      bodyEl.innerHTML = '<div style="padding:32px;text-align:center;color:var(--red);font-size:13px">Failed to load history.</div>';
    }
  }

  window.openAuditDrawer = open;
  window.openWOHistory = function(ref, opts) {
    open({
      entityType: "wo",
      entityId: ref,
      title: ref,
      subtitle: opts && opts.property ? opts.property : "",
      bookmark: opts && opts.bookmark ? opts.bookmark : null,
    });
  };
})();
