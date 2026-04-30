// Unified audit log writer + readers.
//
// Single Postgres table audit_log captures every meaningful state change in the
// system: ops user actions, Axxerion-side data changes, admin events, and logins.
// Reads are served by getEntityHistory (per-WO drawer) and getAuditFeed (system feed).

let pool = null;

function init(pgPool) {
  pool = pgPool;
}

async function ensureTable() {
  if (!pool) return;
  await pool.query(`CREATE TABLE IF NOT EXISTS audit_log (
    id            BIGSERIAL PRIMARY KEY,
    occurred_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    actor         TEXT NOT NULL,
    actor_type    TEXT NOT NULL,
    action        TEXT NOT NULL,
    entity_type   TEXT NOT NULL,
    entity_id     TEXT,
    old_value     JSONB,
    new_value     JSONB,
    metadata      JSONB,
    session_id    TEXT
  )`);
  await pool.query(`CREATE INDEX IF NOT EXISTS idx_audit_entity ON audit_log(entity_type, entity_id, occurred_at DESC)`);
  await pool.query(`CREATE INDEX IF NOT EXISTS idx_audit_actor  ON audit_log(actor, occurred_at DESC)`);
  await pool.query(`CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action, occurred_at DESC)`);
  await pool.query(`CREATE INDEX IF NOT EXISTS idx_audit_time   ON audit_log(occurred_at DESC)`);
}

// Best-effort write — failures logged but never thrown so they can't break
// the originating user action. Audit should be invisible until you query it.
async function logEvent(evt) {
  if (!pool) return;
  try {
    await pool.query(
      `INSERT INTO audit_log (actor, actor_type, action, entity_type, entity_id, old_value, new_value, metadata, session_id)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
      [
        evt.actor || "unknown",
        evt.actorType || "user",
        evt.action,
        evt.entityType,
        evt.entityId || null,
        evt.oldValue == null ? null : JSON.stringify(evt.oldValue),
        evt.newValue == null ? null : JSON.stringify(evt.newValue),
        evt.metadata == null ? null : JSON.stringify(evt.metadata),
        evt.sessionId || null,
      ]
    );
  } catch (e) {
    console.error("[Audit] Insert failed:", e.message, "evt:", evt.action, evt.entityId);
  }
}

// Pull actor info from a request. Pulls from session first (Azure SSO),
// then falls back to req.body.user (legacy ops calls), then "anonymous".
function actorFromReq(req) {
  if (req && req.session && req.session.user && req.session.user.email) {
    return { actor: req.session.user.email, actorType: "user", sessionId: req.sessionID || null };
  }
  if (req && req.body && req.body.user) {
    return { actor: String(req.body.user), actorType: "user", sessionId: null };
  }
  return { actor: "anonymous", actorType: "user", sessionId: null };
}

// Per-entity timeline (drives the per-WO history drawer).
async function getEntityHistory(entityType, entityId, limit = 200) {
  if (!pool) return [];
  const res = await pool.query(
    `SELECT id, occurred_at, actor, actor_type, action, entity_type, entity_id,
            old_value, new_value, metadata
     FROM audit_log
     WHERE entity_type = $1 AND entity_id = $2
     ORDER BY occurred_at DESC
     LIMIT $3`,
    [entityType, entityId, limit]
  );
  return res.rows;
}

// System feed with optional filters (drives /audit page).
async function getAuditFeed({ actor, actorType, action, entityType, entityId, since, until, limit = 100, offset = 0 } = {}) {
  if (!pool) return { rows: [], total: 0 };
  const where = [];
  const params = [];
  const push = (sql, val) => { params.push(val); where.push(sql.replace("$?", "$" + params.length)); };

  if (actor)       push("actor = $?", actor);
  if (actorType)   push("actor_type = $?", actorType);
  if (action)      push("action = $?", action);
  if (entityType)  push("entity_type = $?", entityType);
  if (entityId)    push("entity_id = $?", entityId);
  if (since)       push("occurred_at >= $?", since);
  if (until)       push("occurred_at <= $?", until);

  const whereSql = where.length ? "WHERE " + where.join(" AND ") : "";

  const dataSql = `SELECT id, occurred_at, actor, actor_type, action, entity_type, entity_id,
                          old_value, new_value, metadata
                   FROM audit_log
                   ${whereSql}
                   ORDER BY occurred_at DESC
                   LIMIT ${Math.max(1, Math.min(500, +limit))}
                   OFFSET ${Math.max(0, +offset)}`;
  const countSql = `SELECT COUNT(*)::int AS n FROM audit_log ${whereSql}`;

  const [dataRes, countRes] = await Promise.all([
    pool.query(dataSql, params),
    pool.query(countSql, params),
  ]);

  return { rows: dataRes.rows, total: countRes.rows[0].n };
}

// Diff helper used by Axxerion incremental fetch — emits one event per
// changed tracked field. Tracked fields are the ones Ops actually cares about.
const TRACKED_WO_FIELDS = [
  { key: "Status",                 action: "wo.status_changed" },
  { key: "Vendor",                 action: "wo.vendor_changed" },
  { key: "Scheduled from",         action: "wo.scheduled_from_changed" },
  { key: "Scheduled until",        action: "wo.scheduled_until_changed" },
  { key: "Actual start date",      action: "wo.actual_start_changed" },
  { key: "End",                    action: "wo.actual_end_changed" },
  { key: "Actual end date",        action: "wo.actual_end_changed" },
  { key: "Closed",                 action: "wo.closed_changed" },
  { key: "Vendor Invoice #",       action: "wo.vendor_invoice_changed" },
  { key: "Priority",               action: "wo.priority_changed" },
  { key: "Executor",               action: "wo.executor_changed" },
];

// Emits 0..N events for a single WO update. Caller passes prior snapshot + new snapshot.
async function diffAndLogWO(prior, next) {
  if (!pool || !prior || !next) return;
  const ref = next.Reference || prior.Reference;
  for (const f of TRACKED_WO_FIELDS) {
    const a = prior[f.key] || "";
    const b = next[f.key] || "";
    if (a !== b) {
      await logEvent({
        actor: "axxerion",
        actorType: "axxerion",
        action: f.action,
        entityType: "wo",
        entityId: ref,
        oldValue: a,
        newValue: b,
        metadata: { field: f.key, source: "incremental_fetch" },
      });
    }
  }
}

// New WO appearing for the first time in incremental fetch.
async function logWOCreated(wo) {
  if (!pool || !wo) return;
  await logEvent({
    actor: "axxerion",
    actorType: "axxerion",
    action: "wo.created",
    entityType: "wo",
    entityId: wo.Reference,
    newValue: {
      status: wo.Status,
      vendor: wo.Vendor,
      property: wo.Property,
      priority: wo.Priority,
      created: wo.Created,
    },
    metadata: { source: "incremental_fetch" },
  });
}

// Per-user activity counts across the audited ops actions, over a time window.
// Drives the Activity tab — answers "is the new ops person actually working?"
async function getActivityStats(daysBack = 7) {
  if (!pool) return [];
  const res = await pool.query(`
    SELECT
      actor                                                                    AS email,
      COUNT(*) FILTER (WHERE action = 'call.logged')::int                      AS calls_logged,
      COUNT(*) FILTER (WHERE action = 'appt.set')::int                         AS appts_set,
      COUNT(*) FILTER (WHERE action = 'appt.confirmed')::int                   AS appts_confirmed,
      COUNT(*) FILTER (WHERE action = 'email.sent')::int                       AS emails_sent,
      COUNT(*) FILTER (WHERE action = 'vendor.contact_updated')::int           AS vendor_edits,
      COUNT(*) FILTER (WHERE action = 'note.added')::int                       AS notes_added,
      COUNT(*) FILTER (WHERE action = 'wo.dismissed')::int                     AS dismissals,
      COUNT(*)::int                                                            AS total_actions,
      MAX(occurred_at)                                                         AS last_action,
      COUNT(DISTINCT DATE(occurred_at))::int                                   AS active_days,
      COUNT(DISTINCT entity_id)::int                                           AS distinct_entities
    FROM audit_log
    WHERE actor_type = 'user'
      AND action IN ('call.logged','appt.set','appt.confirmed','email.sent','vendor.contact_updated','note.added','wo.dismissed')
      AND occurred_at >= NOW() - ($1::int || ' days')::interval
    GROUP BY actor
    ORDER BY total_actions DESC
  `, [daysBack]);
  return res.rows;
}

// Aggregated login frequency per user (drives the Logins tab).
async function getLoginStats() {
  if (!pool) return [];
  const res = await pool.query(`
    SELECT
      actor                                                        AS email,
      COUNT(*)::int                                                AS total_logins,
      MAX(occurred_at)                                             AS last_login,
      MIN(occurred_at)                                             AS first_login,
      COUNT(*) FILTER (WHERE occurred_at >= NOW() - INTERVAL '24 hours')::int AS last_24h,
      COUNT(*) FILTER (WHERE occurred_at >= NOW() - INTERVAL '7 days')::int  AS last_7d,
      COUNT(*) FILTER (WHERE occurred_at >= NOW() - INTERVAL '30 days')::int AS last_30d,
      COUNT(DISTINCT DATE(occurred_at))::int                       AS distinct_days
    FROM audit_log
    WHERE action = 'user.login'
    GROUP BY actor
    ORDER BY last_login DESC
  `);
  return res.rows;
}

module.exports = {
  init,
  ensureTable,
  logEvent,
  actorFromReq,
  getEntityHistory,
  getAuditFeed,
  getLoginStats,
  getActivityStats,
  diffAndLogWO,
  logWOCreated,
};
