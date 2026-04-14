// lib/chat-usage.js — Budget check, usage logging, admin helpers (server-only)

const DEFAULT_BUDGET = 25.00;

// Claude Sonnet 4.6 pricing per 1M tokens (USD)
const PRICE_INPUT_MTOK = 3.00;
const PRICE_OUTPUT_MTOK = 15.00;

// Admins — bypass budget check, can access admin page
const ADMIN_EMAILS = [
  "ttran@insitepg.com",
];

function isAdmin(email) {
  if (!email) return false;
  return ADMIN_EMAILS.includes(email.toLowerCase());
}

function estimateCost(inputTokens, outputTokens) {
  return (inputTokens / 1000000) * PRICE_INPUT_MTOK + (outputTokens / 1000000) * PRICE_OUTPUT_MTOK;
}

function startOfMonth() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), 1);
}

async function initChatTables(pool) {
  if (!pool) return;
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS chat_logs (
      id SERIAL PRIMARY KEY,
      user_email TEXT NOT NULL,
      user_name TEXT,
      input_tokens INTEGER NOT NULL DEFAULT 0,
      output_tokens INTEGER NOT NULL DEFAULT 0,
      estimated_cost DOUBLE PRECISION NOT NULL DEFAULT 0,
      tool_call_count INTEGER NOT NULL DEFAULT 0,
      model TEXT,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_chat_logs_user_time ON chat_logs(user_email, created_at)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_chat_logs_time ON chat_logs(created_at)`);

    await pool.query(`CREATE TABLE IF NOT EXISTS user_budgets (
      user_email TEXT PRIMARY KEY,
      user_name TEXT,
      monthly_budget DOUBLE PRECISION NOT NULL DEFAULT ${DEFAULT_BUDGET},
      is_admin BOOLEAN NOT NULL DEFAULT FALSE,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_by TEXT
    )`);

    // Seed admins
    for (const email of ADMIN_EMAILS) {
      await pool.query(
        `INSERT INTO user_budgets (user_email, is_admin) VALUES ($1, TRUE)
         ON CONFLICT (user_email) DO UPDATE SET is_admin = TRUE`,
        [email]
      );
    }
    console.log("[Chat] chat_logs + user_budgets tables ready");
  } catch (e) {
    console.error("[Chat] initChatTables error:", e.message);
  }
}

async function getOrCreateUser(pool, email, name) {
  if (!pool) return { monthly_budget: DEFAULT_BUDGET, is_admin: isAdmin(email) };
  try {
    const existing = await pool.query("SELECT * FROM user_budgets WHERE user_email = $1", [email]);
    if (existing.rows.length) {
      // Sync admin flag from config (in case ADMIN_EMAILS changed)
      const shouldBeAdmin = isAdmin(email);
      if (existing.rows[0].is_admin !== shouldBeAdmin) {
        await pool.query("UPDATE user_budgets SET is_admin = $1 WHERE user_email = $2", [shouldBeAdmin, email]);
        existing.rows[0].is_admin = shouldBeAdmin;
      }
      if (name && existing.rows[0].user_name !== name) {
        await pool.query("UPDATE user_budgets SET user_name = $1 WHERE user_email = $2", [name, email]);
      }
      return existing.rows[0];
    }
    const inserted = await pool.query(
      `INSERT INTO user_budgets (user_email, user_name, monthly_budget, is_admin) VALUES ($1, $2, $3, $4) RETURNING *`,
      [email, name || null, DEFAULT_BUDGET, isAdmin(email)]
    );
    return inserted.rows[0];
  } catch (e) {
    console.error("[Chat] getOrCreateUser error:", e.message);
    return { monthly_budget: DEFAULT_BUDGET, is_admin: isAdmin(email) };
  }
}

async function getMtdCost(pool, email) {
  if (!pool) return 0;
  try {
    const som = startOfMonth();
    const r = await pool.query(
      "SELECT COALESCE(SUM(estimated_cost),0) AS total FROM chat_logs WHERE user_email = $1 AND created_at >= $2",
      [email, som]
    );
    return parseFloat(r.rows[0].total) || 0;
  } catch (e) {
    console.error("[Chat] getMtdCost error:", e.message);
    return 0;
  }
}

async function checkBudget(pool, email, name) {
  const user = await getOrCreateUser(pool, email, name);
  const mtdCost = await getMtdCost(pool, email);
  const budget = parseFloat(user.monthly_budget) || DEFAULT_BUDGET;
  const percentUsed = budget > 0 ? (mtdCost / budget) * 100 : 0;
  const allowed = user.is_admin || mtdCost < budget;
  return { allowed, mtdCost, monthlyBudget: budget, percentUsed, isAdmin: !!user.is_admin };
}

async function logChatUsage(pool, { email, name, inputTokens, outputTokens, toolCallCount, model }) {
  if (!pool) return;
  try {
    const cost = estimateCost(inputTokens, outputTokens);
    await pool.query(
      `INSERT INTO chat_logs (user_email, user_name, input_tokens, output_tokens, estimated_cost, tool_call_count, model)
       VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [email, name || null, inputTokens, outputTokens, cost, toolCallCount || 0, model || null]
    );
  } catch (e) {
    console.error("[Chat] logChatUsage error:", e.message);
  }
}

async function getUsageStats(pool) {
  if (!pool) return { summary: { mtdCost: 0, mtdQueries: 0, activeUsers: 0 }, users: [] };
  const som = startOfMonth();
  const summary = await pool.query(
    `SELECT COALESCE(SUM(estimated_cost),0) AS cost, COUNT(*) AS queries, COUNT(DISTINCT user_email) AS users
     FROM chat_logs WHERE created_at >= $1`,
    [som]
  );
  const users = await pool.query(
    `SELECT b.user_email, b.user_name, b.monthly_budget, b.is_admin,
       COALESCE(SUM(l.estimated_cost),0) AS mtd_cost,
       COALESCE(COUNT(l.id),0) AS mtd_queries
     FROM user_budgets b
     LEFT JOIN chat_logs l ON l.user_email = b.user_email AND l.created_at >= $1
     GROUP BY b.user_email, b.user_name, b.monthly_budget, b.is_admin
     ORDER BY mtd_cost DESC, b.user_email ASC`,
    [som]
  );
  return {
    summary: {
      mtdCost: parseFloat(summary.rows[0].cost) || 0,
      mtdQueries: parseInt(summary.rows[0].queries) || 0,
      activeUsers: parseInt(summary.rows[0].users) || 0,
    },
    users: users.rows.map(r => ({
      email: r.user_email,
      name: r.user_name,
      monthlyBudget: parseFloat(r.monthly_budget),
      isAdmin: !!r.is_admin,
      mtdCost: parseFloat(r.mtd_cost),
      mtdQueries: parseInt(r.mtd_queries),
    })),
  };
}

async function upsertUserBudget(pool, email, name, budget, updatedBy) {
  if (!pool) return;
  if (budget < 0 || budget > 1000) throw new Error("Budget must be between $0 and $1000");
  await pool.query(
    `INSERT INTO user_budgets (user_email, user_name, monthly_budget, updated_by, updated_at)
     VALUES ($1, $2, $3, $4, NOW())
     ON CONFLICT (user_email) DO UPDATE SET monthly_budget = $3, user_name = COALESCE($2, user_budgets.user_name), updated_by = $4, updated_at = NOW()`,
    [email, name || null, budget, updatedBy]
  );
}

async function addAdmin(pool, email, addedBy) {
  if (!pool) return;
  email = email.toLowerCase().trim();
  await pool.query(
    `INSERT INTO user_budgets (user_email, is_admin, updated_by, updated_at)
     VALUES ($1, TRUE, $2, NOW())
     ON CONFLICT (user_email) DO UPDATE SET is_admin = TRUE, updated_by = $2, updated_at = NOW()`,
    [email, addedBy]
  );
}

async function removeAdmin(pool, email, removedBy) {
  if (!pool) return;
  email = email.toLowerCase().trim();
  // Protect hardcoded admins from being removed
  if (ADMIN_EMAILS.includes(email)) throw new Error("Cannot remove a built-in admin");
  await pool.query(
    `UPDATE user_budgets SET is_admin = FALSE, updated_by = $1, updated_at = NOW() WHERE user_email = $2`,
    [removedBy, email]
  );
}

module.exports = {
  DEFAULT_BUDGET,
  PRICE_INPUT_MTOK,
  PRICE_OUTPUT_MTOK,
  isAdmin,
  estimateCost,
  initChatTables,
  getOrCreateUser,
  checkBudget,
  logChatUsage,
  getUsageStats,
  upsertUserBudget,
  addAdmin,
  removeAdmin,
  getMtdCost,
};
