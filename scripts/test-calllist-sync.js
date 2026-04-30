// CLI: end-to-end call directory sync.
// Usage:
//   node scripts/test-calllist-sync.js              # DRY RUN (no GitHub writes)
//   node scripts/test-calllist-sync.js --commit     # Real commits to all 3 repos
//   node scripts/test-calllist-sync.js --force      # Skip xlsx-hash short-circuit
//   node scripts/test-calllist-sync.js --no-db      # Skip Postgres audit trail

require('dotenv').config();
const { Pool } = require('pg');
const { runSync } = require('../lib/calllist-job');

const args = process.argv.slice(2);
const dryRun = !args.includes('--commit');
const force = args.includes('--force');
const useDb = !args.includes('--no-db');

(async () => {
  let pool = null;
  if (useDb && process.env.DATABASE_URL) {
    pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      ssl: process.env.DATABASE_URL.includes('localhost') ? false : { rejectUnauthorized: false },
    });
  }

  console.log(`Mode: ${dryRun ? 'DRY RUN (no GitHub writes)' : 'COMMIT (will push to all 3 repos)'}`);
  if (force) console.log('Force: skipping xlsx-hash short-circuit');
  console.log(`DB audit trail: ${pool ? 'enabled' : 'disabled'}\n`);

  try {
    const result = await runSync({ pool, dryRun, force });
    console.log('\n=== SUMMARY ===');
    console.log(JSON.stringify(
      { action: result.action, summary: result.summary, durationMs: result.durationMs, run_id: result.run_id, xlsx_size_bytes: result.xlsx_size_bytes, counts: result.counts },
      null, 2
    ));

    if (result.publish_results) {
      console.log('\n=== PER-FILE ===');
      for (const r of result.publish_results) {
        const tag = r.action === 'updated' ? '✓ UPDATED   '
                  : r.action === 'would-update' ? '~ would-update'
                  : r.action === 'unchanged' ? '  unchanged '
                  : '✗ ERROR     ';
        console.log(`  ${tag} ${r.repo} :: ${r.filePath}`);
        if (r.commitUrl) console.log(`              ${r.commitUrl}`);
        if (r.error) console.log(`              ${r.error}`);
      }
    }
    process.exit(0);
  } catch (err) {
    console.error('\n✗ SYNC FAILED:', err.message);
    if (err.stack) console.error(err.stack);
    process.exit(1);
  } finally {
    if (pool) await pool.end();
  }
})();
