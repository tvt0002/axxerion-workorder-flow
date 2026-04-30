// Orchestrator: SharePoint xlsx → parse → diff → publish to all 3 dashboard repos.
// Audit trail in Postgres `call_directory_sync_runs`.

const crypto = require('crypto');
const { fetchCallListXlsx } = require('./sharepoint-fetch');
const { parseCallList } = require('./calllist-parser');
const { publishCallList } = require('./calllist-publish');

async function ensureStateTable(pool) {
  if (!pool) return;
  await pool.query(`CREATE TABLE IF NOT EXISTS call_directory_sync_runs (
    id SERIAL PRIMARY KEY,
    run_at TIMESTAMP NOT NULL DEFAULT NOW(),
    xlsx_etag TEXT,
    xlsx_sha256 TEXT,
    xlsx_size_bytes INT,
    stores_count INT,
    corp_sections_count INT,
    people_count INT,
    publish_summary JSONB,
    files_updated INT NOT NULL DEFAULT 0,
    files_unchanged INT NOT NULL DEFAULT 0,
    files_errored INT NOT NULL DEFAULT 0,
    skipped_reason TEXT,
    error TEXT
  )`);
}

async function getLastSuccessfulRunHash(pool) {
  if (!pool) return null;
  const r = await pool.query(
    `SELECT xlsx_sha256 FROM call_directory_sync_runs
     WHERE error IS NULL AND xlsx_sha256 IS NOT NULL
     ORDER BY id DESC LIMIT 1`
  );
  return r.rows[0]?.xlsx_sha256 || null;
}

async function recordRun(pool, row) {
  if (!pool) return null;
  const r = await pool.query(
    `INSERT INTO call_directory_sync_runs (
      xlsx_etag, xlsx_sha256, xlsx_size_bytes,
      stores_count, corp_sections_count, people_count,
      publish_summary, files_updated, files_unchanged, files_errored,
      skipped_reason, error
    ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) RETURNING id, run_at`,
    [
      row.xlsx_etag || null,
      row.xlsx_sha256 || null,
      row.xlsx_size_bytes || null,
      row.stores_count ?? null,
      row.corp_sections_count ?? null,
      row.people_count ?? null,
      row.publish_summary ? JSON.stringify(row.publish_summary) : null,
      row.files_updated || 0,
      row.files_unchanged || 0,
      row.files_errored || 0,
      row.skipped_reason || null,
      row.error || null,
    ]
  );
  return r.rows[0];
}

function summarize(results) {
  const updated = results.filter(r => r.action === 'updated').length;
  const wouldUpdate = results.filter(r => r.action === 'would-update').length;
  const unchanged = results.filter(r => r.action === 'unchanged').length;
  const errored = results.filter(r => r.action === 'error').length;
  return { updated, wouldUpdate, unchanged, errored };
}

// Run one full sync cycle.
// opts: { pool?, dryRun?, force? }
//   - pool: pg Pool — if omitted, no audit trail
//   - dryRun: don't write to GitHub, just compute diffs
//   - force: skip xlsx-hash short-circuit, always parse + publish
async function runSync(opts = {}) {
  const { pool, dryRun = false, force = false } = opts;
  const startedAt = Date.now();
  const log = (...args) => console.log('[calldir-sync]', ...args);

  await ensureStateTable(pool);

  let xlsxResult;
  try {
    log('fetching xlsx from SharePoint…');
    xlsxResult = await fetchCallListXlsx();
  } catch (err) {
    log('FETCH FAILED:', err.message);
    if (pool && !dryRun) {
      await recordRun(pool, { error: `fetch: ${err.message}` });
    }
    throw err;
  }

  const xlsxSha256 = crypto.createHash('sha256').update(xlsxResult.buffer).digest('hex');
  log(`fetched ${xlsxResult.sizeBytes} bytes, sha256=${xlsxSha256.slice(0, 12)}…, lastModified=${xlsxResult.lastModifiedDateTime}`);

  // Short-circuit: if xlsx hash matches last successful run, skip parsing+publishing.
  if (!force && pool) {
    const lastHash = await getLastSuccessfulRunHash(pool);
    if (lastHash === xlsxSha256) {
      log('xlsx unchanged since last successful run — skipping');
      if (!dryRun) {
        await recordRun(pool, {
          xlsx_etag: xlsxResult.eTag,
          xlsx_sha256: xlsxSha256,
          xlsx_size_bytes: xlsxResult.sizeBytes,
          skipped_reason: 'xlsx_hash_unchanged',
        });
      }
      return {
        action: 'skipped',
        reason: 'xlsx_hash_unchanged',
        xlsx_sha256: xlsxSha256,
        durationMs: Date.now() - startedAt,
      };
    }
  }

  // Parse + publish.
  let parsed;
  try {
    parsed = parseCallList(xlsxResult.buffer);
    log(`parsed: ${parsed.stores.length} stores, ${Object.keys(parsed.corporate).length} corp sections, ${parsed.people.length} people`);
  } catch (err) {
    log('PARSE FAILED:', err.message);
    if (pool && !dryRun) {
      await recordRun(pool, {
        xlsx_etag: xlsxResult.eTag,
        xlsx_sha256: xlsxSha256,
        xlsx_size_bytes: xlsxResult.sizeBytes,
        error: `parse: ${err.message}`,
      });
    }
    throw err;
  }

  let publishResults;
  try {
    log(`publishing to ${dryRun ? 'GitHub (DRY RUN)' : 'GitHub'}…`);
    publishResults = await publishCallList(parsed, { dryRun });
  } catch (err) {
    log('PUBLISH FAILED:', err.message);
    if (pool && !dryRun) {
      await recordRun(pool, {
        xlsx_etag: xlsxResult.eTag,
        xlsx_sha256: xlsxSha256,
        xlsx_size_bytes: xlsxResult.sizeBytes,
        stores_count: parsed.stores.length,
        corp_sections_count: Object.keys(parsed.corporate).length,
        people_count: parsed.people.length,
        error: `publish: ${err.message}`,
      });
    }
    throw err;
  }

  const counts = summarize(publishResults);
  log(`publish results: updated=${counts.updated} would-update=${counts.wouldUpdate} unchanged=${counts.unchanged} errored=${counts.errored}`);

  let runRecord = null;
  if (pool && !dryRun) {
    runRecord = await recordRun(pool, {
      xlsx_etag: xlsxResult.eTag,
      xlsx_sha256: xlsxSha256,
      xlsx_size_bytes: xlsxResult.sizeBytes,
      stores_count: parsed.stores.length,
      corp_sections_count: Object.keys(parsed.corporate).length,
      people_count: parsed.people.length,
      publish_summary: publishResults,
      files_updated: counts.updated,
      files_unchanged: counts.unchanged,
      files_errored: counts.errored,
    });
  }

  return {
    action: dryRun ? 'dry-run' : 'completed',
    xlsx_sha256: xlsxSha256,
    xlsx_size_bytes: xlsxResult.sizeBytes,
    counts: parsed.stores.length === 0 ? null : {
      stores: parsed.stores.length,
      corp_sections: Object.keys(parsed.corporate).length,
      people: parsed.people.length,
    },
    publish_results: publishResults,
    summary: counts,
    durationMs: Date.now() - startedAt,
    run_id: runRecord?.id,
  };
}

module.exports = { runSync, ensureStateTable };
