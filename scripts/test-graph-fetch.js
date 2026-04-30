// Smoke test: fetch the call list xlsx from SharePoint, parse it, and print summary.
// Does NOT touch Postgres or push to GitHub. Pure read.

require('dotenv').config();
const { fetchCallListXlsx } = require('../lib/sharepoint-fetch');
const { parseCallList } = require('../lib/calllist-parser');

(async () => {
  try {
    console.log('1. Acquiring token + downloading xlsx via Graph...');
    const t0 = Date.now();
    const result = await fetchCallListXlsx();
    console.log(`   ✓ downloaded ${result.sizeBytes} bytes in ${Date.now() - t0}ms`);
    console.log(`   itemId: ${result.itemId}`);
    console.log(`   eTag: ${result.eTag}`);
    console.log(`   webUrl: ${result.webUrl}`);
    console.log(`   lastModified: ${result.lastModifiedDateTime}`);

    console.log('\n2. Parsing xlsx...');
    const t1 = Date.now();
    const parsed = parseCallList(result.buffer);
    console.log(`   ✓ parsed in ${Date.now() - t1}ms`);
    console.log(`   stores:    ${parsed.stores.length}`);
    console.log(`   corporate: ${Object.keys(parsed.corporate).length} sections`);
    console.log(`   people:    ${parsed.people.length}`);

    console.log('\n✓ End-to-end fetch + parse OK');
    process.exit(0);
  } catch (err) {
    console.error('\n✗ FAILED:', err.message);
    if (err.stack) console.error(err.stack);
    process.exit(1);
  }
})();
