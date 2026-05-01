// Verify lib/calllist-parser.js output matches the Python parser's output.
// Compares against securespace-call-directory/data/*.json (last known-good Python output).

const fs = require('fs');
const path = require('path');
const { parseCallList } = require('../lib/calllist-parser');

const XLSX_PATH = 'C:\\Users\\Timmy\\Downloads\\SecureSpace Store Call List.xlsx';
const PY_OUTPUT_DIR = 'C:\\Users\\Timmy\\securespace-call-directory\\data';

function diffSummary(a, b, label) {
  const aStr = JSON.stringify(a, null, 2);
  const bStr = JSON.stringify(b, null, 2);
  if (aStr === bStr) {
    console.log(`  ${label}: ✓ exact match (${aStr.length} chars)`);
    return true;
  }
  console.log(`  ${label}: ✗ differs (js=${aStr.length} py=${bStr.length})`);
  // Print first diverging chunk for triage
  for (let i = 0; i < Math.min(aStr.length, bStr.length); i++) {
    if (aStr[i] !== bStr[i]) {
      const ctx = Math.max(0, i - 60);
      console.log(`    first diff at char ${i}:`);
      console.log(`      js: ${JSON.stringify(aStr.slice(ctx, i + 60))}`);
      console.log(`      py: ${JSON.stringify(bStr.slice(ctx, i + 60))}`);
      break;
    }
  }
  return false;
}

(async () => {
  console.log(`Reading xlsx: ${XLSX_PATH}`);
  const buf = fs.readFileSync(XLSX_PATH);
  const result = parseCallList(buf);

  console.log(`\nJS parse summary:`);
  console.log(`  stores:    ${result.stores.length}`);
  console.log(`  corporate: ${Object.keys(result.corporate).length} sections`);
  console.log(`  people:    ${result.people.length}`);

  const pyStores = JSON.parse(fs.readFileSync(path.join(PY_OUTPUT_DIR, 'stores.json'), 'utf8'));
  const pyCorporate = JSON.parse(fs.readFileSync(path.join(PY_OUTPUT_DIR, 'corporate.json'), 'utf8'));
  const pyPeople = JSON.parse(fs.readFileSync(path.join(PY_OUTPUT_DIR, 'people.json'), 'utf8'));

  console.log(`\nPython output summary:`);
  console.log(`  stores:    ${pyStores.length}`);
  console.log(`  corporate: ${Object.keys(pyCorporate).length} sections`);
  console.log(`  people:    ${pyPeople.length}`);

  console.log(`\nParity check:`);
  const ok1 = diffSummary(result.stores, pyStores, 'stores.json');
  const ok2 = diffSummary(result.corporate, pyCorporate, 'corporate.json');
  const ok3 = diffSummary(result.people, pyPeople, 'people.json');

  if (ok1 && ok2 && ok3) {
    console.log('\n✓ All three outputs match Python parser exactly.');
    process.exit(0);
  } else {
    console.log('\n✗ Mismatch detected.');
    process.exit(1);
  }
})();
