// Parses "SecureSpace Store Call List.xlsx" buffer into { stores, corporate, people }.
// Logic ported from securespace-call-directory/parse_calllist.py — keep behavior parity.

const XLSX = require('xlsx');

const STORE_HEADER_RE = /^(.+?)\s+L(\d{3})\s*$/;
const PHONE_RE = /(\(?\d{3}[)\.\-\s]?\s?\d{3}[\.\-\s]?\d{4}(?:\s*(?:ext\.?|x)\s*\d+)?)/i;
const EMAIL_RE = /[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/;

const CORP_HEADERS = new Set([
  'HR PARTNERS', 'OPS MANAGEMENT TEAM', 'TRAINING PARTNERS',
  'CUSTOMER RESOLUTION', 'TRANSITIONS TEAM', 'SECURITY TEAM',
  'FACILITIES TEAM', 'VIRTUAL OPS TEAM', 'ASSET MANAGEMENT',
  'MARKETING', 'IT SUPPORT',
]);

function clean(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().replace(/\xa0/g, ' ').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function extractPhone(s) {
  if (!s) return '';
  const m = String(s).match(PHONE_RE);
  return m ? m[1].trim() : '';
}

function extractEmail(s) {
  if (!s) return '';
  const m = String(s).match(EMAIL_RE);
  return m ? m[0].trim() : '';
}

function parseCallList(xlsxBuffer) {
  const wb = XLSX.read(xlsxBuffer, { type: 'buffer' });
  const sheetName = wb.SheetNames.includes('Sheet1') ? 'Sheet1' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Take first 6 columns of every row, cleaned.
  const rows = raw.map(r => {
    const arr = Array.isArray(r) ? r : [];
    return [0, 1, 2, 3, 4, 5].map(i => clean(arr[i]));
  });

  // Find STORES section split.
  let storesStartIdx = null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].startsWith('STORES')) {
      storesStartIdx = i;
      break;
    }
  }

  // --- Parse corporate section ---
  const corporate = {};
  let currentCorpSection = null;
  const corpRows = storesStartIdx == null ? rows : rows.slice(0, storesStartIdx);
  for (const row of corpRows) {
    const [a, b, c, d, e] = row;
    if (!row.some(x => x)) continue;
    const aUpper = a.toUpperCase();
    if (CORP_HEADERS.has(aUpper)) {
      currentCorpSection = a;
      if (!corporate[currentCorpSection]) {
        corporate[currentCorpSection] = {
          team_email: extractEmail(e) || extractEmail(b),
          members: [],
        };
      }
      continue;
    }
    if (a.startsWith('Name:')) continue;
    if (currentCorpSection && a && !a.includes('@') && !a.startsWith('SecureSpace')) {
      const person = {
        name: a,
        role: b,
        timezone: c,
        phone: extractPhone(d) || d,
        email: extractEmail(e),
      };
      if (person.name) corporate[currentCorpSection].members.push(person);
    }
  }

  // --- Parse stores section ---
  const stores = [];
  let currentStore = null;
  let i = storesStartIdx != null ? storesStartIdx + 1 : 0;
  while (i < rows.length) {
    const row = rows[i];
    const a = row[0];
    const m = a.match(STORE_HEADER_RE);
    if (m) {
      if (currentStore) stores.push(currentStore);
      currentStore = {
        code: `L${m[2]}`,
        name: m[1].trim(),
        state: '',
        timezone: '',
        address: '',
        hours_raw: '',
        store_email: '',
        store_office_phone: '',
        store_cell_phone: '',
        computers_onsite: '',
        staff: [],
        supervisors: [],
        account_rep: null,
      };
      i++;
      continue;
    }
    if (currentStore == null) {
      i++;
      continue;
    }

    const lowerA = a.toLowerCase();

    if (lowerA.startsWith('state/time zone')) {
      const val = a.includes(':') ? a.split(':').slice(1).join(':').trim() : '';
      const [stateVal, ...rest] = val.split('/');
      currentStore.state = stateVal || currentStore.state;
      const tz = rest.join('/').trim();
      currentStore.timezone = tz || currentStore.timezone;
      if (row[4]) currentStore.address = row[4];
      if (row[5]) currentStore.hours_raw = (currentStore.hours_raw + ' ' + row[5]).trim();
      i++;
      continue;
    }
    if (lowerA.startsWith('store office')) {
      currentStore.store_office_phone = extractPhone(a);
      if (row[4] && row[4].includes('@')) {
        currentStore.store_email = extractEmail(row[4]);
      } else if (row[4] && !currentStore.address) {
        currentStore.address = row[4];
      }
      if (row[5]) currentStore.hours_raw = (currentStore.hours_raw + ' ' + row[5]).trim();
      i++;
      continue;
    }
    if (lowerA.startsWith('store cell')) {
      currentStore.store_cell_phone = extractPhone(a);
      if (row[5] && row[5].includes('Computers')) {
        const parts = row[5].split(':');
        currentStore.computers_onsite = parts.length > 1 ? parts.slice(1).join(':').trim() : '';
      } else if (row[5]) {
        currentStore.hours_raw = (currentStore.hours_raw + ' ' + row[5]).trim();
      }
      i++;
      continue;
    }

    // Person row: col A = name, col B = role
    if (a && row[1]) {
      const role = row[1];
      const person = {
        name: a,
        role,
        store_code_ref: row[2],
        phone: extractPhone(row[3]) || row[3],
        email: extractEmail(row[4]),
      };
      const roleL = role.toLowerCase();
      if (roleL.includes('account rep')) {
        currentStore.account_rep = person;
      } else if (
        roleL.includes('area manager') ||
        roleL.includes('district manager') ||
        roleL.includes('dm') ||
        roleL.includes('am ') ||
        roleL.includes('sr. area') ||
        role === 'AM' ||
        role === 'DM'
      ) {
        currentStore.supervisors.push(person);
      } else {
        currentStore.staff.push(person);
      }
    }
    i++;
  }
  if (currentStore) stores.push(currentStore);

  // --- Build flat people directory ---
  const people = new Map();
  function addPerson(p, ctx) {
    const key = (p.email || p.name || '').toLowerCase();
    if (!key) return;
    let rec = people.get(key);
    if (!rec) {
      rec = {
        name: p.name,
        email: p.email || '',
        phone: p.phone || '',
        roles: [],
        stores: [],
      };
      people.set(key, rec);
    }
    rec.roles.push(p.role || '');
    if (ctx) rec.stores.push(ctx);
  }
  for (const s of stores) {
    for (const p of s.staff) addPerson(p, s.code);
    for (const p of s.supervisors) addPerson(p, s.code);
    if (s.account_rep) addPerson(s.account_rep, s.code);
  }
  for (const [section, data] of Object.entries(corporate)) {
    for (const p of data.members) addPerson(p, section);
  }

  return {
    stores,
    corporate,
    people: Array.from(people.values()),
  };
}

module.exports = { parseCallList };
