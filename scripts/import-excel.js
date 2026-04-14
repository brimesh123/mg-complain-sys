'use strict';
/**
 * One-time import: reads "Demo sheet.xlsx" (Sheet3) and populates the SQLite database.
 * Run ONCE with:  npm run import
 */

const path = require('path');
const XLSX = require('xlsx');
const { DatabaseSync } = require('node:sqlite');

const EXCEL_PATH = path.join(__dirname, '..', 'Demo sheet.xlsx');
const DB_PATH    = path.join(__dirname, '..', 'data.db');

const db = new DatabaseSync(DB_PATH);
db.exec("PRAGMA foreign_keys = ON;");
db.exec("PRAGMA journal_mode = WAL;");

db.exec(`
  CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nsn INTEGER UNIQUE NOT NULL,
    osn TEXT,
    party_name TEXT NOT NULL,
    new_party_name TEXT,
    contact_no TEXT,
    address TEXT,
    area TEXT,
    install_date TEXT,
    new_date TEXT,
    status TEXT DEFAULT 'ON',
    notes TEXT,
    created_at TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')),
    updated_at TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
  );
`);

const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true });
const ws = wb.Sheets['Sheet3'];
if (!ws) { console.error('Sheet3 not found in the Excel file.'); process.exit(1); }

const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

const insert = db.prepare(`
  INSERT OR IGNORE INTO customers
    (nsn, osn, party_name, new_party_name, contact_no, address, area, install_date, new_date, status)
  VALUES (?,?,?,?,?,?,?,?,?,?)
`);

let imported = 0;
for (const r of rows) {
  const nsn = r['NSN'] || r['NEW SN'] || r['nsn'];
  if (!nsn || isNaN(Number(nsn))) continue;

  const partyName = String(r['PARTY NAME'] || '').trim();
  const newParty  = String(r['NEW NAME /NEW PARTY'] || r['NEW PARTY'] || partyName).trim();
  const osn       = String(r['OSN'] || '').trim() || null;
  const contact   = r['CONTACT NO'] ? String(r['CONTACT NO']).trim() : null;
  const address   = r['ADDRESS']    ? String(r['ADDRESS']).trim()    : null;
  const area      = r['AREA']       ? String(r['AREA']).trim()       : null;
  const status    = r['ON/OFF'] === 'OFF' ? 'OFF' : 'ON';

  let installDate = null;
  const raw = r['INSTALL DATE'];
  if (raw instanceof Date)              installDate = raw.toISOString().split('T')[0];
  else if (raw && typeof raw === 'number') {
    const d = XLSX.SSF.parse_date_code(raw);
    if (d) installDate = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  } else if (raw) {
    installDate = String(raw).trim() || null;
  }

  let newDate = null;
  const rawNew = r['NEW DATE'];
  if (rawNew instanceof Date && rawNew.getFullYear() > 1970) newDate = rawNew.toISOString().split('T')[0];

  const result = insert.run(Number(nsn), osn, partyName, newParty, contact, address, area, installDate, newDate, status);
  if (result.changes > 0) imported++;
}

console.log(`\n  Import complete.`);
console.log(`  Rows in sheet:   ${rows.length}`);
console.log(`  Customers added: ${imported}`);
console.log(`  Already existed: ${rows.length - imported}\n`);
db.close();
