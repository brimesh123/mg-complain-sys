'use strict';

// ─── Prevent stray errors (e.g. EBUSY from whatsapp-web.js cleanup) from crashing server ──
process.on('uncaughtException',     err => console.error('[Process] Uncaught exception:', err.message));
process.on('unhandledRejection',    err => console.error('[Process] Unhandled rejection:', err?.message || err));

const express = require('express');
const { DatabaseSync } = require('node:sqlite');
const path    = require('path');
const wa      = require('./services/whatsapp');
const fs      = require('fs');

const app = express();
const PORT = process.env.PORT || 3050;

// ─── Data directory (writable persistent disk on cloud, local folder on PC) ──
const DATA_DIR = process.env.DATA_DIR || __dirname;
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ─── Database Setup ───────────────────────────────────────────────────────────
const db = new DatabaseSync(path.join(DATA_DIR, 'data.db'));
db.exec("PRAGMA journal_mode = WAL;");
db.exec("PRAGMA foreign_keys = ON;");

db.exec(`
  CREATE TABLE IF NOT EXISTS customers (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    nsn         INTEGER UNIQUE NOT NULL,
    osn         TEXT,
    party_name  TEXT NOT NULL,
    new_party_name TEXT,
    contact_no  TEXT,
    address     TEXT,
    area        TEXT,
    install_date TEXT,
    new_date    TEXT,
    status      TEXT DEFAULT 'ON' CHECK(status IN ('ON','OFF')),
    notes       TEXT,
    created_at  TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')),
    updated_at  TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
  );

  CREATE TABLE IF NOT EXISTS engineers (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL,
    contact    TEXT,
    active     INTEGER DEFAULT 1,
    created_at TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
  );

  CREATE TABLE IF NOT EXISTS complaints (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    complaint_no   TEXT UNIQUE NOT NULL,
    customer_id    INTEGER NOT NULL REFERENCES customers(id) ON DELETE RESTRICT,
    complaint_type TEXT NOT NULL,
    remarks        TEXT,
    engineer_id    INTEGER REFERENCES engineers(id) ON DELETE SET NULL,
    status         TEXT DEFAULT 'Open'   CHECK(status IN ('Open','In Progress','Resolved','Closed')),
    priority       TEXT DEFAULT 'Normal' CHECK(priority IN ('Low','Normal','High','Urgent')),
    created_at     TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')),
    resolved_at    TEXT,
    updated_at     TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
  );

  CREATE TABLE IF NOT EXISTS complaint_types (
    id     INTEGER PRIMARY KEY AUTOINCREMENT,
    name   TEXT UNIQUE NOT NULL,
    active INTEGER DEFAULT 1
  );

  INSERT OR IGNORE INTO complaint_types (name) VALUES
    ('No Signal'),
    ('Low / Weak Signal'),
    ('Internet Not Working'),
    ('Slow Speed'),
    ('Set Top Box Issue'),
    ('Cable Damage / Cut'),
    ('Payment Issue'),
    ('New Connection Request'),
    ('Relocation Request'),
    ('Device Not Working'),
    ('Other');

  CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
  );

  CREATE TABLE IF NOT EXISTS logs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    event_type  TEXT NOT NULL,
    description TEXT NOT NULL,
    created_at  TEXT DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
  );
`);

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization');
  if (req.method === 'OPTIONS') { res.sendStatus(204); return; }
  next();
});
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Helpers ──────────────────────────────────────────────────────────────────
function ok(res, data, meta = {}) {
  res.json({ success: true, data, ...meta });
}

function fail(res, msg, code = 400) {
  res.status(code).json({ success: false, message: msg });
}

function run(sql, ...params) {
  return db.prepare(sql).run(...params);
}
function get(sql, ...params) {
  return db.prepare(sql).get(...params);
}
function all(sql, ...params) {
  return db.prepare(sql).all(...params);
}

function logEvent(type, desc) {
  try { run(`INSERT INTO logs (event_type,description) VALUES (?,?)`, type, desc); }
  catch(e) { console.error('[log]', e.message); }
}

function nextComplaintNo() {
  const year = new Date().getFullYear();
  const prefix = `CMP-${year}-`;
  const row = get(
    `SELECT complaint_no FROM complaints WHERE complaint_no LIKE ? ORDER BY id DESC LIMIT 1`,
    prefix + '%'
  );
  if (!row) return prefix + '0001';
  const last = parseInt(row.complaint_no.split('-')[2], 10);
  return prefix + String(last + 1).padStart(4, '0');
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
app.get('/api/dashboard/stats', (_req, res) => {
  try {
    const stats = get(`
      SELECT
        (SELECT COUNT(*) FROM customers)                                               AS total_customers,
        (SELECT COUNT(*) FROM customers WHERE status='ON')                             AS active_customers,
        (SELECT COUNT(*) FROM customers WHERE status='OFF')                            AS inactive_customers,
        (SELECT COUNT(*) FROM complaints)                                              AS total_complaints,
        (SELECT COUNT(*) FROM complaints WHERE status='Open')                          AS open_complaints,
        (SELECT COUNT(*) FROM complaints WHERE status='In Progress')                   AS inprogress_complaints,
        (SELECT COUNT(*) FROM complaints WHERE status IN ('Resolved','Closed'))        AS resolved_complaints,
        (SELECT COUNT(*) FROM complaints WHERE DATE(created_at,'localtime')=DATE('now','localtime')) AS today_complaints,
        (SELECT COUNT(*) FROM complaints WHERE status IN ('Resolved','Closed')
           AND DATE(resolved_at,'localtime')=DATE('now','localtime'))                  AS resolved_today
    `);

    const recent = all(`
      SELECT c.id, c.complaint_no, c.complaint_type, c.status, c.priority,
             c.created_at, cu.new_party_name, cu.nsn, cu.area, e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id = cu.id
      LEFT JOIN engineers e ON c.engineer_id = e.id
      ORDER BY c.created_at DESC LIMIT 8
    `);

    const byEngineer = all(`
      SELECT e.name, COUNT(c.id) AS total,
             SUM(CASE WHEN c.status='Open' THEN 1 ELSE 0 END) AS open
      FROM complaints c
      JOIN engineers e ON c.engineer_id = e.id
      WHERE c.status IN ('Open','In Progress')
      GROUP BY e.id ORDER BY total DESC LIMIT 6
    `);

    ok(res, { stats, recent, byEngineer });
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Customer Lookup ──────────────────────────────────────────────────────────
app.get('/api/customers/lookup', (req, res) => {
  try {
    const q = (req.query.q || '').trim();
    if (!q) return ok(res, []);
    const rows = all(`
      SELECT * FROM customers
      WHERE CAST(nsn AS TEXT) LIKE ?
         OR LOWER(osn) LIKE LOWER(?)
         OR LOWER(new_party_name) LIKE LOWER(?)
         OR LOWER(party_name) LIKE LOWER(?)
      ORDER BY nsn ASC LIMIT 10
    `, `${q}%`, `%${q}%`, `%${q}%`, `%${q}%`);
    ok(res, rows);
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Customers CRUD ───────────────────────────────────────────────────────────
app.get('/api/customers', (req, res) => {
  try {
    const { search = '', area = '', status = '', page = 1, limit = 50 } = req.query;
    const offset = (parseInt(page) - 1) * parseInt(limit);
    const like = `%${search}%`;

    let where = `WHERE (CAST(c.nsn AS TEXT) LIKE ? OR LOWER(c.new_party_name) LIKE LOWER(?) OR LOWER(c.osn) LIKE LOWER(?))`;
    const params = [like, like, like];
    if (area)   { where += ` AND c.area=?`;   params.push(area); }
    if (status) { where += ` AND c.status=?`; params.push(status); }

    const total = get(`SELECT COUNT(*) AS n FROM customers c ${where}`, ...params).n;
    const rows  = all(`
      SELECT c.*,
             (SELECT COUNT(*) FROM complaints WHERE customer_id=c.id) AS complaint_count,
             (SELECT COUNT(*) FROM complaints WHERE customer_id=c.id AND status IN ('Open','In Progress')) AS open_count
      FROM customers c ${where}
      ORDER BY c.nsn ASC LIMIT ? OFFSET ?
    `, ...params, parseInt(limit), offset);

    ok(res, rows, { total, page: parseInt(page), limit: parseInt(limit) });
  } catch(e) { fail(res, e.message, 500); }
});

app.get('/api/customers/:id', (req, res) => {
  try {
    const row = get(`SELECT * FROM customers WHERE id=?`, req.params.id);
    if (!row) return fail(res, 'Customer not found', 404);
    ok(res, row);
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/customers', (req, res) => {
  try {
    const { nsn, osn, party_name, new_party_name, contact_no, address, area, install_date, new_date, status, notes } = req.body;
    if (!nsn || !party_name) return fail(res, 'NSN and Party Name are required');
    const result = run(`
      INSERT INTO customers (nsn,osn,party_name,new_party_name,contact_no,address,area,install_date,new_date,status,notes)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)
    `, nsn, osn||null, party_name, new_party_name||party_name, contact_no||null, address||null, area||null, install_date||null, new_date||null, status||'ON', notes||null);
    logEvent('customer-add', `Customer added — ${new_party_name||party_name} (NSN: ${nsn})`);
    scheduleSyncWrite();
    ok(res, get(`SELECT * FROM customers WHERE id=?`, result.lastInsertRowid));
  } catch(e) {
    if (e.message.includes('UNIQUE')) return fail(res, `Serial number ${req.body.nsn} already exists`);
    fail(res, e.message, 500);
  }
});

app.put('/api/customers/:id', (req, res) => {
  try {
    const existing = get(`SELECT id FROM customers WHERE id=?`, req.params.id);
    if (!existing) return fail(res, 'Customer not found', 404);
    const { nsn, osn, party_name, new_party_name, contact_no, address, area, install_date, new_date, status, notes } = req.body;
    run(`
      UPDATE customers SET nsn=?,osn=?,party_name=?,new_party_name=?,contact_no=?,address=?,
        area=?,install_date=?,new_date=?,status=?,notes=?,
        updated_at=strftime('%Y-%m-%dT%H:%M:%SZ','now')
      WHERE id=?
    `, nsn, osn||null, party_name, new_party_name||party_name, contact_no||null, address||null, area||null, install_date||null, new_date||null, status||'ON', notes||null, req.params.id);
    const updCust = get(`SELECT * FROM customers WHERE id=?`, req.params.id);
    logEvent('customer-update', `Customer updated — ${updCust.new_party_name} (NSN: ${updCust.nsn})`);
    scheduleSyncWrite();
    ok(res, updCust);
  } catch(e) {
    if (e.message.includes('UNIQUE')) return fail(res, `Serial number ${req.body.nsn} already exists`);
    fail(res, e.message, 500);
  }
});

app.delete('/api/customers/:id', (req, res) => {
  try {
    const row = get(`SELECT id, nsn, COALESCE(new_party_name,party_name) AS name FROM customers WHERE id=?`, req.params.id);
    if (!row) return fail(res, 'Customer not found', 404);
    const linked = get(`SELECT COUNT(*) AS n FROM complaints WHERE customer_id=?`, req.params.id).n;
    if (linked > 0) return fail(res, `Cannot delete: ${linked} complaint(s) linked to this customer`);
    run(`DELETE FROM customers WHERE id=?`, req.params.id);
    logEvent('customer-delete', `Customer deleted — ${row.name} (NSN: ${row.nsn})`);
    scheduleSyncWrite();
    ok(res, { id: parseInt(req.params.id) });
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Complaints CRUD ──────────────────────────────────────────────────────────
const complaintCols = `
  c.id, c.complaint_no, c.complaint_type, c.status, c.priority,
  c.remarks, c.created_at, c.resolved_at, c.updated_at,
  cu.id AS customer_id, cu.nsn, cu.osn, cu.new_party_name, cu.contact_no,
  cu.address, cu.area, cu.status AS connection_status,
  e.id AS engineer_id, e.name AS engineer_name
`;

app.get('/api/complaints', (req, res) => {
  try {
    const { status = '', engineer_id = '', date_from = '', date_to = '', search = '', page = 1, limit = 50 } = req.query;
    const offset = (parseInt(page) - 1) * parseInt(limit);
    const conds = [], params = [];

    if (status)      { conds.push(`c.status=?`);                    params.push(status); }
    if (engineer_id) { conds.push(`c.engineer_id=?`);               params.push(engineer_id); }
    if (date_from)   { conds.push(`DATE(c.created_at)>=?`);          params.push(date_from); }
    if (date_to)     { conds.push(`DATE(c.created_at)<=?`);          params.push(date_to); }
    if (search)      {
      conds.push(`(c.complaint_no LIKE ? OR LOWER(cu.new_party_name) LIKE LOWER(?) OR CAST(cu.nsn AS TEXT) LIKE ?)`);
      params.push(`%${search}%`, `%${search}%`, `%${search}%`);
    }

    const join  = `FROM complaints c JOIN customers cu ON c.customer_id=cu.id LEFT JOIN engineers e ON c.engineer_id=e.id`;
    const where = conds.length ? 'WHERE ' + conds.join(' AND ') : '';

    const total = get(`SELECT COUNT(*) AS n ${join} ${where}`, ...params).n;
    const rows  = all(`SELECT ${complaintCols} ${join} ${where} ORDER BY c.created_at DESC LIMIT ? OFFSET ?`, ...params, parseInt(limit), offset);

    ok(res, rows, { total, page: parseInt(page), limit: parseInt(limit) });
  } catch(e) { fail(res, e.message, 500); }
});

app.get('/api/complaints/:id', (req, res) => {
  try {
    const row = get(`
      SELECT c.*, cu.nsn, cu.osn, cu.new_party_name, cu.contact_no,
             cu.address, cu.area, cu.status AS connection_status, cu.install_date,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id=cu.id
      LEFT JOIN engineers e ON c.engineer_id=e.id
      WHERE c.id=?
    `, req.params.id);
    if (!row) return fail(res, 'Complaint not found', 404);
    ok(res, row);
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/complaints', (req, res) => {
  try {
    const { customer_id, remarks, engineer_id, priority } = req.body;
    const complaint_type = req.body.complaint_type || 'General';
    if (!customer_id) return fail(res, 'Customer is required');
    if (!get(`SELECT id FROM customers WHERE id=?`, customer_id)) return fail(res, 'Customer not found');

    const complaint_no = nextComplaintNo();
    const result = run(`
      INSERT INTO complaints (complaint_no, customer_id, complaint_type, remarks, engineer_id, priority)
      VALUES (?,?,?,?,?,?)
    `, complaint_no, customer_id, complaint_type, remarks||null, engineer_id||null, priority||'Normal');

    const row = get(`
      SELECT c.*, cu.nsn, cu.osn, cu.new_party_name, cu.contact_no, cu.address, cu.area,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id=cu.id
      LEFT JOIN engineers e ON c.engineer_id=e.id
      WHERE c.id=?
    `, result.lastInsertRowid);
    logEvent('complaint-new', `Complaint ${row.complaint_no} logged — ${row.new_party_name} · ${row.complaint_type}`);
    scheduleSyncWrite();
    ok(res, row);
  } catch(e) { fail(res, e.message, 500); }
});

app.put('/api/complaints/:id', (req, res) => {
  try {
    const existing = get(`SELECT * FROM complaints WHERE id=?`, req.params.id);
    if (!existing) return fail(res, 'Complaint not found', 404);

    const { complaint_type, remarks, engineer_id, status, priority } = req.body;
    const resolved_at = (status === 'Resolved' || status === 'Closed') && !existing.resolved_at
      ? new Date().toISOString()
      : existing.resolved_at;

    run(`
      UPDATE complaints SET complaint_type=?, remarks=?, engineer_id=?, status=?, priority=?,
        resolved_at=?, updated_at=strftime('%Y-%m-%dT%H:%M:%SZ','now')
      WHERE id=?
    `, complaint_type, remarks||null, engineer_id||null, status, priority||'Normal', resolved_at, req.params.id);

    const row = get(`
      SELECT c.*, cu.nsn, cu.osn, cu.new_party_name, cu.contact_no, cu.address, cu.area,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id=cu.id
      LEFT JOIN engineers e ON c.engineer_id=e.id
      WHERE c.id=?
    `, req.params.id);
    logEvent('complaint-update', `Complaint ${row.complaint_no} updated — ${row.status} · ${row.complaint_type}`);
    scheduleSyncWrite();
    ok(res, row);
  } catch(e) { fail(res, e.message, 500); }
});

app.delete('/api/complaints/:id', (req, res) => {
  try {
    const cRow = get(`SELECT id, complaint_no FROM complaints WHERE id=?`, req.params.id);
    if (!cRow) return fail(res, 'Complaint not found', 404);
    run(`DELETE FROM complaints WHERE id=?`, req.params.id);
    logEvent('complaint-delete', `Complaint ${cRow.complaint_no} deleted`);
    scheduleSyncWrite();
    ok(res, { id: parseInt(req.params.id) });
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Engineers CRUD ───────────────────────────────────────────────────────────
app.get('/api/engineers', (_req, res) => {
  try {
    const rows = all(`
      SELECT e.*,
             COUNT(c.id) AS total_complaints,
             SUM(CASE WHEN c.status IN ('Open','In Progress') THEN 1 ELSE 0 END) AS open_complaints
      FROM engineers e
      LEFT JOIN complaints c ON c.engineer_id=e.id
      GROUP BY e.id ORDER BY e.name
    `);
    ok(res, rows);
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/engineers', (req, res) => {
  try {
    const { name, contact } = req.body;
    if (!name) return fail(res, 'Engineer name is required');
    const result = run(`INSERT INTO engineers (name, contact) VALUES (?,?)`, name.trim(), contact||null);
    logEvent('engineer-add', `Engineer added — ${name.trim()}`);
    scheduleSyncWrite();
    ok(res, get(`SELECT * FROM engineers WHERE id=?`, result.lastInsertRowid));
  } catch(e) { fail(res, e.message, 500); }
});

app.put('/api/engineers/:id', (req, res) => {
  try {
    if (!get(`SELECT id FROM engineers WHERE id=?`, req.params.id)) return fail(res, 'Engineer not found', 404);
    const { name, contact, active } = req.body;
    run(`UPDATE engineers SET name=?,contact=?,active=? WHERE id=?`, name.trim(), contact||null, active !== undefined ? active : 1, req.params.id);
    logEvent('engineer-update', `Engineer updated — ${name.trim()} (${active === 0 ? 'Inactive' : 'Active'})`);
    scheduleSyncWrite();
    ok(res, get(`SELECT * FROM engineers WHERE id=?`, req.params.id));
  } catch(e) { fail(res, e.message, 500); }
});

app.delete('/api/engineers/:id', (req, res) => {
  try {
    const engRow = get(`SELECT id, name FROM engineers WHERE id=?`, req.params.id);
    if (!engRow) return fail(res, 'Engineer not found', 404);
    run(`UPDATE engineers SET active=0 WHERE id=?`, req.params.id);
    logEvent('engineer-update', `Engineer deactivated — ${engRow.name}`);
    scheduleSyncWrite();
    ok(res, { id: parseInt(req.params.id) });
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Reference Data ───────────────────────────────────────────────────────────
app.get('/api/areas', (_req, res) => {
  try {
    const rows = all(`SELECT DISTINCT area FROM customers WHERE area IS NOT NULL ORDER BY area`);
    ok(res, rows.map(r => r.area));
  } catch(e) { fail(res, e.message, 500); }
});

app.get('/api/complaint-types', (_req, res) => {
  try {
    ok(res, all(`SELECT * FROM complaint_types WHERE active=1 ORDER BY id`));
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/complaint-types', (req, res) => {
  try {
    const { name } = req.body;
    if (!name) return fail(res, 'Type name is required');
    const result = run(`INSERT INTO complaint_types (name) VALUES (?)`, name.trim());
    ok(res, get(`SELECT * FROM complaint_types WHERE id=?`, result.lastInsertRowid));
  } catch(e) {
    if (e.message.includes('UNIQUE')) return fail(res, 'Type already exists');
    fail(res, e.message, 500);
  }
});

app.get('/api/complaints/stats', (_req, res) => {
  try {
    const row = get(`
      SELECT
        (SELECT COUNT(*) FROM complaints
           WHERE DATE(created_at,'localtime')=DATE('now','localtime'))                         AS today,
        (SELECT COUNT(*) FROM complaints
           WHERE DATE(created_at,'localtime')=DATE('now','localtime','-1 day'))                AS yesterday,
        (SELECT COUNT(*) FROM complaints
           WHERE strftime('%Y-%m',created_at,'localtime')=strftime('%Y-%m','now','localtime')) AS this_month,
        (SELECT COUNT(*) FROM complaints)                                                       AS total
    `);
    ok(res, row);
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Reports ──────────────────────────────────────────────────────────────────
app.get('/api/reports/complaints', (req, res) => {
  try {
    const { search, nsn, date_from, date_to, status } = req.query;
    const conditions = [];
    const params = [];
    if (search)    { conditions.push(`(cu.new_party_name LIKE ? OR cu.party_name LIKE ?)`); params.push(`%${search}%`, `%${search}%`); }
    if (nsn)       { conditions.push(`cu.nsn = ?`);                                         params.push(nsn); }
    if (date_from) { conditions.push(`DATE(c.created_at,'localtime') >= ?`);                params.push(date_from); }
    if (date_to)   { conditions.push(`DATE(c.created_at,'localtime') <= ?`);                params.push(date_to); }
    if (status)    { conditions.push(`c.status = ?`);                                       params.push(status); }
    const where = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';
    const rows = all(`
      SELECT c.id, c.complaint_no, c.complaint_type, c.remarks, c.status, c.priority,
             c.created_at, c.resolved_at,
             cu.nsn, cu.new_party_name, cu.contact_no, cu.address, cu.area,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id = cu.id
      LEFT JOIN engineers e ON c.engineer_id = e.id
      ${where}
      ORDER BY c.created_at DESC
      LIMIT 500
    `, ...params);
    ok(res, rows);
  } catch(e) { fail(res, e.message, 500); }
});

app.get('/api/customers/:id/complaints', (req, res) => {
  try {
    const rows = all(`
      SELECT c.id, c.complaint_no, c.complaint_type, c.status, c.priority,
             c.created_at, c.resolved_at, e.name AS engineer_name
      FROM complaints c
      LEFT JOIN engineers e ON c.engineer_id=e.id
      WHERE c.customer_id=? ORDER BY c.created_at DESC
    `, req.params.id);
    ok(res, rows);
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Excel Export / Import ───────────────────────────────────────────────────
const XLSX = require('xlsx');

// Export customers → XLSX (matches original Sheet3 format)
app.get('/api/export/customers', (_req, res) => {
  try {
    const rows = all(`SELECT * FROM customers ORDER BY nsn`);
    const data = rows.map(r => ({
      'NSN':                   r.nsn,
      'OSN':                   r.osn          || '',
      'PARTY NAME':            r.party_name,
      'NEW NAME / NEW PARTY':  r.new_party_name || r.party_name,
      'CONTACT NO':            r.contact_no   || '',
      'ADDRESS':               r.address      || '',
      'AREA':                  r.area         || '',
      'INSTALL DATE':          r.install_date || '',
      'ON/OFF':                r.status,
      'NOTES':                 r.notes        || '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [8,8,28,28,14,40,14,14,8,20].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Customers');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename="customers_${_today()}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch(e) { fail(res, e.message, 500); }
});

// Export complaints → XLSX
app.get('/api/export/complaints', (req, res) => {
  try {
    const { date_from = '', date_to = '', status = '' } = req.query;
    const conds = [], params = [];
    if (status)    { conds.push(`c.status=?`);           params.push(status); }
    if (date_from) { conds.push(`DATE(c.created_at)>=?`); params.push(date_from); }
    if (date_to)   { conds.push(`DATE(c.created_at)<=?`); params.push(date_to); }
    const where = conds.length ? 'WHERE ' + conds.join(' AND ') : '';

    const rows = all(`
      SELECT c.complaint_no, c.complaint_type, c.status, c.priority, c.remarks,
             c.created_at, c.resolved_at,
             cu.nsn, cu.osn, cu.new_party_name, cu.contact_no, cu.area, cu.address,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id = cu.id
      LEFT JOIN engineers e ON c.engineer_id = e.id
      ${where}
      ORDER BY c.created_at DESC
    `, ...params);

    const data = rows.map(r => ({
      'Complaint No':   r.complaint_no,
      'NSN':            r.nsn,
      'OSN':            r.osn            || '',
      'Customer':       r.new_party_name,
      'Contact':        r.contact_no     || '',
      'Area':           r.area           || '',
      'Address':        r.address        || '',
      'Complaint Type': r.complaint_type,
      'Priority':       r.priority,
      'Status':         r.status,
      'Engineer':       r.engineer_name  || '',
      'Remarks':        r.remarks        || '',
      'Logged On':      r.created_at  ? _fmtDate(r.created_at)  : '',
      'Resolved On':    r.resolved_at ? _fmtDate(r.resolved_at) : '',
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [16,7,7,28,13,12,36,22,8,12,18,30,18,18].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Complaints');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', `attachment; filename="complaints_${_today()}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch(e) { fail(res, e.message, 500); }
});

// Import customers from uploaded XLSX (base64 JSON body)
app.post('/api/import/customers', (req, res) => {
  try {
    const { data: b64 } = req.body;
    if (!b64) return fail(res, 'No file data received');
    const buf = Buffer.from(b64, 'base64');
    const wb  = XLSX.read(buf, { cellDates: true });
    // Accept first sheet or named Customer/Sheet3
    const sheetName = wb.SheetNames.find(n => /customer|sheet3/i.test(n)) || wb.SheetNames[0];
    const ws   = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

    const stmt = db.prepare(`
      INSERT OR IGNORE INTO customers
        (nsn, osn, party_name, new_party_name, contact_no, address, area, install_date, status)
      VALUES (?,?,?,?,?,?,?,?,?)
    `);

    let added = 0, skipped = 0;
    for (const r of rows) {
      const nsn = r['NSN'] || r['NEW SN'] || r['nsn'];
      if (!nsn || isNaN(Number(nsn))) { skipped++; continue; }
      const pname  = String(r['PARTY NAME'] || r['NEW NAME / NEW PARTY'] || '').trim();
      if (!pname) { skipped++; continue; }
      const npname = String(r['NEW NAME / NEW PARTY'] || r['NEW PARTY'] || pname).trim();
      const osn    = String(r['OSN'] || '').trim() || null;
      const contact= r['CONTACT NO'] ? String(r['CONTACT NO']).trim() : null;
      const addr   = r['ADDRESS']    ? String(r['ADDRESS']).trim()    : null;
      const area   = r['AREA']       ? String(r['AREA']).trim()       : null;
      const status = r['ON/OFF'] === 'OFF' ? 'OFF' : 'ON';
      let idate = null;
      const raw = r['INSTALL DATE'];
      if (raw instanceof Date)             idate = raw.toISOString().split('T')[0];
      else if (raw && typeof raw === 'number') {
        const d = XLSX.SSF.parse_date_code(raw);
        if (d) idate = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
      }
      const result = stmt.run(Number(nsn), osn, pname, npname, contact, addr, area, idate, status);
      result.changes > 0 ? added++ : skipped++;
    }
    if (added > 0) {
      logEvent('import', `Bulk import — ${added} customers added from Excel (${skipped} skipped)`);
      scheduleSyncWrite();
    }
    ok(res, { added, skipped, total: rows.length });
  } catch(e) { fail(res, e.message, 500); }
});

function _today() { return new Date().toISOString().slice(0,10); }
function _fmtDate(iso) {
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2,'0')}-${d.toLocaleString('en-IN',{month:'short'})}-${d.getFullYear()}`;
}

// ─── Sync (DB ↔ sync.xlsx live backup) ───────────────────────────────────────
const SYNC_PATH = path.join(DATA_DIR, 'sync.xlsx');
let _syncTimer   = null;
let _importTimer = null;
let _ourWrite    = false;
let _lastExport  = null;
let _lastImport  = null;

function writeSyncXlsx() {
  try {
    _ourWrite = true;
    const wb = XLSX.utils.book_new();

    // Sheet 1: Customers
    const custs = all(`SELECT * FROM customers ORDER BY nsn`);
    const wsCust = XLSX.utils.json_to_sheet(custs.map(r => ({
      'NSN':                  r.nsn,
      'OSN':                  r.osn            || '',
      'PARTY NAME':           r.party_name,
      'NEW NAME / NEW PARTY': r.new_party_name || r.party_name,
      'CONTACT NO':           r.contact_no     || '',
      'ADDRESS':              r.address        || '',
      'AREA':                 r.area           || '',
      'INSTALL DATE':         r.install_date   || '',
      'ON/OFF':               r.status,
      'NOTES':                r.notes          || '',
    })));
    wsCust['!cols'] = [8,8,28,28,14,40,14,14,8,20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsCust, 'Customers');

    // Sheet 2: Complaints (export-only — never imported back, managed by the app)
    const comps = all(`
      SELECT c.complaint_no, c.complaint_type, c.status, c.priority, c.remarks,
             c.created_at, c.resolved_at, cu.nsn, cu.new_party_name, cu.contact_no, cu.area,
             e.name AS engineer_name
      FROM complaints c
      JOIN customers cu ON c.customer_id=cu.id
      LEFT JOIN engineers e ON c.engineer_id=e.id
      ORDER BY c.created_at DESC
    `);
    const wsComp = XLSX.utils.json_to_sheet(comps.map(r => ({
      'Complaint No':   r.complaint_no,
      'NSN':            r.nsn,
      'Customer':       r.new_party_name,
      'Contact':        r.contact_no    || '',
      'Area':           r.area          || '',
      'Complaint Type': r.complaint_type,
      'Priority':       r.priority,
      'Status':         r.status,
      'Engineer':       r.engineer_name || '',
      'Remarks':        r.remarks       || '',
      'Logged On':      r.created_at    ? _fmtDate(r.created_at)  : '',
      'Resolved On':    r.resolved_at   ? _fmtDate(r.resolved_at) : '',
    })));
    wsComp['!cols'] = [16,7,28,13,12,22,8,12,18,30,18,18].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsComp, 'Complaints');

    // Sheet 3: Engineers
    const engs = all(`SELECT * FROM engineers ORDER BY name`);
    const wsEng = XLSX.utils.json_to_sheet(engs.map(r => ({
      'Name':    r.name,
      'Contact': r.contact || '',
      'Active':  r.active ? 'YES' : 'NO',
    })));
    wsEng['!cols'] = [22,14,8].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsEng, 'Engineers');

    XLSX.writeFile(wb, SYNC_PATH);
    _lastExport = new Date().toISOString();
    console.log(`  [sync] Saved → sync.xlsx  (${custs.length} customers, ${comps.length} complaints, ${engs.length} engineers)`);
  } catch(e) {
    console.error(`  [sync] Export error:`, e.message);
  } finally {
    setTimeout(() => { _ourWrite = false; }, 800);
  }
}

function scheduleSyncWrite() {
  clearTimeout(_syncTimer);
  _syncTimer = setTimeout(writeSyncXlsx, 1500);
}

function importFromXlsx() {
  try {
    if (!fs.existsSync(SYNC_PATH)) return { custChanged: 0, engChanged: 0 };
    const wb = XLSX.read(fs.readFileSync(SYNC_PATH), { cellDates: true });
    let custChanged = 0, engChanged = 0;

    // Customers — upsert by NSN (existing rows updated, new rows inserted)
    const wsCust = wb.Sheets['Customers'];
    if (wsCust) {
      const upsert = db.prepare(`
        INSERT INTO customers (nsn,osn,party_name,new_party_name,contact_no,address,area,install_date,status,notes)
        VALUES (?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(nsn) DO UPDATE SET
          osn=excluded.osn, party_name=excluded.party_name,
          new_party_name=excluded.new_party_name, contact_no=excluded.contact_no,
          address=excluded.address, area=excluded.area, install_date=excluded.install_date,
          status=excluded.status, notes=excluded.notes,
          updated_at=strftime('%Y-%m-%dT%H:%M:%SZ','now')
      `);
      for (const r of XLSX.utils.sheet_to_json(wsCust, { defval: null })) {
        const nsn = r['NSN'];
        if (!nsn || isNaN(Number(nsn))) continue;
        const pname = String(r['PARTY NAME'] || '').trim();
        if (!pname) continue;
        const npname  = String(r['NEW NAME / NEW PARTY'] || pname).trim();
        const osn     = String(r['OSN']        || '').trim() || null;
        const contact = r['CONTACT NO'] ? String(r['CONTACT NO']).trim() : null;
        const addr    = r['ADDRESS']    ? String(r['ADDRESS']).trim()    : null;
        const area    = r['AREA']       ? String(r['AREA']).trim()       : null;
        const notes   = r['NOTES']      ? String(r['NOTES']).trim()      : null;
        const status  = r['ON/OFF'] === 'OFF' ? 'OFF' : 'ON';
        let idate = null;
        const raw = r['INSTALL DATE'];
        if (raw instanceof Date) idate = raw.toISOString().split('T')[0];
        else if (raw && typeof raw === 'number') {
          const d = XLSX.SSF.parse_date_code(raw);
          if (d) idate = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
        } else if (raw && typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(raw)) {
          idate = raw;
        }
        if (upsert.run(Number(nsn), osn, pname, npname, contact, addr, area, idate, status, notes).changes > 0) custChanged++;
      }
    }

    // Engineers — match by name (case-insensitive), insert if new
    const wsEng = wb.Sheets['Engineers'];
    if (wsEng) {
      for (const r of XLSX.utils.sheet_to_json(wsEng, { defval: null })) {
        const name = String(r['Name'] || '').trim();
        if (!name) continue;
        const contact = r['Contact'] ? String(r['Contact']).trim() : null;
        const active  = String(r['Active'] || '').toUpperCase() === 'NO' ? 0 : 1;
        const existing = get(`SELECT id FROM engineers WHERE LOWER(name)=LOWER(?)`, name);
        if (existing) {
          if (run(`UPDATE engineers SET contact=?,active=? WHERE id=?`, contact, active, existing.id).changes > 0) engChanged++;
        } else {
          run(`INSERT INTO engineers (name,contact,active) VALUES (?,?,?)`, name, contact, active);
          engChanged++;
        }
      }
    }

    _lastImport = new Date().toISOString();
    console.log(`  [sync] Imported ← sync.xlsx  (${custChanged} customers, ${engChanged} engineers changed)`);
    return { custChanged, engChanged };
  } catch(e) {
    console.error(`  [sync] Import error:`, e.message);
    return { custChanged: 0, engChanged: 0, error: e.message };
  }
}

// Watch sync.xlsx for external changes (e.g. user edits it in Excel and saves)
fs.watchFile(SYNC_PATH, { interval: 3000, persistent: false }, (curr, prev) => {
  if (_ourWrite) return;
  if (curr.mtime.getTime() <= prev.mtime.getTime()) return;
  clearTimeout(_importTimer);
  _importTimer = setTimeout(() => {
    console.log('  [sync] External edit detected — importing from sync.xlsx...');
    const result = importFromXlsx();
    if ((result.custChanged + result.engChanged) > 0) setTimeout(scheduleSyncWrite, 500);
  }, 1500);
});

app.get('/api/sync/status', (_req, res) => {
  try {
    let xlsxMtime = null, xlsxSize = null;
    const exists = fs.existsSync(SYNC_PATH);
    if (exists) {
      const s = fs.statSync(SYNC_PATH);
      xlsxMtime = s.mtime.toISOString();
      xlsxSize  = s.size;
    }
    ok(res, { lastExport: _lastExport, lastImport: _lastImport, xlsxExists: exists, xlsxMtime, xlsxSize });
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/sync/import', (_req, res) => {
  try {
    if (!fs.existsSync(SYNC_PATH)) return fail(res, 'sync.xlsx not found. Save any change in the app first to create it.');
    const result = importFromXlsx();
    if ((result.custChanged + result.engChanged) > 0)
      logEvent('sync', `Synced from XLSX — ${result.custChanged} customers, ${result.engChanged} engineers updated`);
    scheduleSyncWrite();
    ok(res, { ...result, lastImport: _lastImport });
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/sync/export', (_req, res) => {
  try {
    writeSyncXlsx();
    ok(res, { exported: true, lastExport: _lastExport });
  } catch(e) { fail(res, e.message, 500); }
});

// ─── Activity Logs ────────────────────────────────────────────────────────────
app.get('/api/logs', (_req, res) => {
  try {
    const rows = all(`SELECT * FROM logs ORDER BY created_at DESC LIMIT 500`);
    ok(res, rows);
  } catch(e) { fail(res, e.message, 500); }
});

app.delete('/api/logs', (_req, res) => {
  try {
    run(`DELETE FROM logs`);
    ok(res, { cleared: true });
  } catch(e) { fail(res, e.message, 500); }
});

// Write initial sync.xlsx 3 seconds after server starts
setTimeout(writeSyncXlsx, 3000);

// ─── WhatsApp Routes ──────────────────────────────────────────────────────────

// Status & QR
app.get('/api/whatsapp/status', (_req, res) => {
  try {
    const target = get(`SELECT value FROM settings WHERE key='wa_target_id'`);
    const tname  = get(`SELECT value FROM settings WHERE key='wa_target_name'`);
    ok(res, {
      status:      wa.status,
      hasQr:       !!wa.qrUrl,
      account:     wa.info ? (wa.info.pushname || wa.info.wid?.user || null) : null,
      target_id:   target?.value || null,
      target_name: tname?.value  || null,
      lastError:   wa.lastError  || null,
    });
  } catch(e) { fail(res, e.message, 500); }
});

app.get('/api/whatsapp/qr', (_req, res) => {
  if (!wa.qrUrl) return fail(res, 'No QR code available right now. Try connecting first.', 404);
  ok(res, { qr: wa.qrUrl });
});

// Connect / disconnect
app.post('/api/whatsapp/connect', (_req, res) => {
  try {
    if (wa.status === 'ready') return ok(res, { status: wa.status });
    wa.connect();
    ok(res, { status: wa.status });
  } catch(e) { fail(res, e.message, 500); }
});

app.post('/api/whatsapp/disconnect', async (_req, res) => {
  try {
    await wa.disconnect();
    ok(res, { status: wa.status });
  } catch(e) { fail(res, e.message, 500); }
});

// Groups & contacts
app.get('/api/whatsapp/groups', async (_req, res) => {
  try {
    const groups = await wa.getGroups();
    ok(res, groups);
  } catch(e) {
    console.error('[WA] getGroups error:', e.message);
    fail(res, e.message, 400);
  }
});

app.get('/api/whatsapp/contacts', async (_req, res) => {
  try {
    ok(res, await wa.getContacts());
  } catch(e) { fail(res, e.message, 400); }
});

// Save / get default notification target (group or individual)
app.post('/api/whatsapp/target', (req, res) => {
  try {
    const { id, name } = req.body;
    if (!id) return fail(res, 'Target ID is required');
    run(`INSERT OR REPLACE INTO settings (key, value) VALUES ('wa_target_id',   ?)`, id);
    run(`INSERT OR REPLACE INTO settings (key, value) VALUES ('wa_target_name', ?)`, name || id);
    ok(res, { id, name: name || id });
  } catch(e) { fail(res, e.message, 500); }
});

// Send message
app.post('/api/whatsapp/send', async (req, res) => {
  try {
    const { chat_id, message } = req.body;
    if (!chat_id || !message) return fail(res, 'chat_id and message are required');
    await wa.sendText(chat_id, message);
    ok(res, { sent: true, chat_id });
  } catch(e) { fail(res, e.message, 400); }
});

// ─── Fallback to SPA ──────────────────────────────────────────────────────────
app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n  Mogal Complaint System running at http://localhost:${PORT}\n`);
});
