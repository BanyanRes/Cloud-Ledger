require('dotenv').config();
const express = require('express');
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-secret';
const DB_PATH = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'cloudledger.db');
const UPLOAD_DIR = path.resolve(path.dirname(DB_PATH), 'attachments');

const dataDir = path.dirname(DB_PATH);
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Multer config
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_'))
});
const upload = multer({ storage, limits: { fileSize: 20 * 1024 * 1024 } }); // 20MB max
const memUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'Viewer', created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS entities (
    id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL, created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    code TEXT NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL,
    subtype TEXT DEFAULT '', bank_acct INTEGER DEFAULT 0, UNIQUE(entity_id, code)
  );
  CREATE TABLE IF NOT EXISTS journal_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    entry_num INTEGER NOT NULL, date TEXT NOT NULL, memo TEXT NOT NULL,
    created_by TEXT NOT NULL, created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS journal_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT, entry_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL, debit REAL DEFAULT 0, credit REAL DEFAULT 0
  );
  CREATE TABLE IF NOT EXISTS journal_attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
    filename TEXT NOT NULL, original_name TEXT NOT NULL,
    mime_type TEXT, size INTEGER, created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS bank_transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    bank_account_code TEXT NOT NULL, date TEXT NOT NULL,
    description TEXT, amount REAL NOT NULL,
    account_code TEXT, memo TEXT,
    status TEXT DEFAULT 'pending', je_id INTEGER,
    batch_id TEXT, created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS cleared_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL, entry_id INTEGER NOT NULL,
    line_index INTEGER NOT NULL, reconciliation_id INTEGER,
    UNIQUE(entity_id, account_code, entry_id, line_index)
  );
  CREATE TABLE IF NOT EXISTS reconciliations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL, statement_date TEXT NOT NULL,
    statement_balance REAL NOT NULL, book_balance REAL NOT NULL,
    cleared_count INTEGER DEFAULT 0, completed_by TEXT NOT NULL,
    completed_at TEXT DEFAULT (datetime('now'))
  );
  CREATE INDEX IF NOT EXISTS idx_accounts_entity ON accounts(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_entity ON journal_entries(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_date ON journal_entries(entity_id, date);
  CREATE INDEX IF NOT EXISTS idx_jl_entry ON journal_lines(entry_id);
  CREATE INDEX IF NOT EXISTS idx_bt_entity ON bank_transactions(entity_id, bank_account_code);
  CREATE INDEX IF NOT EXISTS idx_ja_entry ON journal_attachments(entry_id);
`);

const userCount = db.prepare('SELECT COUNT(*) as c FROM users').get();
if (userCount.c === 0) {
  db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?, ?, ?, ?)').run('Admin', 'admin@company.com', bcrypt.hashSync('admin', 10), 'Admin');
}

const DEFAULT_COA = [
  {code:"10000",name:"Cash",type:"Asset",subtype:"Current Asset",bank:1},
  {code:"10100",name:"Operating Checking",type:"Asset",subtype:"Current Asset",bank:1},
  {code:"10200",name:"Savings Account",type:"Asset",subtype:"Current Asset",bank:1},
  {code:"11000",name:"Accounts Receivable",type:"Asset",subtype:"Current Asset",bank:0},
  {code:"12000",name:"Inventory",type:"Asset",subtype:"Current Asset",bank:0},
  {code:"13000",name:"Prepaid Expenses",type:"Asset",subtype:"Current Asset",bank:0},
  {code:"15000",name:"Property & Equipment",type:"Asset",subtype:"Fixed Asset",bank:0},
  {code:"15100",name:"Accumulated Depreciation",type:"Asset",subtype:"Fixed Asset",bank:0},
  {code:"20000",name:"Accounts Payable",type:"Liability",subtype:"Current Liability",bank:0},
  {code:"21000",name:"Accrued Liabilities",type:"Liability",subtype:"Current Liability",bank:0},
  {code:"22000",name:"Unearned Revenue",type:"Liability",subtype:"Current Liability",bank:0},
  {code:"25000",name:"Notes Payable",type:"Liability",subtype:"Long-term Liability",bank:0},
  {code:"30000",name:"Common Stock / Member's Capital",type:"Equity",subtype:"Equity",bank:0},
  {code:"31000",name:"Retained Earnings",type:"Equity",subtype:"Equity",bank:0},
  {code:"32000",name:"Additional Paid-in Capital",type:"Equity",subtype:"Equity",bank:0},
  {code:"40000",name:"Revenue",type:"Revenue",subtype:"Operating Revenue",bank:0},
  {code:"41000",name:"Service Revenue",type:"Revenue",subtype:"Operating Revenue",bank:0},
  {code:"42000",name:"Interest Income",type:"Revenue",subtype:"Other Revenue",bank:0},
  {code:"50000",name:"Cost of Goods Sold",type:"Expense",subtype:"COGS",bank:0},
  {code:"60000",name:"Salaries Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"61000",name:"Rent Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"62000",name:"Utilities Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"63000",name:"Depreciation Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"64000",name:"Insurance Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"65000",name:"Office Supplies Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"66000",name:"Marketing Expense",type:"Expense",subtype:"Operating Expense",bank:0},
  {code:"70000",name:"Interest Expense",type:"Expense",subtype:"Other Expense",bank:0},
];

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(morgan('short'));
app.use(express.json({ limit: '10mb' }));
if (process.env.NODE_ENV === 'production') app.use(express.static(path.join(__dirname, '..', 'client', 'dist')));

function auth(req, res, next) {
  const token = req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  try { req.user = jwt.verify(token, JWT_SECRET); next(); } catch { return res.status(401).json({ error: 'Invalid token' }); }
}
function requireRole(...roles) { return (req, res, next) => { if (!roles.includes(req.user.role) && req.user.role !== 'Admin') return res.status(403).json({ error: 'Forbidden' }); next(); }; }

// ═══ Auth ═══
app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: 'Email and password required' });
  const user = db.prepare('SELECT * FROM users WHERE email = ?').get(email?.toLowerCase());
  if (!user) return res.status(401).json({ error: 'No account found with email: ' + email.toLowerCase() });
  if (!bcrypt.compareSync(password, user.password_hash)) return res.status(401).json({ error: 'Incorrect password for ' + email.toLowerCase() });
  const token = jwt.sign({ id: user.id, name: user.name, email: user.email, role: user.role }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
});
app.post('/api/auth/signup', (req, res) => {
  const { name, email, password, role } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'All fields required' });
  if (password.length < 3) return res.status(400).json({ error: 'Password too short' });
  if (db.prepare('SELECT id FROM users WHERE email = ?').get(email.toLowerCase())) return res.status(400).json({ error: 'Email exists' });
  const r = db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?, ?, ?, ?)').run(name, email.toLowerCase(), bcrypt.hashSync(password, 10), ['Admin','Accountant','Viewer'].includes(role)?role:'Viewer');
  res.json({ id: r.lastInsertRowid });
});
app.get('/api/auth/me', auth, (req, res) => res.json(req.user));
app.post('/api/auth/change-password', auth, (req, res) => {
  const { current_password, new_password } = req.body;
  if (!new_password || new_password.length < 3) return res.status(400).json({ error: 'Too short' });
  const user = db.prepare('SELECT * FROM users WHERE id = ?').get(req.user.id);
  if (!bcrypt.compareSync(current_password, user.password_hash)) return res.status(400).json({ error: 'Current password incorrect' });
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(bcrypt.hashSync(new_password, 10), req.user.id);
  res.json({ success: true });
});
app.post('/api/auth/forgot-password', (req, res) => {
  const user = db.prepare('SELECT id FROM users WHERE email = ?').get(req.body.email?.toLowerCase());
  if (!user) return res.status(404).json({ error: 'No account with that email' });
  const temp = 'Reset' + Math.random().toString(36).slice(2, 8);
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(bcrypt.hashSync(temp, 10), user.id);
  res.json({ temp_password: temp });
});
app.post('/api/auth/admin-reset-password', auth, requireRole('Admin'), (req, res) => {
  const { user_id, new_password } = req.body;
  if (!new_password || new_password.length < 3) return res.status(400).json({ error: 'Too short' });
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(bcrypt.hashSync(new_password, 10), user_id);
  res.json({ success: true });
});

// ═══ Users ═══
app.get('/api/users', auth, requireRole('Admin'), (req, res) => res.json(db.prepare('SELECT id, name, email, role, created_at FROM users').all()));
app.delete('/api/users/:id', auth, requireRole('Admin'), (req, res) => { if (+req.params.id === req.user.id) return res.status(400).json({ error: 'Cannot delete self' }); db.prepare('DELETE FROM users WHERE id = ?').run(req.params.id); res.json({ success: true }); });
app.put('/api/users/:id', auth, requireRole('Admin'), (req, res) => { db.prepare('UPDATE users SET name = ?, role = ? WHERE id = ?').run(req.body.name, req.body.role, req.params.id); res.json({ success: true }); });

// ═══ Entities ═══
app.get('/api/entities', auth, (req, res) => res.json(db.prepare('SELECT * FROM entities ORDER BY code').all()));
app.post('/api/entities', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { code, name } = req.body; if (!code||!name) return res.status(400).json({ error: 'Required' });
  try { const r = db.prepare('INSERT INTO entities (code, name) VALUES (?, ?)').run(code.toUpperCase(), name); const eid = r.lastInsertRowid;
    const ins = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
    db.transaction(() => { for (const a of DEFAULT_COA) ins.run(eid, a.code, a.name, a.type, a.subtype, a.bank); })();
    res.json({ id: eid }); } catch(e) { if (e.message.includes('UNIQUE')) return res.status(400).json({ error: 'Code exists' }); throw e; }
});
app.post('/api/entities/bulk', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { entities } = req.body; if (!Array.isArray(entities)) return res.status(400).json({ error: 'Invalid' });
  const insE = db.prepare('INSERT OR IGNORE INTO entities (code, name) VALUES (?, ?)');
  const insA = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
  const created = []; db.transaction(() => { for (const e of entities) { if (!e.code||!e.name) continue; const r = insE.run(e.code.toUpperCase(), e.name);
    if (r.changes > 0) { const eid = r.lastInsertRowid; for (const a of DEFAULT_COA) insA.run(eid, a.code, a.name, a.type, a.subtype, a.bank); created.push({ id: eid, code: e.code.toUpperCase(), name: e.name }); } } })();
  res.json({ created, count: created.length });
});
app.delete('/api/entities/:id', auth, requireRole('Admin'), (req, res) => { db.prepare('DELETE FROM entities WHERE id = ?').run(req.params.id); res.json({ success: true }); });

// ═══ Accounts ═══
app.get('/api/entities/:eid/accounts', auth, (req, res) => res.json(db.prepare('SELECT * FROM accounts WHERE entity_id = ? ORDER BY code').all(req.params.eid)));
app.post('/api/entities/:eid/accounts', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { code, name, type, subtype, bank_acct } = req.body; if (!code||!name||!type) return res.status(400).json({ error: 'Required' });
  try { const r = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)').run(req.params.eid, code, name, type, subtype||'', bank_acct?1:0);
    res.json({ id: r.lastInsertRowid, code, name, type, subtype: subtype||'', bank_acct: bank_acct?1:0, entity_id: +req.params.eid }); }
  catch(e) { if (e.message.includes('UNIQUE')) return res.status(400).json({ error: 'Code exists' }); throw e; }
});
app.delete('/api/entities/:eid/accounts/:code', auth, requireRole('Admin','Accountant'), (req, res) => {
  if (db.prepare('SELECT COUNT(*) as c FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id WHERE je.entity_id=? AND jl.account_code=?').get(req.params.eid, req.params.code).c > 0)
    return res.status(400).json({ error: 'Has transactions' });
  db.prepare('DELETE FROM accounts WHERE entity_id=? AND code=?').run(req.params.eid, req.params.code); res.json({ success: true });
});

// ═══ Journal Entries ═══
app.get('/api/entities/:eid/entries', auth, (req, res) => {
  const { from, to } = req.query; let sql = 'SELECT * FROM journal_entries WHERE entity_id = ?'; const params = [req.params.eid];
  if (from) { sql += ' AND date >= ?'; params.push(from); } if (to) { sql += ' AND date <= ?'; params.push(to); }
  sql += ' ORDER BY date DESC, id DESC';
  const entries = db.prepare(sql).all(...params);
  const lineStmt = db.prepare('SELECT * FROM journal_lines WHERE entry_id = ?');
  const attachStmt = db.prepare('SELECT id, original_name, mime_type, size FROM journal_attachments WHERE entry_id = ?');
  res.json(entries.map(e => ({ ...e, lines: lineStmt.all(e.id), attachments: attachStmt.all(e.id) })));
});

app.post('/api/entities/:eid/entries', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { date, memo, lines } = req.body; if (!date||!memo||!lines||lines.length<2) return res.status(400).json({ error: 'Invalid' });
  const tDr = lines.reduce((s,l) => s+(l.debit||0), 0); const tCr = lines.reduce((s,l) => s+(l.credit||0), 0);
  if (Math.abs(tDr-tCr) > 0.005) return res.status(400).json({ error: 'Must balance' });
  const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
  const result = db.transaction(() => {
    const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)').run(req.params.eid, num, date, memo, req.user.name);
    for (const l of lines) db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(r.lastInsertRowid, l.account_code, l.debit||0, l.credit||0);
    return r.lastInsertRowid;
  })();
  res.json({ id: result, entry_num: num });
});

app.delete('/api/entities/:eid/entries/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const atts = db.prepare('SELECT filename FROM journal_attachments WHERE entry_id=?').all(req.params.id);
  atts.forEach(a => { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} });
  db.prepare('DELETE FROM journal_entries WHERE id=? AND entity_id=?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

app.put('/api/entities/:eid/entries/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { date, memo, lines } = req.body;
  if (!date || !memo || !lines || lines.length < 2) return res.status(400).json({ error: 'Invalid entry' });
  const tDr = lines.reduce((s, l) => s + (l.debit || 0), 0);
  const tCr = lines.reduce((s, l) => s + (l.credit || 0), 0);
  if (Math.abs(tDr - tCr) > 0.005) return res.status(400).json({ error: 'Must balance' });
  const entry = db.prepare('SELECT * FROM journal_entries WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!entry) return res.status(404).json({ error: 'Entry not found' });
  db.transaction(() => {
    db.prepare('UPDATE journal_entries SET date=?, memo=? WHERE id=?').run(date, memo, req.params.id);
    db.prepare('DELETE FROM journal_lines WHERE entry_id=?').run(req.params.id);
    for (const l of lines) db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(req.params.id, l.account_code, l.debit || 0, l.credit || 0);
  })();
  res.json({ success: true, entry_num: entry.entry_num });
});

// ═══ Journal Attachments ═══
app.post('/api/entities/:eid/entries/:id/attachments', auth, requireRole('Admin','Accountant'), upload.array('files', 10), (req, res) => {
  if (!req.files || req.files.length === 0) return res.status(400).json({ error: 'No files' });
  const ins = db.prepare('INSERT INTO journal_attachments (entry_id, filename, original_name, mime_type, size) VALUES (?,?,?,?,?)');
  const results = [];
  for (const f of req.files) {
    const r = ins.run(req.params.id, f.filename, f.originalname, f.mimetype, f.size);
    results.push({ id: r.lastInsertRowid, original_name: f.originalname, mime_type: f.mimetype, size: f.size });
  }
  res.json(results);
});

app.get('/api/attachments/:id/download', (req, res) => {
  // Accept token from query param (for <a> links) or header
  const token = req.query.token || req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  try { jwt.verify(token, JWT_SECRET); } catch { return res.status(401).json({ error: 'Invalid token' }); }

  const att = db.prepare('SELECT * FROM journal_attachments WHERE id=?').get(req.params.id);
  if (!att) return res.status(404).json({ error: 'Not found' });
  const filepath = path.resolve(UPLOAD_DIR, att.filename);
  if (!fs.existsSync(filepath)) return res.status(404).json({ error: 'File missing' });

  // For PDFs and images, display inline; otherwise download
  const inlineTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'];
  const disposition = inlineTypes.includes(att.mime_type) ? 'inline' : 'attachment';
  res.setHeader('Content-Disposition', disposition + '; filename="' + att.original_name + '"');
  res.setHeader('Content-Type', att.mime_type || 'application/octet-stream');
  res.sendFile(filepath, err => { if (err && !res.headersSent) res.status(500).json({ error: 'Failed to send file' }); });
});

app.delete('/api/attachments/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const att = db.prepare('SELECT * FROM journal_attachments WHERE id=?').get(req.params.id);
  if (att) { try { fs.unlinkSync(path.join(UPLOAD_DIR, att.filename)); } catch {} }
  db.prepare('DELETE FROM journal_attachments WHERE id=?').run(req.params.id);
  res.json({ success: true });
});

// ═══ Bank Transaction Upload & Coding ═══
app.get('/api/entities/:eid/bank-transactions', auth, (req, res) => {
  const { bank_account, status } = req.query;
  let sql = 'SELECT * FROM bank_transactions WHERE entity_id = ?'; const params = [req.params.eid];
  if (bank_account) { sql += ' AND bank_account_code = ?'; params.push(bank_account); }
  if (status) { sql += ' AND status = ?'; params.push(status); }
  sql += ' ORDER BY date, id';
  res.json(db.prepare(sql).all(...params));
});

app.post('/api/entities/:eid/bank-transactions/upload', auth, requireRole('Admin','Accountant'), memUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  const bankAccount = req.body.bank_account;
  if (!bankAccount) return res.status(400).json({ error: 'Bank account required' });

  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (rows.length === 0) return res.status(400).json({ error: 'No data rows found' });

    // Auto-detect columns (flexible matching)
    const cols = Object.keys(rows[0]);
    const findCol = (...names) => cols.find(c => names.some(n => c.toLowerCase().includes(n.toLowerCase())));
    const dateCol = findCol('date', 'trans date', 'posting date', 'post date', 'transaction date');
    const descCol = findCol('description', 'desc', 'memo', 'narrative', 'details', 'payee', 'name');
    const amountCol = findCol('amount', 'net', 'total');
    const debitCol = findCol('debit', 'withdrawal', 'dr');
    const creditCol = findCol('credit', 'deposit', 'cr');

    if (!dateCol) return res.status(400).json({ error: 'Could not find a date column. Found columns: ' + cols.join(', ') });

    const batchId = 'batch-' + Date.now();
    const ins = db.prepare('INSERT INTO bank_transactions (entity_id, bank_account_code, date, description, amount, batch_id) VALUES (?,?,?,?,?,?)');
    let count = 0;

    db.transaction(() => {
      for (const row of rows) {
        let dateVal = row[dateCol];
        if (dateVal instanceof Date) dateVal = dateVal.toISOString().slice(0, 10);
        else if (typeof dateVal === 'string') {
          // Try to parse common date formats
          const d = new Date(dateVal);
          if (!isNaN(d.getTime())) dateVal = d.toISOString().slice(0, 10);
          else continue; // skip unparseable dates
        } else if (typeof dateVal === 'number') {
          // Excel serial date
          const d = new Date((dateVal - 25569) * 86400000);
          dateVal = d.toISOString().slice(0, 10);
        } else continue;

        const desc = String(row[descCol] || '').trim();
        let amount = 0;
        if (amountCol && row[amountCol] !== '') {
          amount = parseFloat(String(row[amountCol]).replace(/[,$]/g, '')) || 0;
        } else if (debitCol || creditCol) {
          const dr = parseFloat(String(row[debitCol] || '0').replace(/[,$]/g, '')) || 0;
          const cr = parseFloat(String(row[creditCol] || '0').replace(/[,$]/g, '')) || 0;
          amount = cr - dr; // deposits positive, withdrawals negative
        }
        if (amount === 0) continue;

        ins.run(req.params.eid, bankAccount, dateVal, desc, amount, batchId);
        count++;
      }
    })();

    res.json({ count, batch_id: batchId, columns_detected: { date: dateCol, description: descCol, amount: amountCol, debit: debitCol, credit: creditCol } });
  } catch (e) {
    res.status(400).json({ error: 'Failed to parse file: ' + e.message });
  }
});

app.put('/api/entities/:eid/bank-transactions/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { account_code, memo } = req.body;
  db.prepare('UPDATE bank_transactions SET account_code=?, memo=?, status=? WHERE id=? AND entity_id=?')
    .run(account_code || null, memo || null, account_code ? 'coded' : 'pending', req.params.id, req.params.eid);
  res.json({ success: true });
});

app.post('/api/entities/:eid/bank-transactions/post', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { transaction_ids } = req.body;
  if (!transaction_ids || transaction_ids.length === 0) return res.status(400).json({ error: 'No transactions' });

  const txns = db.prepare(`SELECT * FROM bank_transactions WHERE entity_id=? AND id IN (${transaction_ids.map(()=>'?').join(',')}) AND status='coded'`)
    .all(req.params.eid, ...transaction_ids);
  if (txns.length === 0) return res.status(400).json({ error: 'No coded transactions to post' });

  const results = [];
  db.transaction(() => {
    for (const t of txns) {
      const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
      const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
        .run(req.params.eid, num, t.date, t.memo || t.description, req.user.name);
      const jeId = r.lastInsertRowid;

      if (t.amount > 0) { // deposit: debit bank, credit coded account
        db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(jeId, t.bank_account_code, t.amount, 0);
        db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(jeId, t.account_code, 0, t.amount);
      } else { // payment: credit bank, debit coded account
        db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(jeId, t.account_code, Math.abs(t.amount), 0);
        db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)').run(jeId, t.bank_account_code, 0, Math.abs(t.amount));
      }

      db.prepare('UPDATE bank_transactions SET status=?, je_id=? WHERE id=?').run('posted', jeId, t.id);
      results.push({ txn_id: t.id, je_id: jeId, entry_num: num });
    }
  })();
  res.json({ posted: results.length, results });
});

app.delete('/api/entities/:eid/bank-transactions/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM bank_transactions WHERE id=? AND entity_id=? AND status != ?').run(req.params.id, req.params.eid, 'posted');
  res.json({ success: true });
});

app.delete('/api/entities/:eid/bank-transactions/batch/:batchId', auth, requireRole('Admin','Accountant'), (req, res) => {
  const r = db.prepare('DELETE FROM bank_transactions WHERE entity_id=? AND batch_id=? AND status != ?').run(req.params.eid, req.params.batchId, 'posted');
  res.json({ deleted: r.changes });
});

// ═══ Balances (with soft close) ═══
app.get('/api/entities/:eid/balances', auth, (req, res) => {
  const { as_of, from, to, close_pl_before } = req.query;
  let dateFilter = ''; const params = [req.params.eid];
  if (as_of) { dateFilter = ' AND je.date <= ?'; params.push(as_of); }
  else if (from && to) { dateFilter = ' AND je.date >= ? AND je.date <= ?'; params.push(from, to); }
  else if (from) { dateFilter = ' AND je.date >= ?'; params.push(from); }
  else if (to) { dateFilter = ' AND je.date <= ?'; params.push(to); }

  if (close_pl_before && as_of) {
    const priorPL = db.prepare(`SELECT a.type, SUM(jl.debit) as td, SUM(jl.credit) as tc FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? AND je.date<? AND a.type IN ('Revenue','Expense') GROUP BY a.type`).all(req.params.eid, close_pl_before);
    let priorNI = 0; priorPL.forEach(r => { if (r.type==='Revenue') priorNI+=(r.tc-r.td); if (r.type==='Expense') priorNI-=(r.td-r.tc); });
    const bsRows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? AND je.date<=? AND (a.type NOT IN ('Revenue','Expense') OR je.date>=?) GROUP BY jl.account_code`).all(req.params.eid, as_of, close_pl_before);
    const results = bsRows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; let bal=isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit);
      if (r.account_code==='31000') bal+=priorNI;
      return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:bal, total_debit:r.total_debit, total_credit:r.total_credit }; });
    if (Math.abs(priorNI)>0.005 && !results.find(r=>r.code==='31000')) {
      const re=db.prepare('SELECT * FROM accounts WHERE entity_id=? AND code=?').get(req.params.eid,'31000');
      if (re) results.push({ code:'31000', name:re.name, type:re.type, subtype:re.subtype, bank_acct:0, balance:priorNI, total_debit:0, total_credit:0 });
    }
    return res.json(results);
  }

  const rows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=?${dateFilter} GROUP BY jl.account_code`).all(...params);
  res.json(rows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit), total_debit:r.total_debit, total_credit:r.total_credit }; }));
});

// ═══ Bank Rec ═══
app.get('/api/entities/:eid/reconciliations', auth, (req, res) => res.json(db.prepare('SELECT * FROM reconciliations WHERE entity_id=? ORDER BY completed_at DESC').all(req.params.eid)));
app.get('/api/entities/:eid/cleared/:accountCode', auth, (req, res) => {
  const m={}; db.prepare('SELECT entry_id, line_index FROM cleared_items WHERE entity_id=? AND account_code=?').all(req.params.eid, req.params.accountCode).forEach(c=>{m[c.entry_id+'-'+c.line_index]=true;}); res.json(m);
});
app.post('/api/entities/:eid/reconciliations', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { account_code, statement_date, statement_balance, book_balance, cleared_keys } = req.body;
  if (!account_code||!statement_date||statement_balance==null) return res.status(400).json({ error: 'Missing fields' });
  const result = db.transaction(() => {
    const r = db.prepare('INSERT INTO reconciliations (entity_id, account_code, statement_date, statement_balance, book_balance, cleared_count, completed_by) VALUES (?,?,?,?,?,?,?)').run(req.params.eid, account_code, statement_date, statement_balance, book_balance, cleared_keys?.length||0, req.user.name);
    if (cleared_keys) for (const k of cleared_keys) { const [eid,li]=k.split('-').map(Number); db.prepare('INSERT OR IGNORE INTO cleared_items (entity_id, account_code, entry_id, line_index, reconciliation_id) VALUES (?,?,?,?,?)').run(req.params.eid, account_code, eid, li, r.lastInsertRowid); }
    return r.lastInsertRowid;
  })(); res.json({ id: result });
});

// ═══ Summary ═══
app.get('/api/summary', auth, (req, res) => {
  res.json(db.prepare('SELECT * FROM entities ORDER BY code').all().map(e => {
    const rows = db.prepare(`SELECT a.type, SUM(jl.debit) as td, SUM(jl.credit) as tc FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? GROUP BY a.type`).all(e.id);
    const bt={}; rows.forEach(r=>{const isDr=r.type==='Asset'||r.type==='Expense'; bt[r.type]=isDr?(r.td-r.tc):(r.tc-r.td);});
    return { ...e, assets:bt.Asset||0, liabilities:bt.Liability||0, revenue:bt.Revenue||0, expenses:bt.Expense||0, net_income:(bt.Revenue||0)-(bt.Expense||0), entry_count: db.prepare('SELECT COUNT(*) as c FROM journal_entries WHERE entity_id=?').get(e.id).c };
  }));
});

if (process.env.NODE_ENV === 'production') app.get('*', (req, res) => res.sendFile(path.join(__dirname, '..', 'client', 'dist', 'index.html')));
app.listen(PORT, '0.0.0.0', () => console.log(`CloudLedger on port ${PORT}`));
