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
const pdfParse = require('pdf-parse');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-secret';
const DB_PATH = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'cloudledger.db');
const UPLOAD_DIR = path.resolve(path.dirname(DB_PATH), 'attachments');
const WORKPAPERS_DIR = path.resolve(path.dirname(DB_PATH), 'entity_files');

const dataDir = path.dirname(DB_PATH);
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(WORKPAPERS_DIR)) fs.mkdirSync(WORKPAPERS_DIR, { recursive: true });

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
    id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE,
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
    created_by TEXT NOT NULL, created_at TEXT DEFAULT (datetime('now')),
    updated_by TEXT, updated_at TEXT
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
  CREATE TABLE IF NOT EXISTS bank_transaction_splits (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    txn_id INTEGER NOT NULL REFERENCES bank_transactions(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    amount REAL NOT NULL,
    memo TEXT
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
  CREATE TABLE IF NOT EXISTS entity_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    folder_path TEXT NOT NULL DEFAULT '',
    stored_filename TEXT NOT NULL,
    original_name TEXT NOT NULL,
    size INTEGER NOT NULL,
    mime_type TEXT,
    uploaded_by TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS entity_folders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    folder_path TEXT NOT NULL,
    created_by TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(entity_id, folder_path)
  );
  CREATE INDEX IF NOT EXISTS idx_accounts_entity ON accounts(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_entity ON journal_entries(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_date ON journal_entries(entity_id, date);
  CREATE INDEX IF NOT EXISTS idx_jl_entry ON journal_lines(entry_id);
  CREATE INDEX IF NOT EXISTS idx_bt_entity ON bank_transactions(entity_id, bank_account_code);
  CREATE INDEX IF NOT EXISTS idx_bts_txn ON bank_transaction_splits(txn_id);
  CREATE INDEX IF NOT EXISTS idx_ja_entry ON journal_attachments(entry_id);
  CREATE INDEX IF NOT EXISTS idx_ef_entity ON entity_files(entity_id, folder_path);
  CREATE TABLE IF NOT EXISTS billcom_config (
    entity_id INTEGER PRIMARY KEY,
    environment TEXT NOT NULL DEFAULT 'sandbox',
    api_base_url TEXT NOT NULL,
    username TEXT NOT NULL,
    password_enc TEXT NOT NULL,
    org_id TEXT NOT NULL,
    dev_key_enc TEXT NOT NULL,
    default_ap_account TEXT,
    last_tested_at TEXT,
    last_test_status TEXT,
    last_test_message TEXT,
    updated_by TEXT,
    updated_at TEXT,
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE TABLE IF NOT EXISTS billcom_account_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    billcom_account_id TEXT NOT NULL,
    billcom_account_name TEXT,
    cl_account_code TEXT NOT NULL,
    created_at TEXT,
    UNIQUE(entity_id, billcom_account_id),
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_bam_entity ON billcom_account_map(entity_id);
  CREATE TABLE IF NOT EXISTS billcom_sync_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    sync_type TEXT NOT NULL,
    billcom_id TEXT,
    cl_entry_id INTEGER,
    status TEXT NOT NULL,
    message TEXT,
    created_at TEXT,
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_bsl_entity ON billcom_sync_log(entity_id, created_at);
  CREATE INDEX IF NOT EXISTS idx_bsl_billcom_id ON billcom_sync_log(billcom_id);
`);

const userCount = db.prepare('SELECT COUNT(*) as c FROM users').get();
if (userCount.c === 0) {
  const adminName = process.env.ADMIN_NAME || 'Admin';
  const adminEmail = (process.env.ADMIN_EMAIL || 'admin@company.com').toLowerCase();
  const adminPassword = process.env.ADMIN_PASSWORD || 'admin';
  db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?, ?, ?, ?)').run(adminName, adminEmail, bcrypt.hashSync(adminPassword, 10), 'Admin');
  console.log('Default admin created: ' + adminEmail);
}

// Schema migrations for columns added after initial release
const jeCols = db.prepare("PRAGMA table_info(journal_entries)").all().map(c => c.name);
if (!jeCols.includes('updated_by')) db.exec("ALTER TABLE journal_entries ADD COLUMN updated_by TEXT");
if (!jeCols.includes('updated_at')) db.exec("ALTER TABLE journal_entries ADD COLUMN updated_at TEXT");

// Phase 3: default_cash_account on billcom_config (for payment JEs)
const bcCfgCols = db.prepare("PRAGMA table_info(billcom_config)").all().map(c => c.name);
if (!bcCfgCols.includes('default_cash_account')) db.exec("ALTER TABLE billcom_config ADD COLUMN default_cash_account TEXT");

// === Bill.com integration helpers ===
const cryptoMod = require('crypto');
const BILLCOM_ENC_KEY = process.env.BILLCOM_ENCRYPTION_KEY || '';
function billcomEncrypt(plaintext) {
  if (!BILLCOM_ENC_KEY || BILLCOM_ENC_KEY.length !== 64) throw new Error('BILLCOM_ENCRYPTION_KEY missing or invalid');
  const key = Buffer.from(BILLCOM_ENC_KEY, 'hex');
  const iv = cryptoMod.randomBytes(12);
  const cipher = cryptoMod.createCipheriv('aes-256-gcm', key, iv);
  const ct = Buffer.concat([cipher.update(String(plaintext), 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return iv.toString('hex') + ':' + tag.toString('hex') + ':' + ct.toString('hex');
}
function billcomDecrypt(enc) {
  if (!BILLCOM_ENC_KEY || BILLCOM_ENC_KEY.length !== 64) throw new Error('BILLCOM_ENCRYPTION_KEY missing or invalid');
  if (!enc) return '';
  const parts = enc.split(':');
  if (parts.length !== 3) throw new Error('Malformed encrypted blob');
  const key = Buffer.from(BILLCOM_ENC_KEY, 'hex');
  const iv = Buffer.from(parts[0], 'hex');
  const tag = Buffer.from(parts[1], 'hex');
  const ct = Buffer.from(parts[2], 'hex');
  const decipher = cryptoMod.createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString('utf8');
}
function maskSecret(s) {
  if (!s) return '';
  return s.length <= 4 ? '****' : '****' + s.slice(-4);
}
const BILLCOM_BASE_URLS = {
  sandbox: 'https://gateway.stage.bill.com/connect/v3',
  production: 'https://gateway.prod.bill.com/connect/v3'
};
async function billcomLogin({ username, password, orgId, devKey, baseUrl }) {
  const url = (baseUrl || BILLCOM_BASE_URLS.sandbox) + '/login';
  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password, organizationId: orgId, devKey })
  });
  const text = await resp.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw: text }; }
  if (!resp.ok) {
    const errMsg = Array.isArray(data) ? data.map(e => e.message).join('; ') : (data.message || text);
    throw new Error('HTTP ' + resp.status + ': ' + errMsg);
  }
  return data;
}

async function billcomListAccounts({ sessionId, devKey, baseUrl }) {
  // Bill.com v3 API: GET /v3/classifications/chart-of-accounts
  // base already includes /connect/v3, so we append the resource path
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const out = [];
  let nextPage = null;
  const max = 100; // v3 default cap
  // Paginate via nextPage token if present in the response
  while (true) {
    const params = new URLSearchParams({ max: String(max) });
    if (nextPage) params.set('nextPage', nextPage);
    const url = base + '/classifications/chart-of-accounts?' + params.toString();
    const resp = await fetch(url, {
      method: 'GET',
      headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' }
    });
    const text = await resp.text();
    console.log('[billcom COA] HTTP ' + resp.status + ' from ' + url + ' :: ' + text.slice(0, 500));
    let json; try { json = JSON.parse(text); } catch { throw new Error('Non-JSON response (HTTP ' + resp.status + '): ' + text.slice(0, 200)); }
    if (!resp.ok) {
      const msg = (Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join('; ') : (json.message || json.error_message || ('HTTP ' + resp.status + ' body=' + text.slice(0, 200))));
      throw new Error('Bill.com error: ' + msg);
    }
    // v3 typically returns { results: [...], nextPage: "..." }
    const items = Array.isArray(json.results) ? json.results : (Array.isArray(json) ? json : []);
    out.push(...items);
    nextPage = json.nextPage || null;
    if (!nextPage) break;
    if (out.length > 10000) break; // safety
  }
  return out;
}

// Generic paginated GET for v3 list endpoints. Used for bills + payments.
async function billcomListV3({ sessionId, devKey, baseUrl, resourcePath, extraParams }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const out = [];
  let nextPage = null;
  const max = 100;
  let pageCount = 0;
  while (true) {
    const params = new URLSearchParams({ max: String(max), ...(extraParams || {}) });
    if (nextPage) params.set('nextPage', nextPage);
    const url = base + resourcePath + '?' + params.toString();
    const resp = await fetch(url, {
      method: 'GET',
      headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' }
    });
    const text = await resp.text();
    let json; try { json = JSON.parse(text); } catch { throw new Error('Non-JSON response (HTTP ' + resp.status + '): ' + text.slice(0, 200)); }
    if (!resp.ok) {
      const msg = (Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join('; ') : (json.message || ('HTTP ' + resp.status + ' body=' + text.slice(0, 200))));
      throw new Error('Bill.com error: ' + msg);
    }
    const items = Array.isArray(json.results) ? json.results : (Array.isArray(json) ? json : []);
    out.push(...items);
    nextPage = json.nextPage || null;
    pageCount++;
    if (!nextPage) break;
    if (out.length > 10000) break;
    if (pageCount > 100) break;
  }
  return out;
}

async function billcomListBills(args) {
  return billcomListV3({ ...args, resourcePath: '/bills' });
}

async function billcomListPayments(args) {
  return billcomListV3({ ...args, resourcePath: '/payments' });
}

async function billcomGetById({ sessionId, devKey, baseUrl, resourcePath, id }) {
  const url = (baseUrl || BILLCOM_BASE_URLS.sandbox) + resourcePath + '/' + encodeURIComponent(id);
  const resp = await fetch(url, {
    method: 'GET',
    headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' }
  });
  const text = await resp.text();
  let json; try { json = JSON.parse(text); } catch { throw new Error('Non-JSON detail (HTTP ' + resp.status + '): ' + text.slice(0, 200)); }
  if (!resp.ok) {
    const msg = (Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join('; ') : (json.message || ('HTTP ' + resp.status)));
    throw new Error('detail HTTP ' + resp.status + ': ' + msg);
  }
  return json;
}


// One-time recovery: earlier versions of the replace endpoint accidentally saved files
// to WORKPAPERS_DIR root (or to an "undefined" subdir) instead of WORKPAPERS_DIR/<entity_id>/.
// Walk those locations, match files to DB rows by stored_filename, and move them home.
try {
  const recoverFromDir = dir => {
    if (!fs.existsSync(dir)) return 0;
    let moved = 0;
    for (const fname of fs.readdirSync(dir)) {
      const src = path.join(dir, fname);
      try { if (!fs.statSync(src).isFile()) continue; } catch { continue; }
      const row = db.prepare('SELECT entity_id FROM entity_files WHERE stored_filename = ?').get(fname);
      if (!row) continue;
      const targetDir = path.join(WORKPAPERS_DIR, String(row.entity_id));
      try { fs.mkdirSync(targetDir, { recursive: true }); } catch {}
      const dst = path.join(targetDir, fname);
      if (fs.existsSync(dst)) continue;
      try { fs.renameSync(src, dst); moved++; }
      catch (e) { console.error('Recovery move failed for', fname, e.message); }
    }
    return moved;
  };
  const a = recoverFromDir(WORKPAPERS_DIR);
  const b = recoverFromDir(path.join(WORKPAPERS_DIR, 'undefined'));
  if (a + b > 0) console.log('[workpapers recovery] Moved ' + (a + b) + ' orphaned file(s) to correct entity directories');
  // Best-effort: remove the empty "undefined" dir if it's now empty
  try {
    const undefDir = path.join(WORKPAPERS_DIR, 'undefined');
    if (fs.existsSync(undefDir) && fs.readdirSync(undefDir).length === 0) fs.rmdirSync(undefDir);
  } catch {}
} catch (e) { console.error('Workpapers recovery routine failed:', e); }

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
if (process.env.NODE_ENV === 'production') app.use(express.static(path.join(__dirname, '..', 'client', 'dist'), {
  setHeaders: (res, filePath) => {
    // Never cache index.html — forces browsers to always re-fetch it (and pick up new asset URLs)
    if (filePath.endsWith('index.html')) {
      res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    }
  }
}));

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
app.get('/api/auth/me', auth, (req, res) => {
  const user = db.prepare('SELECT id, name, email, role FROM users WHERE id = ?').get(req.user.id);
  if (!user) return res.status(404).json({ error: 'User not found' });
  res.json(user);
});
app.put('/api/auth/profile', auth, (req, res) => {
  const { name, email } = req.body;
  if (!name || !email) return res.status(400).json({ error: 'Name and email required' });
  const existing = db.prepare('SELECT id FROM users WHERE email = ? AND id != ?').get(email.toLowerCase(), req.user.id);
  if (existing) return res.status(400).json({ error: 'Email already in use by another account' });
  db.prepare('UPDATE users SET name = ?, email = ? WHERE id = ?').run(name, email.toLowerCase(), req.user.id);
  const updated = db.prepare('SELECT id, name, email, role FROM users WHERE id = ?').get(req.user.id);
  res.json(updated);
});
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
  const { name } = req.body; if (!name) return res.status(400).json({ error: 'Name required' });
  // Auto-generate a code from the name (used internally for sorting/uniqueness)
  const baseCode = name.replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 8) || 'ENT';
  let code = baseCode; let n = 1;
  while (db.prepare('SELECT id FROM entities WHERE code = ?').get(code)) { code = baseCode + n; n++; }
  try { const r = db.prepare('INSERT INTO entities (code, name) VALUES (?, ?)').run(code, name); const eid = r.lastInsertRowid;
    const ins = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
    db.transaction(() => { for (const a of DEFAULT_COA) ins.run(eid, a.code, a.name, a.type, a.subtype, a.bank); })();
    res.json({ id: eid, code, name }); } catch(e) { throw e; }
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

// Import trial balance: replaces COA and posts a beginning-balance JE
// Account type derived from code: <=19999 Asset, <=29999 Liability, <=39999 Equity, <=49999 Revenue, 50000-69999 Expense, >=70000 Revenue
app.post('/api/entities/:eid/import-tb', auth, requireRole('Admin','Accountant'), memUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const eid = +req.params.eid;
  const asOfDate = req.body.as_of_date || '2024-12-31';
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (rows.length === 0) return res.status(400).json({ error: 'No data rows found in file' });

    const cols = Object.keys(rows[0]);
    const norm = c => String(c).toLowerCase().trim();
    const findCol = (patterns, exclude = []) => {
      const pool = cols.filter(c => !exclude.includes(c));
      for (const pat of patterns) { const hit = pool.find(c => norm(c) === pat); if (hit) return hit; }
      for (const pat of patterns) { const hit = pool.find(c => norm(c).includes(pat)); if (hit) return hit; }
      return null;
    };
    const codeCol = findCol(['account number','account #','account code','acct number','acct code','acct','code','number']);
    const nameCol = findCol(['account name','account description','acct name','description','name'], [codeCol]);
    const amtCol  = findCol(['balance','ending balance','amount','total'], [codeCol, nameCol].filter(Boolean));
    const drCol   = findCol(['debit'], [codeCol, nameCol, amtCol].filter(Boolean));
    const crCol   = findCol(['credit'], [codeCol, nameCol, amtCol].filter(Boolean));

    if (!codeCol) return res.status(400).json({ error: 'Could not find account number/code column. Found: ' + cols.join(', ') });
    if (!nameCol) return res.status(400).json({ error: 'Could not find account name column. Found: ' + cols.join(', ') });
    if (!amtCol && !drCol && !crCol) return res.status(400).json({ error: 'Could not find amount or debit/credit columns. Found: ' + cols.join(', ') });

    // Parse rows into accounts
    const typeFromCode = (codeStr) => {
      const n = parseInt(String(codeStr).replace(/[^0-9]/g, ''), 10);
      if (isNaN(n)) return null;
      if (n <= 19999) return 'Asset';
      if (n <= 29999) return 'Liability';
      if (n <= 39999) return 'Equity';
      if (n <= 49999) return 'Revenue';
      if (n <= 69999) return 'Expense';
      return 'Revenue';
    };

    const parsed = [];
    for (const row of rows) {
      const code = String(row[codeCol] || '').trim();
      const name = String(row[nameCol] || '').trim();
      if (!code || !name) continue;
      const type = typeFromCode(code);
      if (!type) continue;
      let dr = 0, cr = 0, amt = null;
      if (drCol || crCol) {
        // Separate debit and credit columns: take them at face value, do NOT flip by account type
        dr = parseFloat(String(row[drCol] || '0').replace(/[,$()]/g, '')) || 0;
        cr = parseFloat(String(row[crCol] || '0').replace(/[,$()]/g, '')) || 0;
      } else if (amtCol) {
        const raw = String(row[amtCol] || '').trim();
        const isParen = /^\(.*\)$/.test(raw);
        let v = parseFloat(raw.replace(/[,$()]/g, '')) || 0;
        if (isParen) v = -v;
        amt = v;
      }
      parsed.push({ code, name, type, dr, cr, amt });
    }

    if (parsed.length === 0) return res.status(400).json({ error: 'No valid rows found. Check that account codes are numeric.' });

    // If using a single signed amount column, detect the sign convention:
    //   "debit-positive" (a.k.a. signed TB): debits +, credits -. Sum of all amounts ≈ 0 when balanced.
    //   "natural-side": positive means the account's normal side (Asset/Expense+ = debit, L/E/Rev+ = credit).
    let signMode = 'debit-positive';
    if (parsed.some(p => p.amt !== null)) {
      const sumSigned = parsed.reduce((s, p) => s + (p.amt || 0), 0);
      signMode = Math.abs(sumSigned) < 0.01 ? 'debit-positive' : 'natural';
    }

    // Build journal lines for opening balance JE
    const lines = [];
    let totalDr = 0, totalCr = 0;
    for (const p of parsed) {
      let dr = 0, cr = 0;
      if (p.amt === null) {
        // Came from separate debit/credit columns
        dr = p.dr; cr = p.cr;
      } else if (signMode === 'debit-positive') {
        if (p.amt >= 0) dr = p.amt; else cr = -p.amt;
      } else {
        // natural-side: positive = natural balance side
        const isDebitNatural = p.type === 'Asset' || p.type === 'Expense';
        if (isDebitNatural) { if (p.amt >= 0) dr = p.amt; else cr = -p.amt; }
        else { if (p.amt >= 0) cr = p.amt; else dr = -p.amt; }
      }
      if (Math.abs(dr) < 0.005 && Math.abs(cr) < 0.005) continue;
      lines.push({ account_code: p.code, debit: dr, credit: cr });
      totalDr += dr; totalCr += cr;
    }

    // Check if balanced; if not, plug to retained earnings (by code OR name containing "retained earnings")
    const diff = +(totalDr - totalCr).toFixed(2);
    let plugAdded = false;
    if (Math.abs(diff) > 0.005) {
      const reAcct = parsed.find(p => p.code === '31000')
        || parsed.find(p => p.type === 'Equity' && /retain(ed)?\s*earning/i.test(p.name));
      if (reAcct) {
        const retainedCode = reAcct.code;
        const existing = lines.find(l => l.account_code === retainedCode);
        if (existing) {
          if (diff > 0) existing.credit += diff; else existing.debit += -diff;
        } else {
          if (diff > 0) lines.push({ account_code: retainedCode, debit: 0, credit: diff });
          else lines.push({ account_code: retainedCode, debit: -diff, credit: 0 });
        }
        plugAdded = true;
      } else {
        return res.status(400).json({ error: 'Trial balance does not balance (off by ' + diff.toFixed(2) + ') and no Retained Earnings account was found to plug the difference. Add an equity account with "Retained Earnings" in the name.' });
      }
    }

    db.transaction(() => {
      // Remove any prior opening-balance imports so the new import replaces them cleanly
      // (otherwise re-importing would stack on top of the previous opening balances)
      db.prepare("DELETE FROM journal_entries WHERE entity_id = ? AND memo = 'Opening balance from imported trial balance'").run(eid);

      // Replace chart of accounts
      db.prepare('DELETE FROM accounts WHERE entity_id = ?').run(eid);
      const insAcct = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
      // Build set of unique accounts (parsed + retained earnings if needed)
      const allCodes = new Set(parsed.map(p => p.code));
      for (const p of parsed) {
        const isBank = p.type === 'Asset' && /cash|bank|checking|savings/i.test(p.name);
        insAcct.run(eid, p.code, p.name, p.type, '', isBank ? 1 : 0);
      }
      // Make sure retained earnings exists if we plugged
      if (plugAdded && !allCodes.has('31000')) {
        insAcct.run(eid, '31000', 'Retained Earnings', 'Equity', '', 0);
      }

      // Create the opening balance JE
      if (lines.length > 0) {
        const lastNum = db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(eid);
        const entryNum = (lastNum.m || 0) + 1;
        const jeRes = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?, ?, ?, ?, ?)').run(eid, entryNum, asOfDate, 'Opening balance from imported trial balance', req.user.name || req.user.email);
        const jeId = jeRes.lastInsertRowid;
        const insLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?, ?, ?, ?)');
        for (const l of lines) insLine.run(jeId, l.account_code, l.debit, l.credit);
      }
    })();

    res.json({ success: true, accounts_imported: parsed.length, lines: lines.length, total_debit: totalDr, total_credit: totalCr, plug_added: plugAdded });
  } catch (e) {
    res.status(400).json({ error: 'Failed to import trial balance: ' + e.message });
  }
});

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

app.put('/api/entities/:eid/accounts/:code', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { new_code, name, type, subtype, bank_acct } = req.body;
  const oldCode = req.params.code;
  const eid = req.params.eid;
  const acct = db.prepare('SELECT * FROM accounts WHERE entity_id=? AND code=?').get(eid, oldCode);
  if (!acct) return res.status(404).json({ error: 'Account not found' });
  const updatedCode = new_code || oldCode;
  const updatedName = name !== undefined ? name : acct.name;
  const updatedType = type || acct.type;
  const updatedSubtype = subtype !== undefined ? subtype : acct.subtype;
  const updatedBank = bank_acct !== undefined ? (bank_acct ? 1 : 0) : acct.bank_acct;
  if (updatedCode !== oldCode) {
    const existing = db.prepare('SELECT id FROM accounts WHERE entity_id=? AND code=?').get(eid, updatedCode);
    if (existing) return res.status(400).json({ error: 'Account code ' + updatedCode + ' already exists' });
  }
  db.transaction(() => {
    if (updatedCode !== oldCode) {
      // Update code in all related tables
      db.prepare('UPDATE journal_lines SET account_code=? WHERE account_code=? AND entry_id IN (SELECT id FROM journal_entries WHERE entity_id=?)').run(updatedCode, oldCode, eid);
      db.prepare('UPDATE bank_transactions SET bank_account_code=? WHERE bank_account_code=? AND entity_id=?').run(updatedCode, oldCode, eid);
      db.prepare('UPDATE bank_transactions SET account_code=? WHERE account_code=? AND entity_id=?').run(updatedCode, oldCode, eid);
      db.prepare('UPDATE cleared_items SET account_code=? WHERE account_code=? AND entity_id=?').run(updatedCode, oldCode, eid);
    }
    db.prepare('UPDATE accounts SET code=?, name=?, type=?, subtype=?, bank_acct=? WHERE entity_id=? AND code=?')
      .run(updatedCode, updatedName, updatedType, updatedSubtype, updatedBank, eid, oldCode);
  })();
  res.json({ success: true, code: updatedCode });
});

// ═══ Journal Entries ═══
app.get('/api/entities/:eid/entries', auth, (req, res) => {
  const { from, to } = req.query; let sql = 'SELECT * FROM journal_entries WHERE entity_id = ?'; const params = [req.params.eid];
  if (from) { sql += ' AND date >= ?'; params.push(from); } if (to) { sql += ' AND date <= ?'; params.push(to); }
  sql += ' ORDER BY entry_num ASC';
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
    db.prepare("UPDATE journal_entries SET date=?, memo=?, updated_by=?, updated_at=datetime('now') WHERE id=?")
      .run(date, memo, req.user.name || req.user.email, req.params.id);
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
  const txns = db.prepare(sql).all(...params);
  if (txns.length === 0) return res.json([]);
  const splits = db.prepare(`SELECT * FROM bank_transaction_splits WHERE txn_id IN (${txns.map(()=>'?').join(',')}) ORDER BY id`).all(...txns.map(t=>t.id));
  const splitMap = {}; for (const s of splits) { (splitMap[s.txn_id] = splitMap[s.txn_id] || []).push(s); }
  res.json(txns.map(t => ({ ...t, splits: splitMap[t.id] || [] })));
});

app.post('/api/entities/:eid/bank-transactions/upload', auth, requireRole('Admin','Accountant'), memUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  const bankAccount = req.body.bank_account;
  if (!bankAccount) return res.status(400).json({ error: 'Bank account required' });

  const isPdf = req.file.mimetype === 'application/pdf' || req.file.originalname.toLowerCase().endsWith('.pdf');

  try {
    let rows = []; // each: { date: 'YYYY-MM-DD', description: string, amount: number }

    if (isPdf) {
      // ── PDF bank statement parsing ──
      const data = await pdfParse(req.file.buffer);
      const text = data.text || '';
      const lines = text.split(/\n/).map(l => l.trim()).filter(Boolean);

      // Common bank statement date patterns
      const dateRx = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/;
      const altDateRx = /^(\w{3,9})\s+(\d{1,2}),?\s+(\d{4})/; // "January 15, 2024"
      // Match amounts: optional sign, optional $, digits with commas, decimal, OR amounts in parens like (1,234.56)
      const numRx = /\(?\$?[\d,]+\.\d{2}\)?|[\-\+]\$?[\d,]+\.\d{2}/g;

      // Keywords that indicate money going OUT (withdrawal/debit/payment)
      const withdrawalRx = /withdraw|payment|debit|paid out|ach debit|wire out|check\b|chk\b|pmt\b/i;
      // Keywords that indicate money coming IN (deposit/credit)
      const depositRx = /deposit|credit|interest\s*(paid|capitali|earned|income)|ach credit|wire in|xfer in/i;

      const parseDate = s => {
        let m = s.match(dateRx);
        if (m) {
          let [, mm, dd, yy] = m;
          if (yy.length === 2) yy = (+yy > 50 ? '19' : '20') + yy;
          const d = new Date(+yy, +mm - 1, +dd);
          if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
        }
        m = s.match(altDateRx);
        if (m) {
          const d = new Date(m[0]);
          if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
        }
        return null;
      };

      const parseAmt = s => {
        const trimmed = s.trim();
        const isParen = /^\(.*\)$/.test(trimmed);
        const clean = trimmed.replace(/[$,\s()]/g, '');
        let v = parseFloat(clean) || 0;
        if (isParen) v = -v;
        if (trimmed.startsWith('-') && v > 0) v = -v;
        return v;
      };

      let lastDate = null;
      for (const line of lines) {
        const dt = parseDate(line);
        if (dt) lastDate = dt;
        else if (!lastDate) continue;

        const amounts = line.match(numRx);
        if (!amounts || amounts.length === 0) continue;

        const lineDate = parseDate(line);
        if (!lineDate && !dt) continue;
        const usedDate = lineDate || lastDate;

        // Extract description: everything between the date and the first amount
        const firstAmtIdx = line.indexOf(amounts[0]);
        let desc = '';
        if (lineDate) {
          const dateMatch = line.match(dateRx) || line.match(altDateRx);
          if (dateMatch) desc = line.substring(dateMatch.index + dateMatch[0].length, firstAmtIdx).trim();
        }
        if (!desc) desc = line.substring(0, firstAmtIdx).replace(dateRx, '').replace(altDateRx, '').trim();

        // Clean up description: remove trailing/leading parens, extra whitespace
        desc = desc.replace(/[\(\)]+$/g, '').replace(/^[\(\)]+/g, '').trim();

        // Determine the transaction amount
        let amount = 0;
        if (amounts.length === 1) {
          amount = parseAmt(amounts[0]);
        } else if (amounts.length === 2) {
          amount = parseAmt(amounts[0]);
        } else if (amounts.length >= 3) {
          const a1 = parseAmt(amounts[0]);
          const a2 = parseAmt(amounts[1]);
          amount = a1 !== 0 ? a1 : a2;
        }

        // If amount is positive but description indicates a withdrawal, flip it
        // If amount is negative but description indicates a deposit, flip it
        if (amount > 0 && withdrawalRx.test(desc) && !depositRx.test(desc)) {
          amount = -amount;
        } else if (amount < 0 && depositRx.test(desc) && !withdrawalRx.test(desc)) {
          amount = -amount;
        }

        if (amount === 0 || !desc) continue;

        rows.push({ date: usedDate, description: desc.substring(0, 500), amount });
      }

      if (rows.length === 0) {
        return res.status(400).json({
          error: 'Could not extract transactions from this PDF. The parser found ' + lines.length + ' text lines but no recognizable transaction rows. Try exporting as CSV or Excel from your bank instead.',
          pdf_lines_preview: lines.slice(0, 30)
        });
      }
    } else {
      // ── CSV / Excel parsing (existing logic) ──
      const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const xlRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (xlRows.length === 0) return res.status(400).json({ error: 'No data rows found' });

      const cols = Object.keys(xlRows[0]);
      const norm = s => String(s).toLowerCase().trim();
      const findCol = (...names) => cols.find(c => { const cn = norm(c); return names.some(n => cn.includes(norm(n))); });
      const findColExact = (...names) => cols.find(c => { const cn = norm(c); return names.some(n => cn === norm(n)); });
      const dateCol = findCol('date', 'trans date', 'posting date', 'post date', 'transaction date');
      const descCol = findCol('description', 'desc', 'memo', 'narrative', 'details', 'payee', 'name');
      const debitCol = findColExact('debit','debits','dr','withdrawal','withdrawals','withdrawn','paid out','money out','out','outflow','outflows','payment','payments','charge','charges')
                    || findCol('debit','withdrawal','withdrawn','paid out','money out','outflow','dr ','dr.');
      const creditCol = findColExact('credit','credits','cr','deposit','deposits','paid in','money in','in','inflow','inflows','receipt','receipts')
                    || findCol('credit','deposit','paid in','money in','inflow','receipt','cr ','cr.');
      const amountCol = (debitCol && creditCol) ? null : findCol('amount', 'net', 'total');

      if (!dateCol) return res.status(400).json({ error: 'Could not find a date column. Found columns: ' + cols.join(', ') });

      const parseNum = v => { if (v === '' || v == null) return 0; const s = String(v).trim(); if (!s) return 0; const isParen = /^\(.*\)$/.test(s); const n = parseFloat(s.replace(/[,$()\s]/g, '')) || 0; return isParen ? -Math.abs(n) : n; };

      for (const row of xlRows) {
        let dateVal = row[dateCol];
        if (dateVal instanceof Date) dateVal = dateVal.toISOString().slice(0, 10);
        else if (typeof dateVal === 'string') {
          const d = new Date(dateVal);
          if (!isNaN(d.getTime())) dateVal = d.toISOString().slice(0, 10);
          else continue;
        } else if (typeof dateVal === 'number') {
          const d = new Date((dateVal - 25569) * 86400000);
          dateVal = d.toISOString().slice(0, 10);
        } else continue;

        const desc = String(row[descCol] || '').trim();
        let amount = 0;
        if (debitCol && creditCol) {
          const dr = Math.abs(parseNum(row[debitCol]));
          const cr = Math.abs(parseNum(row[creditCol]));
          amount = cr - dr;
        } else if (amountCol && row[amountCol] !== '' && row[amountCol] != null) {
          amount = parseNum(row[amountCol]);
        } else if (debitCol) {
          amount = -Math.abs(parseNum(row[debitCol]));
        } else if (creditCol) {
          amount = Math.abs(parseNum(row[creditCol]));
        }
        if (amount === 0) continue;

        rows.push({ date: dateVal, description: desc, amount });
      }

      if (rows.length === 0) return res.status(400).json({ error: 'No valid transaction rows found' });
    }

    // ── Insert into database (shared path for both PDF and CSV/XLSX) ──
    const batchId = 'batch-' + Date.now();
    const ins = db.prepare('INSERT INTO bank_transactions (entity_id, bank_account_code, date, description, amount, batch_id) VALUES (?,?,?,?,?,?)');
    let count = 0;
    db.transaction(() => {
      for (const r of rows) {
        ins.run(req.params.eid, bankAccount, r.date, r.description, r.amount, batchId);
        count++;
      }
    })();

    res.json({ count, batch_id: batchId, format: isPdf ? 'pdf' : 'csv/xlsx' });
  } catch (e) {
    res.status(400).json({ error: 'Failed to parse file: ' + e.message });
  }
});

app.put('/api/entities/:eid/bank-transactions/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { account_code, memo } = req.body;
  // Setting a single account_code clears any existing splits
  db.transaction(() => {
    db.prepare('DELETE FROM bank_transaction_splits WHERE txn_id=?').run(req.params.id);
    db.prepare('UPDATE bank_transactions SET account_code=?, memo=?, status=? WHERE id=? AND entity_id=?')
      .run(account_code || null, memo || null, account_code ? 'coded' : 'pending', req.params.id, req.params.eid);
  })();
  res.json({ success: true });
});

// Set multiple splits for a single bank transaction.
// Splits: [{ account_code, amount, memo }]. Amounts are positive numbers, must sum to abs(txn.amount).
app.put('/api/entities/:eid/bank-transactions/:id/splits', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { splits } = req.body;
  const txn = db.prepare('SELECT * FROM bank_transactions WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status === 'posted') return res.status(400).json({ error: 'Cannot edit a posted transaction' });
  if (!Array.isArray(splits) || splits.length === 0) return res.status(400).json({ error: 'At least one split required' });
  for (const s of splits) {
    if (!s.account_code) return res.status(400).json({ error: 'Each split needs an account' });
    if (!(Number(s.amount) > 0)) return res.status(400).json({ error: 'Each split amount must be greater than zero' });
  }
  const total = splits.reduce((sum, s) => sum + Number(s.amount), 0);
  const target = Math.abs(txn.amount);
  if (Math.abs(total - target) > 0.005) return res.status(400).json({ error: 'Splits total ' + total.toFixed(2) + ' does not match transaction amount ' + target.toFixed(2) });

  db.transaction(() => {
    db.prepare('DELETE FROM bank_transaction_splits WHERE txn_id=?').run(txn.id);
    const ins = db.prepare('INSERT INTO bank_transaction_splits (txn_id, account_code, amount, memo) VALUES (?,?,?,?)');
    for (const s of splits) ins.run(txn.id, s.account_code, Number(s.amount), s.memo || null);
    // Clear the single-code field and mark coded
    db.prepare('UPDATE bank_transactions SET account_code=NULL, status=? WHERE id=?').run('coded', txn.id);
  })();
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
      const splits = db.prepare('SELECT * FROM bank_transaction_splits WHERE txn_id=? ORDER BY id').all(t.id);
      const hasSplits = splits.length > 0;
      // If no splits and no single account code, skip (shouldn't happen for 'coded' status, but defend anyway)
      if (!hasSplits && !t.account_code) continue;

      const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
      const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
        .run(req.params.eid, num, t.date, t.memo || t.description, req.user.name);
      const jeId = r.lastInsertRowid;
      const insLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)');
      const abs = Math.abs(t.amount);

      if (t.amount > 0) {
        // Deposit: debit bank once, credit each coded account
        insLine.run(jeId, t.bank_account_code, abs, 0);
        if (hasSplits) { for (const s of splits) insLine.run(jeId, s.account_code, 0, s.amount); }
        else { insLine.run(jeId, t.account_code, 0, abs); }
      } else {
        // Payment: debit each coded account, credit bank once
        if (hasSplits) { for (const s of splits) insLine.run(jeId, s.account_code, s.amount, 0); }
        else { insLine.run(jeId, t.account_code, abs, 0); }
        insLine.run(jeId, t.bank_account_code, 0, abs);
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
    // Find the retained earnings account for this entity: code 31000, or any Equity account named "Retained Earnings"
    let reAcct = db.prepare("SELECT * FROM accounts WHERE entity_id=? AND code='31000'").get(req.params.eid)
      || db.prepare("SELECT * FROM accounts WHERE entity_id=? AND type='Equity' AND LOWER(name) LIKE '%retained earning%' ORDER BY code LIMIT 1").get(req.params.eid);
    // Auto-create 39000 Retained Earnings if no RE account exists and there is prior-period P&L to close
    if (!reAcct && Math.abs(priorNI) > 0.005) {
      try {
        db.prepare("INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, '39000', 'Retained Earnings', 'Equity', '', 0)").run(req.params.eid);
        reAcct = db.prepare("SELECT * FROM accounts WHERE entity_id=? AND code='39000'").get(req.params.eid);
      } catch (e) { console.error('Auto-create RE 39000 failed:', e.message); }
    }
    const reCode = reAcct ? reAcct.code : null;
    const bsRows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? AND je.date<=? AND (a.type NOT IN ('Revenue','Expense') OR je.date>=?) GROUP BY jl.account_code`).all(req.params.eid, as_of, close_pl_before);
    const results = bsRows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; let bal=isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit);
      if (reCode && r.account_code===reCode) bal+=priorNI;
      return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:bal, total_debit:r.total_debit, total_credit:r.total_credit }; });
    if (Math.abs(priorNI)>0.005 && reCode && !results.find(r=>r.code===reCode)) {
      results.push({ code:reCode, name:reAcct.name, type:reAcct.type, subtype:reAcct.subtype, bank_acct:0, balance:priorNI, total_debit:0, total_credit:0 });
    }
    return res.json(results);
  }

  const rows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=?${dateFilter} GROUP BY jl.account_code`).all(...params);
  res.json(rows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit), total_debit:r.total_debit, total_credit:r.total_credit }; }));
});

// ═══ Entity Workpapers (files + folders) ═══
// Disk-based uploader that routes files into per-entity directories
const workpaperStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    // Support both /entities/:eid/files (upload) and /entity-files/:id (replace, where req.entityId is stashed)
    const rawEid = req.params.eid != null ? req.params.eid : req.entityId;
    const eid = String(rawEid != null ? rawEid : '').replace(/[^0-9]/g, '');
    if (!eid) { cb(new Error('Missing entity_id for upload destination')); return; }
    const entityDir = path.join(WORKPAPERS_DIR, eid);
    try { fs.mkdirSync(entityDir, { recursive: true }); cb(null, entityDir); }
    catch (e) { console.error('workpapers mkdir failed:', entityDir, e); cb(e); }
  },
  filename: (req, file, cb) => {
    const safe = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
    cb(null, Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + safe);
  }
});
const workpaperUpload = multer({ storage: workpaperStorage, limits: { fileSize: 100 * 1024 * 1024 } }); // 100MB

// Normalize a folder path: trim, collapse slashes, no leading/trailing slash, no .., no empty segments
const normFolderPath = p => {
  if (!p) return '';
  const parts = String(p).split('/').map(s => s.trim()).filter(s => s && s !== '.' && s !== '..');
  return parts.join('/');
};

app.get('/api/entities/:eid/files', auth, (req, res) => {
  const files = db.prepare('SELECT id, folder_path, original_name, size, mime_type, uploaded_by, created_at FROM entity_files WHERE entity_id=? ORDER BY folder_path, original_name').all(req.params.eid);
  const folders = db.prepare('SELECT folder_path, created_by, created_at FROM entity_folders WHERE entity_id=? ORDER BY folder_path').all(req.params.eid);
  // Collect every distinct folder path from both tables plus all ancestor paths
  const folderSet = new Set();
  const addAncestors = p => { if (!p) return; const parts = p.split('/'); for (let i = 1; i <= parts.length; i++) folderSet.add(parts.slice(0, i).join('/')); };
  files.forEach(f => addAncestors(f.folder_path));
  folders.forEach(f => addAncestors(f.folder_path));
  res.json({ files, folders: Array.from(folderSet).sort() });
});

// Multer middleware that reports its own errors instead of bubbling to the default handler
const workpaperUploadMw = (req, res, next) => {
  workpaperUpload.array('files', 20)(req, res, err => {
    if (err) {
      console.error('workpaper upload error:', err);
      return res.status(400).json({ error: 'Upload failed: ' + (err.message || err.code || 'unknown error') });
    }
    next();
  });
};

app.post('/api/entities/:eid/files', auth, requireRole('Admin','Accountant'), workpaperUploadMw, (req, res) => {
  if (!req.files || req.files.length === 0) return res.status(400).json({ error: 'No files received by server. Check that the browser attached files to the "files" field.' });
  const folder = normFolderPath(req.body.folder_path);
  const ins = db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)');
  const inserted = [];
  db.transaction(() => {
    for (const f of req.files) {
      const r = ins.run(req.params.eid, folder, f.filename, f.originalname, f.size, f.mimetype || null, req.user.name || req.user.email);
      inserted.push({ id: r.lastInsertRowid, original_name: f.originalname, size: f.size });
    }
  })();
  res.json({ uploaded: inserted.length, files: inserted });
});

// Download — uses token query param like journal attachments
app.get('/api/entity-files/:id/download', (req, res) => {
  try {
    const token = req.query.token;
    if (!token) return res.status(401).json({ error: 'Token required' });
    jwt.verify(token, JWT_SECRET);
  } catch { return res.status(401).json({ error: 'Invalid token' }); }
  const f = db.prepare('SELECT * FROM entity_files WHERE id=?').get(req.params.id);
  if (!f) return res.status(404).json({ error: 'Not found' });
  const filepath = path.resolve(WORKPAPERS_DIR, String(f.entity_id), f.stored_filename);
  if (!fs.existsSync(filepath)) return res.status(404).json({ error: 'File missing on disk' });
  const inlineTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'];
  const disposition = inlineTypes.includes(f.mime_type) ? 'inline' : 'attachment';
  res.setHeader('Content-Disposition', disposition + '; filename="' + f.original_name.replace(/"/g, '') + '"');
  res.setHeader('Content-Type', f.mime_type || 'application/octet-stream');
  res.sendFile(filepath, err => { if (err && !res.headersSent) res.status(500).json({ error: 'Send failed' }); });
});

app.delete('/api/entity-files/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const f = db.prepare('SELECT * FROM entity_files WHERE id=?').get(req.params.id);
  if (!f) return res.status(404).json({ error: 'Not found' });
  try { fs.unlinkSync(path.join(WORKPAPERS_DIR, String(f.entity_id), f.stored_filename)); } catch {}
  db.prepare('DELETE FROM entity_files WHERE id=?').run(req.params.id);
  res.json({ success: true });
});

// Replace (version) a workpaper file — swaps file on disk, keeps same DB row id and folder location
app.put('/api/entity-files/:id', auth, requireRole('Admin','Accountant'), (req, res, next) => {
  const f = db.prepare('SELECT * FROM entity_files WHERE id=?').get(req.params.id);
  if (!f) return res.status(404).json({ error: 'Not found' });
  // Stash entity_id so workpaperStorage.destination can route the file to the right entity dir
  req.entityId = f.entity_id;
  // Re-use workpaperStorage so the new file lands in the correct entity dir
  workpaperUpload.single('file')(req, res, err => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    // Delete the old file from disk (best-effort)
    try { fs.unlinkSync(path.join(WORKPAPERS_DIR, String(f.entity_id), f.stored_filename)); } catch {}
    const uploader = req.user.name || req.user.email;
    const now = new Date().toISOString();
    db.prepare('UPDATE entity_files SET stored_filename=?, original_name=?, size=?, mime_type=?, uploaded_by=?, created_at=? WHERE id=?')
      .run(req.file.filename, req.file.originalname, req.file.size, req.file.mimetype, uploader, now, f.id);
    const updated = db.prepare('SELECT id, folder_path, original_name, size, mime_type, uploaded_by, created_at FROM entity_files WHERE id=?').get(f.id);
    res.json({ success: true, file: updated });
  });
});

app.post('/api/entities/:eid/folders', auth, requireRole('Admin','Accountant'), (req, res) => {
  const folder = normFolderPath(req.body.folder_path);
  if (!folder) return res.status(400).json({ error: 'Folder path required' });
  try {
    db.prepare('INSERT INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)').run(req.params.eid, folder, req.user.name || req.user.email);
    res.json({ success: true, folder_path: folder });
  } catch (e) {
    if (e.message.includes('UNIQUE')) return res.json({ success: true, folder_path: folder }); // already exists — no-op
    throw e;
  }
});

app.delete('/api/entities/:eid/folders', auth, requireRole('Admin','Accountant'), (req, res) => {
  const folder = normFolderPath(req.query.folder_path);
  if (!folder) return res.status(400).json({ error: 'Folder path required' });
  // Safety: only allow deleting a folder when it has no files or subfolders underneath it
  const childFiles = db.prepare("SELECT COUNT(*) as c FROM entity_files WHERE entity_id=? AND (folder_path=? OR folder_path LIKE ?)").get(req.params.eid, folder, folder + '/%').c;
  const childFolders = db.prepare("SELECT COUNT(*) as c FROM entity_folders WHERE entity_id=? AND folder_path LIKE ?").get(req.params.eid, folder + '/%').c;
  if (childFiles > 0 || childFolders > 0) return res.status(400).json({ error: 'Folder is not empty' });
  db.prepare('DELETE FROM entity_folders WHERE entity_id=? AND folder_path=?').run(req.params.eid, folder);
  res.json({ success: true });
});

app.put('/api/entity-files/:id/move', auth, requireRole('Admin','Accountant'), (req, res) => {
  const folder = normFolderPath(req.body.folder_path);
  const f = db.prepare('SELECT * FROM entity_files WHERE id=?').get(req.params.id);
  if (!f) return res.status(404).json({ error: 'Not found' });
  db.prepare('UPDATE entity_files SET folder_path=? WHERE id=?').run(folder, req.params.id);
  res.json({ success: true });
});

// Rename a folder — updates every file under it, every nested folder row, and the folder row itself.
app.put('/api/entities/:eid/folders/rename', auth, requireRole('Admin','Accountant'), (req, res) => {
  const oldPath = normFolderPath(req.body.old_path);
  const newPath = normFolderPath(req.body.new_path);
  if (!oldPath || !newPath) return res.status(400).json({ error: 'Both old_path and new_path are required' });
  if (oldPath === newPath) return res.json({ success: true, unchanged: true });
  const eid = req.params.eid;

  // Guard: can't rename into a subpath of itself
  if (newPath === oldPath || newPath.startsWith(oldPath + '/')) return res.status(400).json({ error: 'Cannot move a folder into itself' });

  // Guard: target folder must not already exist
  const collision = db.prepare('SELECT 1 FROM entity_folders WHERE entity_id=? AND folder_path=?').get(eid, newPath)
    || db.prepare('SELECT 1 FROM entity_files WHERE entity_id=? AND folder_path=? LIMIT 1').get(eid, newPath)
    || db.prepare('SELECT 1 FROM entity_folders WHERE entity_id=? AND folder_path LIKE ? LIMIT 1').get(eid, newPath + '/%')
    || db.prepare('SELECT 1 FROM entity_files WHERE entity_id=? AND folder_path LIKE ? LIMIT 1').get(eid, newPath + '/%');
  if (collision) return res.status(400).json({ error: 'A folder with that name already exists at the target location' });

  try {
    db.transaction(() => {
      // Rename the folder row itself (if present)
      db.prepare('UPDATE entity_folders SET folder_path=? WHERE entity_id=? AND folder_path=?').run(newPath, eid, oldPath);
      // Rename every nested folder row: oldPath/anything -> newPath/anything
      const nestedFolders = db.prepare('SELECT id, folder_path FROM entity_folders WHERE entity_id=? AND folder_path LIKE ?').all(eid, oldPath + '/%');
      const updFolder = db.prepare('UPDATE entity_folders SET folder_path=? WHERE id=?');
      for (const nf of nestedFolders) updFolder.run(newPath + nf.folder_path.slice(oldPath.length), nf.id);
      // Rename files in the folder itself
      db.prepare('UPDATE entity_files SET folder_path=? WHERE entity_id=? AND folder_path=?').run(newPath, eid, oldPath);
      // Rename files in nested folders
      const nestedFiles = db.prepare('SELECT id, folder_path FROM entity_files WHERE entity_id=? AND folder_path LIKE ?').all(eid, oldPath + '/%');
      const updFile = db.prepare('UPDATE entity_files SET folder_path=? WHERE id=?');
      for (const nf of nestedFiles) updFile.run(newPath + nf.folder_path.slice(oldPath.length), nf.id);
    })();
    res.json({ success: true, old_path: oldPath, new_path: newPath });
  } catch (e) {
    res.status(400).json({ error: 'Rename failed: ' + e.message });
  }
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

// === Bill.com integration routes ===
app.get('/api/billcom/config/:entity_id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT entity_id, environment, api_base_url, username, password_enc, org_id, dev_key_enc, default_ap_account, default_cash_account, last_tested_at, last_test_status, last_test_message, updated_by, updated_at FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
  if (!row) return res.json({ configured: false });
  let pwLast4 = '', keyLast4 = '';
  try { pwLast4 = maskSecret(billcomDecrypt(row.password_enc)); } catch {}
  try { keyLast4 = maskSecret(billcomDecrypt(row.dev_key_enc)); } catch {}
  res.json({
    configured: true,
    entity_id: row.entity_id,
    environment: row.environment,
    api_base_url: row.api_base_url,
    username: row.username,
    password_masked: pwLast4,
    org_id: row.org_id,
    dev_key_masked: keyLast4,
    default_ap_account: row.default_ap_account,
    default_cash_account: row.default_cash_account,
    last_tested_at: row.last_tested_at,
    last_test_status: row.last_test_status,
    last_test_message: row.last_test_message,
    updated_by: row.updated_by,
    updated_at: row.updated_at
  });
});

app.put('/api/billcom/config/:entity_id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { environment, username, password, org_id, dev_key, default_ap_account, default_cash_account } = req.body || {};
  if (!environment || !username || !org_id) return res.status(400).json({ error: 'environment, username, org_id required' });
  if (!['sandbox','production'].includes(environment)) return res.status(400).json({ error: 'environment must be sandbox or production' });
  const baseUrl = BILLCOM_BASE_URLS[environment];
  const existing = db.prepare('SELECT password_enc, dev_key_enc FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
  let pwEnc, keyEnc;
  try {
    pwEnc = password ? billcomEncrypt(password) : (existing ? existing.password_enc : null);
    keyEnc = dev_key ? billcomEncrypt(dev_key) : (existing ? existing.dev_key_enc : null);
  } catch (e) { return res.status(500).json({ error: 'Encryption failed: ' + e.message }); }
  if (!pwEnc || !keyEnc) return res.status(400).json({ error: 'password and dev_key required for first save' });
  const now = new Date().toISOString();
  const updater = req.user.name || req.user.email;
  if (existing) {
    db.prepare('UPDATE billcom_config SET environment=?, api_base_url=?, username=?, password_enc=?, org_id=?, dev_key_enc=?, default_ap_account=?, default_cash_account=?, updated_by=?, updated_at=? WHERE entity_id=?')
      .run(environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, updater, now, req.params.entity_id);
  } else {
    db.prepare('INSERT INTO billcom_config (entity_id, environment, api_base_url, username, password_enc, org_id, dev_key_enc, default_ap_account, default_cash_account, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)')
      .run(req.params.entity_id, environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, updater, now);
  }
  res.json({ success: true });
});

app.delete('/api/billcom/config/:entity_id', auth, requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM billcom_config WHERE entity_id = ?').run(req.params.entity_id);
  res.json({ success: true });
});

app.post('/api/billcom/config/:entity_id/test', auth, requireRole('Admin','Accountant'), async (req, res) => {
  const row = db.prepare('SELECT environment, api_base_url, username, password_enc, org_id, dev_key_enc FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
  if (!row) return res.status(404).json({ error: 'No config found for this entity' });
  let password, devKey;
  try {
    password = billcomDecrypt(row.password_enc);
    devKey   = billcomDecrypt(row.dev_key_enc);
  } catch (e) { return res.status(500).json({ error: 'Decryption failed: ' + e.message }); }
  const now = new Date().toISOString();
  try {
    const result = await billcomLogin({
      username: row.username, password, orgId: row.org_id, devKey, baseUrl: row.api_base_url
    });
    const sessionLen = (result.sessionId || '').length;
    const msg = 'Login OK. sessionId received (' + sessionLen + ' chars), userId=' + (result.userId || 'n/a');
    db.prepare('UPDATE billcom_config SET last_tested_at=?, last_test_status=?, last_test_message=? WHERE entity_id=?')
      .run(now, 'success', msg, req.params.entity_id);
    res.json({ success: true, message: msg, organizationId: result.organizationId, userId: result.userId });
  } catch (e) {
    const msg = e.message || 'Unknown error';
    db.prepare('UPDATE billcom_config SET last_tested_at=?, last_test_status=?, last_test_message=? WHERE entity_id=?')
      .run(now, 'failed', msg, req.params.entity_id);
    res.status(400).json({ success: false, error: msg });
  }
});

// ── Bill.com Phase 2: Chart of Accounts + Mappings ──

app.get('/api/billcom/accounts/:entity_id', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
  const row = db.prepare('SELECT environment, api_base_url, username, password_enc, org_id, dev_key_enc FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
  if (!row) return res.status(404).json({ error: 'No Bill.com config for this entity. Save credentials first.' });
  let password, devKey;
  try {
    password = billcomDecrypt(row.password_enc);
    devKey   = billcomDecrypt(row.dev_key_enc);
  } catch (e) { return res.status(500).json({ error: 'Decryption failed: ' + e.message }); }
  try {
    const login = await billcomLogin({ username: row.username, password, orgId: row.org_id, devKey, baseUrl: row.api_base_url });
    if (!login.sessionId) return res.status(502).json({ error: 'Bill.com login returned no sessionId' });
    const accounts = await billcomListAccounts({ sessionId: login.sessionId, devKey, baseUrl: row.api_base_url });
    if (accounts.length > 0) console.log('[billcom COA] sample shape: ' + JSON.stringify(accounts[0]).slice(0, 600));
    // Return a slim shape - tolerate both v2 and v3 field naming
    const slim = accounts.map(a => ({
      id: a.id,
      name: a.name,
      accountNumber: a.accountNumber || a.number || '',
      accountType: a.accountType || a.type || '',
      description: a.description || '',
      isActive: (a.isActive === '1' || a.isActive === true || a.active === true || a.status === 'ACTIVE')
    }));
    res.json({ accounts: slim, count: slim.length });
  } catch (e) {
    res.status(400).json({ error: e.message || 'Failed to list Bill.com accounts' });
  }
});

app.get('/api/billcom/mappings/:entity_id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const rows = db.prepare('SELECT id, billcom_account_id, billcom_account_name, cl_account_code, created_at FROM billcom_account_map WHERE entity_id = ? ORDER BY id').all(req.params.entity_id);
  res.json({ mappings: rows });
});

app.put('/api/billcom/mappings/:entity_id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });
  const mappings = Array.isArray(req.body && req.body.mappings) ? req.body.mappings : null;
  if (!mappings) return res.status(400).json({ error: 'Body must include mappings: array' });
  // Validate each row
  for (const m of mappings) {
    if (!m.billcom_account_id || !m.cl_account_code) {
      return res.status(400).json({ error: 'Each mapping needs billcom_account_id and cl_account_code' });
    }
  }
  const now = new Date().toISOString();
  const tx = db.transaction((rows) => {
    db.prepare('DELETE FROM billcom_account_map WHERE entity_id = ?').run(entityId);
    const ins = db.prepare('INSERT INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)');
    for (const m of rows) {
      ins.run(entityId, String(m.billcom_account_id), m.billcom_account_name || null, String(m.cl_account_code), now);
    }
  });
  try {
    tx(mappings);
    res.json({ saved: mappings.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// TEMP: Phase 6 verification. Creates a payment for an unpaid bill.
// Remove after verification.
app.post('/api/billcom/_seed_payment/:entity_id', auth, requireRole('Admin'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured' });
  let session, devKey;
  try {
    const pw = billcomDecrypt(cfg.password_enc);
    devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password: pw, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (ex) { return res.status(502).json({ error: 'login: ' + ex.message }); }
  const headers = { 'Content-Type': 'application/json', 'sessionId': session.sessionId, 'devKey': devKey, 'Accept': 'application/json' };
  const base = cfg.api_base_url;
  const dryRun = req.body && req.body.dry_run;
  const reqBillId = req.body && req.body.billId;
  const reqFundingId = req.body && req.body.fundingAccountId;

  // 1. List bills.
  let bills = [];
  try {
    const lr = await fetch(base + '/bills?max=100', { headers });
    const lt = await lr.text(); let lj; try { lj = JSON.parse(lt); } catch { lj = null; }
    bills = (lj && (Array.isArray(lj.results) ? lj.results : (Array.isArray(lj) ? lj : []))) || [];
  } catch (ex) { return res.status(502).json({ error: 'bill list: ' + ex.message }); }

  // 2. Pick target bill.
  let chosen = null;
  if (reqBillId) chosen = bills.find(b => b.id === reqBillId);
  else chosen = bills.find(b => {
    const ps = String(b.paymentStatus || '').toUpperCase();
    return (ps === 'UNPAID' || ps === '') && Number(b.amount || 0) > 0;
  });

  // 3. Discover funding accounts via /payments/options.
  let fundingOptions = null, fundingAccountId = null, fundingAccountType = 'BANK_ACCOUNT';
  if (chosen) {
    try {
      const ur = base + '/payments/options?vendorId=' + encodeURIComponent(chosen.vendorId) + '&amount=' + encodeURIComponent(String(Number(chosen.amount || 0)));
      const opt = await fetch(ur, { headers });
      const ot = await opt.text(); let oj; try { oj = JSON.parse(ot); } catch { oj = null; }
      fundingOptions = oj;
      // Bill.com nests bank list under fundingOptions.banks (no fundingAccount wrapper).
      const nested = oj && oj.fundingOptions && oj.fundingOptions.banks;
      const flat = oj && Array.isArray(oj.options) ? oj.options : (Array.isArray(oj) ? oj : []);
      let pickedBank = null;
      if (Array.isArray(nested)) pickedBank = nested.find(b => b && b.available && b.id);
      if (!pickedBank && flat.length) {
        const avail = flat.find(o => o && o.available && o.fundingAccount && o.fundingAccount.id);
        if (avail) pickedBank = { id: avail.fundingAccount.id, type: avail.fundingAccount.type };
      }
      if (pickedBank) { fundingAccountId = pickedBank.id; fundingAccountType = pickedBank.type || fundingAccountType; }
    } catch (ex) { fundingOptions = { error: ex.message }; }
  }

  if (dryRun) {
    return res.json({
      dry_run: true,
      bill_count: bills.length,
      bills: bills.map(b => ({ id: b.id, vendorId: b.vendorId, amount: b.amount, paymentStatus: b.paymentStatus })),
      chosen: chosen && { id: chosen.id, vendorId: chosen.vendorId, amount: chosen.amount },
      fundingOptions: fundingOptions,
      fundingAccountId: fundingAccountId
    });
  }

  if (!chosen) return res.status(502).json({ error: 'no eligible bill found' });

  if (!fundingAccountId) {
    // Try direct list of bank funding accounts.
    try {
      const br = await fetch(base + '/funding-accounts/banks?max=20', { headers });
      const bt = await br.text(); let bj; try { bj = JSON.parse(bt); } catch { bj = null; }
      const banks = bj && (Array.isArray(bj.results) ? bj.results : (Array.isArray(bj) ? bj : []));
      const verified = (banks || []).find(b => b && (b.status === 'VERIFIED' || b.status === 'ACTIVE' || b.verified === true));
      if (verified) { fundingAccountId = verified.id; fundingAccountType = 'BANK_ACCOUNT'; }
    } catch {}
  }
  if (reqFundingId) { fundingAccountId = reqFundingId; fundingAccountType = (req.body && req.body.fundingAccountType) || fundingAccountType; }
  if (!fundingAccountId) return res.status(502).json({ error: 'no funding account found', fundingOptions: fundingOptions });

  // 4. Create payment.
  const today = new Date().toISOString().slice(0, 10);
  const body = {
    vendorId: chosen.vendorId,
    billId: chosen.id,
    processDate: today,
    amount: Number(chosen.amount || 0),
    fundingAccount: { type: fundingAccountType, id: fundingAccountId },
    processingOptions: { requestPayFaster: false, createBill: false }
  };
  try {
    const pr = await fetch(base + '/payments', { method: 'POST', headers, body: JSON.stringify(body) });
    const pt = await pr.text(); let pj; try { pj = JSON.parse(pt); } catch { pj = null; }
    if (!pr.ok) return res.status(502).json({ error: 'payment create failed: HTTP ' + pr.status + ' :: ' + pt.slice(0, 1000), payload: body });
    res.json({ success: true, chosenBillId: chosen.id, fundingAccountId: fundingAccountId, payment: pj });
  } catch (ex) { return res.status(502).json({ error: 'payment exception: ' + ex.message }); }
});

// Phase 5: Push CloudLedger COA to Bill.com and auto-create mappings.
app.post('/api/billcom/push-coa/:entity_id', auth, requireRole('Admin'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured' });

  const body = req.body || {};
  let rows;
  if (Array.isArray(body.codes) && body.codes.length > 0) {
    const placeholders = body.codes.map(function(){ return '?'; }).join(',');
    rows = db.prepare('SELECT code, name, type, subtype FROM accounts WHERE entity_id = ? AND code IN (' + placeholders + ')').all(entityId, ...body.codes);
  } else if (body.all_expenses) {
    rows = db.prepare("SELECT code, name, type, subtype FROM accounts WHERE entity_id = ? AND type = 'Expense' ORDER BY code").all(entityId);
  } else {
    return res.status(400).json({ error: 'Provide either codes:[...] or all_expenses:true' });
  }
  if (rows.length === 0) return res.status(404).json({ error: 'No matching CL accounts found' });

  let session, devKey, existing;
  try {
    const pw = billcomDecrypt(cfg.password_enc);
    devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password: pw, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
    existing = await billcomListAccounts({ sessionId: session.sessionId, devKey, baseUrl: cfg.api_base_url });
  } catch (ex) { return res.status(502).json({ error: 'login or list failed: ' + ex.message }); }

  const byName = new Map();
  const byNum = new Map();
  for (const a of existing) {
    if (a && a.name) byName.set(String(a.name).trim().toLowerCase(), a);
    const num = a && (a.accountNumber || a.number);
    if (num) byNum.set(String(num), a);
  }

  function mapType(clType, subtype) {
    const t = String(clType || '').toLowerCase();
    const s = String(subtype || '').toLowerCase();
    if (t === 'expense') {
      if (s === 'cogs' || s.indexOf('cost of goods') >= 0) return 'COST_OF_GOODS_SOLD';
      return 'EXPENSE';
    }
    if (t === 'income' || t === 'revenue') return 'INCOME';
    if (t === 'asset') {
      if (s.indexOf('fixed') >= 0) return 'FIXED_ASSET';
      if (s.indexOf('bank') >= 0) return 'BANK';
      return 'OTHER_ASSET';
    }
    if (t === 'liability') {
      if (s.indexOf('current') >= 0) return 'LIABILITY';
      return 'OTHER_LIABILITY';
    }
    if (t === 'equity') return 'EQUITY';
    return 'OTHER_EXPENSE';
  }

  const headers = { 'Content-Type': 'application/json', 'sessionId': session.sessionId, 'devKey': devKey, 'Accept': 'application/json' };
  const base = cfg.api_base_url;

  const existingMap = db.prepare('SELECT cl_account_code, billcom_account_id FROM billcom_account_map WHERE entity_id = ?').all(entityId);
  const mappedClCodes = new Set(existingMap.map(function(m){ return m.cl_account_code; }));

  const out = { pushed: [], skipped_existing: [], mapped_only: [], errors: [] };

  for (const row of rows) {
    try {
      let bcAccount = byNum.get(row.code) || byName.get(String(row.name).trim().toLowerCase());
      if (bcAccount) {
        if (!mappedClCodes.has(row.code)) {
          db.prepare('INSERT INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)')
            .run(entityId, bcAccount.id, bcAccount.name, row.code, new Date().toISOString());
          out.mapped_only.push({ code: row.code, name: row.name, billcom_id: bcAccount.id });
        } else {
          out.skipped_existing.push({ code: row.code, name: row.name, billcom_id: bcAccount.id });
        }
        continue;
      }
      const payload = {
        name: row.name,
        account: { accountNumber: row.code, type: mapType(row.type, row.subtype) }
      };
      const r = await fetch(base + '/classifications/chart-of-accounts', { method: 'POST', headers, body: JSON.stringify(payload) });
      const txt = await r.text();
      let j; try { j = JSON.parse(txt); } catch { j = null; }
      if (!r.ok || !j || !j.id) {
        out.errors.push({ code: row.code, name: row.name, status: r.status, error: txt.slice(0, 400) });
        continue;
      }
      db.prepare('INSERT INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)')
        .run(entityId, j.id, j.name || row.name, row.code, new Date().toISOString());
      out.pushed.push({ code: row.code, name: row.name, billcom_id: j.id });
    } catch (ex) {
      out.errors.push({ code: row.code, name: row.name, error: ex.message });
    }
  }
  res.json(out);
});

// Phase 3: Bill.com sync (bills + payments -> JEs)
app.get('/api/billcom/sync-log/:entity_id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  const limit = Math.min(parseInt(req.query.limit) || 50, 500);
  const rows = db.prepare(
    'SELECT id, sync_type, billcom_id, cl_entry_id, status, message, created_at FROM billcom_sync_log WHERE entity_id = ? ORDER BY id DESC LIMIT ?'
  ).all(entityId, limit);
  res.json({ logs: rows });
});

app.post('/api/billcom/sync/:entity_id', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });

  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });
  const entity = db.prepare('SELECT * FROM entities WHERE id = ?').get(entityId);
  if (!entity) return res.status(404).json({ error: 'Entity not found' });

  const apAccount = cfg.default_ap_account;
  const cashAccount = cfg.default_cash_account;
  if (!apAccount) return res.status(400).json({ error: 'default_ap_account not set in Bill.com config' });
  if (!cashAccount) return res.status(400).json({ error: 'default_cash_account not set in Bill.com config' });

  const apExists = db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, apAccount);
  const cashExists = db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, cashAccount);
  if (!apExists) return res.status(400).json({ error: 'AP account ' + apAccount + ' does not exist on entity' });
  if (!cashExists) return res.status(400).json({ error: 'Cash account ' + cashAccount + ' does not exist on entity' });

  const mapRows = db.prepare('SELECT billcom_account_id, billcom_account_name, cl_account_code FROM billcom_account_map WHERE entity_id = ?').all(entityId);
  const mapById = new Map(mapRows.map(r => [String(r.billcom_account_id), r]));

  let session;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    const devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) {
    return res.status(502).json({ error: 'Bill.com login failed: ' + e.message });
  }
  const listArgs = { sessionId: session.sessionId, devKey: billcomDecrypt(cfg.dev_key_enc), baseUrl: cfg.api_base_url };

  let bills, payments;
  try {
    bills = await billcomListBills(listArgs);
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch bills: ' + e.message });
  }
  try {
    payments = await billcomListPayments(listArgs);
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch payments: ' + e.message });
  }

  const pick = (obj, ...keys) => { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; };
  // Bills: process when approval is complete (APPROVED) or not gated (UNASSIGNED, no policy),
  // or when already paid; skip while waiting on approvers (ASSIGNED) or denied (DENIED).
  const isBillEligible = (bill) => {
    const a = String(pick(bill, 'approvalStatus', 'status') || '').toUpperCase();
    const p = String(pick(bill, 'paymentStatus') || '').toUpperCase();
    if (a === 'DENIED' || a === 'ASSIGNED') return false;
    if (a === 'APPROVED' || a === 'UNASSIGNED' || a === '') return true;
    if (p === 'PAID' || p === 'PARTIAL_PAID' || p === 'PARTIALLY_PAID') return true;
    return false;
  };
  // Payments: only process if actually disbursed (or scheduled to be).
  const isPaymentEligible = (pay) => {
    const s = String(pick(pay, 'paymentStatus', 'status') || '').toUpperCase();
    return s === 'PAID' || s === 'SCHEDULED' || s === 'PROCESSING' || s === 'SENT';
  };

  const result = {
    bills: { synced: 0, skipped: 0, errors: 0, details: [] },
    payments: { synced: 0, skipped: 0, errors: 0, details: [] },
    missing_mappings: []
  };
  const missingMap = new Map();
  const now = new Date().toISOString();
  const actor = req.user.name || req.user.email || 'system';

  const logSync = db.prepare(
    'INSERT INTO billcom_sync_log (entity_id, sync_type, billcom_id, cl_entry_id, status, message, created_at) VALUES (?,?,?,?,?,?,?)'
  );
  const alreadySynced = db.prepare(
    "SELECT 1 FROM billcom_sync_log WHERE entity_id = ? AND sync_type = ? AND billcom_id = ? AND status = 'success' LIMIT 1"
  );

  for (const bill of bills) {
    const billId = String(pick(bill, 'id') || '');
    if (!billId) continue;
    if (!isBillEligible(bill)) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'not approved' });
      continue;
    }
    if (alreadySynced.get(entityId, 'bill', billId)) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'already synced' });
      continue;
    }

    // List endpoint omits some fields (notably chartOfAccountId on line items); hydrate from detail.
    let detail = bill;
    try {
      detail = await billcomGetById({ ...listArgs, resourcePath: '/bills', id: billId });
    } catch (e) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'detail fetch failed: ' + e.message, now);
      result.bills.details.push({ id: billId, status: 'error', reason: 'detail fetch failed: ' + e.message });
      continue;
    }
    const invoiceDate = pick(detail, 'invoiceDate', 'invoice_date') || pick(pick(detail, 'invoice') || {}, 'invoiceDate', 'invoice_date') || pick(detail, 'dueDate');
    const billNumber = pick(detail, 'invoiceNumber', 'invoice_number') || pick(pick(detail, 'invoice') || {}, 'invoiceNumber', 'invoice_number') || billId;
    const lineItems = pick(detail, 'lineItems', 'line_items', 'billLineItems') || [];

    if (!invoiceDate || !Array.isArray(lineItems) || lineItems.length === 0) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'missing invoiceDate or lineItems', now);
      result.bills.details.push({ id: billId, status: 'error', reason: 'missing invoiceDate or lineItems' });
      continue;
    }

    const debitLines = [];
    const billMissing = [];
    let totalDr = 0;
    for (const li of lineItems) {
      const cls = pick(li, 'classifications') || {};
      const acctId = String(pick(li, 'chartOfAccountId', 'chart_of_account_id', 'accountId', 'account_id') || pick(cls, 'chartOfAccountId', 'chart_of_account_id', 'accountId', 'account_id') || '');
      const amt = Number(pick(li, 'amount', 'value') || 0);
      if (!acctId) {
        billMissing.push({ id: '(none)', name: '(no chartOfAccount on line)' });
        continue;
      }
      const mapping = mapById.get(acctId);
      if (!mapping) {
        const name = pick(li, 'chartOfAccountName', 'chart_of_account_name', 'accountName') || pick(cls, 'chartOfAccountName', 'chart_of_account_name', 'accountName') || acctId;
        billMissing.push({ id: acctId, name });
        continue;
      }
      debitLines.push({ account_code: mapping.cl_account_code, debit: amt, credit: 0 });
      totalDr += amt;
    }

    if (billMissing.length > 0) {
      result.bills.errors++;
      const missingNames = billMissing.map(m => m.name).join(', ');
      logSync.run(entityId, 'bill', billId, null, 'error', 'missing GL mapping(s) for: ' + missingNames, now);
      result.bills.details.push({ id: billId, status: 'error', reason: 'missing mappings: ' + missingNames });
      for (const mm of billMissing) {
        const existing = missingMap.get(mm.id);
        if (existing) existing.affected_bills++;
        else missingMap.set(mm.id, { billcom_account_id: mm.id, name: mm.name, affected_bills: 1 });
      }
      continue;
    }

    if (totalDr <= 0) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'bill total is zero', now);
      result.bills.details.push({ id: billId, status: 'error', reason: 'zero total' });
      continue;
    }

    const lines = [...debitLines, { account_code: apAccount, debit: 0, credit: totalDr }];
    const memo = 'Bill.com bill #' + billNumber;

    try {
      const insertedId = db.transaction(() => {
        const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
        const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
          .run(entityId, num, invoiceDate, memo, actor);
        for (const l of lines) {
          db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)')
            .run(r.lastInsertRowid, l.account_code, l.debit, l.credit);
        }
        logSync.run(entityId, 'bill', billId, r.lastInsertRowid, 'success', 'created JE #' + num, now);
        return r.lastInsertRowid;
      })();
      result.bills.synced++;
      result.bills.details.push({ id: billId, status: 'success', cl_entry_id: insertedId });
    } catch (e) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'JE insert failed: ' + e.message, now);
      result.bills.details.push({ id: billId, status: 'error', reason: e.message });
    }
  }

  for (const pay of payments) {
    const payId = String(pick(pay, 'id') || '');
    if (!payId) continue;
    if (!isPaymentEligible(pay)) {
      result.payments.skipped++;
      result.payments.details.push({ id: payId, status: 'skip', reason: 'not approved' });
      continue;
    }
    if (alreadySynced.get(entityId, 'payment', payId)) {
      result.payments.skipped++;
      result.payments.details.push({ id: payId, status: 'skip', reason: 'already synced' });
      continue;
    }

    const processDate = pick(pay, 'processDate', 'process_date', 'paymentDate', 'payment_date');
    const amount = Number(pick(pay, 'amount', 'paymentAmount', 'totalAmount') || 0);
    const payNumber = pick(pay, 'paymentNumber', 'payment_number', 'referenceNumber') || payId;

    if (!processDate || amount <= 0) {
      result.payments.errors++;
      logSync.run(entityId, 'payment', payId, null, 'error', 'missing processDate or zero amount', now);
      result.payments.details.push({ id: payId, status: 'error', reason: 'missing processDate or zero amount' });
      continue;
    }

    const lines = [
      { account_code: apAccount, debit: amount, credit: 0 },
      { account_code: cashAccount, debit: 0, credit: amount }
    ];
    const memo = 'Bill.com payment #' + payNumber;

    try {
      const insertedId = db.transaction(() => {
        const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
        const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
          .run(entityId, num, processDate, memo, actor);
        for (const l of lines) {
          db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)')
            .run(r.lastInsertRowid, l.account_code, l.debit, l.credit);
        }
        logSync.run(entityId, 'payment', payId, r.lastInsertRowid, 'success', 'created JE #' + num, now);
        return r.lastInsertRowid;
      })();
      result.payments.synced++;
      result.payments.details.push({ id: payId, status: 'success', cl_entry_id: insertedId });
    } catch (e) {
      result.payments.errors++;
      logSync.run(entityId, 'payment', payId, null, 'error', 'JE insert failed: ' + e.message, now);
      result.payments.details.push({ id: payId, status: 'error', reason: e.message });
    }
  }

  result.missing_mappings = Array.from(missingMap.values());
  res.json(result);
});

if (process.env.NODE_ENV === 'production') app.get('*', (req, res) => {
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.sendFile(path.join(__dirname, '..', 'client', 'dist', 'index.html'));
});
app.listen(PORT, '0.0.0.0', () => console.log(`CloudLedger on port ${PORT}`));
