require('dotenv').config();
const express = require('express');
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-to-a-real-secret-in-production';
const DB_PATH = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'cloudledger.db');

// Ensure data directory exists
const dataDir = path.dirname(DB_PATH);
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

// ─── Database Setup ───
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');     // Better concurrent access
db.pragma('foreign_keys = ON');

// Create tables
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'Viewer',
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS entities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    code TEXT NOT NULL,
    name TEXT NOT NULL,
    type TEXT NOT NULL,
    subtype TEXT DEFAULT '',
    bank_acct INTEGER DEFAULT 0,
    UNIQUE(entity_id, code)
  );

  CREATE TABLE IF NOT EXISTS journal_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    entry_num INTEGER NOT NULL,
    date TEXT NOT NULL,
    memo TEXT NOT NULL,
    created_by TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS journal_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    debit REAL DEFAULT 0,
    credit REAL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS cleared_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    entry_id INTEGER NOT NULL,
    line_index INTEGER NOT NULL,
    reconciliation_id INTEGER,
    UNIQUE(entity_id, account_code, entry_id, line_index)
  );

  CREATE TABLE IF NOT EXISTS reconciliations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    statement_date TEXT NOT NULL,
    statement_balance REAL NOT NULL,
    book_balance REAL NOT NULL,
    cleared_count INTEGER DEFAULT 0,
    completed_by TEXT NOT NULL,
    completed_at TEXT DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_accounts_entity ON accounts(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_entity ON journal_entries(entity_id);
  CREATE INDEX IF NOT EXISTS idx_jl_entry ON journal_lines(entry_id);
  CREATE INDEX IF NOT EXISTS idx_cleared_entity ON cleared_items(entity_id, account_code);
  CREATE INDEX IF NOT EXISTS idx_rec_entity ON reconciliations(entity_id);
`);

// Seed default admin if no users exist
const userCount = db.prepare('SELECT COUNT(*) as c FROM users').get();
if (userCount.c === 0) {
  const hash = bcrypt.hashSync('admin', 10);
  db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?, ?, ?, ?)').run('Admin', 'admin@company.com', hash, 'Admin');
  console.log('Default admin created: admin@company.com / admin');
}

// Default COA template
const DEFAULT_COA = [
  { code:"1000",name:"Cash",type:"Asset",subtype:"Current Asset",bank:1 },
  { code:"1010",name:"Operating Checking",type:"Asset",subtype:"Current Asset",bank:1 },
  { code:"1020",name:"Savings Account",type:"Asset",subtype:"Current Asset",bank:1 },
  { code:"1100",name:"Accounts Receivable",type:"Asset",subtype:"Current Asset",bank:0 },
  { code:"1200",name:"Inventory",type:"Asset",subtype:"Current Asset",bank:0 },
  { code:"1300",name:"Prepaid Expenses",type:"Asset",subtype:"Current Asset",bank:0 },
  { code:"1500",name:"Property & Equipment",type:"Asset",subtype:"Fixed Asset",bank:0 },
  { code:"1510",name:"Accumulated Depreciation",type:"Asset",subtype:"Fixed Asset",bank:0 },
  { code:"2000",name:"Accounts Payable",type:"Liability",subtype:"Current Liability",bank:0 },
  { code:"2100",name:"Accrued Liabilities",type:"Liability",subtype:"Current Liability",bank:0 },
  { code:"2200",name:"Unearned Revenue",type:"Liability",subtype:"Current Liability",bank:0 },
  { code:"2500",name:"Notes Payable",type:"Liability",subtype:"Long-term Liability",bank:0 },
  { code:"3000",name:"Common Stock / Member's Capital",type:"Equity",subtype:"Equity",bank:0 },
  { code:"3100",name:"Retained Earnings",type:"Equity",subtype:"Equity",bank:0 },
  { code:"3200",name:"Additional Paid-in Capital",type:"Equity",subtype:"Equity",bank:0 },
  { code:"4000",name:"Revenue",type:"Revenue",subtype:"Operating Revenue",bank:0 },
  { code:"4100",name:"Service Revenue",type:"Revenue",subtype:"Operating Revenue",bank:0 },
  { code:"4200",name:"Interest Income",type:"Revenue",subtype:"Other Revenue",bank:0 },
  { code:"5000",name:"Cost of Goods Sold",type:"Expense",subtype:"COGS",bank:0 },
  { code:"6000",name:"Salaries Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6100",name:"Rent Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6200",name:"Utilities Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6300",name:"Depreciation Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6400",name:"Insurance Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6500",name:"Office Supplies Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"6600",name:"Marketing Expense",type:"Expense",subtype:"Operating Expense",bank:0 },
  { code:"7000",name:"Interest Expense",type:"Expense",subtype:"Other Expense",bank:0 },
];

// ─── Middleware ───
app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(morgan('short'));
app.use(express.json({ limit: '10mb' }));

// Serve static frontend in production
if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, '..', 'client', 'dist')));
}

// Auth middleware
function auth(req, res, next) {
  const token = req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token provided' });
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch { return res.status(401).json({ error: 'Invalid token' }); }
}

function requireRole(...roles) {
  return (req, res, next) => {
    if (!roles.includes(req.user.role) && req.user.role !== 'Admin') {
      return res.status(403).json({ error: 'Insufficient permissions' });
    }
    next();
  };
}

// ─── Auth Routes ───
app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body;
  const user = db.prepare('SELECT * FROM users WHERE email = ?').get(email?.toLowerCase());
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid credentials' });
  }
  const token = jwt.sign({ id: user.id, name: user.name, email: user.email, role: user.role }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
});

app.post('/api/auth/signup', (req, res) => {
  const { name, email, password, role } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'All fields required' });
  if (password.length < 3) return res.status(400).json({ error: 'Password too short' });
  const existing = db.prepare('SELECT id FROM users WHERE email = ?').get(email.toLowerCase());
  if (existing) return res.status(400).json({ error: 'Email already exists' });
  const hash = bcrypt.hashSync(password, 10);
  const validRole = ['Admin', 'Accountant', 'Viewer'].includes(role) ? role : 'Viewer';
  const result = db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?, ?, ?, ?)').run(name, email.toLowerCase(), hash, validRole);
  res.json({ id: result.lastInsertRowid, message: 'Account created' });
});

app.get('/api/auth/me', auth, (req, res) => { res.json(req.user); });

// ─── User Routes ───
app.get('/api/users', auth, requireRole('Admin'), (req, res) => {
  const users = db.prepare('SELECT id, name, email, role, created_at FROM users').all();
  res.json(users);
});

app.delete('/api/users/:id', auth, requireRole('Admin'), (req, res) => {
  if (parseInt(req.params.id) === req.user.id) return res.status(400).json({ error: 'Cannot delete yourself' });
  db.prepare('DELETE FROM users WHERE id = ?').run(req.params.id);
  res.json({ success: true });
});

app.put('/api/users/:id', auth, requireRole('Admin'), (req, res) => {
  const { name, role } = req.body;
  db.prepare('UPDATE users SET name = ?, role = ? WHERE id = ?').run(name, role, req.params.id);
  res.json({ success: true });
});

// ─── Entity Routes ───
app.get('/api/entities', auth, (req, res) => {
  const entities = db.prepare('SELECT * FROM entities ORDER BY code').all();
  res.json(entities);
});

app.post('/api/entities', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const { code, name } = req.body;
  if (!code || !name) return res.status(400).json({ error: 'Code and name required' });
  try {
    const result = db.prepare('INSERT INTO entities (code, name) VALUES (?, ?)').run(code.toUpperCase(), name);
    const entityId = result.lastInsertRowid;
    // Seed default COA
    const insert = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
    const seedCOA = db.transaction(() => {
      for (const a of DEFAULT_COA) insert.run(entityId, a.code, a.name, a.type, a.subtype, a.bank);
    });
    seedCOA();
    res.json({ id: entityId });
  } catch (e) {
    if (e.message.includes('UNIQUE')) return res.status(400).json({ error: 'Entity code already exists' });
    throw e;
  }
});

app.post('/api/entities/bulk', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const { entities } = req.body;
  if (!Array.isArray(entities) || entities.length === 0) return res.status(400).json({ error: 'No entities provided' });
  const insertEntity = db.prepare('INSERT OR IGNORE INTO entities (code, name) VALUES (?, ?)');
  const insertAcct = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
  const created = [];
  const bulkInsert = db.transaction(() => {
    for (const e of entities) {
      if (!e.code || !e.name) continue;
      const r = insertEntity.run(e.code.toUpperCase(), e.name);
      if (r.changes > 0) {
        const eid = r.lastInsertRowid;
        for (const a of DEFAULT_COA) insertAcct.run(eid, a.code, a.name, a.type, a.subtype, a.bank);
        created.push({ id: eid, code: e.code.toUpperCase(), name: e.name });
      }
    }
  });
  bulkInsert();
  res.json({ created, count: created.length });
});

app.delete('/api/entities/:id', auth, requireRole('Admin'), (req, res) => {
  db.prepare('DELETE FROM entities WHERE id = ?').run(req.params.id);
  res.json({ success: true });
});

// ─── Chart of Accounts Routes ───
app.get('/api/entities/:eid/accounts', auth, (req, res) => {
  const accounts = db.prepare('SELECT * FROM accounts WHERE entity_id = ? ORDER BY code').all(req.params.eid);
  res.json(accounts);
});

app.post('/api/entities/:eid/accounts', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const { code, name, type, subtype, bank_acct } = req.body;
  if (!code || !name || !type) return res.status(400).json({ error: 'Code, name, and type required' });
  try {
    const r = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)').run(req.params.eid, code, name, type, subtype || '', bank_acct ? 1 : 0);
    res.json({ id: r.lastInsertRowid });
  } catch (e) {
    if (e.message.includes('UNIQUE')) return res.status(400).json({ error: 'Account code already exists' });
    throw e;
  }
});

app.delete('/api/entities/:eid/accounts/:code', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const hasEntries = db.prepare('SELECT COUNT(*) as c FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id WHERE je.entity_id = ? AND jl.account_code = ?').get(req.params.eid, req.params.code);
  if (hasEntries.c > 0) return res.status(400).json({ error: 'Cannot delete account with transactions' });
  db.prepare('DELETE FROM accounts WHERE entity_id = ? AND code = ?').run(req.params.eid, req.params.code);
  res.json({ success: true });
});

// ─── Journal Entry Routes ───
app.get('/api/entities/:eid/entries', auth, (req, res) => {
  const entries = db.prepare('SELECT * FROM journal_entries WHERE entity_id = ? ORDER BY date DESC, id DESC').all(req.params.eid);
  const lineStmt = db.prepare('SELECT * FROM journal_lines WHERE entry_id = ?');
  const result = entries.map(e => ({ ...e, lines: lineStmt.all(e.id) }));
  res.json(result);
});

app.post('/api/entities/:eid/entries', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const { date, memo, lines } = req.body;
  if (!date || !memo || !lines || lines.length < 2) return res.status(400).json({ error: 'Invalid entry' });
  const totalDr = lines.reduce((s, l) => s + (l.debit || 0), 0);
  const totalCr = lines.reduce((s, l) => s + (l.credit || 0), 0);
  if (Math.abs(totalDr - totalCr) > 0.005) return res.status(400).json({ error: 'Entry must balance' });

  // Get next entry number for this entity
  const last = db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(req.params.eid);
  const entryNum = (last.m || 0) + 1;

  const insertEntry = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?, ?, ?, ?, ?)');
  const insertLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?, ?, ?, ?)');

  const result = db.transaction(() => {
    const r = insertEntry.run(req.params.eid, entryNum, date, memo, req.user.name);
    for (const l of lines) insertLine.run(r.lastInsertRowid, l.account_code, l.debit || 0, l.credit || 0);
    return r.lastInsertRowid;
  })();

  res.json({ id: result, entry_num: entryNum });
});

app.delete('/api/entities/:eid/entries/:id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  db.prepare('DELETE FROM journal_entries WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// ─── Bank Reconciliation Routes ───
app.get('/api/entities/:eid/reconciliations', auth, (req, res) => {
  const recs = db.prepare('SELECT * FROM reconciliations WHERE entity_id = ? ORDER BY completed_at DESC').all(req.params.eid);
  res.json(recs);
});

app.get('/api/entities/:eid/cleared/:accountCode', auth, (req, res) => {
  const cleared = db.prepare('SELECT entry_id, line_index FROM cleared_items WHERE entity_id = ? AND account_code = ?').all(req.params.eid, req.params.accountCode);
  const map = {};
  cleared.forEach(c => { map[c.entry_id + '-' + c.line_index] = true; });
  res.json(map);
});

app.post('/api/entities/:eid/reconciliations', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const { account_code, statement_date, statement_balance, book_balance, cleared_keys } = req.body;
  if (!account_code || !statement_date || statement_balance == null) return res.status(400).json({ error: 'Missing fields' });

  const insertRec = db.prepare('INSERT INTO reconciliations (entity_id, account_code, statement_date, statement_balance, book_balance, cleared_count, completed_by) VALUES (?, ?, ?, ?, ?, ?, ?)');
  const insertCleared = db.prepare('INSERT OR IGNORE INTO cleared_items (entity_id, account_code, entry_id, line_index, reconciliation_id) VALUES (?, ?, ?, ?, ?)');

  const result = db.transaction(() => {
    const r = insertRec.run(req.params.eid, account_code, statement_date, statement_balance, book_balance, cleared_keys?.length || 0, req.user.name);
    const recId = r.lastInsertRowid;
    if (cleared_keys) {
      for (const key of cleared_keys) {
        const [entryId, lineIdx] = key.split('-').map(Number);
        insertCleared.run(req.params.eid, account_code, entryId, lineIdx, recId);
      }
    }
    return recId;
  })();

  res.json({ id: result });
});

// ─── Reports (server-computed for large datasets) ───
app.get('/api/entities/:eid/balances', auth, (req, res) => {
  const rows = db.prepare(`
    SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct,
      SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit
    FROM journal_lines jl
    JOIN journal_entries je ON jl.entry_id = je.id
    JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    WHERE je.entity_id = ?
    GROUP BY jl.account_code
  `).all(req.params.eid);

  const balances = rows.map(r => {
    const isDebitNormal = r.type === 'Asset' || r.type === 'Expense';
    const balance = isDebitNormal ? (r.total_debit - r.total_credit) : (r.total_credit - r.total_debit);
    return { code: r.account_code, name: r.name, type: r.type, subtype: r.subtype, bank_acct: r.bank_acct, balance, total_debit: r.total_debit, total_credit: r.total_credit };
  });
  res.json(balances);
});

// ─── Summary across all entities ───
app.get('/api/summary', auth, (req, res) => {
  const entities = db.prepare('SELECT * FROM entities ORDER BY code').all();
  const summaries = entities.map(e => {
    const rows = db.prepare(`
      SELECT a.type, SUM(jl.debit) as td, SUM(jl.credit) as tc
      FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id
      JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
      WHERE je.entity_id = ? GROUP BY a.type
    `).all(e.id);

    const byType = {};
    rows.forEach(r => {
      const isDebit = r.type === 'Asset' || r.type === 'Expense';
      byType[r.type] = isDebit ? (r.td - r.tc) : (r.tc - r.td);
    });

    const entryCount = db.prepare('SELECT COUNT(*) as c FROM journal_entries WHERE entity_id = ?').get(e.id);

    return {
      ...e, assets: byType.Asset || 0, liabilities: byType.Liability || 0,
      revenue: byType.Revenue || 0, expenses: byType.Expense || 0,
      net_income: (byType.Revenue || 0) - (byType.Expense || 0),
      entry_count: entryCount.c
    };
  });
  res.json(summaries);
});

// ─── Catch-all: serve frontend ───
if (process.env.NODE_ENV === 'production') {
  app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, '..', 'client', 'dist', 'index.html'));
  });
}

// ─── Start ───
app.listen(PORT, '0.0.0.0', () => {
  console.log(`CloudLedger running on port ${PORT}`);
  console.log(`Database: ${DB_PATH}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
});
