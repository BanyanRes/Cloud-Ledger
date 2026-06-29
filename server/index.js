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
const turnkey = require('./turnkey');
const requisition = require('./requisition');
const { rollForward } = require('./requisition_rollforward');
const { verifyRollforward } = require('./requisition_rollforward_verify');
const { saveRequisitionOutputs } = require('./requisition_workpaper_save');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-secret';
const DB_PATH = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'cloudledger.db');
const RESEND_API_KEY = process.env.RESEND_API_KEY || '';
const RESET_FROM_EMAIL = process.env.RESET_FROM_EMAIL || 'CloudLedger <onboarding@resend.dev>';
const APP_URL = process.env.APP_URL || '';
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
const memUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 64 * 1024 * 1024 } });
// Roll-forward sends the period's invoices (including each PDF's base64 bytes) in
// a large `invoices` text field. multer's default fieldSize is only 1MB, which
// silently fails the request for a normal multi-invoice period. Allow a big text
// field (and a comfortable file size for the workbook) on that route only.
const reqRollUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024, fieldSize: 80 * 1024 * 1024, fields: 50 } });

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
  CREATE TABLE IF NOT EXISTS password_reset_tokens (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    token_hash TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    used_at TEXT,
    created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE INDEX IF NOT EXISTS idx_prt_token_hash ON password_reset_tokens(token_hash);
  CREATE TABLE IF NOT EXISTS user_entity_access (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    created_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (user_id, entity_id)
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
  CREATE INDEX IF NOT EXISTS idx_accounts_entity_code ON accounts(entity_id, code);
  CREATE INDEX IF NOT EXISTS idx_je_entity ON journal_entries(entity_id);
  CREATE INDEX IF NOT EXISTS idx_je_date ON journal_entries(entity_id, date);
  CREATE INDEX IF NOT EXISTS idx_jl_entry ON journal_lines(entry_id);
  CREATE INDEX IF NOT EXISTS idx_jl_account_code ON journal_lines(account_code);
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
  -- Bill.com dimension maps: accountingClassId -> CL class (investor),
  -- jobId -> CL location (deal). Only mapped ids carry a dimension onto synced
  -- JE lines; an unmapped id (e.g. a workflow-status class) syncs as NULL, so
  -- the map itself is the filter — no hardcoded skip list needed.
  CREATE TABLE IF NOT EXISTS billcom_class_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    billcom_class_id TEXT NOT NULL,
    billcom_class_name TEXT,
    cl_class_id INTEGER NOT NULL,
    created_at TEXT,
    UNIQUE(entity_id, billcom_class_id),
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_bccm_entity ON billcom_class_map(entity_id);
  CREATE TABLE IF NOT EXISTS billcom_location_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    billcom_job_id TEXT NOT NULL,
    billcom_job_name TEXT,
    cl_location_id INTEGER NOT NULL,
    created_at TEXT,
    UNIQUE(entity_id, billcom_job_id),
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_bclm_entity ON billcom_location_map(entity_id);
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
  -- ==========================================================
  -- API keys for system-to-system integrations (e.g., Turnkey Rail)
  -- ==========================================================
  CREATE TABLE IF NOT EXISTS api_keys (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key_hash TEXT NOT NULL UNIQUE,
    key_prefix TEXT NOT NULL,
    name TEXT NOT NULL,
    scopes TEXT NOT NULL DEFAULT '',
    last_used_at TEXT,
    created_by TEXT,
    created_at TEXT NOT NULL,
    revoked_at TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_api_keys_hash ON api_keys(key_hash);
  -- ==========================================================
  -- Turnkey Rail integration (mirrors billcom pattern)
  -- ==========================================================
  CREATE TABLE IF NOT EXISTS turnkey_config (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    enabled INTEGER NOT NULL DEFAULT 0,
    webhook_secret_enc TEXT,
    updated_by TEXT,
    updated_at TEXT
  );
  -- Maps Turnkey projects -> CloudLedger entities, with all POC account codes
  CREATE TABLE IF NOT EXISTS turnkey_project_map (
    turnkey_project_id INTEGER PRIMARY KEY,
    cl_entity_id INTEGER NOT NULL,
    cash_account_code TEXT,
    billcom_clearing_code TEXT,
    ar_owner_code TEXT,
    costs_in_excess_code TEXT,
    cip_code TEXT,
    ap_sub_code TEXT,
    billings_uncompleted_code TEXT,
    billings_in_excess_code TEXT,
    revenue_code TEXT,
    cost_of_construction_code TEXT,
    created_at TEXT,
    FOREIGN KEY (cl_entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_tpm_entity ON turnkey_project_map(cl_entity_id);
  -- Maps Turnkey subcontractor IDs to per-entity vendor sub-accounts (optional)
  CREATE TABLE IF NOT EXISTS turnkey_vendor_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cl_entity_id INTEGER NOT NULL,
    turnkey_vendor_id INTEGER NOT NULL,
    vendor_name TEXT,
    ap_sub_account_code TEXT,
    created_at TEXT,
    UNIQUE(cl_entity_id, turnkey_vendor_id),
    FOREIGN KEY (cl_entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_tvm_entity ON turnkey_vendor_map(cl_entity_id);
  -- Sync event audit log + idempotency
  CREATE TABLE IF NOT EXISTS turnkey_sync_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cl_entity_id INTEGER NOT NULL,
    sync_type TEXT NOT NULL,
    turnkey_id TEXT,
    cl_entry_id INTEGER,
    status TEXT NOT NULL,
    message TEXT,
    payload_json TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY (cl_entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_tsl_entity ON turnkey_sync_log(cl_entity_id, created_at);
  CREATE INDEX IF NOT EXISTS idx_tsl_turnkey_id ON turnkey_sync_log(sync_type, turnkey_id);
  -- ==========================================================
  -- Requisition Report / Invoice Packet (development entities only)
  -- ==========================================================
  -- One row per invoice read by the roll-forward flow. The PDF bytes are stored
  -- inline (file_blob) so the invoice packet can be regenerated later without
  -- depending on Bill.com. req_number is filled in when a roll-forward that
  -- includes this invoice succeeds; until then it is NULL (read-but-not-yet-rolled).
  CREATE TABLE IF NOT EXISTS requisition_invoice (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id),
    req_number INTEGER,
    vendor TEXT,
    bill_number TEXT,
    amount REAL,
    invoice_date TEXT,
    cost_code TEXT,
    cost_code_name TEXT,
    confidence TEXT,
    original_name TEXT,
    mime_type TEXT,
    file_blob BLOB,
    created_at TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_reqinv_entity ON requisition_invoice(entity_id);
  CREATE INDEX IF NOT EXISTS idx_reqinv_req ON requisition_invoice(entity_id, req_number);
  CREATE TABLE IF NOT EXISTS requisition_coding_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id),
    vendor_norm TEXT NOT NULL,
    bill_signature TEXT,
    cost_category TEXT,
    cost_code TEXT,
    bank_cost_category TEXT,
    gl_coding TEXT,
    cost_code_name TEXT,
    req_number INTEGER,
    weight INTEGER NOT NULL DEFAULT 1,
    created_at TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_reqch_lookup ON requisition_coding_history(entity_id, vendor_norm);
  -- Cost-code catalog per development entity. This is the master list of cost
  -- codes that drives the Budget-to-Actual report and the canonical
  -- (cost_category / cost_code_name / bank_cost_category) spelling used when a
  -- coded line is written. Seeded from prior workbooks / Invoice Logs.
  CREATE TABLE IF NOT EXISTS requisition_coa_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id),
    cost_code TEXT NOT NULL,
    cost_code_name TEXT,
    cost_category TEXT,
    bank_cost_category TEXT,
    gl_coding TEXT,
    budget_amount REAL,
    sort_order INTEGER,
    created_at TEXT,
    UNIQUE(entity_id, cost_code)
  );
  CREATE INDEX IF NOT EXISTS idx_reqcoa_entity ON requisition_coa_map(entity_id);
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

// Entity type: 'accounting' (default, standard ledger entity) | 'development' | 'shell' (tracks location + investor/class dimensions)
// (real-estate development project; unlocks Requisition Report / Invoice Packet features)
const entCols = db.prepare("PRAGMA table_info(entities)").all().map(c => c.name);
if (!entCols.includes('entity_type')) {
  db.exec("ALTER TABLE entities ADD COLUMN entity_type TEXT NOT NULL DEFAULT 'accounting'");
  console.log('[db migrate] entities.entity_type added (default accounting)');
}
// display_id: short user-facing identifier (e.g. "0005 B1a") used as a filename
// prefix for requisition invoice packets. Optional; falls back to entity name.
if (!entCols.includes('display_id')) {
  db.exec("ALTER TABLE entities ADD COLUMN display_id TEXT");
  console.log('[db migrate] entities.display_id added');
}

// Phase 3: default_cash_account on billcom_config (for payment JEs)
const bcCfgCols = db.prepare("PRAGMA table_info(billcom_config)").all().map(c => c.name);
if (!bcCfgCols.includes('default_cash_account')) db.exec("ALTER TABLE billcom_config ADD COLUMN default_cash_account TEXT");
// Phase 7 (payment reconcile): clearing account for the QBO-style two-leg model.
// Leg 1 relieves AP into the clearing account; Leg 2 moves the clearing balance
// to operating cash as a process-date lump sum (mirrors Bill.com Money Out Clearing).
if (!bcCfgCols.includes('default_clearing_account')) db.exec("ALTER TABLE billcom_config ADD COLUMN default_clearing_account TEXT");
if (!bcCfgCols.includes('sync_cutoff_date')) db.exec("ALTER TABLE billcom_config ADD COLUMN sync_cutoff_date TEXT");

// Bank-transaction matching: link a bank line to an already-posted JE instead of
// creating a new one. matched_entry_id holds the JE id; status becomes 'matched'.
const btCols = db.prepare("PRAGMA table_info(bank_transactions)").all().map(c => c.name);
if (!btCols.includes('matched_entry_id')) db.exec("ALTER TABLE bank_transactions ADD COLUMN matched_entry_id INTEGER");

// === Turnkey Rail integration v2 migrations ===
// Job costing: journal_lines.project_id tags each line to a Turnkey project,
// so a SINGLE company entity can hold ALL projects with proper job-level cost
// dimension. Reconciles to WIP schedule reports.
const jlCols = db.prepare("PRAGMA table_info(journal_lines)").all().map(c => c.name);
if (!jlCols.includes('project_id')) {
  db.exec("ALTER TABLE journal_lines ADD COLUMN project_id TEXT");
  db.exec("CREATE INDEX IF NOT EXISTS idx_jl_project ON journal_lines(project_id)");
  console.log('[db migrate] journal_lines.project_id added');
}
// GL detail import: per-line narrative explaining why the transaction was booked.
// Existing routes default it to '' so this is a safe additive migration.
if (!jlCols.includes('description')) {
  db.exec("ALTER TABLE journal_lines ADD COLUMN description TEXT DEFAULT ''");
  console.log('[db migrate] journal_lines.description added');
}
// Analytical dimensions on journal lines: class (e.g. investor tracking) and
// location (e.g. deal/asset on which pre-deal costs are capitalized). Both are
// normalized master tables scoped per entity, referenced by nullable FKs on
// journal_lines. Dimensions are LINE attributes, never JE grouping keys, so a
// single balanced JE may carry many different investors/deals across its lines.
db.exec(`
  CREATE TABLE IF NOT EXISTS dim_classes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    kind TEXT DEFAULT 'investor',
    UNIQUE(entity_id, name)
  );
  CREATE TABLE IF NOT EXISTS dim_locations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    kind TEXT DEFAULT '',
    UNIQUE(entity_id, name)
  );
  CREATE TABLE IF NOT EXISTS dim_projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    code TEXT,
    kind TEXT DEFAULT 'project',
    UNIQUE(entity_id, name)
  );
  -- Investor capital commitments (informational only; never posts to the GL).
  -- Links to dim_classes (kind='investor'). Tracks committed + called-to-date;
  -- uncalled and ownership % are computed on read.
  CREATE TABLE IF NOT EXISTS investor_commitments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    class_id INTEGER NOT NULL REFERENCES dim_classes(id) ON DELETE CASCADE,
    commitment_amount REAL NOT NULL DEFAULT 0,
    called_amount REAL NOT NULL DEFAULT 0,
    commit_date TEXT,
    notes TEXT,
    created_at TEXT,
    updated_at TEXT,
    UNIQUE(entity_id, class_id)
  );
  CREATE INDEX IF NOT EXISTS idx_invcommit_entity ON investor_commitments(entity_id);
  -- Saved/memorized report configurations (QBO-style). Shared per entity:
  -- every user with access to the entity sees all of its saved reports.
  -- config_json holds the report-specific settings (accounts, group-by, dates, etc.).
  CREATE TABLE IF NOT EXISTS memorized_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    report_type TEXT NOT NULL,
    name TEXT NOT NULL,
    config_json TEXT NOT NULL DEFAULT '{}',
    created_by INTEGER,
    created_by_name TEXT,
    created_at TEXT,
    updated_at TEXT,
    UNIQUE(entity_id, report_type, name)
  );
  CREATE INDEX IF NOT EXISTS idx_memrep_entity ON memorized_reports(entity_id);
`);
if (!jlCols.includes('class_id')) {
  db.exec("ALTER TABLE journal_lines ADD COLUMN class_id INTEGER");
  db.exec("CREATE INDEX IF NOT EXISTS idx_jl_class ON journal_lines(class_id)");
  console.log('[db migrate] journal_lines.class_id added');
}
if (!jlCols.includes('location_id')) {
  db.exec("ALTER TABLE journal_lines ADD COLUMN location_id INTEGER");
  db.exec("CREATE INDEX IF NOT EXISTS idx_jl_location ON journal_lines(location_id)");
  console.log('[db migrate] journal_lines.location_id added');
}
if (!jlCols.includes('project_id')) {
  db.exec("ALTER TABLE journal_lines ADD COLUMN project_id INTEGER");
  db.exec("CREATE INDEX IF NOT EXISTS idx_jl_project ON journal_lines(project_id)");
  console.log('[db migrate] journal_lines.project_id added');
}
// Dimension code columns (name was the only label originally; code added for reporting/sorting)
const dcCols = db.prepare("PRAGMA table_info(dim_classes)").all().map(c => c.name);
if (!dcCols.includes('code')) { db.exec("ALTER TABLE dim_classes ADD COLUMN code TEXT"); console.log('[db migrate] dim_classes.code added'); }
const dlCols = db.prepare("PRAGMA table_info(dim_locations)").all().map(c => c.name);
if (!dlCols.includes('code')) { db.exec("ALTER TABLE dim_locations ADD COLUMN code TEXT"); console.log('[db migrate] dim_locations.code added'); }
// turnkey_project_map redesigned: no longer stores per-project account codes
// (single COA on the company entity now). We add cl_entity_id linking to the
// COMPANY entity (same for all projects). Keep existing rows on upgrade.
const tpmCols = db.prepare("PRAGMA table_info(turnkey_project_map)").all().map(c => c.name);
if (!tpmCols.includes('project_code')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN project_code TEXT");
  console.log('[db migrate] turnkey_project_map.project_code added');
}
if (!tpmCols.includes('project_name')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN project_name TEXT");
  console.log('[db migrate] turnkey_project_map.project_name added');
}
if (!tpmCols.includes('contract_amount')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN contract_amount REAL");
  console.log('[db migrate] turnkey_project_map.contract_amount added');
}
if (!tpmCols.includes('total_estimated_costs')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN total_estimated_costs REAL");
  console.log('[db migrate] turnkey_project_map.total_estimated_costs added');
}
// turnkey_config gets a default_entity_id: the company entity that holds all
// projects. Admin sets this once before enabling integration.
const tcCols = db.prepare("PRAGMA table_info(turnkey_config)").all().map(c => c.name);
if (!tcCols.includes('default_entity_id')) {
  db.exec("ALTER TABLE turnkey_config ADD COLUMN default_entity_id INTEGER");
  console.log('[db migrate] turnkey_config.default_entity_id added');
}

// === Accounts Receivable (customer invoicing) schema ===
db.exec(`
  CREATE TABLE IF NOT EXISTS ar_customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    email TEXT,
    address TEXT,
    terms_days INTEGER DEFAULT 30,
    active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(entity_id, name)
  );
  CREATE TABLE IF NOT EXISTS ar_invoice_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    customer_id INTEGER NOT NULL REFERENCES ar_customers(id) ON DELETE CASCADE,
    memo TEXT,
    frequency TEXT NOT NULL DEFAULT 'monthly',
    day_of_month INTEGER DEFAULT 1,
    next_run TEXT,
    ar_account_code TEXT NOT NULL DEFAULT '11000',
    active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS ar_template_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER NOT NULL REFERENCES ar_invoice_templates(id) ON DELETE CASCADE,
    description TEXT NOT NULL,
    qty REAL DEFAULT 1,
    rate REAL DEFAULT 0,
    revenue_account_code TEXT NOT NULL,
    class_id INTEGER,
    location_id INTEGER,
    sort INTEGER DEFAULT 0
  );
  CREATE TABLE IF NOT EXISTS ar_invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    customer_id INTEGER REFERENCES ar_customers(id) ON DELETE SET NULL,
    template_id INTEGER REFERENCES ar_invoice_templates(id) ON DELETE SET NULL,
    invoice_num TEXT NOT NULL,
    invoice_date TEXT NOT NULL,
    due_date TEXT,
    customer_name TEXT,
    customer_email TEXT,
    customer_address TEXT,
    memo TEXT,
    subtotal REAL DEFAULT 0,
    total REAL DEFAULT 0,
    ar_account_code TEXT,
    status TEXT NOT NULL DEFAULT 'draft',
    je_id INTEGER REFERENCES journal_entries(id) ON DELETE SET NULL,
    pay_je_id INTEGER REFERENCES journal_entries(id) ON DELETE SET NULL,
    pdf_file_id INTEGER,
    sent_at TEXT,
    paid_at TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(entity_id, invoice_num)
  );
  CREATE TABLE IF NOT EXISTS ar_invoice_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL REFERENCES ar_invoices(id) ON DELETE CASCADE,
    description TEXT NOT NULL,
    qty REAL DEFAULT 1,
    rate REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    revenue_account_code TEXT NOT NULL,
    class_id INTEGER,
    location_id INTEGER,
    sort INTEGER DEFAULT 0
  );
  CREATE INDEX IF NOT EXISTS idx_ar_customers_ent ON ar_customers(entity_id);
  CREATE INDEX IF NOT EXISTS idx_ar_templates_ent ON ar_invoice_templates(entity_id);
  CREATE INDEX IF NOT EXISTS idx_ar_invoices_ent ON ar_invoices(entity_id);
  CREATE INDEX IF NOT EXISTS idx_ar_invoices_status ON ar_invoices(entity_id, status);
`);
console.log('[db] AR schema ready');


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

// Generic v3 classification list (paginated) for any resource under
// /classifications, e.g. 'accounting-classes', 'jobs', 'locations'. Returns the
// raw items. Mirrors billcomListAccounts' pagination + error handling.
async function billcomListClassification({ sessionId, devKey, baseUrl, resource }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const out = [];
  let nextPage = null;
  while (true) {
    const params = new URLSearchParams({ max: '100' });
    if (nextPage) params.set('nextPage', nextPage);
    const url = base + '/classifications/' + resource + '?' + params.toString();
    const resp = await fetch(url, { method: 'GET', headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' } });
    const text = await resp.text();
    let json; try { json = JSON.parse(text); } catch { throw new Error('Non-JSON (' + resource + ', HTTP ' + resp.status + '): ' + text.slice(0, 150)); }
    if (!resp.ok) {
      const msg = Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join('; ') : (json.message || ('HTTP ' + resp.status));
      throw new Error('Bill.com error (' + resource + '): ' + msg);
    }
    const items = Array.isArray(json.results) ? json.results : (Array.isArray(json) ? json : []);
    out.push(...items);
    nextPage = json.nextPage || null;
    if (!nextPage || out.length > 10000) break;
  }
  return out;
}

// Hard wall-clock timeout for a single Bill.com fetch. AbortController proved
// unreliable at cutting fetches in this runtime, so race the fetch against a
// timer and surface a clear error instead of letting Railway 502 at its gateway.
function billcomFetch(url, opts, ms) {
  return Promise.race([
    fetch(url, opts),
    new Promise((_, rej) => setTimeout(() => rej(new Error('Bill.com request timed out after ' + (ms || 15000) + 'ms')), ms || 15000)),
  ]);
}

// Generic paginated GET for v3 list endpoints. Used for bills + payments.
// maxItems (optional) caps total rows fetched so a sync can stay bounded.
async function billcomListV3({ sessionId, devKey, baseUrl, resourcePath, extraParams, maxItems }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const out = [];
  let nextPage = null;
  const max = 100;
  let pageCount = 0;
  while (true) {
    const params = new URLSearchParams({ max: String(max), ...(extraParams || {}) });
    if (nextPage) params.set('nextPage', nextPage);
    const url = base + resourcePath + '?' + params.toString();
    const resp = await billcomFetch(url, {
      method: 'GET',
      headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' }
    }, 15000);
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
    if (maxItems && out.length >= maxItems) break;
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
  const resp = await billcomFetch(url, {
    method: 'GET',
    headers: { 'sessionId': sessionId, 'devKey': devKey, 'Accept': 'application/json' }
  }, 12000);
  const text = await resp.text();
  let json; try { json = JSON.parse(text); } catch { throw new Error('Non-JSON detail (HTTP ' + resp.status + '): ' + text.slice(0, 200)); }
  if (!resp.ok) {
    const msg = (Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join('; ') : (json.message || ('HTTP ' + resp.status)));
    throw new Error('detail HTTP ' + resp.status + ': ' + msg);
  }
  return json;
}

// Paginated vendor list -> used to resolve vendorId on bills to a display name
// for the AP Aging report. Mirrors billcomListV3 pagination/error handling.
async function billcomListVendors(args) {
  return billcomListV3({ ...args, resourcePath: '/vendors' });
}

// Bill.com v3 /bills ignores offset pagination (nextPage + start both return
// the same first 100 rows). The only working way to retrieve the full set is
// to filter by dueDate windows (filters=dueDate:gte:X,dueDate:lt:Y — comma = AND).
// We walk month-sized windows across [fromYM, toYM], union + dedupe by id.
// Each window for CLRF returns well under 100 rows, so nothing truncates.
async function billcomListBillsWindowed({ sessionId, devKey, baseUrl, fromDate, toDate }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const hdr = { sessionId, devKey, Accept: "application/json" };
  const addMonth = (d) => { const [Y, M] = d.split("-"); let yy = +Y, mm = +M + 1; if (mm > 12) { mm = 1; yy++; } return yy + "-" + String(mm).padStart(2, "0") + "-01"; };
  // normalize to first-of-month window starts
  const startYM = fromDate.slice(0, 7) + "-01";
  const endExclusive = addMonth(toDate.slice(0, 7) + "-01"); // include the toDate month fully
  const byId = new Map();
  let win = startYM;
  let guard = 0;
  while (win < endExclusive && guard < 240) {
    guard++;
    const winEnd = addMonth(win);
    const filt = "dueDate:gte:" + win + ",dueDate:lt:" + winEnd;
    const url = base + "/bills?max=100&filters=" + encodeURIComponent(filt);
    let json;
    try {
      const resp = await billcomFetch(url, { method: "GET", headers: hdr }, 20000);
      const text = await resp.text();
      try { json = JSON.parse(text); } catch { throw new Error("Non-JSON bills window (HTTP " + resp.status + ")"); }
      if (!resp.ok) { const msg = Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join("; ") : (json.message || ("HTTP " + resp.status)); throw new Error("bills window: " + msg); }
    } catch (e) { throw new Error("bills window " + win + ": " + e.message); }
    const results = Array.isArray(json.results) ? json.results : [];
    for (const b of results) { const id = b && b.id; if (id != null && !byId.has(String(id))) byId.set(String(id), b); }
    // Safety: if a single month ever returns the 100 cap, narrow it would be needed;
    // log so we know to split finer. (CLRF volume is far below this.)
    if (results.length >= 100) console.log("[ap-aging] WARNING window " + win + " hit 100-row cap; may be truncated");
    win = winEnd;
  }
  return Array.from(byId.values());
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
// Entity access control: Admin bypasses; otherwise user must have either no allowlist (full access) or a matching row.
function userHasEntityAccess(userId, userRole, entityId) {
  if (userRole === 'Admin') return true;
  const allow = db.prepare('SELECT entity_id FROM user_entity_access WHERE user_id = ?').all(userId);
  if (allow.length === 0) return true; // empty allowlist = all entities
  return allow.some(r => r.entity_id === parseInt(entityId));
}
function listAccessibleEntityIds(userId, userRole) {
  if (userRole === 'Admin') return null; // null = all
  const allow = db.prepare('SELECT entity_id FROM user_entity_access WHERE user_id = ?').all(userId);
  if (allow.length === 0) return null; // empty = all
  return allow.map(r => r.entity_id);
}
function requireEntityAccess(paramName) {
  return (req, res, next) => {
    const eid = parseInt(req.params[paramName || 'eid']);
    if (!eid) return res.status(400).json({ error: 'Invalid entity id' });
    if (!userHasEntityAccess(req.user.id, req.user.role, eid)) return res.status(403).json({ error: 'No access to this entity' });
    next();
  };
}

function requireRole(...roles) { return (req, res, next) => { if (!roles.includes(req.user.role) && req.user.role !== 'Admin') return res.status(403).json({ error: 'Forbidden' }); next(); }; }

// Gate Requisition/Invoice-Packet features to development-project entities only.
// Reads the entity id from the named route param (default 'entity_id'); rejects
// non-development entities so accounting entities never expose these endpoints.
function requireDevelopmentEntity(paramName) {
  return (req, res, next) => {
    const eid = parseInt(req.params[paramName || 'entity_id']);
    if (!eid) return res.status(400).json({ error: 'Invalid entity id' });
    const ent = db.prepare('SELECT entity_type FROM entities WHERE id = ?').get(eid);
    if (!ent) return res.status(404).json({ error: 'Entity not found' });
    if (ent.entity_type !== 'development') return res.status(403).json({ error: 'Requisition features are only available for development-project entities' });
    next();
  };
}

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
// Send a password-reset email via Resend's HTTP API (no extra dependency).
async function sendResetEmail(toEmail, resetUrl) {
  if (!RESEND_API_KEY) {
    console.warn('[reset] RESEND_API_KEY not set — cannot send email. Reset URL was: ' + resetUrl);
    return { ok: false, skipped: true };
  }
  try {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + RESEND_API_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        from: RESET_FROM_EMAIL,
        to: [toEmail],
        subject: 'Reset your CloudLedger password',
        html: '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px">'
          + '<h2 style="color:#1d4ed8;margin:0 0 16px">Reset your CloudLedger password</h2>'
          + '<p style="color:#334155;font-size:14px;line-height:1.5">We received a request to reset your password. Click the button below to choose a new one. This link expires in 1 hour.</p>'
          + '<p style="margin:24px 0"><a href="' + resetUrl + '" style="background:#2563eb;color:#fff;text-decoration:none;padding:11px 22px;border-radius:6px;font-size:14px;font-weight:600;display:inline-block">Reset password</a></p>'
          + '<p style="color:#64748b;font-size:12px;line-height:1.5">If the button does not work, copy and paste this link:<br>' + resetUrl + '</p>'
          + '<p style="color:#94a3b8;font-size:12px;margin-top:24px">Did not request this? You can safely ignore this email; your password will not change.</p>'
          + '</div>',
      }),
    });
    if (!r.ok) { const t = await r.text(); console.error('[reset] Resend error ' + r.status + ': ' + t); return { ok: false, status: r.status, body: t }; }
    return { ok: true };
  } catch (e) {
    console.error('[reset] send failed: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// Request a password reset. Always returns a neutral response (never reveals
// whether an account exists). Generates a single-use token (1h expiry) and emails a link.
app.post('/api/auth/forgot-password', async (req, res) => {
  const email = (req.body.email || '').toLowerCase().trim();
  const neutral = { ok: true, message: 'If an account exists for that email, a reset link has been sent.' };
  if (!email) return res.json(neutral);
  const user = db.prepare('SELECT id, email FROM users WHERE email = ?').get(email);
  if (!user) return res.json(neutral); // do not disclose non-existence
  const rawToken = require('crypto').randomBytes(32).toString('hex');
  const tokenHash = require('crypto').createHash('sha256').update(rawToken).digest('hex');
  const expiresAt = new Date(Date.now() + 60 * 60 * 1000).toISOString(); // 1 hour
  // Invalidate any prior unused tokens for this user, then store the new one.
  db.prepare('UPDATE password_reset_tokens SET used_at = ? WHERE user_id = ? AND used_at IS NULL').run(new Date().toISOString(), user.id);
  db.prepare('INSERT INTO password_reset_tokens (user_id, token_hash, expires_at) VALUES (?, ?, ?)').run(user.id, tokenHash, expiresAt);
  const base = APP_URL || (req.protocol + '://' + req.get('host'));
  const resetUrl = base.replace(/\/$/, '') + '/?reset_token=' + rawToken;
  await sendResetEmail(user.email, resetUrl);
  res.json(neutral);
});

// Complete a password reset using the emailed token. Validates token, sets new password, burns token.
app.post('/api/auth/reset-password', (req, res) => {
  const { token, new_password } = req.body || {};
  if (!token || !new_password) return res.status(400).json({ error: 'Token and new password required' });
  if (String(new_password).length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
  const tokenHash = require('crypto').createHash('sha256').update(String(token)).digest('hex');
  const row = db.prepare('SELECT * FROM password_reset_tokens WHERE token_hash = ?').get(tokenHash);
  if (!row) return res.status(400).json({ error: 'Invalid or expired reset link' });
  if (row.used_at) return res.status(400).json({ error: 'This reset link has already been used' });
  if (new Date(row.expires_at).getTime() < Date.now()) return res.status(400).json({ error: 'This reset link has expired' });
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(bcrypt.hashSync(String(new_password), 10), row.user_id);
  db.prepare('UPDATE password_reset_tokens SET used_at = ? WHERE id = ?').run(new Date().toISOString(), row.id);
  res.json({ ok: true, message: 'Password updated. You can now sign in.' });
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

// Entity access management (Admin only)
app.get('/api/users/:id/entity-access', auth, requireRole('Admin'), (req, res) => {
  const rows = db.prepare('SELECT entity_id FROM user_entity_access WHERE user_id = ? ORDER BY entity_id').all(req.params.id);
  res.json({ user_id: parseInt(req.params.id), entity_ids: rows.map(r => r.entity_id) });
});
app.put('/api/users/:id/entity-access', auth, requireRole('Admin'), (req, res) => {
  const userId = parseInt(req.params.id);
  const targetUser = db.prepare('SELECT id, role FROM users WHERE id = ?').get(userId);
  if (!targetUser) return res.status(404).json({ error: 'User not found' });
  if (targetUser.role === 'Admin') return res.status(400).json({ error: 'Admins always have all-entity access; cannot restrict' });
  const ids = Array.isArray(req.body.entity_ids) ? req.body.entity_ids.map(n => parseInt(n)).filter(n => Number.isInteger(n)) : [];
  const tx = db.transaction(() => {
    db.prepare('DELETE FROM user_entity_access WHERE user_id = ?').run(userId);
    const ins = db.prepare('INSERT INTO user_entity_access (user_id, entity_id) VALUES (?, ?)');
    for (const eid of ids) ins.run(userId, eid);
  });
  tx();
  res.json({ user_id: userId, entity_ids: ids });
});



// ═══ Entities ═══
app.get('/api/entities', auth, (req, res) => {
  const ids = listAccessibleEntityIds(req.user.id, req.user.role);
  if (ids === null) return res.json(db.prepare('SELECT * FROM entities ORDER BY code').all());
  if (ids.length === 0) return res.json([]);
  const placeholders = ids.map(() => '?').join(',');
  res.json(db.prepare('SELECT * FROM entities WHERE id IN (' + placeholders + ') ORDER BY code').all(...ids));
});
app.post('/api/entities', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { name } = req.body; if (!name) return res.status(400).json({ error: 'Name required' });
  const entityType = ['development','shell'].includes(req.body.entity_type) ? req.body.entity_type : 'accounting';
  const displayId = (req.body.display_id || '').trim() || null;
  // Auto-generate a code from the name (used internally for sorting/uniqueness)
  const baseCode = name.replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 8) || 'ENT';
  let code = baseCode; let n = 1;
  while (db.prepare('SELECT id FROM entities WHERE code = ?').get(code)) { code = baseCode + n; n++; }
  try { const r = db.prepare('INSERT INTO entities (code, name, entity_type, display_id) VALUES (?, ?, ?, ?)').run(code, name, entityType, displayId); const eid = r.lastInsertRowid;
    const ins = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
    db.transaction(() => { for (const a of DEFAULT_COA) ins.run(eid, a.code, a.name, a.type, a.subtype, a.bank); })();
    res.json({ id: eid, code, name, entity_type: entityType, display_id: displayId }); } catch(e) { throw e; }
});
// Update an entity (currently: name and/or entity_type)
app.put('/api/entities/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const id = parseInt(req.params.id);
  if (!id) return res.status(400).json({ error: 'Invalid entity id' });
  const ent = db.prepare('SELECT * FROM entities WHERE id = ?').get(id);
  if (!ent) return res.status(404).json({ error: 'Entity not found' });
  const name = req.body.name !== undefined ? req.body.name : ent.name;
  if (!name) return res.status(400).json({ error: 'Name required' });
  let entityType = ent.entity_type;
  if (req.body.entity_type !== undefined) {
    if (!['development','accounting','shell'].includes(req.body.entity_type)) return res.status(400).json({ error: 'entity_type must be development, accounting, or shell' });
    entityType = req.body.entity_type;
  }
  let displayId = ent.display_id;
  if (req.body.display_id !== undefined) displayId = (req.body.display_id || '').trim() || null;
  db.prepare('UPDATE entities SET name = ?, entity_type = ?, display_id = ? WHERE id = ?').run(name, entityType, displayId, id);
  res.json({ id, code: ent.code, name, entity_type: entityType, display_id: displayId });
});
app.post('/api/entities/bulk', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { entities } = req.body; if (!Array.isArray(entities)) return res.status(400).json({ error: 'Invalid' });
  const insE = db.prepare('INSERT OR IGNORE INTO entities (code, name) VALUES (?, ?)');
  const insA = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
  const created = []; db.transaction(() => { for (const e of entities) { if (!e.code||!e.name) continue; const r = insE.run(e.code.toUpperCase(), e.name);
    if (r.changes > 0) { const eid = r.lastInsertRowid; for (const a of DEFAULT_COA) insA.run(eid, a.code, a.name, a.type, a.subtype, a.bank); created.push({ id: eid, code: e.code.toUpperCase(), name: e.name }); } } })();
  res.json({ created, count: created.length });
});
app.delete('/api/entities/:id', auth, requireRole('Admin'), (req, res) => {
  const id = parseInt(req.params.id);
  if (!id) return res.status(400).json({ error: 'Invalid entity id' });
  // Guard: never delete the entity that Turnkey Rail is configured to use.
  const tkCfg = db.prepare('SELECT default_entity_id FROM turnkey_config WHERE id = 1').get();
  if (tkCfg && tkCfg.default_entity_id === id) {
    return res.status(400).json({ error: 'This entity is configured as the Turnkey Rail default entity. Reassign Turnkey config before deleting.' });
  }
  // Tables that reference entities WITHOUT ON DELETE CASCADE must be cleared first
  // (billcom_* keyed by entity_id; turnkey_* keyed by cl_entity_id). Tables that
  // already cascade (accounts, journal_entries, journal_lines, bank_transactions,
  // entity_files, etc.) are removed automatically by the final entities delete.
  try {
    db.transaction(() => {
      db.prepare('DELETE FROM billcom_account_map WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM billcom_sync_log WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM billcom_config WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM turnkey_vendor_map WHERE cl_entity_id = ?').run(id);
      db.prepare('DELETE FROM turnkey_sync_log WHERE cl_entity_id = ?').run(id);
      db.prepare('DELETE FROM turnkey_project_map WHERE cl_entity_id = ?').run(id);
      // Requisition tables (none cascade): invoices stored inline, plus coding history.
      db.prepare('DELETE FROM requisition_invoice WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM requisition_coding_history WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM requisition_coa_map WHERE entity_id = ?').run(id);
      db.prepare('DELETE FROM entities WHERE id = ?').run(id);
    })();
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: 'Delete failed: ' + e.message });
  }
});

// Import trial balance: replaces COA and posts a beginning-balance JE
// Account type derived from code: <=19999 Asset, <=29999 Liability, <=39999 Equity, <=49999 Revenue, 50000-69999 Expense, >=70000 Revenue
app.post('/api/entities/:eid/import-tb', auth, requireEntityAccess(), requireRole('Admin','Accountant'), memUpload.single('file'), (req, res) => {
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
      // (otherwise re-importing would stack on top of the previous opening balances).
      // Also remove any prior GL-detail import: importing a TB replaces GL history and
      // vice-versa (latest import wins), so the two never double-count on one entity.
      db.prepare("DELETE FROM journal_entries WHERE entity_id = ? AND memo = 'Opening balance from imported trial balance'").run(eid);
      db.prepare("DELETE FROM journal_entries WHERE entity_id = ? AND memo LIKE 'GL detail import%'").run(eid);

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

// ═══ General Ledger Detail import ═══
// GL detail reports are one row per transaction and every accounting package lays
// them out differently, so import is a two-step "preview → map → import" flow:
//   1) POST /import-gl/preview  parses the raw grid, auto-detects columns + whether
//      account number & name are fused in one cell, and returns a preview.
//   2) POST /import-gl  takes the user-confirmed column mapping and posts the data.
// Importing GL detail replaces any prior TB import AND prior GL import on the entity
// (latest import wins), so the two never double-count.

// Account type from a numeric code prefix — same convention as the TB importer.
function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function glTypeFromCode(codeStr) {
  const digits = String(codeStr).replace(/[^0-9]/g, '');
  const n = parseInt(digits, 10);
  if (isNaN(n)) return null;
  // Charts of accounts vary in code width. The classic 4–5 digit scheme uses the
  // absolute number (1xxxx Asset, 2xxxx Liability, …). Wider schemes (e.g. CLRF's
  // 6-digit codes like 120100, 551100) classify by the LEADING digit instead, so
  // for any code with more than 5 digits we key off the first digit rather than
  // the magnitude (which would otherwise overflow every threshold → Revenue).
  if (digits.length > 5) {
    switch (digits[0]) {
      case '1': return 'Asset';
      case '2': return 'Liability';
      case '3': return 'Equity';
      case '4': return 'Revenue';
      case '5':
      case '6':
      case '7':
      case '8':
      case '9': return 'Expense';
      default: return null;
    }
  }
  if (n <= 19999) return 'Asset';
  if (n <= 29999) return 'Liability';
  if (n <= 39999) return 'Equity';
  if (n <= 49999) return 'Revenue';
  if (n <= 69999) return 'Expense';
  return 'Revenue';
}

// Split a fused "code + name" cell (e.g. "1000 · Cash", "1000 - Cash", "1000: Cash",
// "1000 Cash", "Cash (1000)") into { code, name }. Returns null if no leading/trailing
// numeric code can be isolated. `delimiter` (optional) forces a specific separator.
function splitCodeName(raw, delimiter) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  if (delimiter && delimiter !== 'auto') {
    const idx = s.indexOf(delimiter);
    if (idx >= 0) {
      const code = s.slice(0, idx).trim();
      const name = s.slice(idx + delimiter.length).trim();
      if (code) return { code, name };
    }
  }
  // Trailing parenthesized code: "Cash (1000)"
  let m = s.match(/^(.*?)[\s]*\((\d[\w-]*)\)\s*$/);
  if (m && m[2]) return { code: m[2].trim(), name: m[1].trim() };
  // Leading code with a separator: "1000 · Cash", "1000-Cash", "1000: Cash", "1000 | Cash"
  m = s.match(/^(\d[\w-]*?)\s*[·:|.\-–—]\s*(.+)$/);
  if (m) return { code: m[1].trim(), name: m[2].trim() };
  // Leading code separated by whitespace: "1000 Cash Operating"
  m = s.match(/^(\d[\w-]*)\s+(.+)$/);
  if (m) return { code: m[1].trim(), name: m[2].trim() };
  return null;
}

// Heuristic: does a column look like a fused code+name? Sample non-empty values and
// see if most of them split cleanly AND none is a plain number (which would be a code-only col).
function looksFused(values) {
  const sample = values.map(v => String(v == null ? '' : v).trim()).filter(Boolean).slice(0, 50);
  if (sample.length === 0) return false;
  let split = 0, plainNum = 0;
  for (const v of sample) {
    if (/^[\d.,()-]+$/.test(v)) plainNum++;
    else if (splitCodeName(v)) split++;
  }
  return plainNum === 0 && split >= Math.ceil(sample.length * 0.6);
}

const GL_NUM = s => { const raw = String(s == null ? '' : s).trim(); const neg = /^\(.*\)$/.test(raw); let v = parseFloat(raw.replace(/[,$()\s]/g, '')) || 0; return neg ? -Math.abs(v) : v; };

// Sage Intacct's "General Ledger report" export is an HTML document with a .xls
// extension, not a real spreadsheet — XLSX.read() chokes on the <!DOCTYPE>. It
// lays out one big table: an account section-header row ("<code> - <name>
// (Balance forward ...)"), then per-transaction rows, then a "Totals for ..."
// row, repeated per account. We flatten it into standard rows the normal GL
// mapper understands. Returns null if the buffer is not Intacct HTML.
function glParseIntacctHtml(buffer) {
  const text = buffer.toString('utf8');
  const lower = text.slice(0, 4000).toLowerCase();
  if (!(lower.includes('<html') || lower.includes('<!doctype') || lower.includes('<table'))) return null;
  // Only treat as Intacct GL if the account-section "Balance forward" marker is present.
  if (!/balance forward/i.test(text)) return null;
  const stripTags = s => s.replace(/<[^>]+>/g, '');
  const decode = s => s.replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
  const clean = s => decode(stripTags(s)).replace(/\s+/g, ' ').trim();
  // Scope to the report body so CSS/script before it can't pollute parsing.
  const bodyAt = text.indexOf('report_body');
  const scope = bodyAt >= 0 ? text.slice(bodyAt) : text;
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const tdRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  const HDR = /^(\S.*?)\s+-\s+([\s\S]*?)\s+\(Balance forward/i;
  const out = [];
  let cur = null; // {code, name}
  let m;
  while ((m = trRe.exec(scope))) {
    const cells = []; let c;
    tdRe.lastIndex = 0;
    while ((c = tdRe.exec(m[1]))) cells.push(clean(c[1]));
    if (!cells.length) continue;
    // Account section header: 2 cells, first contains "Balance forward".
    if (cells.length >= 2 && /balance forward/i.test(cells[0])) {
      const hm = HDR.exec(cells[0]);
      if (hm) cur = { code: hm[1].trim(), name: hm[2].trim() };
      continue;
    }
    if (/^totals for/i.test(cells[0])) continue; // account subtotal
    // Transaction row: 9 cells, first is a date, inside an account section.
    if (cur && cells.length >= 9 && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cells[0])) {
      out.push({
        'Account Number': cur.code,
        'Account Name': cur.name,
        'Posted dt.': cells[0],
        'Doc dt.': cells[1],
        'Memo/Description': cells[3],
        'Project': cells[4],
        'JNL': cells[5],
        'Debit': cells[6],
        'Credit': cells[7],
      });
    }
  }
  if (!out.length) return null;
  const columns = ['Account Number', 'Account Name', 'Posted dt.', 'Doc dt.', 'Memo/Description', 'Project', 'JNL', 'Debit', 'Credit'];
  return { columns, rows: out };
}

function glReadGrid(buffer) {
  // Intacct HTML GL export (.xls that's really HTML) — flatten it first.
  const intacct = glParseIntacctHtml(buffer);
  if (intacct) return intacct;
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  // header:1 → array-of-arrays so we can tolerate junk/blank header rows and pick the
  // row with the most non-empty cells as the header.
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false, raw: false });
  if (!aoa.length) return { columns: [], rows: [] };
  let hdrIdx = 0, best = -1;
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const filled = aoa[i].filter(c => String(c).trim()).length;
    if (filled > best) { best = filled; hdrIdx = i; }
  }
  const headerRow = aoa[hdrIdx].map((c, i) => { const t = String(c).trim(); return t || ('Column ' + (i + 1)); });
  const rows = [];
  for (let i = hdrIdx + 1; i < aoa.length; i++) {
    const obj = {}; let any = false;
    headerRow.forEach((h, j) => { const v = aoa[i][j]; obj[h] = v == null ? '' : v; if (String(v).trim()) any = true; });
    if (any) rows.push(obj);
  }
  return { columns: headerRow, rows };
}

function glAutoMap(columns, rows) {
  const norm = c => String(c).toLowerCase().trim();
  const find = (patterns, used) => {
    const pool = columns.filter(c => !used.includes(c));
    for (const pat of patterns) { const hit = pool.find(c => norm(c) === pat); if (hit) return hit; }
    for (const pat of patterns) { const hit = pool.find(c => norm(c).includes(pat)); if (hit) return hit; }
    return null;
  };
  const used = [];
  const push = c => { if (c) used.push(c); return c || null; };
  const acctNum  = push(find(['account number', 'account #', 'account code', 'acct number', 'acct code', 'gl account', 'account', 'acct', 'code'], used));
  const date     = push(find(['transaction date', 'trans date', 'posting date', 'post date', 'posted dt.', 'posted dt', 'doc dt.', 'date'], used));
  const debit    = push(find(['debit', 'dr'], used));
  const credit   = push(find(['credit', 'cr'], used));
  const acctName = push(find(['account name', 'account description', 'acct name', 'account title'], used));
  const desc     = push(find(['description', 'memo/description', 'detail', 'narrative', 'line description'], used));
  const memo     = push(find(['memo', 'note', 'notes', 'reference detail'], used));
  const ref      = push(find(['reference', 'ref', 'document number', 'doc number', 'doc #', 'entry number', 'journal number', 'transaction number', 'num', 'voucher'], used));
  const running  = push(find(['running balance', 'balance', 'ending balance', 'cumulative'], used));
  // Analytical dimensions: project (Intacct project / QBO class), class (investor), location (deal/asset).
  const project  = push(find(['project', 'project id', 'project code', 'job', 'job id'], used));
  const klass    = push(find(['item class', 'class', 'investor'], used));
  const location = push(find(['location', 'deal', 'property', 'asset'], used));
  // Detect a fused code+name column when no separate name was found.
  let fused = false, fusedCol = null;
  if (acctNum) {
    const vals = rows.map(r => r[acctNum]);
    if (!acctName && looksFused(vals)) { fused = true; fusedCol = acctNum; }
  }
  return { account_number: acctNum, account_name: acctName, transaction_date: date, description: desc, memo, debit, credit, reference: ref, running_balance: running, project, class: klass, location, fused, fused_column: fusedCol };
}

app.post('/api/entities/:eid/import-gl/preview', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), memUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { columns, rows } = glReadGrid(req.file.buffer);
    if (!columns.length || !rows.length) return res.status(400).json({ error: 'No data rows found in file' });
    const suggested = glAutoMap(columns, rows);
    res.json({
      columns,
      total_rows: rows.length,
      suggested,
      preview: rows.slice(0, 20),
    });
  } catch (e) {
    res.status(400).json({ error: 'Failed to read file: ' + e.message });
  }
});

app.post('/api/entities/:eid/import-gl', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), memUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const eid = +req.params.eid;
  let mapping;
  try { mapping = JSON.parse(req.body.mapping || '{}'); } catch { return res.status(400).json({ error: 'Invalid mapping' }); }
  const asOfLabel = req.body.as_of_date || new Date().toISOString().slice(0, 10);
  try {
    const { columns, rows } = glReadGrid(req.file.buffer);
    if (!rows.length) return res.status(400).json({ error: 'No data rows found in file' });

    const m = mapping;
    if (!m.account_number && !m.fused) return res.status(400).json({ error: 'Account Number column must be mapped' });
    if (!m.transaction_date) return res.status(400).json({ error: 'Transaction Date column must be mapped' });
    if (!m.debit && !m.credit) return res.status(400).json({ error: 'Debit and/or Credit columns must be mapped' });
    // Account name is optional when an account-number column is present: names are
    // backfilled from the number's row (or parsed from a fused "code name" cell),
    // and any account still missing a name falls back to using its code as the name.
    if (!m.account_number && !m.account_name && !m.fused) return res.status(400).json({ error: 'Map an Account Name column, or enable code+name splitting' });
    if (!m.description && !m.memo) return res.status(400).json({ error: 'A Description or Memo column must be mapped' });

    const isoDate = (v) => {
      if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
      const s = String(v == null ? '' : v).trim();
      if (!s) return null;
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      const d = new Date(s);
      return isNaN(d) ? null : d.toISOString().slice(0, 10);
    };

    // Parse each row into a normalized line.
    const parsedLines = [];
    // Name-candidate columns: when the account-number cell is blank we look here to
    // recover a "code name" or at least an account name, regardless of whether the
    // user mapped an account-name column. Prefer an explicitly mapped account_name,
    // then common name-bearing headers (e.g. QBO's "Full name").
    const nameCandidates = [];
    if (m.account_name) nameCandidates.push(m.account_name);
    for (const c of columns) {
      if (nameCandidates.includes(c)) continue;
      if ([m.account_number, m.transaction_date, m.debit, m.credit, m.reference, m.running_balance, m.class, m.location, m.project].includes(c)) continue;
      if (/full name|account name|account|name|acct/i.test(c)) nameCandidates.push(c);
    }
    const acctNames = new Map(); // code -> name
    const classNames = new Set();    // distinct class values encountered
    const locationNames = new Set(); // distinct location values encountered
    const projectNames = new Set();  // distinct project values encountered
    let skipped = 0;
    for (const row of rows) {
      let code, name;
      if (m.fused) {
        const sp = splitCodeName(row[m.fused_column || m.account_number], m.fused_delimiter);
        if (!sp) { skipped++; continue; }
        code = sp.code; name = sp.name;
        if (m.account_name && String(row[m.account_name] || '').trim()) name = String(row[m.account_name]).trim();
      } else {
        code = String(row[m.account_number] || '').trim();
        name = String(row[m.account_name] || '').trim();
        // Fallback: some exports leave the account-number cell blank but carry the
        // account elsewhere (e.g. QBO's "Full name" holds "551100 Expense" or just
        // "Retained Earnings"). Search the name-candidate columns for a usable value,
        // independent of whether an account-name column was explicitly mapped.
        if (!code) {
          let cand = name;
          if (!cand) {
            for (const c of nameCandidates) { const v = String(row[c] || '').trim(); if (v) { cand = v; break; } }
          }
          if (cand) {
            const sp = splitCodeName(cand);
            if (sp) { code = sp.code; name = sp.name; }
            else { name = cand; }
          }
        }
      }
      // Name backfill: we have a numeric code but no descriptive name (the account-
      // number column held only digits). Recover a name from the candidate columns —
      // e.g. QBO's "Full name" carries "120100 Investment in CLIP". If a candidate
      // splits into the same code, take its name; otherwise use the candidate text
      // (minus a leading copy of the code) as the name.
      if (code && !name) {
        for (const c of nameCandidates) {
          const v = String(row[c] || '').trim();
          if (!v) continue;
          const sp = splitCodeName(v);
          if (sp && sp.code === code && sp.name) { name = sp.name; break; }
          if (!sp && v !== code) { name = v; break; }
          if (sp && sp.name && sp.code !== code) { name = sp.name; break; }
        }
      }
      // Last resort: an account with a name but no resolvable numeric code (e.g.
      // "Retained Earnings"). Use the name itself as the code so the line is kept
      // and its JE still balances; dropping it would unbalance the entry.
      if (!code && name) code = name;
      if (!code) { skipped++; continue; }
      const date = isoDate(row[m.transaction_date]);
      if (!date) { skipped++; continue; }
      const dr = m.debit ? Math.abs(GL_NUM(row[m.debit])) : 0;
      const cr = m.credit ? Math.abs(GL_NUM(row[m.credit])) : 0;
      if (dr < 0.005 && cr < 0.005) { skipped++; continue; }
      const descParts = [];
      if (m.description && String(row[m.description] || '').trim()) descParts.push(String(row[m.description]).trim());
      if (m.memo && String(row[m.memo] || '').trim()) descParts.push(String(row[m.memo]).trim());
      const description = descParts.join(' — ');
      const ref = m.reference ? String(row[m.reference] || '').trim() : '';
      const running = m.running_balance ? GL_NUM(row[m.running_balance]) : null;
      const className = m.class ? String(row[m.class] || '').trim() : '';
      const locationName = m.location ? String(row[m.location] || '').trim() : '';
      const projectName = m.project ? String(row[m.project] || '').trim() : '';
      if (className) classNames.add(className);
      if (locationName) locationNames.add(locationName);
      if (projectName) projectNames.add(projectName);
      if (name && !acctNames.has(code)) acctNames.set(code, name);
      else if (!acctNames.has(code)) acctNames.set(code, '');
      parsedLines.push({ code, date, dr, cr, description, ref, running, className, locationName, projectName });
    }

    if (!parsedLines.length) return res.status(400).json({ error: 'No valid transaction rows found. Check your column mapping and that amounts are numeric.' });

    // Group into journal entries. If a reference column is mapped, group by date+ref into
    // balanced entries; otherwise group by transaction date (one JE per date). A single
    // JE may never span multiple dates — every entry shares one posting date.
    const groups = new Map();
    const useRef = !!m.reference && parsedLines.some(l => l.ref);
    // With a reference column, lines that HAVE a ref group by date+ref; lines that
    // LACK a ref (e.g. QBO bills/payments/expenses with a blank Num) group by date
    // alone, so same-day reference-less activity forms one balanced entry instead of
    // many one-line groups. Without a reference column, everything groups by date.
    parsedLines.forEach((l) => {
      const key = useRef ? (l.ref ? (l.date + '||' + l.ref) : (l.date + '||__noref__')) : l.date;
      if (!groups.has(key)) groups.set(key, { date: l.date, ref: l.ref, lines: [] });
      groups.get(key).lines.push(l);
    });

    // Balance gate: every JE must balance (debits == credits) on its own.
    // For reference grouping each date+ref group must net to zero; for date grouping
    // each date must net to zero. A balanced source GL guarantees this, so any
    // out-of-balance group signals a single-sided export or a misdated line —
    // refuse the import and report the offending groups rather than post garbage.
    {
      const unbalanced = [];
      for (const g of groups.values()) {
        const dr = g.lines.reduce((s, l) => s + l.dr, 0);
        const cr = g.lines.reduce((s, l) => s + l.cr, 0);
        if (Math.abs(dr - cr) > 0.01) {
          unbalanced.push({
            date: g.date,
            ...(useRef ? { reference: g.ref || '(none)' } : {}),
            debit: +dr.toFixed(2),
            credit: +cr.toFixed(2),
            difference: +(dr - cr).toFixed(2),
            lines: g.lines.length,
          });
        }
      }
      if (unbalanced.length) {
        unbalanced.sort((a, b) => Math.abs(b.difference) - Math.abs(a.difference));
        return res.status(400).json({
          error: useRef
            ? ('Import halted: ' + unbalanced.length + ' reference group(s) do not balance (debits ≠ credits). A balanced general ledger should net to zero within every transaction.')
            : ('Import halted: ' + unbalanced.length + ' date(s) do not balance (debits ≠ credits). When the whole GL balances, every date must balance too — an out-of-balance date usually means a single-sided export or a misdated line.'),
          grouping: useRef ? 'by_reference' : 'by_date',
          unbalanced_groups: unbalanced.slice(0, 50),
          unbalanced_count: unbalanced.length,
        });
      }
    }

    // Running-balance verification: compare each account's last imported running balance
    // (in file order) to the net debit-credit we computed for that account.
    let verification = null;
    if (m.running_balance) {
      const lastRun = new Map();   // code -> last running value seen
      const netByCode = new Map(); // code -> sum(dr-cr)
      for (const l of parsedLines) {
        if (l.running !== null && !isNaN(l.running)) lastRun.set(l.code, l.running);
        netByCode.set(l.code, (netByCode.get(l.code) || 0) + (l.dr - l.cr));
      }
      const mismatches = [];
      for (const [code, run] of lastRun) {
        const net = +(netByCode.get(code) || 0).toFixed(2);
        // running balance is on the account's natural side; compare on magnitude with sign
        if (Math.abs(net - +run.toFixed(2)) > 0.01 && Math.abs((-net) - +run.toFixed(2)) > 0.01) {
          mismatches.push({ code, computed: net, reported: +run.toFixed(2) });
        }
      }
      verification = { checked: lastRun.size, matched: lastRun.size - mismatches.length, mismatches: mismatches.slice(0, 25) };
    }

    const result = db.transaction(() => {
      // Latest import wins: clear prior TB import and prior GL import on this entity.
      db.prepare("DELETE FROM journal_entries WHERE entity_id = ? AND memo = 'Opening balance from imported trial balance'").run(eid);
      db.prepare("DELETE FROM journal_entries WHERE entity_id = ? AND memo LIKE 'GL detail import%'").run(eid);

      // Rebuild COA from the accounts encountered in the GL.
      db.prepare('DELETE FROM accounts WHERE entity_id = ?').run(eid);
      const insAcct = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)');
      // For accounts with no numeric code (code == name fallback), infer the type
      // from common equity/P&L keywords rather than defaulting everything to Asset.
      const typeFromName = (nm) => {
        const s = String(nm || '').toLowerCase();
        if (/retained earnings|equity|capital|contribution|distribution|member|partner|accumulated/.test(s)) return 'Equity';
        if (/payable|liabilit|accrued|due to|note payable|loan/.test(s)) return 'Liability';
        if (/receivable|due from|cash|bank|prepaid|investment|asset/.test(s)) return 'Asset';
        if (/income|revenue|gain/.test(s)) return 'Revenue';
        if (/expense|cost|fee|loss/.test(s)) return 'Expense';
        return 'Equity';
      };
      for (const [code, nm] of acctNames) {
        const type = glTypeFromCode(code) || typeFromName(nm || code);
        const name = nm || code;
        const isBank = type === 'Asset' && /cash|bank|checking|savings/i.test(name);
        insAcct.run(eid, code, name, type, '', isBank ? 1 : 0);
      }

      const insJE = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)');
      const insLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, description, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?,?)');

      // Resolve analytical dimensions: upsert each distinct class/location for this
      // entity and build name->id maps. Dimensions accumulate (unique per name), so
      // re-imports reuse existing ids rather than duplicating.
      const classId = new Map();
      const locationId = new Map();
      const projectId = new Map();
      if (classNames.size) {
        const insClass = db.prepare("INSERT OR IGNORE INTO dim_classes (entity_id, name, kind) VALUES (?, ?, 'investor')");
        const getClass = db.prepare('SELECT id FROM dim_classes WHERE entity_id = ? AND name = ?');
        for (const nm of classNames) { insClass.run(eid, nm); classId.set(nm, getClass.get(eid, nm).id); }
      }
      if (locationNames.size) {
        const insLoc = db.prepare("INSERT OR IGNORE INTO dim_locations (entity_id, name, kind) VALUES (?, ?, '')");
        const getLoc = db.prepare('SELECT id FROM dim_locations WHERE entity_id = ? AND name = ?');
        for (const nm of locationNames) { insLoc.run(eid, nm); locationId.set(nm, getLoc.get(eid, nm).id); }
      }
      if (projectNames.size) {
        // Intacct project values are codes (e.g. P-10100.001). Resolve against the
        // existing project catalog BY CODE first so we reuse the catalog row (and its
        // real name like "G&A") instead of creating a duplicate. Only when the code is
        // genuinely new do we create a placeholder row (name = code) to be named later.
        const getProjByCode = db.prepare('SELECT id FROM dim_projects WHERE entity_id = ? AND code = ?');
        const insProjByCode = db.prepare("INSERT INTO dim_projects (entity_id, name, code, kind) VALUES (?, ?, ?, 'project')");
        for (const nm of projectNames) {
          const found = getProjByCode.get(eid, nm);
          if (found) { projectId.set(nm, found.id); }
          else { const r = insProjByCode.run(eid, nm, nm); projectId.set(nm, r.lastInsertRowid); }
        }
      }

      let entryNum = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(eid).m || 0);
      let jeCount = 0, lineCount = 0, totalDr = 0, totalCr = 0;

      // Stable order: by date, then reference.
      const ordered = [...groups.values()].sort((a, b) => (a.date + (a.ref || '')).localeCompare(b.date + (b.ref || '')));
      for (const g of ordered) {
        entryNum++;
        const memo = useRef
          ? ('GL detail import' + (g.ref ? ' — ' + g.ref : ''))
          : ('GL detail import (' + g.date + ')');
        // Every group is single-date by construction, so the JE date is the group date.
        const jeDate = g.date;
        const r = insJE.run(eid, entryNum, jeDate, memo, req.user.name || req.user.email);
        for (const l of g.lines) {
          insLine.run(r.lastInsertRowid, l.code, l.dr, l.cr, l.description,
            l.projectName ? (projectId.get(l.projectName) || null) : null,
            l.className ? (classId.get(l.className) || null) : null,
            l.locationName ? (locationId.get(l.locationName) || null) : null);
          lineCount++; totalDr += l.dr; totalCr += l.cr;
        }
        jeCount++;
      }
      return { jeCount, lineCount, totalDr, totalCr, accounts: acctNames.size, classes: classNames.size, locations: locationNames.size, projects: projectNames.size };
    })();

    // Post-commit read-back: re-query the persisted counts on a fresh statement,
    // OUTSIDE the transaction, so the response proves the data actually landed on
    // this entity rather than merely reporting what the transaction intended to
    // write. If these disagree with `result`, the import did not persist and the
    // caller is told so explicitly instead of seeing a false success.
    const persisted = {
      entries: db.prepare('SELECT COUNT(*) AS c FROM journal_entries WHERE entity_id = ?').get(eid).c,
      accounts: db.prepare('SELECT COUNT(*) AS c FROM accounts WHERE entity_id = ?').get(eid).c,
      lines: db.prepare(
        'SELECT COUNT(*) AS c FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id WHERE je.entity_id = ?'
      ).get(eid).c,
    };
    const persistedOk = persisted.entries >= result.jeCount && persisted.lines >= result.lineCount;

    res.json({
      success: true,
      entity_id: eid,
      grouping: useRef ? 'by_reference' : 'by_date',
      entries_created: result.jeCount,
      lines_imported: result.lineCount,
      accounts_imported: result.accounts,
      classes_imported: result.classes,
      locations_imported: result.locations,
      rows_skipped: skipped,
      total_debit: +result.totalDr.toFixed(2),
      total_credit: +result.totalCr.toFixed(2),
      balanced: Math.abs(result.totalDr - result.totalCr) < 0.01,
      persisted,
      persisted_ok: persistedOk,
      verification,
    });
  } catch (e) {
    res.status(400).json({ error: 'Failed to import GL detail: ' + e.message });
  }
});

// ═══ Accounts ═══
app.get('/api/entities/:eid/accounts', auth, requireEntityAccess(), (req, res) => res.json(db.prepare('SELECT * FROM accounts WHERE entity_id = ? ORDER BY code').all(req.params.eid)));

// === Analytical dimensions (class = investor, location = deal/asset) ===
// List dimension values with how many lines reference each.
app.get('/api/entities/:eid/classes', auth, requireEntityAccess(), (req, res) =>
  res.json(db.prepare(`SELECT c.id, c.name, c.code, c.kind, COUNT(jl.id) AS line_count
    FROM dim_classes c LEFT JOIN journal_lines jl ON jl.class_id = c.id
    WHERE c.entity_id = ? GROUP BY c.id ORDER BY c.name`).all(req.params.eid)));

app.get('/api/entities/:eid/locations', auth, requireEntityAccess(), (req, res) =>
  res.json(db.prepare(`SELECT l.id, l.name, l.code, l.kind, COUNT(jl.id) AS line_count
    FROM dim_locations l LEFT JOIN journal_lines jl ON jl.location_id = l.id
    WHERE l.entity_id = ? GROUP BY l.id ORDER BY l.name`).all(req.params.eid)));

app.get('/api/entities/:eid/projects', auth, requireEntityAccess(), (req, res) =>
  res.json(db.prepare(`SELECT p.id, p.name, p.code, p.kind, COUNT(jl.id) AS line_count
    FROM dim_projects p LEFT JOIN journal_lines jl ON jl.project_id = p.id
    WHERE p.entity_id = ? GROUP BY p.id ORDER BY p.code, p.name`).all(req.params.eid)));

// ── Dimension CRUD (locations + classes). name required; code optional. ──
app.post('/api/entities/:eid/locations', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim(); if (!name) return res.status(400).json({ error: 'Name required' });
  const code = (req.body.code || '').trim() || null;
  try { const r = db.prepare('INSERT INTO dim_locations (entity_id, name, code, kind) VALUES (?, ?, ?, ?)').run(req.params.eid, name, code, req.body.kind || '');
    res.json({ id: r.lastInsertRowid, name, code, kind: req.body.kind || '', line_count: 0 }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A location with that name already exists' }); throw e; }
});
app.patch('/api/entities/:eid/locations/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM dim_locations WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const name = req.body.name !== undefined ? (req.body.name || '').trim() : row.name; if (!name) return res.status(400).json({ error: 'Name required' });
  const code = req.body.code !== undefined ? ((req.body.code || '').trim() || null) : row.code;
  const kind = req.body.kind !== undefined ? (req.body.kind || '') : row.kind;
  try { db.prepare('UPDATE dim_locations SET name = ?, code = ?, kind = ? WHERE id = ? AND entity_id = ?').run(name, code, kind, req.params.id, req.params.eid);
    res.json({ id: Number(req.params.id), name, code, kind }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A location with that name already exists' }); throw e; }
});
app.delete('/api/entities/:eid/locations/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const used = db.prepare('SELECT COUNT(*) AS n FROM journal_lines WHERE location_id = ?').get(req.params.id).n;
  if (used > 0) return res.status(409).json({ error: 'Location is used on ' + used + ' journal line(s); reassign or remove those first' });
  db.prepare('DELETE FROM dim_locations WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

app.post('/api/entities/:eid/classes', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim(); if (!name) return res.status(400).json({ error: 'Name required' });
  const code = (req.body.code || '').trim() || null;
  try { const r = db.prepare('INSERT INTO dim_classes (entity_id, name, code, kind) VALUES (?, ?, ?, ?)').run(req.params.eid, name, code, req.body.kind || 'investor');
    res.json({ id: r.lastInsertRowid, name, code, kind: req.body.kind || 'investor', line_count: 0 }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A class with that name already exists' }); throw e; }
});
app.patch('/api/entities/:eid/classes/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM dim_classes WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const name = req.body.name !== undefined ? (req.body.name || '').trim() : row.name; if (!name) return res.status(400).json({ error: 'Name required' });
  const code = req.body.code !== undefined ? ((req.body.code || '').trim() || null) : row.code;
  const kind = req.body.kind !== undefined ? (req.body.kind || 'investor') : row.kind;
  try { db.prepare('UPDATE dim_classes SET name = ?, code = ?, kind = ? WHERE id = ? AND entity_id = ?').run(name, code, kind, req.params.id, req.params.eid);
    res.json({ id: Number(req.params.id), name, code, kind }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A class with that name already exists' }); throw e; }
});
app.delete('/api/entities/:eid/classes/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const used = db.prepare('SELECT COUNT(*) AS n FROM journal_lines WHERE class_id = ?').get(req.params.id).n;
  if (used > 0) return res.status(409).json({ error: 'Class is used on ' + used + ' journal line(s); reassign or remove those first' });
  db.prepare('DELETE FROM dim_classes WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// ── Investor commitments (informational; never posts to GL). Linked to dim_classes
//    (kind='investor'). Uncalled = commitment - called; pct_called and ownership_pct
//    (commitment / total commitments) are computed on read. ──
app.get('/api/entities/:eid/commitments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const rows = db.prepare(`SELECT ic.id, ic.class_id, c.name AS investor, c.code AS investor_code,
      ic.commitment_amount, ic.called_amount, ic.commit_date, ic.notes
    FROM investor_commitments ic JOIN dim_classes c ON c.id = ic.class_id
    WHERE ic.entity_id = ? ORDER BY c.name`).all(req.params.eid);
  const totalCommit = rows.reduce((s2, r) => s2 + (r.commitment_amount || 0), 0);
  const out = rows.map(r => {
    const uncalled = (r.commitment_amount || 0) - (r.called_amount || 0);
    return {
      ...r,
      uncalled_amount: uncalled,
      pct_called: r.commitment_amount ? (r.called_amount || 0) / r.commitment_amount : 0,
      ownership_pct: totalCommit ? (r.commitment_amount || 0) / totalCommit : 0,
    };
  });
  const totals = {
    commitment_amount: totalCommit,
    called_amount: rows.reduce((s2, r) => s2 + (r.called_amount || 0), 0),
    uncalled_amount: rows.reduce((s2, r) => s2 + ((r.commitment_amount || 0) - (r.called_amount || 0)), 0),
  };
  res.json({ entity_id: parseInt(req.params.eid), investors: out, totals });
});
app.post('/api/entities/:eid/commitments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const classId = parseInt(req.body.class_id);
  if (!classId) return res.status(400).json({ error: 'class_id (investor) is required' });
  const cls = db.prepare('SELECT id FROM dim_classes WHERE id = ? AND entity_id = ?').get(classId, req.params.eid);
  if (!cls) return res.status(400).json({ error: 'Investor class not found in this entity' });
  const commitment = Number(req.body.commitment_amount || 0);
  const called = Number(req.body.called_amount || 0);
  const now = new Date().toISOString();
  try {
    const r = db.prepare(`INSERT INTO investor_commitments (entity_id, class_id, commitment_amount, called_amount, commit_date, notes, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)`).run(req.params.eid, classId, commitment, called, req.body.commit_date || null, req.body.notes || null, now, now);
    res.json({ id: r.lastInsertRowid });
  } catch (e) {
    if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'This investor already has a commitment row; edit it instead' });
    throw e;
  }
});
app.patch('/api/entities/:eid/commitments/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM investor_commitments WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const commitment = req.body.commitment_amount !== undefined ? Number(req.body.commitment_amount) : row.commitment_amount;
  const called = req.body.called_amount !== undefined ? Number(req.body.called_amount) : row.called_amount;
  const commitDate = req.body.commit_date !== undefined ? (req.body.commit_date || null) : row.commit_date;
  const notes = req.body.notes !== undefined ? (req.body.notes || null) : row.notes;
  db.prepare('UPDATE investor_commitments SET commitment_amount = ?, called_amount = ?, commit_date = ?, notes = ?, updated_at = ? WHERE id = ? AND entity_id = ?')
    .run(commitment, called, commitDate, notes, new Date().toISOString(), req.params.id, req.params.eid);
  res.json({ success: true });
});
app.delete('/api/entities/:eid/commitments/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM investor_commitments WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// ── Memorized reports (saved report configurations; shared per entity). ──
app.get('/api/entities/:eid/memorized-reports', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const rows = db.prepare(`SELECT id, report_type, name, config_json, created_by, created_by_name, created_at, updated_at
    FROM memorized_reports WHERE entity_id = ? ORDER BY report_type, name`).all(req.params.eid);
  res.json(rows.map(r => ({ ...r, config: (() => { try { return JSON.parse(r.config_json); } catch { return {}; } })() })));
});
app.post('/api/entities/:eid/memorized-reports', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim();
  const reportType = (req.body.report_type || '').trim();
  if (!name) return res.status(400).json({ error: 'Name is required' });
  if (!reportType) return res.status(400).json({ error: 'report_type is required' });
  const configJson = JSON.stringify(req.body.config || {});
  const now = new Date().toISOString();
  try {
    const r = db.prepare(`INSERT INTO memorized_reports (entity_id, report_type, name, config_json, created_by, created_by_name, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)`).run(req.params.eid, reportType, name, configJson, req.user.id, req.user.name || null, now, now);
    res.json({ id: r.lastInsertRowid });
  } catch (e) {
    if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A saved report of this type already has that name; pick another name' });
    throw e;
  }
});
app.patch('/api/entities/:eid/memorized-reports/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM memorized_reports WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const name = req.body.name !== undefined ? (req.body.name || '').trim() : row.name;
  if (!name) return res.status(400).json({ error: 'Name is required' });
  const configJson = req.body.config !== undefined ? JSON.stringify(req.body.config) : row.config_json;
  try {
    db.prepare('UPDATE memorized_reports SET name = ?, config_json = ?, updated_at = ? WHERE id = ? AND entity_id = ?')
      .run(name, configJson, new Date().toISOString(), req.params.id, req.params.eid);
    res.json({ success: true });
  } catch (e) {
    if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A saved report of this type already has that name; pick another name' });
    throw e;
  }
});
app.delete('/api/entities/:eid/memorized-reports/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM memorized_reports WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// ── Project dimension CRUD (Intacct-style projects; QBO-class equivalent). name required; code optional. ──
app.post('/api/entities/:eid/projects', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim(); if (!name) return res.status(400).json({ error: 'Name required' });
  const code = (req.body.code || '').trim() || null;
  try { const r = db.prepare('INSERT INTO dim_projects (entity_id, name, code, kind) VALUES (?, ?, ?, ?)').run(req.params.eid, name, code, req.body.kind || 'project');
    res.json({ id: r.lastInsertRowid, name, code, kind: req.body.kind || 'project', line_count: 0 }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A project with that name already exists' }); throw e; }
});
app.patch('/api/entities/:eid/projects/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM dim_projects WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const name = req.body.name !== undefined ? (req.body.name || '').trim() : row.name; if (!name) return res.status(400).json({ error: 'Name required' });
  const code = req.body.code !== undefined ? ((req.body.code || '').trim() || null) : row.code;
  const kind = req.body.kind !== undefined ? (req.body.kind || 'project') : row.kind;
  try { db.prepare('UPDATE dim_projects SET name = ?, code = ?, kind = ? WHERE id = ? AND entity_id = ?').run(name, code, kind, req.params.id, req.params.eid);
    res.json({ id: Number(req.params.id), name, code, kind }); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A project with that name already exists' }); throw e; }
});
app.delete('/api/entities/:eid/projects/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const used = db.prepare('SELECT COUNT(*) AS n FROM journal_lines WHERE project_id = ?').get(req.params.id).n;
  if (used > 0) return res.status(409).json({ error: 'Project is used on ' + used + ' journal line(s); reassign or remove those first' });
  db.prepare('DELETE FROM dim_projects WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// Bulk project catalog upsert. Body: { projects:[{code,name}], apply_all?:bool }.
// Sync semantics: for each row, match by code within the entity — update the name
// if the code exists, else create it. Never deletes (keeps journal_lines.project_id
// links intact). apply_all fans the same catalog out to every non-CLRF accounting/
// development entity (CLRF uses Location/Investor, not projects). County Line Rail
// Fund (code COUNTYLI1) is always excluded.
app.post('/api/entities/:eid/projects/bulk', auth, requireRole('Admin','Accountant'), (req, res) => {
  const rows = Array.isArray(req.body.projects) ? req.body.projects : [];
  const clean = rows
    .map(r => ({ code: String(r.code == null ? '' : r.code).trim(), name: String(r.name == null ? '' : r.name).trim() }))
    .filter(r => r.code && r.name);
  if (!clean.length) return res.status(400).json({ error: 'No valid {code,name} rows provided' });

  // Resolve target entities.
  let targets;
  if (req.body.apply_all) {
    targets = db.prepare(
      "SELECT id, name, code FROM entities WHERE entity_type IN ('accounting','development') AND code != 'COUNTYLI1'"
    ).all();
  } else {
    const e = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(req.params.eid);
    if (!e) return res.status(404).json({ error: 'Entity not found' });
    if (e.code === 'COUNTYLI1') return res.status(400).json({ error: 'County Line Rail Fund does not use projects' });
    targets = [e];
  }

  const findByCode = db.prepare('SELECT id, name FROM dim_projects WHERE entity_id = ? AND code = ?');
  const nameOwner = db.prepare('SELECT id, code FROM dim_projects WHERE entity_id = ? AND name = ?');
  const updName = db.prepare('UPDATE dim_projects SET name = ?, code = ? WHERE id = ?');
  const ins = db.prepare("INSERT INTO dim_projects (entity_id, name, code, kind) VALUES (?, ?, ?, 'project')");

  // dim_projects has UNIQUE(entity_id, name). The catalog legitimately reuses a name
  // across different codes (e.g. "Entrada 1" is both code "Entrada 1" and "P-20100.001").
  // Code is the real key, so when a name is already held by a DIFFERENT code we
  // disambiguate by suffixing the code — both rows survive and the code stays primary.
  // Every write is also wrapped so one bad row can never abort the whole batch.
  let created = 0, updated = 0, skipped = 0, failed = 0;
  const perEntity = [];
  for (const ent of targets) {
    let c = 0, u = 0, s = 0, f = 0;
    const run = db.transaction(() => {
      for (const { code, name } of clean) {
        try {
          const existing = findByCode.get(ent.id, code);
          if (existing) {
            if (existing.name === name) { s++; continue; }
            const owner = nameOwner.get(ent.id, name);
            const finalName = (owner && owner.id !== existing.id) ? (name + ' (' + code + ')') : name;
            updName.run(finalName, code, existing.id); u++;
          } else {
            const owner = nameOwner.get(ent.id, name);
            const finalName = owner ? (name + ' (' + code + ')') : name;
            ins.run(ent.id, finalName, code); c++;
          }
        } catch (e) { f++; }
      }
    });
    run();
    created += c; updated += u; skipped += s; failed += f;
    perEntity.push({ entity_id: ent.id, entity: ent.name, created: c, updated: u, skipped: s, failed: f });
  }
  res.json({ ok: true, entities: targets.length, created, updated, skipped, failed, perEntity });
});


// ══════════════ Accounts Receivable: customers ══════════════
app.get('/api/entities/:eid/ar/customers', auth, requireEntityAccess(), (req, res) =>
  res.json(db.prepare('SELECT * FROM ar_customers WHERE entity_id = ? ORDER BY active DESC, name').all(req.params.eid)));

app.post('/api/entities/:eid/ar/customers', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim(); if (!name) return res.status(400).json({ error: 'Name required' });
  const email = (req.body.email || '').trim() || null;
  const address = (req.body.address || '').trim() || null;
  const terms = Number.isFinite(+req.body.terms_days) ? +req.body.terms_days : 30;
  try { const r = db.prepare('INSERT INTO ar_customers (entity_id, name, email, address, terms_days) VALUES (?, ?, ?, ?, ?)').run(req.params.eid, name, email, address, terms);
    res.json(db.prepare('SELECT * FROM ar_customers WHERE id = ?').get(r.lastInsertRowid)); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A customer with that name already exists' }); throw e; }
});

app.patch('/api/entities/:eid/ar/customers/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM ar_customers WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const name = req.body.name !== undefined ? (req.body.name || '').trim() : row.name; if (!name) return res.status(400).json({ error: 'Name required' });
  const email = req.body.email !== undefined ? ((req.body.email || '').trim() || null) : row.email;
  const address = req.body.address !== undefined ? ((req.body.address || '').trim() || null) : row.address;
  const terms = req.body.terms_days !== undefined && Number.isFinite(+req.body.terms_days) ? +req.body.terms_days : row.terms_days;
  const active = req.body.active !== undefined ? (req.body.active ? 1 : 0) : row.active;
  try { db.prepare('UPDATE ar_customers SET name = ?, email = ?, address = ?, terms_days = ?, active = ? WHERE id = ? AND entity_id = ?').run(name, email, address, terms, active, req.params.id, req.params.eid);
    res.json(db.prepare('SELECT * FROM ar_customers WHERE id = ?').get(req.params.id)); }
  catch(e) { if (String(e.message).includes('UNIQUE')) return res.status(409).json({ error: 'A customer with that name already exists' }); throw e; }
});

app.delete('/api/entities/:eid/ar/customers/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const inv = db.prepare('SELECT COUNT(*) AS n FROM ar_invoices WHERE customer_id = ?').get(req.params.id).n;
  if (inv > 0) return res.status(409).json({ error: 'Customer has ' + inv + ' invoice(s); deactivate instead of deleting' });
  db.prepare('DELETE FROM ar_customers WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

// Dimension balance report: net (debit-credit) per dimension value, optionally
// restricted to a set of account codes and/or as-of date. Used for
// "capitalized deal cost by location" (accounts=investment accts, dim=location)
// and investor-level balances (dim=class). Query params:
//   dim=class|location (default location)
//   accounts=120100,120200 (comma list; omit for all accounts)
//   account_prefix=120 (alternative to accounts; matches code prefix)
//   kind=deal (restrict to dimension values of this kind)
//   as_of=YYYY-MM-DD (entries on/before this date)
app.get('/api/entities/:eid/dimension-balances', auth, requireEntityAccess(), (req, res) => {
  const eid = req.params.eid;
  const dim = req.query.dim === 'class' ? 'class' : req.query.dim === 'project' ? 'project' : 'location';
  const dimTable = dim === 'class' ? 'dim_classes' : dim === 'project' ? 'dim_projects' : 'dim_locations';
  const dimCol = dim === 'class' ? 'class_id' : dim === 'project' ? 'project_id' : 'location_id';
  const params = [eid];
  let acctClause = '';
  if (req.query.accounts) {
    const codes = String(req.query.accounts).split(',').map(s => s.trim()).filter(Boolean);
    if (codes.length) { acctClause = ` AND jl.account_code IN (${codes.map(() => '?').join(',')})`; params.push(...codes); }
  } else if (req.query.account_prefix) {
    acctClause = ' AND jl.account_code LIKE ?'; params.push(String(req.query.account_prefix) + '%');
  }
  let dateClause = '';
  if (req.query.as_of) { dateClause = ' AND je.date <= ?'; params.push(req.query.as_of); }
  let kindClause = '';
  if (req.query.kind) { kindClause = ' AND d.kind = ?'; params.push(req.query.kind); }
  const rows = db.prepare(`
    SELECT d.id, d.name, d.kind,
           SUM(jl.debit) AS total_debit, SUM(jl.credit) AS total_credit,
           SUM(jl.debit - jl.credit) AS net, COUNT(jl.id) AS line_count
    FROM journal_lines jl
    JOIN journal_entries je ON jl.entry_id = je.id
    JOIN ${dimTable} d ON d.id = jl.${dimCol}
    WHERE je.entity_id = ?${acctClause}${dateClause}${kindClause}
    GROUP BY d.id ORDER BY net DESC
  `).all(...params);
  const total = rows.reduce((s, r) => s + (r.net || 0), 0);
  res.json({
    dimension: dim,
    rows: rows.map(r => ({ id: r.id, name: r.name, kind: r.kind,
      total_debit: +(r.total_debit || 0).toFixed(2), total_credit: +(r.total_credit || 0).toFixed(2),
      net: +(r.net || 0).toFixed(2), line_count: r.line_count })),
    total_net: +total.toFixed(2),
  });
});

// Pivot report: dimension (class/location/project) × account matrix. Rows are
// dimension members, columns are accounts, cells are the net (debit-credit) sum.
// Used for PCAP-style letters: totals by investor class across contribution /
// accumulated accounts. Accepts the same account selection (accounts=csv or
// account_prefix) and as_of as dimension-balances.
app.get('/api/entities/:eid/pivot', auth, requireEntityAccess(), (req, res) => {
  const eid = req.params.eid;
  const dim = req.query.dim === 'location' ? 'location' : req.query.dim === 'project' ? 'project' : 'class';
  const dimTable = dim === 'class' ? 'dim_classes' : dim === 'project' ? 'dim_projects' : 'dim_locations';
  const dimCol = dim === 'class' ? 'class_id' : dim === 'project' ? 'project_id' : 'location_id';
  const params = [eid];
  let acctClause = '';
  if (req.query.accounts) {
    const codes = String(req.query.accounts).split(',').map(s => s.trim()).filter(Boolean);
    if (codes.length) { acctClause = ` AND jl.account_code IN (${codes.map(() => '?').join(',')})`; params.push(...codes); }
  } else if (req.query.account_prefix) {
    acctClause = ' AND jl.account_code LIKE ?'; params.push(String(req.query.account_prefix) + '%');
  }
  let dateClause = '';
  if (req.query.from) { dateClause += ' AND je.date >= ?'; params.push(req.query.from); }
  if (req.query.to) { dateClause += ' AND je.date <= ?'; params.push(req.query.to); }
  else if (req.query.as_of) { dateClause += ' AND je.date <= ?'; params.push(req.query.as_of); }
  const rows = db.prepare(`
    SELECT d.id AS dim_id, d.name AS dim_name,
           jl.account_code, a.name AS account_name, a.type AS account_type,
           SUM(jl.debit - jl.credit) AS net
    FROM journal_lines jl
    JOIN journal_entries je ON jl.entry_id = je.id
    JOIN ${dimTable} d ON d.id = jl.${dimCol}
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    WHERE je.entity_id = ?${acctClause}${dateClause}
    GROUP BY d.id, jl.account_code
  `).all(...params);
  // Assemble columns (accounts seen) and a row per dimension member.
  const colMap = new Map(); // code -> {code,name}
  const rowMap = new Map(); // dim_id -> {id,name,cells:{code:net}, total}
  for (const r of rows) {
    if (!colMap.has(r.account_code)) colMap.set(r.account_code, { code: r.account_code, name: r.account_name || '' });
    if (!rowMap.has(r.dim_id)) rowMap.set(r.dim_id, { id: r.dim_id, name: r.dim_name, cells: {}, total: 0 });
    const net = +(r.net || 0).toFixed(2);
    rowMap.get(r.dim_id).cells[r.account_code] = net;
    rowMap.get(r.dim_id).total = +(rowMap.get(r.dim_id).total + net).toFixed(2);
  }
  const columns = [...colMap.values()].sort((a, b) => a.code.localeCompare(b.code));
  const outRows = [...rowMap.values()].sort((a, b) => a.name.localeCompare(b.name));
  // Column totals + grand total.
  const colTotals = {}; let grand = 0;
  for (const c of columns) { colTotals[c.code] = +outRows.reduce((s, r) => s + (r.cells[c.code] || 0), 0).toFixed(2); grand += colTotals[c.code]; }
  res.json({ dimension: dim, columns, rows: outRows, column_totals: colTotals, grand_total: +grand.toFixed(2) });
});

app.post('/api/entities/:eid/accounts', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { code, name, type, subtype, bank_acct } = req.body; if (!code||!name||!type) return res.status(400).json({ error: 'Required' });
  try { const r = db.prepare('INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, ?, ?, ?, ?, ?)').run(req.params.eid, code, name, type, subtype||'', bank_acct?1:0);
    res.json({ id: r.lastInsertRowid, code, name, type, subtype: subtype||'', bank_acct: bank_acct?1:0, entity_id: +req.params.eid }); }
  catch(e) { if (e.message.includes('UNIQUE')) return res.status(400).json({ error: 'Code exists' }); throw e; }
});
app.delete('/api/entities/:eid/accounts/:code', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  if (db.prepare('SELECT COUNT(*) as c FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id WHERE je.entity_id=? AND jl.account_code=?').get(req.params.eid, req.params.code).c > 0)
    return res.status(400).json({ error: 'Has transactions' });
  db.prepare('DELETE FROM accounts WHERE entity_id=? AND code=?').run(req.params.eid, req.params.code); res.json({ success: true });
});

app.put('/api/entities/:eid/accounts/:code', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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
app.get('/api/entities/:eid/entries', auth, requireEntityAccess(), (req, res) => {
  const { from, to } = req.query; let sql = 'SELECT * FROM journal_entries WHERE entity_id = ?'; const params = [req.params.eid];
  if (from) { sql += ' AND date >= ?'; params.push(from); } if (to) { sql += ' AND date <= ?'; params.push(to); }
  sql += ' ORDER BY entry_num ASC';
  const entries = db.prepare(sql).all(...params);
  const lineStmt = db.prepare(`SELECT jl.*, dp.name AS project_name, dp.code AS project_code,
      dc.name AS class_name, dl.name AS location_name
    FROM journal_lines jl
    LEFT JOIN dim_projects dp ON dp.id = jl.project_id
    LEFT JOIN dim_classes dc ON dc.id = jl.class_id
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    WHERE jl.entry_id = ?`);
  const attachStmt = db.prepare('SELECT id, original_name, mime_type, size FROM journal_attachments WHERE entry_id = ?');
  res.json(entries.map(e => ({ ...e, lines: lineStmt.all(e.id), attachments: attachStmt.all(e.id) })));
});

// GL detail (flat lines) for a printable/exportable general ledger, optionally
// filtered by location or class. Returns one row per journal line with its entry
// date/num/memo, account code+name+type, dr/cr, description, and dimension names,
// plus a running balance per account (ordered by date, entry_num). When a
// location_id/class_id is given, only lines carrying that tag are returned — by
// design an untagged line is not part of any location's ledger.
app.get('/api/entities/:eid/gl-detail', auth, requireEntityAccess(), (req, res) => {
  const { from, to, location_id, class_id, project_id, account_code } = req.query;
  const params = [req.params.eid];
  let where = '';
  if (from) { where += ' AND je.date >= ?'; params.push(from); }
  if (to) { where += ' AND je.date <= ?'; params.push(to); }
  if (location_id) { where += ' AND jl.location_id = ?'; params.push(location_id); }
  if (class_id) { where += ' AND jl.class_id = ?'; params.push(class_id); }
  if (project_id) { where += ' AND jl.project_id = ?'; params.push(project_id); }
  if (account_code) { where += ' AND jl.account_code = ?'; params.push(account_code); }
  const rows = db.prepare(`
    SELECT jl.id AS line_id, je.id AS entry_id, je.entry_num, je.date, je.memo,
           jl.account_code, a.name AS account_name, a.type AS account_type,
           jl.debit, jl.credit, jl.description,
           dc.name AS class_name, dl.name AS location_name, dp.name AS project_name, dp.code AS project_code
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    LEFT JOIN dim_classes dc ON dc.id = jl.class_id
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    LEFT JOIN dim_projects dp ON dp.id = jl.project_id
    WHERE je.entity_id = ?${where}
    ORDER BY jl.account_code, je.date, je.entry_num, jl.id
  `).all(...params);
  // Running balance per account (natural side: Asset/Expense are debit-positive).
  const run = new Map();
  const out = rows.map(r => {
    const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
    const delta = isDr ? (r.debit - r.credit) : (r.credit - r.debit);
    const bal = (run.get(r.account_code) || 0) + delta;
    run.set(r.account_code, bal);
    return {
      line_id: r.line_id, entry_id: r.entry_id, entry_num: r.entry_num, date: r.date, memo: r.memo,
      account_code: r.account_code, account_name: r.account_name, account_type: r.account_type,
      debit: +(r.debit || 0).toFixed(2), credit: +(r.credit || 0).toFixed(2),
      description: r.description || '', class_name: r.class_name || '', location_name: r.location_name || '',
      running_balance: +bal.toFixed(2),
    };
  });
  const totalDr = +out.reduce((s, r) => s + r.debit, 0).toFixed(2);
  const totalCr = +out.reduce((s, r) => s + r.credit, 0).toFixed(2);
  res.json({ lines: out, count: out.length, total_debit: totalDr, total_credit: totalCr });
});

app.post('/api/entities/:eid/entries', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { date, memo, lines } = req.body; if (!date||!memo||!lines||lines.length<2) return res.status(400).json({ error: 'Invalid' });
  const tDr = lines.reduce((s,l) => s+(l.debit||0), 0); const tCr = lines.reduce((s,l) => s+(l.credit||0), 0);
  if (Math.abs(tDr-tCr) > 0.005) return res.status(400).json({ error: 'Must balance' });
  const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
  const result = db.transaction(() => {
    const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)').run(req.params.eid, num, date, memo, req.user.name);
    for (const l of lines) db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, description, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?,?)').run(r.lastInsertRowid, l.account_code, l.debit||0, l.credit||0, l.description||'', l.project_id||null, l.class_id||null, l.location_id||null);
    return r.lastInsertRowid;
  })();
  res.json({ id: result, entry_num: num });
});

// ─── Bulk journal-entry upload (one journal LINE per row) ────────────────────
// Spreadsheet layout (header names matched case-insensitively, fuzzy):
//   Date | Account # | Account Description | Debit | Credit | Memo? | Location? | Class?
// Each row is one journal line: an amount goes in Debit OR Credit for the given
// account. Lines are grouped into journal entries by DATE — all lines sharing a
// date form one entry, which must balance (sum debits == sum credits). Memo is
// optional; the first non-empty memo in a date group is used for the entry.
// Required columns: Date, Account #, and a Debit and Credit column.
function parseBulkJE(buffer, eid) {
  const { columns, rows } = glReadGrid(buffer);
  const norm = c => String(c).toLowerCase().replace(/[^a-z0-9]/g, '');
  const pick = (cands, exclude = []) => {
    for (const want of cands) {
      const hit = columns.find(c => norm(c) === want && !exclude.includes(c));
      if (hit) return hit;
    }
    for (const want of cands) {
      const hit = columns.find(c => norm(c).includes(want) && !exclude.includes(c));
      if (hit) return hit;
    }
    return null;
  };
  const colDate    = pick(['date', 'postingdate', 'transactiondate', 'gldate']);
  const colAcct    = pick(['account', 'accountnumber', 'acct', 'acctnumber', 'accountno', 'glaccount', 'code']);
  const colAcctDesc= pick(['accountdescription', 'accountname', 'acctdescription', 'acctname']);
  const colDebit   = pick(['debit', 'dr', 'debitamount']);
  const colCredit  = pick(['credit', 'cr', 'creditamount']);
  const colMemo    = pick(['memo', 'entrymemo', 'description', 'desc'], [colAcctDesc].filter(Boolean));
  const colLoc     = pick(['location', 'locationname', 'locationcode', 'deal']);
  const colClass   = pick(['class', 'investor', 'investorclass', 'classname', 'classcode']);

  const missing = [];
  if (!colDate) missing.push('Date');
  if (!colAcct) missing.push('Account #');
  if (!colDebit) missing.push('Debit');
  if (!colCredit) missing.push('Credit');
  if (missing.length) {
    return { error: 'Missing required column(s): ' + missing.join(', ') + '. Found columns: ' + columns.join(', ') };
  }

  const accountRows = db.prepare('SELECT code, name FROM accounts WHERE entity_id=?').all(eid);
  const accounts = new Map(accountRows.map(a => [String(a.code), a.name]));
  const locRows = db.prepare('SELECT id, name, code FROM dim_locations WHERE entity_id=?').all(eid);
  const classRows = db.prepare('SELECT id, name, code FROM dim_classes WHERE entity_id=?').all(eid);
  const dimLookup = (rs) => { const m = new Map(); for (const r of rs) { if (r.name) m.set(String(r.name).toLowerCase().trim(), r.id); if (r.code) m.set(String(r.code).toLowerCase().trim(), r.id); } return m; };
  const locMap = dimLookup(locRows);
  const classMap = dimLookup(classRows);

  const toISO = (v) => {
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    const s = String(v == null ? '' : v).trim();
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/); // MM/DD/YYYY
    if (m) { let [, mo, da, yr] = m; if (yr.length === 2) yr = '20' + yr; return yr + '-' + mo.padStart(2, '0') + '-' + da.padStart(2, '0'); }
    const d = new Date(s);
    if (!isNaN(d)) return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    return null;
  };

  // 1) Parse each row into a line (with per-row errors).
  const lines = [];
  rows.forEach((row, idx) => {
    const rowNum = idx + 1;
    const errors = [];
    const dateISO = toISO(row[colDate]);
    const acctCode = String(row[colAcct] == null ? '' : row[colAcct]).trim().replace(/\.0$/, '');
    const memo = colMemo ? String(row[colMemo] == null ? '' : row[colMemo]).trim() : '';
    const debit = Math.abs(GL_NUM(row[colDebit]));
    const credit = Math.abs(GL_NUM(row[colCredit]));

    if (!dateISO) errors.push('invalid or missing Date');
    if (!acctCode) errors.push('missing Account #');
    else if (!accounts.has(acctCode)) errors.push('account ' + acctCode + ' not in chart of accounts');
    if (debit > 0 && credit > 0) errors.push('row has both a Debit and a Credit — put the amount in one column');
    if (!(debit > 0) && !(credit > 0)) errors.push('row has no Debit or Credit amount');

    let location_id = null;
    if (colLoc) { const raw = String(row[colLoc] == null ? '' : row[colLoc]).trim(); if (raw) { const id = locMap.get(raw.toLowerCase()); if (id) location_id = id; else errors.push('location "' + raw + '" not found for this entity'); } }
    let class_id = null;
    if (colClass) { const raw = String(row[colClass] == null ? '' : row[colClass]).trim(); if (raw) { const id = classMap.get(raw.toLowerCase()); if (id) class_id = id; else errors.push('class "' + raw + '" not found for this entity'); } }

    lines.push({
      row: rowNum, date: dateISO, account_code: acctCode,
      account_name: accounts.get(acctCode) || '', memo,
      debit: +debit.toFixed(2), credit: +credit.toFixed(2),
      location_id, class_id, errors,
    });
  });

  // 2) Group lines into entries by date (only lines with a usable date group).
  const groups = new Map(); // dateISO -> { date, lines:[], rows:[] }
  for (const ln of lines) {
    const key = ln.date || ('__row' + ln.row); // ungrouped (bad date) lines become singletons
    if (!groups.has(key)) groups.set(key, { date: ln.date, lines: [], rows: [] });
    groups.get(key).lines.push(ln);
    groups.get(key).rows.push(ln.row);
  }

  const entries = [];
  for (const g of groups.values()) {
    const lineErrors = g.lines.some(l => l.errors.length > 0);
    const tDr = g.lines.reduce((s, l) => s + l.debit, 0);
    const tCr = g.lines.reduce((s, l) => s + l.credit, 0);
    const balanced = Math.abs(tDr - tCr) <= 0.005;
    const memo = (g.lines.find(l => l.memo) || {}).memo || '';
    const entryErrors = [];
    if (!g.date) entryErrors.push('invalid or missing Date');
    if (g.lines.length < 2) entryErrors.push('a journal entry needs at least 2 lines on the same date');
    if (!balanced) entryErrors.push('does not balance (debits ' + tDr.toFixed(2) + ' \u2260 credits ' + tCr.toFixed(2) + ')');
    entries.push({
      date: g.date, memo, rows: g.rows,
      lines: g.lines.map(l => ({ row: l.row, account_code: l.account_code, account_name: l.account_name, debit: l.debit, credit: l.credit, location_id: l.location_id, class_id: l.class_id, errors: l.errors })),
      total_debit: +tDr.toFixed(2), total_credit: +tCr.toFixed(2),
      valid: !lineErrors && entryErrors.length === 0,
      errors: entryErrors,
    });
  }
  entries.sort((a, b) => (a.date || '').localeCompare(b.date || '') || (a.rows[0] - b.rows[0]));

  const valid = entries.filter(e => e.valid).length;
  return {
    columns,
    mapped: { date: colDate, account: colAcct, account_desc: colAcctDesc, debit: colDebit, credit: colCredit, memo: colMemo, location: colLoc, class: colClass },
    entries, total: entries.length, valid, invalid: entries.length - valid,
    line_count: lines.length,
  };
}

// Preview: parse the uploaded sheet and return validated rows (nothing posted).
app.post('/api/entities/:eid/entries/bulk/preview', auth, requireEntityAccess(), requireRole('Admin','Accountant'), memUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const result = parseBulkJE(req.file.buffer, req.params.eid);
    if (result.error) return res.status(400).json({ error: result.error });
    res.json(result);
  } catch (e) {
    res.status(400).json({ error: 'Failed to read spreadsheet: ' + e.message });
  }
});

// Commit: post the confirmed entries. Body: { entries: [{date, memo, debit_account,
// credit_account, amount, line_description, location_id, class_id}] }. Each becomes
// one balanced 2-line journal entry. All-or-nothing within a single transaction.
app.post('/api/entities/:eid/entries/bulk', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const eid = req.params.eid;
  const list = Array.isArray(req.body && req.body.entries) ? req.body.entries : [];
  if (!list.length) return res.status(400).json({ error: 'No entries to post' });

  const accounts = new Set(db.prepare('SELECT code FROM accounts WHERE entity_id=?').all(eid).map(a => String(a.code)));
  // Re-validate server-side; never trust the client. Each entry is { date, memo,
  // lines:[{account_code, debit, credit, location_id?, class_id?}] } and must balance.
  for (const [i, e] of list.entries()) {
    if (!e.date) return res.status(400).json({ error: 'Entry ' + (i + 1) + ': missing date' });
    const lines = Array.isArray(e.lines) ? e.lines : [];
    if (lines.length < 2) return res.status(400).json({ error: 'Entry ' + (i + 1) + ': needs at least 2 lines' });
    let tDr = 0, tCr = 0;
    for (const l of lines) {
      if (!accounts.has(String(l.account_code))) return res.status(400).json({ error: 'Entry ' + (i + 1) + ': account ' + l.account_code + ' not in chart of accounts' });
      tDr += Math.abs(Number(l.debit) || 0); tCr += Math.abs(Number(l.credit) || 0);
    }
    if (Math.abs(tDr - tCr) > 0.005) return res.status(400).json({ error: 'Entry ' + (i + 1) + ': does not balance' });
  }

  const posted = db.transaction(() => {
    let num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(eid).m || 0);
    const insEntry = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)');
    const insLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, description, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?,?)');
    const ids = [];
    for (const e of list) {
      num += 1;
      const r = insEntry.run(eid, num, e.date, e.memo || '', req.user.name || req.user.email);
      for (const l of e.lines) {
        insLine.run(r.lastInsertRowid, String(l.account_code), +Math.abs(Number(l.debit) || 0).toFixed(2), +Math.abs(Number(l.credit) || 0).toFixed(2), '', null, l.class_id || null, l.location_id || null);
      }
      ids.push({ id: r.lastInsertRowid, entry_num: num });
    }
    return ids;
  })();

  res.json({ posted: posted.length, entries: posted });
});

app.delete('/api/entities/:eid/entries/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const atts = db.prepare('SELECT filename FROM journal_attachments WHERE entry_id=?').all(req.params.id);
  atts.forEach(a => { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} });
  db.prepare('DELETE FROM journal_entries WHERE id=? AND entity_id=?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

app.put('/api/entities/:eid/entries/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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
    for (const l of lines) db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, description, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?,?)').run(req.params.id, l.account_code, l.debit || 0, l.credit || 0, l.description || '', l.project_id || null, l.class_id || null, l.location_id || null);
  })();
  res.json({ success: true, entry_num: entry.entry_num });
});

// ═══ Journal Attachments ═══
app.post('/api/entities/:eid/entries/:id/attachments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), upload.array('files', 10), (req, res) => {
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
app.get('/api/entities/:eid/bank-transactions', auth, requireEntityAccess(), (req, res) => {
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

app.post('/api/entities/:eid/bank-transactions/upload', auth, requireEntityAccess(), requireRole('Admin','Accountant'), memUpload.single('file'), async (req, res) => {
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

app.put('/api/entities/:eid/bank-transactions/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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
app.put('/api/entities/:eid/bank-transactions/:id/splits', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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

app.post('/api/entities/:eid/bank-transactions/post', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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

// Bank matching (Q4): find already-posted JEs that a pending bank line could be
// matched to, instead of creating a new JE. A candidate is a JE that hits this
// bank account with a net effect on the bank line equal to the bank txn amount
// (deposit => debit to bank; payment => credit to bank), dated within ±7 days,
// and not already linked to another bank transaction. Exact amount, one-to-one.
app.get('/api/entities/:eid/bank-transactions/:id/match-candidates', auth, requireEntityAccess(), (req, res) => {
  const t = db.prepare('SELECT * FROM bank_transactions WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!t) return res.status(404).json({ error: 'Transaction not found' });
  const WINDOW_DAYS = 7;
  const lo = new Date(t.date); lo.setDate(lo.getDate() - WINDOW_DAYS);
  const hi = new Date(t.date); hi.setDate(hi.getDate() + WINDOW_DAYS);
  const loS = lo.toISOString().slice(0, 10), hiS = hi.toISOString().slice(0, 10);
  const abs = +Math.abs(t.amount).toFixed(2);
  // Net effect on the bank account line per JE = debit - credit. A deposit
  // (amount>0) should have a positive bank-line net = abs; a payment, -abs.
  const wantNet = t.amount > 0 ? abs : -abs;
  // JE ids already linked to a bank txn (matched or posted) — exclude them.
  const linked = new Set(db.prepare("SELECT je_id FROM bank_transactions WHERE entity_id=? AND je_id IS NOT NULL").all(req.params.eid).map(r => r.je_id)
    .concat(db.prepare("SELECT matched_entry_id FROM bank_transactions WHERE entity_id=? AND matched_entry_id IS NOT NULL").all(req.params.eid).map(r => r.matched_entry_id)));
  const rows = db.prepare(`
    SELECT je.id, je.entry_num, je.date, je.memo,
           SUM(jl.debit - jl.credit) AS bank_net
    FROM journal_entries je
    JOIN journal_lines jl ON jl.entry_id = je.id AND jl.account_code = ?
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ?
    GROUP BY je.id
  `).all(t.bank_account_code, req.params.eid, loS, hiS);
  const candidates = rows
    .filter(r => Math.abs((r.bank_net || 0) - wantNet) < 0.005 && !linked.has(r.id))
    .map(r => ({ je_id: r.id, entry_num: r.entry_num, date: r.date, memo: r.memo, bank_net: +(r.bank_net || 0).toFixed(2),
      date_diff: Math.round((new Date(r.date) - new Date(t.date)) / 86400000) }))
    .sort((a, b) => Math.abs(a.date_diff) - Math.abs(b.date_diff));
  res.json({ transaction: { id: t.id, date: t.date, amount: t.amount, description: t.description }, candidates });
});

// Match a bank transaction to an existing JE (no new JE created).
app.post('/api/entities/:eid/bank-transactions/:id/match', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { je_id } = req.body;
  if (!je_id) return res.status(400).json({ error: 'je_id required' });
  const t = db.prepare("SELECT * FROM bank_transactions WHERE id=? AND entity_id=?").get(req.params.id, req.params.eid);
  if (!t) return res.status(404).json({ error: 'Transaction not found' });
  if (t.status === 'posted') return res.status(400).json({ error: 'Already posted as its own JE' });
  const je = db.prepare('SELECT id, entry_num FROM journal_entries WHERE id=? AND entity_id=?').get(je_id, req.params.eid);
  if (!je) return res.status(404).json({ error: 'Journal entry not found' });
  const already = db.prepare("SELECT id FROM bank_transactions WHERE entity_id=? AND id!=? AND (je_id=? OR matched_entry_id=?)").get(req.params.eid, req.params.id, je_id, je_id);
  if (already) return res.status(400).json({ error: 'That JE is already linked to another bank transaction' });
  db.prepare("UPDATE bank_transactions SET status='matched', matched_entry_id=?, je_id=? WHERE id=?").run(je_id, je_id, req.params.id);
  res.json({ matched: true, je_id, entry_num: je.entry_num });
});

// Unmatch: revert a matched bank transaction back to pending.
app.post('/api/entities/:eid/bank-transactions/:id/unmatch', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const t = db.prepare("SELECT * FROM bank_transactions WHERE id=? AND entity_id=?").get(req.params.id, req.params.eid);
  if (!t) return res.status(404).json({ error: 'Transaction not found' });
  if (t.status !== 'matched') return res.status(400).json({ error: 'Not a matched transaction' });
  db.prepare("UPDATE bank_transactions SET status=?, matched_entry_id=NULL, je_id=NULL WHERE id=?").run(t.account_code ? 'coded' : 'pending', req.params.id);
  res.json({ unmatched: true });
});

app.delete('/api/entities/:eid/bank-transactions/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM bank_transactions WHERE id=? AND entity_id=? AND status != ?').run(req.params.id, req.params.eid, 'posted');
  res.json({ success: true });
});

app.delete('/api/entities/:eid/bank-transactions/batch/:batchId', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const r = db.prepare('DELETE FROM bank_transactions WHERE entity_id=? AND batch_id=? AND status != ?').run(req.params.eid, req.params.batchId, 'posted');
  res.json({ deleted: r.changes });
});

// ═══ Balances (with soft close) ═══
app.get('/api/entities/:eid/balances', auth, requireEntityAccess(), (req, res) => {
  const { as_of, from, to, close_pl_before, location_id, class_id } = req.query;
  let dateFilter = ''; const params = [req.params.eid];
  if (as_of) { dateFilter = ' AND je.date <= ?'; params.push(as_of); }
  else if (from && to) { dateFilter = ' AND je.date >= ? AND je.date <= ?'; params.push(from, to); }
  else if (from) { dateFilter = ' AND je.date >= ?'; params.push(from); }
  else if (to) { dateFilter = ' AND je.date <= ?'; params.push(to); }
  // Dimension filter: restrict to lines tagged with a specific location/class. Used
  // for a per-location trial balance. Only location-tagged lines are picked up, by
  // design — an untagged line belongs to no location's TB.
  let dimFilter = '';
  if (location_id) { dimFilter += ' AND jl.location_id = ?'; params.push(location_id); }
  if (class_id) { dimFilter += ' AND jl.class_id = ?'; params.push(class_id); }

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

  const rows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=?${dateFilter}${dimFilter} GROUP BY jl.account_code`).all(...params);
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

app.get('/api/entities/:eid/files', auth, requireEntityAccess(), (req, res) => {
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

app.post('/api/entities/:eid/files', auth, requireEntityAccess(), requireRole('Admin','Accountant'), workpaperUploadMw, (req, res) => {
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

app.post('/api/entities/:eid/folders', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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

app.delete('/api/entities/:eid/folders', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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
app.put('/api/entities/:eid/folders/rename', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
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
app.get('/api/entities/:eid/reconciliations', auth, requireEntityAccess(), (req, res) => res.json(db.prepare('SELECT * FROM reconciliations WHERE entity_id=? ORDER BY completed_at DESC').all(req.params.eid)));
app.get('/api/entities/:eid/cleared/:accountCode', auth, requireEntityAccess(), (req, res) => {
  const m={}; db.prepare('SELECT entry_id, line_index FROM cleared_items WHERE entity_id=? AND account_code=?').all(req.params.eid, req.params.accountCode).forEach(c=>{m[c.entry_id+'-'+c.line_index]=true;}); res.json(m);
});
app.post('/api/entities/:eid/reconciliations', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { account_code, statement_date, statement_balance, book_balance, cleared_keys } = req.body;
  if (!account_code||!statement_date||statement_balance==null) return res.status(400).json({ error: 'Missing fields' });
  const result = db.transaction(() => {
    const r = db.prepare('INSERT INTO reconciliations (entity_id, account_code, statement_date, statement_balance, book_balance, cleared_count, completed_by) VALUES (?,?,?,?,?,?,?)').run(req.params.eid, account_code, statement_date, statement_balance, book_balance, cleared_keys?.length||0, req.user.name);
    if (cleared_keys) for (const k of cleared_keys) { const [eid,li]=k.split('-').map(Number); db.prepare('INSERT OR IGNORE INTO cleared_items (entity_id, account_code, entry_id, line_index, reconciliation_id) VALUES (?,?,?,?,?)').run(req.params.eid, account_code, eid, li, r.lastInsertRowid); }
    return r.lastInsertRowid;
  })(); res.json({ id: result });
});

// Bank reconciliation report — QBO-style summary + cleared/uncleared detail for a
// single completed reconciliation. Assembled from the stored reconciliation row,
// its cleared_items, and the account's journal lines. Returns structured JSON the
// client renders as a printable report.
app.get('/api/entities/:eid/reconciliations/:id/report', auth, requireEntityAccess(), (req, res) => {
  const eid = req.params.eid;
  const rec = db.prepare('SELECT * FROM reconciliations WHERE id=? AND entity_id=?').get(req.params.id, eid);
  if (!rec) return res.status(404).json({ error: 'Reconciliation not found' });
  const entity = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
  const acct = db.prepare('SELECT code, name FROM accounts WHERE entity_id=? AND code=?').get(eid, rec.account_code);

  // All journal lines hitting this bank account (signed: debit - credit, since a
  // bank account is an Asset). Each carries the entry date, number, and memo.
  const lineRows = db.prepare(`
    SELECT je.id AS entry_id, je.entry_num, je.date, je.memo,
           jl.id AS line_id, jl.debit, jl.credit,
           (SELECT COUNT(*) FROM journal_lines x WHERE x.entry_id=je.id AND x.id<=jl.id) - 1 AS line_index
    FROM journal_lines jl
    JOIN journal_entries je ON jl.entry_id = je.id
    WHERE je.entity_id = ? AND jl.account_code = ?
    ORDER BY je.date, je.id, jl.id
  `).all(eid, rec.account_code);

  // Which (entry_id,line_index) were cleared as part of THIS reconciliation.
  const clearedSet = new Set(
    db.prepare('SELECT entry_id, line_index FROM cleared_items WHERE reconciliation_id=?')
      .all(req.params.id).map(c => c.entry_id + '-' + c.line_index)
  );

  const mapLine = (r) => ({
    date: r.date,
    type: 'Journal',
    ref_no: r.entry_num,
    payee: r.memo || '',
    amount: round2((r.debit || 0) - (r.credit || 0)),
  });

  const clearedLines = lineRows.filter(r => clearedSet.has(r.entry_id + '-' + r.line_index)).map(mapLine);
  const paymentsCleared = clearedLines.filter(l => l.amount < 0).sort((a,b)=> (a.date<b.date?-1:a.date>b.date?1:b.amount-a.amount));
  const depositsCleared = clearedLines.filter(l => l.amount > 0).sort((a,b)=> (a.date<b.date?-1:a.date>b.date?1:a.amount-b.amount));

  // Uncleared = lines on/before statement date not cleared in any reconciliation,
  // and lines dated after the statement date (cleared or not), mirroring QBO's
  // "uncleared transactions after <date>" register reconciliation.
  const everCleared = new Set(
    db.prepare('SELECT entry_id, line_index FROM cleared_items WHERE entity_id=? AND account_code=?')
      .all(eid, rec.account_code).map(c => c.entry_id + '-' + c.line_index)
  );
  const afterDate = lineRows.filter(r => r.date > rec.statement_date);
  const clearedAfter = afterDate.filter(r => everCleared.has(r.entry_id + '-' + r.line_index)).map(mapLine);
  const unclearedAfter = afterDate.filter(r => !everCleared.has(r.entry_id + '-' + r.line_index)).map(mapLine);
  const unclearedThrough = lineRows
    .filter(r => r.date <= rec.statement_date && !everCleared.has(r.entry_id + '-' + r.line_index))
    .map(mapLine);

  const sum = (arr) => round2(arr.reduce((s, l) => s + l.amount, 0));
  const paymentsTotal = sum(paymentsCleared);
  const depositsTotal = sum(depositsCleared);
  const endingBalance = round2(rec.statement_balance);
  const beginningBalance = round2(endingBalance - depositsTotal - paymentsTotal);
  const registerAtStmt = round2(rec.book_balance);
  const clearedAfterTotal = sum(clearedAfter);
  const unclearedAfterTotal = sum(unclearedAfter);
  const registerAsOfReport = round2(registerAtStmt + clearedAfterTotal + unclearedAfterTotal);

  res.json({
    entity_name: entity ? entity.name : '',
    account_code: rec.account_code,
    account_name: acct ? acct.name : '',
    statement_date: rec.statement_date,
    reconciled_on: rec.completed_at,
    reconciled_by: rec.completed_by,
    summary: {
      beginning_balance: beginningBalance,
      payments_count: paymentsCleared.length,
      payments_total: paymentsTotal,
      deposits_count: depositsCleared.length,
      deposits_total: depositsTotal,
      ending_balance: endingBalance,
      register_at_statement_date: registerAtStmt,
      cleared_after_count: clearedAfter.length,
      cleared_after_total: clearedAfterTotal,
      uncleared_after_count: unclearedAfter.length,
      uncleared_after_total: unclearedAfterTotal,
      uncleared_through_count: unclearedThrough.length,
      uncleared_through_total: sum(unclearedThrough),
      register_as_of_report: registerAsOfReport,
    },
    payments_cleared: paymentsCleared,
    deposits_cleared: depositsCleared,
    uncleared_through: unclearedThrough,
    cleared_after: clearedAfter,
    uncleared_after: unclearedAfter,
  });
});

// ═══ Summary ═══
app.get('/api/summary', auth, (req, res) => {
  const ids = listAccessibleEntityIds(req.user.id, req.user.role);
  let entities;
  if (ids === null) entities = db.prepare('SELECT * FROM entities ORDER BY code').all();
  else if (ids.length === 0) entities = [];
  else {
    const placeholders = ids.map(() => '?').join(',');
    entities = db.prepare('SELECT * FROM entities WHERE id IN (' + placeholders + ') ORDER BY code').all(...ids);
  }
  res.json(entities.map(e => {
    const rows = db.prepare(`SELECT a.type, SUM(jl.debit) as td, SUM(jl.credit) as tc FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? GROUP BY a.type`).all(e.id);
    const bt={}; rows.forEach(r=>{const isDr=r.type==='Asset'||r.type==='Expense'; bt[r.type]=isDr?(r.td-r.tc):(r.tc-r.td);});
    return { ...e, assets:bt.Asset||0, liabilities:bt.Liability||0, revenue:bt.Revenue||0, expenses:bt.Expense||0, net_income:(bt.Revenue||0)-(bt.Expense||0), entry_count: db.prepare('SELECT COUNT(*) as c FROM journal_entries WHERE entity_id=?').get(e.id).c };
  }));
});

// === Turnkey Rail integration routes ===

// All routes use API key auth, NOT JWT.
const turnkeyAuth = turnkey.apiKeyAuth(db);

// Health check (no auth — useful for Turnkey to verify connectivity)
app.get('/api/turnkey/health', (req, res) => {
  res.json({ status: 'ok', integration: 'turnkey-rail', timestamp: new Date().toISOString() });
});

// === Turnkey integration config (admin via JWT) ===
//
// Sets the company entity (the single "Turnkey Rail" entity that holds all
// project activity). Must be set before any project linking or sync events.
app.get('/api/admin/turnkey/config', auth, requireRole('Admin'), (req, res) => {
  const row = db.prepare('SELECT * FROM turnkey_config WHERE id = 1').get();
  res.json(row || { id: 1, enabled: 0, default_entity_id: null });
});

app.put('/api/admin/turnkey/config', auth, requireRole('Admin'), (req, res) => {
  const enabled = req.body.enabled ? 1 : 0;
  const entityId = req.body.default_entity_id != null ? Number(req.body.default_entity_id) : null;
  // Validate the entity exists
  if (entityId != null) {
    const ent = db.prepare('SELECT id FROM entities WHERE id = ?').get(entityId);
    if (!ent) return res.status(400).json({ error: 'default_entity_id refers to a non-existent entity' });
  }
  const now = new Date().toISOString();
  db.prepare(
    'INSERT INTO turnkey_config (id, enabled, default_entity_id, updated_by, updated_at) ' +
    'VALUES (1, ?, ?, ?, ?) ' +
    'ON CONFLICT(id) DO UPDATE SET enabled = excluded.enabled, default_entity_id = excluded.default_entity_id, ' +
    '  updated_by = excluded.updated_by, updated_at = excluded.updated_at'
  ).run(enabled, entityId, req.user.email, now);
  // Seed POC chart of accounts on the configured entity (idempotent)
  if (entityId != null) {
    const added = turnkey.seedPOCAccountsIfMissing(db, entityId);
    return res.json({ ok: true, enabled, default_entity_id: entityId, poc_accounts_added: added });
  }
  res.json({ ok: true, enabled, default_entity_id: entityId });
});

// === WIP Schedule (Job Schedule) endpoint ===
// Returns JSON. as_of query param defaults to today.
app.get('/api/turnkey/wip-schedule', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
    ? req.query.as_of
    : new Date().toISOString().slice(0, 10);
  try {
    const schedule = turnkey.computeWipSchedule(db, asOf);
    res.json(schedule);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// List linked Turnkey projects (for the manual journal-entry project tagger).
app.get('/api/admin/turnkey/projects', auth, requireRole('Admin','Accountant'), (req, res) => {
  try {
    const rows = db.prepare('SELECT turnkey_project_id, project_code, project_name FROM turnkey_project_map ORDER BY project_code').all();
    res.json(rows);
  } catch (e) {
    res.json([]); // table may not exist if integration unconfigured
  }
});

// Admin/UI WIP schedule (JWT auth) — same data as the API-key route, for the in-app report.
app.get('/api/admin/turnkey/wip-schedule', auth, requireRole('Admin'), (req, res) => {
  const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
    ? req.query.as_of
    : new Date().toISOString().slice(0, 10);
  try {
    const schedule = turnkey.computeWipSchedule(db, asOf);
    res.json(schedule);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Export the WIP schedule as Excel (.xlsx). Uses the existing 'xlsx' lib.
app.get('/api/turnkey/wip-schedule.xlsx', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
    ? req.query.as_of
    : new Date().toISOString().slice(0, 10);
  try {
    const schedule = turnkey.computeWipSchedule(db, asOf);
    const header = [
      'Job #', 'Job Name', 'Contract Amount', 'Revised Contract',
      'Costs to Date', 'Est Cost to Complete', 'Est Total Cost',
      'Est Gross Profit', '% Complete', 'Earned Revenue',
      'Billed to Date', 'Over/(Under) Billing'
    ];
    const dataRows = schedule.rows.map(r => [
      r.project_code || r.turnkey_project_id, r.project_name || '',
      r.contract_amount, r.revised_contract,
      r.costs_to_date, r.estimated_cost_to_complete, r.estimated_total_cost,
      r.estimated_gross_profit, r.percent_complete / 100, // store as fraction; format will render %
      r.earned_revenue, r.billed_to_date, r.over_under_billing,
    ]);
    const t = schedule.total;
    const totalRow = [
      'TOTAL', '',
      t.contract_amount, t.revised_contract,
      t.costs_to_date, t.estimated_cost_to_complete, t.estimated_total_cost,
      t.estimated_gross_profit, '',
      t.earned_revenue, t.billed_to_date, t.over_under_billing,
    ];
    const aoa = [
      ['Turnkey Rail — WIP Schedule'],
      ['As of:', asOf],
      [],
      header,
      ...dataRows,
      totalRow,
    ];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Column widths
    ws['!cols'] = [
      { wch: 12 }, { wch: 28 }, { wch: 16 }, { wch: 16 },
      { wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 16 },
      { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 20 },
    ];
    // Number formats for numeric cols
    const moneyFmt = '#,##0.00;(#,##0.00)';
    const pctFmt = '0.0%';
    const numericCols = [2,3,4,5,6,7,9,10,11];
    const headerRowIdx = 3; // 0-based, '$Job #' is row 4 in spreadsheet (after title+as_of+blank)
    const firstDataRow = headerRowIdx + 1;
    for (let i = 0; i < dataRows.length + 1; i++) { // +1 for total row
      for (const c of numericCols) {
        const addr = XLSX.utils.encode_cell({ r: firstDataRow + i, c });
        if (ws[addr]) ws[addr].z = moneyFmt;
      }
      // % column (index 8) — only data rows, not total
      if (i < dataRows.length) {
        const addr = XLSX.utils.encode_cell({ r: firstDataRow + i, c: 8 });
        if (ws[addr]) ws[addr].z = pctFmt;
      }
    }
    // Title cell — bold-ish via merge
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }];
    XLSX.utils.book_append_sheet(wb, ws, 'WIP Schedule');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="WIP_Schedule_' + asOf + '.xlsx"');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// === API key management (these use JWT/Admin, not API key) ===

// List API keys (no raw key visible, only metadata)
app.get('/api/admin/api-keys', auth, requireRole('Admin'), (req, res) => {
  const rows = db.prepare(
    'SELECT id, key_prefix, name, scopes, last_used_at, created_by, created_at, revoked_at FROM api_keys ORDER BY id DESC'
  ).all();
  res.json(rows);
});

// Create a new API key. Returns the raw key ONCE — admin must save it.
app.post('/api/admin/api-keys', auth, requireRole('Admin'), (req, res) => {
  const name = (req.body.name || '').trim();
  const scopes = (req.body.scopes || 'turnkey:sync').trim();
  if (!name) return res.status(400).json({ error: 'name required' });
  const rawKey = turnkey.generateApiKey();
  const hash = turnkey.hashApiKey(rawKey);
  const prefix = turnkey.apiKeyPrefix(rawKey);
  const now = new Date().toISOString();
  const result = db.prepare(
    'INSERT INTO api_keys (key_hash, key_prefix, name, scopes, created_by, created_at) VALUES (?, ?, ?, ?, ?, ?)'
  ).run(hash, prefix, name, scopes, req.user.email, now);
  res.json({ id: result.lastInsertRowid, raw_key: rawKey, key_prefix: prefix, name, scopes,
             warning: 'Save this key now. It will never be shown again.' });
});

// Revoke an API key
app.post('/api/admin/api-keys/:id/revoke', auth, requireRole('Admin'), (req, res) => {
  db.prepare('UPDATE api_keys SET revoked_at = ? WHERE id = ?').run(new Date().toISOString(), req.params.id);
  res.json({ success: true });
});

// === Project linking (Turnkey calls these with its API key) ===

// Register a Turnkey project as a job dimension on the company entity.
// Body: { turnkey_project_id, project_code, project_name,
//         contract_amount?, total_estimated_costs? }
// Update-on-conflict so this can also be called to refresh contract/estimate.
app.post('/api/turnkey/projects/link', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  try {
    const { turnkey_project_id, project_code, project_name,
            contract_amount, total_estimated_costs } = req.body;
    if (!turnkey_project_id || !project_code || !project_name) {
      return res.status(400).json({ error: 'turnkey_project_id, project_code, project_name required' });
    }
    const map = turnkey.linkProject(db, {
      turnkey_project_id, project_code, project_name,
      contract_amount, total_estimated_costs,
    });
    res.json({ ok: true, project_map: map });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Get project mapping (for Turnkey to verify the link exists)
app.get('/api/turnkey/projects/:id', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(req.params.id);
  if (!map) return res.status(404).json({ error: 'Project not linked' });
  res.json(map);
});

// === Sync event endpoints ===
// All accept JSON payload; all return { ok, cl_entry_id, idempotent } on success.

function syncRoute(syncFn) {
  return (req, res) => {
    try {
      const result = syncFn(db, req.body || {});
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  };
}

app.post('/api/turnkey/sync/sub-payapp-approved',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncSubPayAppApproved));

app.post('/api/turnkey/sync/sub-payapp-paid',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncSubPayAppPaid));

app.post('/api/turnkey/sync/owner-payapp-issued',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncOwnerPayAppIssued));

app.post('/api/turnkey/sync/owner-payment-received',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncOwnerPaymentReceived));

app.post('/api/turnkey/sync/month-end-poc',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncMonthEndPOC));

// View sync log for a project (last 50 events)
app.get('/api/turnkey/sync-log/:turnkey_project_id', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const map = db.prepare('SELECT cl_entity_id FROM turnkey_project_map WHERE turnkey_project_id = ?').get(req.params.turnkey_project_id);
  if (!map) return res.status(404).json({ error: 'Project not linked' });
  const rows = db.prepare(
    'SELECT id, sync_type, turnkey_id, cl_entry_id, status, message, created_at FROM turnkey_sync_log ' +
    'WHERE cl_entity_id = ? ORDER BY id DESC LIMIT 50'
  ).all(map.cl_entity_id);
  res.json(rows);
});

// === Bill.com integration routes ===
app.get('/api/billcom/config/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT entity_id, environment, api_base_url, username, password_enc, org_id, dev_key_enc, default_ap_account, default_cash_account, default_clearing_account, sync_cutoff_date, last_tested_at, last_test_status, last_test_message, updated_by, updated_at FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
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
    default_clearing_account: row.default_clearing_account,
    sync_cutoff_date: row.sync_cutoff_date,
    last_tested_at: row.last_tested_at,
    last_test_status: row.last_test_status,
    last_test_message: row.last_test_message,
    updated_by: row.updated_by,
    updated_at: row.updated_at
  });
});

app.put('/api/billcom/config/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), (req, res) => {
  const { environment, username, password, org_id, dev_key, default_ap_account, default_cash_account, default_clearing_account, sync_cutoff_date } = req.body || {};
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
    db.prepare('UPDATE billcom_config SET environment=?, api_base_url=?, username=?, password_enc=?, org_id=?, dev_key_enc=?, default_ap_account=?, default_cash_account=?, default_clearing_account=?, sync_cutoff_date=?, updated_by=?, updated_at=? WHERE entity_id=?')
      .run(environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, default_clearing_account || null, sync_cutoff_date || null, updater, now, req.params.entity_id);
  } else {
    db.prepare('INSERT INTO billcom_config (entity_id, environment, api_base_url, username, password_enc, org_id, dev_key_enc, default_ap_account, default_cash_account, default_clearing_account, sync_cutoff_date, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)')
      .run(req.params.entity_id, environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, default_clearing_account || null, sync_cutoff_date || null, updater, now);
  }
  res.json({ success: true });
});

app.delete('/api/billcom/config/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM billcom_config WHERE entity_id = ?').run(req.params.entity_id);
  res.json({ success: true });
});

app.post('/api/billcom/config/:entity_id/test', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), async (req, res) => {
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

app.get('/api/billcom/accounts/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
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

app.get('/api/billcom/mappings/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  const rows = db.prepare('SELECT id, billcom_account_id, billcom_account_name, cl_account_code, created_at FROM billcom_account_map WHERE entity_id = ? ORDER BY id').all(req.params.entity_id);
  res.json({ mappings: rows });
});

// Map-only auto-populate of the GL account map: match existing Bill.com accounts
// to CL account codes by account number (name fallback) and write matches. Never
// creates anything in Bill.com (unlike push-coa). Unmatched CL codes are reported.
app.post('/api/billcom/mappings/:entity_id/auto', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(eid);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });
  let session, devKey, bcAccounts;
  try {
    const pw = billcomDecrypt(cfg.password_enc);
    devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password: pw, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
    bcAccounts = await billcomListAccounts({ sessionId: session.sessionId, devKey, baseUrl: cfg.api_base_url });
  } catch (e) { return res.status(502).json({ error: 'login or account list failed: ' + e.message }); }

  const norm = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const bcByNum = new Map();
  const bcByName = new Map();
  for (const a of bcAccounts) {
    const num = a && (a.accountNumber || a.number || (a.account && a.account.accountNumber));
    if (num) bcByNum.set(String(num), a);
    if (a && a.name) bcByName.set(norm(a.name), a);
  }
  const clAccounts = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ?').all(eid);
  const now = new Date().toISOString();
  const result = { matched: [], unmatched: [] };
  try {
    const tx = db.transaction(() => {
      db.prepare('DELETE FROM billcom_account_map WHERE entity_id = ?').run(eid);
      const ins = db.prepare('INSERT INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)');
      for (const c of clAccounts) {
        const bc = bcByNum.get(String(c.code)) || bcByName.get(norm(c.name));
        if (bc && bc.id) { ins.run(eid, String(bc.id), bc.name || null, String(c.code), now); result.matched.push({ code: c.code, name: c.name, billcom_id: bc.id }); }
        else result.unmatched.push({ code: c.code, name: c.name });
      }
    });
    tx();
  } catch (e) { return res.status(500).json({ error: e.message }); }
  res.json({ matched: result.matched.length, unmatched: result.unmatched.length, unmatched_codes: result.unmatched, billcom_account_count: bcAccounts.length });
});

app.put('/api/billcom/mappings/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
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

// Bill.com dimension maps (class=investor via accountingClassId, location=deal
// via jobId). GET returns current maps; POST auto-populates by name-matching
// Bill.com classes/jobs to CloudLedger's class/location dimensions, writing only
// confident matches. Unmatched (incl. workflow-status classes) are reported but
// not written, so they sync as null by design. Matches are editable via PUT.
app.get('/api/billcom/dimension-maps/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const classes = db.prepare('SELECT billcom_class_id, billcom_class_name, cl_class_id FROM billcom_class_map WHERE entity_id = ? ORDER BY id').all(eid);
  const locations = db.prepare('SELECT billcom_job_id, billcom_job_name, cl_location_id FROM billcom_location_map WHERE entity_id = ? ORDER BY id').all(eid);
  res.json({ classes, locations });
});

app.post('/api/billcom/dimension-maps/:entity_id/auto', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(eid);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });
  let session, devKey;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) { return res.status(502).json({ error: 'Bill.com login failed: ' + e.message }); }
  const args = { sessionId: session.sessionId, devKey, baseUrl: cfg.api_base_url };

  let bcClasses, bcJobs;
  try { bcClasses = await billcomListClassification({ ...args, resource: 'accounting-classes' }); }
  catch (e) { return res.status(502).json({ error: 'fetch classes failed: ' + e.message }); }
  try { bcJobs = await billcomListClassification({ ...args, resource: 'jobs' }); }
  catch (e) { return res.status(502).json({ error: 'fetch jobs failed: ' + e.message }); }

  // CL dimensions for this entity (class = investor, location = deal).
  const clClasses = db.prepare('SELECT id, name FROM dim_classes WHERE entity_id = ?').all(eid);
  const clLocs = db.prepare('SELECT id, name FROM dim_locations WHERE entity_id = ?').all(eid);
  const norm = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const clClassByName = new Map(clClasses.map(c => [norm(c.name), c.id]));
  const clLocByName = new Map(clLocs.map(l => [norm(l.name), l.id]));
  const nameOf = (o) => o.name || o.shortName || o.description || '';
  const idOf = (o) => String(o.id || '');

  const now = new Date().toISOString();
  const result = { classes: { matched: [], unmatched: [] }, locations: { matched: [], unmatched: [] } };

  const tx = db.transaction(() => {
    // Workflow-status classes are not investors. Skip them even when CL happens
    // to have a same-named class, so re-running /auto never re-pollutes the map.
    const WORKFLOW_CLASS_NAMES = new Set(['pay', 'hold', 'already paid', 'paid', 'on hold']);
    db.prepare('DELETE FROM billcom_class_map WHERE entity_id = ?').run(eid);
    const insC = db.prepare('INSERT INTO billcom_class_map (entity_id, billcom_class_id, billcom_class_name, cl_class_id, created_at) VALUES (?,?,?,?,?)');
    for (const c of bcClasses) {
      const nm = nameOf(c);
      if (WORKFLOW_CLASS_NAMES.has(norm(nm))) { result.classes.unmatched.push({ billcom_class_id: idOf(c), name: nm, skipped: 'workflow status' }); continue; }
      const clId = clClassByName.get(norm(nm));
      if (clId) { insC.run(eid, idOf(c), nm, clId, now); result.classes.matched.push({ billcom: nm, cl_class_id: clId }); }
      else result.classes.unmatched.push({ billcom_class_id: idOf(c), name: nm });
    }
    db.prepare('DELETE FROM billcom_location_map WHERE entity_id = ?').run(eid);
    const insL = db.prepare('INSERT INTO billcom_location_map (entity_id, billcom_job_id, billcom_job_name, cl_location_id, created_at) VALUES (?,?,?,?,?)');
    for (const j of bcJobs) {
      const nm = nameOf(j); const clId = clLocByName.get(norm(nm));
      if (clId) { insL.run(eid, idOf(j), nm, clId, now); result.locations.matched.push({ billcom: nm, cl_location_id: clId }); }
      else result.locations.unmatched.push({ billcom_job_id: idOf(j), name: nm });
    }
  });
  try { tx(); } catch (e) { return res.status(500).json({ error: e.message }); }
  res.json(result);
});

// Upsert manual dimension-map rows without disturbing existing ones. Body:
// { classes?: [{billcom_class_id, billcom_class_name?, cl_class_id}],
//   locations?: [{billcom_job_id, billcom_job_name?, cl_location_id}] }.
// Used to add name-mismatch matches the auto step couldn't make (e.g. Bill.com
// "Buna" -> CL "CLR Buna Property Owner LLC"). A null cl id deletes the mapping.
app.put('/api/billcom/dimension-maps/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const classes = Array.isArray(req.body && req.body.classes) ? req.body.classes : [];
  const locations = Array.isArray(req.body && req.body.locations) ? req.body.locations : [];
  const now = new Date().toISOString();
  try {
    const tx = db.transaction(() => {
      const upC = db.prepare('INSERT INTO billcom_class_map (entity_id, billcom_class_id, billcom_class_name, cl_class_id, created_at) VALUES (?,?,?,?,?) ON CONFLICT(entity_id, billcom_class_id) DO UPDATE SET cl_class_id=excluded.cl_class_id, billcom_class_name=excluded.billcom_class_name');
      const delC = db.prepare('DELETE FROM billcom_class_map WHERE entity_id = ? AND billcom_class_id = ?');
      for (const c of classes) {
        if (!c.billcom_class_id) continue;
        if (c.cl_class_id == null) delC.run(eid, String(c.billcom_class_id));
        else upC.run(eid, String(c.billcom_class_id), c.billcom_class_name || null, parseInt(c.cl_class_id), now);
      }
      const upL = db.prepare('INSERT INTO billcom_location_map (entity_id, billcom_job_id, billcom_job_name, cl_location_id, created_at) VALUES (?,?,?,?,?) ON CONFLICT(entity_id, billcom_job_id) DO UPDATE SET cl_location_id=excluded.cl_location_id, billcom_job_name=excluded.billcom_job_name');
      const delL = db.prepare('DELETE FROM billcom_location_map WHERE entity_id = ? AND billcom_job_id = ?');
      for (const l of locations) {
        if (!l.billcom_job_id) continue;
        if (l.cl_location_id == null) delL.run(eid, String(l.billcom_job_id));
        else upL.run(eid, String(l.billcom_job_id), l.billcom_job_name || null, parseInt(l.cl_location_id), now);
      }
    });
    tx();
    const classCount = db.prepare('SELECT COUNT(*) c FROM billcom_class_map WHERE entity_id = ?').get(eid).c;
    const locCount = db.prepare('SELECT COUNT(*) c FROM billcom_location_map WHERE entity_id = ?').get(eid).c;
    res.json({ ok: true, class_map_rows: classCount, location_map_rows: locCount });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Phase 5: Push CloudLedger COA to Bill.com and auto-create mappings.
app.post('/api/billcom/push-coa/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured' });

  const body = req.body || {};
  let rows;
  if (Array.isArray(body.codes) && body.codes.length > 0) {
    const placeholders = body.codes.map(function(){ return '?'; }).join(',');
    rows = db.prepare('SELECT code, name, type, subtype FROM accounts WHERE entity_id = ? AND code IN (' + placeholders + ')').all(entityId, ...body.codes);
  } else if (body.all) {
    // Every account regardless of type (asset/liability/income/equity/expense).
    // mapType() below maps each CL type to the right Bill.com account type.
    rows = db.prepare("SELECT code, name, type, subtype FROM accounts WHERE entity_id = ? ORDER BY code").all(entityId);
  } else if (body.all_expenses) {
    rows = db.prepare("SELECT code, name, type, subtype FROM accounts WHERE entity_id = ? AND type = 'Expense' ORDER BY code").all(entityId);
  } else {
    return res.status(400).json({ error: 'Provide codes:[...], all:true, or all_expenses:true' });
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
app.get('/api/billcom/sync-log/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  const limit = Math.min(parseInt(req.query.limit) || 50, 500);
  const rows = db.prepare(
    'SELECT id, sync_type, billcom_id, cl_entry_id, status, message, created_at FROM billcom_sync_log WHERE entity_id = ? ORDER BY id DESC LIMIT ?'
  ).all(entityId, limit);
  res.json({ logs: rows });
});

app.post('/api/billcom/sync/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
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
  // Dimension maps: only mapped Bill.com class/job ids carry a CL class/location
  // onto the synced line; unmapped ids (incl. workflow-status classes) -> null.
  const classMap = new Map(db.prepare('SELECT billcom_class_id, cl_class_id FROM billcom_class_map WHERE entity_id = ?').all(entityId).map(r => [String(r.billcom_class_id), r.cl_class_id]));
  const locMap = new Map(db.prepare('SELECT billcom_job_id, cl_location_id FROM billcom_location_map WHERE entity_id = ?').all(entityId).map(r => [String(r.billcom_job_id), r.cl_location_id]));

  let session;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    const devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) {
    return res.status(502).json({ error: 'Bill.com login failed: ' + e.message });
  }
  const listArgs = { sessionId: session.sessionId, devKey: billcomDecrypt(cfg.dev_key_enc), baseUrl: cfg.api_base_url };

  // Bounded sync: process at most maxBills per invocation so the request always
  // returns well under Railway's gateway ceiling. Dedup via billcom_sync_log makes
  // repeated runs safe + incremental — re-run until processed === 0. Caller may
  // override with body.max_bills (clamped 1..100).
  const maxBills = Math.max(1, Math.min(100, parseInt((req.body && req.body.max_bills) || 25)));
  const deadline = Date.now() + 230000; // stop starting new work past ~3.8m, safely under the 300s gateway cap

  // Opening-balance cutoff: all balances before this date were booked via the
  // opening journal entry (from the GL detail report), so importing bills/payments
  // dated before it would double-count. Skip anything earlier. Configurable via
  // config.sync_cutoff_date or body.cutoff_date; defaults to 2026-01-01.
  const cutoffDate = String((req.body && req.body.cutoff_date) || cfg.sync_cutoff_date || '2026-01-01');

  let bills, payments;
  try {
    bills = await billcomListBills({ ...listArgs, maxItems: 500 });
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch bills: ' + e.message });
  }
  try {
    payments = await billcomListPayments({ ...listArgs, maxItems: 500 });
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

  let billsProcessed = 0;
  result.bills.budget_reached = false;
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
    // Cheap cutoff check on the list object's date (avoids an expensive detail
    // fetch for the many pre-cutoff bills). A detail-based check below is the
    // safety net for when the list omits a date.
    const listDate = pick(bill, 'invoiceDate', 'invoice_date', 'dueDate') || pick(pick(bill, 'invoice') || {}, 'invoiceDate', 'invoice_date');
    if (listDate && String(listDate) < cutoffDate) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'before cutoff ' + cutoffDate + ' (date ' + listDate + ')' });
      continue;
    }
    // Bounded work: stop starting new bills once the per-run budget or time
    // deadline is hit. Remaining bills are picked up on the next sync run.
    if (billsProcessed >= maxBills || Date.now() > deadline) {
      result.bills.budget_reached = true;
      break;
    }
    billsProcessed++;

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

    // Skip bills dated before the opening-balance cutoff (already in the opening JE).
    if (String(invoiceDate) < cutoffDate) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'before cutoff ' + cutoffDate + ' (date ' + invoiceDate + ')' });
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
      debitLines.push({
        account_code: mapping.cl_account_code, debit: amt, credit: 0,
        class_id: classMap.get(String(pick(cls, 'accountingClassId', 'classId') || '')) || null,
        location_id: locMap.get(String(pick(cls, 'jobId', 'locationId') || '')) || null,
      });
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

    const lines = [...debitLines, { account_code: apAccount, debit: 0, credit: totalDr, class_id: null, location_id: null }];
    const memo = 'Bill.com bill #' + billNumber;

    try {
      const insertedId = db.transaction(() => {
        const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
        const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
          .run(entityId, num, invoiceDate, memo, actor);
        for (const l of lines) {
          db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, class_id, location_id) VALUES (?,?,?,?,?,?)')
            .run(r.lastInsertRowid, l.account_code, l.debit, l.credit, l.class_id || null, l.location_id || null);
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

  result.payments.budget_reached = false;
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
    if (Date.now() > deadline) { result.payments.budget_reached = true; break; }

    const processDate = pick(pay, 'processDate', 'process_date', 'paymentDate', 'payment_date');
    const amount = Number(pick(pay, 'amount', 'paymentAmount', 'totalAmount') || 0);
    const payNumber = pick(pay, 'paymentNumber', 'payment_number', 'referenceNumber') || payId;

    if (!processDate || amount <= 0) {
      result.payments.errors++;
      logSync.run(entityId, 'payment', payId, null, 'error', 'missing processDate or zero amount', now);
      result.payments.details.push({ id: payId, status: 'error', reason: 'missing processDate or zero amount' });
      continue;
    }

    // Skip payments dated before the opening-balance cutoff (already in the opening JE).
    if (String(processDate) < cutoffDate) {
      result.payments.skipped++;
      result.payments.details.push({ id: payId, status: 'skip', reason: 'before cutoff ' + cutoffDate + ' (date ' + processDate + ')' });
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

// ───────────────────────────────────────────────────────────────────────────
// Payment Reconcile (Phase 7, Option 2) — QBO-style two-leg payment sync built
// on the READ path only (the MFA/BDC_1361 wall blocks writing/initiating
// payments via the API, but reading payment status works fine).
//
// Bill.com data model (verified against entity 40 production):
//   • /payments rows carry: billPayments[] = [{ billId, amount }] (one payment
//     can settle several bills), processDate (the date Bill.com pulls the bank),
//     status (PAID / VOID), voidInfo[], cancelRequestSubmitted.
//   • A bill's invoiceNumber is null at the top level (nested under invoice),
//     so we link payments to CL bill JEs by billId via billcom_sync_log, NOT by
//     invoice number.
//
// Two legs, mirroring how Bill.com posts into QuickBooks Online:
//   Leg 1 (per bill relieved): Dr AP (202000) / Cr Clearing (1072) — relieves
//          the open payable and ages the bill out of the AP-aging report.
//   Leg 2 (per processDate, lump sum): Dr Clearing (1072) / Cr Cash (100200) —
//          one funds-transfer entry per process date, matching the single
//          batch ACH withdrawal on the bank statement.
//
// Only truly disbursed payments relieve: status === 'PAID', not voided, not
// cancel-requested, and processDate <= as_of (a future-dated PAID is scheduled,
// not yet pulled). Payments whose bill has no synced CL JE (e.g. pre-cutover
// bills already in the opening balance) are skipped — relieving them would
// double-count against the opening JE. Idempotent + incremental via
// billcom_sync_log (sync_type 'payment' and 'funds_transfer'); safe to re-run.
// ───────────────────────────────────────────────────────────────────────────
app.post('/api/billcom/payment-reconcile/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });

  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });

  const apAccount = cfg.default_ap_account;
  const clearingAccount = cfg.default_clearing_account;
  const cashAccount = cfg.default_cash_account;
  if (!apAccount) return res.status(400).json({ error: 'default_ap_account not set in Bill.com config' });
  if (!clearingAccount) return res.status(400).json({ error: 'default_clearing_account not set in Bill.com config' });
  if (!cashAccount) return res.status(400).json({ error: 'default_cash_account not set in Bill.com config' });

  for (const [label, code] of [['AP', apAccount], ['clearing', clearingAccount], ['cash', cashAccount]]) {
    const ok = db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, code);
    if (!ok) return res.status(400).json({ error: label + ' account ' + code + ' does not exist on entity' });
  }

  // dry_run=true previews the JEs that WOULD be created without writing them.
  const dryRun = !!(req.body && (req.body.dry_run === true || req.body.dry_run === 'true'));
  const asOf = String((req.body && req.body.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.body.as_of)) ? req.body.as_of : new Date().toISOString().slice(0, 10));
  // Anything dated before the opening-balance cutoff is already in the opening JE.
  const cutoffDate = String((req.body && req.body.cutoff_date) || cfg.sync_cutoff_date || '2026-01-01');

  let session;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    const devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) {
    return res.status(502).json({ error: 'Bill.com login failed: ' + e.message });
  }
  const listArgs = { sessionId: session.sessionId, devKey: billcomDecrypt(cfg.dev_key_enc), baseUrl: cfg.api_base_url };
  const pick = (o, ...ks) => { for (const k of ks) if (o && o[k] != null) return o[k]; return null; };

  let payments;
  try {
    payments = await billcomListPayments({ ...listArgs, maxItems: 5000 });
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch payments: ' + e.message });
  }

  // billId -> CL bill JE id (only bills we actually synced into CL can be relieved).
  const billEntryByBillcomId = new Map();
  for (const r of db.prepare(
    "SELECT billcom_id, cl_entry_id FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill' AND status = 'success' AND cl_entry_id IS NOT NULL"
  ).all(entityId)) {
    if (r.billcom_id != null) billEntryByBillcomId.set(String(r.billcom_id), r.cl_entry_id);
  }

  const alreadySynced = db.prepare(
    "SELECT 1 FROM billcom_sync_log WHERE entity_id = ? AND sync_type = ? AND billcom_id = ? AND status = 'success' LIMIT 1"
  );
  const logSync = db.prepare(
    'INSERT INTO billcom_sync_log (entity_id, sync_type, billcom_id, cl_entry_id, status, message, created_at) VALUES (?,?,?,?,?,?,?)'
  );
  const now = new Date().toISOString();
  const actor = req.user.name || req.user.email || 'system';

  // A payment is "disbursed" (relieves AP) only when truly paid out, not voided,
  // not pending cancellation, and not future-dated relative to as_of.
  const isDisbursed = (pay) => {
    const s = String(pick(pay, 'status', 'paymentStatus') || '').toUpperCase();
    if (s !== 'PAID') return false;
    const voids = pick(pay, 'voidInfo');
    if (Array.isArray(voids) && voids.length > 0) return false;
    if (pick(pay, 'cancelRequestSubmitted') === true) return false;
    const pd = pick(pay, 'processDate', 'process_date', 'paymentDate');
    if (!pd) return false;
    if (String(pd) > asOf) return false;        // scheduled, not yet pulled
    if (String(pd) < cutoffDate) return false;  // already in opening balance
    return true;
  };

  const result = {
    dry_run: dryRun, as_of: asOf, cutoff_date: cutoffDate,
    accounts: { ap: apAccount, clearing: clearingAccount, cash: cashAccount },
    payments_fetched: payments.length,
    leg1: { relieved: 0, skipped: 0, errors: 0, amount: 0, details: [] },
    leg2: { transfers: 0, skipped: 0, errors: 0, amount: 0, details: [] },
  };

  // ── Leg 1: relieve each bill settled by a disbursed payment.
  //    Group disbursed amounts by processDate for Leg 2 lump sums.
  const transferByDate = new Map(); // processDate -> total disbursed amount (this run)
  const insertJE = (date, memo, lines) => {
    const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
    const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
      .run(entityId, num, date, memo, actor);
    for (const l of lines) {
      db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)')
        .run(r.lastInsertRowid, l.account_code, l.debit, l.credit);
    }
    return { id: r.lastInsertRowid, num };
  };

  for (const pay of payments) {
    const payId = String(pick(pay, 'id') || '');
    if (!payId) continue;
    const processDate = String(pick(pay, 'processDate', 'process_date', 'paymentDate') || '');
    const payNum = pick(pay, 'transactionNumber', 'confirmationNumber', 'paymentNumber') || payId;

    if (!isDisbursed(pay)) {
      result.leg1.skipped++;
      result.leg1.details.push({ id: payId, status: 'skip', reason: 'not disbursed/in-window', processDate, payStatus: pick(pay, 'status') });
      continue;
    }

    // Settle each bill this payment covers (billPayments[] is the authoritative
    // allocation; fall back to the top-level billId for single-bill payments).
    let allocations = Array.isArray(pick(pay, 'billPayments')) ? pick(pay, 'billPayments') : null;
    if (!allocations || !allocations.length) {
      const bid = pick(pay, 'billId');
      const amt = Number(pick(pay, 'amount', 'paymentAmount') || 0);
      allocations = bid ? [{ billId: bid, amount: amt }] : [];
    }

    for (const alloc of allocations) {
      const billId = String(pick(alloc, 'billId', 'bill_id') || '');
      const amount = Number(pick(alloc, 'amount') || 0);
      if (!billId || amount <= 0) { result.leg1.skipped++; continue; }

      // Dedup key: one relief per (payment, bill) allocation.
      const dedupId = payId + ':' + billId;
      if (alreadySynced.get(entityId, 'payment', dedupId)) {
        result.leg1.skipped++;
        result.leg1.details.push({ id: dedupId, status: 'skip', reason: 'already reconciled' });
        // Still counts toward the clearing balance already moved in a prior run's
        // Leg 2, so do NOT add to transferByDate here.
        continue;
      }

      const billEntryId = billEntryByBillcomId.get(billId);
      if (!billEntryId) {
        // Bill not synced into CL (pre-cutover/opening-balance bill) — skip to
        // avoid double-relieving against the opening JE.
        result.leg1.skipped++;
        result.leg1.details.push({ id: dedupId, status: 'skip', reason: 'bill not synced to CL (likely pre-cutover)', billId });
        continue;
      }

      const memo = 'Bill.com payment ' + payNum + ' — relieve bill ' + billId;
      const lines = [
        { account_code: apAccount, debit: amount, credit: 0 },
        { account_code: clearingAccount, debit: 0, credit: amount },
      ];

      if (dryRun) {
        result.leg1.relieved++; result.leg1.amount += amount;
        result.leg1.details.push({ id: dedupId, status: 'would_create', date: processDate, amount, billEntryId });
        transferByDate.set(processDate, (transferByDate.get(processDate) || 0) + amount);
      } else {
        try {
          const je = db.transaction(() => {
            const created = insertJE(processDate, memo, lines);
            logSync.run(entityId, 'payment', dedupId, created.id, 'success', 'relieved bill ' + billId + ' (JE #' + created.num + ')', now);
            return created;
          })();
          result.leg1.relieved++; result.leg1.amount += amount;
          result.leg1.details.push({ id: dedupId, status: 'created', cl_entry_id: je.id, date: processDate, amount });
          transferByDate.set(processDate, (transferByDate.get(processDate) || 0) + amount);
        } catch (e) {
          result.leg1.errors++;
          logSync.run(entityId, 'payment', dedupId, null, 'error', e.message, now);
          result.leg1.details.push({ id: dedupId, status: 'error', reason: e.message });
        }
      }
    }
  }

  // ── Leg 2: one funds-transfer JE per processDate that had NEW relief this run.
  //    Dedup key = 'ft:' + processDate. If a transfer already exists for that
  //    date (a prior run), we must ADD to it rather than skip — otherwise the
  //    clearing account would not fully drain. We post a top-up for the delta.
  for (const [pd, amount] of transferByDate.entries()) {
    if (amount <= 0.005) continue;
    const dedupId = 'ft:' + pd;
    const prior = db.prepare(
      "SELECT cl_entry_id FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'funds_transfer' AND billcom_id = ? AND status = 'success'"
    ).all(entityId, dedupId);
    // Sum what we've already transferred for this date so we only top up the delta.
    let priorTransferred = 0;
    for (const p of prior) {
      if (p.cl_entry_id) {
        const row = db.prepare('SELECT SUM(debit) AS d FROM journal_lines WHERE entry_id = ? AND account_code = ?').get(p.cl_entry_id, clearingAccount);
        priorTransferred += (row && row.d) || 0;
      }
    }
    const delta = amount; // amount here is only NEW relief from this run (prior runs' relief wasn't added to transferByDate)
    if (delta <= 0.005) { result.leg2.skipped++; continue; }

    const memo = 'Bill.com funds transfer — ' + pd + ' batch';
    const lines = [
      { account_code: clearingAccount, debit: delta, credit: 0 },
      { account_code: cashAccount, debit: 0, credit: delta },
    ];

    if (dryRun) {
      result.leg2.transfers++; result.leg2.amount += delta;
      result.leg2.details.push({ date: pd, status: 'would_create', amount: delta, prior_transferred: priorTransferred });
    } else {
      try {
        const je = db.transaction(() => {
          const created = insertJE(pd, memo, lines);
          logSync.run(entityId, 'funds_transfer', dedupId, created.id, 'success', 'funds transfer ' + pd + ' $' + delta.toFixed(2) + ' (JE #' + created.num + ')', now);
          return created;
        })();
        result.leg2.transfers++; result.leg2.amount += delta;
        result.leg2.details.push({ date: pd, status: 'created', cl_entry_id: je.id, amount: delta });
      } catch (e) {
        result.leg2.errors++;
        logSync.run(entityId, 'funds_transfer', dedupId, null, 'error', e.message, now);
        result.leg2.details.push({ date: pd, status: 'error', reason: e.message });
      }
    }
  }

  result.leg1.amount = Math.round(result.leg1.amount * 100) / 100;
  result.leg2.amount = Math.round(result.leg2.amount * 100) / 100;
  res.json(result);
});

// ───────────────────────────────────────────────────────────────────────────
// AP Aging Detail (Q5 / Weaver) — read open bills straight from Bill.com and
// bucket by days past due. Read-only: no JEs are written. Available to
// Accountant + Admin. The MFA/BDC_1361 block only affects payment *sync*;
// reading bills works. Buckets match Weaver's sample: current, 1-30, 31-60,
// 61-90, 91+ (relative to as_of, default today).
// ───────────────────────────────────────────────────────────────────────────
app.get('/api/billcom/ap-aging/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });

  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  const apAccount = (cfg && cfg.default_ap_account) ? String(cfg.default_ap_account) : '202000';
  const asOf = String((req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of)) ? req.query.as_of : new Date().toISOString().slice(0, 10));

  // ── 1. Pull all GL activity on the AP account through the as-of date. This is
  //    the authoritative AP record: credits = bills, debits = payments/relief.
  //    The report is built from here so it ALWAYS ties to the GL balance.
  const glLines = db.prepare(
    `SELECT jl.id AS line_id, je.id AS entry_id, je.entry_num, je.date, je.memo,
            jl.debit, jl.credit, jl.description
       FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id
      WHERE je.entity_id = ? AND jl.account_code = ? AND je.date <= ?
      ORDER BY je.date ASC, je.entry_num ASC, jl.id ASC`
  ).all(entityId, apAccount, asOf);

  const glBalance = glLines.reduce((s, l) => s + (l.credit || 0) - (l.debit || 0), 0);

  // ── 2. Which entries are Bill.com-synced bills? (vs. imported/manual GL entries)
  //    A 202000 credit whose entry is linked in billcom_sync_log is a synced
  //    invoice → aged Bill.com row. Everything else → GL column.
  const syncedEntryIds = new Set();
  const billcomIdByEntry = new Map();
  try {
    const rows = db.prepare(
      "SELECT cl_entry_id, billcom_id FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill' AND status = 'success' AND cl_entry_id IS NOT NULL"
    ).all(entityId);
    for (const r of rows) { syncedEntryIds.add(r.cl_entry_id); billcomIdByEntry.set(r.cl_entry_id, String(r.billcom_id)); }
  } catch (e) { /* sync log optional */ }

  // ── 3. FIFO net: apply debits (payments/relief) against the oldest open
  //    credits (bills), so what remains are the genuinely open items as of date.
  //    A payment may post before its bill is imported (interleaving in the GL),
  //    so any debit not fully absorbed by the current queue is CARRIED FORWARD
  //    and relieves later credits. Without this carry-forward the netting
  //    discards over-relief and overstates open AP — the open total must equal
  //    the GL balance (credits − debits) exactly.
  const openItems = []; // { line_id, entry_id, entry_num, date, memo, description, amount }
  let creditQueue = []; // FIFO of open credit slices
  let unappliedDebit = 0; // payment carried forward to relieve future bills
  for (const l of glLines) {
    if ((l.credit || 0) > 0.005) {
      let remaining = l.credit;
      if (unappliedDebit > 0.005) {
        const take = Math.min(remaining, unappliedDebit);
        remaining -= take; unappliedDebit -= take;
      }
      if (remaining > 0.005) creditQueue.push({ ...l, remaining });
    }
    if ((l.debit || 0) > 0.005) {
      let pay = l.debit;
      while (pay > 0.005 && creditQueue.length) {
        const head = creditQueue[0];
        const take = Math.min(head.remaining, pay);
        head.remaining -= take; pay -= take;
        if (head.remaining <= 0.005) creditQueue.shift();
      }
      if (pay > 0.005) unappliedDebit += pay; // carry forward to later bills
    }
  }
  for (const c of creditQueue) {
    if (c.remaining > 0.005) openItems.push({
      line_id: c.line_id, entry_id: c.entry_id, entry_num: c.entry_num,
      date: c.date, memo: c.memo || '', description: c.description || '', amount: c.remaining,
    });
  }

  // ── 4. Enrich Bill.com-synced items with vendor + due date from Bill.com.
  //    Only attempt the (slow) Bill.com pull if there ARE synced items to enrich.
  const hasSynced = openItems.some(it => syncedEntryIds.has(it.entry_id));
  const billByNum = new Map(); // invoiceNumber -> { vendor, dueDate }
  const vendorById = new Map();
  let billcomError = null;
  if (hasSynced && cfg) {
    try {
      const password = billcomDecrypt(cfg.password_enc);
      const devKey = billcomDecrypt(cfg.dev_key_enc);
      const session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
      const listArgs = { sessionId: session.sessionId, devKey, baseUrl: cfg.api_base_url };
      const pick = (o, ...ks) => { for (const k of ks) if (o && o[k] != null) return o[k]; return null; };
      let vendors = [];
      try { vendors = await billcomListVendors({ ...listArgs, maxItems: 5000 }); } catch (e) {}
      for (const v of vendors) { const id = String(pick(v, 'id') || ''); const n = pick(v, 'name', 'vendorName', 'companyName'); if (id && n) vendorById.set(id, n); }
      // dueDate-window paginator: the only reliable way to page Bill.com /bills
      const bills = await billcomListBillsWindowed({ ...listArgs, fromDate: '2024-01-01', toDate: asOf });
      for (const b of bills) {
        const num = pick(b, 'invoiceNumber', 'invoice_number') || pick(pick(b, 'invoice') || {}, 'invoiceNumber');
        if (!num) continue;
        const vid = String(pick(b, 'vendorId', 'vendor_id') || (pick(b, 'vendor') || {}).id || '');
        billByNum.set(String(num), { vendor: vendorById.get(vid) || pick(pick(b, 'vendor') || {}, 'name') || null, dueDate: pick(b, 'dueDate', 'due_date') || null });
      }
    } catch (e) { billcomError = e.message; }
  }

  // Invoice # is stored in the import memo as "... — <invoiceNum>".
  const invNumFromMemo = (memo) => { const m = String(memo || '').match(/—\s*(.+?)\s*$/); return m ? m[1].trim() : null; };

  // ── 5. Build buckets for Bill.com invoices; sum GL column for the rest.
  const buckets = ['current', 'd1_30', 'd31_60', 'd61_90', 'd91_plus'];
  const emptyBuckets = () => ({ current: 0, d1_30: 0, d31_60: 0, d61_90: 0, d91_plus: 0, gl: 0, total: 0 });
  const bucketOf = (d) => d <= 0 ? 'current' : d <= 30 ? 'd1_30' : d <= 60 ? 'd31_60' : d <= 90 ? 'd61_90' : 'd91_plus';
  const dayDiff = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);

  const byVendor = new Map();
  const glRows = [];
  const grand = emptyBuckets();

  for (const it of openItems) {
    const num = invNumFromMemo(it.memo);
    const enrich = (num && billByNum.get(String(num))) || null;
    const isBillcom = syncedEntryIds.has(it.entry_id) && enrich;
    if (isBillcom) {
      const vname = enrich.vendor || ('Vendor');
      const dueDate = enrich.dueDate || it.date;
      const dpd = dayDiff(asOf, it.date); // age by invoice (GL line) date
      const bk = bucketOf(dpd);
      if (!byVendor.has(vname)) byVendor.set(vname, { vendor: vname, rows: [], subtotal: emptyBuckets() });
      const grp = byVendor.get(vname);
      grp.rows.push({ date: it.date, type: 'Bill', num: String(num), vendor: vname, due_date: dueDate, past_due_days: Math.max(0, dpd), amount: it.amount, bucket: bk });
      grp.subtotal[bk] += it.amount; grp.subtotal.total += it.amount;
      grand[bk] += it.amount; grand.total += it.amount;
    } else {
      // GL column: imported/manual entry, not aged, no vendor/invoice
      glRows.push({ date: it.date, entry_num: it.entry_num, entry_id: it.entry_id, memo: it.memo, description: it.description, amount: it.amount });
      grand.gl += it.amount; grand.total += it.amount;
    }
  }

  // Net overpayment: if payments exceeded all bills (202000 is net-debit at the
  // as-of date), the leftover unapplied debit is a prepaid/overpayment balance.
  // Surface it as a negative GL line so the report still ties to the GL balance.
  if (unappliedDebit > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Net prepaid / overpayment (payments exceed open bills)', description: '', amount: -unappliedDebit });
    grand.gl -= unappliedDebit; grand.total -= unappliedDebit;
  }

  const glTotal = glRows.reduce((s, r) => s + r.amount, 0);
  const vendorsOut = Array.from(byVendor.values())
    .sort((a, b) => a.vendor.localeCompare(b.vendor))
    .map(g => ({ ...g, rows: g.rows.sort((x, y) => String(x.date).localeCompare(String(y.date))) }));

  // Reconciliation: report total should equal the GL balance by construction.
  const reportTotal = grand.total;
  const reconDiff = Math.round((reportTotal - glBalance) * 100) / 100;

  res.json({
    entity_id: entityId,
    as_of: asOf,
    ap_account: apAccount,
    source: 'gl',
    bucket_labels: { current: 'Current', d1_30: '1-30', d31_60: '31-60', d61_90: '61-90', d91_plus: '91+', gl: 'GL' },
    bucket_order: buckets,
    vendors: vendorsOut,
    gl_rows: glRows.sort((a, b) => String(a.date).localeCompare(String(b.date))),
    gl_total: glTotal,
    grand_total: grand,
    gl_balance: glBalance,
    recon_diff: reconDiff,
    bill_count: vendorsOut.reduce((n, g) => n + g.rows.length, 0),
    gl_entry_count: glRows.length,
    billcom_error: billcomError,
  });
});












// ═══════════════════════════════════════════════════════════════════════════
// Requisition / Invoice-Packet API (development-project entities only)
// Every route is gated: auth → entity access → development-entity check.
// ═══════════════════════════════════════════════════════════════════════════
const reqGuards = (param) => [auth, requireEntityAccess(param || 'entity_id'), requireDevelopmentEntity(param || 'entity_id')];

// Seed coding history (and optionally the cost-code catalog) from prior Invoice
// Logs. Body: { lines: [{vendor, bill_number, cost_category, cost_code,
// bank_cost_category, gl_coding, cost_code_name, req_number, weight}],
// coa?: [{cost_code, cost_code_name, cost_category, bank_cost_category,
// gl_coding, budget_amount, sort_order}], replace?: bool }.
app.post('/api/requisition/:entity_id/seed-history', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const body = req.body || {};
  const lines = Array.isArray(body.lines) ? body.lines : [];
  const coa = Array.isArray(body.coa) ? body.coa : [];
  const tx = db.transaction(() => {
    if (body.replace) {
      db.prepare('DELETE FROM requisition_coding_history WHERE entity_id = ?').run(eid);
      if (coa.length) db.prepare('DELETE FROM requisition_coa_map WHERE entity_id = ?').run(eid);
    }
    for (const ln of lines) {
      requisition.recordHistory(db, eid, ln, ln.req_number, ln.weight);
    }
    if (coa.length) {
      const now = new Date().toISOString();
      const up = db.prepare(
        'INSERT INTO requisition_coa_map (entity_id, cost_code, cost_code_name, cost_category, bank_cost_category, gl_coding, budget_amount, sort_order, created_at) ' +
        'VALUES (?,?,?,?,?,?,?,?,?) ' +
        'ON CONFLICT(entity_id, cost_code) DO UPDATE SET ' +
        'cost_code_name=excluded.cost_code_name, cost_category=excluded.cost_category, ' +
        'bank_cost_category=excluded.bank_cost_category, gl_coding=excluded.gl_coding, ' +
        'budget_amount=excluded.budget_amount, sort_order=excluded.sort_order'
      );
      coa.forEach((c, i) => {
        if (c.cost_code == null || c.cost_code === '') return;
        up.run(eid, String(c.cost_code), c.cost_code_name || null, c.cost_category || null,
          c.bank_cost_category || null, c.gl_coding || null,
          c.budget_amount != null ? Number(c.budget_amount) : null,
          c.sort_order != null ? Number(c.sort_order) : i, now);
      });
    }
  });
  tx();
  res.json({ seeded_history: lines.length, seeded_coa: coa.length });
});

// Cost-code -> cost-code-name catalog for an entity, used by the Requisition UI
// to auto-fill the Cost Code Name when a code is typed. Primary source is the
// curated requisition_coa_map (canonical spelling, seeded from prior workbooks);
// if that is empty for this entity, fall back to distinct code/name pairs seen
// on previously-saved requisition invoices so the field still auto-fills.
app.get('/api/requisition/:entity_id/coa-map', ...reqGuards(), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const map = {};
  const rows = db.prepare(
    'SELECT cost_code, cost_code_name, cost_category, bank_cost_category, gl_coding ' +
    'FROM requisition_coa_map WHERE entity_id = ? AND cost_code IS NOT NULL'
  ).all(eid);
  for (const r of rows) {
    if (r.cost_code == null || r.cost_code === '') continue;
    map[String(r.cost_code).trim()] = {
      cost_code_name: r.cost_code_name || '',
      cost_category: r.cost_category || '',
      bank_cost_category: r.bank_cost_category || '',
      gl_coding: r.gl_coding || '',
    };
  }
  // Fallback: fill in names from invoice history, preferring the most recent
  // name for a given code. This also REPAIRS curated-map entries whose code is
  // present but whose name is blank (the SRN map seeded codes with empty names),
  // so a code still auto-fills its name when invoice history has one.
  const inv = db.prepare(
    'SELECT cost_code, cost_code_name FROM requisition_invoice ' +
    "WHERE entity_id = ? AND cost_code IS NOT NULL AND TRIM(COALESCE(cost_code_name,'')) <> '' " +
    'ORDER BY req_number DESC, id DESC'
  ).all(eid);
  for (const r of inv) {
    const code = String(r.cost_code).trim();
    if (!code) continue;
    const existing = map[code];
    if (existing && (existing.cost_code_name || '').trim() !== '') continue; // keep curated name
    if (existing) { existing.cost_code_name = r.cost_code_name; continue; } // fill blank curated name
    map[code] = { cost_code_name: r.cost_code_name || '', cost_category: '', bank_cost_category: '', gl_coding: '' };
  }
  res.json({ map });
});


// Body: { lines: [{vendor, bill_number, amount?, invoice_date?, ...}] }.
// Returns per-line { confidence, cost_code, coding, candidates } plus a summary.
app.post('/api/requisition/:entity_id/predict', ...reqGuards(), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const lines = Array.isArray(req.body && req.body.lines) ? req.body.lines : [];
  const index = requisition.buildHistoryIndex(db, eid);
  let high = 0, review = 0, neu = 0;
  const results = lines.map((ln) => {
    const p = requisition.predict(ln, index);
    if (p.confidence === 'high') high++;
    else if (p.confidence === 'review') review++;
    else neu++;
    return {
      vendor: ln.vendor,
      bill_number: ln.bill_number,
      amount: ln.amount != null ? ln.amount : null,
      confidence: p.confidence,
      cost_code: p.cost_code,
      coding: p.coding,
      candidates: p.candidates,
    };
  });
  res.json({
    total: lines.length,
    summary: { high, review, new: neu, auto_coverage: lines.length ? high / lines.length : 0 },
    lines: results,
  });
});

// ─── R4: Stored invoice download ─────────────────────────────────────────────
// Serve the inline-stored PDF/image bytes for one saved invoice, so the invoice
// packet (and manual review) can pull the original document back out of the DB.
app.get('/api/requisition/invoice/:id/download', (req, res) => {
  const token = req.query.token || req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  try { jwt.verify(token, JWT_SECRET); } catch { return res.status(401).json({ error: 'Invalid token' }); }
  const f = db.prepare('SELECT * FROM requisition_invoice WHERE id = ?').get(req.params.id);
  if (!f || !f.file_blob) return res.status(404).json({ error: 'Not found' });
  const isPdf = (f.mime_type === 'application/pdf') || /\.pdf$/i.test(f.original_name || '');
  res.setHeader('Content-Disposition', (isPdf ? 'inline' : 'attachment') + '; filename="' + (f.original_name || 'invoice') + '"');
  res.setHeader('Content-Type', f.mime_type || (isPdf ? 'application/pdf' : 'application/octet-stream'));
  res.send(Buffer.from(f.file_blob));
});

// Purge orphaned requisition invoices (never included in a successful roll-forward,
// i.e. req_number IS NULL). These are leftovers from older builds that saved every
// read invoice; the current flow only persists invoices at roll-forward time.
app.delete('/api/requisition/:entity_id/orphan-invoices', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const before = db.prepare('SELECT COUNT(*) c FROM requisition_invoice WHERE entity_id = ? AND req_number IS NULL').get(eid).c;
    const info = db.prepare('DELETE FROM requisition_invoice WHERE entity_id = ? AND req_number IS NULL').run(eid);
    res.json({ deleted: info.changes, matched: before });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── R4: Read one invoice PDF with Claude (Haiku) and pre-fill fields ─────────
// Upload a single invoice (PDF or image). The model extracts vendor / bill
// number / amount / invoice date; we then run the validated coding engine to
// suggest a cost code. The client renders this as an editable card the user
// corrects before it joins the requisition. Requires ANTHROPIC_API_KEY in env.
//
// multipart/form-data: invoice (file, required)
// Returns: { vendor, bill_number, amount, invoice_date, cost_code,
//            cost_code_name, confidence, candidates, model }
app.post('/api/requisition/:entity_id/read-invoice', ...reqGuards(), requireRole('Admin', 'Accountant'), memUpload.single('invoice'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No invoice file uploaded (field name: invoice)' });
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(503).json({ error: 'Invoice reading is not configured (ANTHROPIC_API_KEY missing on the server).' });

  const eid = parseInt(req.params.entity_id);
  const mime = req.file.mimetype || '';
  const isPdf = mime === 'application/pdf' || /\.pdf$/i.test(req.file.originalname || '');
  const isImage = /^image\//.test(mime);
  if (!isPdf && !isImage) return res.status(400).json({ error: 'Upload a PDF or image invoice' });

  const b64 = req.file.buffer.toString('base64');
  const source = isPdf
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
    : { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } };

  const instruction =
    'You are reading a single vendor invoice for a real-estate development requisition. ' +
    'Extract these fields and return ONLY a JSON object, no prose, no markdown fences:\n' +
    '{"vendor": string|null, "bill_number": string|null, "amount": number|null, "invoice_date": string|null}\n' +
    '- vendor: the company billing us (the payee/remit-to / "from" party), not our company.\n' +
    '- bill_number: the invoice number or, for pay applications, the application label (e.g. "Pay App #15").\n' +
    '- amount: the total amount due for THIS invoice as a number (no currency symbol or commas). ' +
    'Use the current amount due / total due, not running totals.\n' +
    '- invoice_date: the invoice date in YYYY-MM-DD if determinable, else null.\n' +
    'If a field is not present, use null.';

  let extracted;
  try {
    const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 512,
        messages: [{ role: 'user', content: [source, { type: 'text', text: instruction }] }],
      }),
    });
    if (!apiRes.ok) {
      const t = await apiRes.text();
      return res.status(502).json({ error: 'Invoice reader failed (Anthropic ' + apiRes.status + '): ' + t.slice(0, 300) });
    }
    const data = await apiRes.json();
    const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
    let clean = text.replace(/```json|```/g, '').trim();
    // The model sometimes appends prose after the JSON object (common when a PDF
    // holds multiple invoices), which breaks a whole-string JSON.parse. Extract
    // the first balanced {...} object and parse only that.
    try {
      extracted = JSON.parse(clean);
    } catch (firstErr) {
      const start = clean.indexOf('{');
      if (start === -1) throw firstErr;
      let depth = 0, end = -1, inStr = false, esc = false;
      for (let i = start; i < clean.length; i++) {
        const ch = clean[i];
        if (inStr) {
          if (esc) esc = false;
          else if (ch === '\\') esc = true;
          else if (ch === '"') inStr = false;
        } else if (ch === '"') inStr = true;
        else if (ch === '{') depth++;
        else if (ch === '}') { depth--; if (depth === 0) { end = i; break; } }
      }
      if (end === -1) throw firstErr;
      extracted = JSON.parse(clean.slice(start, end + 1));
    }
  } catch (e) {
    return res.status(502).json({ error: 'Invoice reader error: ' + e.message });
  }

  // Suggest a cost code from history using the validated engine.
  let prediction = { confidence: 'new', cost_code: null, coding: null, candidates: [] };
  try {
    const index = requisition.buildHistoryIndex(db, eid);
    prediction = requisition.predict({ vendor: extracted.vendor, bill_number: extracted.bill_number }, index);
  } catch {}

  const amountNum = extracted.amount != null && extracted.amount !== '' ? Number(String(extracted.amount).replace(/[$,]/g, '')) : null;

  // Do NOT persist here. Reading is exploratory — many invoices are read and then
  // discarded during a session, so saving every one (with its file blob) wastes
  // space. We return the extracted fields plus the original bytes (base64) so the
  // client can hold them in the card and send only the kept invoices at
  // roll-forward time, which is when they're persisted with a req_number.
  const finalAmount = Number.isFinite(amountNum) ? amountNum : null;
  const costCodeName = (prediction.coding && prediction.coding.cost_code_name) || null;

  res.json({
    vendor: extracted.vendor || null,
    bill_number: extracted.bill_number || null,
    amount: finalAmount,
    invoice_date: extracted.invoice_date || null,
    cost_code: prediction.cost_code,
    cost_code_name: costCodeName,
    confidence: prediction.confidence,
    candidates: prediction.candidates || [],
    filename: req.file.originalname,
    // Bytes echoed back so the client can resend at roll-forward (not stored server-side yet).
    original_name: req.file.originalname || null,
    mime_type: mime || null,
    file_b64: b64,
  });
});

// Build the rolled-forward output filename from the prior workbook's name,
// bumping the requisition number and the embedded date to the As-of Date while
// preserving the prior name's exact shape. Two requisition-number conventions
// are supported, anchored so we never touch the leading document code digits:
//   1. A hash token   "...Report #11 01.31.2026.xlsx"   (#<num>)
//   2. An underscore-separated token after the word "Report", as produced by
//      the Workpapers save + manual exports:
//        "0005_B1_County_Line_SRN_Requisition_Report__11_01_31_2026.xlsx"
//      i.e. "Report" + "_"(x1-2) + <reqNum> + "_" + <date>. The plain "#"-less
//      digit run is why the old #-only matcher fell through to the generic
//      fallback for these names.
// The embedded date is matched with the same separator used in the prior name
// (".", "/", "-", or "_") and the same 2- vs 4-digit year width.
// Returns a safe generic name only if no requisition-number anchor is found.
function buildRollforwardFilename(originalName, reqNumber, asOfDate) {
  const fallback = 'Requisition_Report' + (reqNumber ? '_' + String(reqNumber) : '') + '.xlsx';
  if (!originalName || typeof originalName !== 'string') return fallback;
  let base = originalName.replace(/\.[^.]+$/, '');// strip extension

  // Parse the As-of Date once; used by both conventions below.
  let mm, dd, yyyy;
  if (asOfDate) {
    const d = new Date(asOfDate + 'T00:00:00');
    if (!isNaN(d)) {
      mm = String(d.getMonth() + 1).padStart(2, '0');
      dd = String(d.getDate()).padStart(2, '0');
      yyyy = String(d.getFullYear());
    }
  }
  const setReq = reqNumber != null && reqNumber !== '';

  // Convention 2 (underscore form) FIRST, matched as ONE token so the req number
  // and the date can never be confused for one another:
  //   "Report" + "_"(x1-2) + <req> + "_" + MM + "_" + DD + "_" + YYYY|YY
  // Matching the whole run lets us rewrite req + date together and is why the
  // prior #-only matcher (which left this form untouched) produced the wrong name.
  const underBlockRe = /(Report_+)(\d+)_(\d{1,2})_(\d{1,2})_(\d{4}|\d{2})(?!\d)/i;
  const um = base.match(underBlockRe);
  if (um) {
    const pfx = um[1];
    const req = setReq ? String(reqNumber) : um[2];
    let newDate = um[3] + '_' + um[4] + '_' + um[5];
    if (mm) {
      const yr = um[5].length === 2 ? yyyy.slice(-2) : yyyy;
      newDate = mm + '_' + dd + '_' + yr;
    }
    base = base.replace(underBlockRe, pfx + req + '_' + newDate);
    return base + '.xlsx';
  }

  // Convention 1 (hash form): bump "#<num>" then the dotted/slashed/dashed date.
  const hashRe = /#\s*(\d+)/;
  if (!hashRe.test(base)) return fallback; // no anchor of either kind -> don't guess
  if (setReq) base = base.replace(hashRe, '#' + String(reqNumber));
  if (mm) {
    const dateRe = /(\d{1,2})([.\/-])(\d{1,2})\2(\d{4}|\d{2})(?!\d)/;
    base = base.replace(dateRe, (m, _mo, sep, _da, yr) => {
      const year = yr.length === 2 ? yyyy.slice(-2) : yyyy;
      return mm + sep + dd + sep + year;
    });
  }
  return base + '.xlsx';
}

// ─── R4: Roll-forward engine route ───────────────────────────────────────────
// Produce Req#N+1 from an uploaded Req#N workbook + the new period's invoices.
// The engine writes formulas but does not evaluate them; production has no
// headless LibreOffice, so verifyRollforward runs WITHOUT recalc here. That
// gates on the structural identities (A1 prior total, A2 per-code, A3 row count,
// B1 group subtotals, B4 absolute refs) which read amounts/formulas directly and
// need no evaluation. A4/B5 (which need evaluated SUBTOTAL/Dev-Fee results)
// degrade to "not evaluated" and do not block.
//
// multipart/form-data:
//   workbook    : the Req#N .xlsx (required)
//   newCurrent  : JSON string — array of invoice rows for the new period, each
//                 { code, name, vendor, bill, amount, date?, req? } (required)
//   reqNumber   : new requisition number (optional, used in titles)
//   asOfDate    : new period as-of date string (optional, used in titles)
//
// On success streams the rolled-forward .xlsx. On a required-check failure
// returns 422 with the reconciliation detail so the caller can see what broke.
app.post('/api/requisition/:entity_id/rollforward', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res, next) => {
  reqRollUpload.single('workbook')(req, res, (err) => {
    if (err) {
      const tooBig = err.code === 'LIMIT_FIELD_VALUE' || err.code === 'LIMIT_FILE_SIZE';
      return res.status(tooBig ? 413 : 400).json({
        error: tooBig
          ? 'Upload too large: the combined invoices/workbook exceeded the size limit. Try rolling forward with fewer invoices at once, or contact support.'
          : 'Upload failed: ' + err.message,
      });
    }
    next();
  });
}, async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No workbook uploaded (field name: workbook)' });

  let newCurrent;
  try {
    newCurrent = JSON.parse(req.body.newCurrent || '[]');
  } catch (e) {
    return res.status(400).json({ error: 'newCurrent must be valid JSON: ' + e.message });
  }
  if (!Array.isArray(newCurrent)) return res.status(400).json({ error: 'newCurrent must be a JSON array of invoice rows' });

  const meta = {};
  if (req.body.reqNumber != null && req.body.reqNumber !== '') meta.reqNumber = req.body.reqNumber;
  if (req.body.asOfDate) meta.asOfDate = req.body.asOfDate;

  // Force flag: when set, a FAILED required reconciliation no longer blocks the
  // download. The roll-forward still runs and is verified, but instead of a 422
  // we stream the (imperfect) workbook + packet and surface which checks failed
  // in the response headers so the user can fix them by hand. This trades a hard
  // gate for "any prepopulation beats starting from scratch" — the user opts in
  // explicitly (the client only sends force=true after seeing the failure).
  const force = req.body.force === 'true' || req.body.force === '1' || req.body.force === true;

  // Invoices that make up this period, sent by the client (not previously stored).
  // Each: { vendor, bill_number, amount, cost_code, cost_code_name, original_name,
  // mime_type, file_b64 }. On a successful roll-forward we persist them with the
  // new req_number and use their bytes to build the invoice packet. Order is the
  // Current Invoice Log order the user arranged on screen.
  let invoicesIn = [];
  try {
    const parsed = JSON.parse(req.body.invoices || '[]');
    if (Array.isArray(parsed)) invoicesIn = parsed;
  } catch {}

  // Load the uploaded Req#N workbook twice: one mutable copy to roll forward,
  // and one untouched copy to supply the prior-period sheets for reconciliation.
  let workbook, priorBook;
  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    priorBook = new ExcelJS.Workbook();
    await priorBook.xlsx.load(req.file.buffer);
  } catch (e) {
    return res.status(400).json({ error: 'Failed to read workbook (.xlsx expected): ' + e.message });
  }

  const prevSheets = {
    prior: priorBook.getWorksheet('Prior Invoice Log'),
    current: priorBook.getWorksheet('Current Invoice Log'),
  };
  if (!prevSheets.prior || !prevSheets.current) {
    return res.status(400).json({ error: 'Workbook is missing required sheets: Prior Invoice Log and/or Current Invoice Log' });
  }

  try {
    // Mutate `workbook` into Req#N+1. The engine also auto-computes this period's
    // Development Fee (entity rate from the Dev Fee tab) and appends it as a
    // Current-Log line; rfResult.devFee carries the amount + row for the packet.
    const rfResult = rollForward(workbook, newCurrent, meta);

    // Verify WITHOUT recalc (no LibreOffice in prod). Structural required checks
    // gate; A4/B5 degrade to "not evaluated". No callClaude here — a failure is
    // surfaced to the caller rather than auto-repaired in this synchronous route.
    const verification = await verifyRollforward({
      prevSheets,
      nextWorkbook: workbook,
      recalc: null,
      callClaude: null,
    });

    if (!verification.ok && !force) {
      return res.status(422).json({
        error: 'Roll-forward failed reconciliation',
        ok: false,
        summary: verification.finalResult && verification.finalResult.summary,
        unresolved: verification.unresolved,
        checks: verification.finalResult && verification.finalResult.checks,
        note: verification.note,
      });
    }
    // When forced, we proceed past a failed required check. The failure detail is
    // still exposed below via X-Reconcile-Summary / X-Reconcile-Failed so the
    // user can see (and hand-correct) what didn't reconcile in the downloaded file.
    const forcedPastFailure = !verification.ok && force;

    const outBuf = await workbook.xlsx.writeBuffer();

    // Persist this period's invoices now (roll-forward succeeded), stamped with
    // the new requisition number. Build the in-memory rows used for the packet.
    const eidInt = parseInt(req.params.entity_id);
    const rn = (meta.reqNumber != null && meta.reqNumber !== '' && Number.isFinite(parseInt(meta.reqNumber))) ? parseInt(meta.reqNumber) : null;
    let invoiceRows = [];
    if (invoicesIn.length) {
      const ins = db.prepare(
        'INSERT INTO requisition_invoice (entity_id, req_number, vendor, bill_number, amount, invoice_date, cost_code, cost_code_name, confidence, original_name, mime_type, file_blob, created_at) ' +
        'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)'
      );
      const nowIso = new Date().toISOString();
      const tx = db.transaction(() => {
        for (const inv of invoicesIn) {
          let blob = null;
          try { if (inv.file_b64) blob = Buffer.from(inv.file_b64, 'base64'); } catch {}
          const amt = inv.amount != null && inv.amount !== '' ? Number(String(inv.amount).replace(/[$,]/g, '')) : null;
          ins.run(
            eidInt, rn,
            inv.vendor || null,
            inv.bill_number || inv.bill || null,
            Number.isFinite(amt) ? amt : null,
            inv.invoice_date || null,
            inv.cost_code || null,
            inv.cost_code_name || null,
            inv.confidence || null,
            inv.original_name || inv.filename || null,
            inv.mime_type || null,
            blob,
            nowIso
          );
          // Row used for packet building (keep bytes in memory; avoids a re-read).
          invoiceRows.push({
            original_name: inv.original_name || inv.filename || null,
            mime_type: inv.mime_type || null,
            file_blob: blob,
            vendor: inv.vendor || null,
            bill_number: inv.bill_number || inv.bill || null,
            amount: Number.isFinite(amt) ? amt : null,
            cost_code: inv.cost_code || null,
            cost_code_name: inv.cost_code_name || null,
          });
        }
      });
      try { tx(); } catch (e) { console.error('requisition invoice persist failed:', e.message); }
    }

    // Build the output filename from the PRIOR workbook's name, bumping the
    // requisition number and (if present) the embedded date to the As-of Date.
    // e.g. "0005 B1 County Line SRN Requisition Report #11 01.31.2026.xlsx"
    //   -> "0005 B1 County Line SRN Requisition Report #12 02.28.2026.xlsx"
    // Falls back to the generic name if the prior name can't be parsed.
    // Derived up front so the Workpapers auto-save uses the SAME name as the
    // download (otherwise the save fell back to a bare "Req N Report.xlsx").
    const fname = buildRollforwardFilename(req.file.originalname, meta.reqNumber, meta.asOfDate);

    // Auto-save the workbook + a merged invoice packet into the entity's
    // Workpapers under "<year>/Requisition Reports/<Month year>" (best-effort:
    // a save failure is logged but never blocks the user's download).
    try {
      const entRow = db.prepare('SELECT name, display_id FROM entities WHERE id = ?').get(eidInt) || {};
      const packetPrefix = (entRow.display_id && entRow.display_id.trim()) || entRow.name || '';
      const saved = await saveRequisitionOutputs({
        db, workpapersDir: WORKPAPERS_DIR, eid: eidInt,
        reqNumber: meta.reqNumber, asOfDate: meta.asOfDate,
        workbookBuffer: Buffer.from(outBuf), invoices: invoiceRows,
        devFee: rfResult && rfResult.devFee && !rfResult.devFee.error ? rfResult.devFee : null,
        who: (req.user && (req.user.name || req.user.email)) || 'system',
        packetPrefix, workbookFilename: fname,
      });
      if (saved.errors && saved.errors.length) console.error('requisition workpaper save:', saved.errors.join('; '));
      res.setHeader('X-Workpaper-Folder', saved.folder || '');
      res.setHeader('X-Workpaper-Saved', JSON.stringify({ workbook: !!saved.workbook, packet: !!saved.packet }));
      // Expose the saved invoice-packet PDF's entity-file id + name so the client
      // can download the packet into the user's Downloads folder alongside the
      // workbook (the packet is also retained in Workpapers via this same id).
      if (saved.packet && saved.packet.id) {
        res.setHeader('X-Packet-File-Id', String(saved.packet.id));
        res.setHeader('X-Packet-File-Name', String(saved.packet.original_name || 'Invoice Packet.pdf').replace(/[\r\n"]/g, ' '));
      }
    } catch (e) {
      console.error('requisition workpaper save failed:', e.message);
    }

    // fname was derived above (shared with the Workpapers auto-save).
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
    // Expose the verification summary in a header so the client can confirm which
    // checks passed without parsing the binary body.
    res.setHeader('X-Reconcile-Summary', JSON.stringify(verification.finalResult ? verification.finalResult.summary : {}));
    // Also expose any checks that didn't pass (e.g. advisory failures) so the
    // client can show their detail on the success card. Header-safe: strip CR/LF.
    try {
      const failed = ((verification.finalResult && verification.finalResult.checks) || []).filter(c => !c.pass);
      res.setHeader('X-Reconcile-Failed', JSON.stringify(failed).replace(/[\r\n]/g, ' '));
    } catch {}
    // Tell the client this download bypassed a failed required check, so it can
    // flag the file as needing manual correction rather than presenting it as clean.
    if (forcedPastFailure) res.setHeader('X-Reconcile-Forced', '1');
    res.send(Buffer.from(outBuf));
  } catch (e) {
    res.status(500).json({ error: 'Roll-forward error: ' + e.message });
  }
});

// ───────────────────────────────────────────────────────────────────────────
// Workpapers › Management Fee (CLRF) — roll a prior-quarter management-fee
// workpaper forward into the next quarter. The uploaded workbook is the single
// source of truth: investor list, group classification, rate tables, BBR/GCM
// tier splits and the ITD invoice history all carry over. Only two things move
// each quarter: (1) the quarter dates (recomputed: end, day counts, stub %),
// and (2) per-investor commitment changes (entered by the user).
//
// Verified against the real Q2->Q3 CLRF workpaper: Standard fees (47/47),
// BBR/GCM tier fees, USC tiered rate, and the grand total ($609,182.73) all
// reproduce exactly. Parsing is header-driven (column positions differ between
// quarters), not fixed-cell.
// ───────────────────────────────────────────────────────────────────────────
const mgmtFeeUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

function mgmtFindCalcSheets(wb) {
  const calc = wb.worksheets.filter(ws => /mgmt\s*fee\s*calc/i.test(ws.name));
  return calc.length ? calc : wb.worksheets;
}
function mgmtHeaderRow(ws) {
  for (let r = 1; r <= 25; r++) {
    const row = ws.getRow(r);
    let found = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (typeof cell.value === 'string' && /InvestorName/i.test(cell.value)) found = true;
    });
    if (found) {
      const cols = {};
      row.eachCell({ includeEmpty: false }, (cell, c) => { if (cell.value != null) cols[String(cell.value).trim()] = c; });
      return { headerRow: r, cols };
    }
  }
  return null;
}
const mgmtNum = (v) => {
  if (v == null) return null;
  // ExcelJS formula cells: { formula, result } or { sharedFormula, result };
  // rich text: { richText:[...] }; hyperlink: { text }.
  if (typeof v === 'object') {
    if (v.result != null) v = v.result;
    else if (v.text != null) v = v.text;
    else if (Array.isArray(v.richText)) v = v.richText.map(t => t.text).join('');
    else return null;
  }
  const n = Number(String(v).replace(/[$,\s]/g, ''));
  return Number.isFinite(n) ? n : null;
};
function mgmtParseWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  return wb.xlsx.load(buffer).then(() => {
    const sheets = mgmtFindCalcSheets(wb);
    let best = null;
    for (const ws of sheets) { const h = mgmtHeaderRow(ws); if (h) { best = { ws, ...h }; break; } }
    if (!best) throw new Error('no "InvestorName" header found in any Mgmt Fee Calc sheet');
    const { ws, headerRow, cols } = best;
    const nameC = cols['InvestorName'], grpC = cols['Investor Group'];
    const endC = cols['InvestorTotal'] || cols['Investor Total'];
    const totC = cols['Total Quarterly Mgmt Fee'];
    if (!nameC || !grpC) throw new Error('missing InvestorName / Investor Group columns');
    const meta = {};
    for (let r = 1; r <= 16; r++) {
      const label = ws.getRow(r).getCell(1).value, val = ws.getRow(r).getCell(2).value;
      const ls = typeof label === 'string' ? label.toLowerCase() : '';
      if (/inception/.test(ls)) meta.inception = val;
      else if (/quarter start/.test(ls)) meta.quarterStart = val;
      else if (/quarter:/.test(ls)) meta.quarterLabel = val;
    }
    const investors = [];
    for (let r = headerRow + 1; r <= headerRow + 90; r++) {
      const nm = ws.getRow(r).getCell(nameC).value;
      if (nm == null || String(nm).trim() === '') continue;
      investors.push({
        name: String(nm).trim(),
        group: ws.getRow(r).getCell(grpC).value != null ? String(ws.getRow(r).getCell(grpC).value).trim() : '',
        ending_commitment: mgmtNum(endC ? ws.getRow(r).getCell(endC).value : null),
        prior_fee: mgmtNum(totC ? ws.getRow(r).getCell(totC).value : null),
      });
    }
    return { investors, meta, sheetName: ws.name };
  });
}
function mgmtNextQuarter(priorStart) {
  const d = new Date(priorStart);
  let ny = d.getUTCFullYear(), nm = d.getUTCMonth() + 3;
  if (nm > 11) { nm -= 12; ny += 1; }
  const start = new Date(Date.UTC(ny, nm, 1));
  const end = new Date(Date.UTC(ny, nm + 3, 0));
  const daysInQuarter = Math.round((end - start) / 86400000) + 1;
  return { start, end, daysInQuarter, label: 'Q' + (Math.floor(nm / 3) + 1) + ' ' + ny };
}

app.post('/api/workpapers/mgmt-fee/:entity_id/analyze', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  mgmtFeeUpload.single('workbook')(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'No workbook uploaded' });
    try {
      const parsed = await mgmtParseWorkbook(req.file.buffer);
      const priorStart = parsed.meta.quarterStart ? new Date(parsed.meta.quarterStart) : null;
      const next = priorStart ? mgmtNextQuarter(priorStart) : null;
      res.json({
        source_sheet: parsed.sheetName,
        prior_quarter: parsed.meta.quarterLabel || null,
        prior_quarter_start: priorStart ? priorStart.toISOString().slice(0,10) : null,
        inception: parsed.meta.inception ? new Date(parsed.meta.inception).toISOString().slice(0,10) : null,
        next_quarter: next ? { label: next.label, start: next.start.toISOString().slice(0,10), end: next.end.toISOString().slice(0,10), days: next.daysInQuarter } : null,
        investor_count: parsed.investors.length,
        groups: parsed.investors.reduce((a,i)=>{a[i.group]=(a[i.group]||0)+1;return a;},{}),
        investors: parsed.investors.map(i => ({ name: i.name, group: i.group, beginning_commitment: i.ending_commitment, change: 0 })),
      });
    } catch (e) {
      res.status(400).json({ error: 'Could not parse workbook: ' + e.message });
    }
  });
});

// ── Generate: roll the workbook forward into the next quarter and return .xlsx.
// Body: multipart with `workbook` (the prior file) + `changes` JSON
// ([{name, change}]) + optional `quarter_start` override.
app.post('/api/workpapers/mgmt-fee/:entity_id/generate', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  mgmtFeeUpload.single('workbook')(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'No workbook uploaded' });
    let changes = [];
    try { if (req.body.changes) changes = JSON.parse(req.body.changes); } catch { return res.status(400).json({ error: 'Invalid changes JSON' }); }
    const changeByName = new Map(changes.map(c => [String(c.name).trim(), mgmtNum(c.change) || 0]));

    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(req.file.buffer);
      const sheets = mgmtFindCalcSheets(wb);
      let best = null;
      for (const ws of sheets) { const h = mgmtHeaderRow(ws); if (h) { best = { ws, ...h }; break; } }
      if (!best) throw new Error('no calc sheet found');
      const { ws, headerRow, cols } = best;
      const nameC = cols['InvestorName'], grpC = cols['Investor Group'];
      const endC = cols['InvestorTotal'] || cols['Investor Total'];
      const begC = cols['Beginning InvestorTotal'];
      const chgC = cols['Change in  Commitment in Qtr'] || cols['New Commitment in Qtr'] || cols['Change in Commitment in Qtr'];

      // Determine the target quarter. The client sends quarter_start already set
      // to the NEW quarter's start (from analyze's next_quarter.start), so use it
      // as-is; only roll forward from the prior start when no override is given.
      let priorStart = null;
      for (let r = 1; r <= 16; r++) {
        const label = ws.getRow(r).getCell(1).value;
        if (typeof label === 'string' && /quarter start/i.test(label)) priorStart = ws.getRow(r).getCell(2).value;
      }
      let next;
      if (req.body.quarter_start) {
        const s = new Date(req.body.quarter_start);
        const end = new Date(Date.UTC(s.getUTCFullYear(), s.getUTCMonth() + 3, 0));
        next = { start: s, end, daysInQuarter: Math.round((end - s) / 86400000) + 1, label: 'Q' + (Math.floor(s.getUTCMonth() / 3) + 1) + ' ' + s.getUTCFullYear() };
      } else {
        next = mgmtNextQuarter(new Date(priorStart));
      }

      // Rates carried from the workbook's Rates tab (fallback to standard 1.5%).
      const ratesTab = wb.getWorksheet('Rates');
      const grpRate = (group) => {
        // Standard inv-period rate is in Rates!C3; default 1.5%
        if (ratesTab) {
          for (let r = 1; r <= 10; r++) {
            const g = ratesTab.getRow(r).getCell(2).value;
            if (g != null && String(g).trim().toLowerCase() === String(group).trim().toLowerCase()) {
              const rate = mgmtNum(ratesTab.getRow(r).getCell(3).value);
              if (rate != null) return rate;
            }
          }
        }
        return group === 'Affiliated LP' || group === 'Affiliated Investor' ? 0 : 0.015;
      };

      // Tier-based quarterly fee for BBR/GCM. Prefer the workbook's already-
      // computed "Quarterly Fee" totals (F8/F10) since the rate cells are
      // cross-sheet formulas whose results ExcelJS may not surface. Fall back to
      // recomputing from tier commitments (row 4) × rates (row 6).
      const tierQuarterlyFee = (tabName) => {
        const t = wb.getWorksheet(tabName);
        if (!t) return null;
        // F10 = Quarterly Fee (Pro-Rated total); F8 = Quarterly Fee (Whole Quarter).
        const f10 = mgmtNum(t.getRow(10).getCell(6).value);
        if (f10 != null && f10 !== 0) return f10;
        const f8 = mgmtNum(t.getRow(8).getCell(6).value);
        if (f8 != null && f8 !== 0) return f8;
        const t1 = mgmtNum(t.getRow(4).getCell(3).value), t2 = mgmtNum(t.getRow(4).getCell(4).value), t3 = mgmtNum(t.getRow(4).getCell(5).value);
        const r1 = mgmtNum(t.getRow(6).getCell(3).value), r2 = mgmtNum(t.getRow(6).getCell(4).value), r3 = mgmtNum(t.getRow(6).getCell(5).value);
        if ([t1,t2,t3,r1,r2,r3].some(x => x == null)) return null;
        return ((t1*r1)+(t2*r2)+(t3*r3)) / 4;
      };
      const uscQuarterlyFee = () => {
        const u = wb.getWorksheet('USC Rate');
        if (!u) return null;
        // B8 = "Total GCM Fee for Quarter" = B2*B6/4, already computed in the book.
        const b8 = mgmtNum(u.getRow(8).getCell(2).value);
        if (b8 != null && b8 !== 0) return b8;
        const calls = mgmtNum(u.getRow(2).getCell(2).value);
        const rate = mgmtNum(u.getRow(6).getCell(2).value);
        if (calls == null || rate == null) return null;
        return calls * rate / 4;
      };

      const pctQuarter = 1.0; // full quarter (CLRF is long past inception)
      let total = 0;
      const lines = [];
      for (let r = headerRow + 1; r <= headerRow + 90; r++) {
        const nm = ws.getRow(r).getCell(nameC).value;
        if (nm == null || String(nm).trim() === '') continue;
        const name = String(nm).trim();
        const group = ws.getRow(r).getCell(grpC).value != null ? String(ws.getRow(r).getCell(grpC).value).trim() : '';
        const priorEnding = mgmtNum(endC ? ws.getRow(r).getCell(endC).value : null) || 0;
        const change = changeByName.get(name) || 0;
        const newBeginning = priorEnding;
        const newEnding = newBeginning + change;
        let fee = 0;
        if (group === 'Standard') fee = Math.round(newEnding * grpRate('Standard') / 4 * pctQuarter * 100) / 100;
        else if (group === 'USC') { /* USC fee is the tab lump; per-investor split not needed for total */ }
        // BBR/GCM/USC totals are tab-driven lumps; we set them at the group level below.
        lines.push({ r, name, group, newBeginning, change, newEnding, fee });
        if (group === 'Standard') total += fee;
      }
      const bbrFee = tierQuarterlyFee('tblBBRCalc') || 0;
      const gcmFee = tierQuarterlyFee('tblGCMCalc') || 0;
      const uscFee = uscQuarterlyFee() || 0;
      total += bbrFee + gcmFee + uscFee;

      // Write by editing the original .xlsx zip in place (JSZip), NOT a full
      // ExcelJS round-trip. ExcelJS rewrite drops drawings/calcChain and makes
      // Excel prompt to "recover" the file. We only set numeric cells (beginning,
      // change, quarter-start date); ending/fees/totals are formulas the workbook
      // recomputes on open (fullCalcOnLoad), so all other parts stay byte-intact.
      const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
      const excelSerial = (dt) => Math.round((Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
      let qStartRow = null;
      for (let r = 1; r <= 16; r++) { const label = ws.getRow(r).getCell(1).value; if (typeof label === 'string' && /quarter start/i.test(label)) { qStartRow = r; break; } }

      const zip = await JSZip.loadAsync(req.file.buffer);
      let wbXml = await zip.file('xl/workbook.xml').async('string');
      const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
      // Map rId -> worksheet target. Attribute order in .rels varies by writer
      // (Excel emits Id..Type..Target; some libs reorder), so match each
      // attribute independently within each <Relationship .../> element.
      const ridToTarget = {};
      for (const rel of relsXml.match(/<Relationship\b[^>]*\/>/g) || []) {
        const tm = rel.match(/Target="([^"]*worksheets\/sheet\d+\.xml)"/);
        const im = rel.match(/Id="(rId\d+)"/);
        if (tm && im) ridToTarget[im[1]] = tm[1].replace(/^\.?\//, '').replace(/^xl\//, '');
      }
      // Locate the <sheet> element for ws.name without depending on attribute
      // order (Excel: name first; ExcelJS: sheetId first). Grab the whole tag
      // by name, then pull r:id from inside it.
      let calcTarget = null;
      for (const tag of wbXml.match(/<sheet\b[^>]*\/>/g) || []) {
        const nm = tag.match(/\bname="([^"]*)"/);
        if (nm && nm[1] === ws.name) {
          const rid = tag.match(/r:id="(rId\d+)"/);
          if (rid) calcTarget = ridToTarget[rid[1]];
          break;
        }
      }
      if (!calcTarget) throw new Error('could not locate calc sheet xml');

      let sheetXml = await zip.file('xl/' + calcTarget).async('string');
      const setNumber = (ref, val) => {
        // A cell is either self-closing (<c r=".."/>, empty) or open/close
        // (<c r=".."> ... </c>). Decide which form THIS cell is before editing —
        // matching the open/close pattern against a self-closing cell would let
        // the lazy [\s\S]*? run past it and swallow the next cell's </c>,
        // corrupting the sheet (mismatched tags -> Excel "recover" prompt).
        const selfClose = new RegExp('<c r="' + ref + '"([^>]*?)/>');
        const scm = sheetXml.match(selfClose);
        if (scm) {
          sheetXml = sheetXml.replace(selfClose, (m, attrs) =>
            '<c r="' + ref + '"' + attrs.replace(/\s+t="[^"]*"/, '') + '><v>' + val + '</v></c>');
          return;
        }
        const re = new RegExp('(<c r="' + ref + '"[^>]*>)([\\s\\S]*?)(</c>)');
        if (re.test(sheetXml)) {
          sheetXml = sheetXml.replace(re, (m, open, inner, close) => {
            const o = open.replace(/\s+t="[^"]*"/, '');
            const fm = inner.match(/<f[\s\S]*?<\/f>|<f[^>]*\/>/);
            return o + (fm ? fm[0] : '') + '<v>' + val + '</v>' + close;
          });
        }
      };
      for (const ln of lines) {
        if (begC) setNumber(colLetter(begC) + ln.r, ln.newBeginning);
        if (chgC) setNumber(colLetter(chgC) + ln.r, ln.change);
        // ending (endC) is a formula =beg+chg; leave it to recompute.
      }
      if (qStartRow) setNumber('B' + qStartRow, excelSerial(next.start));
      zip.file('xl/' + calcTarget, sheetXml);

      if (/<calcPr/.test(wbXml)) wbXml = wbXml.replace(/<calcPr([^/]*)\/>/, (m, a) => /fullCalcOnLoad/.test(a) ? m : '<calcPr' + a + ' fullCalcOnLoad="1"/>');
      else wbXml = wbXml.replace('</workbook>', '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>');
      zip.file('xl/workbook.xml', wbXml);

      const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      const fname = 'CLRF_Mgmt_Fee_Calc_' + next.label.replace(/\s+/g, '_') + '.xlsx';
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
      res.setHeader('X-Mgmt-Fee-Summary', JSON.stringify({ quarter: next.label, total: Math.round(total*100)/100, bbr: bbrFee, gcm: gcmFee, usc: uscFee, standard: Math.round((total-bbrFee-gcmFee-uscFee)*100)/100 }).replace(/[\r\n]/g, ' '));
      res.send(outBuf);
    } catch (e) {
      res.status(500).json({ error: 'Generate error: ' + e.message });
    }
  });
});

if (process.env.NODE_ENV === 'production') app.get('*', (req, res) => {
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.sendFile(path.join(__dirname, '..', 'client', 'dist', 'index.html'));
});
app.listen(PORT, '0.0.0.0', () => console.log(`CloudLedger on port ${PORT}`));
