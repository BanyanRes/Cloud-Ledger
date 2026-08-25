require('dotenv').config();
const express = require('express');
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const morgan = require('morgan');
const multer = require('multer');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const path = require('path');
const fs = require('fs');
const turnkey = require('./turnkey');
const requisition = require('./requisition');
const { rollForward, findSheet: findReqSheet } = require('./requisition_rollforward');
const { verifyRollforward } = require('./requisition_rollforward_verify');
const { finalizeRequisitionWorkbook } = require('./requisition_preserve');
const { makeDevFeeClaudeCaller } = require('./requisition_devfee');
const { saveRequisitionOutputs, saveBufferToWorkpapers: saveWpBuffer, ensureFolders: ensureWpFolders, purgePriorRequisitionCopies } = require('./requisition_workpaper_save');
const reqDraft = require('./requisition_draft');
const { computeAllocation, buildAllocationWorkbook } = require('./insurance_allocation');
const financials = require('./financials');
const financialsXlsx = require('./financials_xlsx');
const execSummaries = require('./execSummaries');
const ExcelJS = require('exceljs');
const xlsxStyledReport = require('./xlsxStyledReport.js');
const JSZip = require('jszip');

const app = express();
// Railway terminates TLS at a single edge proxy and forwards X-Forwarded-For.
// Trust exactly one proxy hop so req.ip resolves to the real client address
// (not the proxy's), which is what express-rate-limit keys the auth limiter on.
// Without this, rate-limit v7 sees an untrusted XFF and can't distinguish
// clients, so the login limiter never triggers. Trusting 1 hop (not `true`)
// prevents clients from spoofing X-Forwarded-For to evade the limit.
app.set('trust proxy', 1);
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET || JWT_SECRET.length < 32) {
  console.error('FATAL: JWT_SECRET is missing or too short (must be set to a random string of at least 32 characters). Refusing to start.');
  process.exit(1);
}
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
// Auto-save of roll-forward outputs into Workpapers. Held OFF during req-report
// testing (set REQ_AUTOSAVE_WORKPAPERS=1 in the environment to re-enable).
const REQ_AUTOSAVE_WORKPAPERS = process.env.REQ_AUTOSAVE_WORKPAPERS === '1';

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
  -- User groups: bundle users so entity access can be granted to the whole group
  -- at once (e.g. all CLA / CliftonLarsonAllen staff). A user's effective entity
  -- access is the UNION of their individual grants and every group they belong to.
  CREATE TABLE IF NOT EXISTS user_groups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE TABLE IF NOT EXISTS user_group_members (
    group_id INTEGER NOT NULL REFERENCES user_groups(id) ON DELETE CASCADE,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (group_id, user_id)
  );
  CREATE TABLE IF NOT EXISTS user_group_entity_access (
    group_id INTEGER NOT NULL REFERENCES user_groups(id) ON DELETE CASCADE,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    created_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (group_id, entity_id)
  );
  -- Per-user entity EXCLUSIONS (negative overrides). A user keeps their group
  -- membership but specific entities are subtracted from their effective access.
  -- Evaluated as: effective = (individual grants UNION group grants) MINUS
  -- exclusions. Only meaningful for a scoped user; an all-access user (no grants,
  -- no groups) is unaffected by exclusions (see listAccessibleEntityIds).
  CREATE TABLE IF NOT EXISTS user_entity_exclusions (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    created_at TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (user_id, entity_id)
  );
  INSERT OR IGNORE INTO user_groups (name) VALUES ('CLA');
  INSERT OR IGNORE INTO user_groups (name) VALUES ('Weaver');
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
  -- Wire coding notes: a note left during the month so that when the bank
  -- statement is uploaded, the matching wire row is auto-populated with its GL
  -- coding (status 'coded') instead of arriving 'pending'. Matches on amount
  -- within a tolerance and a date window; an optional description keyword can
  -- further narrow it. The note is kept for reference after it fires. When
  -- one_shot=1 the note stops matching further rows once it has grabbed one wire.
  CREATE TABLE IF NOT EXISTS bank_coding_notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    bank_account_code TEXT,            -- NULL = any account for this entity
    note TEXT,                         -- free-text reminder shown in the grid
    match_amount REAL NOT NULL,        -- signed; sign is compared too
    amount_tolerance REAL NOT NULL DEFAULT 0,  -- +/- dollars allowed on the match
    date_from TEXT,                    -- YYYY-MM-DD inclusive (NULL = open)
    date_to TEXT,                      -- YYYY-MM-DD inclusive (NULL = open)
    desc_keyword TEXT,                 -- optional case-insensitive substring
    account_code TEXT,                 -- single-account coding (or use splits)
    splits_json TEXT,                  -- JSON [{account_code,amount,memo,project_id,class_id,location_id}]
    memo TEXT,
    project_id TEXT, class_id INTEGER, location_id INTEGER,
    one_shot INTEGER NOT NULL DEFAULT 1,
    active INTEGER NOT NULL DEFAULT 1,
    matched_count INTEGER NOT NULL DEFAULT 0,
    last_matched_at TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE INDEX IF NOT EXISTS idx_bcn_entity ON bank_coding_notes(entity_id, active);
  -- Supporting documents for a wire coding note (email copy, PDF, Excel, etc).
  -- Mirrors journal_attachments: the bytes live on disk under UPLOAD_DIR, this
  -- row holds the metadata. Cascades when the note is deleted.
  CREATE TABLE IF NOT EXISTS bank_coding_note_attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    note_id INTEGER NOT NULL REFERENCES bank_coding_notes(id) ON DELETE CASCADE,
    filename TEXT NOT NULL, original_name TEXT NOT NULL,
    mime_type TEXT, size INTEGER, created_at TEXT DEFAULT (datetime('now'))
  );
  CREATE INDEX IF NOT EXISTS idx_bcna_note ON bank_coding_note_attachments(note_id);
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
  -- departmentId -> CL project. Banyan enters the PROJECT (e.g. "Van Buren") in
  -- Bill.com's Department field, so department is what carries the project onto a
  -- synced JE line. Confirmed by Jimmy 2026-08-19 after CLA (Dennis Arada) found
  -- July/August bills invisible to the project-scoped Custom Detail report: the
  -- sync wrote class_id and location_id only, so every synced line had a NULL
  -- project and was filtered out of any report scoped to a project.
  CREATE TABLE IF NOT EXISTS billcom_project_map (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    billcom_dept_id TEXT NOT NULL,
    billcom_dept_name TEXT,
    cl_project_id INTEGER NOT NULL,
    created_at TEXT,
    UNIQUE(entity_id, billcom_dept_id),
    FOREIGN KEY (entity_id) REFERENCES entities(id)
  );
  CREATE INDEX IF NOT EXISTS idx_bcpm_entity ON billcom_project_map(entity_id);
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
    retainage_receivable_code TEXT,
    cl_project_id INTEGER,
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
  -- Editable / persistent Requisition Report draft (one open draft per entity).
  -- A draft is created once, edited (add/delete/update invoices) and re-rolled on
  -- each save; its in-progress workbook + packet live here (NOT in Workpapers) so
  -- intermediate versions don't clutter the filing tree. Finalize flips status to
  -- 'finalized', stamps the invoices with req_number, and files the workbook +
  -- packet to Workpapers. The finalized draft's output_blob is next month's
  -- auto-seed base. The partial unique index enforces at-most-one open draft.
  CREATE TABLE IF NOT EXISTS requisition_draft (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id     INTEGER NOT NULL REFERENCES entities(id),
    status        TEXT NOT NULL DEFAULT 'open',   -- 'open' | 'finalized'
    phase         TEXT,                           -- rail-assets stream key (e.g. '2','2a'); NULL/'' = single default stream
    req_number    INTEGER,                        -- target req # (editable while open)
    as_of_date    TEXT,                           -- period end (editable while open)
    base_blob     BLOB,                           -- prior workbook this draft rolls from
    base_name     TEXT,                           -- original filename of the base (for name bump)
    base_sha256   TEXT,                           -- hash of base bytes (upload-guard dedupe)
    output_blob   BLOB,                           -- latest rolled-forward workbook
    output_name   TEXT,                           -- filename for the current output/download
    packet_blob   BLOB,                           -- latest merged invoice packet PDF
    packet_name   TEXT,
    recon_ok      INTEGER,                        -- last verify result (1/0/null)
    recon_summary TEXT,                           -- last verify summary (JSON, for the banner)
    created_at    TEXT,
    updated_at    TEXT,
    finalized_at  TEXT,
    created_by    TEXT
  );
  CREATE UNIQUE INDEX IF NOT EXISTS uq_reqdraft_open ON requisition_draft(entity_id, IFNULL(phase,'')) WHERE status='open';
  CREATE INDEX IF NOT EXISTS idx_reqdraft_entity ON requisition_draft(entity_id, status);
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
// Vendor/payee name for the entry (e.g. Bill.com bill JEs), so account detail can
// show who an A/P transaction is with. Populated at Bill.com sync.
if (!jeCols.includes('vendor')) db.exec("ALTER TABLE journal_entries ADD COLUMN vendor TEXT");
// Document / invoice number for the entry (CLA request, 8/2026): the source
// document's own number - a Bill.com invoiceNumber, a check number, an AR invoice
// number - so every detail report can show a Doc column beside the JE number.
if (!jeCols.includes('doc_number')) { db.exec("ALTER TABLE journal_entries ADD COLUMN doc_number TEXT"); console.log('[db migrate] journal_entries.doc_number added'); }

// Entity type: 'accounting' (default, standard ledger entity) | 'development' | 'shell' (tracks location + investor/class dimensions)
// (real-estate development project; unlocks Requisition Report / Invoice Packet features)
const entCols = db.prepare("PRAGMA table_info(entities)").all().map(c => c.name);
if (!entCols.includes('entity_type')) {
  db.exec("ALTER TABLE entities ADD COLUMN entity_type TEXT NOT NULL DEFAULT 'accounting'");
  console.log('[db migrate] entities.entity_type added (default accounting)');
}
// Per-entity access LEVEL on individual grants: 'full' (create/edit, like an
// Accountant) or 'view' (read-only, like a Viewer). Lets one user be Full on some
// entities and View-only on others. Backfill preserves today's access: a Viewer's
// existing grants become 'view'; everyone else's stay 'full'.
const ueaCols = db.prepare("PRAGMA table_info(user_entity_access)").all().map(c => c.name);
if (!ueaCols.includes('access_level')) {
  db.exec("ALTER TABLE user_entity_access ADD COLUMN access_level TEXT NOT NULL DEFAULT 'full'");
  db.exec("UPDATE user_entity_access SET access_level = 'view' WHERE user_id IN (SELECT id FROM users WHERE role = 'Viewer')");
  console.log('[db migrate] user_entity_access.access_level added (viewers backfilled to view)');
}
// Same per-entity level on GROUP grants: a group can confer Full or View-only
// access to each of its entities. Existing group grants default to 'full'.
const ugeaCols = db.prepare("PRAGMA table_info(user_group_entity_access)").all().map(c => c.name);
if (!ugeaCols.includes('access_level')) {
  db.exec("ALTER TABLE user_group_entity_access ADD COLUMN access_level TEXT NOT NULL DEFAULT 'full'");
  console.log('[db migrate] user_group_entity_access.access_level added (default full)');
}
// display_id: short user-facing identifier (e.g. "0005 B1a") used as a filename
// prefix for requisition invoice packets. Optional; falls back to entity name.
if (!entCols.includes('display_id')) {
  db.exec("ALTER TABLE entities ADD COLUMN display_id TEXT");
  console.log('[db migrate] entities.display_id added');
}

// draft_id: links a requisition_invoice to the editable draft that owns it while
// open. NULL for every pre-existing (already-finalized) invoice, so the orphan
// purge (req_number IS NULL) is narrowed to also require draft_id IS NULL and can
// never delete a live draft's lines.
const reqInvCols = db.prepare("PRAGMA table_info(requisition_invoice)").all().map(c => c.name);
if (!reqInvCols.includes('draft_id')) {
  db.exec("ALTER TABLE requisition_invoice ADD COLUMN draft_id INTEGER");
  db.exec("CREATE INDEX IF NOT EXISTS idx_reqinv_draft ON requisition_invoice(draft_id)");
  console.log('[db migrate] requisition_invoice.draft_id added');
}

// phase: rail-assets requisition stream key (e.g. '2', '2a'). Rail-assets
// entities can run two requisitions for the same month (e.g. Phase 2 and Phase
// 2a); each is its own draft/stream. NULL/'' is the single default stream used
// by every non-rail entity. The open-draft unique index must be phase-scoped so
// two open drafts on different phases can coexist — the original entity-only
// index is dropped and recreated here for existing databases.
const reqDraftCols = db.prepare("PRAGMA table_info(requisition_draft)").all().map(c => c.name);
if (!reqDraftCols.includes('phase')) {
  db.exec("ALTER TABLE requisition_draft ADD COLUMN phase TEXT");
  db.exec("DROP INDEX IF EXISTS uq_reqdraft_open");
  db.exec("CREATE UNIQUE INDEX IF NOT EXISTS uq_reqdraft_open ON requisition_draft(entity_id, IFNULL(phase,'')) WHERE status='open'");
  console.log('[db migrate] requisition_draft.phase added (open-draft index now phase-scoped)');
}

// Phase 3: default_cash_account on billcom_config (for payment JEs)
const bcCfgCols = db.prepare("PRAGMA table_info(billcom_config)").all().map(c => c.name);
if (!bcCfgCols.includes('default_cash_account')) db.exec("ALTER TABLE billcom_config ADD COLUMN default_cash_account TEXT");
// Phase 7 (payment reconcile): clearing account for the QBO-style two-leg model.
// Leg 1 relieves AP into the clearing account; Leg 2 moves the clearing balance
// to operating cash as a process-date lump sum (mirrors Bill.com Money Out Clearing).
if (!bcCfgCols.includes('default_clearing_account')) db.exec("ALTER TABLE billcom_config ADD COLUMN default_clearing_account TEXT");
if (!bcCfgCols.includes('sync_cutoff_date')) db.exec("ALTER TABLE billcom_config ADD COLUMN sync_cutoff_date TEXT");
// A/P aging dedupe: the parsed lines of the last uploaded A/P aging detail (JSON
// array of {vendor, invoice_number, bill_date, amount}), used to skip Bill.com
// bills already booked in the GL. Set from the A/P Aging "Upload aging detail" flow.
if (!bcCfgCols.includes('ap_aging_lines_json')) db.exec("ALTER TABLE billcom_config ADD COLUMN ap_aging_lines_json TEXT");
if (!bcCfgCols.includes('ap_aging_as_of')) db.exec("ALTER TABLE billcom_config ADD COLUMN ap_aging_as_of TEXT");
if (!bcCfgCols.includes('ap_aging_uploaded_at')) db.exec("ALTER TABLE billcom_config ADD COLUMN ap_aging_uploaded_at TEXT");

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
// Bank transactions and their split lines carry the same optional dimensions as
// journal lines (Location / Class / Project), so coding a bank txn (or a split)
// can tag the dimension that then flows onto the posted JE line.
for (const _tbl of ['bank_transactions', 'bank_transaction_splits']) {
  const _cols = db.prepare(`PRAGMA table_info(${_tbl})`).all().map(c => c.name);
  if (!_cols.includes('project_id'))  db.exec(`ALTER TABLE ${_tbl} ADD COLUMN project_id TEXT`);
  if (!_cols.includes('class_id'))    db.exec(`ALTER TABLE ${_tbl} ADD COLUMN class_id INTEGER`);
  if (!_cols.includes('location_id')) db.exec(`ALTER TABLE ${_tbl} ADD COLUMN location_id INTEGER`);
}
console.log('[db migrate] bank_transactions/splits dimension columns ensured');
// AR cash application: a split line can carry the ar_invoice it pays, so posting
// a deposit coded to the A/R control account also records the subledger receipt.
{
  const _sc = db.prepare("PRAGMA table_info(bank_transaction_splits)").all().map(c => c.name);
  if (!_sc.includes('invoice_id')) db.exec('ALTER TABLE bank_transaction_splits ADD COLUMN invoice_id INTEGER');
  console.log('[db migrate] bank_transaction_splits.invoice_id ensured');
}
// Dimension code columns (name was the only label originally; code added for reporting/sorting)
const dcCols = db.prepare("PRAGMA table_info(dim_classes)").all().map(c => c.name);
if (!dcCols.includes('code')) { db.exec("ALTER TABLE dim_classes ADD COLUMN code TEXT"); console.log('[db migrate] dim_classes.code added'); }
const dlCols = db.prepare("PRAGMA table_info(dim_locations)").all().map(c => c.name);
if (!dlCols.includes('code')) { db.exec("ALTER TABLE dim_locations ADD COLUMN code TEXT"); console.log('[db migrate] dim_locations.code added'); }
// Per-entity switch to hide dimension tagging (Location/Class/Project) across the
// UI — the Dimensions manager, report dimension filters, and the Bank Transactions
// coding column. Added because a bulk project catalog was fanned out (apply_all)
// to every accounting/development entity, surfacing an unwanted Dimensions
// selector on entities that don't use dimensions (e.g. SRN, CLR Silsbee). Seeded
// ON for those two on first run; toggle in the DB (hide_dims 0/1) to change later.
const entHideDimCols = db.prepare("PRAGMA table_info(entities)").all().map(c => c.name);
if (!entHideDimCols.includes('hide_dims')) {
  db.exec("ALTER TABLE entities ADD COLUMN hide_dims INTEGER NOT NULL DEFAULT 0");
  const seeded = db.prepare("UPDATE entities SET hide_dims=1 WHERE code IN ('SABINERI','CLRSILSB2')").run();
  console.log('[db migrate] entities.hide_dims added; seeded hidden for ' + seeded.changes + ' entity(ies) (SABINERI, CLRSILSB2)');
}
const bslCols = db.prepare("PRAGMA table_info(billcom_sync_log)").all().map(c => c.name);
if (!bslCols.includes('invoice_number')) { db.exec("ALTER TABLE billcom_sync_log ADD COLUMN invoice_number TEXT"); console.log('[db migrate] billcom_sync_log.invoice_number added'); }
// ── Fund reporting (CLRF, entity 40) ────────────────────────────────────────
// GP vs LP designation on investor classes. Defaults to 'LP'; specific classes
// are tagged 'GP' via the Fund Reporting admin UI. Drives the GP/LP columns of
// the Statement of Assets/Liabilities/Partners' Capital and the Statement of
// Changes in Partners' Capital.
if (!dcCols.includes('partner_type')) {
  db.exec("ALTER TABLE dim_classes ADD COLUMN partner_type TEXT NOT NULL DEFAULT 'LP'");
  console.log('[db migrate] dim_classes.partner_type added');
}
// Schedule of Investments look-through detail that is NOT derivable from the
// fund's trial balance (per-underlying acquisition date, proportional cost and
// fair value, and the holding-company grouping). One row per underlying
// investment; edited in the Fund Reporting admin UI. sort_order controls
// display order; parent_name groups underlyings beneath a holding company
// (e.g. "CLRFI Midco I, LLC").
db.exec(`
  CREATE TABLE IF NOT EXISTS fund_investments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
    parent_name TEXT DEFAULT '',
    name TEXT NOT NULL,
    acquisition_date TEXT,
    cost REAL NOT NULL DEFAULT 0,
    fair_value REAL NOT NULL DEFAULT 0,
    sort_order INTEGER NOT NULL DEFAULT 0,
    notes TEXT,
    created_at TEXT,
    updated_at TEXT
  );
  CREATE INDEX IF NOT EXISTS idx_fundinv_entity ON fund_investments(entity_id);
`);
console.log('[db migrate] fund_investments ensured');
// turnkey_project_map redesigned: no longer stores per-project account codes
// (single COA on the company entity now). We add cl_entity_id linking to the
// COMPANY entity (same for all projects). Keep existing rows on upgrade.
const tpmCols = db.prepare("PRAGMA table_info(turnkey_project_map)").all().map(c => c.name);
if (!tpmCols.includes('project_code')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN project_code TEXT");
  console.log('[db migrate] turnkey_project_map.project_code added');
}
if (!tpmCols.includes('retainage_receivable_code')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN retainage_receivable_code TEXT");
  console.log('[db migrate] turnkey_project_map.retainage_receivable_code added');
}
if (!tpmCols.includes('cl_project_id')) {
  db.exec("ALTER TABLE turnkey_project_map ADD COLUMN cl_project_id INTEGER");
  console.log('[db migrate] turnkey_project_map.cl_project_id added');
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
    if (nextPage) params.set('page', nextPage);
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
  // Dedupe by account id. Bill.com can repeat the same account across paginated
  // responses (nextPage overlap) or return archived + active copies, which bloats
  // the mapping UI ("unusually long list") and makes the mapping save collide on
  // UNIQUE(entity_id, billcom_account_id). One row per id is what every caller wants.
  const _seen = new Set();
  const _dedup = [];
  for (const a of out) {
    const id = a && a.id != null ? String(a.id) : null;
    if (id !== null) { if (_seen.has(id)) continue; _seen.add(id); }
    _dedup.push(a);
  }
  return _dedup;
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
    if (nextPage) params.set('page', nextPage);
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
    if (nextPage) params.set('page', nextPage);
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

async function billcomGetById({ sessionId, devKey, baseUrl, resourcePath, id, extraParams }) {
  const qs = extraParams ? ('?' + new URLSearchParams(extraParams).toString()) : '';
  const url = (baseUrl || BILLCOM_BASE_URLS.sandbox) + resourcePath + '/' + encodeURIComponent(id) + qs;
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

// Fetch bills by updatedTime window instead of dueDate. A bill approved after the
// sync cutoff must have been updated at/after approval, so this captures late,
// back-dated invoices a dueDate window (anchored on the cutoff month) would miss.
// Same month-windowing to dodge v3's broken offset pagination.
async function billcomListBillsByUpdatedWindowed({ sessionId, devKey, baseUrl, fromDate, toDate }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const hdr = { sessionId, devKey, Accept: "application/json" };
  const addMonth = (d) => { const [Y, M] = d.split("-"); let yy = +Y, mm = +M + 1; if (mm > 12) { mm = 1; yy++; } return yy + "-" + String(mm).padStart(2, "0") + "-01"; };
  const startYM = fromDate.slice(0, 7) + "-01";
  const endExclusive = addMonth(toDate.slice(0, 7) + "-01");
  const byId = new Map();
  let win = startYM;
  let guard = 0;
  while (win < endExclusive && guard < 240) {
    guard++;
    const winEnd = addMonth(win);
    const filt = "updatedTime:gte:" + win + ",updatedTime:lt:" + winEnd;
    const url = base + "/bills?max=100&billApprovals=true&filters=" + encodeURIComponent(filt);
    let json;
    try {
      const resp = await billcomFetch(url, { method: "GET", headers: hdr }, 20000);
      const text = await resp.text();
      try { json = JSON.parse(text); } catch { throw new Error("Non-JSON bills window (HTTP " + resp.status + ")"); }
      if (!resp.ok) { const msg = Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join("; ") : (json.message || ("HTTP " + resp.status)); throw new Error("bills window: " + msg); }
    } catch (e) { throw new Error("bills updated-window " + win + ": " + e.message); }
    const results = Array.isArray(json.results) ? json.results : [];
    for (const b of results) { const id = b && b.id; if (id != null && !byId.has(String(id))) byId.set(String(id), b); }
    if (results.length >= 100) console.log("[billcom sync] WARNING updated-window " + win + " hit 100-row cap; may be truncated");
    win = winEnd;
  }
  return Array.from(byId.values());
}

// Approval-complete date of a bill = latest statusChangedTime among approvers that
// have APPROVED it (requires billApprovals=true on the fetch). Falls back to the
// bill createdTime when no approver approval timestamp is present, so legacy
// pre-conversion bills (old createdTime) stay excluded while a newly-entered bill
// is judged by when it was entered.
function billApprovalDate(bill) {
  const aps = Array.isArray(bill && bill.approvers) ? bill.approvers : [];
  let latest = null;
  for (const a of aps) {
    if (String((a && a.status) || "").toUpperCase() !== "APPROVED") continue;
    const t = a.statusChangedTime || null;
    if (t && (!latest || String(t) > String(latest))) latest = t;
  }
  return latest || (bill && bill.createdTime) || null;
}

// TRUE only when EVERY approver on the bill — across all approval layers — has
// APPROVED. Banyan/CLA policy (8/2026): Bill.com carries two approval layers and
// the sync must hold a bill until all approvers in every layer have signed off.
// The bill's overall approvalStatus can read APPROVED before that is true depending
// on how the multi-layer policy resolves, so we check the per-approver list
// directly (requires billApprovals=true on the fetch, i.e. use the DETAIL object).
// Returns false if the approver list is absent or empty, so a bill can never sync
// on missing approver data — approval must be positively demonstrated, never assumed.
function allApproversApproved(bill) {
  const aps = Array.isArray(bill && bill.approvers) ? bill.approvers : [];
  if (aps.length === 0) return false;
  for (const a of aps) {
    const s = String((a && a.status) || '').toUpperCase();
    if (s !== 'APPROVED') return false; // any ASSIGNED / PENDING / DENIED layer holds the bill
  }
  return true;
}

// TRUE when AT LEAST ONE approver has APPROVED (Banyan policy, 8/2026: a single
// approval is sufficient; a bill need not clear every approver or layer). Falls
// back to the bill's overall approvalStatus === 'APPROVED' when no per-approver
// list is present. DENIED bills are filtered earlier and never reach this check.
function anyApproverApproved(bill) {
  const aps = Array.isArray(bill && bill.approvers) ? bill.approvers : [];
  for (const a of aps) {
    if (String((a && a.status) || '').toUpperCase() === 'APPROVED') return true;
  }
  return String((bill && bill.approvalStatus) || '').toUpperCase() === 'APPROVED';
}

// Bill.com v3 /payments has the SAME broken offset pagination as /bills (nextPage
// returns the same first 100 rows), so a plain paged fetch silently caps at the
// first page and never sees newer payments. Mirror the bills approach: walk
// month-sized processDate windows and union+dedupe by id. Each CLRF window is far
// below the 100-row cap. Returns the full payment set across [fromDate, toDate].
async function billcomListPaymentsWindowed({ sessionId, devKey, baseUrl, fromDate, toDate }) {
  const base = (baseUrl || BILLCOM_BASE_URLS.sandbox);
  const hdr = { sessionId, devKey, Accept: "application/json" };
  const addMonth = (d) => { const [Y, M] = d.split("-"); let yy = +Y, mm = +M + 1; if (mm > 12) { mm = 1; yy++; } return yy + "-" + String(mm).padStart(2, "0") + "-01"; };
  const startYM = fromDate.slice(0, 7) + "-01";
  const endExclusive = addMonth(toDate.slice(0, 7) + "-01");
  const byId = new Map();
  let win = startYM;
  let guard = 0;
  while (win < endExclusive && guard < 240) {
    guard++;
    const winEnd = addMonth(win);
    const filt = "processDate:gte:" + win + ",processDate:lt:" + winEnd;
    const url = base + "/payments?max=100&filters=" + encodeURIComponent(filt);
    let json;
    try {
      const resp = await billcomFetch(url, { method: "GET", headers: hdr }, 20000);
      const text = await resp.text();
      try { json = JSON.parse(text); } catch { throw new Error("Non-JSON payments window (HTTP " + resp.status + ")"); }
      if (!resp.ok) { const msg = Array.isArray(json) ? json.map(e => e.message || JSON.stringify(e)).join("; ") : (json.message || ("HTTP " + resp.status)); throw new Error("payments window: " + msg); }
    } catch (e) { throw new Error("payments window " + win + ": " + e.message); }
    const results = Array.isArray(json.results) ? json.results : [];
    for (const p of results) { const id = p && p.id; if (id != null && !byId.has(String(id))) byId.set(String(id), p); }
    if (results.length >= 100) console.log("[billcom-sync] WARNING payments window " + win + " hit 100-row cap; may be truncated");
    win = winEnd;
  }
  return Array.from(byId.values());
}


// ── Bill.com legacy v2 API: read GL Posting Date ──
// The v3 Connect API does NOT return a bill's GL Posting Date, but the classic v2
// API does. We use v2 (same stored credentials) purely to READ glPostingDate and
// key it back to the v3 bills the sync processes. No v2 writes are ever made.
const BILLCOM_V2_BASE = 'https://api.bill.com/api/v2';
function billcomV2Form(obj) { return Object.entries(obj).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&'); }
async function billcomV2Login({ username, password, orgId, devKey }) {
  const r = await billcomFetch(BILLCOM_V2_BASE + '/Login.json', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: billcomV2Form({ userName: username, password, orgId, devKey })
  }, 20000);
  const j = await r.json();
  if (j.response_status !== 0) throw new Error('v2 login: ' + (j.response_message || JSON.stringify(j).slice(0, 200)));
  return j.response_data.sessionId;
}
// Fetch bills from v2 whose glPostingDate is on/after fromDate, paginated. Each v2
// bill carries id, invoiceNumber, amount, vendorId, invoiceDate, glPostingDate.
async function billcomV2ListBillsByGlPosting({ sessionId, devKey, fromDate }) {
  const out = [];
  let start = 0; const max = 999; let guard = 0;
  while (guard < 100) {
    guard++;
    const data = JSON.stringify({ start, max, filters: [{ field: 'glPostingDate', op: '>=', value: fromDate }] });
    const r = await billcomFetch(BILLCOM_V2_BASE + '/List/Bill.json', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: billcomV2Form({ devKey, sessionId, data })
    }, 25000);
    const j = await r.json();
    if (j.response_status !== 0) throw new Error('v2 List/Bill: ' + (j.response_message || JSON.stringify(j).slice(0, 200)));
    const rows = Array.isArray(j.response_data) ? j.response_data : [];
    for (const b of rows) out.push(b);
    if (rows.length < max) break;
    start += max;
  }
  return out;
}
// Build lookup maps so the sync can resolve a v3 bill's GL Posting Date. Keyed by
// BOTH raw id (in case v2/v3 ids coincide) and a stable identity (invoiceNumber|amount).
function billcomBuildGlPostingMap(v2bills) {
  const byId = new Map();
  const byIdent = new Map();
  const identKey = (num, amt) => String(num == null ? '' : num).trim() + '|' + (amt == null ? '' : Number(amt).toFixed(2));
  for (const b of (v2bills || [])) {
    const gp = b && b.glPostingDate ? String(b.glPostingDate).slice(0, 10) : null;
    if (!gp) continue;
    if (b.id != null) byId.set(String(b.id), gp);
    const k = identKey(b.invoiceNumber, b.amount);
    if (!byIdent.has(k)) byIdent.set(k, gp);
  }
  return { byId, byIdent, identKey };
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

// Rate limiter for authentication endpoints. Login and password-reset are the
// prime targets for credential-stuffing and brute-force attempts, so cap each
// client IP to a small number of attempts per window. Successful requests do not
// count against the limit, so legitimate users are never locked out by normal use.
const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 10,                  // 10 attempts per IP per window
  standardHeaders: true,
  legacyHeaders: false,
  skipSuccessfulRequests: true,
  message: { error: 'Too many attempts. Please wait a few minutes and try again.' },
});
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
  const ids = listAccessibleEntityIds(userId, userRole);
  if (ids === null) return true; // null = all entities
  return ids.includes(parseInt(entityId));
}
// A user's effective entity access = the UNION of their individual grants and
// the entities assigned to every group they belong to (e.g. the CLA group).
// Legacy behavior is preserved: a user with NO individual grants AND NO group
// membership has an empty allowlist, which still means "all entities". Once a
// user has either an individual grant or a group membership, they are scoped to
// the union (so putting an all-access user into CLA restricts them to CLA's
// entities). Returns null for "all", otherwise an array of entity ids.
function listAccessibleEntityIds(userId, userRole) {
  if (userRole === 'Admin') return null; // null = all
  const direct = db.prepare('SELECT entity_id FROM user_entity_access WHERE user_id = ?').all(userId);
  const groups = db.prepare('SELECT group_id FROM user_group_members WHERE user_id = ?').all(userId);
  if (direct.length === 0 && groups.length === 0) return null; // empty = all (legacy)
  const groupEnt = groups.length
    ? db.prepare('SELECT DISTINCT entity_id FROM user_group_entity_access WHERE group_id IN (' + groups.map(() => '?').join(',') + ')').all(...groups.map(g => g.group_id))
    : [];
  const set = new Set([...direct.map(r => r.entity_id), ...groupEnt.map(r => r.entity_id)]);
  // Subtract per-user exclusions (negative overrides). Only applied to a scoped
  // user, which we already are here (direct.length || groups.length > 0), so an
  // all-access user is never accidentally narrowed by a stray exclusion row.
  const excl = db.prepare('SELECT entity_id FROM user_entity_exclusions WHERE user_id = ?').all(userId);
  for (const r of excl) set.delete(r.entity_id);
  return [...set];
}
// Effective access level for a user on ONE entity: 'full' (create/edit) or 'view'
// (read-only). Admin is always full. A direct per-user grant carries its own level;
// access that comes only through a group is full; a legacy all-access user (no
// grants at all) inherits from their global role (Viewer => view, else full).
function entityAccessLevel(userId, userRole, entityId) {
  if (userRole === 'Admin') return 'full';
  const row = db.prepare('SELECT access_level FROM user_entity_access WHERE user_id = ? AND entity_id = ?').get(userId, entityId);
  if (row) return row.access_level === 'view' ? 'view' : 'full';
  // Group grants for this entity: Full wins over View across the user's groups.
  const grpLevels = db.prepare('SELECT ga.access_level lvl FROM user_group_entity_access ga JOIN user_group_members m ON m.group_id = ga.group_id WHERE m.user_id = ? AND ga.entity_id = ?').all(userId, entityId);
  if (grpLevels.length) return grpLevels.some(g => g.lvl !== 'view') ? 'full' : 'view';
  const direct = db.prepare('SELECT 1 FROM user_entity_access WHERE user_id = ? LIMIT 1').get(userId);
  const grp = db.prepare('SELECT 1 FROM user_group_members WHERE user_id = ? LIMIT 1').get(userId);
  if (!direct && !grp) return userRole === 'Viewer' ? 'view' : 'full'; // legacy all-access
  return 'full';
}
function requireEntityAccess(paramName) {
  return (req, res, next) => {
    const eid = parseInt(req.params[paramName || 'eid']);
    if (!eid) return res.status(400).json({ error: 'Invalid entity id' });
    if (!userHasEntityAccess(req.user.id, req.user.role, eid)) return res.status(403).json({ error: 'No access to this entity' });
    // Scope the effective role to THIS entity so write gates honor a per-entity
    // 'view' grant even when the user's global role would otherwise permit writes.
    const lvl = entityAccessLevel(req.user.id, req.user.role, eid);
    req.entityRole = req.user.role === 'Admin' ? 'Admin' : (lvl === 'view' ? 'Viewer' : 'Accountant');
    next();
  };
}

// Write/section gate. When an entity-scoped route ran requireEntityAccess first,
// req.entityRole holds the caller's level FOR THAT ENTITY and takes precedence;
// otherwise the caller's global role is used (non-entity routes).
function requireRole(...roles) { return (req, res, next) => { const role = req.entityRole || req.user.role; if (!roles.includes(role) && role !== 'Admin') return res.status(403).json({ error: 'Forbidden' }); next(); }; }

// Gate Requisition/Invoice-Packet features to development-project entities only.
// Reads the entity id from the named route param (default 'entity_id'); rejects
// non-development entities so accounting entities never expose these endpoints.
function requireDevelopmentEntity(paramName) {
  return (req, res, next) => {
    const eid = parseInt(req.params[paramName || 'entity_id']);
    if (!eid) return res.status(400).json({ error: 'Invalid entity id' });
    const ent = db.prepare('SELECT entity_type FROM entities WHERE id = ?').get(eid);
    if (!ent) return res.status(404).json({ error: 'Entity not found' });
    if (ent.entity_type !== 'development' && ent.entity_type !== 'rail_assets') return res.status(403).json({ error: 'Requisition features are only available for development-project and rail-assets entities' });
    next();
  };
}

// ═══ Auth ═══
// Login returns a single generic message for both "no such account" and "wrong
// password" so the endpoint can't be used to enumerate which emails have accounts
// (a login form that says "no account found" confirms every email you try).
// The rate limiter caps brute-force attempts per IP.
app.post('/api/auth/login', authLimiter, (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: 'Email and password required' });
  const user = db.prepare('SELECT * FROM users WHERE email = ?').get(email?.toLowerCase());
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }
  const token = jwt.sign({ id: user.id, name: user.name, email: user.email, role: user.role }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
});
// Account creation is Admin-only. This endpoint previously had no auth guard and
// accepted a caller-chosen `role`, which let anyone reach the API mint themselves
// an Admin account. It now requires a valid Admin JWT (same path as /api/users
// POST) so roles can never be self-assigned.
app.post('/api/auth/signup', auth, requireRole('Admin'), (req, res) => {
  const { name, email, password, role } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'All fields required' });
  if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters' });
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

// ── Per-user UI preferences ────────────────────────────────────────────────
// A single JSON blob per user for client-side layout choices that should follow
// the person across browsers and machines (currently the sidebar's per-category
// item order). Deliberately schema-less: the client owns the shape, so adding a
// new preference never needs a migration. PUT does a shallow merge so one client
// writing navOrder cannot clobber a key it does not know about.
db.exec(`
  CREATE TABLE IF NOT EXISTS user_prefs (
    user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
    prefs TEXT NOT NULL DEFAULT '{}',
    updated_at TEXT DEFAULT (datetime('now'))
  );
`);

function readPrefs(userId) {
  const row = db.prepare('SELECT prefs FROM user_prefs WHERE user_id = ?').get(userId);
  if (!row) return {};
  try { const p = JSON.parse(row.prefs); return (p && typeof p === 'object' && !Array.isArray(p)) ? p : {}; }
  catch { return {}; }
}

app.get('/api/me/prefs', auth, (req, res) => res.json(readPrefs(req.user.id)));

app.put('/api/me/prefs', auth, (req, res) => {
  const patch = req.body;
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return res.status(400).json({ error: 'Body must be a JSON object' });
  const next = Object.assign(readPrefs(req.user.id), patch);
  const json = JSON.stringify(next);
  if (json.length > 100000) return res.status(413).json({ error: 'Preferences too large' });
  db.prepare(`INSERT INTO user_prefs (user_id, prefs, updated_at) VALUES (?, ?, datetime('now'))
    ON CONFLICT(user_id) DO UPDATE SET prefs = excluded.prefs, updated_at = datetime('now')`)
    .run(req.user.id, json);
  res.json(next);
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
app.get('/api/users', auth, requireRole('Admin'), (req, res) => res.json(db.prepare('SELECT id, name, email, role, created_at FROM users ORDER BY name COLLATE NOCASE ASC').all()));
app.delete('/api/users/:id', auth, requireRole('Admin'), (req, res) => { if (+req.params.id === req.user.id) return res.status(400).json({ error: 'Cannot delete self' }); db.prepare('DELETE FROM users WHERE id = ?').run(req.params.id); res.json({ success: true }); });
app.put('/api/users/:id', auth, requireRole('Admin'), (req, res) => { db.prepare('UPDATE users SET name = ?, role = ? WHERE id = ?').run(req.body.name, req.body.role, req.params.id); res.json({ success: true }); });

// Entity access management (Admin only)
app.get('/api/users/:id/entity-access', auth, requireRole('Admin'), (req, res) => {
  const uid = parseInt(req.params.id);
  const rows = db.prepare('SELECT entity_id, access_level FROM user_entity_access WHERE user_id = ? ORDER BY entity_id').all(uid);
  // Also surface access the user gets through group membership, so the modal can
  // show effective access (individual grants alone are misleading once a user is
  // in a group like CLA/Weaver).
  const groups = db.prepare('SELECT g.id, g.name FROM user_groups g JOIN user_group_members m ON m.group_id = g.id WHERE m.user_id = ? ORDER BY g.name').all(uid);
  const groupsDetail = groups.map(g => ({
    id: g.id, name: g.name,
    entity_ids: db.prepare('SELECT entity_id FROM user_group_entity_access WHERE group_id = ? ORDER BY entity_id').all(g.id).map(r => r.entity_id),
  }));
  const groupEntityIds = [...new Set(groupsDetail.flatMap(g => g.entity_ids))];
  const exclusions = db.prepare('SELECT entity_id FROM user_entity_exclusions WHERE user_id = ? ORDER BY entity_id').all(uid).map(r => r.entity_id);
  const user = db.prepare('SELECT role FROM users WHERE id = ?').get(uid);
  // effective === null means "all entities"; otherwise (union of individual + group) minus exclusions.
  let effective;
  if (user && user.role === 'Admin') effective = null;
  else if (rows.length === 0 && groups.length === 0) effective = null;
  else effective = [...new Set([...rows.map(r => r.entity_id), ...groupEntityIds])].filter(id => !exclusions.includes(id));
  const levels = Object.fromEntries(rows.map(r => [r.entity_id, r.access_level === 'view' ? 'view' : 'full']));
  res.json({ user_id: uid, entity_ids: rows.map(r => r.entity_id), levels, groups: groupsDetail, group_entity_ids: groupEntityIds, exclusions, effective });
});
app.put('/api/users/:id/entity-access', auth, requireRole('Admin'), (req, res) => {
  const userId = parseInt(req.params.id);
  const targetUser = db.prepare('SELECT id, role FROM users WHERE id = ?').get(userId);
  if (!targetUser) return res.status(404).json({ error: 'User not found' });
  if (targetUser.role === 'Admin') return res.status(400).json({ error: 'Admins always have all-entity access; cannot restrict' });
  const ids = Array.isArray(req.body.entity_ids) ? req.body.entity_ids.map(n => parseInt(n)).filter(n => Number.isInteger(n)) : [];
  const levels = (req.body.levels && typeof req.body.levels === 'object') ? req.body.levels : {};
  const lvlFor = (eid) => (String(levels[eid] != null ? levels[eid] : (levels[String(eid)] || 'full')) === 'view' ? 'view' : 'full');
  const exclusions = Array.isArray(req.body.exclusions) ? req.body.exclusions.map(n => parseInt(n)).filter(n => Number.isInteger(n)) : [];
  const tx = db.transaction(() => {
    db.prepare('DELETE FROM user_entity_access WHERE user_id = ?').run(userId);
    const ins = db.prepare('INSERT INTO user_entity_access (user_id, entity_id, access_level) VALUES (?, ?, ?)');
    for (const eid of ids) ins.run(userId, eid, lvlFor(eid));
    db.prepare('DELETE FROM user_entity_exclusions WHERE user_id = ?').run(userId);
    const insX = db.prepare('INSERT OR IGNORE INTO user_entity_exclusions (user_id, entity_id) VALUES (?, ?)');
    for (const eid of exclusions) insX.run(userId, eid);
  });
  tx();
  res.json({ user_id: userId, entity_ids: ids, levels, exclusions });
});



// ═══ User Groups (Admin only) ═══
// List groups with member + entity counts.
app.get('/api/groups', auth, requireRole('Admin'), (req, res) => {
  const groups = db.prepare('SELECT id, name, created_at FROM user_groups ORDER BY name').all();
  const mc = db.prepare('SELECT group_id, COUNT(*) n FROM user_group_members GROUP BY group_id').all();
  const ec = db.prepare('SELECT group_id, COUNT(*) n FROM user_group_entity_access GROUP BY group_id').all();
  const mmap = Object.fromEntries(mc.map(r => [r.group_id, r.n]));
  const emap = Object.fromEntries(ec.map(r => [r.group_id, r.n]));
  res.json(groups.map(g => ({ ...g, member_count: mmap[g.id] || 0, entity_count: emap[g.id] || 0 })));
});
// Group detail: member user ids + granted entity ids.
app.get('/api/groups/:id', auth, requireRole('Admin'), (req, res) => {
  const g = db.prepare('SELECT id, name, created_at FROM user_groups WHERE id = ?').get(req.params.id);
  if (!g) return res.status(404).json({ error: 'Group not found' });
  const member_ids = db.prepare('SELECT user_id FROM user_group_members WHERE group_id = ?').all(g.id).map(r => r.user_id);
  const entRows = db.prepare('SELECT entity_id, access_level FROM user_group_entity_access WHERE group_id = ? ORDER BY entity_id').all(g.id);
  const entity_ids = entRows.map(r => r.entity_id);
  const levels = Object.fromEntries(entRows.map(r => [r.entity_id, r.access_level === 'view' ? 'view' : 'full']));
  res.json({ ...g, member_ids, entity_ids, levels });
});
// Create a group.
app.post('/api/groups', auth, requireRole('Admin'), (req, res) => {
  const name = (req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Group name required' });
  try { const r = db.prepare('INSERT INTO user_groups (name) VALUES (?)').run(name); res.json({ id: r.lastInsertRowid, name }); }
  catch { res.status(400).json({ error: 'A group with that name already exists' }); }
});
// Delete a group (memberships + entity grants cascade).
app.delete('/api/groups/:id', auth, requireRole('Admin'), (req, res) => {
  db.prepare('DELETE FROM user_groups WHERE id = ?').run(req.params.id);
  res.json({ success: true });
});
// Replace a group's members.
app.put('/api/groups/:id/members', auth, requireRole('Admin'), (req, res) => {
  const gid = parseInt(req.params.id);
  if (!db.prepare('SELECT id FROM user_groups WHERE id = ?').get(gid)) return res.status(404).json({ error: 'Group not found' });
  const ids = Array.isArray(req.body.user_ids) ? req.body.user_ids.map(n => parseInt(n)).filter(n => Number.isInteger(n)) : [];
  const tx = db.transaction(() => {
    db.prepare('DELETE FROM user_group_members WHERE group_id = ?').run(gid);
    const ins = db.prepare('INSERT OR IGNORE INTO user_group_members (group_id, user_id) VALUES (?, ?)');
    for (const uid of ids) ins.run(gid, uid);
  });
  tx();
  res.json({ group_id: gid, user_ids: ids });
});
// Replace a group's entity access. Every member gains access to these entities.
app.put('/api/groups/:id/entities', auth, requireRole('Admin'), (req, res) => {
  const gid = parseInt(req.params.id);
  if (!db.prepare('SELECT id FROM user_groups WHERE id = ?').get(gid)) return res.status(404).json({ error: 'Group not found' });
  const ids = Array.isArray(req.body.entity_ids) ? req.body.entity_ids.map(n => parseInt(n)).filter(n => Number.isInteger(n)) : [];
  const levels = (req.body.levels && typeof req.body.levels === 'object') ? req.body.levels : {};
  const lvlFor = (eid) => (String(levels[eid] != null ? levels[eid] : (levels[String(eid)] || 'full')) === 'view' ? 'view' : 'full');
  const tx = db.transaction(() => {
    db.prepare('DELETE FROM user_group_entity_access WHERE group_id = ?').run(gid);
    const ins = db.prepare('INSERT OR IGNORE INTO user_group_entity_access (group_id, entity_id, access_level) VALUES (?, ?, ?)');
    for (const eid of ids) ins.run(gid, eid, lvlFor(eid));
  });
  tx();
  res.json({ group_id: gid, entity_ids: ids, levels });
});

// ═══ Entities ═══
app.get('/api/entities', auth, (req, res) => {
  const ids = listAccessibleEntityIds(req.user.id, req.user.role);
  // Attach the caller's per-entity access level so the client can gate write UI.
  const withLevel = (rows) => rows.map(e => ({ ...e, access_level: entityAccessLevel(req.user.id, req.user.role, e.id) }));
  if (ids === null) return res.json(withLevel(db.prepare('SELECT * FROM entities ORDER BY code').all()));
  if (ids.length === 0) return res.json([]);
  const placeholders = ids.map(() => '?').join(',');
  res.json(withLevel(db.prepare('SELECT * FROM entities WHERE id IN (' + placeholders + ') ORDER BY code').all(...ids)));
});
app.post('/api/entities', auth, requireRole('Admin','Accountant'), (req, res) => {
  const { name } = req.body; if (!name) return res.status(400).json({ error: 'Name required' });
  const entityType = ['development','shell','operating','rail_assets'].includes(req.body.entity_type) ? req.body.entity_type : 'accounting';
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
    if (!['development','accounting','shell','operating','rail_assets'].includes(req.body.entity_type)) return res.status(400).json({ error: 'entity_type must be development, accounting, shell, operating, or rail_assets' });
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
  // 7xxxx-9xxxx (5-digit) is ambiguous in these charts: some are other-income
  // accounts (70000 Interest Income), others are other-expense (82000
  // Amortization Expense). Code alone can't disambiguate, so return null here
  // and let the caller fall back to the account name.
  return null;
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
  // First pass: collect every table row's cells.
  const allRows = [];
  let m;
  while ((m = trRe.exec(scope))) {
    const cells = []; let c;
    tdRe.lastIndex = 0;
    while ((c = tdRe.exec(m[1]))) cells.push(clean(c[1]));
    if (cells.length) allRows.push(cells);
  }
  // Locate the transaction header row (names both Debit and Credit) and map
  // columns by NAME. Intacct exports vary in how many dimension columns they
  // include — e.g. a single "Project" (9 cols) vs. "Department" + "Location"
  // (10 cols) — which shifts the Debit/Credit positions. Hardcoding indices
  // misreads the wider layouts (reading the JNL code as the debit), which
  // scrambles every entry so nothing balances.
  let headerNames = null;
  for (const r of allRows) {
    const low = r.map(x => String(x).toLowerCase());
    if (low.some(x => x.includes('debit')) && low.some(x => x.includes('credit'))) {
      headerNames = r.map((c, i) => (String(c).trim() || ('Column ' + (i + 1))));
      break;
    }
  }
  if (!headerNames) return null;
  let dateIdx = headerNames.findIndex(h => /posted|post date|transaction date|^date$/i.test(h));
  if (dateIdx < 0) dateIdx = 0;
  const isTxnDate = s => /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(s == null ? '' : s).trim());
  const out = [];
  let cur = null; // {code, name}
  for (const cells of allRows) {
    // Account section header: first cell contains "Balance forward".
    if (cells.length >= 2 && /balance forward/i.test(cells[0])) {
      const hm = HDR.exec(cells[0]);
      if (hm) cur = { code: hm[1].trim(), name: hm[2].trim() };
      continue;
    }
    if (/^totals for/i.test(cells[0])) continue; // account subtotal
    // Transaction row: the posting-date column holds a date, inside an account section.
    if (cur && isTxnDate(cells[dateIdx])) {
      const obj = { 'Account Number': cur.code, 'Account Name': cur.name };
      headerNames.forEach((h, j) => { obj[h] = (cells[j] != null ? cells[j] : ''); });
      out.push(obj);
    }
  }
  if (!out.length) return null;
  return { columns: ['Account Number', 'Account Name', ...headerNames], rows: out };
}

// RealPage/Yardi-style property "General Ledger" export (.xlsx). Layout: a
// header row (Property, Property Name, Date, Period, Person/Description, Control,
// Reference, Debit, Credit, Balance, Remarks); then per account a "= Beginning
// Balance =" row whose Property cell holds the account code (e.g. "11020-000")
// and whose Person/Description cell holds the account name; then transaction
// rows (Property cell holds the property code, not the account, so there is no
// per-row account column); then an "= Ending Balance =" row (no date; carries
// account totals). Flatten by carrying the code/name down onto each transaction
// row. The Control value (the J-###### document id) becomes the reference so the
// two sides of each entry group into one balanced JE. Returns null otherwise.
function glParseYardiXlsxGrid(aoa) {
  const lc = s => String(s == null ? '' : s).toLowerCase().trim();
  const cell = (r, i) => String((r && r[i] != null) ? r[i] : '').trim();
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 25); i++) {
    const names = (aoa[i] || []).map(lc);
    if (names.some(n => n === 'property') && names.some(n => n.includes('remarks')) &&
        names.some(n => n.includes('debit')) && names.some(n => n.includes('credit'))) { hdrIdx = i; break; }
  }
  if (hdrIdx === -1) return null;
  const hdr = aoa[hdrIdx].map(c => String(c).trim());
  const find = pred => hdr.findIndex(h => pred(lc(h)));
  const dateIdx = find(n => n === 'date');
  const descIdx = find(n => n.includes('description'));
  const ctrlIdx = find(n => n === 'control');
  const debitIdx = find(n => n.includes('debit'));
  const creditIdx = find(n => n.includes('credit'));
  if (dateIdx < 0 || debitIdx < 0 || creditIdx < 0) return null;
  const isDate = v => { const s = String(v == null ? '' : v).trim(); return /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s) || /^\d{4}-\d{1,2}-\d{1,2}/.test(s); };
  const ACCT = /^\d{3,6}-\d{2,4}$/;   // account code e.g. 11020-000, 14105-007
  const rows = [];
  let curCode = null, curName = null, acctHeaders = 0;
  for (let i = hdrIdx + 1; i < aoa.length; i++) {
    const r = aoa[i];
    if (!r) continue;
    const first = cell(r, 0);
    if (ACCT.test(first)) {                       // account section boundary (Beginning Balance row)
      curCode = first;
      const nm = descIdx >= 0 ? cell(r, descIdx) : '';
      if (nm) curName = nm;                        // Person/Description holds the account name here
      acctHeaders++;
      continue;
    }
    const dcell = dateIdx >= 0 ? cell(r, dateIdx) : '';
    if (curCode && isDate(dcell)) {                // transaction row (ending-balance rows have no date)
      rows.push({
        'Account Number': curCode,
        'Account Name': curName || '',
        'Date': r[dateIdx],
        'Reference': ctrlIdx >= 0 ? r[ctrlIdx] : '',
        'Person/Description': descIdx >= 0 ? r[descIdx] : '',
        'Debit': r[debitIdx],
        'Credit': r[creditIdx],
      });
    }
  }
  if (acctHeaders < 2 || rows.length === 0) return null;
  return { columns: ['Account Number', 'Account Name', 'Date', 'Reference', 'Person/Description', 'Debit', 'Credit'], rows };
}

// Sage Intacct's "General Ledger report" saved as a real .xlsx (not the HTML
// variant handled above) lays each account out as a banded section-header row
// ("<code> - <name> (Balance forward ...)"), then its transaction rows, then a
// "Totals for ..." row — with NO account column on the transaction rows. Detect
// that shape from the parsed grid and flatten it, carrying the account code/name
// down onto every transaction row. Returns null for a normal flat table (one
// that already has its own account column) so the generic reader handles it.
function glParseIntacctXlsxGrid(aoa) {
  const lc = s => String(s == null ? '' : s).toLowerCase().trim();
  const cell = (r, i) => String((r && r[i] != null) ? r[i] : '').trim();
  // Transaction header row = the one that names both Debit and Credit.
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 25); i++) {
    const names = (aoa[i] || []).map(lc);
    if (names.some(n => n.includes('debit')) && names.some(n => n.includes('credit'))) { hdrIdx = i; break; }
  }
  if (hdrIdx === -1) return null;
  const headerNames = aoa[hdrIdx].map((c, i) => { const t = String(c).trim(); return t || ('Column ' + (i + 1)); });
  // If a real account column already exists, this is a flat file — let the generic path handle it.
  const hasAcctCol = headerNames.some(h => {
    const n = lc(h);
    return n === 'account' || n === 'code' || n === 'gl account' || n.includes('account number') ||
           n.includes('account #') || n.includes('account code') || n.includes('acct');
  });
  if (hasAcctCol) return null;
  let dateIdx = headerNames.findIndex(h => { const n = lc(h); return n.includes('posted') || n.includes('post date') || n.includes('transaction date') || n === 'date'; });
  if (dateIdx === -1) dateIdx = 0;
  const debitIdx = headerNames.findIndex(h => lc(h).includes('debit'));
  const creditIdx = headerNames.findIndex(h => lc(h).includes('credit'));
  const isDate = v => { const s = String(v == null ? '' : v).trim(); return /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s) || /^\d{4}-\d{1,2}-\d{1,2}/.test(s); };
  const isNum = v => { const s = String(v == null ? '' : v).replace(/[,$()\s]/g, ''); return s !== '' && !isNaN(Number(s)); };
  const ACCT = /^(\S+)\s*-\s*(.+)$/;
  const stripSuffix = s => s.replace(/\s*\((?:balance forward|as of).*$/i, '').trim();
  const rows = [];
  let curCode = null, curName = null, acctHeaders = 0;
  for (let i = hdrIdx + 1; i < aoa.length; i++) {
    const r = aoa[i];
    if (!r) continue;
    const first = cell(r, 0);
    const dcell = cell(r, dateIdx);
    const am = first ? ACCT.exec(first) : null;
    const isAcctHdr = am && /\d/.test(am[1]) && !/\s/.test(am[1]) && !isDate(dcell);
    if (isAcctHdr) { curCode = am[1].trim(); curName = stripSuffix(am[2]); acctHeaders++; continue; }
    if (/^totals for|^grand total/i.test(first)) continue;
    const looksTxn = isDate(dcell) || (debitIdx >= 0 && isNum(r[debitIdx])) || (creditIdx >= 0 && isNum(r[creditIdx]));
    if (looksTxn && curCode != null) {
      const obj = { 'Account Number': curCode, 'Account Name': curName };
      headerNames.forEach((h, j) => { obj[h] = (r[j] == null ? '' : r[j]); });
      rows.push(obj);
    }
  }
  if (acctHeaders < 2 || rows.length === 0) return null;
  return { columns: ['Account Number', 'Account Name', ...headerNames], rows };
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
  // RealPage/Yardi property GL saved as .xlsx (banded, account code in Property col) — flatten if detected.
  const yardi = glParseYardiXlsxGrid(aoa);
  if (yardi) return yardi;
  // Intacct GL report saved as .xlsx (banded account sections) — flatten if detected.
  const banded = glParseIntacctXlsxGrid(aoa);
  if (banded) return banded;
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
        if (/payable|liabilit|accrued|due to|note payable|loan payable/.test(s)) return 'Liability';
        if (/receivable|due from|cash|bank|prepaid|investment|asset/.test(s)) return 'Asset';
        // Check expense keywords BEFORE income so names like 'Interest Expense'
        // or 'Loan Fee Expense' resolve to Expense even though the 7xxxx code is
        // ambiguous. 'Refund' stays income.
        if (/expense|amortization|depreciation|donation/.test(s)) return 'Expense';
        if (/income|revenue|gain|refund|dividend/.test(s)) return 'Revenue';
        if (/cost|fee|loss/.test(s)) return 'Expense';
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
  res.json(db.prepare(`SELECT c.id, c.name, c.code, c.kind, c.partner_type, COUNT(jl.id) AS line_count
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
  let partner_type = row.partner_type || 'LP';
  if (req.body.partner_type !== undefined) partner_type = (String(req.body.partner_type).toUpperCase() === 'GP') ? 'GP' : 'LP';
  try { db.prepare('UPDATE dim_classes SET name = ?, code = ?, kind = ?, partner_type = ? WHERE id = ? AND entity_id = ?').run(name, code, kind, partner_type, req.params.id, req.params.eid);
    res.json({ id: Number(req.params.id), name, code, kind, partner_type }); }
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

// ── Upsert a single investor class's commitment amount by class id (find-or-
//    create). Used by the Fund Reporting "Odyssey commitment" box, which updates
//    a GP commitment that changes monthly. ──
app.put('/api/entities/:eid/commitments/by-class/:classId', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const eid = req.params.eid;
  const classId = parseInt(req.params.classId);
  if (!classId) return res.status(400).json({ error: 'classId is required' });
  const cls = db.prepare('SELECT id, name FROM dim_classes WHERE id = ? AND entity_id = ?').get(classId, eid);
  if (!cls) return res.status(404).json({ error: 'Investor class not found in this entity' });
  const commitment = Number(req.body.commitment_amount || 0);
  const now = new Date().toISOString();
  const existing = db.prepare('SELECT id FROM investor_commitments WHERE entity_id = ? AND class_id = ?').get(eid, classId);
  if (existing) {
    db.prepare('UPDATE investor_commitments SET commitment_amount = ?, updated_at = ? WHERE id = ?').run(commitment, now, existing.id);
    return res.json({ id: existing.id, class_id: classId, commitment_amount: commitment, updated: true });
  }
  const r = db.prepare(`INSERT INTO investor_commitments (entity_id, class_id, commitment_amount, called_amount, commit_date, notes, created_at, updated_at)
    VALUES (?, ?, ?, 0, NULL, NULL, ?, ?)`).run(eid, classId, commitment, now, now);
  res.json({ id: r.lastInsertRowid, class_id: classId, commitment_amount: commitment, created: true });
});

// ── Fund GP/LP allocation preview. Returns the commitment-based ownership split
//    and the resulting net-loss allocation for a given as-of date, so the Fund
//    Reporting UI can show the GP % and GP/LP net-loss split without generating
//    the full PDF. ──
app.get('/api/entities/:eid/fund-allocation', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  try {
    const eid = req.params.eid;
    const asOf = (req.query && req.query.as_of);
    // GP classes and their commitments.
    const classes = db.prepare('SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?').all(eid);
    const gpIds = new Set(classes.filter(c => String(c.partner_type).toUpperCase() === 'GP').map(c => c.id));
    const commits = db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
    const totalCommit = commits.reduce((s, x) => s + (Number(x.commitment_amount) || 0), 0);
    const gpCommit = commits.filter(x => gpIds.has(x.class_id)).reduce((s, x) => s + (Number(x.commitment_amount) || 0), 0);
    const gpShare = totalCommit > 0 ? gpCommit / totalCommit : 0;

    // Per-GP-class commitment detail for display.
    const nameById = new Map(classes.map(c => [c.id, c.name]));
    const gpDetail = commits
      .filter(x => gpIds.has(x.class_id))
      .map(x => ({ class_id: x.class_id, name: nameById.get(x.class_id) || ('Class ' + x.class_id), commitment_amount: Number(x.commitment_amount) || 0 }))
      .sort((a, b) => a.name.localeCompare(b.name));

    // Net loss for the period (only if a valid as-of is supplied).
    let niYtd = null, gpNetLoss = null, lpNetLoss = null;
    if (asOf && /^\d{4}-\d{2}-\d{2}$/.test(asOf)) {
      const ys = asOf.slice(0, 4) + '-01-01';
      const isYtd = computeBalances(eid, { from: ys, to: asOf });
      let ni = 0;
      for (const r of isYtd) { if (r.type === 'Revenue') ni += (r.balance || 0); else if (r.type === 'Expense') ni -= (r.balance || 0); }
      niYtd = Math.round(ni * 100) / 100;
      gpNetLoss = totalCommit > 0 ? Math.round(niYtd * gpShare * 100) / 100 : 0;
      lpNetLoss = Math.round((niYtd - gpNetLoss) * 100) / 100;
    }

    res.json({
      totalCommitment: Math.round(totalCommit * 100) / 100,
      gpCommitment: Math.round(gpCommit * 100) / 100,
      gpSharePct: Math.round(gpShare * 1000000) / 10000, // percent, 4 dp
      gpDetail,
      hasCommitments: totalCommit > 0,
      asOf: asOf || null,
      netLossYtd: niYtd,
      gpNetLoss,
      lpNetLoss,
    });
  } catch (e) {
    res.status(500).json({ error: 'Allocation error: ' + e.message });
  }
});

// ── Fund investments (Schedule of Investments look-through for CLRF-style fund
//    packages). Informational; never posts to the GL. One row per underlying,
//    grouped under an optional parent_name (holding company). ──
app.get('/api/entities/:eid/fund-investments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const rows = db.prepare(`SELECT id, parent_name, name, acquisition_date, cost, fair_value, sort_order, notes
    FROM fund_investments WHERE entity_id = ? ORDER BY sort_order, id`).all(req.params.eid);
  res.json(rows);
});
app.post('/api/entities/:eid/fund-investments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const name = (req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Name is required' });
  const now = new Date().toISOString();
  const r = db.prepare(`INSERT INTO fund_investments (entity_id, parent_name, name, acquisition_date, cost, fair_value, sort_order, notes, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(
    req.params.eid,
    (req.body.parent_name || '').trim(),
    name,
    (req.body.acquisition_date || '').trim() || null,
    Number(req.body.cost) || 0,
    Number(req.body.fair_value) || 0,
    Number(req.body.sort_order) || 0,
    (req.body.notes || '').trim() || null,
    now, now);
  res.json({ id: r.lastInsertRowid });
});
app.patch('/api/entities/:eid/fund-investments/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const row = db.prepare('SELECT * FROM fund_investments WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!row) return res.status(404).json({ error: 'Not found' });
  const b = req.body;
  const name = b.name !== undefined ? (b.name || '').trim() : row.name;
  if (!name) return res.status(400).json({ error: 'Name is required' });
  const parent_name = b.parent_name !== undefined ? (b.parent_name || '').trim() : row.parent_name;
  const acquisition_date = b.acquisition_date !== undefined ? ((b.acquisition_date || '').trim() || null) : row.acquisition_date;
  const cost = b.cost !== undefined ? (Number(b.cost) || 0) : row.cost;
  const fair_value = b.fair_value !== undefined ? (Number(b.fair_value) || 0) : row.fair_value;
  const sort_order = b.sort_order !== undefined ? (Number(b.sort_order) || 0) : row.sort_order;
  const notes = b.notes !== undefined ? ((b.notes || '').trim() || null) : row.notes;
  db.prepare(`UPDATE fund_investments SET parent_name=?, name=?, acquisition_date=?, cost=?, fair_value=?, sort_order=?, notes=?, updated_at=?
    WHERE id=? AND entity_id=?`).run(parent_name, name, acquisition_date, cost, fair_value, sort_order, notes, new Date().toISOString(), req.params.id, req.params.eid);
  res.json({ id: Number(req.params.id), parent_name, name, acquisition_date, cost, fair_value, sort_order, notes });
});
app.delete('/api/entities/:eid/fund-investments/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  db.prepare('DELETE FROM fund_investments WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
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

// AR invoicing engine: per-entity settings, recurring templates, invoices with
// accrual JEs, invoice PDF, review-then-send email, cash receipts, A/R aging.
require('./ar').registerArRoutes(app, {
  db,
  auth,
  requireEntityAccess,
  requireRole,
  workpapersDir: WORKPAPERS_DIR,
  verifyToken: (t) => jwt.verify(t, JWT_SECRET),
  getResendKey: () => RESEND_API_KEY,
  getFromEmail: () => (process.env.AR_FROM_EMAIL || RESET_FROM_EMAIL),
});

// ── Workpapers › Insurance Allocation ───────────────────────────────────────
// Banyan Residential pays the monthly health-insurance premium and allocates it
// across the four commonly-owned entities. The accountant uploads the carrier
// billing invoice + the consolidated billing report; we compute each entity's
// employer/employee split, build the workpaper, file it under Workpapers, and
// return it for download with a summary header for the on-screen view.
app.post('/api/entities/:eid/insurance-allocation',
  auth, requireEntityAccess(), requireRole('Admin', 'Accountant'),
  memUpload.fields([{ name: 'invoice', maxCount: 1 }, { name: 'consolidated', maxCount: 1 }]),
  async (req, res) => {
    try {
      const inv = req.files && req.files.invoice && req.files.invoice[0];
      const con = req.files && req.files.consolidated && req.files.consolidated[0];
      if (!inv || !con) return res.status(400).json({ error: 'Upload both the health-insurance billing invoice and the consolidated billing report.' });

      const result = computeAllocation({ invoiceBuf: inv.buffer, consolidatedBuf: con.buffer });

      // Period folder + filename from the invoice's billing period start date.
      const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      let year = null, monthLabel = 'Current';
      const md = String(result.period || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (md) { year = md[3]; monthLabel = MONTHS[(parseInt(md[1], 10) || 1) - 1] + ' ' + md[3]; }
      const folderPath = `${year || new Date().getFullYear()}/Insurance Allocation`;
      const fname = `Insurance Allocation ${monthLabel}.xlsx`;

      const entRow = db.prepare('SELECT name FROM entities WHERE id=?').get(req.params.eid);
      const wbBuf = Buffer.from(await buildAllocationWorkbook(result, { entityName: (entRow && entRow.name) || '', title: 'Health Insurance Allocation' }));

      // File the workpaper under the entity's Workpapers tree (best-effort).
      let saved = null;
      try {
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        ensureWpFolders(db, req.params.eid, folderPath, who);
        const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        saved = saveWpBuffer(db, WORKPAPERS_DIR, req.params.eid, folderPath, fname, XLSX_MIME, wbBuf, who, { overwrite: true });
        // Retain the two original uploads alongside the workpaper for audit.
        const srcFolder = folderPath + '/Source Documents';
        ensureWpFolders(db, req.params.eid, srcFolder, who);
        saveWpBuffer(db, WORKPAPERS_DIR, req.params.eid, srcFolder, inv.originalname || 'Billing Invoice.xlsx', inv.mimetype || XLSX_MIME, inv.buffer, who, { overwrite: true });
        saveWpBuffer(db, WORKPAPERS_DIR, req.params.eid, srcFolder, con.originalname || 'Consolidated Billing.xlsx', con.mimetype || XLSX_MIME, con.buffer, who, { overwrite: true });
      } catch (e) { console.error('[insurance-allocation] save failed:', e.message); }

      const summary = {
        period: result.period, invoice: result.invoice, subscriberCount: result.subscriberCount,
        entities: result.entities, subtotal: result.subtotal, eligibilityTotal: result.eligibilityTotal,
        employerTotal: result.employerTotal, employeeTotal: result.employeeTotal, totalBilled: result.totalBilled,
        flags: result.flags, unmatched: result.unmatched, reconciled: result.reconciled,
        savedFolder: saved ? folderPath : null, savedName: saved ? saved.original_name : null,
      };
      res.setHeader('X-Alloc-Summary', encodeURIComponent(JSON.stringify(summary)));
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
      return res.send(wbBuf);
    } catch (e) {
      console.error('[insurance-allocation] error:', e.message);
      return res.status(400).json({ error: 'Could not build the allocation: ' + e.message });
    }
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

// Single journal entry by id, with lines + attachments (same shape as the list).
// Lets any drilldown that only holds an entry id (e.g. AP Aging GL rows) open the
// JE modal without needing to prefetch the full entry.
app.get('/api/entities/:eid/entries/:id', auth, requireEntityAccess(), (req, res) => {
  const e = db.prepare('SELECT * FROM journal_entries WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
  if (!e) return res.status(404).json({ error: 'Entry not found' });
  const lines = db.prepare(`SELECT jl.*, dp.name AS project_name, dp.code AS project_code,
      dc.name AS class_name, dl.name AS location_name
    FROM journal_lines jl
    LEFT JOIN dim_projects dp ON dp.id = jl.project_id
    LEFT JOIN dim_classes dc ON dc.id = jl.class_id
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    WHERE jl.entry_id = ?`).all(e.id);
  const attachments = db.prepare('SELECT id, original_name, mime_type, size FROM journal_attachments WHERE entry_id = ?').all(e.id);
  res.json({ ...e, lines, attachments });
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
  if (project_id) { where += ' AND CAST(jl.project_id AS REAL) = CAST(? AS REAL)'; params.push(project_id); }
  if (account_code) { where += ' AND jl.account_code = ?'; params.push(account_code); }
  const rows = db.prepare(`
    SELECT jl.id AS line_id, je.id AS entry_id, je.entry_num, je.date, je.memo,
           je.doc_number, je.vendor,
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
  // Opening balances: when a 'from' date is given, each account's running
  // balance must START at its cumulative balance as of the day BEFORE 'from'
  // (all prior activity), not at 0 — otherwise a mid-life date-range GL detail
  // for a balance-sheet account (e.g. cash 1/1/25–12/31/25) wrongly opens at 0
  // instead of the 12/31/24 ending balance. Same dimension filters apply so the
  // opening ties to the windowed activity. With no 'from' (inception-to-date),
  // opening is 0, which is already correct.
  const opening = new Map();
  if (from) {
    const oParams = [req.params.eid];
    let oWhere = ' AND je.date < ?'; oParams.push(from);
    if (location_id) { oWhere += ' AND jl.location_id = ?'; oParams.push(location_id); }
    if (class_id) { oWhere += ' AND jl.class_id = ?'; oParams.push(class_id); }
    if (project_id) { oWhere += ' AND jl.project_id = ?'; oParams.push(project_id); }
    if (account_code) { oWhere += ' AND jl.account_code = ?'; oParams.push(account_code); }
    const oRows = db.prepare(`
      SELECT jl.account_code, a.type AS account_type,
             SUM(jl.debit) AS td, SUM(jl.credit) AS tc
      FROM journal_lines jl
      JOIN journal_entries je ON je.id = jl.entry_id
      LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
      WHERE je.entity_id = ?${oWhere}
      GROUP BY jl.account_code
    `).all(...oParams);
    for (const r of oRows) {
      const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
      const bal = isDr ? ((r.td || 0) - (r.tc || 0)) : ((r.tc || 0) - (r.td || 0));
      opening.set(r.account_code, +bal.toFixed(2));
    }
  }
  // Running balance per account (natural side: Asset/Expense are debit-positive),
  // seeded with the opening balance so date-range reports carry forward.
  const run = new Map();
  const out = rows.map(r => {
    const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
    const delta = isDr ? (r.debit - r.credit) : (r.credit - r.debit);
    const bal = (run.has(r.account_code) ? run.get(r.account_code) : (opening.get(r.account_code) || 0)) + delta;
    run.set(r.account_code, bal);
    return {
      line_id: r.line_id, entry_id: r.entry_id, entry_num: r.entry_num, date: r.date, memo: r.memo,
      account_code: r.account_code, account_name: r.account_name, account_type: r.account_type,
      debit: +(r.debit || 0).toFixed(2), credit: +(r.credit || 0).toFixed(2),
      description: r.description || '', class_name: r.class_name || '', location_name: r.location_name || '',
      // project_name / project_code were selected but never returned, so every
      // downstream "Project" column (TB GL Detail export, Custom Detail
      // group-by-project) read blank. CLA item 6, 8/2026.
      project_name: r.project_name || '', project_code: r.project_code || '',
      doc_number: r.doc_number || '', vendor: r.vendor || '',
      running_balance: +bal.toFixed(2),
    };
  });
  const totalDr = +out.reduce((s, r) => s + r.debit, 0).toFixed(2);
  const totalCr = +out.reduce((s, r) => s + r.credit, 0).toFixed(2);
  res.json({ lines: out, count: out.length, total_debit: totalDr, total_credit: totalCr, opening_balances: Object.fromEntries(opening) });
});

app.post('/api/entities/:eid/entries', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { date, memo, lines, doc_number } = req.body; if (!date||!memo||!lines||lines.length<2) return res.status(400).json({ error: 'Invalid' });
  const tDr = lines.reduce((s,l) => s+(l.debit||0), 0); const tCr = lines.reduce((s,l) => s+(l.credit||0), 0);
  if (Math.abs(tDr-tCr) > 0.005) return res.status(400).json({ error: 'Must balance' });
  const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
  const result = db.transaction(() => {
    const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, doc_number, created_by) VALUES (?,?,?,?,?,?)').run(req.params.eid, num, date, memo, (doc_number || '').trim() || null, req.user.name);
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

// Backfill journal_entries.doc_number for entries created before the column
// existed. Source of truth is billcom_sync_log.invoice_number; the memo pattern
// "Bill.com bill #<n>" is the fallback for rows the log does not cover. Body:
// { dry_run?: true }. Idempotent - only fills rows where doc_number IS NULL.
app.post('/api/entities/:eid/backfill-doc-numbers', auth, requireEntityAccess(), requireRole('Admin'), (req, res) => {
  const eid = req.params.eid;
  const dryRun = !!(req.body && req.body.dry_run);
  const fromLog = db.prepare(
    "SELECT je.id, je.entry_num, bl.invoice_number AS doc " +
    "FROM billcom_sync_log bl JOIN journal_entries je ON je.id = bl.cl_entry_id " +
    "WHERE bl.entity_id = ? AND bl.status = 'success' AND bl.invoice_number IS NOT NULL " +
    "AND bl.invoice_number <> '' AND (je.doc_number IS NULL OR je.doc_number = '')"
  ).all(eid);
  const fromMemo = db.prepare(
    "SELECT id, entry_num, memo FROM journal_entries " +
    "WHERE entity_id = ? AND (doc_number IS NULL OR doc_number = '') AND memo LIKE 'Bill.com bill #%'"
  ).all(eid);
  const seen = new Set();
  const plan = [];
  for (const r of fromLog) { if (seen.has(r.id)) continue; seen.add(r.id); plan.push({ id: r.id, entry_num: r.entry_num, doc_number: String(r.doc).trim(), source: 'sync_log' }); }
  for (const r of fromMemo) {
    if (seen.has(r.id)) continue;
    const m = String(r.memo || '').match(/^Bill\.com bill #(.+)$/);
    if (!m) continue;
    seen.add(r.id);
    plan.push({ id: r.id, entry_num: r.entry_num, doc_number: m[1].trim(), source: 'memo' });
  }
  if (!dryRun && plan.length) {
    const upd = db.prepare('UPDATE journal_entries SET doc_number = ? WHERE id = ? AND entity_id = ?');
    db.transaction(() => { for (const p of plan) upd.run(p.doc_number, p.id, eid); })();
  }
  res.json({
    dry_run: dryRun,
    would_update: plan.length,
    updated: dryRun ? 0 : plan.length,
    from_sync_log: plan.filter(p => p.source === 'sync_log').length,
    from_memo: plan.filter(p => p.source === 'memo').length,
    sample: plan.slice(0, 25),
  });
});

app.delete('/api/entities/:eid/entries/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const atts = db.prepare('SELECT filename FROM journal_attachments WHERE entry_id=?').all(req.params.id);
  atts.forEach(a => { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} });
  db.prepare('DELETE FROM journal_entries WHERE id=? AND entity_id=?').run(req.params.id, req.params.eid);
  res.json({ success: true });
});

app.put('/api/entities/:eid/entries/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { date, memo, lines, doc_number } = req.body;
  if (!date || !memo || !lines || lines.length < 2) return res.status(400).json({ error: 'Invalid entry' });
  const tDr = lines.reduce((s, l) => s + (l.debit || 0), 0);
  const tCr = lines.reduce((s, l) => s + (l.credit || 0), 0);
  if (Math.abs(tDr - tCr) > 0.005) return res.status(400).json({ error: 'Must balance' });
  const entry = db.prepare('SELECT * FROM journal_entries WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!entry) return res.status(404).json({ error: 'Entry not found' });
  db.transaction(() => {
    // doc_number is only touched when the caller sends the key, so an older client
    // that does not know about the field cannot blank an existing document number.
    if (Object.prototype.hasOwnProperty.call(req.body, 'doc_number')) {
      db.prepare("UPDATE journal_entries SET date=?, memo=?, doc_number=?, updated_by=?, updated_at=datetime('now') WHERE id=?")
        .run(date, memo, String(doc_number || '').trim() || null, req.user.name || req.user.email, req.params.id);
    } else {
      db.prepare("UPDATE journal_entries SET date=?, memo=?, updated_by=?, updated_at=datetime('now') WHERE id=?")
        .run(date, memo, req.user.name || req.user.email, req.params.id);
    }
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

// ── OCR fallback for image-only bank statement PDFs ─────────────────────────
// Some statements have NO embedded text layer — scanned copies, or bank portals
// (like MapleMark) that export the operating-account statement as a page image.
// pdf-parse returns almost nothing for those. Rasterize each page with pdftoppm
// and OCR it with tesseract (both baked into the Docker image), then rebuild
// text lines from the word bounding boxes so a row like "7/02 7,408.05 Bill.com
// ..." comes back on ONE line and downstream parsing sees normal transactions.
function ocrPdfToLines(buffer) {
  const cp = require('child_process');
  const osMod = require('os');
  const dir = fs.mkdtempSync(path.join(osMod.tmpdir(), 'clocr-'));
  // Total wall-clock budget for the whole OCR pass. Railway hard-kills a request
  // at ~300s (returning its own HTML 502, which the client can't parse as JSON);
  // stay well under that so we return a clean JSON error instead of timing out.
  const OCR_BUDGET_MS = 200000;
  const startedAt = Date.now();
  const remaining = () => OCR_BUDGET_MS - (Date.now() - startedAt);
  try {
    const pdfPath = path.join(dir, 'in.pdf');
    fs.writeFileSync(pdfPath, buffer);
    // 200 DPI is enough for statement text and uses ~40% less time/memory than
    // 300 DPI, reducing the risk of an OOM/timeout on a multi-page scanned PDF.
    cp.execFileSync('pdftoppm', ['-r', '200', '-png', pdfPath, path.join(dir, 'pg')], { stdio: 'ignore', timeout: Math.max(15000, remaining()), maxBuffer: 64 * 1024 * 1024 });
    const pngs = fs.readdirSync(dir).filter(f => /^pg.*\.png$/.test(f)).sort();
    const allLines = [];
    for (const png of pngs) {
      if (remaining() < 10000) throw new Error('OCR exceeded time budget after ' + allLines.length + ' rows across ' + pngs.length + ' page(s); the scanned PDF is too large/slow to process in time.');
      const base = path.join(dir, png.replace(/\.png$/, ''));
      cp.execFileSync('tesseract', [path.join(dir, png), base, 'tsv'], { stdio: 'ignore', timeout: Math.max(10000, remaining()), maxBuffer: 64 * 1024 * 1024 });
      const tsv = fs.readFileSync(base + '.tsv', 'utf8');
      const words = [];
      for (const row of tsv.split(/\r?\n/).slice(1)) {
        const c = row.split('\t');
        if (c.length < 12) continue;
        const conf = parseFloat(c[10]); const txt = (c[11] || '').trim();
        if (!txt || !(conf > 0)) continue;
        words.push({ left: +c[6], top: +c[7], text: txt });
      }
      words.sort((a, b) => a.top - b.top || a.left - b.left);
      // Group words into visual rows by y-position (tolerance ~ half a line at 300 DPI).
      const YTOL = 14; let cur = [], cy = null;
      const flush = () => { if (cur.length) { cur.sort((a, b) => a.left - b.left); allLines.push(cur.map(w => w.text).join(' ')); } cur = []; };
      for (const w of words) {
        if (cy === null || Math.abs(w.top - cy) <= YTOL) { cur.push(w); if (cy === null) cy = w.top; }
        else { flush(); cur = [w]; cy = w.top; }
      }
      flush();
    }
    return { lines: allLines, text: allLines.join('\n') };
  } finally {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) {}
  }
}

// ── MapleMark "Platinum Money Market" operating-account statement parser ─────
// Runs on OCR-recovered lines. Two sections: "Deposits and Other Credits"
// (inflow) and "Debits and Other Withdrawals" (outflow); each transaction row
// starts with an M/D date and an amount, and its description can continue on the
// following lines. Stop at the "Daily Balance Summary" table. Reconcile against
// the Summary-of-Activity control totals ("Deposits / Misc Credits <count>
// <total>", "Withdrawals / Misc Debits <count> <total>") plus the beginning/
// ending balance identity; if it doesn't tie out, flag every row for review.
function parseMapleMarkMoneyMarket(lines, text) {
  const rows = [];
  const _toNum = s => parseFloat(String(s).replace(/[^0-9.]/g, '')) || 0;
  // Statement year from a full M/D/YY date in the Summary of Activity block.
  const ym = text.match(/(?:Beginning|Ending)\s*Balance\s*\d{1,2}\/\d{1,2}\/(\d{2})/i);
  const year = ym ? 2000 + (+ym[1]) : (new Date()).getFullYear();
  // Control totals.
  const _cd = text.match(/Deposits\s*\/\s*Misc\s*Credits\s+(\d+)\s+([\d,]+\.\d{2})/i);
  const _cw = text.match(/Withdrawals\s*\/\s*Misc\s*Debits\s+(\d+)\s+([\d,]+\.\d{2})/i);
  const _cb = text.match(/Beginning\s*Balance\s*\d{1,2}\/\d{1,2}\/\d{2}\s+([\d,]+\.\d{2})/i);
  const _ce = text.match(/Ending\s*Balance\s*\d{1,2}\/\d{1,2}\/\d{2}\s+([\d,]+\.\d{2})/i);
  const ctrlDepN = _cd ? +_cd[1] : null, ctrlDepTot = _cd ? _toNum(_cd[2]) : null;
  const ctrlDebN = _cw ? +_cw[1] : null, ctrlDebTot = _cw ? _toNum(_cw[2]) : null;
  const ctrlBeg = _cb ? _toNum(_cb[1]) : null, ctrlEnd = _ce ? _toNum(_ce[1]) : null;

  const dateRx = /^(\d{1,2})\/(\d{1,2})\b/;
  const moneyRx = /^\$?[\d,]+\.\d{2}$/;
  // Page header/footer furniture that repeats at every page break — must never
  // be folded into a transaction's description as a continuation line.
  const noiseRx = /^m?\s*maple\b|^mark$|^bank$|banyan residential llc|rosecrans|el segundo|^page\b|page\s+\d+\s+of\s+\d+|account number|^date\b|platinum money market|contains confidential|member\s*fdic|service mark|intrafi|^\*{2,}\d+|^\d+\s+of\s+\d+$|activity description/i;
  let section = 0; // +1 deposits, -1 debits, 0 none
  let cur = null;
  const push = () => { if (cur && cur.amount !== 0) rows.push(cur); cur = null; };

  for (const raw of lines) {
    const line = raw.replace(/\s+/g, ' ').trim();
    if (/deposits and other credits/i.test(line)) { push(); section = 1; continue; }
    if (/debits and other withdrawals/i.test(line)) { push(); section = -1; continue; }
    if (/daily balance summary|summary of activity/i.test(line)) { push(); section = 0; continue; }
    if (section === 0) continue;
    if (/^date\s+amount\b/i.test(line) || /^amount\b/i.test(line)) continue; // column header
    const dm = line.match(dateRx);
    if (dm) {
      const toks = line.split(' ');
      let amtIdx = -1;
      for (let i = 1; i < toks.length; i++) { if (moneyRx.test(toks[i])) { amtIdx = i; break; } }
      if (amtIdx < 0) continue; // date line without an amount — skip
      push();
      const mag = _toNum(toks[amtIdx]);
      const desc = toks.slice(amtIdx + 1).join(' ').trim();
      const d = new Date(year, +dm[1] - 1, +dm[2]);
      cur = { date: isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10),
        description: desc, amount: section === -1 ? -mag : mag };
      if (!cur.date) cur = null;
    } else if (cur && !noiseRx.test(line)) {
      cur.description = (cur.description + ' ' + line).replace(/\s+/g, ' ').trim();
    }
  }
  push();
  rows.forEach(r => { if (!r.description) r.description = r.amount < 0 ? 'Withdrawal' : 'Deposit'; r.description = r.description.substring(0, 500); });

  const EPS = 0.02;
  const depRows = rows.filter(r => r.amount > 0), debRows = rows.filter(r => r.amount < 0);
  const depTot = depRows.reduce((a, r) => a + r.amount, 0);
  const debTot = debRows.reduce((a, r) => a + Math.abs(r.amount), 0);
  const depOK = ctrlDepTot == null || (Math.abs(depTot - ctrlDepTot) < EPS && (ctrlDepN == null || depRows.length === ctrlDepN));
  const debOK = ctrlDebTot == null || (Math.abs(debTot - ctrlDebTot) < EPS && (ctrlDebN == null || debRows.length === ctrlDebN));
  let idOK = true;
  if (ctrlBeg != null && ctrlEnd != null) idOK = Math.abs((ctrlBeg + depTot - debTot) - ctrlEnd) < EPS;
  const haveCtrl = ctrlDepTot != null || ctrlDebTot != null;
  const reconciled = haveCtrl && depOK && debOK && idOK;
  const ctrlInfo = { deposits: ctrlDepTot, checks: ctrlDebTot, prev: ctrlBeg, curr: ctrlEnd,
    deposit_count: ctrlDepN, check_count: ctrlDebN, parsed_deposits: +depTot.toFixed(2), parsed_checks: +debTot.toFixed(2),
    parsed_deposit_count: depRows.length, parsed_check_count: debRows.length };
  if (!reconciled && haveCtrl) rows.forEach(r => { r.needs_review = true; });
  return { rows, reconciled, ctrlInfo };
}

// ── Coordinate-aware text extraction for column-based statements ────────────
// pdf-parse flattens a PDF to plain text and loses the horizontal position of
// each number, so on a statement with separate Deposits and Withdrawals columns
// you can't tell which column an amount sits in. pdfjs exposes each text item's
// x-position, so we rebuild per-page visual lines whose words keep their x —
// enough to classify an amount by column. Returns pages[] of lines[] of
// { x, x1, t }. Lazy-required so a pdfjs issue never blocks server startup.
async function extractPdfPositionedPages(buffer) {
  const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buffer), isEvalSupported: false, verbosity: 0 }).promise;
  const pages = [];
  for (let pg = 1; pg <= doc.numPages; pg++) {
    const page = await doc.getPage(pg);
    const tc = await page.getTextContent();
    const items = tc.items.map(it => ({ x: it.transform[4], y: it.transform[5], w: it.width || 0, t: (it.str || '').trim() })).filter(i => i.t);
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = []; let cur = [], cy = null;
    for (const it of items) {
      if (cy === null || Math.abs(it.y - cy) <= 3) { cur.push(it); cy = cy === null ? it.y : cy; }
      else { lines.push(cur); cur = [it]; cy = it.y; }
    }
    if (cur.length) lines.push(cur);
    pages.push(lines.map(l => l.sort((a, b) => a.x - b.x).map(w => ({ x: w.x, x1: w.x + w.w, y: w.y, t: w.t }))));
  }
  return pages;
}

// ── UMB Bank "Commercial Checking" statement parser ─────────────────────────
// A two-column statement (Deposits | Withdrawals). Each Transaction Detail row
// linearizes to a bare amount, so we classify it by which column its x-position
// falls under (from the Deposits/Withdrawals header centers). Reconcile against
// the Account Summary (Deposits and Credits, Withdrawals and Debits, Service
// Charges and Fees) and the beginning/ending balance identity.
function parseUMB(pages, flatText) {
  const rows = [];
  const _toNum = s => parseFloat(String(s).replace(/[^0-9.]/g, '')) || 0;
  const moneyRx = /^\$?[\d,]+\.\d{2}$/;
  const MONTHS = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  const ym = flatText.match(/Ending Balance as of\s*\d{1,2}\/\d{1,2}\/(\d{4})/i);
  const year = ym ? +ym[1] : (new Date()).getFullYear();
  const g = rx => { const m = flatText.match(rx); return m ? _toNum(m[1]) : null; };
  const ctrlBeg = g(/Beginning Balance as of\s*\d{1,2}\/\d{1,2}\/\d{4}\s*\$?([\d,]+\.\d{2})/i);
  const ctrlEnd = g(/Ending Balance as of\s*\d{1,2}\/\d{1,2}\/\d{4}\s*\$?([\d,]+\.\d{2})/i);
  const _dm = flatText.match(/Deposits and Credits\s*\((\d+)\)\s*\$?([\d,]+\.\d{2})/i);
  const _wm = flatText.match(/Withdrawals and Debits\s*\((\d+)\)\s*\$?([\d,]+\.\d{2})/i);
  const ctrlDepN = _dm ? +_dm[1] : null, ctrlDep = _dm ? _toNum(_dm[2]) : null;
  const ctrlWd = _wm ? _toNum(_wm[2]) : null;
  const ctrlFees = g(/Service Charges and Fees\s*\$?([\d,]+\.\d{2})/i);
  const expOut = (ctrlWd || 0) + (ctrlFees || 0);
  const expIn = ctrlDep || 0;

  let inDetail = false, depC = null, wdC = null, cur = null;
  const noiseRx = /^(hp property owner|e 64th|aurora|statement (ending|period)|page \d|return service|account (number|title|summary)|commercial|mailstop|p\.o\. box|kansas city|terms and conditions|go paperless|contact information|in case of|beginning balance|ending balance|deposits and|withdrawals and|service charges|total days)/i;
  const push = () => { if (cur && cur.amount !== 0) rows.push(cur); cur = null; };

  for (const lines of pages) {
    for (const words of lines) {
      const joined = words.map(w => w.t).join(' ').replace(/\s+/g, ' ').trim();
      if (/^Transaction Detail/i.test(joined)) { push(); inDetail = true; depC = wdC = null; continue; }
      if (/End of Day|Current Balance|^Totals\b/i.test(joined)) { push(); inDetail = false; continue; }
      if (!inDetail) continue;
      const dh = words.find(w => /^Deposits$/i.test(w.t)); const wh = words.find(w => /^Withdrawals$/i.test(w.t));
      if (dh || wh) { if (dh) depC = (dh.x + dh.x1) / 2; if (wh) wdC = (wh.x + wh.x1) / 2; continue; }
      if (/^Date\b/i.test(joined) || /^Description/i.test(joined)) continue;
      const dtMatch = joined.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\b/i);
      if (dtMatch) {
        const amtWord = [...words].reverse().find(w => moneyRx.test(w.t));
        if (!amtWord) continue;
        push();
        const mag = _toNum(amtWord.t);
        const center = (amtWord.x + amtWord.x1) / 2;
        let sign;
        if (depC != null && wdC != null) sign = Math.abs(center - wdC) <= Math.abs(center - depC) ? -1 : 1;
        else if (wdC != null) sign = center >= wdC - 30 ? -1 : 1;
        else sign = -1;
        const d = new Date(year, MONTHS[dtMatch[1].toLowerCase()], +dtMatch[2]);
        let desc = joined.slice(dtMatch[0].length).replace(amtWord.t, '').replace(/\s+/g, ' ').trim();
        cur = { date: isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10), description: desc, amount: sign * mag };
        if (!cur.date) cur = null;
      } else if (cur && !noiseRx.test(joined) && !/^[\d,]+\.\d{2}$/.test(joined)) {
        cur.description = (cur.description + ' ' + joined).replace(/\s+/g, ' ').trim();
      }
    }
  }
  push();
  rows.forEach(r => { if (!r.description) r.description = r.amount < 0 ? 'Withdrawal' : 'Deposit'; r.description = r.description.substring(0, 500); });

  const EPS = 0.02;
  const depRows = rows.filter(r => r.amount > 0), wdRows = rows.filter(r => r.amount < 0);
  const depTot = depRows.reduce((a, r) => a + r.amount, 0);
  const wdTot = wdRows.reduce((a, r) => a + Math.abs(r.amount), 0);
  const depOK = ctrlDep == null || Math.abs(depTot - expIn) < EPS;
  const wdOK = (ctrlWd == null && ctrlFees == null) || Math.abs(wdTot - expOut) < EPS;
  let idOK = true;
  if (ctrlBeg != null && ctrlEnd != null) idOK = Math.abs((ctrlBeg + depTot - wdTot) - ctrlEnd) < EPS;
  const haveCtrl = ctrlDep != null || ctrlWd != null || ctrlFees != null;
  const reconciled = haveCtrl && depOK && wdOK && idOK;
  const ctrlInfo = { deposits: expIn, checks: expOut, prev: ctrlBeg, curr: ctrlEnd,
    deposit_count: ctrlDepN, parsed_deposits: +depTot.toFixed(2), parsed_checks: +wdTot.toFixed(2),
    parsed_deposit_count: depRows.length, parsed_check_count: wdRows.length };
  if (!reconciled && haveCtrl) rows.forEach(r => { r.needs_review = true; });
  return { rows, reconciled, ctrlInfo };
}

// ── UBS "Business Services Account" brokerage statement parser ──────────────
// This is a sweep/brokerage account, not a checking register — its only monthly
// cash activity is dividend & interest income. Per Jimmy's decision, record that
// income as ONE credit transaction dated the statement-period end.
//
// We read the MONTHLY figure from the "Cash activity summary" table, taking the
// value in the period column ("<Month> <Year> ($)") — NOT the "Year to date ($)"
// column. An earlier version keyed on the page-1 "Sources of your account growth
// during <Year>" box, whose header also reads "Dividend and interest income" /
// "Change in market value"; that box carries the YEAR-TO-DATE income, so a month
// like July 2026 (period $0.06, YTD $0.90) imported $0.90. The YTD identity
// (year-end value + YTD income = closing) happens to tie, which is why the bug
// passed reconciliation. Reconciling on the PERIOD identity
// (period opening + period income = period closing) rejects the YTD figure.
//
// The statement is rotated 90°, so pdf.js reports the label as the row key (its
// x) and the period-vs-YTD column as the token y: each money value shares its
// label's x and sits under either the period header or the YTD header by y.
function parseUBS(pages, flatText) {
  const _num = s => parseFloat(String(s).replace(/[^0-9.]/g, '')) || 0;
  const norm = flatText.replace(/\s+/g, ' ');
  const MONTHS = { january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11 };
  const pm = norm.match(/Business Services Account\s+([A-Za-z]+)\s+(\d{4})/i);
  let year = pm ? +pm[2] : (new Date()).getFullYear();
  let month = pm ? MONTHS[pm[1].toLowerCase()] : null;
  if (month == null) month = 0;
  const day = new Date(year, month + 1, 0).getDate();  // statement-period end
  const date = new Date(year, month, day);
  const dateStr = isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
  const monthName = Object.keys(MONTHS).find(k => MONTHS[k] === month);

  const moneyRx = /^\$?\(?[\d,]+\.\d{2}\)?$/;
  const periodHdrRx = new RegExp(`${monthName}\\s+${year}\\s*\\(\\$\\)`, 'i'); // "July 2026 ($)"
  const ytdHdrRx = /Year to date\s*\(\$\)/i;

  let income = null, opening = null, closing = null, netcash = null;
  for (const lines of pages) {
    const toks = [];
    for (const ln of lines) for (const w of ln) toks.push(w);
    if (!toks.some(w => /Cash activity summary/i.test(w.t))) continue;

    // Period ("<Month> <Year> ($)") header tokens for this block. There may be
    // one in the Change-in-value box and one in Cash-activity; matching on the
    // header y (below) selects the right column regardless of which we anchor on.
    const periodHdrs = toks.filter(w => periodHdrRx.test(w.t));
    if (!periodHdrs.length) continue;

    // Cash-activity row labels. Each label's value shares the label's x; the
    // period value is the money token at that x whose y is nearest a period
    // header's y (vs the YTD header's y).
    const labelTok = rx => toks.find(w => rx.test(w.t));
    const openLbl  = labelTok(/^Opening balances$/i);
    const divLbl   = labelTok(/^Dividend and interest income$/i);
    const netLbl   = labelTok(/^Net cash flow$/i);
    const closeLbl = labelTok(/^Closing balances$/i);
    if (!divLbl) continue;

    const XTOL = 4;
    const valueAtLabel = (lbl, hdrs) => {
      if (!lbl || !hdrs.length) return null;
      const col = toks.filter(w => moneyRx.test(w.t) && Math.abs(w.x - lbl.x) <= XTOL);
      if (!col.length) return null;
      let best = null, bestD = Infinity;
      for (const m of col) for (const h of hdrs) {
        const d = Math.abs(m.y - h.y);
        if (d < bestD) { bestD = d; best = m; }
      }
      return best ? _num(best.t) : null;
    };

    opening = valueAtLabel(openLbl,  periodHdrs);
    income  = valueAtLabel(divLbl,   periodHdrs);
    netcash = valueAtLabel(netLbl,   periodHdrs);
    closing = valueAtLabel(closeLbl, periodHdrs);
    break;
  }

  const rows = [];
  if (income != null && dateStr) rows.push({ date: dateStr, description: 'UBS dividend and interest income', amount: income });
  const EPS = 0.02;
  let reconciled = false;
  // Period identity (cash sweep: no market-value change). Falls back to the
  // Net cash flow line, which equals income for a cash-only account.
  if (opening != null && closing != null && income != null) reconciled = Math.abs((opening + income) - closing) < EPS;
  if (!reconciled && netcash != null && income != null) reconciled = Math.abs(netcash - income) < EPS;
  if (!reconciled) rows.forEach(r => { r.needs_review = true; });
  const ctrlInfo = { deposits: income, checks: 0, prev: opening, curr: closing, net_cash_flow: netcash,
    parsed_deposits: income, parsed_checks: 0 };
  return { rows, reconciled, ctrlInfo };
}

// ═══ Bank Transaction Upload & Coding ═══

// Find the best active coding note for a parsed statement row. Matching keys:
// signed amount within tolerance, date inside [date_from,date_to] (open ends
// allowed), and — if the note carries a desc_keyword — a case-insensitive
// substring hit on the description. When several notes match, the tightest
// amount tolerance wins (most specific), then the smallest date window, then
// oldest note. one_shot notes already at matched_count>0 are skipped.
function findCodingNote(entityId, bankAccount, row) {
  const notes = db.prepare(
    `SELECT * FROM bank_coding_notes
      WHERE entity_id=? AND active=1
        AND (bank_account_code IS NULL OR bank_account_code=?)
        AND (one_shot=0 OR matched_count=0)`
  ).all(entityId, bankAccount);
  const candidates = notes.filter(n => {
    const tol = Number(n.amount_tolerance) || 0;
    if (Math.abs(Number(row.amount) - Number(n.match_amount)) > tol + 1e-6) return false;
    if (n.date_from && row.date < n.date_from) return false;
    if (n.date_to && row.date > n.date_to) return false;
    if (n.desc_keyword && !(row.description || '').toLowerCase().includes(n.desc_keyword.toLowerCase())) return false;
    return true;
  });
  if (!candidates.length) return null;
  const span = n => (n.date_from && n.date_to)
    ? (Date.parse(n.date_to) - Date.parse(n.date_from)) : Number.MAX_SAFE_INTEGER;
  candidates.sort((a, b) =>
    (Number(a.amount_tolerance) - Number(b.amount_tolerance)) || (span(a) - span(b)) || (a.id - b.id));
  return candidates[0];
}

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
    let reconciled = false; // set true if a PDF parse tied to statement control totals
    let ctrlInfo = null;    // {deposits, checks, prev, curr} when available

    if (isPdf) {
      // ── PDF bank statement parsing ──
      const data = await pdfParse(req.file.buffer);
      let text = data.text || '';
      let lines = text.split(/\n/).map(l => l.trim()).filter(Boolean);
      // Image-only PDF (no text layer): pdf-parse yields (almost) nothing. Fall
      // back to OCR to recover the text and per-row line layout.
      if (lines.length < 5) {
        try {
          const ocr = ocrPdfToLines(req.file.buffer);
          if (ocr.lines.length > lines.length) { lines = ocr.lines; text = ocr.text; }
        } catch (e) {
          console.error('[bank upload] OCR fallback failed:', e.message);
          // Surface a clear, actionable JSON error rather than falling through to
          // a generic "no transaction rows" message for a scanned/image PDF.
          return res.status(422).json({ error: 'This looks like a scanned (image-only) PDF and it could not be read in time: ' + e.message + ' Try uploading a text-based PDF export from the bank portal, or a CSV/Excel file if available.' });
        }
      }

      // Bank of Texas (BOKF) statements linearize through pdf-parse with the
      // date, the description, and the amount each on their OWN line, grouped
      // under DEPOSITS / WITHDRAWALS / CHECKS section headers. The generic
      // single-line parser below can't read that shape (it lands on the bare
      // amount line and takes the amount string as the description, and the
      // MM-DD row dates carry no year). Route Bank of Texas statements to a
      // dedicated block parser, detected by the bank's own branding in the text.
      const isBOT = /bank of texas|bankoftexas\.com|\bBOKF\b/i.test(text);
      // MapleMark Bank's IntraFi Cash Service (ICS) statement has a different
      // shape than a normal operating-account statement: a single "Account
      // Transaction Detail" table where every row carries the activity Amount AND
      // a running Balance (two money tokens per row), withdrawals are shown in
      // parentheses, and there are no Deposits/Withdrawals section headers. The
      // generic parser below would grab the balance (the LAST token) as the
      // amount and would also mis-read the Account Summary and per-bank "Summary
      // of Balances" figures as transactions — so ICS gets its own parser.
      const isICS = /intrafi|maplemark|\bICS\b/i.test(text) && /account\s*transaction\s*detail/i.test(text);
      // MapleMark "Platinum Money Market" operating-account statement (image PDF,
      // read via OCR above): two Deposits/Debits sections, not the ICS layout.
      // Detect on the Money Market template markers, NOT on both section headers
      // being present: a month with zero withdrawals has no "Debits and Other
      // Withdrawals" section (and a zero-deposit month has no credits section),
      // so require only MapleMark + the Money Market/Summary-of-Activity header
      // and at least one of the two transaction sections.
      const isMMM = /maple\s*mark/i.test(text)
        && (/platinum money market/i.test(text) || /summary of activity since your last statement/i.test(text))
        && (/deposits and other credits/i.test(text) || /debits and other withdrawals/i.test(text));
      // UMB Bank Commercial Checking: two-column (Deposits | Withdrawals) detail,
      // classified by column position via a coordinate-aware re-read below.
      const isUMB = /transaction detail/i.test(text) && /end of day\s*-?\s*current balance/i.test(text) && /deposits and credits/i.test(text);
      // UBS brokerage statement — record the month's dividend/interest income as
      // one credit. Detected on whitespace-normalized text (UBS statements
      // linearize one word per line through pdf-parse).
      const _norm = text.replace(/\s+/g, ' ');
      const isUBS = /UBS Financial Services/i.test(_norm) && /Business Services Account/i.test(_norm) && /Dividend and\s*interest income/i.test(_norm);
      if (isBOT) {
        // Group each date -> description(s) -> amount block into one transaction;
        // section headers set the sign; stop before the DAILY ACCOUNT BALANCE
        // table and the balancing worksheet; then verify the parse against the
        // statement's own "N Deposits" / "N Checks & Withdrawals" control totals
        // (both the item COUNT and the dollar amount).
        const _y4 = y => { y = String(y); return y.length === 2 ? (+y > 50 ? '19' : '20') + y : y; };
        const _perM = text.match(/statement period[^0-9]*?(\d{1,2})-(\d{1,2})-(\d{2,4})\s*(?:to|through|-)\s*(\d{1,2})-(\d{1,2})-(\d{2,4})/i);
        let _startY = null, _startMo = null, _endY = null;
        if (_perM) { _startMo = +_perM[1]; _startY = +_y4(_perM[3]); _endY = +_y4(_perM[6]); }
        // MM-DD rows get their year from the statement period. If the period
        // straddles a year boundary (Dec->Jan), split by month.
        const _yearFor = mo => { if (_startY == null) return (new Date()).getFullYear(); if (_startY === _endY) return _startY; return mo >= _startMo ? _startY : _endY; };
        const _mkDate = (dd, mo) => { const y = _yearFor(mo); const d = new Date(y, mo - 1, dd); return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10); };
        const _toNum = s => parseFloat(String(s).replace(/[,$]/g, '')) || 0;
        const _depM = text.match(/\+?\s*(\d+)\s*Deposits\D{0,4}(\d[\d,]*\.\d{2})/i);
        const _chkM = text.match(/-?\s*(\d+)\s*Checks?\s*&\s*Withdrawals\D{0,4}(\d[\d,]*\.\d{2})/i);
        const _ctrlDep = _depM ? _toNum(_depM[2]) : null, _ctrlDepN = _depM ? +_depM[1] : null;
        const _ctrlChk = _chkM ? _toNum(_chkM[2]) : null, _ctrlChkN = _chkM ? +_chkM[1] : null;
        const _dateRx = /^(\d{1,2})[-\/](\d{1,2})$/;               // "07-01" (whole line, no year)
        const _amtRx = /^\$?\(?-?[\d,]*\.\d{2}-?\)?$/;             // "128,901.87", "239.50", ".01"
        const _isAmt = l => _amtRx.test(l) && /\d/.test(l);
        const _parseAmt = l => { const neg = /^\(.*\)$/.test(l) || /-$/.test(l); const v = Math.abs(parseFloat(l.replace(/[^0-9.]/g, '')) || 0); return neg ? -v : v; };
        let _section = 0, _cur = null; // section: +1 deposits, -1 withdrawals/checks, 0 none/stop
        const _flush = amt => {
          if (!_cur || amt === 0) { _cur = null; return; } // no amount -> discard incomplete block
          const date = _mkDate(_cur.dd, _cur.mo);
          if (!date) { _cur = null; return; }
          let desc = _cur.desc.join(' ').replace(/\s+/g, ' ').trim();
          if (!desc) desc = '(no description)';
          const amount = _section === -1 ? -Math.abs(amt) : Math.abs(amt);
          rows.push({ date, description: desc.substring(0, 500), amount });
          _cur = null;
        };
        for (const line of lines) {
          if (/^deposits$/i.test(line)) { _flush(0); _section = 1; continue; }
          if (/^withdrawals$/i.test(line)) { _flush(0); _section = -1; continue; }
          if (/^checks\s*(\(|paid|$)/i.test(line)) { _flush(0); _section = -1; continue; } // paper checks paid = outflow
          if (/^(daily\s+account\s+balance|daily\s+balance|service\s+fee\s+balance|balancing\s+your\s+account|average\s+ledger|electronic\s+transfer)/i.test(line)) { _flush(0); _section = 0; continue; }
          if (_section === 0) continue;                            // ignore summary / headers / disclosures
          if (/^date\s*amount$/i.test(line) || /^dateamount$/i.test(line) || /^date\s*balance$/i.test(line) || /^datebalance$/i.test(line)) continue; // column header
          if (/no checks/i.test(line)) continue;
          const dm = line.match(_dateRx);
          if (dm) { _flush(0); _cur = { mo: +dm[1], dd: +dm[2], desc: [] }; continue; }
          if (_isAmt(line)) { if (_cur) _flush(_parseAmt(line)); continue; } // stray amount w/o a date block -> ignore
          if (_cur) _cur.desc.push(line);                          // description line
        }
        _flush(0);
        const _depTot = rows.filter(r => r.amount > 0).reduce((a, r) => a + r.amount, 0);
        const _chkTot = rows.filter(r => r.amount < 0).reduce((a, r) => a + Math.abs(r.amount), 0);
        const _depN = rows.filter(r => r.amount > 0).length, _chkN = rows.filter(r => r.amount < 0).length;
        const _EPS = 0.02;
        const _depOK = _ctrlDep == null || (Math.abs(_depTot - _ctrlDep) < _EPS && (_ctrlDepN == null || _depN === _ctrlDepN));
        const _chkOK = _ctrlChk == null || (Math.abs(_chkTot - _ctrlChk) < _EPS && (_ctrlChkN == null || _chkN === _ctrlChkN));
        reconciled = (_ctrlDep != null || _ctrlChk != null) && _depOK && _chkOK;
        ctrlInfo = { deposits: _ctrlDep, checks: _ctrlChk, prev: null, curr: null,
          parsed_deposits: +_depTot.toFixed(2), parsed_checks: +_chkTot.toFixed(2),
          deposit_count: _ctrlDepN, check_count: _ctrlChkN, parsed_deposit_count: _depN, parsed_check_count: _chkN };
        // If the parse doesn't tie to the statement's own totals, flag every row
        // for manual review rather than importing possibly-wrong figures silently.
        if (!reconciled && (_ctrlDep != null || _ctrlChk != null)) rows.forEach(r => { r.needs_review = true; });
      } else if (isICS) {
        // ── IntraFi Cash Service (ICS) / MapleMark Bank statement ──
        // pdf-parse linearizes each transaction onto its own line, with the
        // inter-word spaces stripped, e.g.:
        //   "07/03/2026Withdrawal($125,107.61)$1,337,507.65"
        //   "07/31/2026Interest Capitalization4,042.821,134,730.68"
        // Every row carries TWO money tokens — the activity Amount THEN the
        // running Balance (which can be glued directly to the amount) — so we
        // take the FIRST money token as the amount and ignore the balance.
        // Withdrawals are parenthesized (outflow); deposits / interest
        // capitalization are bare (inflow). Only rows inside the "Account
        // Transaction Detail" table are transactions: the Account Summary block
        // above it and the per-bank "Summary of Balances" table below it also
        // put money on their lines and must be excluded.
        const _toNum = s => {
          const t = String(s).trim();
          const neg = /^\(.*\)$/.test(t) || /-$/.test(t);
          const v = Math.abs(parseFloat(t.replace(/[^0-9.]/g, '')) || 0);
          return neg ? -v : v;
        };
        // Control totals from the Account Summary block, for reconciliation.
        // pdf-parse strips the spaces, so the value is glued to the label
        // ("Total Program Deposits0.00"); [^\n\d(-]* skips any separator chars
        // (like a "$") without crossing into the next line or into the number.
        const _grab = rx => { const m = text.match(rx); return m ? _toNum(m[1]) : null; };
        const ctrlDeposits    = _grab(/Total\s*Program\s*Deposits[^\n\d(-]*([\d,]+\.\d{2})/i);
        const ctrlWithdrawals = _grab(/Total\s*Program\s*Withdrawals[^\n\d(-]*\(?([\d,]+\.\d{2})\)?/i);
        const ctrlInterest    = _grab(/Interest\s*Capitali[sz]ed[^\n\d(-]*([\d,]+\.\d{2})/i);
        const ctrlPrev        = _grab(/Previous\s*Period\s*Ending\s*Balance[^\n\d(-]*\$?([\d,]+\.\d{2})/i);
        const ctrlCurr        = _grab(/Current\s*Period\s*Ending\s*Balance[^\n\d(-]*\$?([\d,]+\.\d{2})/i);

        // Restrict scanning to the Account Transaction Detail table.
        const _startIdx = lines.findIndex(l => /account\s*transaction\s*detail/i.test(l));
        const _endIdx   = lines.findIndex(l => /summary\s*of\s*balances/i.test(l));
        const _slice = lines.slice(_startIdx >= 0 ? _startIdx + 1 : 0, _endIdx >= 0 ? _endIdx : lines.length);

        const _dateRx  = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/;        // MM/DD/YYYY at line start (no boundary; text is glued)
        const _moneyRx = /\(?\$?-?[\d,]+\.\d{2}-?\)?/g;
        for (const line of _slice) {
          const dm = line.match(_dateRx);
          if (!dm) continue;                                     // only date-led rows are transactions
          const money = line.match(_moneyRx);
          if (!money || !money.length) continue;
          const amount = _toNum(money[0]);                       // FIRST token = amount; the last token is the balance
          if (amount === 0) continue;
          const d = new Date(+dm[3], +dm[1] - 1, +dm[2]);
          if (isNaN(d.getTime())) continue;
          const date = d.toISOString().slice(0, 10);
          // Description = the activity type between the date and the amount.
          let desc = line.slice(dm[0].length, line.indexOf(money[0])).replace(/\s+/g, ' ').trim();
          if (!desc) desc = amount < 0 ? 'Withdrawal' : 'Deposit';
          rows.push({ date, description: desc.substring(0, 500), amount });
        }

        // Reconcile the parse against the statement's own control totals and the
        // balance identity (previous + inflows - outflows = current). ICS treats
        // interest capitalization separately from program deposits, so expected
        // inflow = Total Program Deposits + Interest Capitalized.
        const _EPS = 0.02;
        const _inflow  = rows.filter(r => r.amount > 0).reduce((a, r) => a + r.amount, 0);
        const _outflow = rows.filter(r => r.amount < 0).reduce((a, r) => a + Math.abs(r.amount), 0);
        const _expIn   = (ctrlDeposits || 0) + (ctrlInterest || 0);
        const _haveCtrl = ctrlWithdrawals != null || ctrlDeposits != null || ctrlInterest != null;
        const _inOK  = (ctrlDeposits == null && ctrlInterest == null) || Math.abs(_inflow - _expIn) < _EPS;
        const _outOK = ctrlWithdrawals == null || Math.abs(_outflow - (ctrlWithdrawals || 0)) < _EPS;
        let _idOK = true;
        if (ctrlPrev != null && ctrlCurr != null) _idOK = Math.abs((ctrlPrev + _inflow - _outflow) - ctrlCurr) < _EPS;
        reconciled = _haveCtrl && _inOK && _outOK && _idOK;
        ctrlInfo = { deposits: ctrlDeposits, checks: ctrlWithdrawals, prev: ctrlPrev, curr: ctrlCurr,
          interest: ctrlInterest, parsed_deposits: +_inflow.toFixed(2), parsed_checks: +_outflow.toFixed(2) };
        // If the parse doesn't tie out, flag every row for manual review rather
        // than importing possibly-wrong figures silently.
        if (!reconciled && _haveCtrl) rows.forEach(r => { r.needs_review = true; });
      } else if (isMMM) {
        const _mmm = parseMapleMarkMoneyMarket(lines, text);
        rows = _mmm.rows; reconciled = _mmm.reconciled; ctrlInfo = _mmm.ctrlInfo;
      } else if (isUMB) {
        const _umb = parseUMB(await extractPdfPositionedPages(req.file.buffer), text);
        rows = _umb.rows; reconciled = _umb.reconciled; ctrlInfo = _umb.ctrlInfo;
      } else if (isUBS) {
        const _ubs = parseUBS(await extractPdfPositionedPages(req.file.buffer), text);
        rows = _ubs.rows; reconciled = _ubs.reconciled; ctrlInfo = _ubs.ctrlInfo;
      } else {

      // ── Control totals for post-parse reconciliation ──
      // Most statements print summary control totals we can check the parse
      // against: "Deposits/Credits 3 3,030,021.64 +" and "Checks/Debits 4
      // 2,825,643.75 -", plus Previous/Current balance. pdf-parse sometimes
      // misplaces the space between a row's trailing digits and its amount by one
      // character, stealing a leading digit INTO the amount (e.g. a ref ending
      // "...072926" + " 22,372.55" comes out "...07292" + "622,372.55"). Because
      // the corrupted token is itself well-formed, no regex can catch it — but it
      // makes the deposit/debit total NOT tie. We use these control totals as the
      // arbiter to detect and repair such single-digit corruption after parsing.
      const _num = s => { if (s == null) return null; const neg = /-\s*$/.test(s); const v = parseFloat(String(s).replace(/[,$\s\-]/g, '')); if (!isFinite(v)) return null; return neg ? -v : v; };
      // rx has two capture groups: integer part and 2-digit decimal. Join with a
      // literal dot so "3,030,021" + "64" becomes 3030021.64, not 303002164.
      const _grabAmt = (rx) => { const m = text.match(rx); return m ? _num(m[1] + '.' + m[2]) : null; };
      // "Deposits/Credits <count> <amount> +"  — amount is the LAST money-looking token on the line
      let ctrlDeposits = _grabAmt(/deposits?\s*\/?\s*credits[^\n]*?(\d{1,3}(?:,\d{3})*|\d+)\.(\d{2})\s*\+?/i);
      let ctrlChecks   = _grabAmt(/checks?\s*\/?\s*debits[^\n]*?(\d{1,3}(?:,\d{3})*|\d+)\.(\d{2})\s*-?/i);
      const ctrlPrev     = _grabAmt(/previous\s*balance[^\n]*?(\d{1,3}(?:,\d{3})*|\d+)\.(\d{2})/i);
      const ctrlCurr     = _grabAmt(/current\s*(?:statement\s*)?balance[^\n]*?(\d{1,3}(?:,\d{3})*|\d+)\.(\d{2})/i);
      // pdf-parse strips ALL spaces on some statements, so the summary line's
      // transaction COUNT glues onto the amount: "Deposits/Credits33,030,021.64+"
      // is count 3 + amount 3,030,021.64, but the regex reads 33,030,021.64.
      // The Previous/Current Balance lines have no count column, so they parse
      // cleanly — use the balance identity (prev + deposits - checks = curr) as
      // the arbiter: try trimming 0-3 leading digits off each raw total and
      // accept the UNIQUE pair that satisfies the identity. If zero or multiple
      // pairs tie, keep the raw values (reconciliation then simply won't fire,
      // which falls back to flagging — never a wrong silent repair).
      if (ctrlPrev != null && ctrlCurr != null && ctrlDeposits != null && ctrlChecks != null) {
        const _cands = (v) => {
          const out = [v]; const s = Math.abs(v).toFixed(2); const [ip, dp] = s.split('.');
          for (let k = 1; k <= 3 && k < ip.length; k++) {
            const t = parseFloat(ip.slice(k) + '.' + dp); if (t > 0) out.push(t);
          }
          return out;
        };
        const target = +(ctrlCurr - ctrlPrev).toFixed(2);
        const ties = [];
        for (const d of _cands(ctrlDeposits)) for (const c of _cands(ctrlChecks))
          if (Math.abs(d - c - target) < 0.005) ties.push([d, c]);
        if (ties.length === 1) { ctrlDeposits = ties[0][0]; ctrlChecks = ties[0][1]; }
      }

      // Some PDFs (e.g. Sunflower Bank / First National 1870) linearize through
      // pdf-parse with ALL inter-word spaces stripped, so a transaction row comes
      // out glued: "06/04/26APPFOLIOSAASCOUNTYLINERAILFUND2,880.00-". The parser
      // therefore keys off the date at the START of the line and the money token
      // at the END, and takes whatever is between as the description — this works
      // whether or not the words are space-separated.

      // Cut off everything from the Daily Balance Summary onward: that table and
      // the disclosures that follow also begin lines with dates but are NOT
      // transactions (folding them in would double-count balances as activity).
      // \s* between words so it still matches when pdf-parse strips inter-word
      // spaces ("DailyBalanceSummary").
      const stopIdx = lines.findIndex(l => /daily\s*balance\s*summary|balance\s*summary/i.test(l));
      if (stopIdx >= 0) lines = lines.slice(0, stopIdx);

      // Date at the very start of the line (with or without a following space).
      const dateHeadRx = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/;
      const altDateRx = /^(\w{3,9})\s+(\d{1,2}),?\s+(\d{4})/; // "January 15, 2024"
      // A money token. The integer part is BOUNDED (1-3 digits + comma groups, or
      // up to 7 plain digits) so a long reference number with no decimal — e.g.
      // "RF#105925009476061126" glued in front of "663,719.62" — can't be sucked
      // into the amount. A decimal ".dd" is REQUIRED. Supports a leading sign, a
      // TRAILING minus (bank convention: "2,880.00-"), and parenthesized negatives.
      // LEFT BOUNDARY (?<![\d,]): the integer part must NOT be immediately preceded
      // by a digit or comma, so a reference number glued to the amount with no
      // space (e.g. "...RF#...07292622,372.55") can't donate its trailing digit to
      // the amount (which produced a bogus 622,372.55 instead of 22,372.55).
      const moneyRx = /(?<![\d,])\(?[-+]?\$?(?:\d{1,3}(?:,\d{3})+|\d{1,7})\.\d{2}-?\)?/g;
      // Recovery scan for the fully-glued case where moneyRx (with its left
      // boundary) finds nothing because the amount is fused to a longer digit run.
      // We take the amount as the DECIMAL plus the minimal integer part to its
      // left, then flag the row for manual review since the exact integer/amount
      // split can be ambiguous when everything is glued.
      const gluedMoneyRx = /(\d{1,3}(?:,\d{3})+|\d{1,7})\.\d{2}-?/g;

      // Keywords that indicate money OUT / money IN, used only as a fallback to
      // fix the sign when neither the amount token NOR the section header settles it.
      const withdrawalRx = /withdraw|payment|payable|debit|paid out|ach debit|wire out|xfer\s*(from|out)|check\b|chk\b|pmt\b|svc\b|fee\b/i;
      const depositRx = /deposit|credit|interest\s*(paid|capitali|earned|income)|ach credit|wire in|xfer\s*in/i;

      // Section headers. Statements (e.g. Sunflower Bank / First National 1870)
      // group rows under a section that determines sign far more reliably than
      // the description: a row under "Deposits" is an inflow even if its memo says
      // "AP PAYMENT" (a customer's payment TO us), and a row under "Electronic
      // Transactions"/"Checks" is an outflow. Deposit amounts carry NO sign;
      // debits carry a trailing minus. We track the current section as we walk the
      // linearized text and use it as the PRIMARY sign source. section: +1 inflow,
      // -1 outflow, 0 unknown (fall back to token sign then keywords).
      const sectionInRx = /^(deposits|deposits\s*\/\s*credits|credits|additions)\b/i;
      const sectionOutRx = /^(electronic transactions|checks paid|checks paid electronically|checks\/debits|withdrawals|debits|other debits|checks)\b/i;
      const sectionNeutralRx = /^(commercial\s*checking\s*summary|account\s*summary|daily\s*balance\s*summary|balance\s*summary|overdraft)/i;

      const parseDate = s => {
        let m = s.match(dateHeadRx);
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

      const parseAmt = tok => {
        const t = tok.trim();
        const paren = /^\(.*\)$/.test(t);
        // Trailing-minus form ("2,880.00-") is negative, but only when there's no
        // leading sign already (so we don't double-negate "-2,880.00").
        const trailNeg = /-\)?$/.test(t) && !/^[-+]/.test(t);
        const clean = t.replace(/[$,()\s]/g, '').replace(/-$/, '');
        let v = parseFloat(clean) || 0;
        if (paren || trailNeg) v = -Math.abs(v);
        if (t.startsWith('-') && v > 0) v = -v;
        return v;
      };

      let lastDate = null;
      let section = 0; // +1 = deposits/credits, -1 = debits/checks, 0 = unknown
      for (const line of lines) {
        if (/statement date/i.test(line)) continue; // header, not a transaction
        // Update the section context from any header line BEFORE trying to parse
        // the line as a transaction (header lines carry no money token anyway).
        if (sectionInRx.test(line)) { section = 1; continue; }
        if (sectionOutRx.test(line)) { section = -1; continue; }
        if (sectionNeutralRx.test(line)) { section = 0; continue; }
        const lineDate = parseDate(line);
        if (lineDate) lastDate = lineDate;
        const usedDate = lineDate || lastDate;
        if (!usedDate) continue; // no date context yet — skip preamble lines

        const money = line.match(moneyRx);
        let amtTok, needsReview = false;
        if (money && money.length) {
          // The transaction amount is the LAST money token on the line. A plain
          // transaction row has exactly one; if a row ever carried a running
          // balance too, the activity amount precedes the balance — but these
          // statements put only the amount on the row, so "last" is safe and also
          // avoids grabbing a glued-in figure earlier in the description.
          amtTok = money[money.length - 1];
        } else {
          // No clean money token (its left boundary was a digit) — the amount is
          // fused to a preceding digit run. Recover the trailing decimal amount so
          // the row still imports, but FLAG it: the integer/amount split can be
          // ambiguous when fully glued (e.g. "...07292622,372.55" could read as
          // 22,372.55 or 622,372.55). The recovered value uses the minimal
          // comma-grouped reading; the reviewer confirms it against the statement.
          const glued = line.match(gluedMoneyRx);
          if (!glued || !glued.length) continue;
          amtTok = glued[glued.length - 1];
          needsReview = true;
        }
        let amount = parseAmt(amtTok);
        if (amount === 0) continue;

        // Description = text between the date and the FIRST money token.
        let afterDate = line;
        const dm = line.match(dateHeadRx) || line.match(altDateRx);
        if (dm && dm.index === 0) afterDate = line.slice(dm[0].length);
        // For a clean token, cut the description at the first money token; for a
        // glued token, cut at where the recovered amount starts in the line.
        let firstMoneyIdx = afterDate.search(moneyRx);
        if (firstMoneyIdx < 0) { const gi = afterDate.lastIndexOf(amtTok); firstMoneyIdx = gi; }
        let desc = (firstMoneyIdx > 0 ? afterDate.slice(0, firstMoneyIdx) : afterDate).trim();
        desc = desc.replace(/[()]+$/g, '').replace(/^[()]+/g, '').trim();
        if (!desc) continue;

        // Sign resolution, in priority order:
        //  1. An EXPLICIT sign on the token always wins (trailing "-", parens, or
        //     a leading +/-). e.g. "35,738.74-" is unambiguously an outflow.
        //  2. Otherwise the SECTION header settles it: a bare amount under
        //     "Deposits" is an inflow; under "Electronic Transactions"/"Checks"
        //     it's an outflow. This is what fixes a customer payment whose memo
        //     reads "AP PAYMENT" but which is listed under Deposits.
        //  3. Only when there is NO section context do we fall back to description
        //     keywords (older/loose statements with no section grouping).
        const tokenHadSign = /^[-+]/.test(amtTok.trim()) || /-\)?$/.test(amtTok.trim()) || /^\(.*\)$/.test(amtTok.trim());
        if (!tokenHadSign) {
          if (section === 1) amount = Math.abs(amount);
          else if (section === -1) amount = -Math.abs(amount);
          else {
            if (amount > 0 && withdrawalRx.test(desc) && !depositRx.test(desc)) amount = -amount;
            else if (amount < 0 && depositRx.test(desc) && !withdrawalRx.test(desc)) amount = Math.abs(amount);
          }
        }

        if (amount === 0 || !desc) continue;
        rows.push({ date: usedDate, description: desc.substring(0, 500), amount, needs_review: needsReview });
      }

      // ── Reconcile against control totals; auto-repair stolen-leading-digit rows.
      // pdf-parse can steal one leading digit of a description INTO the amount
      // (a well-formed but wrong token). When the parsed deposit/debit totals do
      // NOT tie to the statement's summary control totals, try trimming a single
      // leading digit off candidate amounts (3+ integer digits) to find the
      // reading that makes BOTH totals reconcile. Applied only when it resolves
      // uniquely; otherwise the row keeps its needs_review flag for manual check.
      const _EPS = 0.02;
      const depTotal = () => rows.filter(r => r.amount > 0).reduce((a, r) => a + r.amount, 0);
      const chkTotal = () => rows.filter(r => r.amount < 0).reduce((a, r) => a + Math.abs(r.amount), 0);
      const depOK = ctrlDeposits == null || Math.abs(depTotal() - ctrlDeposits) < _EPS;
      const chkOK = ctrlChecks == null || Math.abs(chkTotal() - ctrlChecks) < _EPS;
      if ((ctrlDeposits != null || ctrlChecks != null) && (!depOK || !chkOK)) {
        // Candidate rows: integer part has >1 digit so a leading digit can be trimmed.
        const trimCand = (v) => {
          const abs = Math.abs(v); const s = abs.toFixed(2); const [ip, dp] = s.split('.');
          if (ip.length <= 1) return null;
          const t = parseFloat(ip.slice(1) + '.' + dp);
          return t > 0 ? (v < 0 ? -t : t) : null;
        };
        // Try trimming exactly one candidate row (covers the common single-corruption
        // case) and accept the first that makes the relevant side(s) tie.
        for (let i = 0; i < rows.length; i++) {
          const orig = rows[i].amount;
          const cand = trimCand(orig);
          if (cand == null) continue;
          rows[i].amount = cand;
          const nowDepOK = ctrlDeposits == null || Math.abs(depTotal() - ctrlDeposits) < _EPS;
          const nowChkOK = ctrlChecks == null || Math.abs(chkTotal() - ctrlChecks) < _EPS;
          if (nowDepOK && nowChkOK) { rows[i].needs_review = false; break; }
          rows[i].amount = orig; // revert; keep looking
        }
      }
      // Re-evaluate: if totals now tie, clear any residual review flags that were
      // only about amount ambiguity (the batch is proven correct against control).
      const finalDepOK = ctrlDeposits == null || Math.abs(depTotal() - ctrlDeposits) < _EPS;
      const finalChkOK = ctrlChecks == null || Math.abs(chkTotal() - ctrlChecks) < _EPS;
      reconciled = (ctrlDeposits != null || ctrlChecks != null) && finalDepOK && finalChkOK;
      ctrlInfo = { deposits: ctrlDeposits, checks: ctrlChecks, prev: ctrlPrev, curr: ctrlCurr,
        parsed_deposits: +depTotal().toFixed(2), parsed_checks: +chkTotal().toFixed(2) };
      if (reconciled) rows.forEach(r => { r.needs_review = false; });

      } // end generic (non-Bank-of-Texas) PDF parser

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
    const insSplit = db.prepare('INSERT INTO bank_transaction_splits (txn_id, account_code, amount, memo, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?)');
    let count = 0;
    let autoCoded = 0;
    const reviewRows = [];
    db.transaction(() => {
      for (const r of rows) {
        // A glued-amount row (reference number fused to the amount) has an
        // ambiguous integer/amount split; mark its description so it's obvious in
        // the grid that the amount must be verified against the statement.
        const desc = r.needs_review ? ('[VERIFY AMOUNT] ' + r.description) : r.description;
        const info = ins.run(req.params.eid, bankAccount, r.date, desc, r.amount, batchId);
        count++;
        if (r.needs_review) reviewRows.push({ date: r.date, description: r.description, amount: r.amount });

        // ── Wire coding notes: auto-populate coding from a matching note ──
        // A glued/unverified amount is not trustworthy enough to match on, so
        // skip auto-coding those rows.
        if (!r.needs_review) {
          const note = findCodingNote(req.params.eid, bankAccount, r);
          if (note) {
            const txnId = info.lastInsertRowid;
            let splits = null;
            if (note.splits_json) { try { splits = JSON.parse(note.splits_json); } catch { splits = null; } }
            if (Array.isArray(splits) && splits.length) {
              for (const s of splits) {
                insSplit.run(txnId, s.account_code, Math.abs(Number(s.amount)), s.memo || note.memo || null,
                  s.project_id || null, s.class_id || null, s.location_id || null);
              }
              db.prepare("UPDATE bank_transactions SET account_code=NULL, memo=?, status='coded' WHERE id=?")
                .run(note.memo || null, txnId);
            } else if (note.account_code) {
              db.prepare("UPDATE bank_transactions SET account_code=?, memo=?, project_id=?, class_id=?, location_id=?, status='coded' WHERE id=?")
                .run(note.account_code, note.memo || null, note.project_id || null, note.class_id || null, note.location_id || null, txnId);
            }
            db.prepare("UPDATE bank_coding_notes SET matched_count=matched_count+1, last_matched_at=datetime('now') WHERE id=?").run(note.id);
            autoCoded++;
          }
        }
      }
    })();

    res.json({ count, batch_id: batchId, format: isPdf ? 'pdf' : 'csv/xlsx',
      needs_review: reviewRows.length, review_rows: reviewRows,
      auto_coded: autoCoded,
      reconciled, control_totals: ctrlInfo });
  } catch (e) {
    res.status(400).json({ error: 'Failed to parse file: ' + e.message });
  }
});

app.put('/api/entities/:eid/bank-transactions/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { account_code, memo, project_id, class_id, location_id, amount } = req.body;
  // Optional amount correction — used to fix a mis-parsed figure (e.g. a PDF
  // row flagged [VERIFY AMOUNT] where a glued reference number bled a digit into
  // the amount). Allowed ONLY on a not-yet-posted row, since a posted row already
  // has a GL journal entry that would no longer match. Sign is preserved from the
  // supplied value; pass a signed number.
  if (amount !== undefined && amount !== null && amount !== '') {
    const cur = db.prepare('SELECT status FROM bank_transactions WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
    if (!cur) return res.status(404).json({ error: 'Transaction not found' });
    if (cur.status === 'posted') return res.status(400).json({ error: 'Cannot change amount on a posted transaction; unpost it first' });
    const amt = Number(amount);
    if (!isFinite(amt) || amt === 0) return res.status(400).json({ error: 'amount must be a non-zero number' });
    db.prepare('UPDATE bank_transactions SET amount=? WHERE id=? AND entity_id=?').run(amt, req.params.id, req.params.eid);
    // Once the amount is verified/corrected, drop the [VERIFY AMOUNT] flag from
    // the description so the grid no longer shows a stale warning on this row.
    db.prepare("UPDATE bank_transactions SET description = REPLACE(description, '[VERIFY AMOUNT] ', '') WHERE id=? AND entity_id=?").run(req.params.id, req.params.eid);
  }
  // Setting a single account_code clears any existing splits
  db.transaction(() => {
    db.prepare('DELETE FROM bank_transaction_splits WHERE txn_id=?').run(req.params.id);
    db.prepare('UPDATE bank_transactions SET account_code=?, memo=?, project_id=?, class_id=?, location_id=?, status=? WHERE id=? AND entity_id=?')
      .run(account_code || null, memo || null, project_id || null, class_id || null, location_id || null, account_code ? 'coded' : 'pending', req.params.id, req.params.eid);
  })();
  res.json({ success: true });
});

// Set multiple splits for a single bank transaction.
// Splits: [{ account_code, amount, memo }]. Amounts are stored as MAGNITUDES relative
// to the transaction's own direction: a positive split moves money the same way as the
// transaction (a deposit's invoice line, a payment's expense line); a NEGATIVE split
// offsets it (a credit memo applied against a receipt, a refund netted against a
// payment). The signed net of the splits must equal abs(txn.amount).
app.put('/api/entities/:eid/bank-transactions/:id/splits', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { splits } = req.body;
  const txn = db.prepare('SELECT * FROM bank_transactions WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!txn) return res.status(404).json({ error: 'Transaction not found' });
  if (txn.status === 'posted') return res.status(400).json({ error: 'Cannot edit a posted transaction' });
  if (!Array.isArray(splits) || splits.length === 0) return res.status(400).json({ error: 'At least one split required' });
  for (const s of splits) {
    if (!s.account_code) return res.status(400).json({ error: 'Each split needs an account' });
    if (!Number.isFinite(Number(s.amount)) || Number(s.amount) === 0) return res.status(400).json({ error: 'Each split amount must be a non-zero number' });
  }
  // Validate on the NET of the split magnitudes vs the transaction magnitude, so a
  // receipt applied net of a credit memo (34,934.16 - 4,993.96 = 29,940.20) ties out
  // instead of being rejected for dropping the negative line. Require a positive net
  // so the direction still matches the transaction (a deposit stays a deposit).
  const total = splits.reduce((sum, s) => sum + Number(s.amount), 0);
  const target = Math.abs(txn.amount);
  if (!(total > 0)) return res.status(400).json({ error: 'Splits must net to a positive amount matching the transaction (currently ' + total.toFixed(2) + ')' });
  if (Math.abs(total - target) > 0.005) return res.status(400).json({ error: 'Splits total ' + total.toFixed(2) + ' does not match transaction amount ' + target.toFixed(2) });

  db.transaction(() => {
    db.prepare('DELETE FROM bank_transaction_splits WHERE txn_id=?').run(txn.id);
    const ins = db.prepare('INSERT INTO bank_transaction_splits (txn_id, account_code, amount, memo, project_id, class_id, location_id, invoice_id) VALUES (?,?,?,?,?,?,?,?)');
    for (const s of splits) ins.run(txn.id, s.account_code, Number(s.amount), s.memo || null, s.project_id || null, s.class_id || null, s.location_id || null, s.invoice_id || null);
    // Clear the single-code field + its dimensions (dimensions now live per split) and mark coded
    db.prepare('UPDATE bank_transactions SET account_code=NULL, project_id=NULL, class_id=NULL, location_id=NULL, status=? WHERE id=?').run('coded', txn.id);
  })();
  res.json({ success: true });
});

app.post('/api/entities/:eid/bank-transactions/post', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const { transaction_ids } = req.body;
  const arMod = require('./ar');
  if (!transaction_ids || transaction_ids.length === 0) return res.status(400).json({ error: 'No transactions' });

  const txns = db.prepare(`SELECT * FROM bank_transactions WHERE entity_id=? AND id IN (${transaction_ids.map(()=>'?').join(',')}) AND status='coded'`)
    .all(req.params.eid, ...transaction_ids);
  if (txns.length === 0) return res.status(400).json({ error: 'No coded transactions to post' });

  // Guard: never post unless EVERY selected transaction has a valid bank account.
  // A missing/blank/unknown bank_account_code would post the offset line but leave
  // the bank side on a null account, creating an unbalanced entry. Block the whole
  // batch (all-or-nothing) so no imbalance is ever written, and report which
  // transactions need a bank account chosen first.
  const validBankCodes = new Set(db.prepare('SELECT code FROM accounts WHERE entity_id=?').all(req.params.eid).map(a => String(a.code)));
  const badBank = txns.filter(t => !t.bank_account_code || !validBankCodes.has(String(t.bank_account_code)));
  if (badBank.length > 0) {
    return res.status(400).json({
      error: 'Cannot post: ' + badBank.length + ' transaction' + (badBank.length === 1 ? '' : 's') + ' ' + (badBank.length === 1 ? 'has' : 'have') + ' no valid bank account selected. Select a bank account for these before posting.',
      transactions_missing_bank_account: badBank.map(t => ({ id: t.id, date: t.date, description: t.description, amount: t.amount, bank_account_code: t.bank_account_code || null })),
    });
  }
  // Guard: every posting transaction must also have a coded offset (single
  // account or splits), otherwise the offset side would be null.
  const badOffset = txns.filter(t => {
    const hasSplits = db.prepare('SELECT COUNT(*) c FROM bank_transaction_splits WHERE txn_id=?').get(t.id).c > 0;
    return !hasSplits && !t.account_code;
  });
  if (badOffset.length > 0) {
    return res.status(400).json({
      error: 'Cannot post: ' + badOffset.length + ' transaction' + (badOffset.length === 1 ? '' : 's') + ' ' + (badOffset.length === 1 ? 'has' : 'have') + ' no account coded. Code the offsetting account before posting.',
      transactions_missing_account: badOffset.map(t => ({ id: t.id, date: t.date, description: t.description, amount: t.amount })),
    });
  }

  // Guard: a DEPOSIT line coded to an A/R (Accounts Receivable) control account must
  // name the invoice it's clearing. Without an invoice_id the post would credit the
  // A/R control account without applying to any document, so the GL control drifts
  // away from the aging subledger. A/R accounts are detected by name the same way the
  // aging report does (Accounts Receivable, excluding "- Other"/note/interest variants),
  // not by a hard-coded 12000, so other entities' A/R codes are covered too.
  // Scope: deposits (positive amount) only. Whole batch is all-or-nothing.
  const arCodeSet = new Set(
    db.prepare('SELECT code, name FROM accounts WHERE entity_id=?').all(req.params.eid)
      .filter(a => /accounts?\s*receivable/i.test(a.name || '') && !/other|note|interest/i.test(a.name || ''))
      .map(a => String(a.code))
  );
  if (arCodeSet.size > 0) {
    const arNoInvoice = [];
    for (const t of txns) {
      if (!(Number(t.amount) > 0)) continue; // deposits only
      const splits = db.prepare('SELECT * FROM bank_transaction_splits WHERE txn_id=? ORDER BY id').all(t.id);
      if (splits.length > 0) {
        splits.forEach((s, i) => {
          // Only positive (natural-direction) A/R lines clear an invoice; a negative
          // A/R line is a GL reclass/offset, not a cash receipt, so it needs no invoice.
          if (arCodeSet.has(String(s.account_code)) && Number(s.amount) > 0 && !s.invoice_id) {
            arNoInvoice.push({ id: t.id, date: t.date, description: t.description, amount: t.amount, line: i + 1, account_code: s.account_code, split_amount: s.amount });
          }
        });
      } else if (arCodeSet.has(String(t.account_code))) {
        // A single-account A/R deposit has no way to attach an invoice — always block.
        arNoInvoice.push({ id: t.id, date: t.date, description: t.description, amount: t.amount, line: null, account_code: t.account_code, split_amount: t.amount });
      }
    }
    if (arNoInvoice.length > 0) {
      const first = arNoInvoice[0];
      return res.status(400).json({
        error: 'Cannot post: ' + arNoInvoice.length + ' deposit line' + (arNoInvoice.length === 1 ? '' : 's') + ' coded to Accounts Receivable ' + (arNoInvoice.length === 1 ? 'has' : 'have') + ' no invoice attached. '
          + 'Open the deposit (' + first.date + ', $' + Math.abs(Number(first.amount)).toFixed(2) + (first.line ? ', line ' + first.line : '') + ') and apply it to the invoice it pays before posting. '
          + 'An A/R deposit must clear a specific invoice, otherwise the aging report and the GL control account drift apart.',
        transactions_ar_missing_invoice: arNoInvoice,
      });
    }
  }

  // Guard: never over-apply an invoice across the batch. recordArReceipt already
  // rejects a receipt that exceeds an invoice's open balance versus ALREADY-POSTED
  // receipts — but two unposted deposits in the same batch could each try to clear
  // the same invoice (the classic double-apply). Sum every A/R application in this
  // batch by invoice, add what's already been received (posted) on that invoice, and
  // reject before writing anything if the combined total exceeds the invoice total.
  // This is the server backstop behind the split-modal picker's effective-remaining.
  {
    const applyByInvoice = new Map(); // invoice_id -> [{txn_id, date, amount}]
    for (const t of txns) {
      if (!(Number(t.amount) > 0)) continue;
      const splits = db.prepare('SELECT * FROM bank_transaction_splits WHERE txn_id=? ORDER BY id').all(t.id);
      for (const s of splits) {
        if (!s.invoice_id || !(Number(s.amount) > 0)) continue;
        if (!applyByInvoice.has(s.invoice_id)) applyByInvoice.set(s.invoice_id, []);
        applyByInvoice.get(s.invoice_id).push({ txn_id: t.id, date: t.date, amount: Number(s.amount) });
      }
    }
    for (const [invId, apps] of applyByInvoice) {
      const inv = db.prepare('SELECT invoice_num, total, status FROM ar_invoices WHERE id=? AND entity_id=?').get(invId, req.params.eid);
      if (!inv) return res.status(400).json({ error: 'Cannot post: an A/R deposit line references invoice #' + invId + ', which is not an invoice for this entity.' });
      if (inv.status === 'void') return res.status(400).json({ error: 'Cannot post: a deposit line is applied to voided invoice ' + inv.invoice_num + '.' });
      const posted = Number(db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id=?').get(invId).p);
      const batchApply = apps.reduce((s, a) => s + a.amount, 0);
      const invTotal = Number(inv.total || 0);
      if (invTotal >= 0 && (posted + batchApply) - invTotal > 0.005) {
        const openBefore = +(invTotal - posted).toFixed(2);
        const detail = apps.length > 1
          ? apps.length + ' deposits in this batch together apply $' + batchApply.toFixed(2)
          : 'this deposit applies $' + batchApply.toFixed(2);
        return res.status(400).json({
          error: 'Cannot post: invoice ' + inv.invoice_num + ' has only $' + openBefore.toFixed(2) + ' open'
            + (posted > 0.005 ? ' (after $' + posted.toFixed(2) + ' already received)' : '')
            + ', but ' + detail + ' to it. Two deposits can\'t both clear the same invoice — '
            + 'apply one of them to the correct invoice (or reduce the amount) before posting.',
          invoice: { id: invId, invoice_num: inv.invoice_num, total: invTotal, already_received: +posted.toFixed(2), open: openBefore, batch_applied: +batchApply.toFixed(2) },
          applications: apps,
        });
      }
    }
  }

  const results = [];
  try {
  db.transaction(() => {
    for (const t of txns) {
      const splits = db.prepare('SELECT * FROM bank_transaction_splits WHERE txn_id=? ORDER BY id').all(t.id);
      const hasSplits = splits.length > 0;
      // Defense in depth: the batch-level guards above already rejected these,
      // but re-check per transaction so a malformed row can never post a
      // half-sided (unbalanced) entry.
      if (!t.bank_account_code) continue;
      if (!hasSplits && !t.account_code) continue;

      const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id=?').get(req.params.eid).m||0)+1;
      const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
        .run(req.params.eid, num, t.date, t.memo || t.description, req.user.name);
      const jeId = r.lastInsertRowid;
      const insLine = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?)');
      const abs = Math.abs(t.amount);
      // Bank-account line carries no dimension; the coded offset line(s) do.
      const tDims = [t.project_id || null, t.class_id || null, t.location_id || null];
      const sDims = s => [s.project_id || null, s.class_id || null, s.location_id || null];

      // A split amount is a magnitude in the transaction's direction. Its natural
      // side is credit for a deposit, debit for a payment. A NEGATIVE split (a credit
      // memo / offset) posts to the OPPOSITE side as a positive amount, rather than a
      // negative entry on the natural side, so the JE reads cleanly and stays balanced.
      const postSplit = (s, naturalDebit) => {
        const amt = Math.abs(Number(s.amount));
        const onDebit = Number(s.amount) < 0 ? !naturalDebit : naturalDebit;
        insLine.run(jeId, s.account_code, onDebit ? amt : 0, onDebit ? 0 : amt, ...sDims(s));
      };
      if (t.amount > 0) {
        // Deposit: debit bank once, credit each coded account (negative split => debit)
        insLine.run(jeId, t.bank_account_code, abs, 0, null, null, null);
        if (hasSplits) { for (const s of splits) postSplit(s, false); }
        else { insLine.run(jeId, t.account_code, 0, abs, ...tDims); }
      } else {
        // Payment: debit each coded account, credit bank once (negative split => credit)
        if (hasSplits) { for (const s of splits) postSplit(s, true); }
        else { insLine.run(jeId, t.account_code, abs, 0, ...tDims); }
        insLine.run(jeId, t.bank_account_code, 0, abs, null, null, null);
      }

      db.prepare('UPDATE bank_transactions SET status=?, je_id=? WHERE id=?').run('posted', jeId, t.id);
      // Deposit splits coded to an A/R invoice also record a subledger receipt
      // against that invoice, so the aging report clears the specific document
      // while the GL keeps one clean deposit JE (Dr bank / Cr A/R control).
      if (hasSplits && t.amount > 0) {
        for (const s of splits) {
          if (!s.invoice_id) continue;
          if (!(Number(s.amount) > 0)) continue; // negative (credit-memo/offset) line is a GL reclass, not a cash receipt
          arMod.recordArReceipt(db, { entity_id: +req.params.eid, invoice_id: s.invoice_id, date: t.date,
            amount: s.amount, bank_account_code: t.bank_account_code, memo: s.memo || t.memo || null,
            je_id: jeId, created_by: req.user.name });
        }
      }
      results.push({ txn_id: t.id, je_id: jeId, entry_num: num });
    }
  })();
  } catch (e) {
    // The whole batch runs in one db.transaction, so any throw here (e.g. a
    // recordArReceipt over-application backstop) rolls back every JE — nothing
    // half-posts. Surface it as a readable 400 instead of a raw 500.
    return res.status(400).json({ error: 'Post failed, nothing was posted: ' + e.message });
  }
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
  // Only discard rows that aren't yet tied to a journal entry. A 'posted' row has
  // its own JE; a 'matched' row is linked (reconciled) to an existing JE — deleting
  // either would orphan/break that GL link. Both are protected here so a batch
  // discard can never remove a finished row, even if the client miscounts.
  const r = db.prepare("DELETE FROM bank_transactions WHERE entity_id=? AND batch_id=? AND status NOT IN ('posted','matched')").run(req.params.eid, req.params.batchId);
  res.json({ deleted: r.changes });
});

// ═══ Bank Coding Notes (wire pre-coding) ═══
// Leave a note during the month describing how a wire should be coded; on the
// next statement upload the matching row is auto-populated (status 'coded') and
// the note is kept for reference (matched_count/last_matched_at record the hit).
app.get('/api/entities/:eid/bank-coding-notes', auth, requireEntityAccess(), (req, res) => {
  const rows = db.prepare('SELECT * FROM bank_coding_notes WHERE entity_id=? ORDER BY active DESC, id DESC').all(req.params.eid);
  const attStmt = db.prepare('SELECT id, original_name, mime_type, size FROM bank_coding_note_attachments WHERE note_id=? ORDER BY id');
  res.json(rows.map(r => ({ ...r, splits: r.splits_json ? JSON.parse(r.splits_json) : null, attachments: attStmt.all(r.id) })));
});

app.post('/api/entities/:eid/bank-coding-notes', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const b = req.body || {};
  if (b.match_amount === undefined || b.match_amount === null || b.match_amount === '' || !isFinite(Number(b.match_amount)) || Number(b.match_amount) === 0)
    return res.status(400).json({ error: 'match_amount is required and must be a non-zero signed number' });
  const splits = Array.isArray(b.splits) && b.splits.length ? b.splits : null;
  if (!splits && !b.account_code)
    return res.status(400).json({ error: 'Provide account_code or splits for the coding' });
  if (splits) for (const s of splits) {
    if (!s.account_code) return res.status(400).json({ error: 'Each split needs an account_code' });
    if (!(Number(s.amount) > 0)) return res.status(400).json({ error: 'Each split amount must be > 0' });
  }
  const info = db.prepare(
    `INSERT INTO bank_coding_notes
      (entity_id, bank_account_code, note, match_amount, amount_tolerance, date_from, date_to, desc_keyword,
       account_code, splits_json, memo, project_id, class_id, location_id, one_shot, active, created_by)
     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
  ).run(
    req.params.eid, b.bank_account_code || null, b.note || null,
    Number(b.match_amount), Number(b.amount_tolerance) || 0,
    b.date_from || null, b.date_to || null, b.desc_keyword || null,
    splits ? null : b.account_code, splits ? JSON.stringify(splits) : null,
    b.memo || null, b.project_id || null, b.class_id || null, b.location_id || null,
    b.one_shot === false ? 0 : 1, 1, req.user.name
  );
  res.json({ id: info.lastInsertRowid, success: true });
});

app.put('/api/entities/:eid/bank-coding-notes/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const existing = db.prepare('SELECT * FROM bank_coding_notes WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!existing) return res.status(404).json({ error: 'Note not found' });
  const b = req.body || {};
  const splits = Array.isArray(b.splits) ? (b.splits.length ? b.splits : null) : undefined;
  const val = (k, fallback) => (b[k] === undefined ? fallback : b[k]);
  db.prepare(
    `UPDATE bank_coding_notes SET
       bank_account_code=?, note=?, match_amount=?, amount_tolerance=?, date_from=?, date_to=?, desc_keyword=?,
       account_code=?, splits_json=?, memo=?, project_id=?, class_id=?, location_id=?, one_shot=?, active=?
     WHERE id=? AND entity_id=?`
  ).run(
    val('bank_account_code', existing.bank_account_code), val('note', existing.note),
    b.match_amount === undefined ? existing.match_amount : Number(b.match_amount),
    b.amount_tolerance === undefined ? existing.amount_tolerance : Number(b.amount_tolerance),
    val('date_from', existing.date_from), val('date_to', existing.date_to), val('desc_keyword', existing.desc_keyword),
    splits === undefined ? existing.account_code : (splits ? null : val('account_code', existing.account_code)),
    splits === undefined ? existing.splits_json : (splits ? JSON.stringify(splits) : null),
    val('memo', existing.memo), val('project_id', existing.project_id),
    val('class_id', existing.class_id), val('location_id', existing.location_id),
    b.one_shot === undefined ? existing.one_shot : (b.one_shot ? 1 : 0),
    b.active === undefined ? existing.active : (b.active ? 1 : 0),
    req.params.id, req.params.eid
  );
  res.json({ success: true });
});

app.delete('/api/entities/:eid/bank-coding-notes/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  // Remove any supporting-doc files from disk before the row (and its
  // attachment rows via cascade) are deleted.
  const atts = db.prepare('SELECT a.filename FROM bank_coding_note_attachments a JOIN bank_coding_notes n ON n.id=a.note_id WHERE a.note_id=? AND n.entity_id=?').all(req.params.id, req.params.eid);
  atts.forEach(a => { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} });
  const r = db.prepare('DELETE FROM bank_coding_notes WHERE id=? AND entity_id=?').run(req.params.id, req.params.eid);
  res.json({ deleted: r.changes });
});

// ── Wire-note supporting documents (email copy / PDF / Excel / etc) ──
app.post('/api/entities/:eid/bank-coding-notes/:id/attachments', auth, requireEntityAccess(), requireRole('Admin','Accountant'), upload.array('files', 10), (req, res) => {
  const note = db.prepare('SELECT id FROM bank_coding_notes WHERE id=? AND entity_id=?').get(req.params.id, req.params.eid);
  if (!note) { (req.files||[]).forEach(f => { try { fs.unlinkSync(path.join(UPLOAD_DIR, f.filename)); } catch {} }); return res.status(404).json({ error: 'Note not found' }); }
  if (!req.files || req.files.length === 0) return res.status(400).json({ error: 'No files' });
  const ins = db.prepare('INSERT INTO bank_coding_note_attachments (note_id, filename, original_name, mime_type, size) VALUES (?,?,?,?,?)');
  const results = [];
  for (const f of req.files) {
    const r = ins.run(req.params.id, f.filename, f.originalname, f.mimetype, f.size);
    results.push({ id: r.lastInsertRowid, original_name: f.originalname, mime_type: f.mimetype, size: f.size });
  }
  res.json(results);
});

app.get('/api/bank-coding-note-attachments/:id/download', (req, res) => {
  const token = req.query.token || req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  try { jwt.verify(token, JWT_SECRET); } catch { return res.status(401).json({ error: 'Invalid token' }); }
  const att = db.prepare('SELECT * FROM bank_coding_note_attachments WHERE id=?').get(req.params.id);
  if (!att) return res.status(404).json({ error: 'Not found' });
  const filepath = path.resolve(UPLOAD_DIR, att.filename);
  if (!fs.existsSync(filepath)) return res.status(404).json({ error: 'File missing' });
  const inlineTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'];
  const disposition = inlineTypes.includes(att.mime_type) ? 'inline' : 'attachment';
  res.setHeader('Content-Disposition', disposition + '; filename="' + att.original_name + '"');
  res.setHeader('Content-Type', att.mime_type || 'application/octet-stream');
  res.sendFile(filepath, err => { if (err && !res.headersSent) res.status(500).json({ error: 'Failed to send file' }); });
});

app.delete('/api/bank-coding-note-attachments/:id', auth, requireRole('Admin','Accountant'), (req, res) => {
  const att = db.prepare('SELECT * FROM bank_coding_note_attachments WHERE id=?').get(req.params.id);
  if (att) { try { fs.unlinkSync(path.join(UPLOAD_DIR, att.filename)); } catch {} }
  db.prepare('DELETE FROM bank_coding_note_attachments WHERE id=?').run(req.params.id);
  res.json({ success: true });
});

// ═══ Balances (with soft close) ═══
// Core balances computation, shared by the HTTP endpoint and the financial-
// statement generator. Returns natural-signed balances per account for a date
// window, with optional prior-period P&L close into Retained Earnings.
// opts: { as_of, from, to, close_pl_before, location_id, class_id }
function computeBalances(eid, opts = {}) {
  const { as_of, from, to, close_pl_before, location_id, class_id, project_id } = opts;
  let dateFilter = ''; const params = [eid];
  if (as_of) { dateFilter = ' AND je.date <= ?'; params.push(as_of); }
  else if (from && to) { dateFilter = ' AND je.date >= ? AND je.date <= ?'; params.push(from, to); }
  else if (from) { dateFilter = ' AND je.date >= ?'; params.push(from); }
  else if (to) { dateFilter = ' AND je.date <= ?'; params.push(to); }
  let dimFilter = '';
  if (location_id) { dimFilter += ' AND jl.location_id = ?'; params.push(location_id); }
  if (class_id) { dimFilter += ' AND jl.class_id = ?'; params.push(class_id); }
  if (project_id) { dimFilter += ' AND CAST(jl.project_id AS REAL) = CAST(? AS REAL)'; params.push(project_id); }

  if (close_pl_before && as_of) {
    const priorPL = db.prepare(`SELECT a.type, SUM(jl.debit) as td, SUM(jl.credit) as tc FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? AND je.date<? AND a.type IN ('Revenue','Expense') GROUP BY a.type`).all(eid, close_pl_before);
    let priorNI = 0; priorPL.forEach(r => { if (r.type==='Revenue') priorNI+=(r.tc-r.td); if (r.type==='Expense') priorNI-=(r.td-r.tc); });
    let reAcct = db.prepare("SELECT * FROM accounts WHERE entity_id=? AND code='31000'").get(eid)
      || db.prepare("SELECT * FROM accounts WHERE entity_id=? AND type='Equity' AND LOWER(name) LIKE '%retained earning%' ORDER BY code LIMIT 1").get(eid);
    if (!reAcct && Math.abs(priorNI) > 0.005) {
      try {
        db.prepare("INSERT INTO accounts (entity_id, code, name, type, subtype, bank_acct) VALUES (?, '39000', 'Retained Earnings', 'Equity', '', 0)").run(eid);
        reAcct = db.prepare("SELECT * FROM accounts WHERE entity_id=? AND code='39000'").get(eid);
      } catch (e) { console.error('Auto-create RE 39000 failed:', e.message); }
    }
    const reCode = reAcct ? reAcct.code : null;
    const bsRows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=? AND je.date<=? AND (a.type NOT IN ('Revenue','Expense') OR je.date>=?) GROUP BY jl.account_code`).all(eid, as_of, close_pl_before);
    const results = bsRows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; let bal=isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit);
      if (reCode && r.account_code===reCode) bal+=priorNI;
      return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:bal, total_debit:r.total_debit, total_credit:r.total_credit }; });
    if (Math.abs(priorNI)>0.005 && reCode && !results.find(r=>r.code===reCode)) {
      results.push({ code:reCode, name:reAcct.name, type:reAcct.type, subtype:reAcct.subtype, bank_acct:0, balance:priorNI, total_debit:0, total_credit:0 });
    }
    return results;
  }

  const rows = db.prepare(`SELECT jl.account_code, a.type, a.name, a.subtype, a.bank_acct, SUM(jl.debit) as total_debit, SUM(jl.credit) as total_credit FROM journal_lines jl JOIN journal_entries je ON jl.entry_id=je.id JOIN accounts a ON a.entity_id=je.entity_id AND a.code=jl.account_code WHERE je.entity_id=?${dateFilter}${dimFilter} GROUP BY jl.account_code`).all(...params);
  return rows.map(r => { const isDr=r.type==='Asset'||r.type==='Expense'; return { code:r.account_code, name:r.name, type:r.type, subtype:r.subtype, bank_acct:r.bank_acct, balance:isDr?(r.total_debit-r.total_credit):(r.total_credit-r.total_debit), total_debit:r.total_debit, total_credit:r.total_credit }; });
}

app.get('/api/entities/:eid/balances', auth, requireEntityAccess(), (req, res) => {
  res.json(computeBalances(req.params.eid, req.query));
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
  let files = db.prepare('SELECT id, folder_path, original_name, size, mime_type, uploaded_by, created_at FROM entity_files WHERE entity_id=? ORDER BY folder_path, original_name').all(req.params.eid);
  let folders = db.prepare('SELECT folder_path, created_by, created_at FROM entity_folders WHERE entity_id=? ORDER BY folder_path').all(req.params.eid);
  // Admin-only folders: any path whose top segment is "_Admin" is hidden from
  // non-Admin users (files and folder entries alike).
  const isAdmin = req.user && req.user.role === 'Admin';
  if (!isAdmin) {
    const isAdminPath = p => String(p || '').split('/')[0] === '_Admin';
    files = files.filter(f => !isAdminPath(f.folder_path));
    folders = folders.filter(f => !isAdminPath(f.folder_path));
  }
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
  let claims;
  try {
    const token = req.query.token;
    if (!token) return res.status(401).json({ error: 'Token required' });
    claims = jwt.verify(token, JWT_SECRET);
  } catch { return res.status(401).json({ error: 'Invalid token' }); }
  const f = db.prepare('SELECT * FROM entity_files WHERE id=?').get(req.params.id);
  if (!f) return res.status(404).json({ error: 'Not found' });
  // Admin-only folders are downloadable only by Admins.
  if (String(f.folder_path || '').split('/')[0] === '_Admin' && !(claims && claims.role === 'Admin')) {
    return res.status(403).json({ error: 'Admin only' });
  }
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

// Maintenance (Admin only): purge auto-saved Requisition Report workpapers.
// Removes every entity_files row whose folder path is under a "Requisition
// Reports" folder (and its blob on disk), plus the now-empty "Requisition
// Reports" / month folder rows, across ALL entities. Destructive — requires
// confirm=PURGE. Scope is strictly limited to "Requisition Reports" paths.
app.post('/api/admin/purge-requisition-workpapers', auth, requireRole('Admin'), (req, res) => {
  const confirm = (req.body && req.body.confirm) || req.query.confirm;
  if (confirm !== 'PURGE') return res.status(400).json({ error: 'Add confirm=PURGE to run this destructive purge.' });
  const LIKE = '%Requisition Reports%';
  const files = db.prepare('SELECT id, entity_id, folder_path, original_name, stored_filename FROM entity_files WHERE folder_path LIKE ?').all(LIKE);
  const del = db.prepare('DELETE FROM entity_files WHERE id=?');
  const deletedFiles = [];
  for (const f of files) {
    try { fs.unlinkSync(path.join(WORKPAPERS_DIR, String(f.entity_id), f.stored_filename)); } catch {}
    del.run(f.id);
    deletedFiles.push({ entity_id: f.entity_id, folder: f.folder_path, name: f.original_name });
  }
  const folders = db.prepare('SELECT entity_id, folder_path FROM entity_folders WHERE folder_path LIKE ?').all(LIKE);
  db.prepare('DELETE FROM entity_folders WHERE folder_path LIKE ?').run(LIKE);
  res.json({ success: true, deletedFileCount: deletedFiles.length, deletedFolderCount: folders.length, deletedFiles, deletedFolders: folders });
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

// Undo (unpost) a completed bank reconciliation. QBO-style constraint: only the
// most recent reconciliation for an account can be undone, because each later
// reconciliation's beginning balance depends on the cleared state left by earlier
// ones. Undoing deletes the reconciliation row and un-clears its items, so they
// reappear as uncleared in the next reconciliation session. No journal entries
// are touched.
app.delete('/api/entities/:eid/reconciliations/:id', auth, requireEntityAccess(), requireRole('Admin','Accountant'), (req, res) => {
  const eid = req.params.eid, id = parseInt(req.params.id);
  const rec = db.prepare('SELECT * FROM reconciliations WHERE id=? AND entity_id=?').get(id, eid);
  if (!rec) return res.status(404).json({ error: 'Reconciliation not found' });
  const newer = db.prepare(
    'SELECT id, statement_date FROM reconciliations WHERE entity_id=? AND account_code=? AND (statement_date > ? OR (statement_date = ? AND id > ?)) ORDER BY statement_date DESC, id DESC'
  ).all(eid, rec.account_code, rec.statement_date, rec.statement_date, id);
  if (newer.length) return res.status(409).json({ error: 'A newer reconciliation exists for account ' + rec.account_code + ' (statement date ' + newer[0].statement_date + '). Undo reconciliations newest-first.' });
  const result = db.transaction(() => {
    const uncleared = db.prepare('DELETE FROM cleared_items WHERE reconciliation_id=?').run(id).changes;
    db.prepare('DELETE FROM reconciliations WHERE id=?').run(id);
    return uncleared;
  })();
  res.json({ success: true, id, account_code: rec.account_code, statement_date: rec.statement_date, items_uncleared: result });
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

// Direct (non-commitment) costs by cost code, for the Turnkey portal's Cost
// Report "Direct Costs" column. Sums the project's expense-account postings
// (= cost codes) excluding CIP commitments and the POC recognition account.
// GET /api/turnkey/projects/:id/direct-costs[?as_of=YYYY-MM-DD]
app.get('/api/turnkey/projects/:id/direct-costs', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  try {
    const data = turnkey.getDirectCosts(db, { turnkey_project_id: req.params.id, as_of: req.query.as_of });
    res.json(data);
  } catch (e) {
    const code = /not linked/.test(e.message) ? 404 : 500;
    res.status(code).json({ error: e.message });
  }
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
// List entity ids that have a Bill.com configuration (used to filter the setup
// entity dropdown to only Bill.com-enabled entities).
app.get('/api/billcom/entities', auth, requireRole('Admin', 'Accountant'), (req, res) => {
  const ids = db.prepare('SELECT entity_id FROM billcom_config').all().map(r => r.entity_id);
  res.json({ entity_ids: ids });
});
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
  // Auto-fill the Money Out Clearing account when the caller doesn't supply one,
  // by matching the entity's chart of accounts on name (prefer "Money Out
  // Clearing" / "Bill.com Clearing", else any account with "clearing" in the
  // name). So setting up a new entity's Bill.com config wires up the clearing
  // account without hunting for the GL code. An explicit value always wins.
  let clearingAcct = default_clearing_account || null;
  if (!clearingAcct) {
    const cand = db.prepare(
      "SELECT code FROM accounts WHERE entity_id = ? AND lower(name) LIKE '%clearing%' " +
      "ORDER BY (CASE " +
      "WHEN lower(name) LIKE '%money out clearing%' THEN 0 " +
      "WHEN lower(name) LIKE '%bill%com clearing%' THEN 1 " +
      "ELSE 2 END), code LIMIT 1"
    ).get(req.params.entity_id);
    if (cand) clearingAcct = cand.code;
  }
  const now = new Date().toISOString();
  const updater = req.user.name || req.user.email;
  if (existing) {
    db.prepare('UPDATE billcom_config SET environment=?, api_base_url=?, username=?, password_enc=?, org_id=?, dev_key_enc=?, default_ap_account=?, default_cash_account=?, default_clearing_account=?, sync_cutoff_date=?, updated_by=?, updated_at=? WHERE entity_id=?')
      .run(environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, clearingAcct, sync_cutoff_date || null, updater, now, req.params.entity_id);
  } else {
    db.prepare('INSERT INTO billcom_config (entity_id, environment, api_base_url, username, password_enc, org_id, dev_key_enc, default_ap_account, default_cash_account, default_clearing_account, sync_cutoff_date, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)')
      .run(req.params.entity_id, environment, baseUrl, username, pwEnc, org_id, keyEnc, default_ap_account || null, default_cash_account || null, clearingAcct, sync_cutoff_date || null, updater, now);
  }
  res.json({ success: true });
});

// Set ONLY the Bill.com sync cutoff date, without touching stored credentials.
// Driven by the A/P Aging "Upload aging detail" flow: the latest bill date on
// the uploaded GL/prior-system aging report is the last invoice already booked
// in the GL, so we skip anything dated on/before it. The sync engine treats the
// cutoff as EXCLUSIVE (a bill syncs when invoiceDate >= cutoff), so to exclude
// the last booked bill itself the caller stores latestBillDate + 1 day. Creating
// a fresh config row here would be missing required credentials, so this only
// updates an existing config.
app.put('/api/billcom/config/:entity_id/cutoff', auth, requireEntityAccess('entity_id'), requireRole('Admin','Accountant'), (req, res) => {
  const cutoff = (req.body && req.body.sync_cutoff_date) || null;
  if (cutoff !== null && !/^\d{4}-\d{2}-\d{2}$/.test(String(cutoff))) return res.status(400).json({ error: 'sync_cutoff_date must be YYYY-MM-DD or null' });
  const existing = db.prepare('SELECT entity_id FROM billcom_config WHERE entity_id = ?').get(req.params.entity_id);
  if (!existing) return res.status(400).json({ error: 'Bill.com is not configured for this entity yet. Set up the Bill.com connection first, then upload the A/P aging.' });
  // Optionally persist the parsed A/P aging lines alongside the cutoff so the
  // dedupe check can later skip Bill.com bills already booked in the GL. Each
  // line is normalized to {vendor, invoice_number, bill_date, amount}.
  let linesJson = null, asOf = null;
  if (Array.isArray(req.body && req.body.lines)) {
    const clean = req.body.lines.map(l => ({
      vendor: (l && l.vendor != null) ? String(l.vendor) : '',
      invoice_number: (l && (l.invoice_number != null ? l.invoice_number : l.document_no != null ? l.document_no : l.num)) != null ? String(l.invoice_number != null ? l.invoice_number : l.document_no != null ? l.document_no : l.num) : '',
      bill_date: (l && (l.bill_date != null ? l.bill_date : l.invoice_date != null ? l.invoice_date : l.date)) ? String(l.bill_date != null ? l.bill_date : l.invoice_date != null ? l.invoice_date : l.date).slice(0, 10) : null,
      amount: (l && l.amount != null && !isNaN(Number(l.amount))) ? Math.round(Number(l.amount) * 100) / 100 : null,
    })).filter(l => l.amount != null);
    linesJson = JSON.stringify(clean);
    asOf = (req.body.as_of && /^\d{4}-\d{2}-\d{2}$/.test(String(req.body.as_of))) ? String(req.body.as_of) : null;
  }
  const now = new Date().toISOString();
  const updater = req.user.name || req.user.email;
  if (linesJson !== null) {
    db.prepare('UPDATE billcom_config SET sync_cutoff_date=?, ap_aging_lines_json=?, ap_aging_as_of=?, ap_aging_uploaded_at=?, updated_by=?, updated_at=? WHERE entity_id=?')
      .run(cutoff, linesJson, asOf, now, updater, now, req.params.entity_id);
  } else {
    db.prepare('UPDATE billcom_config SET sync_cutoff_date=?, updated_by=?, updated_at=? WHERE entity_id=?')
      .run(cutoff, updater, now, req.params.entity_id);
  }
  res.json({ success: true, sync_cutoff_date: cutoff, aging_lines: linesJson !== null ? JSON.parse(linesJson).length : undefined });
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
      // OR IGNORE: two CL codes can name-match the same Bill.com account; the
      // first mapping wins and the duplicate is skipped rather than throwing UNIQUE.
      const ins = db.prepare('INSERT OR IGNORE INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)');
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
  // The Bill.com account list can contain duplicate ids, so the UI may submit
  // more than one row for the same billcom_account_id -> a UNIQUE(entity_id,
  // billcom_account_id) collision on save. Collapse to one row per id (last wins)
  // and INSERT OR REPLACE so a save can never throw the constraint error.
  const _byId = new Map();
  for (const m of mappings) { if (m && m.billcom_account_id) _byId.set(String(m.billcom_account_id), m); }
  const deduped = [..._byId.values()];
  const tx = db.transaction((rows) => {
    db.prepare('DELETE FROM billcom_account_map WHERE entity_id = ?').run(entityId);
    const ins = db.prepare('INSERT OR REPLACE INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)');
    for (const m of rows) {
      ins.run(entityId, String(m.billcom_account_id), m.billcom_account_name || null, String(m.cl_account_code), now);
    }
  });
  try {
    tx(deduped);
    res.json({ saved: deduped.length });
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
  const projects = db.prepare('SELECT billcom_dept_id, billcom_dept_name, cl_project_id FROM billcom_project_map WHERE entity_id = ? ORDER BY id').all(eid);
  res.json({ classes, locations, projects });
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

  let bcClasses, bcJobs, bcDepts;
  try { bcClasses = await billcomListClassification({ ...args, resource: 'accounting-classes' }); }
  catch (e) { return res.status(502).json({ error: 'fetch classes failed: ' + e.message }); }
  try { bcJobs = await billcomListClassification({ ...args, resource: 'jobs' }); }
  catch (e) { return res.status(502).json({ error: 'fetch jobs failed: ' + e.message }); }
  // Departments carry the PROJECT for Banyan (confirmed by Jimmy 2026-08-19).
  // Tolerate a department fetch failure: an org that does not use departments
  // should still be able to auto-map its classes and jobs.
  try { bcDepts = await billcomListClassification({ ...args, resource: 'departments' }); }
  catch (e) { bcDepts = null; }

  // CL dimensions for this entity (class = investor, location = deal).
  const clClasses = db.prepare('SELECT id, name FROM dim_classes WHERE entity_id = ?').all(eid);
  const clLocs = db.prepare('SELECT id, name FROM dim_locations WHERE entity_id = ?').all(eid);
  const clProjs = db.prepare('SELECT id, code, name FROM dim_projects WHERE entity_id = ?').all(eid);
  const norm = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const clClassByName = new Map(clClasses.map(c => [norm(c.name), c.id]));
  const clLocByName = new Map(clLocs.map(l => [norm(l.name), l.id]));
  // A project matches on its name, its code, or the combined "code - name" label,
  // each with a leading % stripped (Bill.com departments arrive as e.g. "%Van
  // Buren"). A key that two different projects both answer to is AMBIGUOUS and
  // maps to nothing - Banyan has "%Van Buren" and "Van Buren" as separate
  // projects, and silently picking one by query order is how a bill lands on the
  // wrong job. Ambiguous names are reported for a human to resolve, exactly as the
  // chart-of-accounts auto-map does.
  const clProjByKey = new Map();
  const addProjKey = (k, id) => {
    const kk = norm(String(k || '').replace(/^%+/, ''));
    if (!kk) return;
    const cur = clProjByKey.get(kk);
    if (cur === undefined) clProjByKey.set(kk, id);
    else if (cur !== id && cur !== '(AMBIGUOUS)') clProjByKey.set(kk, '(AMBIGUOUS)');
  };
  for (const p of clProjs) {
    addProjKey(p.name, p.id);
    addProjKey(p.code, p.id);
    if (p.code && p.name) { addProjKey(p.code + ' - ' + p.name, p.id); addProjKey(p.code + ' ' + p.name, p.id); }
  }
  const nameOf = (o) => o.name || o.shortName || o.description || '';
  const idOf = (o) => String(o.id || '');

  const now = new Date().toISOString();
  const result = { classes: { matched: [], unmatched: [] }, locations: { matched: [], unmatched: [] }, projects: { matched: [], unmatched: [] } };

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
    // Departments -> projects. Only rebuild the map when the fetch succeeded; a
    // failed fetch must not wipe mappings the user already has.
    if (Array.isArray(bcDepts)) {
      // Snapshot what is mapped today so a hand-made mapping is not lost to a
      // re-run: anything the auto pass cannot match is restored below.
      const priorP = new Map(db.prepare('SELECT billcom_dept_id, billcom_dept_name, cl_project_id FROM billcom_project_map WHERE entity_id = ?').all(eid).map(r => [String(r.billcom_dept_id), r]));
      db.prepare('DELETE FROM billcom_project_map WHERE entity_id = ?').run(eid);
      const insP = db.prepare('INSERT INTO billcom_project_map (entity_id, billcom_dept_id, billcom_dept_name, cl_project_id, created_at) VALUES (?,?,?,?,?)');
      const validProj = new Set(clProjs.map(p => p.id));
      for (const d of bcDepts) {
        const nm = nameOf(d);
        const did = idOf(d);
        const hit = clProjByKey.get(norm(String(nm).replace(/^%+/, '')));
        if (hit && hit !== '(AMBIGUOUS)') {
          insP.run(eid, did, nm, hit, now);
          result.projects.matched.push({ billcom: nm, cl_project_id: hit });
          continue;
        }
        // No confident match. Keep whatever a human had set for this department.
        const kept = priorP.get(did);
        if (kept && validProj.has(kept.cl_project_id)) {
          insP.run(eid, did, nm || kept.billcom_dept_name, kept.cl_project_id, now);
          result.projects.matched.push({ billcom: nm, cl_project_id: kept.cl_project_id, kept_manual: true });
          continue;
        }
        result.projects.unmatched.push({ billcom_dept_id: did, name: nm, ambiguous: hit === '(AMBIGUOUS)' });
      }
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
  const projects = Array.isArray(req.body && req.body.projects) ? req.body.projects : [];
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
      const upP = db.prepare('INSERT INTO billcom_project_map (entity_id, billcom_dept_id, billcom_dept_name, cl_project_id, created_at) VALUES (?,?,?,?,?) ON CONFLICT(entity_id, billcom_dept_id) DO UPDATE SET cl_project_id=excluded.cl_project_id, billcom_dept_name=excluded.billcom_dept_name');
      const delP = db.prepare('DELETE FROM billcom_project_map WHERE entity_id = ? AND billcom_dept_id = ?');
      for (const p of projects) {
        if (!p.billcom_dept_id) continue;
        if (p.cl_project_id == null) delP.run(eid, String(p.billcom_dept_id));
        else upP.run(eid, String(p.billcom_dept_id), p.billcom_dept_name || null, parseInt(p.cl_project_id), now);
      }
    });
    tx();
    const classCount = db.prepare('SELECT COUNT(*) c FROM billcom_class_map WHERE entity_id = ?').get(eid).c;
    const locCount = db.prepare('SELECT COUNT(*) c FROM billcom_location_map WHERE entity_id = ?').get(eid).c;
    const projCount = db.prepare('SELECT COUNT(*) c FROM billcom_project_map WHERE entity_id = ?').get(eid).c;
    res.json({ ok: true, class_map_rows: classCount, location_map_rows: locCount, project_map_rows: projCount });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Retag synced Bill.com entries with the project their Bill.com DEPARTMENT maps
// to. Body: { from?: 'YYYY-MM-DD', to?: 'YYYY-MM-DD', dry_run?: true }.
// Only fills gaps - a line that already carries a project, or an entry that
// already has a doc number, is left alone - so it is safe to re-run.
app.post('/api/billcom/retag-projects/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
 try {
  // `pick` is a local helper in every function that uses it, not a module-level
  // one. Reaching for it here without declaring it threw ReferenceError, and an
  // unhandled rejection in an async handler kills the process.
  const pick = (obj, ...keys) => { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; };
  const eid = parseInt(req.params.entity_id);
  const dryRun = !!(req.body && req.body.dry_run);
  const from = (req.body && req.body.from) || null;
  const to = (req.body && req.body.to) || null;
  // Each bill costs one sequential Bill.com round trip, so a month's worth blows
  // through Railway's gateway timeout. Work in bounded batches and hand back a
  // cursor; the caller loops until remaining is 0.
  const limit = Math.max(1, Math.min(parseInt((req.body && req.body.limit) || 20, 10) || 20, 100));
  const offset = Math.max(0, parseInt((req.body && req.body.offset) || 0, 10) || 0);
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(eid);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });
  const projMap = new Map(db.prepare('SELECT billcom_dept_id, cl_project_id FROM billcom_project_map WHERE entity_id = ?').all(eid).map(r => [String(r.billcom_dept_id), r.cl_project_id]));
  if (!projMap.size) return res.status(400).json({ error: 'No department -> project mappings yet. Run POST /api/billcom/dimension-maps/' + eid + '/auto first.' });

  let where = '';
  const params = [eid];
  if (from) { where += ' AND je.date >= ?'; params.push(from); }
  if (to) { where += ' AND je.date <= ?'; params.push(to); }
  const allSynced = db.prepare(
    "SELECT bl.billcom_id, bl.cl_entry_id, bl.invoice_number, je.date, je.entry_num, je.doc_number " +
    "FROM billcom_sync_log bl JOIN journal_entries je ON je.id = bl.cl_entry_id " +
    "WHERE bl.entity_id = ? AND bl.sync_type = 'bill' AND bl.status = 'success' " +
    "AND bl.cl_entry_id IS NOT NULL" + where + " ORDER BY je.date, je.entry_num"
  ).all(...params);
  const total = allSynced.length;
  const synced = allSynced.slice(offset, offset + limit);
  if (!synced.length) return res.json({ dry_run: dryRun, total_in_window: total, offset, examined: 0, entries_tagged: 0, lines_tagged: 0, docs_filled: 0, next_offset: null, remaining: 0, details: [] });

  let session, devKey;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) { return res.status(502).json({ error: 'Bill.com login failed: ' + e.message }); }
  const args = { sessionId: session.sessionId, devKey, baseUrl: cfg.api_base_url };

  const nextOffset = offset + synced.length;
  const result = { dry_run: dryRun, total_in_window: total, offset, examined: synced.length, next_offset: nextOffset < total ? nextOffset : null, remaining: Math.max(0, total - nextOffset), entries_tagged: 0, lines_tagged: 0, docs_filled: 0, skipped_no_department: 0, errors: [], details: [] };
  const untaggedLines = db.prepare('SELECT id, account_code, debit, credit, project_id, description FROM journal_lines WHERE entry_id = ? ORDER BY id');
  const setProject = db.prepare('UPDATE journal_lines SET project_id = ? WHERE id = ?');
  const setDesc = db.prepare('UPDATE journal_lines SET description = ? WHERE id = ?');
  const setDoc = db.prepare('UPDATE journal_entries SET doc_number = ? WHERE id = ? AND entity_id = ?');

  for (const row of synced) {
    let detail;
    try {
      detail = await billcomGetById({ ...args, resourcePath: '/bills', id: row.billcom_id });
    } catch (e) { result.errors.push({ billcom_id: row.billcom_id, entry_num: row.entry_num, error: e.message }); continue; }
    const lineItems = pick(detail, 'lineItems', 'line_items', 'billLineItems') || [];
    if (!Array.isArray(lineItems) || !lineItems.length) { result.errors.push({ billcom_id: row.billcom_id, entry_num: row.entry_num, error: 'no line items' }); continue; }
    // Bill.com line order is stable and the sync wrote one debit line per line
    // item in that order, so pair them positionally over the DEBIT lines only.
    // The AP credit line is deliberately left undimensioned.
    const clLines = untaggedLines.all(row.cl_entry_id).filter(l => (l.debit || 0) > 0);
    const plan = [];
    let sawDept = false;
    lineItems.forEach((li, i) => {
      const cls = pick(li, 'classifications') || {};
      const deptId = String(pick(cls, 'departmentId', 'deptId', 'department') || pick(li, 'departmentId', 'deptId') || '');
      if (deptId) sawDept = true;
      const projId = deptId ? (projMap.get(deptId) || null) : null;
      const clLine = clLines[i];
      if (!clLine) return;
      const desc = String(pick(li, 'description', 'memo', 'lineItemDescription') || '').trim();
      if (projId && !clLine.project_id) plan.push({ line_id: clLine.id, project_id: projId, description: (!clLine.description && desc) ? desc : null });
      else if (desc && !clLine.description) plan.push({ line_id: clLine.id, project_id: null, description: desc });
    });
    if (!sawDept) result.skipped_no_department++;
    const docToFill = (!row.doc_number && row.invoice_number) ? String(row.invoice_number).trim() : null;
    if (!plan.length && !docToFill) continue;
    if (!dryRun) {
      db.transaction(() => {
        for (const p of plan) {
          if (p.project_id) setProject.run(p.project_id, p.line_id);
          if (p.description) setDesc.run(p.description, p.line_id);
        }
        if (docToFill) setDoc.run(docToFill, row.cl_entry_id, eid);
      })();
    }
    const tagged = plan.filter(p => p.project_id).length;
    if (tagged) { result.entries_tagged++; result.lines_tagged += tagged; }
    if (docToFill) result.docs_filled++;
    if (result.details.length < 100) {
      result.details.push({ entry_num: row.entry_num, date: row.date, billcom_id: row.billcom_id, invoice_number: row.invoice_number, lines_tagged: tagged, doc_number_set: docToFill || null });
    }
  }
  res.json(result);
 } catch (e) {
  console.error('retag-projects failed: ' + (e && e.stack || e));
  if (!res.headersSent) res.status(500).json({ error: e.message });
 }
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
          db.prepare('INSERT OR IGNORE INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)')
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
      db.prepare('INSERT OR IGNORE INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)')
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
    'SELECT id, sync_type, billcom_id, cl_entry_id, status, message, created_at, invoice_number FROM billcom_sync_log WHERE entity_id = ? ORDER BY id DESC LIMIT ?'
  ).all(entityId, limit);
  res.json({ logs: rows });
});

// Un-sync: remove every CloudLedger journal entry that a Bill.com sync created
// for this entity, and clear the entity's sync log so a subsequent (corrected)
// sync re-pulls from scratch. Scoped STRICTLY to entries recorded in
// billcom_sync_log with a cl_entry_id — GL-import entries and manual JEs are
// never touched because they have no sync-log row. Use case: a sync ran with
// the wrong cutoff and duplicated invoices already present from a GL import.
app.post('/api/billcom/unsync/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  const dryRun = !!(req.body && req.body.dry_run);
  // Every JE this entity's sync created, via the authoritative link column.
  const linked = db.prepare(
    "SELECT DISTINCT cl_entry_id FROM billcom_sync_log WHERE entity_id = ? AND cl_entry_id IS NOT NULL"
  ).all(entityId).map(r => r.cl_entry_id);
  // Only those that still exist as JEs on THIS entity (defensive: never delete
  // an id that isn't actually this entity's journal entry).
  const existing = linked.filter(id => db.prepare('SELECT 1 FROM journal_entries WHERE id = ? AND entity_id = ?').get(id, entityId));
  if (dryRun) {
    return res.json({ dry_run: true, entity_id: entityId, would_delete_entries: existing.length, sync_log_rows: db.prepare('SELECT COUNT(*) c FROM billcom_sync_log WHERE entity_id = ?').get(entityId).c });
  }
  let deleted = 0;
  const tx = db.transaction(() => {
    for (const id of existing) {
      const atts = db.prepare('SELECT filename FROM journal_attachments WHERE entry_id = ?').all(id);
      atts.forEach(a => { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} });
      db.prepare('DELETE FROM journal_entries WHERE id = ? AND entity_id = ?').run(id, entityId);
      deleted++;
    }
    // Clear the sync log so dedup doesn't block a corrected re-sync.
    db.prepare('DELETE FROM billcom_sync_log WHERE entity_id = ?').run(entityId);
  });
  tx();
  res.json({ success: true, entity_id: entityId, deleted_entries: deleted, sync_log_cleared: true });
});

// posts the QBO-style clearing-account flow and returns a structured result.
// Used by BOTH the Sync endpoint (payment phase) and the payment-reconcile
// endpoint, so the accounting is identical and dedup can never double-post:
//   Leg 1 (per bill relieved): Dr AP / Cr Clearing (1072) at the payment's
//          processDate — relieves the open payable, ages the bill out of AP.
//   Leg 2 (per processDate, lump sum): Dr Clearing (1072) / Cr Cash — one
//          funds-transfer entry per date, matching the batch ACH on the bank.
// Only truly disbursed payments relieve (PAID, not voided/cancelled, processDate
// within (cutoff, asOf]). A payment whose bill has no synced CL JE is skipped
// (relieving it would double-count against the opening balance). Idempotent via
// billcom_sync_log: sync_type 'payment' keyed payId:billId, 'funds_transfer'
// keyed ft:processDate. Safe to re-run and safe to call from either endpoint.
function performPaymentReconcileCore({ entityId, apAccount, clearingAccount, cashAccount, payments, asOf, cutoffDate, dryRun, actor, now, billInvoiceDateById }) {
  const pick = (o, ...ks) => { for (const k of ks) if (o && o[k] != null) return o[k]; return null; };
  // Optional map of billcom bill id -> invoice date, built by the caller from the
  // bills fetched in the same run. Lets us exclude a payment by its INVOICE date
  // (the conversion rule: invoices on/before the last JE-entered invoice are
  // already in CL and must not be re-imported), independent of when it was paid.
  const invDateOf = (billId) => (billInvoiceDateById && billInvoiceDateById.get(String(billId))) || null;

  // billId -> CL bill JE id (only bills actually synced into CL can be relieved).
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
    'INSERT INTO billcom_sync_log (entity_id, sync_type, billcom_id, cl_entry_id, status, message, created_at, invoice_number) VALUES (?,?,?,?,?,?,?,?)'
  );

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

  const transferByDate = new Map(); // processDate -> NEW disbursed amount this run
  const insertJE = (date, memo, lines) => {
    const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
    const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
      .run(entityId, num, date, memo, 'Bill.com sync');
    for (const l of lines) {
      db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit) VALUES (?,?,?,?)')
        .run(r.lastInsertRowid, l.account_code, l.debit, l.credit);
    }
    return { id: r.lastInsertRowid, num };
  };

  // ── Leg 1: relieve each bill settled by a disbursed payment.
  for (const pay of payments) {
    const payId = String(pick(pay, 'id') || '');
    if (!payId) continue;
    const processDate = String(pick(pay, 'processDate', 'process_date', 'paymentDate') || '');
    const payNum = pick(pay, 'transactionNumber', 'confirmationNumber', 'paymentNumber') || payId;

    if (!isDisbursed(pay)) {
      result.leg1.skipped++;
      result.leg1.details.push({ id: payId, status: 'skip', reason: 'not disbursed/in-window', processDate, payStatus: pick(pay, 'status', 'paymentStatus') });
      continue;
    }

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

      // Conversion-cutoff rule (by INVOICE date, not payment date): if the bill
      // this payment settles is dated on/before the cutoff, the invoice is already
      // in CL (entered via JE pre-conversion) and its payoff belongs to the opening
      // state — so skip the payment no matter when it was actually paid. cutoffDate
      // is exclusive: invoiceDate < cutoffDate means pre-conversion.
      const invDate = invDateOf(billId);
      if (invDate && String(invDate) < cutoffDate) {
        result.leg1.skipped++;
        result.leg1.details.push({ id: payId + ':' + billId, status: 'skip', reason: 'invoice pre-conversion (' + invDate + ' < ' + cutoffDate + ')' });
        continue;
      }

      const dedupId = payId + ':' + billId;
      if (alreadySynced.get(entityId, 'payment', dedupId)) {
        result.leg1.skipped++;
        result.leg1.details.push({ id: dedupId, status: 'skip', reason: 'already reconciled' });
        continue;
      }

      const billEntryId = billEntryByBillcomId.get(billId);
      if (!billEntryId) {
        result.leg1.skipped++;
        result.leg1.details.push({ id: dedupId, status: 'skip', reason: 'bill not synced to CL (likely pre-cutover)', billId });
        continue;
      }

      const memo = 'Bill.com payment ' + payNum + ' \u2014 relieve bill ' + billId;
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
            logSync.run(entityId, 'payment', dedupId, created.id, 'success', 'relieved bill ' + billId + ' (JE #' + created.num + ')', now, null);
            return created;
          })();
          result.leg1.relieved++; result.leg1.amount += amount;
          result.leg1.details.push({ id: dedupId, status: 'created', cl_entry_id: je.id, date: processDate, amount });
          transferByDate.set(processDate, (transferByDate.get(processDate) || 0) + amount);
        } catch (e) {
          result.leg1.errors++;
          logSync.run(entityId, 'payment', dedupId, null, 'error', e.message, now, null);
          result.leg1.details.push({ id: dedupId, status: 'error', reason: e.message });
        }
      }
    }
  }

  // ── Leg 2: one funds-transfer JE per processDate with NEW relief this run.
  for (const [pd, amount] of transferByDate.entries()) {
    if (amount <= 0.005) continue;
    const dedupId = 'ft:' + pd;
    const delta = amount; // only NEW relief from this run
    if (delta <= 0.005) { result.leg2.skipped++; continue; }

    const memo = 'Bill.com funds transfer \u2014 ' + pd + ' batch';
    const lines = [
      { account_code: clearingAccount, debit: delta, credit: 0 },
      { account_code: cashAccount, debit: 0, credit: delta },
    ];

    if (dryRun) {
      result.leg2.transfers++; result.leg2.amount += delta;
      result.leg2.details.push({ date: pd, status: 'would_create', amount: delta });
    } else {
      try {
        const je = db.transaction(() => {
          const created = insertJE(pd, memo, lines);
          logSync.run(entityId, 'funds_transfer', dedupId, created.id, 'success', 'funds transfer ' + pd + ' $' + delta.toFixed(2) + ' (JE #' + created.num + ')', now, null);
          return created;
        })();
        result.leg2.transfers++; result.leg2.amount += delta;
        result.leg2.details.push({ date: pd, status: 'created', cl_entry_id: je.id, amount: delta });
      } catch (e) {
        result.leg2.errors++;
        logSync.run(entityId, 'funds_transfer', dedupId, null, 'error', e.message, now, null);
        result.leg2.details.push({ date: pd, status: 'error', reason: e.message });
      }
    }
  }

  result.leg1.amount = Math.round(result.leg1.amount * 100) / 100;
  result.leg2.amount = Math.round(result.leg2.amount * 100) / 100;
  return result;
}

// ─── A/P aging dedupe matcher ───
// Decide whether a Bill.com bill is already booked in the GL by matching it
// against a line from the last uploaded A/P aging detail. Per spec, amount must
// always agree; identity is then confirmed by invoice number OR vendor+date.
// GL-import aging rows carry neither a vendor nor an invoice number (only a JE
// number + date + amount), so those fall back to date+amount — safe because the
// sync cutoff already excludes anything dated on/before the aging's latest date,
// so a bill that reaches this check is dated later and won't spuriously match.
function billcomNormKey(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function billcomAmtEq(a, b) { return a != null && b != null && Math.abs(Number(a) - Number(b)) < 0.005; }
function billcomDateEq(a, b) { return !!a && !!b && String(a).slice(0, 10) === String(b).slice(0, 10); }
function matchApAgingLine(lines, bill) {
  if (!Array.isArray(lines) || !lines.length) return null;
  const bnum = billcomNormKey(bill.number);
  const bven = billcomNormKey(bill.vendor);
  const bamt = bill.amount;
  const bdate = bill.date;
  for (const ln of lines) {
    if (!billcomAmtEq(bamt, ln.amount)) continue; // amount is required in every case
    const lnum = billcomNormKey(ln.invoice_number);
    const lven = billcomNormKey(ln.vendor);
    if (bnum && lnum && bnum === lnum) return { matched_on: 'invoice number + amount', line: ln };
    if (bven && lven && bven === lven && billcomDateEq(bdate, ln.bill_date)) return { matched_on: 'vendor + date + amount', line: ln };
    if (!lnum && !lven && billcomDateEq(bdate, ln.bill_date)) return { matched_on: 'date + amount', line: ln };
  }
  return null;
}

// Manual dry-run: report Bill.com bills that would sync (dated >= cutoff) but are
// already present in the last uploaded A/P aging detail. No JEs are created. This
// is what the "Check against A/P aging" button calls; the same matching runs
// automatically (and auto-skips) inside the sync itself.
app.post('/api/billcom/ap-aging-check/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  const entityId = parseInt(req.params.entity_id);
  if (!entityId) return res.status(400).json({ error: 'Invalid entity_id' });
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  if (!cfg) return res.status(400).json({ error: 'Bill.com not configured for this entity' });
  let agingLines = [];
  try { agingLines = cfg.ap_aging_lines_json ? JSON.parse(cfg.ap_aging_lines_json) : []; } catch (e) { agingLines = []; }
  if (!agingLines.length) return res.json({ ok: true, aging_uploaded: false, aging_as_of: null, aging_lines: 0, checked_bills: 0, overlaps: [], message: 'No A/P aging has been uploaded for this entity yet. Upload one from the A/P Aging report first.' });
  const pick = (obj, ...keys) => { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; };
  const cutoffDate = String((req.body && req.body.cutoff_date) || cfg.sync_cutoff_date || '2026-01-01');
  let session;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    const devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) { return res.status(502).json({ error: 'Bill.com login failed: ' + e.message }); }
  const listArgs = { sessionId: session.sessionId, devKey: billcomDecrypt(cfg.dev_key_enc), baseUrl: cfg.api_base_url };
  const windowFrom = (cutoffDate.slice(0, 7) + '-01');
  const windowTo = (() => { const d = new Date(); d.setMonth(d.getMonth() + 1); return d.toISOString().slice(0, 10); })();
  let bills;
  try { bills = await billcomListBillsWindowed({ ...listArgs, fromDate: windowFrom, toDate: windowTo }); }
  catch (e) { return res.status(502).json({ error: 'Failed to fetch bills: ' + e.message }); }
  const vendorById = new Map();
  try { const vlist = await billcomListVendors({ ...listArgs, maxItems: 5000 }); for (const v of vlist) { const id = String(pick(v, 'id') || ''); const n = pick(v, 'name', 'vendorName', 'companyName'); if (id && n) vendorById.set(id, n); } } catch (e) {}
  const vendorOf = (obj) => vendorById.get(String(pick(obj, 'vendorId', 'vendor_id') || (pick(obj, 'vendor') || {}).id || '')) || pick(pick(obj, 'vendor') || {}, 'name', 'vendorName') || '';
  let checked = 0;
  const overlaps = [];
  for (const bill of bills) {
    const billId = String(pick(bill, 'id') || '');
    if (!billId) continue;
    const a = String(pick(bill, 'approvalStatus', 'status') || '').toUpperCase();
    if (a === 'DENIED') continue;
    const date = pick(bill, 'invoiceDate', 'invoice_date', 'dueDate') || pick(pick(bill, 'invoice') || {}, 'invoiceDate', 'invoice_date');
    if (date && String(date) < cutoffDate) continue; // only bills that would actually sync
    checked++;
    const number = pick(bill, 'invoiceNumber', 'invoice_number') || pick(pick(bill, 'invoice') || {}, 'invoiceNumber', 'invoice_number') || billId;
    const amount = Number(pick(bill, 'amount', 'amountDue', 'invoiceAmount') || 0) || null;
    const hit = matchApAgingLine(agingLines, { number, date, vendor: vendorOf(bill), amount });
    if (hit) overlaps.push({ id: billId, invoice_number: number, date: date || null, vendor: vendorOf(bill) || '', amount, matched_on: hit.matched_on });
  }
  res.json({ ok: true, aging_uploaded: true, aging_as_of: cfg.ap_aging_as_of || null, aging_lines: agingLines.length, cutoff_date: cutoffDate, checked_bills: checked, overlap_count: overlaps.length, overlaps });
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
  const clearingAccount = cfg.default_clearing_account;
  // Bills require only the AP account. The payment (Money Out Clearing) leg
  // additionally needs the cash + clearing accounts; when either is missing or
  // doesn't exist we skip JUST the payment leg (with a note) instead of failing
  // the whole sync — so a partially-configured entity still gets its bills.
  // Applies to every entity, not just Turnkey.
  if (!apAccount) return res.status(400).json({ error: 'default_ap_account not set in Bill.com config' });
  const apExists = db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, apAccount);
  if (!apExists) return res.status(400).json({ error: 'AP account ' + apAccount + ' does not exist on entity' });
  const cashExists = cashAccount && db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, cashAccount);
  const clearingExists = clearingAccount && db.prepare('SELECT 1 FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, clearingAccount);
  // Reason the payment leg can't run (if any). Bills sync regardless.
  let paymentSkipReason = null;
  if (!clearingAccount) paymentSkipReason = 'Money Out Clearing (default_clearing_account) is not set in Bill.com config';
  else if (!clearingExists) paymentSkipReason = 'Clearing account ' + clearingAccount + ' does not exist on this entity';
  else if (!cashAccount) paymentSkipReason = 'default_cash_account is not set in Bill.com config';
  else if (!cashExists) paymentSkipReason = 'Cash account ' + cashAccount + ' does not exist on this entity';

  const mapRows = db.prepare('SELECT billcom_account_id, billcom_account_name, cl_account_code FROM billcom_account_map WHERE entity_id = ?').all(entityId);
  const mapById = new Map(mapRows.map(r => [String(r.billcom_account_id), r]));
  // Auto-map support: index this entity's GL accounts by normalized name so a
  // Bill.com line whose account name EXACTLY and UNAMBIGUOUSLY matches one GL
  // account can be mapped automatically (no manual mapping step). Names are
  // normalized (lowercase, strip non-alphanumerics) so "Telephone & Internet"
  // matches "Telephone and Internet". A name shared by 2+ GL accounts is left
  // ambiguous and NOT auto-mapped — it falls through to the error path so the
  // accountant decides. Newly auto-mapped ids are persisted so the sync only
  // resolves each once and the mapping shows up in the Account Mapping tab.
  const glNorm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const glByName = new Map(); // normName -> cl_account_code, or the string '(AMBIGUOUS)'
  for (const a of db.prepare('SELECT code, name FROM accounts WHERE entity_id = ?').all(entityId)) {
    const key = glNorm(a.name);
    if (!key) continue;
    if (glByName.has(key)) glByName.set(key, '(AMBIGUOUS)');
    else glByName.set(key, a.code);
  }
  const insAutoMap = db.prepare('INSERT OR IGNORE INTO billcom_account_map (entity_id, billcom_account_id, billcom_account_name, cl_account_code, created_at) VALUES (?,?,?,?,?)');
  // Dimension maps: only mapped Bill.com class/job ids carry a CL class/location
  // onto the synced line; unmapped ids (incl. workflow-status classes) -> null.
  const classMap = new Map(db.prepare('SELECT billcom_class_id, cl_class_id FROM billcom_class_map WHERE entity_id = ?').all(entityId).map(r => [String(r.billcom_class_id), r.cl_class_id]));
  const locMap = new Map(db.prepare('SELECT billcom_job_id, cl_location_id FROM billcom_location_map WHERE entity_id = ?').all(entityId).map(r => [String(r.billcom_job_id), r.cl_location_id]));
  // departmentId -> CL project. Banyan enters the project as the Bill.com
  // Department (confirmed by Jimmy 2026-08-19), so this is what puts a synced
  // bill line onto a project-scoped report.
  const projMap = new Map(db.prepare('SELECT billcom_dept_id, cl_project_id FROM billcom_project_map WHERE entity_id = ?').all(entityId).map(r => [String(r.billcom_dept_id), r.cl_project_id]));
  // Departments seen on synced lines that have no project mapping, so the sync
  // report can say "map this department to a project" instead of silently
  // posting an untagged line. Keyed by department id.
  const unmappedDepts = new Map();

  let session;
  try {
    const password = billcomDecrypt(cfg.password_enc);
    const devKey = billcomDecrypt(cfg.dev_key_enc);
    session = await billcomLogin({ username: cfg.username, password, orgId: cfg.org_id, devKey, baseUrl: cfg.api_base_url });
  } catch (e) {
    return res.status(502).json({ error: 'Bill.com login failed: ' + e.message });
  }
  const listArgs = { sessionId: session.sessionId, devKey: billcomDecrypt(cfg.dev_key_enc), baseUrl: cfg.api_base_url };

  // Bill.com account-id -> account-name lookup, for auto-mapping by name.
  // Bill.com bill line items DON'T carry the account name (only the id), so to
  // match a line's account against a GL account by name we must resolve the id
  // to its name via the chart-of-accounts list. Best-effort: if this fetch
  // fails, auto-mapping simply won't find names and the sync behaves as before.
  // Resolve Bill.com account-id -> name LAZILY. The full chart of accounts is
  // large and paginated (slow), and it is only needed to auto-map account ids
  // that are NOT already in billcom_account_map. Fetch it at most once, and only
  // when an unmapped id is actually encountered — so a fully-mapped entity's sync
  // never pays the COA download (which is what was making the Banyan sync 502).
  const bcNameById = new Map();
  let bcNamesLoaded = false;
  async function ensureBcNames() {
    if (bcNamesLoaded) return;
    bcNamesLoaded = true;
    try {
      console.log('[billcom sync] entity ' + entityId + ': loading Bill.com chart of accounts (unmapped account seen)');
      const bcAll = await billcomListAccounts(listArgs);
      for (const a of bcAll) { if (a && a.id != null) bcNameById.set(String(a.id), a.name || ''); }
      console.log('[billcom sync] entity ' + entityId + ': COA loaded, ' + bcNameById.size + ' accounts');
    } catch (e) { console.log('[billcom sync] account-name lookup failed: ' + e.message); }
  }

  // Bounded sync: process at most maxBills per invocation so the request always
  // returns well under Railway's gateway ceiling. Dedup via billcom_sync_log makes
  // repeated runs safe + incremental — re-run until processed === 0. Caller may
  // override with body.max_bills (clamped 1..1000; default 500 so a normal-size
  // entity finishes in a single pass and the setup fetches run once, not per batch).
  const maxBills = Math.max(1, Math.min(1000, parseInt((req.body && req.body.max_bills) || 500)));
  const deadline = Date.now() + 230000; // stop starting new work past ~3.8m, safely under the 300s gateway cap

  // Opening-balance cutoff: all balances before this date were booked via the
  // opening journal entry (from the GL detail report), so importing bills/payments
  // dated before it would double-count. Skip anything earlier. Configurable via
  // config.sync_cutoff_date or body.cutoff_date; defaults to 2026-01-01.
  const cutoffDate = String((req.body && req.body.cutoff_date) || cfg.sync_cutoff_date || '2026-01-01');
  // A/P aging dedupe: lines from the last uploaded A/P aging detail. Any bill
  // matching one of these is already booked in the GL, so it's auto-skipped and
  // reported. Empty when no aging has been uploaded (feature is a no-op then).
  let agingLines = [];
  try { agingLines = cfg.ap_aging_lines_json ? JSON.parse(cfg.ap_aging_lines_json) : []; } catch (e) { agingLines = []; }

  let bills, payments;
  // Bill.com v3 /bills and /payments IGNORE offset pagination — nextPage returns
  // the same first 100 rows repeatedly, so a plain paged fetch (billcomListBills)
  // silently caps at ~100 stale rows and never sees newer bills. That's why sync
  // reported "0 synced, 500 skipped": the 500 rows were the same old page fetched
  // five times, all already-synced/pre-cutoff, while June bills were never pulled.
  // Fetch via month-windowed filters instead (the method already proven for AP
  // aging), from the cutoff month through one month past today to catch any
  // future-dated due/process dates. Union+dedupe by id.
  const windowFrom = (cutoffDate.slice(0, 7) + '-01');
  const windowTo = (() => { const d = new Date(); d.setMonth(d.getMonth() + 1); return d.toISOString().slice(0, 10); })();
  console.log('[billcom sync] entity ' + entityId + ': cutoff ' + cutoffDate + ', fetching bills ' + windowFrom + '..' + windowTo);
  // Resolve each bill's GL Posting Date from Bill.com's legacy v2 API (v3 does not
  // expose it). Keyed back to v3 by id (v2 and v3 share bill ids), with an
  // invoiceNumber|amount fallback. Best-effort: if v2 is unavailable, fall back to
  // invoice date (prior behavior). Only glPostingDate >= windowFrom is fetched,
  // which covers every bill that could pass the cutoff.
  // EXCEPTION: these entities post by INVOICE DATE, not the GL posting date:
  //   40 = CLRF (County Line Rail Fund); 36 = Turnkey Rail (on CloudLedger from
  //   inception, no Intacct cutover, and GL posting date is only spottily set).
  // Every other Bill.com entity uses the GL posting date.
  const INVOICE_DATE_ENTITIES = new Set([40, 36]);
  const useGlPosting = !INVOICE_DATE_ENTITIES.has(entityId);
  let glMap = { byId: new Map(), byIdent: new Map(), identKey: (n, a) => String(n == null ? '' : n).trim() + '|' + (a == null ? '' : Number(a).toFixed(2)) };
  if (useGlPosting) {
    try {
      const v2pw = billcomDecrypt(cfg.password_enc);
      const v2dk = billcomDecrypt(cfg.dev_key_enc);
      const v2Session = await billcomV2Login({ username: cfg.username, password: v2pw, orgId: cfg.org_id, devKey: v2dk });
      const v2bills = await billcomV2ListBillsByGlPosting({ sessionId: v2Session, devKey: v2dk, fromDate: windowFrom });
      glMap = billcomBuildGlPostingMap(v2bills);
      console.log('[billcom sync] entity ' + entityId + ': v2 glPostingDate map built (' + v2bills.length + ' bills)');
    } catch (e) { console.log('[billcom sync] entity ' + entityId + ': v2 glPostingDate unavailable, using invoice date: ' + e.message); }
  } else {
    console.log('[billcom sync] entity ' + entityId + ': posts by invoice date (GL posting date not used)');
  }
  const glPostingFor = (bill) => {
    const id = String((bill && bill.id) || '');
    if (id && glMap.byId.has(id)) return glMap.byId.get(id);
    const num = (bill && bill.invoiceNumber) || (bill && bill.invoice && bill.invoice.invoiceNumber) || null;
    const amt = bill && bill.amount;
    return glMap.byIdent.get(glMap.identKey(num, amt)) || null;
  };
  try {
    bills = await billcomListBillsByUpdatedWindowed({ ...listArgs, fromDate: windowFrom, toDate: windowTo });
  } catch (e) {
    console.error('[billcom sync] entity ' + entityId + ' bills fetch failed: ' + e.message);
    return res.status(502).json({ error: 'Failed to fetch bills: ' + e.message });
  }
  console.log('[billcom sync] entity ' + entityId + ': ' + bills.length + ' bills fetched, fetching payments');
  try {
    payments = await billcomListPaymentsWindowed({ ...listArgs, fromDate: windowFrom, toDate: windowTo });
  } catch (e) {
    console.error('[billcom sync] entity ' + entityId + ' payments fetch failed: ' + e.message);
    return res.status(502).json({ error: 'Failed to fetch payments: ' + e.message });
  }
  console.log('[billcom sync] entity ' + entityId + ': ' + payments.length + ' payments fetched, processing (max ' + maxBills + ')');

  const pick = (obj, ...keys) => { for (const k of keys) if (obj && obj[k] != null) return obj[k]; return null; };
  // Bills: ONE APPROVAL SUFFICES. A bill syncs once at least one approver has
  // approved (Banyan policy, 8/2026) - it need not clear every approver or layer.
  // Pre-filter passes only APPROVED/APPROVING (both mean >=1 approver has approved);
  // ASSIGNED/UNASSIGNED (zero approvals) and DENIED are held here WITHOUT a detail
  // fetch. The per-approver check below is the authoritative gate. Period = GL posting date.
  const isBillEligible = (bill) => ['APPROVED', 'APPROVING'].includes(String(pick(bill, 'approvalStatus', 'status') || '').toUpperCase());
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
    'INSERT INTO billcom_sync_log (entity_id, sync_type, billcom_id, cl_entry_id, status, message, created_at, invoice_number) VALUES (?,?,?,?,?,?,?,?)'
  );
  const alreadySynced = db.prepare(
    "SELECT 1 FROM billcom_sync_log WHERE entity_id = ? AND sync_type = ? AND billcom_id = ? AND status = 'success' LIMIT 1"
  );

  // Vendor name comes straight off the bill object — the v3 bill (both the list
  // and the detail form) carries vendorName — so we no longer download the entire
  // vendor list (up to 5,000) on every run just to resolve names.
  const vendorOf = (obj) => pick(obj, 'vendorName', 'vendor_name') || pick(pick(obj, 'vendor') || {}, 'name', 'vendorName') || null;
  const backfillVendor = db.prepare("UPDATE journal_entries SET vendor = ? WHERE entity_id = ? AND memo = ? AND (vendor IS NULL OR vendor = '')");

  // Bills approved before the cutoff are a permanent skip. We persist that skip
  // (status 'skip_cutoff') the first time we determine it, so later batches can
  // skip them here CHEAPLY — before the per-batch budget gate and the expensive
  // per-bill detail fetch. Without this, an entity whose window is all pre-cutoff
  // bills re-fetches and re-scans the same 25 bills every batch and never advances.
  const alreadyPreCutoff = db.prepare("SELECT 1 FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill' AND billcom_id = ? AND status = 'skip_cutoff' LIMIT 1");

  // Front-load the per-bill detail fetches in parallel (bounded concurrency) so the
  // posting loop below isn't blocked one network round-trip at a time. Best-effort
  // CACHE only: the loop still applies every skip/gate and fetches inline on a miss,
  // so this changes WHEN details are fetched, never WHICH bills post. Candidates are
  // bills that pass the cheap read-only pre-checks, capped at the per-run budget.
  const detailCache = new Map();
  {
    const candidateIds = [];
    for (const b of bills) {
      const id = String(pick(b, 'id') || '');
      if (!id) continue;
      if (!isBillEligible(b)) continue;
      if (alreadySynced.get(entityId, 'bill', id)) continue;
      if (alreadyPreCutoff.get(entityId, id)) continue;
      candidateIds.push(id);
      if (candidateIds.length >= maxBills) break;
    }
    const CONC = 6;
    for (let i = 0; i < candidateIds.length && Date.now() < deadline; i += CONC) {
      const slice = candidateIds.slice(i, i + CONC);
      const settled = await Promise.all(slice.map(id =>
        billcomGetById({ ...listArgs, resourcePath: '/bills', id, extraParams: { billApprovals: 'true' } })
          .then(d => ({ id, d })).catch(() => ({ id, d: null }))
      ));
      for (const st of settled) if (st.d) detailCache.set(st.id, st.d);
    }
    if (candidateIds.length) console.log('[billcom sync] entity ' + entityId + ': prefetched ' + detailCache.size + '/' + candidateIds.length + ' bill details in parallel');
  }
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
      // Backfill the vendor on the existing JE if it's missing (covers bills
      // synced before vendor was captured). Uses the list object's invoice # +
      // vendor id; falls back to billId, matching the original memo.
      const vn = vendorOf(bill);
      const bnum = pick(bill, 'invoiceNumber', 'invoice_number') || billId;
      if (vn && bnum) { try { backfillVendor.run(vn, entityId, 'Bill.com bill #' + bnum); } catch (e) {} }
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'already synced' });
      continue;
    }
    if (alreadyPreCutoff.get(entityId, billId)) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'approved before cutoff (cached)' });
      continue;
    }
    // Invoice/list date is kept only for the A/P-aging dedupe below; the sync no
    // longer gates on invoice date. Gating is by APPROVAL date (per-bill, after the
    // detail fetch) so back-dated late invoices approved after the cutoff still sync.
    const listDate = pick(bill, 'invoiceDate', 'invoice_date', 'dueDate') || pick(pick(bill, 'invoice') || {}, 'invoiceDate', 'invoice_date');
    // A/P aging dedupe: skip bills already present in the uploaded A/P aging
    // (already booked in the GL). Done here on the list object — before the
    // per-run budget gate and the detail fetch — so overlaps neither consume the
    // batch budget nor need a detail round-trip. Reported in aging_overlaps.
    if (agingLines.length) {
      const listNumber = pick(bill, 'invoiceNumber', 'invoice_number') || pick(pick(bill, 'invoice') || {}, 'invoiceNumber', 'invoice_number') || billId;
      const listVendor = vendorOf(bill) || '';
      const listAmount = Number(pick(bill, 'amount', 'amountDue', 'invoiceAmount') || 0) || null;
      const agingHit = matchApAgingLine(agingLines, { number: listNumber, date: listDate, vendor: listVendor, amount: listAmount });
      if (agingHit) {
        result.bills.skipped++;
        result.bills.aging_skipped = (result.bills.aging_skipped || 0) + 1;
        (result.bills.aging_overlaps = result.bills.aging_overlaps || []).push({ id: billId, invoice_number: listNumber, date: listDate || null, vendor: listVendor, amount: listAmount, matched_on: agingHit.matched_on });
        result.bills.details.push({ id: billId, status: 'skip', reason: 'already in A/P aging (' + agingHit.matched_on + ')' });
        continue;
      }
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
      detail = detailCache.get(billId) || await billcomGetById({ ...listArgs, resourcePath: '/bills', id: billId, extraParams: { billApprovals: 'true' } });
    } catch (e) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'detail fetch failed: ' + e.message, now, null);
      result.bills.details.push({ id: billId, status: 'error', reason: 'detail fetch failed: ' + e.message });
      continue;
    }
    const invoiceDate = pick(detail, 'invoiceDate', 'invoice_date') || pick(pick(detail, 'invoice') || {}, 'invoiceDate', 'invoice_date') || pick(detail, 'dueDate');
    const glPostDay = glPostingFor(detail);
    const postingDate = glPostDay || invoiceDate;
    const postingBasis = glPostDay ? 'GL posting date' : 'invoice date';
    const billNumber = pick(detail, 'invoiceNumber', 'invoice_number') || pick(pick(detail, 'invoice') || {}, 'invoiceNumber', 'invoice_number') || billId;
    const lineItems = pick(detail, 'lineItems', 'line_items', 'billLineItems') || [];

    // Authoritative approval gate: AT LEAST ONE approver must have approved (Banyan
    // policy, 8/2026 - a single approval is sufficient; a bill need not clear every
    // approver or layer). The list-level isBillEligible check above is a cheap
    // pre-filter; this is the real gate, run on the detail object where the
    // per-approver list is populated (billApprovals=true). Held bills are skipped,
    // NOT errored, and NOT persisted - they should sync once an approval lands.
    if (!anyApproverApproved(detail)) {
      result.bills.skipped++;
      result.bills.details.push({ id: billId, status: 'skip', reason: 'awaiting approval (no approver has approved yet)' });
      continue;
    }

    if (!invoiceDate || !Array.isArray(lineItems) || lineItems.length === 0) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'missing invoiceDate or lineItems', now, billNumber);
      result.bills.details.push({ id: billId, status: 'error', reason: 'missing invoiceDate or lineItems' });
      continue;
    }

    // Gate on the bill's DOCUMENT (invoice) date, which is the basis the GL detail
    // import uses to book bills. Any bill dated on/before the cutoff is already in
    // the GL through that date, so skip it. (Previously gated on APPROVAL date,
    // which double-posted bills dated on/before the cutoff but approved afterward —
    // e.g. a 06/30 bill approved 07/31 was in the GL import AND re-synced.)
    const approvedDate = billApprovalDate(detail);
    const approvedDay = approvedDate ? String(approvedDate).slice(0, 10) : null;
    const postDay = postingDate ? String(postingDate).slice(0, 10) : null;
    if (!postDay || postDay <= cutoffDate) {
      result.bills.skipped++;
      // Persist so later batches skip this bill cheaply (before the budget gate
      // and detail fetch) instead of re-fetching it every run.
      try { logSync.run(entityId, 'bill', billId, null, 'skip_cutoff', 'GL posting date ' + (postDay || 'unknown') + ' on/before cutoff ' + cutoffDate + ' (already in GL)', now, billNumber); } catch (e) {}
      result.bills.details.push({ id: billId, status: 'skip', reason: 'GL posting date ' + (postDay || 'unknown') + ' on/before cutoff ' + cutoffDate + ' (already in GL)' });
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
      let mapping = mapById.get(acctId);
      // Only pay for the (slow) Bill.com COA name lookup when this id isn't
      // already mapped and we still need a name to auto-map it.
      if (!mapping) await ensureBcNames();
      // Prefer the name from the Bill.com chart-of-accounts lookup (keyed by id);
      // line items usually omit the name. Fall back to any name on the line.
      const bcName = bcNameById.get(acctId) || pick(li, 'chartOfAccountName', 'chart_of_account_name', 'accountName') || pick(cls, 'chartOfAccountName', 'chart_of_account_name', 'accountName') || '';
      if (!mapping) {
        // Level-2 auto-map: if the Bill.com account name matches exactly one GL
        // account on this entity, use it and remember the mapping. Ambiguous or
        // no match -> fall through to the missing-mapping error below.
        const glCode = bcName ? glByName.get(glNorm(bcName)) : undefined;
        if (glCode && glCode !== '(AMBIGUOUS)') {
          try { insAutoMap.run(entityId, acctId, bcName, glCode, new Date().toISOString()); } catch (e) {}
          mapping = { cl_account_code: glCode, billcom_account_name: bcName, _auto: true };
          mapById.set(acctId, mapping);
        }
      }
      if (!mapping) {
        const name = bcName || acctId;
        const norm = bcName ? glNorm(bcName) : '';
        const ambiguous = norm ? (glByName.get(norm) === '(AMBIGUOUS)') : false;
        billMissing.push({ id: acctId, name, ambiguous });
        continue;
      }
      const deptId = String(pick(cls, 'departmentId', 'deptId', 'department') || pick(li, 'departmentId', 'deptId') || '');
      const projId = deptId ? (projMap.get(deptId) || null) : null;
      if (deptId && !projId && !unmappedDepts.has(deptId)) unmappedDepts.set(deptId, { billcom_dept_id: deptId, affected_bills: 0 });
      if (deptId && !projId) unmappedDepts.get(deptId).affected_bills++;
      debitLines.push({
        account_code: mapping.cl_account_code, debit: amt, credit: 0,
        class_id: classMap.get(String(pick(cls, 'accountingClassId', 'classId') || '')) || null,
        location_id: locMap.get(String(pick(cls, 'jobId', 'locationId') || '')) || null,
        project_id: projId,
        // The line's own text, so the report's Description column reads like the
        // pre-sync imported entries did rather than repeating the bill number.
        description: String(pick(li, 'description', 'memo', 'lineItemDescription') || '').trim(),
      });
      totalDr += amt;
    }

    if (billMissing.length > 0) {
      result.bills.errors++;
      const missingNames = billMissing.map(m => m.name).join(', ');
      logSync.run(entityId, 'bill', billId, null, 'error', 'missing GL mapping(s) for: ' + missingNames, now, billNumber);
      result.bills.details.push({ id: billId, status: 'error', reason: 'missing mappings: ' + missingNames });
      for (const mm of billMissing) {
        const existing = missingMap.get(mm.id);
        if (existing) { existing.affected_bills++; continue; }
        // Enrich each unmapped account for the accountant-facing UI: the real
        // Bill.com account name, whether the name was ambiguous (matched 2+ GL
        // accounts), and the GL account list so the UI can offer a dropdown
        // pre-selected on a close name match. This turns the cryptic id error
        // into an actionable "map this name to this GL account" prompt.
        const norm = mm.name ? glNorm(mm.name) : '';
        let suggestion = null;
        if (norm && !mm.ambiguous) {
          const code = glByName.get(norm);
          if (code && code !== '(AMBIGUOUS)') {
            const acc = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ? AND code = ?').get(entityId, code);
            if (acc) suggestion = { cl_account_code: acc.code, cl_account_name: acc.name };
          }
        }
        missingMap.set(mm.id, {
          billcom_account_id: mm.id,
          billcom_account_name: mm.name,
          name: mm.name,
          ambiguous: !!mm.ambiguous,
          suggested_gl: suggestion,
          affected_bills: 1,
        });
      }
      continue;
    }

    if (totalDr <= 0) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'bill total is zero', now, billNumber);
      result.bills.details.push({ id: billId, status: 'error', reason: 'zero total' });
      continue;
    }

    const lines = [...debitLines, { account_code: apAccount, debit: 0, credit: totalDr, class_id: null, location_id: null, project_id: null, description: '' }];
    const billVendor = vendorOf(detail);
    // Memo reads like the entries the GL import produced: "Bill - <vendor>: <what it was for>".
    // The bill number itself now lives on journal_entries.doc_number, which is
    // what the reports show in their Doc column (CLA request, 8/2026).
    const _lineDesc = (debitLines.find(l => l.description) || {}).description || '';
    const memo = billVendor
      ? ('Bill - ' + billVendor + (_lineDesc ? ': ' + _lineDesc : ''))
      : ('Bill.com bill #' + billNumber);

    try {
      const insertedId = db.transaction(() => {
        const num = (db.prepare('SELECT MAX(entry_num) as m FROM journal_entries WHERE entity_id = ?').get(entityId).m || 0) + 1;
        const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, vendor, doc_number, created_by) VALUES (?,?,?,?,?,?,?)')
          .run(entityId, num, postingDate, memo, billVendor, String(billNumber || '') || null, 'Bill.com sync');
        for (const l of lines) {
          db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, class_id, location_id, project_id, description) VALUES (?,?,?,?,?,?,?,?)')
            .run(r.lastInsertRowid, l.account_code, l.debit, l.credit, l.class_id || null, l.location_id || null, l.project_id || null, l.description || '');
        }
        logSync.run(entityId, 'bill', billId, r.lastInsertRowid, 'success', 'created JE #' + num + ' (posted ' + String(postingDate).slice(0, 10) + ' by ' + postingBasis + ')', now, billNumber);
        return r.lastInsertRowid;
      })();
      result.bills.synced++;
      result.bills.details.push({ id: billId, status: 'success', cl_entry_id: insertedId });
    } catch (e) {
      result.bills.errors++;
      logSync.run(entityId, 'bill', billId, null, 'error', 'JE insert failed: ' + e.message, now, billNumber);
      result.bills.details.push({ id: billId, status: 'error', reason: e.message });
    }
  }

  // Deletion sync (per CLA request, 7/2026): when a bill is deleted in Bill.com it
  // should also be removed from CloudLedger. Bill.com's /bills list returns only
  // active bills, so any bill we previously synced (has a success row in
  // billcom_sync_log with a cl_entry_id) that is NOT in the current list has been
  // deleted upstream — we remove its journal entry (journal_lines cascade) and
  // record a 'bill_deleted' log row so re-runs are idempotent.
  //
  // Guard against false positives. The live fetch is now windowed by dueDate
  // ([windowFrom, windowTo]), so `bills` is the complete set of live bills WITHIN
  // that window — not the entire active list. Deletion mirroring must therefore be
  // scoped to the same window: only a synced bill whose invoice date falls inside
  // [windowFrom, windowTo] and is absent from the live window has truly been
  // deleted in Bill.com. A synced bill dated outside the window simply wasn't
  // fetched this run and must NOT be treated as deleted.
  // Unmapped departments are reported the same way unmapped GL accounts are, so
  // an untagged project is visible work rather than a silent omission.
  if (unmappedDepts.size) {
    result.bills.unmapped_departments = [...unmappedDepts.values()];
  }
  result.bills.deleted = 0;
  const liveIds = new Set(bills.map(b => String(pick(b, 'id') || '')).filter(Boolean));
  {
    // synced bills that still have a live CL entry, keyed by billcom_id -> cl_entry_id,
    // limited to those whose CL entry date is inside the fetched window.
    const syncedBills = db.prepare(
      "SELECT bl.billcom_id, bl.cl_entry_id, je.date AS entry_date " +
      "FROM billcom_sync_log bl JOIN journal_entries je ON je.id = bl.cl_entry_id " +
      "WHERE bl.entity_id = ? AND bl.sync_type = 'bill' AND bl.status = 'success' AND bl.cl_entry_id IS NOT NULL"
    ).all(entityId);
    // a bill already logged as deleted should not be re-processed
    const deletedIds = new Set(db.prepare(
      "SELECT billcom_id FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill_deleted' AND status = 'success'"
    ).all(entityId).map(r => String(r.billcom_id)));
    const entryStillExists = db.prepare('SELECT 1 FROM journal_entries WHERE id = ? AND entity_id = ?');
    const delEntry = db.prepare('DELETE FROM journal_entries WHERE id = ? AND entity_id = ?');
    const delAtts = db.prepare('SELECT filename FROM journal_attachments WHERE entry_id = ?');
    result.bills.deleted_details = [];
    for (const row of syncedBills) {
      const bid = String(row.billcom_id);
      if (liveIds.has(bid)) continue;        // still active in Bill.com
      if (deletedIds.has(bid)) continue;     // already handled on a prior run
      // Only consider deletion when this bill's CL entry date is inside the window
      // we actually fetched — otherwise its absence just means "not in this window".
      if (!(row.entry_date && String(row.entry_date) >= windowFrom && String(row.entry_date) < windowTo)) continue;
      // only delete if the CL entry actually exists (it may have been removed manually)
      if (!entryStillExists.get(row.cl_entry_id, entityId)) {
        logSync.run(entityId, 'bill_deleted', bid, row.cl_entry_id, 'success', 'bill absent in Bill.com; CL entry already gone', now, null);
        continue;
      }
      try {
        // remove any attachment files first, then the entry (lines cascade)
        for (const a of delAtts.all(row.cl_entry_id)) { try { fs.unlinkSync(path.join(UPLOAD_DIR, a.filename)); } catch {} }
        const info = delEntry.run(row.cl_entry_id, entityId);
        if (info.changes > 0) {
          result.bills.deleted++;
          logSync.run(entityId, 'bill_deleted', bid, row.cl_entry_id, 'success', 'bill deleted in Bill.com; removed CL entry ' + row.cl_entry_id, now, null);
          result.bills.deleted_details.push({ id: bid, cl_entry_id: row.cl_entry_id, status: 'deleted' });
        }
      } catch (e) {
        result.bills.errors++;
        logSync.run(entityId, 'bill_deleted', bid, row.cl_entry_id, 'error', 'delete failed: ' + e.message, now, null);
        result.bills.deleted_details.push({ id: bid, cl_entry_id: row.cl_entry_id, status: 'error', reason: e.message });
      }
    }
  }

  // ── Payments: route through the shared two-leg reconcile core so disbursements
  //    post through 1072 Money Out Clearing (Dr AP/Cr 1072, then Dr 1072/Cr Cash)
  //    instead of straight to cash. This replaces the old Dr AP/Cr Cash loop that
  //    (a) skipped 1072 and (b) errored "missing processDate or zero amount" on
  //    payment-list objects. The core reads processDate/billPayments off the
  //    fetched payments, relieves each linked bill, and posts one funds-transfer
  //    per date. It shares dedup keys (payment payId:billId, funds_transfer
  //    ft:date) with the payment-reconcile endpoint, so neither path double-posts.
  // Build billcom bill id -> invoice date from the bills fetched above, so the
  // core can exclude payments by invoice date (conversion-cutoff rule).
  const billInvoiceDateById = new Map();
  for (const b of bills) {
    const bid = pick(b, 'id');
    const idt = pick(b, 'invoiceDate', 'invoice_date', 'dueDate') || pick(pick(b, 'invoice') || {}, 'invoiceDate', 'invoice_date');
    if (bid != null && idt) billInvoiceDateById.set(String(bid), String(idt));
  }
  if (paymentSkipReason) {
    // Can't post the two-leg payment flow without cash + clearing; skip just the
    // payment leg so the bill sync still succeeds. Surface why in the result.
    result.payments.synced = 0;
    result.payments.errors = 0;
    result.payments.skipped = Array.isArray(payments) ? payments.length : 0;
    result.payments.details = [];
    result.payments.skip_reason = paymentSkipReason;
    result.payments.note = 'Payments were not synced — ' + paymentSkipReason + '. Bills synced normally; set the Money Out Clearing (and cash) account in Bill.com config to sync payments.';
  } else {
    const payResult = performPaymentReconcileCore({
      entityId, apAccount, clearingAccount, cashAccount,
      payments, asOf: new Date().toISOString().slice(0, 10), cutoffDate,
      dryRun: false, actor, now, billInvoiceDateById,
    });
    // Fold the two-leg result into the sync response shape the client expects
    // (synced = bill reliefs + funds transfers created; errors/skipped summed).
    result.payments.synced = payResult.leg1.relieved + payResult.leg2.transfers;
    result.payments.skipped = payResult.leg1.skipped + payResult.leg2.skipped;
    result.payments.errors = payResult.leg1.errors + payResult.leg2.errors;
    result.payments.details = [
      ...payResult.leg1.details.map(d => ({ ...d, leg: 'relieve' })),
      ...payResult.leg2.details.map(d => ({ ...d, leg: 'funds_transfer' })),
    ];
    result.payments.clearing_account = clearingAccount;
    result.payments.leg1_amount = payResult.leg1.amount;
    result.payments.leg2_amount = payResult.leg2.amount;
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

  // Fetch payments via the month-windowed method. Bill.com v3 /payments ignores
  // offset pagination (nextPage returns the same first 100 rows), so the old
  // maxItems:5000 fetch silently capped at ~100 and missed newer payments — the
  // same bug fixed on the bills side. Window from cutoff month through one month
  // past as_of to catch every disbursed payment in range.
  const windowFrom = (cutoffDate.slice(0, 7) + '-01');
  const windowTo = (() => { const d = new Date(asOf + 'T00:00:00'); d.setMonth(d.getMonth() + 1); return d.toISOString().slice(0, 10); })();
  let payments;
  try {
    payments = await billcomListPaymentsWindowed({ ...listArgs, fromDate: windowFrom, toDate: windowTo });
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch payments: ' + e.message });
  }

  const now = new Date().toISOString();
  const actor = req.user.name || req.user.email || 'system';

  // Also fetch bills (windowed) to build a bill id -> invoice date map, so the
  // core can apply the conversion-cutoff rule by invoice date, not payment date.
  const billInvoiceDateById = new Map();
  try {
    const bills = await billcomListBillsWindowed({ ...listArgs, fromDate: windowFrom, toDate: windowTo });
    for (const b of bills) {
      const bid = pick(b, 'id');
      const idt = pick(b, 'invoiceDate', 'invoice_date', 'dueDate') || pick(pick(b, 'invoice') || {}, 'invoiceDate', 'invoice_date');
      if (bid != null && idt) billInvoiceDateById.set(String(bid), String(idt));
    }
  } catch (e) {
    return res.status(502).json({ error: 'Failed to fetch bills for invoice-date check: ' + e.message });
  }

  // Delegate to the shared two-leg core (same logic Sync uses), so both paths
  // post identical entries and share dedup keys — re-running either is safe.
  const result = performPaymentReconcileCore({
    entityId, apAccount, clearingAccount, cashAccount,
    payments, asOf, cutoffDate, dryRun, actor, now, billInvoiceDateById,
  });
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
    `SELECT jl.id AS line_id, je.id AS entry_id, je.entry_num, je.date, je.memo, je.vendor,
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
  const invNumByEntry = new Map();
  try {
    const rows = db.prepare(
      "SELECT cl_entry_id, billcom_id, invoice_number FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill' AND status = 'success' AND cl_entry_id IS NOT NULL"
    ).all(entityId);
    for (const r of rows) { syncedEntryIds.add(r.cl_entry_id); billcomIdByEntry.set(r.cl_entry_id, String(r.billcom_id)); if (r.invoice_number) invNumByEntry.set(r.cl_entry_id, String(r.invoice_number)); }
  } catch (e) { /* sync log optional */ }

  // ── 3. Net each payment against the SPECIFIC bill it settled (not FIFO
  //    oldest-first), and treat the imported opening balance as a locked block.
  //    A Bill.com payment JE names its bill ("relieve bill <billcomId>"); we net
  //    it against that bill's own credit and never against the opening GL block.
  //    Imported (non-Bill.com) debits still net FIFO within the imported block.
  //    AP debits paired with cash credits (from bank transaction uploads) reduce
  //    the opening balance directly.
  //    Anything unmatched is carried out as a reconciling line so the report
  //    total still ties to the GL balance exactly.
  const entryIdByBillcomId = new Map(); // billcom_id -> synced-bill cl_entry_id
  for (const [eid, bcid] of billcomIdByEntry.entries()) entryIdByBillcomId.set(String(bcid), eid);
  
  // Identify entries with both AP debit and cash credit (opening balance relief via bank txns)
  const entryHasCashCredit = new Map(); // entry_id -> boolean
  for (const l of glLines) {
    // Cash accounts typically start with 101 or 102
    if ((String(l.account_code).startsWith('101') || String(l.account_code).startsWith('102')) && (l.credit || 0) > 0.005) {
      entryHasCashCredit.set(l.entry_id, true);
    }
  }
  
  const billOpenByEntry = new Map(); // synced-bill entry_id -> { ...line, remaining }
  let glCreditQueue = [];            // FIFO queue of imported/opening credits ONLY
  let glUnappliedDebit = 0;          // imported over-relief carried within the GL block
  let unmatchedPayment = 0;          // Bill.com payment debits not matched to a synced bill
  let openingBalanceRelief = 0;      // AP debits paired with cash credits (opening balance relief)
  const matchedPayments = []; // { targetEntry, debit } - applied AFTER all bill credits are known
  for (const l of glLines) {
    if ((l.credit || 0) > 0.005) {
      if (syncedEntryIds.has(l.entry_id)) {
        const prev = billOpenByEntry.get(l.entry_id);
        billOpenByEntry.set(l.entry_id, { ...l, remaining: (prev ? prev.remaining : 0) + l.credit });
      } else {
        let remaining = l.credit;
        if (glUnappliedDebit > 0.005) { const take = Math.min(remaining, glUnappliedDebit); remaining -= take; glUnappliedDebit -= take; }
        if (remaining > 0.005) glCreditQueue.push({ ...l, remaining });
      }
    }
    if ((l.debit || 0) > 0.005) {
      const mm = /relieve bill (\S+)/.exec(l.memo || '');
      const billId = mm ? mm[1] : null;
      const targetEntry = billId ? entryIdByBillcomId.get(String(billId)) : null;
      if (targetEntry != null) {
        matchedPayments.push({ targetEntry, debit: l.debit }); // net after the pass (a payment can post before its bill's date)
      } else if (billId) {
        unmatchedPayment += l.debit; // payment for a bill not in CL (e.g. pre-cutover) - never touch opening
      } else if (entryHasCashCredit.get(l.entry_id)) {
        // AP debit paired with cash credit (opening balance relief from bank transactions)
        openingBalanceRelief += l.debit;
      } else {
        let pay = l.debit; // imported/manual AP debit: FIFO within the imported block only
        while (pay > 0.005 && glCreditQueue.length) {
          const head = glCreditQueue[0];
          const take = Math.min(head.remaining, pay);
          head.remaining -= take; pay -= take;
          if (head.remaining <= 0.005) glCreditQueue.shift();
        }
        if (pay > 0.005) glUnappliedDebit += pay;
      }
    }
  }
  // Apply matched payments now that every synced bill credit is known - a payment
  // can post before its bill's GL date, so this must run after the pass above.
  for (const mp of matchedPayments) {
    const b = billOpenByEntry.get(mp.targetEntry);
    if (!b) { unmatchedPayment += mp.debit; continue; }
    const take = Math.min(b.remaining, mp.debit);
    b.remaining -= take;
    if (mp.debit - take > 0.005) unmatchedPayment += (mp.debit - take);
  }
  const openItems = []; // { line_id, entry_id, entry_num, date, memo, description, vendor, amount }
  for (const [eid, b] of billOpenByEntry.entries()) {
    if (b.remaining > 0.005) openItems.push({ line_id: b.line_id, entry_id: eid, entry_num: b.entry_num, date: b.date, memo: b.memo || '', description: b.description || '', vendor: b.vendor || '', amount: b.remaining });
  }
  for (const c of glCreditQueue) {
    if (c.remaining > 0.005) openItems.push({ line_id: c.line_id, entry_id: c.entry_id, entry_num: c.entry_num, date: c.date, memo: c.memo || '', description: c.description || '', vendor: c.vendor || '', amount: c.remaining });
  }

  // ── 4. Vendor + invoice number for Bill.com-synced items come from LOCAL data
  //    now (the JE's vendor field, populated at sync, plus the sync log's
  //    invoice_number), so the report builds instantly — no live Bill.com login,
  //    vendor list, or multi-year bill fetch. Due date defaults to the line date
  //    (aging is computed off the line date regardless).
  let billcomError = null;
  const invNumFromMemo = (memo) => { const s = String(memo || ''); const h = s.match(/#\s*([^\s].*?)\s*$/); if (h) return h[1].trim(); const m = s.match(/—\s*(.+?)\s*$/); return m ? m[1].trim() : null; };

  // ── 5. Build buckets for Bill.com invoices; sum GL column for the rest.
  const buckets = ['current', 'd1_30', 'd31_60', 'd61_90', 'd91_plus'];
  const emptyBuckets = () => ({ current: 0, d1_30: 0, d31_60: 0, d61_90: 0, d91_plus: 0, gl: 0, total: 0 });
  const bucketOf = (d) => d <= 0 ? 'current' : d <= 30 ? 'd1_30' : d <= 60 ? 'd31_60' : d <= 90 ? 'd61_90' : 'd91_plus';
  const dayDiff = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);

  const byVendor = new Map();
  const glRows = [];
  const grand = emptyBuckets();

  for (const it of openItems) {
    const num = invNumByEntry.get(it.entry_id) || invNumFromMemo(it.memo) || String(it.entry_num);
    const isBillcom = syncedEntryIds.has(it.entry_id);
    if (isBillcom) {
      const vname = it.vendor || 'Vendor';
      const dueDate = it.date;
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
  if (glUnappliedDebit > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Net prepaid / overpayment (payments exceed open bills)', description: '', amount: -glUnappliedDebit });
    grand.gl -= glUnappliedDebit; grand.total -= glUnappliedDebit;
  }
  if (unmatchedPayment > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Bill.com payment(s) not matched to a synced bill', description: '', amount: -unmatchedPayment });
    grand.gl -= unmatchedPayment; grand.total -= unmatchedPayment;
  }
  if (openingBalanceRelief > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Opening AP balance relief (bank transactions)', description: '', amount: -openingBalanceRelief });
    grand.gl -= openingBalanceRelief; grand.total -= openingBalanceRelief;
  }


  const glTotal = glRows.reduce((s, r) => s + r.amount, 0);
  // Once the GL-sourced A/P nets to zero (e.g. the legacy balance was cleared by a
  // journal entry), drop those GL entries from the report entirely — they no longer
  // represent anything outstanding, so the report shows only real open items.
  if (Math.abs(glTotal) < 0.005) { grand.total -= grand.gl; grand.gl = 0; glRows.length = 0; }
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
    gl_total: glRows.reduce((s, r) => s + r.amount, 0),
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
    const before = db.prepare('SELECT COUNT(*) c FROM requisition_invoice WHERE entity_id = ? AND req_number IS NULL AND draft_id IS NULL').get(eid).c;
    const info = db.prepare('DELETE FROM requisition_invoice WHERE entity_id = ? AND req_number IS NULL AND draft_id IS NULL').run(eid);
    res.json({ deleted: info.changes, matched: before });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// Editable / persistent Requisition Report draft
// ───────────────────────────────────────────────────────────────────────────
// One open draft per entity. Create (auto-seed or first-time upload) → edit its
// invoice list (add/update/delete) → Save/re-roll → Finalize (files to
// Workpapers, one copy only, becomes next month's auto-seed base).
// ═══════════════════════════════════════════════════════════════════════════

// Map a stored requisition_invoice row into a Current-Invoice-Log "newCurrent"
// row the roll-forward engine understands (same shape the legacy client sends).
function draftInvoiceToNewCurrent(inv) {
  const out = {
    code: inv.cost_code || undefined,
    name: inv.cost_code_name || undefined,
    vendor: inv.vendor || undefined,
    bill: inv.bill_number || undefined,
    date: inv.invoice_date || undefined,
  };
  const amt = inv.amount != null && inv.amount !== '' ? Number(String(inv.amount).replace(/[$,]/g, '')) : NaN;
  if (Number.isFinite(amt)) out.amount = amt;
  return out;
}

// Serialize a draft row for the client (never ships the blobs).
function draftForClient(db, draft) {
  if (!draft) return null;
  const invoices = db.prepare(
    'SELECT id, vendor, bill_number, amount, invoice_date, cost_code, cost_code_name, confidence, original_name, mime_type, ' +
    "(file_blob IS NOT NULL) AS has_file FROM requisition_invoice WHERE draft_id = ? ORDER BY id"
  ).all(draft.id);
  let recon = null;
  try { recon = draft.recon_summary ? JSON.parse(draft.recon_summary) : null; } catch (_) {}
  return {
    id: draft.id, entity_id: draft.entity_id, status: draft.status,
    req_number: draft.req_number, as_of_date: draft.as_of_date, phase: draft.phase || '',
    base_name: draft.base_name, output_name: draft.output_name, packet_name: draft.packet_name,
    recon_ok: draft.recon_ok == null ? null : !!draft.recon_ok, recon: recon,
    has_output: draft.output_blob != null,
    created_at: draft.created_at, updated_at: draft.updated_at, finalized_at: draft.finalized_at,
    created_by: draft.created_by,
    invoices,
  };
}

// GET the entity's open draft (or {draft:null}). Powers "reopen and edit".
app.get('/api/requisition/:entity_id/draft', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    // With a phase (?phase=) return that stream's open draft; without one, return
    // the full page state: all open drafts (drafts:[...]) so the UI can list
    // Phase 2 / 2a, plus the reopenable finalized streams (finalized:[...]) — the
    // LATEST finalized report per phase that has no open draft — so the page can
    // offer "Reopen for edits" on exactly those.
    const ent = db.prepare('SELECT entity_type FROM entities WHERE id=?').get(eid) || {};
    const isRail = ent.entity_type === 'rail_assets';
    if (req.query.phase != null && req.query.phase !== '') {
      const draft = reqDraft.getOpenDraft(db, eid, req.query.phase);
      return res.json({ draft: draftForClient(db, draft), is_rail: isRail });
    }
    const all = reqDraft.getOpenDrafts(db, eid).map(d => draftForClient(db, d));
    // Latest finalized per phase, excluding any phase that has an open draft.
    const openPhases = new Set(all.map(d => d.phase || ''));
    const finRows = db.prepare(
      "SELECT * FROM requisition_draft d WHERE entity_id=? AND status='finalized' " +
      "AND finalized_at = (SELECT MAX(finalized_at) FROM requisition_draft x WHERE x.entity_id=d.entity_id AND x.status='finalized' AND IFNULL(x.phase,'')=IFNULL(d.phase,'')) " +
      "ORDER BY IFNULL(phase,''), id DESC"
    ).all(eid);
    const seenPhase = new Set();
    const finalized = [];
    for (const r of finRows) {
      const ph = r.phase || '';
      if (seenPhase.has(ph) || openPhases.has(ph)) continue; // one per phase; skip if a draft is open
      seenPhase.add(ph);
      finalized.push(draftForClient(db, r));
    }
    res.json({ draft: all[0] || null, drafts: all, finalized, is_rail: isRail });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Report what a NEW draft would auto-seed from, so the client can either proceed
// (auto-seed available) or show the first-time upload box (source === null).
app.get('/api/requisition/:entity_id/draft/seed-source', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const ent = db.prepare('SELECT entity_type FROM entities WHERE id=?').get(eid) || {};
    const seed = reqDraft.resolveAutoSeed(db, WORKPAPERS_DIR, eid, req.query.phase);
    res.json({ source: seed ? seed.source : null, name: seed ? seed.name : null, is_rail: ent.entity_type === 'rail_assets' });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// CREATE the open draft. Auto-seeds base from the last finalized Req; if none
// exists, a base workbook upload is required (first time). Upload guard (Option
// B): if a hand-uploaded base matches the filed copy, we STOP and ask unless the
// client confirms which to use (baseChoice=uploaded|filed).
app.post('/api/requisition/:entity_id/draft', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res, next) => {
  reqRollUpload.single('workbook')(req, res, (err) => {
    if (err) return res.status(err.code === 'LIMIT_FILE_SIZE' ? 413 : 400).json({ error: 'Upload failed: ' + err.message });
    next();
  });
}, async (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const ent = db.prepare('SELECT entity_type FROM entities WHERE id=?').get(eid) || {};
    const isRail = ent.entity_type === 'rail_assets';
    // Phase (rail-assets stream key). Ignored for non-rail entities (single stream).
    const phase = isRail ? reqDraft.normPhase(req.body.phase) : '';
    if (reqDraft.getOpenDraft(db, eid, phase)) {
      return res.status(409).json({ error: phase
        ? ('An open draft for Phase ' + phase + ' already exists. Open it to continue, or finalize it first.')
        : 'An open draft already exists for this entity. Open it to continue editing, or finalize it first.' });
    }
    // Rail assets may run at most TWO requisition streams at once.
    if (isRail) {
      const openCount = reqDraft.getOpenDrafts(db, eid).length;
      if (openCount >= 2) return res.status(409).json({ error: 'This rail asset already has two open requisition drafts. Finalize one before starting another.' });
    }
    const reqNumber = req.body.reqNumber != null && req.body.reqNumber !== '' ? parseInt(req.body.reqNumber) : null;
    const asOfDate = req.body.asOfDate || null;
    const baseChoice = req.body.baseChoice || null; // 'uploaded' | 'filed' | null

    let baseBuf = null, baseName = null;
    const uploaded = req.file ? req.file.buffer : null;
    const seed = reqDraft.resolveAutoSeed(db, WORKPAPERS_DIR, eid, phase);

    if (uploaded) {
      // Option B: if the upload matches the filed prior-month copy, ask before proceeding.
      if (!baseChoice) {
        const conflict = reqDraft.checkUploadAgainstFiled(db, WORKPAPERS_DIR, eid, uploaded, reqNumber, asOfDate, phase);
        if (conflict) {
          return res.status(409).json({
            error: 'upload_matches_filed',
            conflict,
            message: conflict.match === 'identical'
              ? `The uploaded file is an exact copy of the requisition report already filed in Workpapers (${conflict.filedName}). Use the uploaded file, or the filed copy?`
              : `The uploaded file matches the period of the filed report (${conflict.filedName}) but its contents differ — it looks like an edited copy. Use the uploaded (edited) version, or the filed copy?`,
          });
        }
      }
      if (baseChoice === 'filed' && seed) { baseBuf = seed.buffer; baseName = seed.name; }
      else { baseBuf = uploaded; baseName = req.file.originalname || 'uploaded_base.xlsx'; }
    } else if (seed) {
      baseBuf = seed.buffer; baseName = seed.name;
    } else {
      return res.status(400).json({ error: 'No prior finalized requisition on file — upload the prior month\'s finalized workbook (field name: workbook) to start.' });
    }

    // Validate the base has the required tabs before we store it.
    try { await reqDraft.loadReqWorkbook(ExcelJS, baseBuf); }
    catch (e) { return res.status(e.status || 400).json({ error: e.message }); }

    const now = new Date().toISOString();
    const who = (req.user && (req.user.name || req.user.email)) || 'system';
    const info = db.prepare(
      'INSERT INTO requisition_draft (entity_id, status, phase, req_number, as_of_date, base_blob, base_name, base_sha256, created_at, updated_at, created_by) ' +
      "VALUES (?, 'open', ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    ).run(eid, phase || null, reqNumber, asOfDate, baseBuf, baseName, reqDraft.sha256(baseBuf), now, now, who);

    const draft = db.prepare('SELECT * FROM requisition_draft WHERE id = ?').get(info.lastInsertRowid);
    res.json({ draft: draftForClient(db, draft), seededFrom: uploaded ? (baseChoice === 'filed' ? 'filed' : 'upload') : (seed ? seed.source : null) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ADD an invoice to the open draft (OCR/coding done client-side, same payload as
// the legacy invoices[] rows). Marks the draft dirty (recon stale until re-roll).
app.post('/api/requisition/:entity_id/draft/invoice', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const draft = reqDraft.getOpenDraft(db, eid, req.body.phase);
    if (!draft) return res.status(404).json({ error: 'No open draft. Start a new requisition first.' });
    const inv = req.body || {};
    const amt = inv.amount != null && inv.amount !== '' ? Number(String(inv.amount).replace(/[$,]/g, '')) : null;
    let blob = null;
    try { if (inv.file_b64) blob = Buffer.from(inv.file_b64, 'base64'); } catch (_) {}
    const info = db.prepare(
      'INSERT INTO requisition_invoice (entity_id, req_number, draft_id, vendor, bill_number, amount, invoice_date, cost_code, cost_code_name, confidence, original_name, mime_type, file_blob, created_at) ' +
      'VALUES (?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
    ).run(
      eid, draft.id,
      inv.vendor || null, inv.bill_number || inv.bill || null,
      Number.isFinite(amt) ? amt : null, inv.invoice_date || null,
      inv.cost_code || null, inv.cost_code_name || null, inv.confidence || null,
      inv.original_name || inv.filename || null, inv.mime_type || null, blob,
      new Date().toISOString()
    );
    db.prepare("UPDATE requisition_draft SET updated_at=?, recon_ok=NULL WHERE id=?").run(new Date().toISOString(), draft.id);
    res.json({ id: info.lastInsertRowid, draft: draftForClient(db, reqDraft.getOpenDraft(db, eid, req.body.phase)) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// UPDATE a draft invoice's coding/amount in place. Marks the draft dirty.
app.put('/api/requisition/:entity_id/draft/invoice/:invoice_id', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const draft = reqDraft.getOpenDraft(db, eid, req.body.phase);
    if (!draft) return res.status(404).json({ error: 'No open draft.' });
    const row = db.prepare('SELECT * FROM requisition_invoice WHERE id=? AND draft_id=?').get(parseInt(req.params.invoice_id), draft.id);
    if (!row) return res.status(404).json({ error: 'Invoice not found on this draft.' });
    const inv = req.body || {};
    const amt = inv.amount != null && inv.amount !== '' ? Number(String(inv.amount).replace(/[$,]/g, '')) : row.amount;
    db.prepare(
      'UPDATE requisition_invoice SET vendor=?, bill_number=?, amount=?, invoice_date=?, cost_code=?, cost_code_name=? WHERE id=?'
    ).run(
      inv.vendor != null ? inv.vendor : row.vendor,
      inv.bill_number != null ? inv.bill_number : (inv.bill != null ? inv.bill : row.bill_number),
      Number.isFinite(amt) ? amt : row.amount,
      inv.invoice_date != null ? inv.invoice_date : row.invoice_date,
      inv.cost_code != null ? inv.cost_code : row.cost_code,
      inv.cost_code_name != null ? inv.cost_code_name : row.cost_code_name,
      row.id
    );
    db.prepare("UPDATE requisition_draft SET updated_at=?, recon_ok=NULL WHERE id=?").run(new Date().toISOString(), draft.id);
    res.json({ draft: draftForClient(db, reqDraft.getOpenDraft(db, eid, req.body.phase)) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// DELETE a draft invoice line. Marks the draft dirty.
app.delete('/api/requisition/:entity_id/draft/invoice/:invoice_id', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const dphase = req.query.phase != null ? req.query.phase : (req.body && req.body.phase);
  try {
    const draft = reqDraft.getOpenDraft(db, eid, dphase);
    if (!draft) return res.status(404).json({ error: 'No open draft.' });
    const info = db.prepare('DELETE FROM requisition_invoice WHERE id=? AND draft_id=?').run(parseInt(req.params.invoice_id), draft.id);
    db.prepare("UPDATE requisition_draft SET updated_at=?, recon_ok=NULL WHERE id=?").run(new Date().toISOString(), draft.id);
    res.json({ deleted: info.changes, draft: draftForClient(db, reqDraft.getOpenDraft(db, eid, dphase)) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// SAVE / re-roll: rebuild the draft's workbook + packet from the CURRENT invoice
// set against the stored base. Overwrites output_blob/packet_blob/recon_*. Does
// NOT touch Workpapers (that happens only at finalize).
app.post('/api/requisition/:entity_id/draft/roll', ...reqGuards(), requireRole('Admin', 'Accountant'), async (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const draft = reqDraft.getOpenDraft(db, eid, req.body.phase);
    if (!draft) return res.status(404).json({ error: 'No open draft.' });
    if (!draft.base_blob) return res.status(400).json({ error: 'Draft has no base workbook.' });

    // Allow header edits (req#, as-of) to be sent with the roll.
    const reqNumber = req.body.reqNumber != null && req.body.reqNumber !== '' ? parseInt(req.body.reqNumber) : draft.req_number;
    const asOfDate = req.body.asOfDate || draft.as_of_date;

    const invoices = db.prepare('SELECT * FROM requisition_invoice WHERE draft_id=? ORDER BY id').all(draft.id);
    const newCurrent = invoices.map(draftInvoiceToNewCurrent).filter(r => Number.isFinite(r.amount));

    const { outBuf, rfResult, verification } = await reqDraft.rollForwardFromBase(
      ExcelJS, Buffer.from(draft.base_blob), eid, newCurrent, { reqNumber, asOfDate }
    );

    const outName = reqDraft.phasedFilename(buildRollforwardFilename(draft.base_name || 'Requisition_Report.xlsx', reqNumber, asOfDate), draft.phase);
    const reconSummary = verification && verification.finalResult ? JSON.stringify({
      ok: !!verification.ok,
      summary: verification.finalResult.summary,
      checks: verification.finalResult.checks,
      unresolved: verification.unresolved,
    }) : null;

    db.prepare(
      'UPDATE requisition_draft SET req_number=?, as_of_date=?, output_blob=?, output_name=?, recon_ok=?, recon_summary=?, updated_at=? WHERE id=?'
    ).run(
      reqNumber, asOfDate, outBuf, outName,
      verification && verification.ok ? 1 : 0, reconSummary,
      new Date().toISOString(), draft.id
    );

    // Keep an in-progress copy of the report visible in Workpapers while the
    // draft is being worked on. It lives in the same month folder as the final
    // report will, but with a [DRAFT] prefix so it is unmistakably provisional,
    // and is overwritten on each Prepare. Finalize deletes it and files the
    // clean report in its place.
    try {
      const draftWho = (req.user && (req.user.name || req.user.email)) || 'system';
      const draftFolder = require('./requisition_workpaper_save').requisitionFolderPath(asOfDate) + '/Drafts';
      const draftFileName = '[DRAFT] ' + String(outName || 'Requisition_Report.xlsx');
      try { ensureWpFolders(db, eid, draftFolder, draftWho); } catch (_e) {}
      saveWpBuffer(
        db, WORKPAPERS_DIR, eid, draftFolder, draftFileName,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        Buffer.from(outBuf), draftWho, { overwrite: true }
      );
    } catch (_e) { /* non-fatal: the draft blob in the DB remains the source of truth */ }

    res.json({
      ok: !!(verification && verification.ok),
      recon: reconSummary ? JSON.parse(reconSummary) : null,
      devFee: rfResult && rfResult.devFee && !rfResult.devFee.error ? rfResult.devFee : null,
      output_name: outName,
      draft: draftForClient(db, reqDraft.getOpenDraft(db, eid, draft.phase)),
    });
  } catch (e) { res.status(e.status || 500).json({ error: e.message }); }
});

// DOWNLOAD the draft's current workbook (must have been rolled at least once).
app.get('/api/requisition/:entity_id/draft/download', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  try {
    const draft = reqDraft.getOpenDraft(db, eid, req.query.phase);
    if (!draft || !draft.output_blob) return res.status(404).json({ error: 'No rolled workbook yet — Save/re-roll the draft first.' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + String(draft.output_name || 'Requisition_Report.xlsx').replace(/[\r\n"]/g, ' ') + '"');
    res.send(Buffer.from(draft.output_blob));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// FINALIZE: lock the draft, stamp its invoices with the req number, file the
// workbook + packet to Workpapers (one copy only), and leave it as next month's
// auto-seed base. Re-rolls first so the filed output reflects the latest edits.
app.post('/api/requisition/:entity_id/draft/finalize', ...reqGuards(), requireRole('Admin', 'Accountant'), async (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const force = req.body.force === true || req.body.force === 'true' || req.body.force === '1';
  try {
    const draft = reqDraft.getOpenDraft(db, eid, req.body.phase);
    if (!draft) return res.status(404).json({ error: 'No open draft.' });
    if (!draft.base_blob) return res.status(400).json({ error: 'Draft has no base workbook.' });

    const reqNumber = draft.req_number;
    const asOfDate = draft.as_of_date;
    const invoices = db.prepare('SELECT * FROM requisition_invoice WHERE draft_id=? ORDER BY id').all(draft.id);
    const newCurrent = invoices.map(draftInvoiceToNewCurrent).filter(r => Number.isFinite(r.amount));

    // Always re-roll at finalize so the filed copy is current.
    const { outBuf, rfResult, verification, workbook } = await reqDraft.rollForwardFromBase(
      ExcelJS, Buffer.from(draft.base_blob), eid, newCurrent, { reqNumber, asOfDate }
    );

    if (!(verification && verification.ok) && !force) {
      const reconSummary = verification && verification.finalResult ? {
        ok: false, summary: verification.finalResult.summary,
        checks: verification.finalResult.checks, unresolved: verification.unresolved,
      } : null;
      // Persist the fresh (failed) output so the user can download + inspect it.
      db.prepare('UPDATE requisition_draft SET output_blob=?, recon_ok=0, recon_summary=?, updated_at=? WHERE id=?')
        .run(outBuf, reconSummary ? JSON.stringify(reconSummary) : null, new Date().toISOString(), draft.id);
      return res.status(422).json({ error: 'Roll-forward failed reconciliation', ok: false, recon: reconSummary });
    }

    const fname = reqDraft.phasedFilename(buildRollforwardFilename(draft.base_name || 'Requisition_Report.xlsx', reqNumber, asOfDate), draft.phase);
    const who = (req.user && (req.user.name || req.user.email)) || 'system';
    // Phase-scoped purge predicate: only sweep files belonging to THIS phase, so
    // Phase 2a's finalize never deletes Phase 2's filed copy in the same folder.
    const phaseMatch = (nm) => reqDraft.phaseMatchesName(nm, draft.phase);

    // Build invoice rows (with bytes) for the packet, ordered like the Current Log.
    let invoiceRows = invoices.map(inv => ({
      original_name: inv.original_name, mime_type: inv.mime_type, file_blob: inv.file_blob ? Buffer.from(inv.file_blob) : null,
      vendor: inv.vendor, bill_number: inv.bill_number, amount: inv.amount,
      cost_code: inv.cost_code, cost_code_name: inv.cost_code_name,
    }));
    try { invoiceRows = orderInvoicesByCurrentLog(workbook, invoiceRows); } catch (_) {}

    const entRow = db.prepare('SELECT name, display_id FROM entities WHERE id = ?').get(eid) || {};
    let packetPrefix = (entRow.display_id && entRow.display_id.trim()) || entRow.name || '';
    { const _pl = reqDraft.phaseLabel(draft.phase); if (_pl) packetPrefix = (packetPrefix + ' ' + _pl).trim(); }

    // One-copy enforcement: clear prior report/packet copies in OTHER month
    // folders (in case a re-finalize moved the period), then let
    // saveRequisitionOutputs overwrite same-named files in the target folder.
    const targetFolder = require('./requisition_workpaper_save').requisitionFolderPath(asOfDate);
    try { purgePriorRequisitionCopies(db, WORKPAPERS_DIR, eid, { keepFolderPath: targetFolder, otherFoldersOnly: true, phaseMatch }); } catch (_) {}

    // Remove the in-progress [DRAFT] copy for this phase from the Drafts subfolder
    // (and any other month folder a re-finalize may have moved it from). The
    // finalized report filed just below replaces it.
    try {
      const draftRows = db.prepare(
        "SELECT ef.id, ef.stored_filename FROM entity_files ef WHERE ef.entity_id=? AND ef.original_name LIKE '[DRAFT] %'"
      ).all(eid);
      for (const dr of draftRows) {
        const bare = dr && dr.stored_filename;
        // Only sweep [DRAFT] files belonging to THIS phase, mirroring phaseMatch
        // so Phase 2a's finalize never removes Phase 2's in-progress copy.
        const nameRow = db.prepare('SELECT original_name FROM entity_files WHERE id=?').get(dr.id);
        if (nameRow && !phaseMatch(nameRow.original_name)) continue;
        try { fs.unlinkSync(path.join(WORKPAPERS_DIR, String(eid), bare)); } catch (_) {}
        db.prepare('DELETE FROM entity_files WHERE id=?').run(dr.id);
      }
    } catch (_) {}

    const saved = await saveRequisitionOutputs({
      db, workpapersDir: WORKPAPERS_DIR, eid,
      reqNumber, asOfDate, workbookBuffer: Buffer.from(outBuf), invoices: invoiceRows,
      devFee: rfResult && rfResult.devFee && !rfResult.devFee.error ? rfResult.devFee : null,
      who, packetPrefix, workbookFilename: fname, saveWorkbook: true,
    });

    // Also sweep any leftover differently-named report/packet in the target
    // folder (e.g. a prior finalize under a different req number), keeping the
    // two we just wrote.
    const keepNames = [saved.workbook && saved.workbook.original_name, saved.packet && saved.packet.original_name].filter(Boolean);
    try { purgePriorRequisitionCopies(db, WORKPAPERS_DIR, eid, { keepFolderPath: targetFolder, keepNames, phaseMatch }); } catch (_) {}

    // Stamp the draft's invoices with the final req number (keep draft_id for provenance).
    const now = new Date().toISOString();
    const reconSummary = verification && verification.finalResult ? JSON.stringify({
      ok: !!verification.ok, summary: verification.finalResult.summary,
      checks: verification.finalResult.checks, unresolved: verification.unresolved,
    }) : null;
    const tx = db.transaction(() => {
      if (reqNumber != null) db.prepare('UPDATE requisition_invoice SET req_number=? WHERE draft_id=?').run(reqNumber, draft.id);
      db.prepare(
        "UPDATE requisition_draft SET status='finalized', output_blob=?, output_name=?, packet_name=?, recon_ok=?, recon_summary=?, finalized_at=?, updated_at=? WHERE id=?"
      ).run(
        outBuf, fname, saved.packet ? saved.packet.original_name : null,
        verification && verification.ok ? 1 : 0, reconSummary, now, now, draft.id
      );
    });
    tx();

    res.json({
      ok: true, folder: saved.folder,
      workbook: saved.workbook ? saved.workbook.original_name : null,
      packet: saved.packet ? saved.packet.original_name : null,
      forced: !(verification && verification.ok),
    });
  } catch (e) { res.status(e.status || 500).json({ error: e.message }); }
});

// Reopen the LATEST finalized requisition of a stream back into an editable
// draft. Only the most-recent finalized report per (entity, phase) may be
// reopened — never an older one that a newer req was already seeded from, which
// would leave that newer req's base stale. The draft's invoices are still linked
// by draft_id (finalize keeps them), so flipping status back to 'open' restores
// the full editable set; re-finalizing replaces the filed copy in place.
app.post('/api/requisition/:entity_id/draft/reopen', ...reqGuards(), requireRole('Admin', 'Accountant'), (req, res) => {
  const eid = parseInt(req.params.entity_id);
  const phase = reqDraft.normPhase(req.body.phase);
  try {
    // Guard: an open draft for this stream already exists — nothing to reopen.
    if (reqDraft.getOpenDraft(db, eid, phase)) {
      return res.status(409).json({ error: 'This stream already has an open draft.' });
    }
    // The latest finalized draft for this exact stream.
    const latest = db.prepare(
      "SELECT * FROM requisition_draft WHERE entity_id=? AND status='finalized' AND IFNULL(phase,'')=? ORDER BY finalized_at DESC, id DESC LIMIT 1"
    ).get(eid, phase);
    if (!latest) return res.status(404).json({ error: 'No finalized requisition to reopen for this stream.' });

    // Flip it back to open. Keep output/packet blobs and recon state as-is; a
    // subsequent Save/re-roll or Re-finalize regenerates them from the invoices.
    db.prepare("UPDATE requisition_draft SET status='open', finalized_at=NULL, updated_at=? WHERE id=?")
      .run(new Date().toISOString(), latest.id);
    const draft = db.prepare('SELECT * FROM requisition_draft WHERE id=?').get(latest.id);
    res.json({ draft: draftForClient(db, draft) });
  } catch (e) { res.status(e.status || 500).json({ error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════════════════


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

// Translate ExcelJS's cryptic shared-formula write error into a clear, actionable
// message. ExcelJS throws e.g. "Shared Formula master must exist above and or left
// of clone for cell F14" when a follower cell points to a master formula that is
// missing or positioned after it — which a roll-forward can trip when a total or
// subtotal row shifts. We locate the exact tab + master cell so the user knows
// precisely what to fix. Returns null when the error is something else.
function explainSharedFormulaWriteError(workbook, rawMsg) {
  const m = /Shared Formula master must exist[^]*?for cell ([A-Z]+\d+)/i.exec(rawMsg || '');
  if (!m) return null;
  const ref = m[1];
  const parse = (a) => { const q = /^([A-Z]+)(\d+)$/.exec(a || ''); if (!q) return null; const col = q[1].split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0); return { col, row: parseInt(q[2], 10) }; };
  const rp = parse(ref) || { col: 0, row: 0 };
  let tab = null, master = null;
  try {
    for (const ws of (workbook ? workbook.worksheets : [])) {
      let cell; try { cell = ws.getCell(ref); } catch (e) { continue; }
      const mdl = cell && cell.model;
      if (!mdl || !mdl.sharedFormula) continue;
      const mp = parse(mdl.sharedFormula);
      const positionedOK = mp && (mp.row < rp.row || (mp.row === rp.row && mp.col < rp.col));
      let masterOk = false;
      try { const mm = ws.getCell(mdl.sharedFormula).model; masterOk = !!(mm && mm.shareType === 'shared'); } catch (e) {}
      if (!positionedOK || !masterOk) { tab = ws.name; master = mdl.sharedFormula; break; }
    }
  } catch (e) {}
  const at = tab ? ` on the "${tab}" tab` : '';
  const expected = master ? ` (its master formula should sit at ${master})` : '';
  const goTab = tab ? ` go to the "${tab}" tab,` : '';
  const message =
    `Roll-forward couldn't save the workbook because of a broken shared formula at cell ${ref}${at}. ` +
    `Excel stores a run of identical formulas as a group: one "master" cell holds the real formula and the cells below it just point to that master${expected}. ` +
    `While rolling the report forward, that master ended up missing or moved below/right of ${ref} — usually because a total or subtotal row was inserted, deleted, or reordered — which Excel's file format doesn't allow, so the save was rejected. ` +
    `To fix it: open the source requisition workbook,${goTab} click cell ${ref}, re-enter its formula (or copy the formula down from the row just above so the group has a valid master again), save, and re-run the roll-forward.`;
  return { message, cell: ref, sheet: tab, master };
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
//
// Reorder packet invoices to match the rolled-forward Current Invoice Log. The
// client sends invoices in on-screen order, but the roll-forward GROUPS the log
// by cost code — so the packet must follow the final log order, not the upload
// order. Reads the output log's leaf rows and sorts the invoice rows to match;
// unmatched rows keep their relative order at the end. Universal, best-effort.
function orderInvoicesByCurrentLog(workbook, invoiceRows) {
  if (!Array.isArray(invoiceRows) || invoiceRows.length < 2) return invoiceRows;
  try {
    const { COL, cellStr, cellNum, cellFormula, applyInvoiceCols } = require('./requisition_reconcile.js');
    const ws = findReqSheet(workbook, 'Current Invoice Log');
    if (!ws) return invoiceRows;
    applyInvoiceCols(ws); // ensure COL matches this workbook's layout
    const order = [];
    const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
    for (let r = 1; r <= last; r++) {
      const amtCell = ws.getCell(r, COL.amount);
      if (cellFormula(amtCell)) continue;                 // skip SUBTOTAL / grand-total rows
      const vendor = cellStr(ws.getCell(r, COL.vendor)).trim();
      const bill = cellStr(ws.getCell(r, COL.bill)).trim();
      const code = cellStr(ws.getCell(r, COL.code)).trim();
      const amt = cellNum(amtCell);
      if (!vendor && !bill && amt == null) continue;       // spacer/blank
      order.push({ code, bill, vendor, amt });
    }
    if (!order.length) return invoiceRows;
    const norm = s => String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
    const used = new Array(order.length).fill(false);
    const posOf = (inv) => {
      const ib = norm(inv.bill_number), ic = norm(inv.cost_code), iv = norm(inv.vendor);
      const ia = inv.amount != null ? Math.round(inv.amount * 100) : null;
      const pick = (pred) => { for (let i = 0; i < order.length; i++) if (!used[i] && pred(order[i])) { used[i] = true; return i; } return -1; };
      let idx = ib && ic ? pick(o => norm(o.code) === ic && norm(o.bill) === ib) : -1;
      if (idx < 0 && ib) idx = pick(o => norm(o.bill) === ib && (ia == null || o.amt == null || Math.round(o.amt * 100) === ia));
      if (idx < 0 && ib) idx = pick(o => norm(o.bill) === ib);
      if (idx < 0 && iv && ia != null) idx = pick(o => norm(o.vendor) === iv && o.amt != null && Math.round(o.amt * 100) === ia);
      if (idx < 0 && iv) idx = pick(o => norm(o.vendor) === iv);
      return idx;
    };
    return invoiceRows
      .map((inv, i) => ({ inv, p: posOf(inv), i }))
      .sort((a, b) => (a.p < 0 ? Infinity : a.p) - (b.p < 0 ? Infinity : b.p) || a.i - b.i)
      .map(x => x.inv);
  } catch (e) { return invoiceRows; }
}

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
  // Per-entity Dev Fee tab style: only these entities collapse Hard/Soft costs
  // into a single "Project costs" line; all others keep their existing Dev Fee
  // layout. Configurable via REQ_DEVFEE_COLLAPSE_ENTITIES (comma-separated entity
  // ids); defaults to County Line Industrial Park (CLIP = entity 42).
  const _dfCollapseIds = (process.env.REQ_DEVFEE_COLLAPSE_ENTITIES || '42,38,39').split(',').map(x => x.trim()).filter(Boolean);
  const _rfEid = String(parseInt(req.params.entity_id));
  meta.collapseDevFeeCosts = _dfCollapseIds.includes(_rfEid);
  // Same CLIP-style entities also fix the report-number header (update the
  // existing "Requisition Report #N" line in place instead of adding a duplicate).
  meta.fixReportNumberHeader = meta.collapseDevFeeCosts;
  // Per-entity development-fee payee (the dev-fee line vendor). Override the map
  // via REQ_DEVFEE_PAYEES as JSON {"<entityId>":"<payee>"}. Entities not listed
  // fall back to the vendor cloned from the prior dev-fee line.
  const _payeeMap = (() => { try { return Object.assign({ '42': 'County Line Rail Interest', '38': 'County Line Rail Interest', '39': 'County Line Rail Interest' }, JSON.parse(process.env.REQ_DEVFEE_PAYEES || '{}')); } catch (e) { return { '42': 'County Line Rail Interest', '38': 'County Line Rail Interest', '39': 'County Line Rail Interest' }; } })();
  if (meta.collapseDevFeeCosts && _payeeMap[_rfEid]) meta.devFeePayee = _payeeMap[_rfEid];

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
    prior: findReqSheet(priorBook, 'Prior Invoice Log'),
    current: findReqSheet(priorBook, 'Current Invoice Log'),
  };
  if (!prevSheets.prior || !prevSheets.current) {
    const present = priorBook.worksheets.map(s => s.name).join(', ');
    const missing = [!prevSheets.prior && 'Prior Invoice Log', !prevSheets.current && 'Current Invoice Log'].filter(Boolean).join(', ');
    return res.status(400).json({ error: 'Workbook is missing required tab(s): ' + missing + '. Tabs found: ' + (present || '(none)') + '. A requisition roll-forward needs "Prior Invoice Log", "Current Invoice Log", and "Budget to Actual" tabs.' });
  }

  try {
    // Mutate `workbook` into Req#N+1. The engine also auto-computes this period's
    // Development Fee by LEARNING the project's method from the prior report's Dev
    // Fee tab (formula parse, else Claude), back-validated against the prior fee;
    // rfResult.devFee carries the amount + row (or needsReview when it couldn't be
    // confirmed). Inject a Claude caller for the fallback when an API key is set.
    const devFeeCaller = process.env.ANTHROPIC_API_KEY ? makeDevFeeClaudeCaller() : null;
    const rfResult = await rollForward(workbook, newCurrent, { ...meta, callClaude: devFeeCaller });

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

    let outBuf = await workbook.xlsx.writeBuffer();
    // Finalize the workbook: force full recalc on open (so the Dev Fee tab,
    // B2A SUMIF columns, subtotals and grand total recompute from the rolled-
    // forward data instead of showing ExcelJS's stale cache) and re-inject any
    // external links ExcelJS dropped. Best-effort: unchanged buffer on failure.
    outBuf = await finalizeRequisitionWorkbook(req.file.buffer, Buffer.from(outBuf));

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
    // Always build the invoice packet (so it can be downloaded); only auto-save
    // the WORKBOOK into the Workpapers tree when the flag is on (saveWorkbook).
    try {
      const entRow = db.prepare('SELECT name, display_id FROM entities WHERE id = ?').get(eidInt) || {};
      const packetPrefix = (entRow.display_id && entRow.display_id.trim()) || entRow.name || '';
      // Reorder the packet to follow the Current Invoice Log (grouped by cost code).
      invoiceRows = orderInvoicesByCurrentLog(workbook, invoiceRows);
      const saved = await saveRequisitionOutputs({
        db, workpapersDir: WORKPAPERS_DIR, eid: eidInt,
        reqNumber: meta.reqNumber, asOfDate: meta.asOfDate,
        workbookBuffer: Buffer.from(outBuf), invoices: invoiceRows,
        devFee: rfResult && rfResult.devFee && !rfResult.devFee.error ? rfResult.devFee : null,
        who: (req.user && (req.user.name || req.user.email)) || 'system',
        packetPrefix, workbookFilename: fname,
        saveWorkbook: REQ_AUTOSAVE_WORKPAPERS,
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
    // Surface how the development fee was determined (or that it needs manual
    // entry) so the client can show it on the success card. Header-safe JSON.
    try {
      if (rfResult && rfResult.devFee) {
        const d = rfResult.devFee;
        const devFeeHeader = {
          amount: d.amount != null ? d.amount : null,
          base: d.base != null ? d.base : null,
          rate_text: d.rateText || null,
          source: d.source || null,            // 'formula:E15' | 'claude' | 'none'
          needs_review: !!d.needsReview,
          note: d.note || null,
          prior: d.prior || null,              // { base, fee } observed in prior period
          validated: d.validation ? !!d.validation.ok : null,
        };
        res.setHeader('X-Dev-Fee', JSON.stringify(devFeeHeader).replace(/[\r\n]/g, ' '));
      }
    } catch {}
    res.send(Buffer.from(outBuf));
  } catch (e) {
    // A userFacing error (e.g. a missing required tab) is a client-input problem,
    // not a server fault — return 400 with its message so the UI shows the actual
    // cause instead of a generic 500.
    if (e && e.userFacing) return res.status(400).json({ error: e.message });
    const sf = explainSharedFormulaWriteError(workbook, e && e.message);
    if (sf) return res.status(400).json({ error: sf.message, code: 'SHARED_FORMULA', cell: sf.cell, sheet: sf.sheet });
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

// ── Mgmt-fee Q-roll-forward engine ──────────────────────────────────────────
// Clones the current "Mgmt Fee Calc QX YY" + "QX YY Recalc" tabs into the NEXT
// quarter, shifts every quarter reference one forward, repoints the Invoice and
// ITD aggregators, and preserves the prior tabs intact. Operates at the zip/XML
// level so drawings, comments, and styles stay byte-intact (no ExcelJS rewrite,
// which would drop those parts and trigger Excel's repair prompt).
function mgmtShiftQuarterText(s) {
  return s.replace(/Q([1-4])('?\s?)(\d{2}|20\d{2})/g, (m, q, sep, yr) => {
    let qn = +q, y = yr.length === 2 ? 2000 + +yr : +yr;
    let nq = qn === 4 ? 1 : qn + 1, ny = qn === 4 ? y + 1 : y;
    let nyr = yr.length === 2 ? String(ny).slice(2) : String(ny);
    return 'Q' + nq + sep + nyr;
  });
}
async function mgmtRollForward(inputBuf, commitmentChanges = []) {
  const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  // Per-investor commitment change for the new quarter, keyed by trimmed name.
  // Entered by the user (e.g. a full redemption is a negative change equal to the
  // prior ending, zeroing the investor's commitment so no fee is calculated).
  const changeByName = {};
  for (const c of (commitmentChanges || [])) {
    if (c && c.name != null) changeByName[String(c.name).trim()] = Number(c.change) || 0;
  }
  const excelSerial = (dt) => Math.round((Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);

  // Identify the current calc tab + quarter via ExcelJS (for header/cols/endings).
  const ewb = new ExcelJS.Workbook();
  await ewb.xlsx.load(inputBuf);
  const calcSheets = mgmtFindCalcSheets(ewb);
  let cur = null;
  for (const ws of calcSheets) { const h = mgmtHeaderRow(ws); if (h) { cur = { ws, ...h }; break; } }
  if (!cur) throw new Error('no calc sheet found');
  const curName = cur.ws.name;
  const qm = curName.match(/Mgmt Fee Calc Q([1-4]) ?(\d{2})/i);
  if (!qm) throw new Error('calc tab name not in "Mgmt Fee Calc QX YY" form: ' + curName);
  const curQ = +qm[1], curYY = +qm[2];
  const nextQ = curQ === 4 ? 1 : curQ + 1, nextYY = curQ === 4 ? curYY + 1 : curYY;
  const prevQ = curQ === 1 ? 4 : curQ - 1, prevYY = curQ === 1 ? curYY - 1 : curYY;
  const NEW = `Mgmt Fee Calc Q${nextQ} ${nextYY}`;
  const PREVCALC = `Mgmt Fee Calc Q${prevQ} ${prevYY}`;
  const CURRECALC = `Q${curQ} ${curYY} Recalc`;
  const NEWRECALC = `Q${nextQ} ${nextYY} Recalc`;
  // next quarter start date (first day of the quarter's first month)
  const nextStart = new Date(Date.UTC(2000 + nextYY, (nextQ - 1) * 3, 1));

  // Prior-quarter references can't be retargeted by sheet name alone: an older
  // quarter tab may use a DIFFERENT column layout (e.g. legacy Q2 has
  // A=InvestorNo, B=InvestorName, T=ITD-before, V=ITD-after, while the current
  // layout has A=InvestorName, S=ITD-before, U=ITD-after). A blind name swap
  // leaves formulas pointing at the wrong columns -> #N/A. So when shifting a
  // reference from PREVCALC up to curName, remap each column letter by its
  // header text (row 17). Same-layout quarters yield an identity map, so this is
  // always safe.
  const colLetterOf = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const headerMapOf = (sheetName) => {
    const w = ewb.getWorksheet(sheetName); const map = {};
    if (w) w.getRow(17).eachCell({ includeEmpty: false }, (c, col) => { if (c.value != null) { const t = String(c.value).trim(); if (!(t in map)) map[t] = colLetterOf(col); } });
    return map;
  };
  const prevHdr = headerMapOf(PREVCALC), curHdr = headerMapOf(curName);
  const prevColByLetter = {}; for (const k in prevHdr) prevColByLetter[prevHdr[k]] = k;
  const prevToCurCol = {}; for (const col in prevColByLetter) { const h = prevColByLetter[col]; if (curHdr[h]) prevToCurCol[col] = curHdr[h]; }
  // Shift a formula's PREVCALC refs up to curName WITH column remap; curName self
  // refs and other sheet names are untouched here.
  const shiftPrevRefs = (xml) => xml.replace(new RegExp("'" + PREVCALC.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'!((?:\\$?[A-Z]{1,2}\\$?\\d*)(?::\\$?[A-Z]{1,2}\\$?\\d*)?)", 'g'), (m, ref) => {
    const nr = ref.replace(/(\$?)([A-Z]{1,2})(\$?\d*)/g, (mm, d1, col, rest) => d1 + (prevToCurCol[col] || col) + rest);
    return "'" + curName + "'!" + nr;
  });

  const { ws, headerRow, cols } = cur;
  const nameC = cols['InvestorName'], endC = cols['InvestorTotal'] || cols['Investor Total'];
  const begC = cols['Beginning InvestorTotal'];
  const chgC = cols['Change in  Commitment in Qtr'] || cols['Change in Commitment in Qtr'] || cols['New Commitment in Qtr'];
  const endingByRow = {};
  const nameByRow = {};
  for (let r = headerRow + 1; r <= headerRow + 90; r++) {
    const nm = ws.getRow(r).getCell(nameC).value;
    if (nm == null || String(nm).trim() === '') continue;
    endingByRow[r] = mgmtNum(ws.getRow(r).getCell(endC).value) || 0;
    nameByRow[r] = String(nm).trim();
  }
  let qStartRow = null;
  for (let r = 1; r <= 16; r++) { const l = ws.getRow(r).getCell(1).value; if (typeof l === 'string' && /quarter start/i.test(l)) { qStartRow = r; break; } }

  // zip-level clone + retarget
  const zip = await JSZip.loadAsync(inputBuf);
  let wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rid2tgt = {};
  for (const rel of relsXml.match(/<Relationship\b[^>]*\/>/g) || []) { const im = rel.match(/Id="(rId\d+)"/), tm = rel.match(/Target="([^"]+)"/); if (im && tm) rid2tgt[im[1]] = tm[1]; }
  const name2info = {};
  for (const tag of wbXml.match(/<sheet\b[^>]*\/>/g) || []) {
    const nm = tag.match(/\bname="([^"]*)"/), rid = tag.match(/r:id="(rId\d+)"/), sid = tag.match(/sheetId="(\d+)"/);
    if (nm && rid) { const name = nm[1].replace(/&apos;/g, "'").replace(/&gt;/g, '>').replace(/&amp;/g, '&'); name2info[name] = { rid: rid[1], sheetId: sid ? +sid[1] : 0, target: rid2tgt[rid[1]].replace(/^\/?xl\//, '') }; }
  }
  if (!name2info[NEW] === false) throw new Error('target tab ' + NEW + ' already exists — already rolled forward?');
  // Capture the ORIGINAL sheet order (by name). Inserting new <sheet> tags shifts
  // every later sheet's positional index, so any definedName scoped with
  // localSheetId="N" (a 0-based sheet index) would silently point at the wrong
  // sheet afterward — which Excel flags as corruption. We remap those indices by
  // name once the final order is known (see end of function).
  const origOrder = [...wbXml.matchAll(/<sheet name="([^"]*)"/g)].map(m => m[1].replace(/&apos;/g, "'").replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
  const allSheetNums = Object.values(name2info).map(i => +i.target.match(/sheet(\d+)\.xml/)[1]);
  let nextSheetNum = Math.max(...allSheetNums) + 1;
  const allRids = (relsXml.match(/Id="rId(\d+)"/g) || []).map(x => +x.match(/\d+/)[0]);
  let nextRid = Math.max(...allRids) + 1;
  const reName = (s, from, to) => s.split("'" + from + "'").join("'" + to + "'");

  const copyWorksheet = async (srcName, transform) => {
    const src = name2info[srcName];
    if (!src) throw new Error('source tab not found: ' + srcName);
    const srcNum = +src.target.match(/sheet(\d+)\.xml/)[1];
    const newNum = nextSheetNum++, newRid = 'rId' + (nextRid++);
    let sx = await zip.file('xl/' + src.target).async('string'); if (transform) sx = transform(sx);
    zip.file('xl/worksheets/sheet' + newNum + '.xml', sx);
    const relsPath = 'xl/worksheets/_rels/sheet' + srcNum + '.xml.rels';
    if (zip.file(relsPath)) {
      let rx = await zip.file(relsPath).async('string'); const deps = [];
      rx = rx.replace(/Target="([^"]+)"/g, (m, t) => {
        const mm = t.match(/(comments|vmlDrawing|drawing)(\d+)\.(xml|vml)/); if (!mm) return m;
        const kind = mm[1], ext = mm[3], dir = (kind === 'vmlDrawing' || kind === 'drawing') ? 'drawings/' : '';
        let k = 1; while (zip.file('xl/' + dir + kind + k + '.' + ext) || deps.find(d => d.newName === kind + k + '.' + ext)) k++;
        const newName = kind + k + '.' + ext, oldName = kind + mm[2] + '.' + ext;
        deps.push({ oldName, newName, dir }); return 'Target="' + t.replace(oldName, newName) + '"';
      });
      for (const d of deps) {
        const srcPath = Object.keys(zip.files).find(p => p.endsWith('/' + d.oldName) || p === 'xl/' + d.dir + d.oldName);
        if (srcPath) { const buf = await zip.file(srcPath).async('nodebuffer'); zip.file(srcPath.replace(d.oldName, d.newName), buf); }
      }
      zip.file('xl/worksheets/_rels/sheet' + newNum + '.xml.rels', rx);
    }
    return { num: newNum, rid: newRid };
  };

  const q4calc = await copyWorksheet(curName, (sx) => shiftPrevRefs(reName(sx, curName, NEW)));
  let q4recalc = null;
  if (name2info[CURRECALC]) q4recalc = await copyWorksheet(CURRECALC, (sx) => reName(sx, curName, NEW));

  // register new sheets
  let extraRels = '<Relationship Id="' + q4calc.rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + q4calc.num + '.xml"/>';
  if (q4recalc) extraRels += '<Relationship Id="' + q4recalc.rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + q4recalc.num + '.xml"/>';
  zip.file('xl/_rels/workbook.xml.rels', relsXml.replace('</Relationships>', extraRels + '</Relationships>'));
  const maxSid = Math.max(...Object.values(name2info).map(i => i.sheetId));
  const curTag = wbXml.match(new RegExp('<sheet name="' + curName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'))[0];
  wbXml = wbXml.replace(curTag, '<sheet name="' + NEW + '" sheetId="' + (maxSid + 1) + '" r:id="' + q4calc.rid + '"/>' + curTag);
  if (q4recalc && name2info[CURRECALC]) {
    const rTag = wbXml.match(new RegExp('<sheet name="' + CURRECALC.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'))[0];
    wbXml = wbXml.replace(rTag, '<sheet name="' + NEWRECALC + '" sheetId="' + (maxSid + 2) + '" r:id="' + q4recalc.rid + '"/>' + rTag);
  }

  // edit NEW calc numeric cells: beginning = prior ending, change = 0, date, B1 label
  let ncXml = await zip.file('xl/worksheets/sheet' + q4calc.num + '.xml').async('string');
  const setNum = (ref, val) => {
    const selfClose = new RegExp('<c r="' + ref + '"([^>]*?)/>');
    const scm = ncXml.match(selfClose);
    if (scm) { ncXml = ncXml.replace(selfClose, (m, attrs) => '<c r="' + ref + '"' + attrs.replace(/\s+t="[^"]*"/, '') + '><v>' + val + '</v></c>'); return; }
    const re = new RegExp('(<c r="' + ref + '"[^>]*>)([\\s\\S]*?)(</c>)');
    if (re.test(ncXml)) ncXml = ncXml.replace(re, (m, open, inner, close) => { const o = open.replace(/\s+t="[^"]*"/, ''); return o + '<v>' + val + '</v>' + close; });
  };
  for (const r in endingByRow) {
    if (begC) setNum(colLetter(begC) + r, endingByRow[r]);
    // Apply the user's per-investor commitment change (default 0). Beginning is the
    // prior ending; InvestorTotal (E) = Beginning + Change, and every fee column
    // keys off E, so a change that zeroes the commitment yields zero fee. A full
    // redemption (change = -beginning) makes E = 0 and the quarterly fee 0.
    if (chgC) setNum(colLetter(chgC) + r, changeByName[nameByRow[r]] || 0);
  }
  if (qStartRow) setNum('B' + qStartRow, excelSerial(nextStart));
  // Normalize the catch-up column. A subsequent-closing investor gets a one-time
  // catch-up fee in the quarter they join, sometimes entered as a MANUAL formula
  // that doesn't key off this-quarter's new commitment (e.g. James Bloomingdale's
  // "=N97*($E$8+$E$9)" charges every quarter, not just the join quarter). On
  // roll-forward that would re-bill the catch-up. The standard catch-up formula
  // (used by ~all rows) is gated on D>0, so it self-zeroes once D resets to 0.
  // So in the NEW tab we replace any non-standard catch-up cell with the standard
  // template (row-number adjusted), making the catch-up correctly 0 next quarter.
  const catchC = cols['Catch-up Management Fee'];
  if (catchC) {
    const cl = colLetter(catchC);
    const cellRe = (r) => new RegExp('<c r="' + cl + r + '"[^>]*>([\\s\\S]*?)</c>');
    // Pick the standard template: the most common catch-up formula shape
    // (an =IF(AND(...="GCM"...>0)...) gated on D). Sample from a known-normal row.
    const fOf = (r) => { const m = ncXml.match(cellRe(r)); if (!m) return null; const fm = m[1].match(/<f[^>]*>([\s\S]*?)<\/f>/); return fm ? fm[1] : null; };
    // NB: <f> content is raw OOXML — no leading "=", and ">" is stored as "&gt;".
    const isStandard = (f) => !!f && /^IF\(AND\(/.test(f) && /="GCM"/.test(f) && /\$D\d+&gt;0/.test(f);
    // find a template row + its row number to substitute
    let tmpl = null, tmplRow = null;
    for (let r = headerRow + 1; r <= headerRow + 90; r++) { const f = fOf(r); if (isStandard(f)) { tmpl = f; tmplRow = r; break; } }
    if (tmpl) {
      for (let r = headerRow + 1; r <= headerRow + 90; r++) {
        const m = ncXml.match(cellRe(r)); if (!m) continue;
        const fm = m[1].match(/<f[^>]*>([\s\S]*?)<\/f>/); if (!fm) continue; // no formula -> leave
        if (isStandard(fm[1])) continue; // already standard
        // build standard formula for THIS row by retargeting the template's row number
        const rowFixed = tmpl.replace(new RegExp('(\\$[A-Z]{1,2})' + tmplRow + '\\b', 'g'), '$1' + r);
        const open = m[0].match(/^<c r="[^"]*"[^>]*>/)[0].replace(/\s+t="[^"]*"/, '');
        const replacement = open + '<f>' + rowFixed + '</f></c>';
        ncXml = ncXml.replace(cellRe(r), () => replacement); // function form: avoids $-substitution in replacement
      }
    }
  }
  // Reset the Transfers column. A transfer (one investor assigning its interest
  // to another, e.g. Stewart Tate -> his revocable trust) is a one-time event in
  // the quarter it happens: the calc tab's Transfers cell carries a hardcoded
  // amount (and a "=-Tnn" mirror on the receiving row), and ITD-after = SUM(S:T)
  // nets it out that quarter. Rolling that forward re-applies the transfer every
  // quarter, so the inception-to-date fee total spirals negative. Transfers must
  // be 0 next quarter (the prior-quarter ending ITD already reflects the moved
  // balance), so we blank every non-empty Transfers cell in the new tab.
  const xferC = cols['Transfers'];
  if (xferC) {
    const xl = colLetter(xferC);
    for (let r = headerRow + 1; r <= headerRow + 90; r++) {
      // Match the self-closing form FIRST. A self-closing cell like
      // `<c r="T18" s="136"/>` must not fall through to the open/close branch:
      // there, `[^>]*>` stops at the `>` of the self-close and `[\s\S]*?</c>`
      // then greedily swallows the NEXT cell up to its </c> (e.g. U18's
      // SUM formula + shared-formula master), corrupting it and triggering
      // Excel's repair prompt. So try `<c .../>` first; only if it's a real
      // open tag (its content not ending in `/`) do we match `<c ...>...</c>`.
      const re = new RegExp('<c r="' + xl + r + '"[^>]*/>|<c r="' + xl + r + '"[^>]*[^/]>[\\s\\S]*?</c>');
      const m = ncXml.match(re); if (!m) continue;
      // only touch cells that actually hold a transfer (a value or a formula); skip already-empty
      if (/<f[^>]*>/.test(m[0]) || /<v>/.test(m[0])) {
        const open = m[0].match(/^<c r="[^"]*"[^>]*?(?=\/?>)/)[0].replace(/\s+t="[^"]*"/, '');
        ncXml = ncXml.replace(re, () => open + '><v>0</v></c>');
      }
    }
  }
  ncXml = ncXml.replace(/<c r="B1"[^>]*>[\s\S]*?<\/c>/, '<c r="B1" t="inlineStr"><is><t>Q' + nextQ + ' 20' + nextYY + '</t></is></c>');
  zip.file('xl/worksheets/sheet' + q4calc.num + '.xml', ncXml);

  // Freeze the now-prior quarter's S column to static values to break the circular
  // reference. The current tab (curName) becomes the prior quarter after roll-forward.
  // Its S column ("ITD Fees + Current Quarter (before transfers)") holds live
  // XLOOKUP('ITD Recalc'!B:B) formulas. Once the new quarter is added, ITD Recalc
  // is rewritten to read the new tab's columns, so 'ITD Recalc' <-> curName becomes
  // a genuine cycle (Excel raises the circular-reference warning). A past quarter
  // should be a fixed historical record, not a live driver — so we replace each
  // S-cell formula in curName with its cached numeric value (which we already have
  // from the ExcelJS-loaded workbook `ws`). Cells with no cached value (inactive
  // sponsor rows whose current-quarter fee is 0) are frozen to 0; their S equals
  // prior-ITD + R where R = 0, and they carry no live ITD contribution.
  {
    const sCol = cols['ITD Fees + Current Quarter (before transfers)'];
    if (sCol) {
      const sL = colLetter(sCol);
      let curSx = await zip.file('xl/' + name2info[curName].target).async('string');
      for (let r = headerRow + 1; r <= headerRow + 90; r++) {
        const ref = sL + r;
        // only touch cells whose formula references ITD Recalc (the cycle source)
        const re = new RegExp('<c r="' + ref + '"([^>]*?)>(?:(?!</c>)[\\s\\S])*?</c>');
        const m = curSx.match(re);
        if (!m || !/ITD Recalc/.test(m[0])) continue;
        let v = mgmtNum(ws.getRow(r).getCell(sCol).value);
        if (v == null) v = 0;
        const attrs = m[1].replace(/\s+t="[^"]*"/, '');
        curSx = curSx.replace(re, '<c r="' + ref + '"' + attrs + '><v>' + v + '</v></c>');
      }
      zip.file('xl/' + name2info[curName].target, curSx);
    }
  }

  for (const tab of ['ITD Recalc']) {
    if (!name2info[tab]) continue;
    let sx = await zip.file('xl/' + name2info[tab].target).async('string');
    sx = shiftPrevRefs(reName(sx, curName, NEW));
    // Same one-time-transfer reset as the calc tab: ITD Recalc has a "Transfers"
    // column (header row 1) whose cells are "=-Bnn" (transferor) / "=-Gnn"
    // (transferee). Left intact they re-subtract the prior ITD every quarter,
    // driving ITD-after negative. Blank them so the rolled-forward ITD only
    // reflects transfers that genuinely occurred in the new quarter (none yet).
    if (tab === 'ITD Recalc') {
      // ITD Recalc's Transfers column cells are "=-Bnn" (transferor) / "=-Gnn"
      // (transferee). Detecting that formula shape is layout-robust; zero them so
      // the rolled-forward ITD only reflects transfers in the new quarter (none).
      sx = sx.replace(/(<c r="[A-Z]+\d+"[^>]*>)<f>=?-[BG]\d+<\/f>(?:<v>[^<]*<\/v>)?(<\/c>)/g, (m, open, close) => {
        const o = open.replace(/\s+t="[^"]*"/, '');
        return o + '<v>0</v>' + close;
      });
      // Normalize column B ("Prior Quarter ITD"). Every investor row should carry
      // the same XLOOKUP into the prior quarter's ITD-after column, but the source
      // workbook sometimes has a few rows hardcoded to a static value (e.g. B81/B82
      // for James Bloomingdale and the Stewart Tate trust were frozen to 0), which
      // drops their prior ITD on roll-forward and breaks the transfer carry-forward.
      // Rebuild any B cell that lost its formula, using a sibling B cell's formula
      // as the row-adjusted template. Only touch rows that (a) have a real investor
      // name in column A and (b) currently lack a <f> in column B.
      {
        // shared strings for resolving column-A investor names
        const ssXmlR = await zip.file('xl/sharedStrings.xml').async('string');
        const ssArr = [];
        for (const si of ssXmlR.match(/<si>[\s\S]*?<\/si>/g) || []) { const t = (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(x => x.replace(/<[^>]+>/g, '')).join(''); ssArr.push(t); }
        // find a template: the first column-B cell that still has a formula
        let tmplRow = null, tmplF = null;
        for (let r = 2; r <= 200; r++) {
          const mm = sx.match(new RegExp('<c r="B' + r + '"[^>]*><f[^>]*>([\\s\\S]*?)<\\/f>'));
          if (mm) { tmplRow = r; tmplF = mm[1]; break; }
        }
        // resolve which column A "Checker"/blank rows to skip: only fix rows whose
        // A cell is a nonempty shared string that isn't the "Checker" sentinel.
        if (tmplF) {
          const aName = (r) => {
            const am = sx.match(new RegExp('<c r="A' + r + '"([^>]*)>(?:<v>(\\d+)<\\/v>)?'));
            if (!am) return null;
            if (/t="s"/.test(am[1]) && am[2] != null) return (ssArr[+am[2]] || '').trim();
            return null;
          };
          for (let r = 2; r <= 200; r++) {
            const nm = aName(r);
            if (!nm || /^checker$/i.test(nm)) continue;
            // does B{r} already have a formula? if so skip
            const bWhole = sx.match(new RegExp('<c r="B' + r + '"[^>]*?(?:/>|>[\\s\\S]*?<\\/c>)'));
            if (!bWhole) continue;
            if (/<f[^>]*>/.test(bWhole[0])) continue;
            // rebuild B{r} with the template formula retargeted to this row
            const rowF = tmplF.replace(new RegExp('([A-Z]{1,2})' + tmplRow + '\\b', 'g'), '$1' + r);
            const open = bWhole[0].match(/^<c r="B\d+"[^>]*?(?=\/?>)/)[0].replace(/\s+t="[^"]*"/, '');
            sx = sx.replace(bWhole[0], open + '><f>' + rowF + '</f></c>');
          }
        }
      }
    }
    zip.file('xl/' + name2info[tab].target, sx);
  }

  // ITD Mgmt Fee: a manually-grown pivot of per-investor fees, one dated column
  // per fee event, with Grand Total in the last column. Unlike ITD Recalc we must
  // NOT shiftPrevRefs here: that would repoint the existing current-quarter column
  // (e.g. the Q3 column) at the new quarter, overwriting the prior quarter's fees
  // and dropping a quarter out of the ITD history. Instead we PRESERVE every
  // existing column and INSERT a new dated column for the new quarter, immediately
  // left of Grand Total, pulling each investor's base quarterly fee (calc col O).
  if (name2info['ITD Mgmt Fee']) {
    const tgtMF = 'xl/' + name2info['ITD Mgmt Fee'].target;
    let sx = await zip.file(tgtMF).async('string');
    // Locate the header row (has the date serials) and the Grand Total column.
    // Grand Total is the last header cell carrying the "Grand Total" string; the
    // dated columns sit to its left. We detect it structurally: find the row whose
    // first column A holds "Grand Total" (the totals row) to bound data rows, and
    // find the header row as the one directly above the first data row.
    // From the workbook's stable layout: header row 5, data 6..85, totals row 86,
    // tie row 87, Grand Total in column R (18). We derive these instead of trusting
    // fixed letters by scanning, so the code survives layout drift.
    const colToNum = (a) => { let n = 0; for (const ch of a) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
    const numToCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    // find totals row: the <row> containing a cell whose shared-string is "Grand Total" in column A
    // (fallback to scanning for a SUM row). We read sharedStrings to resolve A-cell text.
    let ssXmlMF = await zip.file('xl/sharedStrings.xml').async('string');
    const ssArrMF = [];
    for (const si of ssXmlMF.match(/<si>[\s\S]*?<\/si>/g) || []) { const t = (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(x => x.replace(/<[^>]+>/g, '')).join(''); ssArrMF.push(t); }
    // Grand Total column = the header cell whose text = "Grand Total"; header row = its row.
    let gtCol = null, hdrRowMF = null;
    for (const cm of sx.matchAll(/<c r="([A-Z]+)(\d+)"[^>]*\bt="s"[^>]*><v>(\d+)<\/v><\/c>/g)) {
      if (ssArrMF[+cm[3]] === 'Grand Total') { gtCol = colToNum(cm[1]); hdrRowMF = +cm[2]; break; }
    }
    if (gtCol && hdrRowMF) {
      // data rows run from hdrRow+1 up to the totals row (the row whose col-A text is "Grand Total")
      let totalsRow = null;
      for (const rm of sx.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
        const aCell = rm[2].match(/<c r="A\d+"[^>]*\bt="s"[^>]*><v>(\d+)<\/v><\/c>/);
        if (aCell && ssArrMF[+aCell[1]] === 'Grand Total' && +rm[1] > hdrRowMF) { totalsRow = +rm[1]; break; }
      }
      const firstData = hdrRowMF + 1;
      const lastData = totalsRow ? totalsRow - 1 : hdrRowMF + 80;
      const prevLastDated = numToCol(gtCol - 1);   // column just left of Grand Total
      const newL = numToCol(gtCol);                // vacated slot the new column lands in
      const movedGtL = numToCol(gtCol + 1);        // Grand Total after the shift
      // Shift bare (non-sheet-qualified) column refs >= gtCol inside a formula body.
      // Sheet-qualified refs like 'Mgmt Fee Calc Q4 26'!O:O must be left alone.
      const shiftColsInFormula = (f) => f.replace(/(\$?)([A-Z]{1,3})(?=\$?\d|:|$|[^A-Za-z0-9(])/g, (m, dollar, col, off, str) => {
        const prevChar = str[off - 1];
        if (prevChar === '!' || prevChar === "'") return m;
        if (prevChar && /[A-Za-z0-9_]/.test(prevChar)) return m;
        const cn = colToNum(col);
        return cn >= gtCol ? dollar + numToCol(cn + 1) : m;
      });
      const shiftRefAttr = (attrs) => attrs.replace(/\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, (rm, c1, r1, c2, r2) => {
        const n1 = colToNum(c1), n2 = colToNum(c2);
        return 'ref="' + (n1 >= gtCol ? numToCol(n1 + 1) : c1) + r1 + ':' + (n2 >= gtCol ? numToCol(n2 + 1) : c2) + r2 + '"';
      });
      // 1) Shift EVERY column >= gtCol one to the right (Grand Total AND any
      //    annotation/comment columns to its right move together, so nothing
      //    collides). Cell refs, shared ref="" attrs, and formula bodies all shift.
      sx = sx.replace(/<c r="([A-Z]+)(\d+)"([^>]*?)(\/>|>[\s\S]*?<\/c>)/g, (m, col, row, attrs, tail) => {
        const cn = colToNum(col);
        const newCol = cn >= gtCol ? numToCol(cn + 1) : col;
        const newAttrs = shiftRefAttr(attrs);
        const newTail = tail.replace(/<f([^>]*)>([\s\S]*?)<\/f>/g, (fm, fa, ff) =>
          '<f' + shiftRefAttr(fa) + '>' + shiftColsInFormula(ff) + '</f>');
        return '<c r="' + newCol + row + '"' + newAttrs + newTail;
      });
      // 2) Extend Grand Total's per-row SUM to include the new column: the range
      //    end was prevLastDated (unchanged by the shift), widen it to newL.
      sx = sx.replace(new RegExp('(<c r="' + movedGtL + '\\d+"[^>]*>)([\\s\\S]*?)(<\\/c>)', 'g'), (m, open, body, close) => {
        const nb = body.replace(/<f([^>]*)>([\s\S]*?)<\/f>/g, (fm, fa, ff) =>
          '<f' + fa + '>' + ff.replace(new RegExp('SUM\\(([A-Z]+\\d+):' + prevLastDated + '(\\d+)\\)', 'g'), 'SUM($1:' + newL + '$2)') + '</f>');
        return open + nb + close;
      });
      // 3) Insert the new dated column cells at the vacated gtCol (newL), placed
      //    immediately before the moved Grand Total cell in each row.
      const addCell = (rowNum, cellXml) => {
        const re = new RegExp('(<row r="' + rowNum + '"[^>]*>)([\\s\\S]*?)(</row>)');
        if (!re.test(sx)) return;
        sx = sx.replace(re, (m, open, body, close) => {
          const movedRe = new RegExp('(<c r="' + movedGtL + '\\d+"[^>]*?(?:/>|>[\\s\\S]*?</c>))');
          if (movedRe.test(body)) return open + body.replace(movedRe, cellXml + '$1') + close;
          return open + body + cellXml + close;
        });
      };
      const styleMatch = sx.match(new RegExp('<c r="' + prevLastDated + hdrRowMF + '"[^>]*\\bs="(\\d+)"'));
      const dateStyle = styleMatch ? styleMatch[1] : '72';
      addCell(hdrRowMF, '<c r="' + newL + hdrRowMF + '" s="' + dateStyle + '"><v>' + excelSerial(nextStart) + '</v></c>');
      for (let rr = firstData; rr <= lastData; rr++) {
        const rowM = sx.match(new RegExp('<row r="' + rr + '"[^>]*>[\\s\\S]*?</row>'));
        if (!rowM || !new RegExp('<c r="A' + rr + '"[^>]*>').test(rowM[0])) continue;
        const valStyle = (rowM[0].match(new RegExp('<c r="B' + rr + '"[^>]*\\bs="(\\d+)"')) || [null, '16'])[1] || '16';
        // Wrap in IFERROR so investors absent from the new calc tab show 0, not #N/A.
        const fml = '_xlfn.IFERROR(_xlfn.XLOOKUP(A' + rr + ",'" + NEW + "'!A:A,'" + NEW + "'!O:O),0)";
        addCell(rr, '<c r="' + newL + rr + '" s="' + valStyle + '"><f>' + fml + '</f></c>');
      }
      if (totalsRow) {
        const totStyle = (sx.match(new RegExp('<c r="' + movedGtL + totalsRow + '"[^>]*\\bs="(\\d+)"')) || [null, '74'])[1] || '74';
        addCell(totalsRow, '<c r="' + newL + totalsRow + '" s="' + totStyle + '"><f>SUM(' + newL + firstData + ':' + newL + lastData + ')</f></c>');
      }
      // Tie-check row: the Grand-Total column's tie cell compares the ITD grand
      // total against the current calc tab's ITD figure (e.g. +S86-'...Q3 26'!S16).
      // The reference points at curName (the quarter being rolled forward); repoint
      // it to NEW so it validates against the new quarter's calc tab (...'Q4 26'!S16).
      sx = sx.replace(new RegExp('(<c r="' + movedGtL + '\\d+"[^>]*>)([\\s\\S]*?)(<\\/c>)', 'g'), (m, open, body, close) => {
        if (!body.includes("'" + curName + "'")) return m;
        return open + body.split("'" + curName + "'").join("'" + NEW + "'") + close;
      });
      // Normalize the totals row: the source workbook has inconsistent SUM ranges
      // per column (e.g. B86 =SUM(B6:B82), E86 =SUM(E6:E83), others 6:85), which
      // drops the late-inserted investor rows (83-85) from some column totals and
      // makes the Grand Total tie out by an amount equal to those rows' fees.
      // Widen every totals-row SUM in this sheet to the full investor range
      // (firstData:lastData). Covers plain and shared-master formulas alike.
      if (totalsRow) {
        const rowRe = new RegExp('(<row r="' + totalsRow + '"[^>]*>)([\\s\\S]*?)(</row>)');
        sx = sx.replace(rowRe, (m, open, body, close) => {
          const nb = body.replace(/<f([^>]*)>SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)<\/f>/g,
            (fm, fa, c1, r1, c2, r2) => '<f' + fa + '>SUM(' + c1 + firstData + ':' + c2 + lastData + ')</f>');
          return open + nb + close;
        });
      }
      // 4) Recompute each row's spans upper bound from its actual max column, then
      //    widen the sheet dimension and shift <col> width entries >= gtCol.
      sx = sx.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, (m, rnum, rattrs, body) => {
        let maxc = 0, minc = 1e9;
        for (const cm of body.matchAll(/<c r="([A-Z]+)\d+"/g)) { const c = colToNum(cm[1]); if (c > maxc) maxc = c; if (c < minc) minc = c; }
        if (!maxc) return m;
        const lo = minc === 1e9 ? 1 : minc;
        const nattrs = /spans="/.test(rattrs) ? rattrs.replace(/spans="\d+:\d+"/, 'spans="' + lo + ':' + maxc + '"') : rattrs;
        return '<row r="' + rnum + '"' + nattrs + '>' + body + '</row>';
      });
      sx = sx.replace(/<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"\/>/, (m, a, cL, cR) => {
        const n = colToNum(cL); return '<dimension ref="' + a + ':' + numToCol(n >= gtCol ? n + 1 : n) + cR + '"/>';
      });
      sx = sx.replace(/<col min="(\d+)" max="(\d+)"/g, (m, mn, mx) =>
        '<col min="' + (+mn >= gtCol ? +mn + 1 : +mn) + '" max="' + (+mx >= gtCol ? +mx + 1 : +mx) + '"');
      // Give the newly-inserted column (at gtCol) its own width def with bestFit so
      // Excel auto-sizes it and the amounts aren't hidden behind ###. Reuse the
      // width of the dated column just to its left (prevLastDated) when available.
      {
        const prevColN = gtCol - 1;
        // find the specific <col> covering prevColN to copy its width
        let width = '13.26953125';
        for (const cm of sx.matchAll(/<col min="(\d+)" max="(\d+)"[^>]*?width="([\d.]+)"[^>]*?\/>/g)) {
          if (+cm[1] <= prevColN && prevColN <= +cm[2]) { width = cm[3]; break; }
        }
        const newColDef = '<col min="' + gtCol + '" max="' + gtCol + '" width="' + width + '" bestFit="1" customWidth="1"/>';
        if (/<cols>/.test(sx) && !new RegExp('<col min="' + gtCol + '" max="' + gtCol + '"').test(sx)) {
          sx = sx.replace('</cols>', newColDef + '</cols>');
        }
      }
    }
    zip.file(tgtMF, sx);
  }


  // Invoice: repoint formulas cur->NEW; shift dynamic quarter labels; keep annual axis + one-offs
  if (name2info['Invoice']) {
    let invXml = await zip.file('xl/' + name2info['Invoice'].target).async('string');
    invXml = reName(invXml, curName, NEW);
    let ssXml = await zip.file('xl/sharedStrings.xml').async('string');
    const ssList = [];
    for (const si of ssXml.match(/<si>[\s\S]*?<\/si>/g) || []) { const t = (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map(x => x.replace(/<[^>]+>/g, '')).join(''); ssList.push(t.replace(/&amp;/g, '&').replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&apos;/g, "'")); }
    const invCell = {};
    for (const m of invXml.matchAll(/<c r="([A-Z]+)(\d+)"[^>]*\bt="s"[^>]*><v>(\d+)<\/v><\/c>/g)) { const t = ssList[+m[3]]; if (t != null) invCell[m[1] + m[2]] = { col: m[1], row: +m[2], txt: t }; }
    const qOf = (t) => { const mm = t.match(/Q([1-4])'?\s?(\d{2}|20\d{2})/); return mm ? { q: +mm[1], y: mm[2].length === 2 ? 2000 + +mm[2] : +mm[2] } : null; };
    const annualAxis = new Set(); const byCol = {};
    for (const k in invCell) { (byCol[invCell[k].col] = byCol[invCell[k].col] || []).push(invCell[k]); }
    for (const col in byCol) {
      const cells = byCol[col].sort((a, b) => a.row - b.row);
      for (let i = 0; i < cells.length; i++) {
        let run = [cells[i]], prev = qOf(cells[i].txt); if (!prev) continue;
        for (let j = i + 1; j < cells.length; j++) { if (cells[j].row !== cells[j - 1].row + 1) break; const c = qOf(cells[j].txt); if (!c) break; if (c.y === prev.y && c.q === prev.q + 1) { run.push(cells[j]); prev = c; } else break; }
        if (run.length >= 3) run.forEach(c => annualAxis.add(c.col + c.row));
      }
    }
    // The "ITD Activity" ledger (row 21 down) is a manually maintained,
    // inception-to-date transaction table. It carries its own quarter labels
    // (e.g. "Q3'26 Expense") but those are historical entries, not a current-
    // quarter pointer — the accountant adds a new row each quarter by hand. So we
    // never shift anything at row 21 or below; only the summary/YTD area above it.
    const ITD_ACTIVITY_ROW = 21;
    const shiftAllowed = (ref, txt) => {
      const rowNum = parseInt((ref.match(/\d+/) || ['0'])[0], 10);
      if (rowNum >= ITD_ACTIVITY_ROW) return false;
      if (annualAxis.has(ref)) return false;
      if (/Legacy Knight|Bloomingdale|capital call|dtd /i.test(txt)) return false;
      return /Q[1-4]'?\s?2?6\b/.test(txt) && /(Management|Mangement|Mgmt|Catch[- ]?Up|Expense|ITD|Payment)/i.test(txt);
    };
    invXml = invXml.replace(/<c r="([A-Z]+\d+)"([^>]*)\bt="s"([^>]*)><v>(\d+)<\/v><\/c>/g, (m, ref, a1, a2, idx) => {
      const txt = ssList[+idx]; if (txt == null || !shiftAllowed(ref, txt)) return m;
      const shifted = mgmtShiftQuarterText(txt).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      return '<c r="' + ref + '"' + a1 + a2 + ' t="inlineStr"><is><t xml:space="preserve">' + shifted + '</t></is></c>';
    });
    zip.file('xl/' + name2info['Invoice'].target, invXml);
  }

  // content types for the new sheets + their cloned comments/drawings
  let ct = await zip.file('[Content_Types].xml').async('string');
  const addOv = (p, t) => { if (!ct.includes('PartName="' + p + '"')) ct = ct.replace('</Types>', '<Override PartName="' + p + '" ContentType="' + t + '"/></Types>'); };
  addOv('/xl/worksheets/sheet' + q4calc.num + '.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  if (q4recalc) addOv('/xl/worksheets/sheet' + q4recalc.num + '.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  for (const p of Object.keys(zip.files)) {
    if (/xl\/comments\d+\.xml$/.test(p)) addOv('/' + p, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
    if (/xl\/drawings\/drawing\d+\.xml$/.test(p)) addOv('/' + p, 'application/vnd.openxmlformats-officedocument.drawing+xml');
  }
  zip.file('[Content_Types].xml', ct);
  // The workbook contains a pre-existing circular reference: ITD Recalc B85 =
  // C84 - <curQ>!G16, reached through whole-column XLOOKUP ranges (e.g. a calc
  // tab's S column reads 'ITD Recalc'!B:B, whose B85 reads the calc tab's G16
  // total). The source file masks it by saving cached values and never forcing a
  // recompute. We need the opposite — the new quarter's totals must recompute on
  // open so the accountant sees correct figures without pressing F9 — so we set
  // fullCalcOnLoad. To stop that recompute from raising Excel's circular-reference
  // warning, we also enable iterative calculation (the same toggle as Excel's
  // "Enable iterative calculation" option), which resolves the benign cycle by
  // convergence. Verified: under iteration the totals settle to the correct
  // values (Q4 beginning ties to Q3 ending) with zero error cells. calcChain is
  // dropped so Excel rebuilds it cleanly for the new sheets.
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    ct = ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''); zip.file('[Content_Types].xml', ct);
    // Also drop the workbook->calcChain relationship. Removing the part and its
    // content-type Override is not enough: xl/_rels/workbook.xml.rels still holds
    // a <Relationship .../calcChain Target="calcChain.xml"/> that now points at a
    // deleted part. That dangling rel is exactly what triggers Excel's "we found
    // a problem with some content" repair prompt on open. Re-read the rels we
    // already rewrote (new-sheet rels appended) and strip the calcChain entry.
    let wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    wbRels = wbRels.replace(/<Relationship[^>]*Type="[^"]*\/calcChain"[^>]*\/>/g, '');
    zip.file('xl/_rels/workbook.xml.rels', wbRels);
  }
  const calcAttrs = ' iterate="1" iterateCount="100" iterateDelta="0.001" fullCalcOnLoad="1"';
  if (/<calcPr/.test(wbXml)) wbXml = wbXml.replace(/<calcPr([^/]*)\/>/, (m, a) => '<calcPr' + a.replace(/\s+(iterate|iterateCount|iterateDelta|fullCalcOnLoad)="[^"]*"/g, '') + calcAttrs + '/>');
  else wbXml = wbXml.replace('</workbook>', '<calcPr calcId="0"' + calcAttrs + '/></workbook>');
  // Remap localSheetId indices to the FINAL sheet order (see origOrder note above).
  const finalOrder = [...wbXml.matchAll(/<sheet name="([^"]*)"/g)].map(m => m[1].replace(/&apos;/g, "'").replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
  const newIdxByName = {}; finalOrder.forEach((n, i) => { newIdxByName[n] = i; });
  const idxRemap = {}; origOrder.forEach((n, i) => { if (newIdxByName[n] != null) idxRemap[i] = newIdxByName[n]; });
  wbXml = wbXml.replace(/localSheetId="(\d+)"/g, (m, d) => { const n = idxRemap[+d]; return 'localSheetId="' + (n == null ? d : n) + '"'; });
  zip.file('xl/workbook.xml', wbXml);

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { outBuf, label: `Q${nextQ} ${nextYY}`, newTab: NEW, investors: Object.keys(endingByRow).length };
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
    try {
      let changes = [];
      try { changes = JSON.parse(req.body.changes || '[]'); } catch { changes = []; }
      const { outBuf, label, newTab, investors } = await mgmtRollForward(req.file.buffer, changes);
      const fname = 'CLRF_Mgmt_Fee_Calc_' + label.replace(/\s+/g, '_') + '.xlsx';
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
      res.setHeader('X-Mgmt-Fee-Summary', JSON.stringify({ quarter: label, new_tab: newTab, investors }).replace(/[\r\n]/g, ' '));
      res.send(outBuf);
    } catch (e) {
      res.status(500).json({ error: 'Generate error: ' + e.message });
    }
  });
});

// ═══ CLRF workpaper: Schedule of Fees Paid to the GP & Affiliates ═══
// Quarterly. Reads the four CLRF portfolio-company ledgers, builds the schedule,
// saves one copy per period under the entity's workpaper folder, and returns the
// .xlsx. See server/gpfees.js for the development-fee measurement rule.
require('./gpfees').registerGpFeesRoutes(app, {
  db,
  auth,
  requireEntityAccess,
  requireRole,
  workpapersDir: WORKPAPERS_DIR,
  computeBalances: (eid, opts) => computeBalances(eid, opts),
});

// ═══ CLIP Development Costs ═══
// GET /api/entities/:eid/dev-costs?as_of=YYYY-MM-DD — total capitalized
// development cost (Total Long Term Investments + Total Other Assets), the
// figure that feeds the CLIP line of the CLRF valuation schedule. See
// server/devcosts.js for the account set and exclusions.
require('./devcosts').registerDevCostsRoutes(app, {
  auth,
  requireEntityAccess,
  computeBalances: (eid, opts) => computeBalances(eid, opts),
});

// ═══ CLRF Investment & Valuation workpapers ═══
// POST /api/workpapers/investment-valuation/:entity_id/generate { quarter_end }
// One run produces TWO workbooks under Workpapers/Investment & Valuation/<Qn YYYY>:
// the Investment workpaper (portfolio TBs, NWC/loans, hypothetical-liquidation
// waterfall, SOLVED valuations under the frozen-unrealized-gain convention) and
// the Valuation workbook generated against those solved amounts so its Summary
// matches the investment workpaper's Valuations tab exactly.
// See server/invval.js and server/valuation.js.
require('./invval').registerInvValRoutes(app, {
  db,
  auth,
  requireEntityAccess,
  requireRole,
  workpapersDir: WORKPAPERS_DIR,
  computeBalances: (eid, opts) => computeBalances(eid, opts),
});

// ═══ Intercompany (IC Mapping + IC Reconciliation) ═══
// A top-level section, peer to A/R and A/P. Two things live here:
//   • IC Mapping   — setup: which GL account on which entity faces which other
//                     entity, and what kind of intercompany balance it is.
//                     A person confirms every row; name parsing only suggests.
//   • IC Reconcile — Due from / Due to (transactional, netted by counterparty
//                     pair) and Investment / Contributed capital (structural).
// A counterparty outside the selected group is tagged "no elim" and is never
// eliminated — see server/intercompany.js for why that rule exists.
// All endpoints are Admin/Accountant; entity access is checked per group member
// inside the module because these reports span several entities at once.
require('./intercompany').registerIntercompanyRoutes(app, {
  db,
  auth,
  requireRole,
  userHasEntityAccess,
  computeBalances: (eid, opts) => computeBalances(eid, opts),
  workpapersDir: WORKPAPERS_DIR,
});

// === Org structure (ownership tree) ===
// Who owns whom, from the legal org charts, including holding companies that
// hold investment balances but keep no ledger here. Drives two things the
// entities table cannot express: effective ownership / NCI through a
// multi-level chain, and the look-through tie of a fund's investment account
// to the contributed capital of the operating company it funded (CLRF's
// investment names the operating company; the operating company's capital
// names an intermediate shell, so neither side names the other).
// See server/orgstructure.js.
require('./orgstructure').registerOrgStructureRoutes(app, {
  db,
  auth,
  requireRole,
  userHasEntityAccess,
  computeBalances: (eid, opts) => computeBalances(eid, opts),
});

// ═══ Financial Statements package generator ═══
// Generates GL-derived financial statements (Balance Sheet, Operations, Cash
// Flows, Members' Equity) for an entity as of a date, on a monthly/quarterly/
// annually basis, then assembles a single merged PDF package:
//   cover -> executive summary (uploaded) -> GL statements -> requisition report
//   (uploaded, with Current/Prior Invoice Log pages stripped).
const finStmtUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 64 * 1024 * 1024 } });
const finStmtFields = finStmtUpload.fields([{ name: 'execSummary', maxCount: 1 }, { name: 'reqReport', maxCount: 2 }]);

// Preview endpoint: returns the numeric statements + tie-out checks as JSON,
// so the UI can show balance-sheet / cash-flow tie-outs before generating.
app.post('/api/workpapers/financial-statements/:entity_id/preview', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.entity_id;
    const asOf = (req.body && req.body.as_of) || (req.query && req.query.as_of);
    const period = ((req.body && req.body.period) || (req.query && req.query.period) || 'monthly');
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name, code FROM entities WHERE id=?').get(eid);
    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const s = await financials.buildStatements(getBalances, { asOf, period, entityName: ent ? ent.name : ('Entity ' + eid), entityCode: ent ? ent.code : '' });
    res.json({
      meta: s.meta,
      checks: s.checks,
      totals: {
        totalAssets: s.balanceSheet.totalAssets, totalLiabEquity: s.balanceSheet.totalLiabEquity,
        totalEquity: s.balanceSheet.totalEquity, netIncomeYtd: s.operations.netIncome.ytd,
        cashEnd: s.cashFlow.cashEnd, cashFlowTieOut: s.cashFlow.tieOut,
      },
    });
  } catch (e) {
    res.status(500).json({ error: 'Preview error: ' + e.message });
  }
});

// Trailing 12 Months P&L — 12 monthly columns (oldest→newest) + a Total column.
// Returns the full matrix as JSON; the client renders it and exports to Excel.
app.get('/api/entities/:eid/ttm-pl', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.eid;
    const asOf = (req.query && req.query.as_of);
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const out = await financials.buildTtmPL(getBalances, { asOf, entityName: ent ? ent.name : ('Entity ' + eid) });
    res.json(out);
  } catch (e) {
    res.status(500).json({ error: 'TTM P&L error: ' + e.message });
  }
});

// Trailing 12 Months — Claude-powered "Items Needing Attention" analysis.
// Takes an already-built TTM P&L matrix and returns { generatedBy, lastMonthLabel,
// summary, findings[], hasFindings }. Throws on a missing key or an API/parse
// failure so callers can decide whether to surface or swallow the error.
async function buildTtmAnalysis(d, entityName, asOf) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    const e = new Error('Analysis is not configured (ANTHROPIC_API_KEY missing on the server).');
    e.code = 'NO_API_KEY';
    throw e;
  }

  // Compact, model-friendly rendering of the matrix: one line per account with
  // its 12 monthly values and the trailing-12 total, grouped by section.
  const monthLabels = d.meta.months.map(m => m.label);
  const fmtRow = (name, vals, total) => name + ' | ' + vals.map(v => Math.round(Number(v) || 0)).join(', ') + ' | TTM: ' + Math.round(Number(total) || 0);
  const lines = [];
  lines.push('Months (oldest to newest): ' + monthLabels.join(', '));
  lines.push('');
  lines.push('REVENUE:');
  d.revenue.forEach(l => lines.push('  ' + fmtRow(l.name, l.vals, l.total)));
  lines.push('  ' + fmtRow('TOTAL REVENUE', d.totRev.vals, d.totRev.total));
  if (d.hasCogs) {
    lines.push('COST OF REVENUE:');
    d.cogs.forEach(l => lines.push('  ' + fmtRow(l.name, l.vals, l.total)));
    lines.push('  ' + fmtRow('TOTAL COST OF REVENUE', d.totCogs.vals, d.totCogs.total));
    lines.push('  ' + fmtRow('GROSS PROFIT', d.grossProfit.vals, d.grossProfit.total));
  }
  lines.push('OPERATING EXPENSES:');
  d.opexGroups.forEach(g => {
    lines.push(' ' + g.title + ':');
    g.lines.forEach(l => lines.push('    ' + fmtRow(l.name, l.vals, l.total)));
    lines.push('    ' + fmtRow('Total ' + g.title, g.subtotal.vals, g.subtotal.total));
  });
  lines.push('  ' + fmtRow('TOTAL OPERATING EXPENSES', d.totOpex.vals, d.totOpex.total));
  lines.push('  ' + fmtRow('NET INCOME (LOSS)', d.netIncome.vals, d.netIncome.total));
  const matrixText = lines.join('\n');

  const instruction =
    'You are a senior accountant reviewing a Trailing 12 Months profit-and-loss report for ' + entityName +
    ', ending ' + asOf + '. Each line shows 12 monthly amounts (oldest to newest) then the trailing-12-month total. ' +
    'Revenue and expenses are shown as positive magnitudes; Net Income is revenue minus expenses.\n\n' +
    'Identify the items that NEED ATTENTION — focus on UNFAVORABLE things a CAO would want flagged: ' +
    'expenses trending or spiking up, revenue declining or dropping to zero, unusual one-off movements, ' +
    'volatile lines, negative gross profit or net losses, and anything that looks like a possible posting gap ' +
    '(e.g. a normally-active account suddenly at zero in the latest month). Ignore trivially small amounts. ' +
    'Do NOT flag favorable movements as problems (e.g. expenses going down or revenue going up are good).\n\n' +
    'Return ONLY a JSON object, no prose, no markdown fences, in exactly this shape:\n' +
    '{"summary": string, "findings": [{"account": string, "reason": string}]}\n' +
    '- summary: 1-2 sentence plain-English overview of the entity\'s trailing-12 performance and the main concern.\n' +
    '- findings: the items needing attention, ORDERED BY IMPORTANCE (most important/material first). ' +
    'Each finding has "account" (the account or line-item name exactly as shown) and "reason" (one concise ' +
    'sentence saying why it needs attention, citing the specific numbers/months). ' +
    'Aim for the 3-8 most material items; return an empty array if nothing warrants attention.\n\n' +
    'THE REPORT:\n' + matrixText;

  const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      messages: [{ role: 'user', content: [{ type: 'text', text: instruction }] }],
    }),
  });
  if (!apiRes.ok) {
    const t = await apiRes.text();
    const e = new Error('Analysis failed (Anthropic ' + apiRes.status + '): ' + t.slice(0, 300));
    e.code = 'ANTHROPIC_ERROR';
    e.status = apiRes.status;
    throw e;
  }
  const data = await apiRes.json();
  const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  let clean = text.replace(/```json|```/g, '').trim();
  let parsed;
  try {
    parsed = JSON.parse(clean);
  } catch (firstErr) {
    const start = clean.indexOf('{');
    if (start === -1) throw firstErr;
    let depth = 0, end = -1, inStr = false, esc = false;
    for (let i = start; i < clean.length; i++) {
      const ch = clean[i];
      if (inStr) { if (esc) esc = false; else if (ch === '\\') esc = true; else if (ch === '"') inStr = false; }
      else if (ch === '"') inStr = true;
      else if (ch === '{') depth++;
      else if (ch === '}') { depth--; if (depth === 0) { end = i; break; } }
    }
    if (end === -1) throw firstErr;
    parsed = JSON.parse(clean.slice(start, end + 1));
  }
  const findings = Array.isArray(parsed.findings) ? parsed.findings.map(f => {
    const account = String(f.account || f.title || '').slice(0, 200);
    const reason = String(f.reason || f.detail || '').slice(0, 500);
    // title/detail retained for the Excel export and older clients.
    return { account, reason, title: account, detail: reason };
  }) : [];
  return {
    generatedBy: 'claude-haiku-4-5-20251001',
    lastMonthLabel: monthLabels[monthLabels.length - 1] || '',
    summary: String(parsed.summary || '').slice(0, 1000),
    findings,
    hasFindings: findings.length > 0,
  };
}

// Trailing 12 Months P&L — styled Excel export (12 monthly columns + Total).
// Amounts are comma-styled; the month-header row and every subtotal/grand-total
// row are underlined (bottom border). Built with ExcelJS server-side because the
// client SheetJS community build cannot write cell borders/styles.
// Styled workbook for any client report. The client keeps building its rows and
// live formulas exactly as before and adds a style spec naming the subtotal and
// grand-total rows; this returns the .xlsx with the underlines drawn. Needed
// because the community SheetJS build the client exports with cannot write cell
// borders (CLA items 2-4, 8/2026).
app.post('/api/xlsx-styled', auth, async (req, res) => {
  try {
    const spec = req.body || {};
    if (!Array.isArray(spec.rows) || !spec.rows.length) return res.status(400).json({ error: 'rows required' });
    // A report is a few hundred rows; anything far larger is a client bug, and
    // building it would tie up the process.
    if (spec.rows.length > 50000) return res.status(413).json({ error: 'too many rows' });
    const buf = await xlsxStyledReport.buildStyledWorkbookBuffer(spec);
    const fn = String(spec.filename || 'Report.xlsx').replace(/[^A-Za-z0-9._-]+/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fn + '"');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error('xlsx-styled failed: ' + e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/entities/:eid/ttm-pl.xlsx', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.eid;
    const asOf = (req.body && req.body.as_of) || (req.query && req.query.as_of);
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
    const entityName = ent ? ent.name : ('Entity ' + eid);
    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const d = await financials.buildTtmPL(getBalances, { asOf, entityName });

    const nCols = d.meta.months.length; // 12
    const NUMFMT = '#,##0.00;(#,##0.00)';
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Trailing 12 Months');
    const totalCol = 2 + nCols; // col A = account, B..(1+n) = months, last = Total

    // Title block.
    ws.getCell(1, 1).value = entityName;
    ws.getCell(1, 1).font = { bold: true, size: 13 };
    ws.getCell(2, 1).value = d.meta.title;
    ws.getCell(2, 1).font = { bold: true, size: 11 };
    ws.getCell(3, 1).value = d.meta.periodLabel;
    ws.getCell(3, 1).font = { italic: true, size: 10 };

    // Header row (row 5): Account, month labels, Total — all underlined.
    const headerRowIdx = 5;
    const hdr = ['', ...d.meta.months.map(m => m.label), d.meta.totalLabel || 'Total'];
    for (let c = 1; c <= totalCol; c++) {
      const cell = ws.getCell(headerRowIdx, c);
      cell.value = hdr[c - 1];
      cell.font = { bold: true, size: 10 };
      cell.alignment = { horizontal: c === 1 ? 'left' : 'right' };
      cell.border = { bottom: { style: 'thin' } };
    }

    let r = headerRowIdx + 1;
    // Emit one worksheet row from a display line. underline: 'single'|'double'|none.
    const emit = (label, vals, total, opt = {}) => {
      const rowIdx = r++;
      const nameCell = ws.getCell(rowIdx, 1);
      const indent = opt.indent || 0;
      nameCell.value = label;
      nameCell.font = { bold: !!(opt.bold || opt.header || opt.sub), size: 10 };
      nameCell.alignment = { indent };
      if (opt.header) return rowIdx; // header/label row: no numbers
      for (let i = 0; i < nCols; i++) {
        const cell = ws.getCell(rowIdx, 2 + i);
        cell.value = (vals && vals[i] != null) ? Number(vals[i]) : null;
        cell.numFmt = NUMFMT;
        cell.font = { bold: !!opt.bold, size: 10 };
      }
      const tcell = ws.getCell(rowIdx, totalCol);
      tcell.value = total == null ? null : Number(total);
      tcell.numFmt = NUMFMT;
      tcell.font = { bold: true, size: 10 };
      if (opt.underline) {
        const style = opt.underline === 'double' ? 'double' : 'thin';
        for (let c = 2; c <= totalCol; c++) ws.getCell(rowIdx, c).border = { bottom: { style } };
      }
      return rowIdx;
    };

    // Revenue
    emit('Revenue', null, null, { header: true });
    d.revenue.forEach(l => emit(l.name, l.vals, l.total, { indent: 1 }));
    emit('Total Revenue', d.totRev.vals, d.totRev.total, { bold: true, underline: 'single' });
    // Cost of Revenue (only if present)
    if (d.hasCogs) {
      emit('Cost of Revenue', null, null, { header: true });
      d.cogs.forEach(l => emit(l.name, l.vals, l.total, { indent: 1 }));
      emit('Total Cost of Revenue', d.totCogs.vals, d.totCogs.total, { bold: true, underline: 'single' });
      emit('Gross Profit', d.grossProfit.vals, d.grossProfit.total, { bold: true, underline: 'single' });
    }
    // Operating Expenses, grouped
    emit('Operating Expenses', null, null, { header: true });
    d.opexGroups.forEach(g => {
      if (d.opexGroups.length > 1) {
        emit(g.title, null, null, { indent: 1, sub: true });
        g.lines.forEach(l => emit(l.name, l.vals, l.total, { indent: 2 }));
        emit('Total ' + g.title, g.subtotal.vals, g.subtotal.total, { indent: 1, underline: 'single' });
      } else {
        g.lines.forEach(l => emit(l.name, l.vals, l.total, { indent: 1 }));
      }
    });
    emit('Total Operating Expenses', d.totOpex.vals, d.totOpex.total, { bold: true, underline: 'single' });
    // Net Income (grand total) — double underline.
    emit('Net Income (Loss)', d.netIncome.vals, d.netIncome.total, { bold: true, underline: 'double' });
    // -- Analysis: Items Needing Attention (generated automatically) -----------
    // The on-screen section was removed; the export now always asks Claude for
    // the review and appends it below the report. A failure here must not break
    // the export, so it is swallowed and the workbook ships without the block.
    let a = null;
    try {
      a = await buildTtmAnalysis(d, entityName, asOf);
    } catch (e) {
      console.error('TTM export analysis skipped: ' + e.message);
    }
    if (a && (a.summary || (a.findings && a.findings.length))) {
      const DASH = String.fromCharCode(8212);  // em dash
      r += 1; // blank spacer row
      const titleRow = r++;
      const tCell = ws.getCell(titleRow, 1);
      tCell.value = 'Items Needing Attention';
      tCell.font = { bold: true, size: 11 };
      tCell.border = { bottom: { style: 'thin' } };
      if (a.summary) {
        const sRow = r++;
        const sCell = ws.getCell(sRow, 1);
        sCell.value = a.summary;
        sCell.font = { italic: true, size: 10, color: { argb: 'FF444444' } };
        sCell.alignment = { wrapText: false };
      }
      const writeFinding = (text) => {
        const rowIdx = r++;
        const c = ws.getCell(rowIdx, 1);
        c.value = text;
        c.font = { size: 10 };
        c.alignment = { indent: 1 };
      };
      if (a.findings && a.findings.length) {
        // Numbered by importance (findings arrive already ordered most-important first).
        a.findings.forEach((it, i) => {
          const account = it.account || it.title || '';
          const reason = it.reason || it.detail || '';
          const sep = account && reason ? (' ' + DASH + ' ') : '';
          writeFinding((i + 1) + '. ' + account + sep + reason);
        });
      } else {
        writeFinding('Nothing flagged this period.');
      }
      const genRow = r++;
      const gCell = ws.getCell(genRow, 1);
      gCell.value = 'Analysis generated by Claude.';
      gCell.font = { italic: true, size: 8, color: { argb: 'FF999999' } };
    }

    // Column widths: account column wide, numeric columns comfortable.
    ws.getColumn(1).width = 34;
    for (let c = 2; c <= totalCol; c++) ws.getColumn(c).width = 14;

    const buf = await wb.xlsx.writeBuffer();
    const fname = (entityName.replace(/[^a-zA-Z0-9]+/g, '_')) + '_Trailing_12_Months_' + asOf + '.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(500).json({ error: 'TTM export error: ' + e.message });
  }
});

// Trailing 12 Months — Claude-powered analysis of items needing attention.
// On-demand (POST). Sends the 12-month P&L matrix to Claude Haiku and returns a
// structured set of findings + a short narrative. Requires ANTHROPIC_API_KEY.
app.post('/api/entities/:eid/ttm-pl/analyze', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.eid;
    const asOf = (req.body && req.body.as_of) || (req.query && req.query.as_of);
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
    const entityName = ent ? ent.name : ('Entity ' + eid);
    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const d = await financials.buildTtmPL(getBalances, { asOf, entityName });
    res.json(await buildTtmAnalysis(d, entityName, asOf));
  } catch (e) {
    if (e.code === 'NO_API_KEY') return res.status(503).json({ error: e.message });
    if (e.code === 'ANTHROPIC_ERROR') return res.status(502).json({ error: e.message });
    res.status(500).json({ error: 'Analysis error: ' + e.message });
  }
});

// Read an entity's stored default Executive Summary (single-page PDF) from its
// Workpapers tree, or null if none has been uploaded/split for it yet.
function readStoredExecSummary(eid) {
  const row = db.prepare(
    'SELECT stored_filename FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name=? ORDER BY id DESC LIMIT 1'
  ).get(eid, execSummaries.DEFAULT_FOLDER, execSummaries.DEFAULT_FILENAME);
  if (!row) return null;
  try {
    const p = path.resolve(WORKPAPERS_DIR, String(eid), row.stored_filename);
    return fs.readFileSync(p);
  } catch (_) { return null; }
}

// Persist a single-page Executive Summary PDF as an entity's stored default,
// overwriting any prior default. Returns the entity_files row info.
function writeStoredExecSummary(eid, buffer, who) {
  ensureWpFolders(db, eid, execSummaries.DEFAULT_FOLDER, who);
  return saveWpBuffer(db, WORKPAPERS_DIR, eid, execSummaries.DEFAULT_FOLDER,
    execSummaries.DEFAULT_FILENAME, 'application/pdf', buffer, who, { overwrite: true });
}

app.post('/api/workpapers/financial-statements/:entity_id/generate', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
  finStmtFields(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    try {
      const eid = req.params.entity_id;
      const asOf = req.body.as_of;
      const period = req.body.period || 'monthly';
      if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
      const ent = db.prepare('SELECT name, code FROM entities WHERE id=?').get(eid);
      const entityName = ent ? ent.name : ('Entity ' + eid);
      const entityCode = ent ? ent.code : '';

      const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
      const statements = await financials.buildStatements(getBalances, { asOf, period, entityName, entityCode });

      const files = req.files || {};
      const execSummaryBytes = files.execSummary && files.execSummary[0] ? files.execSummary[0].buffer : null;
      // Up to two requisition reports (rail-assets entities may pair two). Each
      // becomes its own section + Table-of-Contents entry, auto-numbered when
      // more than one is present.
      const reqFiles = (files.reqReport || []);
      const reqSheetNames = [].concat(req.body.req_sheet || []); // optional per-file sheet override(s)
      const reqReports = reqFiles.map((f, i) => ({
        bytes: f.buffer,
        name: f.originalname,
        sheet: reqSheetNames[i] || undefined,
      }));
      // Back-compat single-report fields (first file), kept so an older client
      // or any other caller of generatePackage still works.
      const reqReportBytes = reqReports[0] ? reqReports[0].bytes : null;
      const reqReportName = reqReports[0] ? reqReports[0].name : null;
      const reqSheetName = reqReports[0] ? reqReports[0].sheet : undefined;

      // When the user uploads an exec summary with this generate call, it both
      // goes into THIS package and becomes the entity's new stored default.
      // Otherwise fall back to the stored default (if any) for the merge.
      const who = req.user ? (req.user.name || req.user.email) : 'system';
      let storedDefaultBytes = null;
      if (execSummaryBytes) {
        try { writeStoredExecSummary(eid, execSummaryBytes, who); }
        catch (e) { console.error('exec-summary persist failed:', e.message); }
      } else {
        storedDefaultBytes = readStoredExecSummary(eid);
      }

      const { bytes, info } = await financials.generatePackage({ statements, execSummaryBytes, storedDefaultBytes, reqReports, reqReportBytes, reqReportName, reqSheetName });

      const mm = asOf.slice(5, 7), yyyy = asOf.slice(0, 4);
      const safeName = entityName.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
      const fname = safeName + '_Financial_Statements_' + mm + '_' + yyyy + '.pdf';
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
      res.setHeader('X-Financials-Summary', JSON.stringify({
        pages: info.pages, sections: info.sections, warnings: info.warnings,
        reqRemoved: info.reqRemoved || [], reqKept: info.reqKept, reqTotal: info.reqTotal,
        reqReports: info.reqReports || null,
        reqConvertedFromXlsx: info.reqConvertedFromXlsx || false, reqSheetUsed: info.reqSheetUsed,
        balanceSheetTies: info.balanceSheetTies, cashFlowTies: info.cashFlowTies, cashFlowDiff: info.cashFlowDiff,
        execSummarySource: info.execSummarySource,
        checks: statements.checks,
      }).replace(/[\r\n]/g, ' '));
      res.send(Buffer.from(bytes));
    } catch (e) {
      res.status(500).json({ error: 'Generate error: ' + e.message });
    }
  });
});

// Excel export of the SAME GL-derived statements the PDF renders, formatted to
// mirror the PDF: one worksheet per statement (Balance Sheet, Statements of
// Operations, Statement of Cash Flows, Statement of Changes in Members' Equity),
// same titles/headers/indentation/subtotal rules. Built from buildStatements()
// so the numbers are identical to the PDF by construction.
app.get('/api/workpapers/financial-statements/:entity_id/excel', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.entity_id;
    const asOf = req.query && req.query.as_of;
    const period = (req.query && req.query.period) || 'monthly';
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name, code FROM entities WHERE id=?').get(eid);
    const entityName = ent ? ent.name : ('Entity ' + eid);
    const entityCode = ent ? ent.code : '';
    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const statements = await financials.buildStatements(getBalances, { asOf, period, entityName, entityCode });
    const buf = await financialsXlsx.buildStatementsWorkbook(statements);
    const mm = asOf.slice(5, 7), yyyy = asOf.slice(0, 4);
    const safeName = entityName.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    const fname = safeName + '_Financial_Statements_' + mm + '_' + yyyy + '.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
    res.setHeader('X-Financials-Summary', JSON.stringify({ checks: statements.checks }).replace(/[\r\n]/g, ' '));
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(500).json({ error: 'Excel export error: ' + e.message });
  }
});

// Extract plain text per page from a PDF buffer (for entity matching). Lazy pdfjs.
async function extractPdfPageTexts(buffer) {
  const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(buffer), isEvalSupported: false, verbosity: 0 }).promise;
  const out = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    out.push(tc.items.map(i => i.str).join(' ').replace(/\s+/g, ' ').trim());
  }
  return out;
}

// Split a multi-page master Executive Summary PDF into single-page defaults,
// one per entity, matched by the entity title text on each page. Each matched
// page becomes that entity's stored default (Executive Summary/executive_summary.pdf,
// overwriting any prior). The uploaded master is archived in an Admin-only
// Workpapers folder. Admin only.
//
// Body (multipart): master = the combined PDF. Optional master_home_eid = the
// entity whose Workpapers tree stores the archived master (defaults to the first
// matched entity, or the caller-supplied value).
app.post('/api/admin/exec-summaries/split', auth, requireRole('Admin'), (req, res) => {
  memUpload.single('master')(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    try {
      if (!req.file || !req.file.buffer) return res.status(400).json({ error: 'master PDF is required' });
      const who = req.user ? (req.user.name || req.user.email) : 'system';
      const { PDFDocument } = require('pdf-lib');

      const masterDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
      const pageCount = masterDoc.getPageCount();
      let pageTexts = [];
      try { pageTexts = await extractPdfPageTexts(req.file.buffer); }
      catch (e) { return res.status(400).json({ error: 'Could not read master PDF text: ' + e.message }); }

      const results = [];
      const usedEntities = new Set();
      let firstMatchedEid = null;

      for (let i = 0; i < pageCount; i++) {
        const text = pageTexts[i] || '';
        const def = execSummaries.matchPageToSummary(text);
        if (!def) { results.push({ page: i + 1, matched: false, reason: 'no entity title matched' }); continue; }
        const eid = def.match.entityId;
        if (!eid) { results.push({ page: i + 1, matched: false, key: def.key, reason: 'no entityId configured' }); continue; }
        if (usedEntities.has(eid)) { results.push({ page: i + 1, matched: false, key: def.key, reason: 'entity already matched on an earlier page' }); continue; }

        // Extract this single page into its own PDF.
        const single = await PDFDocument.create();
        const [pg] = await single.copyPages(masterDoc, [i]);
        single.addPage(pg);
        const singleBytes = Buffer.from(await single.save());

        try {
          const saved = writeStoredExecSummary(eid, singleBytes, who);
          usedEntities.add(eid);
          if (firstMatchedEid == null) firstMatchedEid = eid;
          const ent = db.prepare('SELECT name, code FROM entities WHERE id=?').get(eid);
          results.push({ page: i + 1, matched: true, key: def.key, entity_id: eid, entity_name: ent ? ent.name : null, entity_code: ent ? ent.code : null, file_id: saved.id });
        } catch (e) {
          results.push({ page: i + 1, matched: false, key: def.key, entity_id: eid, reason: 'save failed: ' + e.message });
        }
      }

      // Archive the master in an Admin-only folder under a home entity.
      const homeEid = req.body.master_home_eid || firstMatchedEid;
      let archived = null;
      if (homeEid) {
        try {
          ensureWpFolders(db, homeEid, execSummaries.ADMIN_MASTER_FOLDER, who);
          const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
          const masterName = 'Executive Summaries master (' + stamp + ').pdf';
          const a = saveWpBuffer(db, WORKPAPERS_DIR, homeEid, execSummaries.ADMIN_MASTER_FOLDER, masterName, 'application/pdf', req.file.buffer, who, { overwrite: false });
          archived = { entity_id: homeEid, folder: execSummaries.ADMIN_MASTER_FOLDER, file_id: a.id, original_name: a.original_name };
        } catch (e) { archived = { error: 'archive failed: ' + e.message }; }
      }

      const matchedCount = results.filter(r => r.matched).length;
      res.json({ pages: pageCount, matched: matchedCount, unmatched: pageCount - matchedCount, results, archived });
    } catch (e) {
      res.status(500).json({ error: 'Split error: ' + e.message });
    }
  });
});

// Fund financial-statement package (CLRF-style limited-partnership fund). Builds
// the five-statement PDF from the GL plus the fund_investments config and the
// GP/LP partner_type tags on investor classes. Admin/Accountant.
app.get('/api/entities/:eid/fund-statements.pdf', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
  try {
    const eid = req.params.eid;
    const asOf = (req.query && req.query.as_of);
    if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
    const ent = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
    const entityName = ent ? ent.name : ('Entity ' + eid);

    const investments = db.prepare(`SELECT id, parent_name, name, acquisition_date, cost, fair_value, sort_order
      FROM fund_investments WHERE entity_id = ? ORDER BY sort_order, id`).all(eid);
    const partnerClasses = db.prepare(`SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?`).all(eid);
    const commitments = db.prepare(`SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?`).all(eid);

    const getBalances = (o) => Promise.resolve(computeBalances(eid, o));
    const model = await financials.buildFundStatements({ asOf, entityName, getBalances, investments, partnerClasses, commitments });
    const bytes = await financials.renderFundStatementsPdf(model, []);

    const mm = asOf.slice(5, 7), yyyy = asOf.slice(0, 4);
    const safeName = entityName.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    const fname = safeName + '_Fund_Financial_Statements_' + mm + '_' + yyyy + '.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
    res.setHeader('X-Fund-Tie', JSON.stringify(model._tie).replace(/[\r\n]/g, ' '));
    res.send(Buffer.from(bytes));
  } catch (e) {
    res.status(500).json({ error: 'Fund statements error: ' + e.message });
  }
});


// ── Database backup (online snapshot) ───────────────────────────────────────
// GET /api/admin/backup — returns a consistent copy of the entire SQLite
// database. db.backup() takes a live-safe snapshot even while writes are in
// flight (no downtime, no locking). Auth accepts either the BACKUP_TOKEN env
// var (Authorization: Bearer <token> or ?token=) for unattended automation, or
// a normal Admin JWT for a browser-initiated download.
const BACKUP_TOKEN = process.env.BACKUP_TOKEN || '';
function backupAuth(req, res, next) {
  const hdr = req.headers['authorization'] || '';
  const token = hdr.startsWith('Bearer ') ? hdr.slice(7) : ((req.query && req.query.token) || '');
  if (BACKUP_TOKEN && token) {
    const a = Buffer.from(String(token));
    const b = Buffer.from(BACKUP_TOKEN);
    if (a.length === b.length && cryptoMod.timingSafeEqual(a, b)) return next();
  }
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    if (payload.role !== 'Admin') return res.status(403).json({ error: 'Forbidden' });
    req.user = payload;
    return next();
  } catch (e) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
}
app.get('/api/admin/backup', backupAuth, async (req, res) => {
  const stamp = new Date().toISOString().slice(0, 10);
  // Write the snapshot to the OS temp dir (large ephemeral overlay), NOT the
  // data volume — the volume only has room for the live DB, and a same-volume
  // copy would fail with SQLITE_FULL.
  const tmp = path.join(require('os').tmpdir(), '_cl_backup_tmp_' + Date.now() + '.db');
  try {
    await db.backup(tmp);
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', 'attachment; filename="cloudledger-' + stamp + '.db"');
    const stream = fs.createReadStream(tmp);
    stream.pipe(res);
    const cleanup = () => { try { fs.unlinkSync(tmp); } catch (e) {} };
    stream.on('close', cleanup);
    stream.on('error', () => { cleanup(); if (!res.headersSent) res.status(500).end(); });
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) {}
    console.error('[backup] failed:', e);
    if (!res.headersSent) res.status(500).json({ error: 'Backup failed: ' + e.message });
  }
});

// Any unmatched /api/* request (any method) returns a JSON 404 rather than
// falling through to the SPA shell (GET) or Express's default HTML 404 (POST/PUT/
// DELETE) — both of which make the client's res.json() throw "Unexpected token '<'".
app.use('/api', (req, res) => res.status(404).json({ error: 'Not found: ' + req.method + ' ' + req.originalUrl }));

if (process.env.NODE_ENV === 'production') app.get('*', (req, res) => {
  // Never serve the SPA shell for an unmatched API route — returning index.html
  // with a 200 makes the client's res.json() throw "Unexpected token '<'". An
  // unknown/API path must return a JSON 404 so the caller sees a real error.
  if (req.path.startsWith('/api/')) return res.status(404).json({ error: 'Not found: ' + req.method + ' ' + req.path });
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.sendFile(path.join(__dirname, '..', 'client', 'dist', 'index.html'));
});

// Global error handler (must be LAST). Without it, any error thrown in middleware
// or a route — e.g. multer rejecting an upload (file too large / bad multipart),
// or a synchronous throw — falls through to Express's DEFAULT handler, which
// renders an HTML error page. For an /api/* request that HTML reaches the client's
// res.json() and throws "Unexpected token '<', '<!DOCTYPE ...". Return JSON instead.
app.use((err, req, res, next) => {
  if (res.headersSent) return next(err);
  console.error('[error handler]', req.method, req.originalUrl, '-', err && err.message ? err.message : err);
  // Map common upload errors to sensible statuses.
  let status = err && err.status ? err.status : 500;
  let message = (err && err.message) ? err.message : 'Internal server error';
  if (err && err.code === 'LIMIT_FILE_SIZE') { status = 413; message = 'File is too large. Please upload a smaller file.'; }
  else if (err && err.name === 'MulterError') { status = 400; message = 'Upload error: ' + message; }
  if (req.path && req.path.startsWith('/api/')) return res.status(status).json({ error: message });
  // Non-API request: keep it simple, plain text (never the SPA shell for an error).
  res.status(status).type('text/plain').send(message);
});

app.listen(PORT, '0.0.0.0', () => console.log(`CloudLedger on port ${PORT}`));
