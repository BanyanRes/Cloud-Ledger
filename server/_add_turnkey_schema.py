import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/index.js'
src = open(p, encoding='utf-8').read()

# Insert new schemas right after billcom_sync_log index
old = "  CREATE INDEX IF NOT EXISTS idx_bsl_billcom_id ON billcom_sync_log(billcom_id);"
new = old + """
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
  CREATE INDEX IF NOT EXISTS idx_tsl_turnkey_id ON turnkey_sync_log(sync_type, turnkey_id);"""

if old not in src:
    print('MISSING anchor'); sys.exit(1)
src = src.replace(old, new, 1)
open(p, 'w', encoding='utf-8').write(src)
print('schema OK')
