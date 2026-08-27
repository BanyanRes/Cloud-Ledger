import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/index.js'
src = open(p, encoding='utf-8').read()

# Insert new migrations after the billcom_config default_cash_account block
old = "if (!bcCfgCols.includes('default_cash_account')) db.exec(\"ALTER TABLE billcom_config ADD COLUMN default_cash_account TEXT\");"
new = old + """

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
"""

if old not in src:
    print('MISSING anchor'); sys.exit(1)
src = src.replace(old, new, 1)
open(p, 'w', encoding='utf-8').write(src)
print('OK')
