import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()

# 1. Replace linkProject — no entity creation, just record mapping + seed COA ONCE on the company entity if needed.
old_link = """// Project setup: link Turnkey project to a CL entity, seed POC accounts
function linkProject(db, params) {
  const turnkey_project_id = params.turnkey_project_id;
  const project_code = params.project_code;
  const project_name = params.project_name;
  const existing = db.prepare(
    'SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?'
  ).get(turnkey_project_id);
  if (existing) return existing;

  const now = new Date().toISOString();
  const entityResult = db.prepare(
    'INSERT INTO entities (name, code, created_at) VALUES (?, ?, ?)'
  ).run(project_name, project_code, now);
  const cl_entity_id = entityResult.lastInsertRowid;

  const insertAcct = db.prepare(
    'INSERT INTO accounts (entity_id, code, name, type) VALUES (?, ?, ?, ?)'
  );
  for (var i = 0; i < POC_ACCOUNTS.length; i++) {
    const a = POC_ACCOUNTS[i];
    insertAcct.run(cl_entity_id, a.code, a.name, a.type);
  }

  const codeMap = {};
  POC_ACCOUNTS.forEach(function (a) { codeMap[a.name] = a.code; });

  db.prepare(
    'INSERT INTO turnkey_project_map (' +
    'turnkey_project_id, cl_entity_id,' +
    'cash_account_code, billcom_clearing_code, ar_owner_code,' +
    'costs_in_excess_code, cip_code, ap_sub_code,' +
    'billings_uncompleted_code, billings_in_excess_code,' +
    'revenue_code, cost_of_construction_code, created_at' +
    ') VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).run(
    turnkey_project_id, cl_entity_id,
    codeMap['Cash'],
    codeMap['Bill.com Clearing'],
    codeMap['Accounts Receivable - Owner'],
    codeMap['Costs in Excess of Billings'],
    codeMap['Construction-in-Progress'],
    codeMap['Accounts Payable - Subcontractors'],
    codeMap['Billings on Uncompleted Contracts'],
    codeMap['Billings in Excess of Costs'],
    codeMap['Construction Revenue'],
    codeMap['Cost of Construction'],
    now
  );
  return db.prepare(
    'SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?'
  ).get(turnkey_project_id);
}"""

new_link = """// Resolve the company entity (the single "Turnkey Rail" entity that holds
// all project activity). Reads turnkey_config.default_entity_id; returns null
// if not set (caller must validate).
function getCompanyEntityId(db) {
  const row = db.prepare('SELECT default_entity_id FROM turnkey_config WHERE id = 1').get();
  return row ? row.default_entity_id : null;
}

// Seed the standard 5-digit POC chart of accounts on the company entity if any
// are missing. Idempotent: existing codes are left alone (so names you renamed
// stay renamed).
function seedPOCAccountsIfMissing(db, cl_entity_id) {
  const existing = new Set(
    db.prepare('SELECT code FROM accounts WHERE entity_id = ?').all(cl_entity_id).map(r => r.code)
  );
  const ins = db.prepare(
    'INSERT INTO accounts (entity_id, code, name, type) VALUES (?, ?, ?, ?)'
  );
  let added = 0;
  for (const a of POC_ACCOUNTS) {
    if (!existing.has(a.code)) {
      ins.run(cl_entity_id, a.code, a.name, a.type);
      added++;
    }
  }
  return added;
}

// Project setup: register a Turnkey project as a job dimension on the company
// entity. Does NOT create a new CL entity. Also seeds POC accounts on the
// company entity (idempotent) so the first project setup brings the standard
// COA in if not already present.
function linkProject(db, params) {
  const turnkey_project_id = params.turnkey_project_id;
  const project_code = params.project_code;
  const project_name = params.project_name;
  const contract_amount = params.contract_amount != null ? Number(params.contract_amount) : null;
  const total_estimated_costs = params.total_estimated_costs != null ? Number(params.total_estimated_costs) : null;

  const cl_entity_id = getCompanyEntityId(db);
  if (!cl_entity_id) {
    throw new Error('Company entity not configured. POST /api/turnkey/config first with default_entity_id.');
  }

  seedPOCAccountsIfMissing(db, cl_entity_id);

  const existing = db.prepare(
    'SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?'
  ).get(turnkey_project_id);

  const now = new Date().toISOString();
  if (existing) {
    db.prepare(
      'UPDATE turnkey_project_map SET ' +
      'project_code = COALESCE(?, project_code), ' +
      'project_name = COALESCE(?, project_name), ' +
      'contract_amount = COALESCE(?, contract_amount), ' +
      'total_estimated_costs = COALESCE(?, total_estimated_costs) ' +
      'WHERE turnkey_project_id = ?'
    ).run(project_code, project_name, contract_amount, total_estimated_costs, turnkey_project_id);
  } else {
    // Resolve POC account codes from the COA (so renamed names don't matter)
    const acctsByCode = {};
    db.prepare('SELECT code FROM accounts WHERE entity_id = ?').all(cl_entity_id)
      .forEach(r => { acctsByCode[r.code] = true; });
    const pick = (code) => acctsByCode[code] ? code : null;

    db.prepare(
      'INSERT INTO turnkey_project_map (' +
      'turnkey_project_id, cl_entity_id,' +
      'cash_account_code, billcom_clearing_code, ar_owner_code,' +
      'costs_in_excess_code, cip_code, ap_sub_code,' +
      'billings_uncompleted_code, billings_in_excess_code,' +
      'revenue_code, cost_of_construction_code,' +
      'project_code, project_name, contract_amount, total_estimated_costs,' +
      'created_at' +
      ') VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
    ).run(
      turnkey_project_id, cl_entity_id,
      pick('10000'), pick('10100'), pick('12000'),
      pick('12500'), pick('13000'), pick('21000'),
      pick('23000'), pick('24000'),
      pick('40000'), pick('50000'),
      project_code, project_name, contract_amount, total_estimated_costs,
      now
    );
  }

  return db.prepare(
    'SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?'
  ).get(turnkey_project_id);
}"""

assert old_link in src, 'old linkProject pattern not found'
src = src.replace(old_link, new_link, 1)

open(p, 'w', encoding='utf-8').write(src)
print('OK linkProject replaced')
