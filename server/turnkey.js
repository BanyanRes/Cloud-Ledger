// Turnkey Rail integration for CloudLedger
// Mirrors the Bill.com integration pattern (config + map + sync_log).
// Auth: API keys (system-to-system), not JWTs. All sync events idempotent.

const crypto = require('crypto');
const bcrypt = require('bcryptjs');

// POC standard chart of accounts (5-digit codes; names can be renamed freely)
// POC chart of accounts — picks 5-digit codes that don't collide with the
// CloudLedger DEFAULT_COA (which already takes 10000, 10100, 11000, 12000,
// 13000, 15000, 20000, 21000, 40000, 50000, etc.).
const POC_ACCOUNTS = [
  { code: '10000', name: 'Cash',                                type: 'Asset' },     // shared w/ default
  { code: '10500', name: 'Bill.com Clearing',                   type: 'Asset' },     // new (avoids 10100)
  { code: '11500', name: 'Accounts Receivable - Owner',         type: 'Asset' },     // new (avoids 11000)
  { code: '14500', name: 'Costs in Excess of Billings',         type: 'Asset' },     // new (avoids 12500)
  { code: '15500', name: 'Construction-in-Progress',            type: 'Asset' },     // new (avoids 13000)
  { code: '20500', name: 'Accounts Payable - Subcontractors',   type: 'Liability' }, // new (avoids 20000)
  { code: '23000', name: 'Billings on Uncompleted Contracts',   type: 'Liability' }, // new (free slot)
  { code: '24000', name: 'Billings in Excess of Costs',         type: 'Liability' }, // new (free slot)
  { code: '45000', name: 'Construction Revenue',                type: 'Revenue' },   // new (avoids 40000)
  { code: '55000', name: 'Cost of Construction',                type: 'Expense' },   // new (avoids 50000)
];

// API key helpers
function generateApiKey() {
  return 'tkr_' + crypto.randomBytes(16).toString('hex');
}
function hashApiKey(rawKey) {
  return bcrypt.hashSync(rawKey, 10);
}
function compareApiKey(rawKey, hash) {
  try { return bcrypt.compareSync(rawKey, hash); } catch (e) { return false; }
}
function apiKeyPrefix(rawKey) {
  return rawKey.slice(0, 12);
}

// Auth middleware: validates an API key + attaches scopes to req.apiKey
function apiKeyAuth(db) {
  return function (req, res, next) {
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (!token || !token.startsWith('tkr_')) {
      return res.status(401).json({ error: 'API key required' });
    }
    const keys = db.prepare("SELECT * FROM api_keys WHERE revoked_at IS NULL").all();
    const match = keys.find(function (k) { return compareApiKey(token, k.key_hash); });
    if (!match) return res.status(401).json({ error: 'Invalid API key' });
    db.prepare('UPDATE api_keys SET last_used_at = ? WHERE id = ?')
      .run(new Date().toISOString(), match.id);
    req.apiKey = { id: match.id, name: match.name, scopes: match.scopes.split(',').filter(Boolean) };
    next();
  };
}

function requireScope(scope) {
  return function (req, res, next) {
    if (!req.apiKey) return res.status(401).json({ error: 'API key required' });
    if (!req.apiKey.scopes.includes(scope) && !req.apiKey.scopes.includes('*')) {
      return res.status(403).json({ error: 'Scope required: ' + scope });
    }
    next();
  };
}

// Resolve the company entity (the single "Turnkey Rail" entity that holds
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
      pick('10000'), pick('10500'), pick('11500'),
      pick('14500'), pick('15500'), pick('20500'),
      pick('23000'), pick('24000'),
      pick('45000'), pick('55000'),
      project_code, project_name, contract_amount, total_estimated_costs,
      now
    );
  }

  return db.prepare(
    'SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?'
  ).get(turnkey_project_id);
}

// Idempotency + logging
function findExistingSync(db, args) {
  return db.prepare(
    "SELECT cl_entry_id FROM turnkey_sync_log " +
    "WHERE sync_type = ? AND turnkey_id = ? AND status = 'success' AND cl_entry_id IS NOT NULL " +
    "ORDER BY id DESC LIMIT 1"
  ).get(args.sync_type, String(args.turnkey_id));
}

function logSync(db, args) {
  db.prepare(
    'INSERT INTO turnkey_sync_log ' +
    '(cl_entity_id, sync_type, turnkey_id, cl_entry_id, status, message, payload_json, created_at) ' +
    'VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
  ).run(
    args.cl_entity_id, args.sync_type,
    args.turnkey_id != null ? String(args.turnkey_id) : null,
    args.cl_entry_id || null, args.status, args.message || null,
    args.payload ? JSON.stringify(args.payload) : null,
    new Date().toISOString()
  );
}

// Journal entry helper (balanced).
// Schema: journal_entries(entity_id, entry_num, date, memo, created_by, created_at)
//         journal_lines(entry_id, account_code, debit, credit)
// memo is NOT NULL; entry_num auto-incremented per entity.
function postJE(db, args) {
  const lines = args.lines;
  let totalDr = 0, totalCr = 0;
  for (var i = 0; i < lines.length; i++) {
    totalDr += Number(lines[i].debit || 0);
    totalCr += Number(lines[i].credit || 0);
  }
  if (Math.abs(totalDr - totalCr) > 0.005) {
    throw new Error('Unbalanced JE: Dr ' + totalDr + ' vs Cr ' + totalCr);
  }
  // Next entry_num for this entity
  const nextRow = db.prepare(
    'SELECT COALESCE(MAX(entry_num), 0) + 1 AS n FROM journal_entries WHERE entity_id = ?'
  ).get(args.cl_entity_id);
  const entryNum = nextRow.n;

  // Build memo: prefix with reference if provided
  const memo = (args.reference ? '[' + args.reference + '] ' : '') + (args.memo || '');

  const now = new Date().toISOString();
  const result = db.prepare(
    'INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by, created_at) ' +
    'VALUES (?, ?, ?, ?, ?, ?)'
  ).run(
    args.cl_entity_id, entryNum, args.date, memo,
    args.created_by || 'turnkey-sync', now
  );
  const entryId = result.lastInsertRowid;
  const insertLine = db.prepare(
    'INSERT INTO journal_lines (entry_id, account_code, debit, credit, project_id) ' +
    'VALUES (?, ?, ?, ?, ?)'
  );
  for (var j = 0; j < lines.length; j++) {
    const l = lines[j];
    insertLine.run(
      entryId, l.account_code,
      Number(l.debit || 0), Number(l.credit || 0),
      l.project_id != null ? String(l.project_id) : null
    );
  }
  return entryId;
}

// ===== Sync event handlers =====

// Event 1: Sub pay app approved -> Dr. CIP / Cr. AP-Sub
function syncSubPayAppApproved(db, payload) {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(payload.turnkey_project_id);
  if (!map) throw new Error('Project ' + payload.turnkey_project_id + ' not linked');

  const existing = findExistingSync(db, { sync_type: 'sub_payapp_approved', turnkey_id: payload.payapp_id });
  if (existing) return { cl_entry_id: existing.cl_entry_id, idempotent: true };

  const entryId = postJE(db, {
    cl_entity_id: map.cl_entity_id,
    date: payload.date,
    memo: 'Sub pay app approved - ' + (payload.vendor_name || ''),
    reference: 'TKR-PA-' + payload.payapp_id,
    lines: [
      { account_code: map.cip_code, debit: payload.amount, credit: 0, description: 'CIP accrual - ' + payload.vendor_name, project_id: payload.turnkey_project_id },
      { account_code: map.ap_sub_code, debit: 0, credit: payload.amount, description: 'AP - ' + payload.vendor_name, project_id: payload.turnkey_project_id },
    ],
  });

  logSync(db, {
    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_approved',
    turnkey_id: payload.payapp_id, cl_entry_id: entryId, status: 'success', payload: payload,
  });
  return { cl_entry_id: entryId, idempotent: false };
}

// Event 2: Sub pay app paid (wire/ACH/check or bill_com) -> Dr. AP-Sub / Cr. Cash or Clearing
function syncSubPayAppPaid(db, payload) {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(payload.turnkey_project_id);
  if (!map) throw new Error('Project ' + payload.turnkey_project_id + ' not linked');

  const existing = findExistingSync(db, { sync_type: 'sub_payapp_paid', turnkey_id: payload.payapp_id });
  if (existing) return { cl_entry_id: existing.cl_entry_id, idempotent: true };

  const creditCode = payload.payment_method === 'bill_com' ? map.billcom_clearing_code : map.cash_account_code;

  const entryId = postJE(db, {
    cl_entity_id: map.cl_entity_id,
    date: payload.date,
    memo: 'Sub pay app paid - ' + (payload.vendor_name || '') + ' (' + payload.payment_method + ')',
    reference: 'TKR-PMT-' + payload.payapp_id,
    lines: [
      { account_code: map.ap_sub_code, debit: payload.amount, credit: 0, description: 'AP clear - ' + payload.vendor_name, project_id: payload.turnkey_project_id },
      { account_code: creditCode, debit: 0, credit: payload.amount, description: 'Cash/Clearing out - ' + payload.payment_method, project_id: payload.turnkey_project_id },
    ],
  });

  logSync(db, {
    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_paid',
    turnkey_id: payload.payapp_id, cl_entry_id: entryId, status: 'success', payload: payload,
  });
  return { cl_entry_id: entryId, idempotent: false };
}

// Event 3: Owner pay app issued -> Dr. AR / Cr. Billings on Uncompleted
function syncOwnerPayAppIssued(db, payload) {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(payload.turnkey_project_id);
  if (!map) throw new Error('Project ' + payload.turnkey_project_id + ' not linked');

  const existing = findExistingSync(db, { sync_type: 'owner_payapp_issued', turnkey_id: payload.payapp_id });
  if (existing) return { cl_entry_id: existing.cl_entry_id, idempotent: true };

  const entryId = postJE(db, {
    cl_entity_id: map.cl_entity_id,
    date: payload.date,
    memo: 'Owner pay app issued',
    reference: 'TKR-OWPA-' + payload.payapp_id,
    lines: [
      { account_code: map.ar_owner_code, debit: payload.amount, credit: 0, description: 'AR - Owner', project_id: payload.turnkey_project_id },
      { account_code: map.billings_uncompleted_code, debit: 0, credit: payload.amount, description: 'Billings on uncompleted contract', project_id: payload.turnkey_project_id },
    ],
  });

  logSync(db, {
    cl_entity_id: map.cl_entity_id, sync_type: 'owner_payapp_issued',
    turnkey_id: payload.payapp_id, cl_entry_id: entryId, status: 'success', payload: payload,
  });
  return { cl_entry_id: entryId, idempotent: false };
}

// Event 4: Owner payment received -> Dr. Cash / Cr. AR
function syncOwnerPaymentReceived(db, payload) {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(payload.turnkey_project_id);
  if (!map) throw new Error('Project ' + payload.turnkey_project_id + ' not linked');

  const existing = findExistingSync(db, { sync_type: 'owner_payment_received', turnkey_id: payload.payapp_id });
  if (existing) return { cl_entry_id: existing.cl_entry_id, idempotent: true };

  const entryId = postJE(db, {
    cl_entity_id: map.cl_entity_id,
    date: payload.date,
    memo: 'Owner payment received',
    reference: 'TKR-RCT-' + payload.payapp_id,
    lines: [
      { account_code: map.cash_account_code, debit: payload.amount, credit: 0, description: 'Cash from owner', project_id: payload.turnkey_project_id },
      { account_code: map.ar_owner_code, debit: 0, credit: payload.amount, description: 'AR clear', project_id: payload.turnkey_project_id },
    ],
  });

  logSync(db, {
    cl_entity_id: map.cl_entity_id, sync_type: 'owner_payment_received',
    turnkey_id: payload.payapp_id, cl_entry_id: entryId, status: 'success', payload: payload,
  });
  return { cl_entry_id: entryId, idempotent: false };
}

// Event 5: Month-end POC adjustment (cost-to-cost)
//   pct = CIP_balance / total_estimated_costs
//   earned_revenue = pct * contract_amount
//   recognized_cost = CIP_balance
//   Reclass billings -> revenue; post over/under-billing as needed
function syncMonthEndPOC(db, payload) {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(payload.turnkey_project_id);
  if (!map) throw new Error('Project ' + payload.turnkey_project_id + ' not linked');

  const turnkey_id = payload.turnkey_project_id + ':' + payload.period_end_date;
  const existing = findExistingSync(db, { sync_type: 'month_end_poc', turnkey_id: turnkey_id });
  if (existing) return { cl_entry_id: existing.cl_entry_id, idempotent: true };

  function getBal(code) {
    const row = db.prepare(
      'SELECT COALESCE(SUM(l.debit), 0) - COALESCE(SUM(l.credit), 0) AS bal ' +
      'FROM journal_lines l ' +
      'JOIN journal_entries e ON e.id = l.entry_id ' +
      'WHERE e.entity_id = ? AND l.account_code = ? AND l.project_id = ? AND e.date <= ?'
    ).get(map.cl_entity_id, code, String(payload.turnkey_project_id), payload.period_end_date);
    return Number(row.bal || 0);
  }
  const cipBalance = getBal(map.cip_code);
  const billingsBalance = -getBal(map.billings_uncompleted_code);

  const totalEst = payload.total_estimated_costs;
  const pct = totalEst > 0 ? Math.min(cipBalance / totalEst, 1) : 0;
  const earnedRevenue = Math.round(pct * payload.contract_amount * 100) / 100;
  const recognizedCost = cipBalance;

  const pid = payload.turnkey_project_id;
  const lines = [];
  if (recognizedCost > 0.005) {
    lines.push({ account_code: map.cost_of_construction_code, debit: recognizedCost, credit: 0, description: 'POC cost recognition', project_id: pid });
    lines.push({ account_code: map.cip_code, debit: 0, credit: recognizedCost, description: 'CIP relief', project_id: pid });
  }
  if (earnedRevenue > 0.005) {
    lines.push({ account_code: map.billings_uncompleted_code, debit: earnedRevenue, credit: 0, description: 'Reclass billings to revenue', project_id: pid });
    lines.push({ account_code: map.revenue_code, debit: 0, credit: earnedRevenue, description: 'POC revenue recognition', project_id: pid });
  }
  const diff = Math.round((earnedRevenue - billingsBalance) * 100) / 100;
  if (Math.abs(diff) > 0.005) {
    if (diff > 0) {
      lines.push({ account_code: map.costs_in_excess_code, debit: diff, credit: 0, description: 'Under-billing (costs in excess)', project_id: pid });
      lines.push({ account_code: map.billings_uncompleted_code, debit: 0, credit: diff, description: 'Under-billing offset', project_id: pid });
    } else {
      lines.push({ account_code: map.billings_uncompleted_code, debit: -diff, credit: 0, description: 'Over-billing offset', project_id: pid });
      lines.push({ account_code: map.billings_in_excess_code, debit: 0, credit: -diff, description: 'Over-billing (billings in excess)', project_id: pid });
    }
  }

  if (lines.length === 0) {
    logSync(db, {
      cl_entity_id: map.cl_entity_id, sync_type: 'month_end_poc',
      turnkey_id: turnkey_id, cl_entry_id: null, status: 'skipped',
      message: 'No POC activity this period', payload: payload,
    });
    return { cl_entry_id: null, idempotent: false, skipped: true };
  }

  const entryId = postJE(db, {
    cl_entity_id: map.cl_entity_id,
    date: payload.period_end_date,
    memo: 'Month-end POC adjustment (' + (pct*100).toFixed(1) + '% complete)',
    reference: 'TKR-POC-' + payload.period_end_date,
    lines: lines,
  });

  logSync(db, {
    cl_entity_id: map.cl_entity_id, sync_type: 'month_end_poc',
    turnkey_id: turnkey_id, cl_entry_id: entryId, status: 'success',
    message: 'pct=' + (pct*100).toFixed(2) + '% earned=' + earnedRevenue + ' cost=' + recognizedCost + ' diff=' + diff,
    payload: payload,
  });
  return { cl_entry_id: entryId, idempotent: false, pct: pct, earnedRevenue: earnedRevenue, recognizedCost: recognizedCost, diff: diff };
}

module.exports = {
  POC_ACCOUNTS: POC_ACCOUNTS,
  generateApiKey: generateApiKey,
  hashApiKey: hashApiKey,
  compareApiKey: compareApiKey,
  apiKeyPrefix: apiKeyPrefix,
  apiKeyAuth: apiKeyAuth,
  requireScope: requireScope,
  linkProject: linkProject,
  syncSubPayAppApproved: syncSubPayAppApproved,
  syncSubPayAppPaid: syncSubPayAppPaid,
  syncOwnerPayAppIssued: syncOwnerPayAppIssued,
  syncOwnerPaymentReceived: syncOwnerPaymentReceived,
  syncMonthEndPOC: syncMonthEndPOC,
};

// =====================================================================
// WIP Schedule (Job Schedule) generator
// ---------------------------------------------------------------------
// For each linked Turnkey project, computes the standard WIP row:
//   contract_amount, approved_co (currently rolled into contract),
//   revised_contract, costs_to_date (CIP debit by project_id),
//   estimated_cost_to_complete (= total_est - costs_to_date),
//   estimated_total_cost (= total_estimated_costs),
//   estimated_gross_profit, percent_complete (cost-to-cost),
//   earned_revenue, billed_to_date (billings + AR cleared),
//   over_under_billing.
// Reconciles to the GL: sum of (CIP+COC) per project = costs_to_date.
// =====================================================================

function computeWipRow(db, map, asOfDate) {
  const eid = map.cl_entity_id;
  const pid = String(map.turnkey_project_id);

  // SQL helper: net balance of an account on this entity, scoped to project_id, as of date
  function bal(code) {
    if (!code) return 0;
    const row = db.prepare(
      'SELECT COALESCE(SUM(l.debit), 0) - COALESCE(SUM(l.credit), 0) AS bal ' +
      'FROM journal_lines l ' +
      'JOIN journal_entries e ON e.id = l.entry_id ' +
      'WHERE e.entity_id = ? AND l.account_code = ? AND l.project_id = ? AND e.date <= ?'
    ).get(eid, code, pid, asOfDate);
    return Number(row.bal || 0);
  }
  // For revenue/expense (income statement), we just take debits - credits (cumulative).
  function actDr(code) {
    if (!code) return 0;
    const row = db.prepare(
      'SELECT COALESCE(SUM(l.debit), 0) AS dr, COALESCE(SUM(l.credit), 0) AS cr ' +
      'FROM journal_lines l JOIN journal_entries e ON e.id = l.entry_id ' +
      'WHERE e.entity_id = ? AND l.account_code = ? AND l.project_id = ? AND e.date <= ?'
    ).get(eid, code, pid, asOfDate);
    return { dr: Number(row.dr || 0), cr: Number(row.cr || 0) };
  }

  // Costs to date = CIP debits + COC debits (recognized cost already moved out
  // of CIP at month-end stays counted as cost incurred on the job)
  const cipAct = actDr(map.cip_code);
  const cocAct = actDr(map.cost_of_construction_code);
  const costsToDate = cipAct.dr + cocAct.dr - cipAct.cr; // cipAct.cr is the CIP->COC relief

  const contract = Number(map.contract_amount || 0);
  const totalEst = Number(map.total_estimated_costs || 0);
  // Estimated cost to complete: simple model = totalEst - costsToDate (floor 0)
  const estCostToComplete = Math.max(totalEst - costsToDate, 0);
  const estTotalCost = costsToDate + estCostToComplete;
  const estGrossProfit = contract - estTotalCost;
  const pctComplete = estTotalCost > 0 ? Math.min(costsToDate / estTotalCost, 1) : 0;
  const earnedRevenue = Math.round(pctComplete * contract * 100) / 100;

  // Billed to date = sum of credits on Billings on Uncompleted AR cycle = cumulative gross billings
  // We sum AR debits (every owner pay app issued posted Dr.AR)
  const arAct = actDr(map.ar_owner_code);
  const billedToDate = arAct.dr; // every issued owner pay app

  const overUnder = Math.round((billedToDate - earnedRevenue) * 100) / 100;

  return {
    turnkey_project_id: map.turnkey_project_id,
    project_code: map.project_code,
    project_name: map.project_name,
    contract_amount: contract,
    revised_contract: contract, // CO uplifts already baked into contract_amount at sync time
    costs_to_date: Math.round(costsToDate * 100) / 100,
    estimated_cost_to_complete: Math.round(estCostToComplete * 100) / 100,
    estimated_total_cost: Math.round(estTotalCost * 100) / 100,
    estimated_gross_profit: Math.round(estGrossProfit * 100) / 100,
    percent_complete: Math.round(pctComplete * 10000) / 100, // as %, 2 decimals
    earned_revenue: earnedRevenue,
    billed_to_date: Math.round(billedToDate * 100) / 100,
    over_under_billing: overUnder,
    over_under_label: overUnder >= 0 ? 'over' : 'under',
  };
}

function computeWipSchedule(db, asOfDate) {
  const maps = db.prepare(
    'SELECT * FROM turnkey_project_map ORDER BY project_code'
  ).all();
  const rows = maps.map(m => computeWipRow(db, m, asOfDate));
  const total = rows.reduce(function (acc, r) {
    acc.contract_amount += r.contract_amount;
    acc.revised_contract += r.revised_contract;
    acc.costs_to_date += r.costs_to_date;
    acc.estimated_cost_to_complete += r.estimated_cost_to_complete;
    acc.estimated_total_cost += r.estimated_total_cost;
    acc.estimated_gross_profit += r.estimated_gross_profit;
    acc.earned_revenue += r.earned_revenue;
    acc.billed_to_date += r.billed_to_date;
    acc.over_under_billing += r.over_under_billing;
    return acc;
  }, {
    contract_amount: 0, revised_contract: 0, costs_to_date: 0,
    estimated_cost_to_complete: 0, estimated_total_cost: 0,
    estimated_gross_profit: 0, earned_revenue: 0, billed_to_date: 0,
    over_under_billing: 0,
  });
  return { as_of_date: asOfDate, rows: rows, total: total };
}

module.exports.getCompanyEntityId = getCompanyEntityId;
module.exports.seedPOCAccountsIfMissing = seedPOCAccountsIfMissing;
module.exports.computeWipRow = computeWipRow;
module.exports.computeWipSchedule = computeWipSchedule;
