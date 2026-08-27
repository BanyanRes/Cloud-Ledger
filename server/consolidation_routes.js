// ═══════════════════════════════════════════════════════════════════════════
// Consolidation — trial-balance parsing, the Braker seed, and the HTTP surface.
//
// The engine (schema, mapping roll-up, eliminations, consolidating columns)
// lives in server/consolidation.js. This file is the edge: reading a property
// manager's spreadsheet, standing Braker up the first time the server boots
// with this code, and the routes the UI calls.
//
// SCOPE: Braker and HP only. Every route resolves its group through
// scopedGroup(), which refuses any parent entity outside those two.
// ═══════════════════════════════════════════════════════════════════════════

const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const C = require('./consolidation');
const { r2, monthEnd, monthStart, yearStart } = C._helpers;
const isDrType = t => t === 'Asset' || t === 'Expense';

// ════════════════════════════ TB parsing ════════════════════════════

// Property-management exports are not clean tables. The Foxtail file carries
// four preamble rows, a header split across two rows ("Balance / Forward",
// "Ending / balance"), and unlabelled code and name columns. So the sheet is
// read as an array of arrays and the header is located rather than assumed,
// with a positional fallback for exactly that layout.
//
// Returns lines in DEBIT-POSITIVE terms — assets and expenses positive,
// liabilities, equity and revenue negative — which is how a trial balance
// reads and what lets the stored rows be tied back to the file line by line.
function parseOperatingTb(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw Object.assign(new Error('The workbook has no sheets'), { status: 400 });
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
  if (!grid.length) throw Object.assign(new Error('No rows found in the file'), { status: 400 });

  const CODE_RX = /^[0-9]{3,}([-.][0-9A-Za-z]+)*$/;
  const num = (v) => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const s = String(v).trim();
    const paren = /^\(.*\)$/.test(s);
    const n = parseFloat(s.replace(/[,$()\s]/g, ''));
    if (isNaN(n)) return 0;
    return paren ? -n : n;
  };

  // First data row: a row whose first code-shaped cell has a non-numeric label
  // beside it.
  let firstData = -1, codeIdx = -1, nameIdx = -1;
  for (let i = 0; i < grid.length && firstData < 0; i++) {
    const row = grid[i] || [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] != null && CODE_RX.test(String(row[c]).trim())) {
        const nm = row.slice(c + 1).findIndex(v => v != null && String(v).trim() !== '' && isNaN(Number(v)));
        if (nm >= 0) { firstData = i; codeIdx = c; nameIdx = c + 1 + nm; }
        break;
      }
    }
  }
  if (firstData < 0) {
    throw Object.assign(new Error('Could not find any account-code rows. Check that a column holds account numbers.'), { status: 400 });
  }

  // Header text per column: everything above the first data row joined, so a
  // header split across two rows reads as one string.
  const joined = [];
  for (let c = 0; c < 60; c++) {
    let s = '';
    for (let i = 0; i < firstData; i++) {
      const v = (grid[i] || [])[c];
      s += ' ' + (v == null ? '' : String(v));
    }
    joined[c] = s.toLowerCase();
  }
  const findIdx = (rx) => joined.findIndex((h, i) => i > nameIdx && rx.test(h));
  let fwdIdx = findIdx(/forward|beginning|opening|prior/);
  let drIdx = findIdx(/debit/);
  let crIdx = findIdx(/credit/);
  let endIdx = findIdx(/ending|closing|current\s*balance/);
  if (endIdx < 0) {
    const b = findIdx(/balance/);
    if (b >= 0 && b !== fwdIdx) endIdx = b;
  }

  // Positional fallback: code, name, forward, debit, credit, ending.
  if (fwdIdx < 0 && drIdx < 0 && crIdx < 0 && endIdx < 0) {
    const probe = grid[firstData] || [];
    const numeric = [];
    for (let c = nameIdx + 1; c < probe.length; c++) {
      if (probe[c] != null && String(probe[c]).trim() !== '' && !isNaN(Number(String(probe[c]).replace(/[,$()]/g, '')))) numeric.push(c);
    }
    if (numeric.length >= 4) { fwdIdx = numeric[0]; drIdx = numeric[1]; crIdx = numeric[2]; endIdx = numeric[3]; }
    else if (numeric.length >= 1) { endIdx = numeric[numeric.length - 1]; }
  }
  if (endIdx < 0 && (fwdIdx < 0 || drIdx < 0 || crIdx < 0)) {
    throw Object.assign(new Error('Could not find an ending-balance column, or a forward/debit/credit set to derive one from.'), { status: 400 });
  }

  const lines = [];
  for (let i = firstData; i < grid.length; i++) {
    const row = grid[i] || [];
    const code = row[codeIdx] == null ? '' : String(row[codeIdx]).trim();
    if (!code || !CODE_RX.test(code)) continue;              // subtotal and total rows
    const name = row[nameIdx] == null ? '' : String(row[nameIdx]).trim();
    if (/^total\b/i.test(name)) continue;
    const forward = fwdIdx >= 0 ? num(row[fwdIdx]) : 0;
    const debit = drIdx >= 0 ? num(row[drIdx]) : 0;
    const credit = crIdx >= 0 ? num(row[crIdx]) : 0;
    const ending = endIdx >= 0 ? num(row[endIdx]) : r2(forward + debit - credit);
    lines.push({ source_code: code, source_name: name, forward: r2(forward), debit: r2(debit), credit: r2(credit), ending: r2(ending) });
  }
  if (!lines.length) throw Object.assign(new Error('No account rows parsed. Check the account-code column.'), { status: 400 });

  // A trial balance sums to zero in debit-positive terms. A file on the
  // natural-side convention only foots once the credit-side accounts are
  // flipped — detected by trying it, not assumed.
  const leadType = (code) => {
    const n = parseInt(String(code).replace(/[^0-9]/g, '').slice(0, 1), 10);
    if (n === 1) return 'Asset';
    if (n === 2) return 'Liability';
    if (n === 3) return 'Equity';
    if (n === 4) return 'Revenue';
    return 'Expense';
  };
  const sumAsIs = r2(lines.reduce((s, l) => s + l.ending, 0));
  const sumFlip = r2(lines.reduce((s, l) => s + (isDrType(leadType(l.source_code)) ? l.ending : -l.ending), 0));
  let signMode = 'debit-positive';
  if (Math.abs(sumAsIs) > 0.02 && Math.abs(sumFlip) < 0.02) {
    signMode = 'natural';
    for (const l of lines) {
      if (!isDrType(leadType(l.source_code))) {
        l.forward = r2(-l.forward);
        l.ending = r2(-l.ending);
        const d = l.debit; l.debit = l.credit; l.credit = d;
      }
    }
  }
  const residual = r2(lines.reduce((s, l) => s + l.ending, 0));
  const activityResidual = r2(lines.reduce((s, l) => s + (l.debit - l.credit), 0));
  return {
    lines, signMode, residual, activityResidual,
    hasActivity: drIdx >= 0 && crIdx >= 0,
    hasForward: fwdIdx >= 0,
    balanced: Math.abs(residual) < 0.02,
  };
}

// ══════════════════════════════ Braker seed ══════════════════════════════

// Foxtail account -> statement line, verified against the July 2026 package:
// all 68 income-statement lines and every balance-sheet line agreed.
// Keyed on the FULL source code — 51030-000 Bonuses and 51030-001 Quarterly
// Bonuses are different statement lines, and the CPA package has them crossed.
// [ source code, target code, target name, target type ]
const BRAKER_MAP = [
  // Balance sheet
  ['11020-000', '10119', 'Cash - Operating I', 'Asset'],
  ['11086-000', '10903', 'Restricted Cash', 'Asset'],
  ['12010-000', '12000', 'Accounts Receivable', 'Asset'],
  ['13085-000', '13000', 'Prepaid Expense', 'Asset'],
  ['21010-000', '20000', 'Accounts Payable', 'Liability'],
  ['22005-000', '21008', 'Accrued Management Fees - Operating', 'Liability'],
  ['22010-000', '21009', 'Accrued Expenses - Operating', 'Liability'],
  ['22020-000', '21006', 'Accrued Property Taxes', 'Liability'],
  ['22025-000', '21004', 'Accrued Other', 'Liability'],
  ['23010-000', '21010', 'Deferred Revenue', 'Liability'],
  ['23030-000', '24000', 'Security Deposit', 'Liability'],
  ['32000-000', '32000', 'Contributed Capital - Braker Propco. LLC', 'Equity'],
  ['34010-000', '39000', 'Retained Earnings', 'Equity'],
  // Revenue and the contra-revenue block
  ['41000-000', '40001', 'Gross Potential Rent', 'Revenue'],
  ['41010-000', '40100', 'Loss/Gain to Lease', 'Revenue'],
  ['41100-000', '40200', 'Vacancy Loss', 'Revenue'],
  ['41110-000', '40201', 'Employee Units', 'Revenue'],
  ['41091-000', '40400', 'Rent Concession', 'Revenue'],
  ['41120-000', '40480', 'Model Apartment', 'Revenue'],
  ['43200-000', '41111', 'Non Refundable Pet Fees', 'Revenue'],
  ['43020-000', '41113', 'Fee - Application', 'Revenue'],
  ['43063-000', '41115', 'Fee - Community', 'Revenue'],
  ['43201-000', '41117', 'Income - Pet Rent', 'Revenue'],
  ['43257-000', '41119', 'Cable Rebill', 'Revenue'],
  ['43261-000', '41120', 'Pest Control Rebill', 'Revenue'],
  ['43262-000', '41121', 'Trash Rebill', 'Revenue'],
  ['43263-000', '41122', 'Trash Remove Door To Door', 'Revenue'],
  ['43264-001', '41123', 'Water Rebill', 'Revenue'],
  ['43135-000', '41124', 'Late Charges', 'Revenue'],
  ['43160-000', '41125', 'Lock Key Income', 'Revenue'],
  ['43180-000', '41126', 'NSF Fee Income', 'Revenue'],
  ['43190-000', '41127', 'Parking', 'Revenue'],
  ['43210-000', '41128', 'Accelerated Rent', 'Revenue'],
  ['43215-000', '41129', 'Renter Insurance Fees', 'Revenue'],
  ['43125-000', '70000', 'Interest Income', 'Revenue'],
  // Payroll
  ['51010-000', '60000', 'Salaries & Wages', 'Expense'],
  ['51020-000', '60008', 'Leasing Consultant I', 'Expense'],
  ['58250-000', '60010', 'Incentives', 'Expense'],
  ['51040-000', '60016', 'Maintenance Personnel II', 'Expense'],
  ['51090-000', '60104', '401k Expenses', 'Expense'],
  ['51110-000', '60105', 'Employee Benefits', 'Expense'],
  // Facilities
  ['53182-000', '61034', 'Trash Valet', 'Expense'],
  ['53105-000', '61056', 'Landscape Maintenance', 'Expense'],
  ['53105-001', '61058', 'Landscape Maintenance Rebill', 'Expense'],
  ['53140-000', '61064', 'Pest Control Services', 'Expense'],
  ['54126-000', '61084', 'Follow Up Services', 'Expense'],
  ['58025-000', '61085', 'Other Leasing', 'Expense'],
  ['52260-000', '61087', 'Janitorial Expenses', 'Expense'],
  ['54105-000', '61093', 'Refreshment Supplies', 'Expense'],
  ['53150-000', '61169', 'Pool Contract', 'Expense'],
  ['53180-000', '61170', 'Trash Removal Contract', 'Expense'],
  ['53230-000', '61171', 'Other Services Contract', 'Expense'],
  ['54122-000', '61172', 'Resident Retention', 'Expense'],
  // Utilities
  ['59020-000', '61147', 'Electric Common Areas', 'Expense'],
  ['59030-000', '61148', 'Electric Models', 'Expense'],
  ['59040-000', '61149', 'Electric Vacant', 'Expense'],
  ['59110-000', '61152', 'Water', 'Expense'],
  ['59100-000', '61153', 'Utility Rebill Service Fees', 'Expense'],
  ['59112-000', '61163', 'Sewer', 'Expense'],
  ['53055-000', '61165', 'Equipment Contract', 'Expense'],
  ['53060-000', '61166', 'Fire Alarm Contract', 'Expense'],
  ['53070-000', '61167', 'Fire Protection Contract', 'Expense'],
  ['53090-000', '61168', 'Janitorial Contract', 'Expense'],
  // Legal and accounting
  ['58205-000', '63000', 'Accounting', 'Expense'],
  ['54050-000', '63150', 'Broker Commission Fee', 'Expense'],
  // Office
  ['58290-000', '60451', 'Training', 'Expense'],
  ['58238-000', '60452', 'Compliance Monitor', 'Expense'],
  ['54010-000', '63025', 'Professional Fees', 'Expense'],
  ['61030-000', '63026', 'Management Fees', 'Expense'],
  ['51030-000', '63045', 'Bonuses', 'Expense'],
  ['51030-001', '63046', 'Qtrly Bonuses', 'Expense'],
  ['54038-000', '63100', 'Consulting Fee', 'Expense'],
  ['54025-000', '67001', 'Website', 'Expense'],
  ['58240-000', '67002', 'Computer Expense', 'Expense'],
  ['71585-000', '67003', 'Computer Hardware', 'Expense'],
  ['58090-000', '67004', 'Telephone Expense', 'Expense'],
  ['58035-000', '67005', 'Copy Machine', 'Expense'],
  ['58115-000', '67006', 'Software Licenses / Maintenance Fees', 'Expense'],
  ['58253-000', '67011', 'Employee Recognition', 'Expense'],
  ['58284-000', '67012', 'Technology Fee', 'Expense'],
  ['54090-000', '67152', 'Other Administration', 'Expense'],
  ['54002-000', '67200', 'Office Expense', 'Expense'],
  ['58100-000', '67250', 'Postage & Delivery', 'Expense'],
  ['58080-000', '67251', 'Print Material', 'Expense'],
  ['58110-000', '67300', 'Telephone & Internet', 'Expense'],
  ['53030-000', '67301', 'Cable TV Contract', 'Expense'],
  ['71800-000', '67400', 'Advertising & Marketing', 'Expense'],
  ['54012-000', '67401', 'Advertising - Other', 'Expense'],
  ['71688-000', '67403', 'Marketing', 'Expense'],
  ['54040-000', '67405', 'Traditional - Printing Costs', 'Expense'],
  ['54035-000', '67467', 'Social Media', 'Expense'],
  ['58225-000', '68110', 'Bank Fees', 'Expense'],
  // Taxes and insurance
  ['51120-000', '65000', 'Insurance - Liability', 'Expense'],
  ['63010-000', '68055', 'Property Insurance', 'Expense'],
  // Advertising and promotion
  ['58305-000', '67455', 'Uniforms', 'Expense'],
  ['58107-000', '67457', 'Referrals - Resident', 'Expense'],
  // Other operating expense
  ['72617-000', '68306', 'Lease Up - Other', 'Expense'],
  // Other expense
  ['62020-000', '68061', 'State Franchise Tax', 'Expense'],
  ['71640-000', '75131', 'Equipment', 'Expense'],
  ['71670-000', '75132', 'Golf Carts', 'Expense'],
  ['71530-000', '75133', 'Brochures', 'Expense'],
  ['58320-000', '70350', 'Other Expense', 'Expense'],
  // Seven Foxtail accounts have no line in the CPA package. Three of them
  // carry balances totalling 1,416.10, and the package's operating column net
  // loss of 500,668.76 includes them even though no printed line does — so its
  // consolidating statement of operations does not foot to its own net income.
  // Mapping all seven makes the operating column internally consistent. These
  // target codes are additions to the statement chart, flagged for review.
  ['52700-000', '61089', 'Other Make-Ready Expenses', 'Expense'],       // 413.76
  ['71580-000', '75134', 'Clubhouse', 'Expense'],                       // 729.72
  ['72615-000', '68307', 'Lease Up - Contract Optimization', 'Expense'],// 272.62
  ['43290-000', '41130', 'Miscellaneous Income', 'Revenue'],            // 0.01
  ['51070-000', '60106', 'Payroll Taxes', 'Expense'],                   // nil so far
  ['53185-000', '61035', 'Trash Removal Rebill', 'Expense'],            // nil so far
  ['59105-000', '61154', 'Utility Rebill Service Fee Reimbursement', 'Expense'],
];

// The development-entity accounts that were debited when cash went to Foxtail,
// with the cumulative amounts as at July 2026. Traced to the funding requests
// in the requisition invoice log:
//   11690 Other FF&E           33,250.00  Q1 2026 funding request
//   12597 Other Marketing      55,197.00  Q1 funding request
//   13425 Operating Shortfall  238,779.06 Q4 2025 119,204.00 + Q1 69,575.06
//                                         + Q2 2026 50,000.00
// Total 327,226.06 — the operating ledger's contributed capital to the penny.
// 13425 exists only to carry funding, so it eliminates in full and maintains
// itself. The other two are mixed accounts and carry a stated amount, which the
// user raises as further transfers post.
// [ account code, name, mode, amount ]
const BRAKER_FUNDING = [
  ['11690', 'Other FF&E', 'amount', 33250.00],
  ['12597', 'Other Marketing', 'amount', 55197.00],
  ['13425', 'Operating Shortfall', 'full', 238779.06],
];

function findEntity(db, rx, codes) {
  const rows = db.prepare('SELECT id, name, code FROM entities').all();
  return rows.find(e => codes.includes(String(e.code || '').toUpperCase()))
    || rows.find(e => rx.test(String(e.name || '')));
}

// Idempotent. Creates only what is missing and never overwrites a mapping or an
// amount a user has since edited.
function seedBraker(db) {
  const parent = findEntity(db, /braker\s*qoz\s*business/i, ['BRAKERQO1']);
  const propco = findEntity(db, /braker\s*prop\s*co/i, ['BRAKERPR']);
  const oper = findEntity(db, /braker\s*operating/i, ['BRAKEROP']);
  if (!parent || !propco || !oper) return { seeded: false, reason: 'Braker entities are not all present' };

  const now = new Date().toISOString();
  let g = C.groupForParent(db, parent.id);
  if (!g) {
    db.prepare('INSERT INTO consol_groups (parent_entity_id, scope_key, name, created_at, created_by) VALUES (?,?,?,?,?)')
      .run(parent.id, 'braker', 'Braker QOZ Business, LLC', now, 'seed');
    g = C.groupForParent(db, parent.id);
  }
  const addMember = (eid, label, source, order) => {
    try {
      db.prepare('INSERT INTO consol_members (group_id, entity_id, label, source, sort_order) VALUES (?,?,?,?,?)')
        .run(g.id, eid, label, source, order);
    } catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }
  };
  addMember(parent.id, 'Braker QOZ Business', 'ledger', 0);
  addMember(propco.id, 'Braker Prop Co', 'ledger', 1);
  addMember(oper.id, 'Braker Operating', 'tb', 2);

  const insMap = db.prepare(`INSERT INTO operating_tb_map
    (entity_id, source_code, target_code, target_name, target_type, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?)`);
  let mapped = 0;
  for (const [src, tc, tn, tt] of BRAKER_MAP) {
    try { insMap.run(oper.id, src, tc, tn, tt, now, 'seed'); mapped++; }
    catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }
  }

  try {
    db.prepare(`INSERT INTO consol_investment_pairs
      (group_id, label, holder_entity_id, holder_account_code, issuer_entity_id, issuer_account_code, sort_order, created_at, created_by)
      VALUES (?,?,?,?,?,?,?,?,?)`)
      .run(g.id, 'Investment in Braker PropCo LLC / Contributed Capital - Braker QOZ Business LLC',
        parent.id, '19056', propco.id, '34167', 0, now, 'seed');
  } catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }

  const insFund = db.prepare(`INSERT INTO consol_funding_accounts
    (group_id, entity_id, account_code, account_name, mode, amount, notes, created_at, created_by, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?,?,?,?,?)`);
  let funds = 0;
  for (const [code, name, mode, amt] of BRAKER_FUNDING) {
    try {
      insFund.run(g.id, propco.id, code, name, mode, amt,
        'Seeded from the July 2026 requisition invoice log', now, 'seed', now, 'seed');
      funds++;
    } catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }
  }
  try {
    db.prepare('INSERT INTO consol_funding_capital (group_id, entity_id, account_code) VALUES (?,?,?)')
      .run(g.id, oper.id, '32000');
  } catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }

  return { seeded: true, group_id: g.id, parent: parent.id, propco: propco.id, operating: oper.id, mapped, funds };
}

// ══════════════════════════════ Routes ══════════════════════════════

function registerConsolidationRoutes(app, deps) {
  const { db, auth, requireRole, computeBalances, userHasEntityAccess, workpapersDir } = deps;
  C.ensureSchema(db);
  try {
    const s = seedBraker(db);
    if (s.seeded) console.log('[consol] Braker group ready (group ' + s.group_id + ', +' + s.mapped + ' map rows, +' + s.funds + ' funding accounts)');
    else console.log('[consol] Braker seed skipped: ' + s.reason);
  } catch (e) { console.error('[consol] seed failed:', e.message); }

  const gate = [auth, requireRole('Admin', 'Accountant')];
  const who = req => (req.user && (req.user.name || req.user.email)) || null;
  const tbUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
  const fail = (res, e) => res.status(e && e.status ? e.status : 500).json({ error: (e && e.message) || 'Server error' });

  // Every read spans several ledgers at once, so access is checked per member
  // rather than by the usual single-:eid middleware. A user who cannot see one
  // column cannot run the group's schedules.
  //
  // A trial-balance column has no ledger here and need not have an entity row
  // at all — the property manager's books are not a CloudLedger entity. Ids
  // with no entity are skipped rather than refused, so deleting the placeholder
  // entity cannot lock the group out of its own schedules.
  function assertAccess(req, ids) {
    for (const eid of ids) {
      const exists = db.prepare('SELECT 1 FROM entities WHERE id = ?').get(eid);
      if (!exists) continue;
      if (!userHasEntityAccess(req.user.id, req.user.role, eid)) {
        throw Object.assign(new Error('No access to entity ' + eid), { status: 403 });
      }
    }
  }

  // Resolve a group from its parent entity, refusing anything outside Braker
  // and HP so the feature cannot be pointed at an unrelated entity.
  function scopedGroup(req) {
    const eid = Number(req.params.parent_eid);
    const ent = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(eid);
    if (!ent) throw Object.assign(new Error('Entity not found'), { status: 404 });
    if (!C.scopeKeyFor(ent)) {
      throw Object.assign(new Error('Consolidation from an uploaded operating trial balance is set up for Braker and HP only.'), { status: 400 });
    }
    const group = C.groupForParent(db, eid);
    if (!group) throw Object.assign(new Error('No consolidation group is configured for ' + ent.name), { status: 404 });
    const columns = C.columnsOf(db, group);
    assertAccess(req, columns.map(c => c.entity_id));
    return { ent, group, columns };
  }

  // A column's display name. A trial-balance column may have no entity row at
  // all, so the member's own label is preferred and the entity table is only a
  // fallback. Never falls back to a bare id, which would print on a schedule.
  const memberLabel = (eid) => {
    const m = db.prepare('SELECT label FROM consol_members WHERE entity_id = ? AND label IS NOT NULL AND label <> \'\' LIMIT 1').get(eid);
    if (m && m.label) return m.label;
    const e = db.prepare('SELECT name FROM entities WHERE id = ?').get(eid);
    return e ? e.name : ('Entity ' + eid);
  };
  const entName = (eid) => {
    const e = db.prepare('SELECT name FROM entities WHERE id = ?').get(eid);
    return e ? e.name : memberLabel(eid);
  };

  // Which entities offer consolidation at all — drives the page's picker.
  app.get('/api/consolidation/groups', ...gate, (req, res) => {
    try {
      const out = [];
      for (const g of db.prepare('SELECT * FROM consol_groups ORDER BY id').all()) {
        const ent = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(g.parent_entity_id);
        if (!ent || !C.scopeKeyFor(ent)) continue;
        if (!userHasEntityAccess(req.user.id, req.user.role, g.parent_entity_id)) continue;
        out.push({
          group_id: g.id, parent_entity_id: g.parent_entity_id, parent_name: ent.name, scope_key: g.scope_key,
          columns: C.columnsOf(db, g).map(c => ({ entity_id: c.entity_id, label: c.label || entName(c.entity_id), source: c.source })),
        });
      }
      res.json(out);
    } catch (e) { fail(res, e); }
  });

  // One group's setup: columns, uploaded months, mapping coverage, both rules.
  app.get('/api/consolidation/:parent_eid', ...gate, (req, res) => {
    try {
      const { ent, group, columns } = scopedGroup(req);
      const cols = columns.map(c => {
        const o = { entity_id: c.entity_id, label: c.label || entName(c.entity_id), entity_name: entName(c.entity_id), source: c.source };
        if (c.source === 'tb') {
          o.tb_months = C.tbMonths(db, c.entity_id);
          o.map_rows = db.prepare('SELECT COUNT(*) AS n FROM operating_tb_map WHERE entity_id = ?').get(c.entity_id).n;
        }
        return o;
      });
      res.json({
        group_id: group.id, parent_entity_id: group.parent_entity_id, parent_name: ent.name, scope_key: group.scope_key,
        columns: cols,
        investment_pairs: db.prepare('SELECT * FROM consol_investment_pairs WHERE group_id = ? ORDER BY sort_order, id').all(group.id),
        funding_accounts: db.prepare('SELECT * FROM consol_funding_accounts WHERE group_id = ? ORDER BY entity_id, account_code').all(group.id),
        funding_capital: db.prepare('SELECT * FROM consol_funding_capital WHERE group_id = ?').all(group.id),
      });
    } catch (e) { fail(res, e); }
  });

  // ── Operating trial balance ──
  app.post('/api/consolidation/:parent_eid/operating-tb', ...gate, tbUpload.single('file'), (req, res) => {
    try {
      const { group, columns } = scopedGroup(req);
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const entity_id = Number(req.body.entity_id) || (columns.find(c => c.source === 'tb') || {}).entity_id;
      const as_of = String(req.body.as_of || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(as_of)) return res.status(400).json({ error: 'as_of must be YYYY-MM-DD' });
      if (!columns.find(c => c.entity_id === entity_id && c.source === 'tb')) {
        return res.status(400).json({ error: 'entity_id must be the operating member of this group' });
      }
      // Stored on the month end whatever day was typed, so one month can only
      // ever hold one trial balance.
      const eom = monthEnd(as_of);

      const parsed = parseOperatingTb(req.file.buffer);
      const now = new Date().toISOString();
      const tx = db.transaction(() => {
        db.prepare('DELETE FROM operating_tb WHERE entity_id = ? AND as_of = ?').run(entity_id, eom);
        const ins = db.prepare(`INSERT INTO operating_tb
          (entity_id, as_of, source_code, source_name, forward, debit, credit, ending, uploaded_at, uploaded_by, filename)
          VALUES (?,?,?,?,?,?,?,?,?,?,?)`);
        for (const l of parsed.lines) {
          ins.run(entity_id, eom, l.source_code, l.source_name, l.forward, l.debit, l.credit, l.ending, now, who(req), req.file.originalname || null);
        }
      });
      tx();

      // File the original next to what was parsed from it, so the source
      // document is one click from the schedule that used it. It goes on the
      // PARENT's workpapers, not the trial-balance column's: the property
      // manager is not a CloudLedger entity, so that column may have no entity
      // row and files written against it would be unreachable.
      let workpaper = null;
      try {
        if (workpapersDir) {
          const fileEid = group.parent_entity_id;
          const folder = 'Operating TBs/' + eom;
          const by = who(req) || 'system';
          for (const fp of ['Operating TBs', folder]) {
            try { db.prepare('INSERT INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)').run(fileEid, fp, by); }
            catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }
          }
          const ext = (String(req.file.originalname || '').match(/\.[A-Za-z0-9]+$/) || ['.xlsx'])[0];
          const originalName = eom + ' ' + memberLabel(entity_id) + ' TB' + ext;
          const dir = path.join(workpapersDir, String(fileEid));
          fs.mkdirSync(dir, { recursive: true });
          for (const old of db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(fileEid, folder, originalName)) {
            try { fs.unlinkSync(path.join(dir, old.stored_filename)); } catch (e) {}
            db.prepare('DELETE FROM entity_files WHERE id = ?').run(old.id);
          }
          const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + originalName.replace(/[^a-zA-Z0-9._-]/g, '_');
          fs.writeFileSync(path.join(dir, stored), req.file.buffer);
          db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)')
            .run(fileEid, folder, stored, originalName, req.file.size, req.file.mimetype || null, by);
          workpaper = { entity_id: fileEid, folder_path: folder, file_name: originalName };
        }
      } catch (e) { console.error('[consol] TB workpaper filing failed:', e.message); }

      const unmapped = C.unmappedFor(db, entity_id, eom);
      res.json({
        success: true, entity_id, as_of: eom, count: parsed.lines.length,
        sign_mode: parsed.signMode, residual: parsed.residual, balanced: parsed.balanced,
        has_activity_columns: parsed.hasActivity, has_forward_column: parsed.hasForward,
        activity_residual: parsed.activityResidual,
        unmapped_count: unmapped.length,
        unmapped_nonzero: unmapped.filter(u => Math.abs(u.ending) > 0.004),
        workpaper,
      });
    } catch (e) { fail(res, e); }
  });

  app.get('/api/consolidation/:parent_eid/operating-tb', ...gate, (req, res) => {
    try {
      const { columns } = scopedGroup(req);
      const entity_id = Number(req.query.entity_id) || (columns.find(c => c.source === 'tb') || {}).entity_id;
      const months = C.tbMonths(db, entity_id);
      const as_of = req.query.as_of ? monthEnd(String(req.query.as_of)) : (months[0] || {}).as_of;
      if (!as_of) return res.json({ entity_id, as_of: null, months, lines: [], unmapped: [] });
      const map = C.mapFor(db, entity_id);
      const lines = C.tbAt(db, entity_id, as_of).map(l => {
        const m = map.get(String(l.source_code));
        return {
          source_code: l.source_code, source_name: l.source_name,
          forward: r2(l.forward), debit: r2(l.debit), credit: r2(l.credit), ending: r2(l.ending),
          target_code: m ? m.target_code : null, target_name: m ? m.target_name : null, target_type: m ? m.target_type : null,
        };
      }).sort((a, b) => String(a.source_code).localeCompare(String(b.source_code)));
      res.json({ entity_id, as_of, months, lines, unmapped: C.unmappedFor(db, entity_id, as_of) });
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/consolidation/:parent_eid/operating-tb', ...gate, (req, res) => {
    try {
      const { columns } = scopedGroup(req);
      const entity_id = Number(req.query.entity_id);
      const as_of = monthEnd(String(req.query.as_of || ''));
      if (!columns.find(c => c.entity_id === entity_id && c.source === 'tb')) {
        return res.status(400).json({ error: 'Not the operating member of this group' });
      }
      const r = db.prepare('DELETE FROM operating_tb WHERE entity_id = ? AND as_of = ?').run(entity_id, as_of);
      res.json({ success: true, deleted: r.changes });
    } catch (e) { fail(res, e); }
  });

  // ── Mapping ──
  app.get('/api/consolidation/:parent_eid/map', ...gate, (req, res) => {
    try {
      const { columns } = scopedGroup(req);
      const entity_id = Number(req.query.entity_id) || (columns.find(c => c.source === 'tb') || {}).entity_id;
      res.json({
        entity_id,
        rows: db.prepare('SELECT * FROM operating_tb_map WHERE entity_id = ? ORDER BY source_code').all(entity_id),
        log: db.prepare('SELECT * FROM operating_tb_map_log WHERE entity_id = ? ORDER BY id DESC LIMIT 100').all(entity_id),
      });
    } catch (e) { fail(res, e); }
  });

  app.put('/api/consolidation/:parent_eid/map', ...gate, (req, res) => {
    try {
      const { columns } = scopedGroup(req);
      const entity_id = Number(req.body.entity_id) || (columns.find(c => c.source === 'tb') || {}).entity_id;
      if (!columns.find(c => c.entity_id === entity_id && c.source === 'tb')) {
        return res.status(400).json({ error: 'Not the operating member of this group' });
      }
      const rows = Array.isArray(req.body.rows) ? req.body.rows : [];
      if (!rows.length) return res.status(400).json({ error: 'rows is required' });
      const now = new Date().toISOString();
      const by = who(req);
      let changed = 0;
      const tx = db.transaction(() => {
        for (const r of rows) {
          const src = String(r.source_code || '').trim();
          if (!src) continue;
          const prev = db.prepare('SELECT * FROM operating_tb_map WHERE entity_id = ? AND source_code = ?').get(entity_id, src);
          if (r.remove) {
            if (prev) {
              db.prepare('DELETE FROM operating_tb_map WHERE id = ?').run(prev.id);
              db.prepare('INSERT INTO operating_tb_map_log (entity_id, source_code, old_target, new_target, changed_at, changed_by) VALUES (?,?,?,?,?,?)')
                .run(entity_id, src, prev.target_code, null, now, by);
              changed++;
            }
            continue;
          }
          const tc = String(r.target_code || '').trim();
          const tn = String(r.target_name || '').trim();
          const tt = String(r.target_type || '').trim();
          if (!tc || !tn || !['Asset', 'Liability', 'Equity', 'Revenue', 'Expense'].includes(tt)) {
            throw Object.assign(new Error('Each row needs target_code, target_name and a target_type of Asset, Liability, Equity, Revenue or Expense (' + src + ')'), { status: 400 });
          }
          if (prev) {
            db.prepare('UPDATE operating_tb_map SET target_code=?, target_name=?, target_type=?, notes=?, updated_at=?, updated_by=? WHERE id=?')
              .run(tc, tn, tt, r.notes != null ? r.notes : prev.notes, now, by, prev.id);
          } else {
            db.prepare(`INSERT INTO operating_tb_map (entity_id, source_code, target_code, target_name, target_type, notes, updated_at, updated_by)
              VALUES (?,?,?,?,?,?,?,?)`).run(entity_id, src, tc, tn, tt, r.notes || null, now, by);
          }
          if (!prev || prev.target_code !== tc) {
            db.prepare('INSERT INTO operating_tb_map_log (entity_id, source_code, old_target, new_target, changed_at, changed_by) VALUES (?,?,?,?,?,?)')
              .run(entity_id, src, prev ? prev.target_code : null, tc, now, by);
          }
          changed++;
        }
      });
      tx();
      res.json({ success: true, changed });
    } catch (e) { fail(res, e); }
  });

  // ── Funding accounts: the offsetting development-entity accounts ──
  // Nothing in the ledger says which account a transfer to the property manager
  // was charged to, so the list is maintained here.
  app.get('/api/consolidation/:parent_eid/funding-accounts', ...gate, (req, res) => {
    try {
      const { group } = scopedGroup(req);
      const rows = db.prepare('SELECT * FROM consol_funding_accounts WHERE group_id = ? ORDER BY entity_id, account_code').all(group.id);
      res.json({
        funding_accounts: rows.map(r => Object.assign({}, r, { entity_name: entName(r.entity_id) })),
        funding_capital: db.prepare('SELECT * FROM consol_funding_capital WHERE group_id = ?').all(group.id)
          .map(r => Object.assign({}, r, { entity_name: entName(r.entity_id) })),
      });
    } catch (e) { fail(res, e); }
  });

  app.put('/api/consolidation/:parent_eid/funding-accounts', ...gate, (req, res) => {
    try {
      const { group, columns } = scopedGroup(req);
      const rows = Array.isArray(req.body.rows) ? req.body.rows : [];
      if (!rows.length) return res.status(400).json({ error: 'rows is required' });
      const now = new Date().toISOString();
      const by = who(req);
      let changed = 0;
      const tx = db.transaction(() => {
        for (const r of rows) {
          const eid = Number(r.entity_id);
          const code = String(r.account_code || '').trim();
          if (!eid || !code) continue;
          if (!columns.find(c => c.entity_id === eid && c.source === 'ledger')) {
            throw Object.assign(new Error('entity ' + eid + ' is not a ledger member of this group'), { status: 400 });
          }
          const prev = db.prepare('SELECT * FROM consol_funding_accounts WHERE group_id=? AND entity_id=? AND account_code=?').get(group.id, eid, code);
          if (r.remove) {
            if (prev) { db.prepare('DELETE FROM consol_funding_accounts WHERE id=?').run(prev.id); changed++; }
            continue;
          }
          const acct = db.prepare('SELECT name FROM accounts WHERE entity_id=? AND code=?').get(eid, code);
          if (!acct && !prev) {
            throw Object.assign(new Error('Account ' + code + ' does not exist on ' + entName(eid)), { status: 400 });
          }
          const mode = r.mode === 'full' ? 'full' : 'amount';
          const amount = r2(r.amount);
          const name = String(r.account_name || (acct ? acct.name : '') || (prev ? prev.account_name : '') || '').trim();
          if (prev) {
            db.prepare('UPDATE consol_funding_accounts SET account_name=?, mode=?, amount=?, notes=?, updated_at=?, updated_by=? WHERE id=?')
              .run(name, mode, amount, r.notes != null ? r.notes : prev.notes, now, by, prev.id);
          } else {
            db.prepare(`INSERT INTO consol_funding_accounts (group_id, entity_id, account_code, account_name, mode, amount, notes, created_at, created_by, updated_at, updated_by)
              VALUES (?,?,?,?,?,?,?,?,?,?,?)`).run(group.id, eid, code, name, mode, amount, r.notes || null, now, by, now, by);
          }
          changed++;
        }
      });
      tx();
      res.json({ success: true, changed });
    } catch (e) { fail(res, e); }
  });

  // ── Eliminations and the consolidating schedules ──
  function asOfOf(req) {
    const as_of = String((req.query && req.query.as_of) || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(as_of)) throw Object.assign(new Error('as_of (YYYY-MM-DD) is required'), { status: 400 });
    return monthEnd(as_of);
  }

  // The eliminating entries in debit/credit form, with each rule's own residual.
  app.get('/api/consolidation/:parent_eid/eliminations', ...gate, (req, res) => {
    try {
      const { group } = scopedGroup(req);
      const asOf = asOfOf(req);
      const o = { as_of: asOf, close_pl_before: yearStart(asOf) };
      const { rules, adjustments } = C.computeEliminations(db, group, o, computeBalances, null);
      const entries = adjustments.map(a => {
        // The operating column's accounts live in an uploaded trial balance, not
        // in `accounts`, so the type the engine carried with the adjustment is
        // authoritative and the chart is only a fallback for the name.
        const acct = db.prepare('SELECT name, type FROM accounts WHERE entity_id=? AND code=?').get(a.entity_id, a.code);
        const type = a.type || (acct ? acct.type : null);
        // Which side removes the balance depends on the side the account sits
        // on. Taking out a positive asset or expense is a CREDIT; taking out
        // positive liability, equity or revenue is a DEBIT. A negative balance
        // reverses that.
        const isDr = type === 'Asset' || type === 'Expense';
        const removeWithCredit = isDr ? a.amount > 0 : a.amount < 0;
        const mag = r2(Math.abs(a.amount));
        return {
          entity_id: a.entity_id, entity_name: entName(a.entity_id), code: a.code,
          name: acct ? acct.name : null, type,
          debit: removeWithCredit ? 0 : mag,
          credit: removeWithCredit ? mag : 0,
        };
      });
      const totalDr = r2(entries.reduce((s, e2) => s + e2.debit, 0));
      const totalCr = r2(entries.reduce((s, e2) => s + e2.credit, 0));
      res.json({
        as_of: asOf, rules, entries,
        total_debit: totalDr, total_credit: totalCr, balanced: Math.abs(r2(totalDr - totalCr)) < 0.02,
      });
    } catch (e) { fail(res, e); }
  });

  // Both consolidating schedules, one column per member plus eliminations and
  // the consolidated cross-foot. The balance sheet is as of the month end; the
  // statement of operations is returned for the month AND year to date, each
  // labelled, because the same caption carrying two different periods is what
  // makes the CPA package's two schedules disagree.
  app.get('/api/consolidation/:parent_eid/schedules', ...gate, (req, res) => {
    try {
      const { ent, group, columns } = scopedGroup(req);
      const asOf = asOfOf(req);
      const bs = C.buildColumns(db, group, { as_of: asOf, close_pl_before: yearStart(asOf) }, computeBalances);
      const month = C.buildColumns(db, group, { from: monthStart(asOf), to: asOf }, computeBalances);
      const ytd = C.buildColumns(db, group, { from: yearStart(asOf), to: asOf }, computeBalances);

      const sumBy = (built, pred, pick) => r2(built.accounts.filter(pred).reduce((s, a) => s + pick(a), 0));
      const totals = (built) => {
        const out = { by_entity: {}, elimination: 0, consolidated: 0 };
        for (const c of columns) {
          out.by_entity[c.entity_id] = {
            assets: sumBy(built, a => a.type === 'Asset', a => a.byEntity[c.entity_id] || 0),
            liabilities: sumBy(built, a => a.type === 'Liability', a => a.byEntity[c.entity_id] || 0),
            equity: sumBy(built, a => a.type === 'Equity', a => a.byEntity[c.entity_id] || 0),
            revenue: sumBy(built, a => a.type === 'Revenue', a => a.byEntity[c.entity_id] || 0),
            expense: sumBy(built, a => a.type === 'Expense', a => a.byEntity[c.entity_id] || 0),
          };
        }
        out.elimination = {
          assets: sumBy(built, a => a.type === 'Asset', a => a.elimination),
          liabilities: sumBy(built, a => a.type === 'Liability', a => a.elimination),
          equity: sumBy(built, a => a.type === 'Equity', a => a.elimination),
        };
        out.consolidated = {
          assets: sumBy(built, a => a.type === 'Asset', a => a.consolidated),
          liabilities: sumBy(built, a => a.type === 'Liability', a => a.consolidated),
          equity: sumBy(built, a => a.type === 'Equity', a => a.consolidated),
          revenue: sumBy(built, a => a.type === 'Revenue', a => a.consolidated),
          expense: sumBy(built, a => a.type === 'Expense', a => a.consolidated),
        };
        return out;
      };

      res.json({
        parent_name: ent.name, as_of: asOf,
        columns: columns.map(c => ({ entity_id: c.entity_id, label: c.label || entName(c.entity_id), source: c.source })),
        balance_sheet: { period: 'As of ' + asOf, accounts: bs.accounts, totals: totals(bs), rules: bs.rules, unavailable: bs.unavailable },
        operations_month: { period: 'Month ended ' + asOf, accounts: month.accounts, totals: totals(month), rules: month.rules, unavailable: month.unavailable },
        operations_ytd: { period: 'Year to date ' + asOf, accounts: ytd.accounts, totals: totals(ytd), rules: ytd.rules, unavailable: ytd.unavailable },
      });
    } catch (e) { fail(res, e); }
  });
}

module.exports = { registerConsolidationRoutes, parseOperatingTb, seedBraker, BRAKER_MAP, BRAKER_FUNDING };
