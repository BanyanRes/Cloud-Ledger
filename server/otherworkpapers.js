// ─── CLRF workpaper: Other Workpapers (balance-sheet account support) ─────────
//
// A single quarterly workbook that reproduces the seven Weaver balance-sheet
// support workpapers for County Line Rail Fund I, LP (CL entity 40), each on its
// own tab in Weaver's format, fully GL-derived from the CL ledger:
//
//   1. Due Fr (To) Port Co   — interco reconciliation by property (101100/211100)
//   2. Interest Receivable   — account 120010 transaction detail
//   3. Prepaid Expenses      — 150200 (advisory) + 150300 (insurance) roll-forward
//   4. Other Assets          — account 180100 transaction detail, grouped
//   5. AP Recon              — account 202000 (accounts payable) vs Bill.com
//   6. Accrual & Sub Cash    — accruals booked in the period (210600 / 210000)
//   7. Distributions Payable — account 230100 transaction detail
//
// Every tab ties to the CL general ledger by construction (same account groupings
// the fund Balance Sheet uses in financials.js), so each ties to the Statement of
// Assets, Liabilities and Partners' Capital. Built like the other CLRF workpapers
// (buildData -> buildWorkbook -> saveToWorkpapers), registered by index.js.
//
// NOTE on granularity: CL's ledger carries these balances but not every source
// tag Weaver's QBO/Bill.com workpapers show (per-vendor AP names, project tags on
// pre-2026 Other Assets costs, per-investor distribution lines). Totals tie; the
// sub-groupings reflect what the CL ledger actually carries (by location/class,
// and the GL transaction detail itself).
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const FUND_EID = 40;

// Account groupings — identical to the fund Balance Sheet (server/financials.js).
const ACCT = {
  dueFrom:    (c) => /^1011/.test(c) || /^1012/.test(c),   // 101100 Due From Port Co
  dueToPort:  (c) => /^2111/.test(c),                       // 211100 Due to Portfolio Company
  intRecv:    (c) => /^12001/.test(c),                      // 120010 Interest Receivable
  prepaidAdv: (c) => /^1502/.test(c),                       // 150200 Prepaid Advisory Fees
  prepaidIns: (c) => /^1503/.test(c),                       // 150300 Prepaid Insurance
  otherAsset: (c) => /^1801/.test(c) || /^1800/.test(c),    // 180100 Other Assets
  ap:         (c) => /^(2020|2100|2101|2102|2103|2104|2105|2107|2109|211[2-9])/.test(c),
  tradeAP:    (c) => /^2020/.test(c),                       // 202000 Accounts Payable (Bill.com trade AP)
  accrued:    (c) => /^2100/.test(c),                        // 210000 Accrued Expenses
  mgmtPay:    (c) => /^2106/.test(c),                        // 210600 Payable - Management Fees
  distPay:    (c) => /^230/.test(c),                         // 230100 Distributions Payable / Due to members
};

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

function resolveQuarter(quarterEnd) {
  if (!isDate(quarterEnd)) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const [y, m, d] = quarterEnd.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) throw new Error('quarter_end must be a quarter end date. Received ' + quarterEnd);
  const q = m / 3;
  return {
    label: y + '-Q' + q, year: String(y), quarter: 'Q' + q,
    ys: y + '-01-01', qs: y + '-' + String(m - 2).padStart(2, '0') + '-01',
    end: quarterEnd, prior_ye: (y - 1) + '-12-31',
  };
}

// ── GL transaction detail (mirrors the /gl-detail route), with running balance on
// the account's natural side. Optional `from` (inclusive); always up to `to`.
function glDetail(db, eid, { from, to, match }) {
  const rows = db.prepare(`
    SELECT je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number,
           je.vendor AS vendor, je.memo AS memo,
           jl.account_code AS account_code, a.name AS account_name, a.type AS account_type,
           jl.debit AS debit, jl.credit AS credit, jl.description AS description,
           dc.name AS class_name, dl.name AS location_name, dp.name AS project_name
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    LEFT JOIN dim_classes dc ON dc.id = jl.class_id
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    LEFT JOIN dim_projects dp ON dp.id = jl.project_id
    WHERE je.entity_id = ? ${from ? 'AND je.date >= ?' : ''} AND je.date <= ?
    ORDER BY jl.account_code, je.date, je.entry_num, jl.id
  `).all(...(from ? [eid, from, to] : [eid, to]));
  // Opening balance (before `from`) per account on natural side.
  const opening = new Map();
  if (from) {
    const oRows = db.prepare(`
      SELECT jl.account_code AS account_code, a.type AS account_type,
             SUM(jl.debit) AS td, SUM(jl.credit) AS tc
      FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
      LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
      WHERE je.entity_id = ? AND je.date < ? GROUP BY jl.account_code
    `).all(eid, from);
    for (const r of oRows) {
      const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
      opening.set(r.account_code, r2(isDr ? (r.td || 0) - (r.tc || 0) : (r.tc || 0) - (r.td || 0)));
    }
  }
  const run = new Map();
  const out = [];
  for (const r of rows) {
    if (!match(String(r.account_code))) continue;
    const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
    const delta = isDr ? (r.debit - r.credit) : (r.credit - r.debit);
    const bal = r2((run.has(r.account_code) ? run.get(r.account_code) : (opening.get(r.account_code) || 0)) + delta);
    run.set(r.account_code, bal);
    out.push({
      date: r.date, entry_num: r.entry_num || '', doc_number: r.doc_number || '',
      vendor: r.vendor || '', memo: r.memo || '', description: r.description || '',
      account_code: r.account_code, account_name: r.account_name || '',
      debit: r2(r.debit || 0), credit: r2(r.credit || 0),
      signed: r2(isDr ? (r.debit - r.credit) : (r.credit - r.debit)), balance: bal,
      class_name: r.class_name || '', location_name: r.location_name || '', project_name: r.project_name || '',
    });
  }
  return out;
}

function balanceMap(computeBalances, eid, asOf) {
  const rows = computeBalances(eid, { as_of: asOf }) || [];
  const m = new Map();
  for (const r of rows) m.set(String(r.code), { name: r.name, balance: r2(r.balance) });
  return m;
}
const sumWhere = (bmap, match) => { let s = 0; for (const [c, v] of bmap) if (match(c)) s = r2(s + (v.balance || 0)); return s; };

function buildData(ctx, quarter, opts = {}) {
  const { db, computeBalances } = ctx;
  const eid = opts.entity_id || FUND_EID;
  const flags = []; // exceptions/reviews surfaced on the Summary tab
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  const bEnd = balanceMap(computeBalances, eid, quarter.end);
  const bBeg = balanceMap(computeBalances, eid, quarter.prior_ye);

  // Property map: the CL locations that represent the four rail properties.
  const PROPS = ['CLIP', 'Buna', 'SRN', 'Silsbee'];

  // 1. Due From/To Port Co — net (due-from asset minus due-to liability) by property.
  const dfrom = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: ACCT.dueFrom });
  const dto = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: ACCT.dueToPort });
  // Beginning net by property (prior-YE balances aren't split by location in the
  // balance map, so derive the beginning location split from ITD detail < ys).
  const dfromITD = glDetail(db, eid, { to: quarter.prior_ye, match: ACCT.dueFrom });
  const dtoITD = glDetail(db, eid, { to: quarter.prior_ye, match: ACCT.dueToPort });
  const netByProp = (fromRows, toRows) => {
    const m = {}; PROPS.forEach((p) => (m[p] = 0));
    for (const r of fromRows) { const p = matchProp(r.location_name, PROPS); if (p) m[p] = r2(m[p] + r.signed); }
    for (const r of toRows) { const p = matchProp(r.location_name, PROPS); if (p) m[p] = r2(m[p] - r.signed); } // due-to reduces net
    return m;
  };
  const dueBeg = netByProp(dfromITD, dtoITD);
  // Period JE activity by JE# and property (net effect on due-from/(to)).
  const dueJEs = {};
  const addJE = (r, sign) => {
    const p = matchProp(r.location_name, PROPS); if (!p) return;
    const key = r.entry_num || r.doc_number || r.date;
    dueJEs[key] = dueJEs[key] || { je: key }; PROPS.forEach((pp) => (dueJEs[key][pp] = dueJEs[key][pp] || 0));
    dueJEs[key][p] = r2(dueJEs[key][p] + sign * r.signed);
  };
  for (const r of dfrom) addJE(r, 1);
  for (const r of dto) addJE(r, -1);
  const dueEnd = {}; PROPS.forEach((p) => (dueEnd[p] = r2(dueBeg[p] + Object.values(dueJEs).reduce((a, j) => a + (j[p] || 0), 0))));
  const dueData = { PROPS, beg: dueBeg, jes: Object.values(dueJEs).filter((j) => PROPS.some((p) => Math.abs(j[p]) >= 0.005)), end: dueEnd };
  if (eid === FUND_EID) { dueData.recon = buildDueRecon(ctx, quarter, dueData); for (const f of dueData.recon.flags) flags.push(f); }

  // 2. Interest Receivable — detail.
  const intRows = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: ACCT.intRecv });
  const intBal = sumWhere(bEnd, ACCT.intRecv);
  const interestTie = { gl: intBal, tied: Math.abs(intBal) < 0.005 };
  if (!interestTie.tied) flags.push({ severity: 'review', wp: 'Interest Receivable', message: 'Interest receivable ' + fmt(intBal) + ' — confirm it agrees to the portfolio company book (Silsbee).' });

  // 3. Prepaid Expenses — roll-forward (advisory + insurance).
  const prepaid = {
    advBeg: sumWhere(bBeg, ACCT.prepaidAdv), advEnd: sumWhere(bEnd, ACCT.prepaidAdv),
    insBeg: sumWhere(bBeg, ACCT.prepaidIns), insEnd: sumWhere(bEnd, ACCT.prepaidIns),
  };
  prepaid.advAmort = r2(prepaid.advEnd - prepaid.advBeg);
  prepaid.insAmort = r2(prepaid.insEnd - prepaid.insBeg);

  // 4. Other Assets — inception-to-date detail, grouped by location (falls back to
  // one "Other Assets" group for untagged lines).
  const oaRows = glDetail(db, eid, { to: quarter.end, match: ACCT.otherAsset });
  const oaGroups = groupByLocation(oaRows);
  const oaTotal = sumWhere(bEnd, ACCT.otherAsset);

  // 5. AP Recon — full supporting schedules: the 202000 GL detail for the quarter
  // with debit/credit offset tagging (each paid bill and its payment cancel; the
  // remaining unmatched bills are the open invoices), the Bill.com A/P detail, and
  // a CL A/P detail, all reconciling to the 202000 balance.
  const apGl = sumWhere(bEnd, ACCT.tradeAP);
  const apRecon = buildApRecon(db, eid, quarter, apGl);

  // 6. Accrual & Subsequent Cash Disbursement — JEs that credited the accrual
  // accounts (mgmt-fee payable, accrued expenses) during the period.
  const accrRows = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: (c) => ACCT.mgmtPay(c) || ACCT.accrued(c) })
    .filter((r) => r.credit > 0);
  const mgmtPayBal = sumWhere(bEnd, ACCT.mgmtPay);
  const accruedBal = sumWhere(bEnd, ACCT.accrued);

  // 7. Distributions Payable — detail for 230x.
  const distRows = glDetail(db, eid, { to: quarter.end, match: ACCT.distPay });
  const distBal = sumWhere(bEnd, ACCT.distPay);

  return {
    quarter, entity_name: ent ? ent.name : ('entity ' + eid),
    due: dueData,
    flags,
    interest: { rows: intRows, balance: intBal, tie: interestTie },
    prepaid,
    otherAssets: { groups: oaGroups, total: oaTotal },
    ap: apRecon,
    accrual: { rows: accrRows, mgmtPay: mgmtPayBal, accrued: accruedBal },
    dist: { rows: distRows, balance: distBal },
    ties: {
      due_to_portfolio: sumWhere(bEnd, ACCT.dueToPort), due_from_portfolio: sumWhere(bEnd, ACCT.dueFrom),
      due_net: r2(sumWhere(bEnd, ACCT.dueFrom) - sumWhere(bEnd, ACCT.dueToPort)),
      interest_receivable: intBal, prepaid_advisory: prepaid.advEnd, prepaid_insurance: prepaid.insEnd,
      other_assets: oaTotal, accounts_payable: apGl, accrued_expenses: accruedBal,
      management_fees_payable: mgmtPayBal, distributions_payable: distBal,
    },
  };
}

function matchProp(loc, props) {
  const s = String(loc || '').toLowerCase();
  for (const p of props) if (s.includes(p.toLowerCase()) || (p === 'SRN' && /sabine|s&n|srn/.test(s))) return p;
  return null;
}
function groupByLocation(rows) {
  const g = new Map();
  for (const r of rows) {
    const k = r.location_name || r.project_name || 'Other Assets';
    if (!g.has(k)) g.set(k, []);
    g.get(k).push(r);
  }
  const out = [];
  for (const [name, rs] of g) out.push({ name, rows: rs, total: r2(rs.reduce((a, x) => a + x.signed, 0)) });
  out.sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
  return out;
}

// ─── AP Recon offset engine ──────────────────────────────────────────────────
// Weaver's AP recon workpaper lists the full 202000 GL detail for the quarter,
// tags each debit/credit that offsets another with a shared letter, and whatever
// is left uncancelled equals the ending A/P balance = the open invoices per the
// Bill.com A/P detail. This reproduces that: a symmetric matcher cancels equal-
// and-opposite items in either order (payments against bills, reversals, and the
// lump within-AP reclass JEs), leaving the unmatched credits as the open bills
// and the unmatched debits (which relieve the beginning balance) tagged X.

// Offset-letter sequence: a..z then A..Z (skipping x/X, which flags beginning-
// balance items), then double letters.
function letterSeq(n) {
  const S = 'abcdefghijklmnopqrstuvwyzABCDEFGHIJKLMNOPQRSTUVWYZ'.split('');
  if (n < S.length) return S[n];
  const a = Math.floor(n / S.length) - 1, b = n % S.length;
  return (a >= 0 && a < S.length ? S[a] : 'z') + S[b];
}

// Best-effort invoice number from an aging/GL line.
function parseInv(row) {
  const d = row.doc_number != null ? String(row.doc_number).trim() : '';
  if (d) return d;
  const m = String(row.memo || row.description || '').match(/#\s*([A-Za-z0-9._-]{3,})|—\s*([A-Za-z0-9._-]{3,})/);
  return m ? (m[1] || m[2]) : '';
}

// Sum an array of {amount} by a key function.
function groupSum(rows, keyFn) {
  const m = new Map();
  for (const x of rows) { const k = keyFn(x) || '(unnamed)'; m.set(k, r2((m.get(k) || 0) + (x.amount || 0))); }
  return m;
}

function buildApRecon(db, eid, q, apGlBal) {
  const r2c = (n) => Math.round((Number(n) || 0) * 100);

  // Beginning balance (signed liability) before the quarter start.
  const bRow = db.prepare(
    'SELECT COALESCE(SUM(jl.credit),0) tc, COALESCE(SUM(jl.debit),0) td '
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + "WHERE je.entity_id = ? AND je.date < ? AND jl.account_code LIKE '2020%'"
  ).get(eid, q.qs);
  const begin = r2((bRow.tc || 0) - (bRow.td || 0));

  // Period 202000 lines, in GL order (matches the /gl-detail route ordering).
  const lines = db.prepare(
    'SELECT je.id AS entry_id, je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number, '
    + 'je.vendor AS vendor, je.memo AS memo, jl.id AS line_id, jl.debit AS debit, jl.credit AS credit, jl.description AS description '
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + "WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? AND jl.account_code LIKE '2020%' "
    + 'ORDER BY je.date, je.entry_num, jl.id'
  ).all(eid, q.qs, q.end);

  // Offset account(s) for each entry (the non-AP lines in the same JE).
  const offStmt = db.prepare(
    'SELECT jl.account_code AS code, a.name AS name FROM journal_lines jl '
    + 'LEFT JOIN accounts a ON a.entity_id = ? AND a.code = jl.account_code '
    + "WHERE jl.entry_id = ? AND jl.account_code NOT LIKE '2020%'"
  );
  const offCache = new Map();
  const offsetFor = (entryId) => {
    if (offCache.has(entryId)) return offCache.get(entryId);
    const rs = offStmt.all(eid, entryId);
    let v;
    if (!rs.length) v = '';
    else if (rs.length === 1) v = rs[0].code + (rs[0].name ? ' ' + rs[0].name : '');
    else v = '-Split-';
    offCache.set(entryId, v); return v;
  };

  // Amount -> vendor/invoice lookup, from the real vendored bill credit lines.
  const vmap = new Map();
  for (const l of lines) {
    const cr = r2(l.credit || 0);
    if (cr > 0 && l.vendor && String(l.vendor).trim()) {
      const k = r2c(cr);
      if (!vmap.has(k)) vmap.set(k, {
        vendor: String(l.vendor).trim(), invoice: l.doc_number != null ? String(l.doc_number) : '',
        date: l.date, num: l.entry_num != null ? String(l.entry_num) : '', memo: l.memo || l.description || '',
      });
    }
  }
  const vlookup = (amt, row) => {
    const hit = vmap.get(r2c(amt));
    if (hit) return hit;
    const mm = String((row.memo || row.description) || '').match(/^Bill\s*-\s*([^:]+):/i);
    return {
      vendor: mm ? mm[1].trim() : (row.vendor || ''), invoice: row.doc_number != null ? String(row.doc_number) : '',
      date: row.date, num: row.entry_num != null ? String(row.entry_num) : '', memo: row.memo || row.description || '',
    };
  };

  // Symmetric offset matcher. Each open lot carries a sign (+1 credit / -1 debit).
  // A new line cancels an existing opposite-sign lot (or subset of up to three)
  // of equal amount, regardless of order; both sides then share one letter.
  const open = []; // { amtc, sign, idx }
  const tag = new Array(lines.length).fill(null);
  let pair = 0;
  const findOpp = (amtc, sign) => {
    const want = -sign, pos = [];
    for (let i = 0; i < open.length; i++) if (open[i].sign === want) pos.push(i);
    for (const i of pos) if (open[i].amtc === amtc) return [i];
    for (let a = 0; a < pos.length; a++) for (let b = a + 1; b < pos.length; b++)
      if (open[pos[a]].amtc + open[pos[b]].amtc === amtc) return [pos[a], pos[b]];
    for (let a = 0; a < pos.length; a++) for (let b = a + 1; b < pos.length; b++) for (let c = b + 1; c < pos.length; c++)
      if (open[pos[a]].amtc + open[pos[b]].amtc + open[pos[c]].amtc === amtc) return [pos[a], pos[b], pos[c]];
    return null;
  };
  lines.forEach((l, i) => {
    const dr = r2(l.debit || 0), cr = r2(l.credit || 0);
    const sign = cr > 0 ? 1 : -1, amtc = r2c(cr > 0 ? cr : dr);
    if (amtc === 0) { tag[i] = ''; return; }
    const m = findOpp(amtc, sign);
    if (m) {
      const lt = letterSeq(pair++);
      tag[i] = lt;
      m.forEach((p) => (tag[open[p].idx] = lt));
      m.sort((a, b) => b - a).forEach((p) => open.splice(p, 1));
    } else {
      open.push({ amtc, sign, idx: i });
    }
  });
  // Unmatched debit lots relieve the beginning balance -> X; unmatched credit
  // lots are the open invoices at period end (left unlettered).
  const openBillIdx = [];
  let xTotal = 0;
  for (const o of open) {
    if (o.sign < 0) { tag[o.idx] = 'X'; xTotal = r2(xTotal + o.amtc / 100); }
    else { tag[o.idx] = ''; openBillIdx.push(o.idx); }
  }

  // GL rows with running balance from the beginning balance.
  let bal = begin;
  const glRows = lines.map((l, i) => {
    const dr = r2(l.debit || 0), cr = r2(l.credit || 0);
    bal = r2(bal + cr - dr);
    return {
      date: l.date, num: (l.entry_num != null ? String(l.entry_num) : '') || String(l.doc_number || ''),
      vendor: l.vendor || '', offset: offsetFor(l.entry_id), memo: l.description || l.memo || '',
      debit: dr, credit: cr, balance: bal, letter: tag[i],
    };
  });

  // Open invoices per GL (unmatched credits) with vendor/invoice detail.
  const openBills = openBillIdx.map((i) => {
    const l = lines[i], amt = r2(l.credit || 0), v = vlookup(amt, l);
    return { vendor: v.vendor, invoice: v.invoice, date: v.date || l.date, num: v.num, amount: amt, memo: v.memo };
  });
  openBills.sort((a, b) => b.amount - a.amount);
  const openTotal = r2(openBills.reduce((a, x) => a + x.amount, 0));

  // Bill.com A/P aging when uploaded; otherwise the open GL bills ARE the Bill.com
  // open invoices (CLRF A/P is synced from Bill.com).
  let billcom = null, billcomSource = 'gl';
  try {
    const cfg = db.prepare('SELECT ap_aging_lines_json, ap_aging_as_of FROM billcom_config WHERE entity_id = ?').get(eid);
    if (cfg && cfg.ap_aging_lines_json) {
      const arr = JSON.parse(cfg.ap_aging_lines_json) || [];
      const norm = arr.map((x) => ({
        vendor: x.vendor || '', invoice: x.invoice_number || parseInv(x) || '',
        date: x.bill_date || '', amount: r2(x.amount || 0),
      })).filter((x) => Math.abs(x.amount) >= 0.005);
      if (norm.length) { billcom = { asOf: cfg.ap_aging_as_of || '', lines: norm, total: r2(norm.reduce((a, x) => a + x.amount, 0)) }; billcomSource = 'aging'; }
    }
  } catch (e) { billcom = null; }
  const billcomLines = billcom ? billcom.lines : openBills.map((b) => ({ vendor: b.vendor, invoice: b.invoice, date: b.date, amount: b.amount }));
  const billcomTotal = billcom ? billcom.total : openTotal;
  const billcomAsOf = billcom ? billcom.asOf : q.end;

  // Reconciliation by vendor: Bill.com open vs GL open.
  const glByV = groupSum(openBills, (x) => String(x.vendor || '').trim());
  const bcByV = groupSum(billcomLines, (x) => String(x.vendor || '').trim());
  const vnames = new Set([...glByV.keys(), ...bcByV.keys()]);
  const byVendor = [...vnames].map((k) => {
    const gl = glByV.get(k) || 0, bc = bcByV.get(k) || 0;
    return { vendor: k, gl, billcom: bc, diff: r2(bc - gl) };
  }).sort((a, b) => b.gl - a.gl);

  return {
    gl: apGlBal, begin,
    ending: r2(begin + glRows.reduce((a, r) => a + r.credit - r.debit, 0)),
    totalDebit: r2(glRows.reduce((a, r) => a + r.debit, 0)),
    totalCredit: r2(glRows.reduce((a, r) => a + r.credit, 0)),
    glRows, openBills, openTotal, xTotal,
    billcomLines, billcomTotal, billcomAsOf, billcomSource,
    byVendor,
  };
}

// ─── Cross-entity intercompany recon + exception flags ───────────────────────
// A workpaper's job is to prove the GL balance is CORRECT against an independent
// source, not just reprint the GL. For the intercompany balances that source is
// the counterparty's own ledger (CLRF's portfolio companies are entities here),
// so we compute both sides and flag any leg that does not mirror.

// Portfolio-property token -> the CL entity that carries the mirror balance.
const PORT_ENTITY = [
  { token: 'Silsbee', re: /silsbee property owner/i },
  { token: 'Buna', re: /buna property owner/i },
  { token: 'SRN', re: /sabine river|northern railroad/i },
  { token: 'CLIP', re: /clip property owner/i },
];
function resolvePortfolioEntities(db) {
  const ents = db.prepare('SELECT id, name FROM entities').all();
  const map = {};
  for (const pe of PORT_ENTITY) {
    const hit = ents.find((e) => pe.re.test(String(e.name)));
    if (hit) map[pe.token] = { id: hit.id, name: hit.name };
  }
  return map;
}

// A counterparty's net position owed TO the fund, from its CLRF-facing loan/due
// accounts: a payable-to-fund liability counts positive, a due-from-fund asset
// negative. Contributed capital / equity is excluded — that is the investment
// recon, not the due-from/(to) recon.
function counterpartyOwedToFund(computeBalances, cpEid, asOf) {
  const rows = computeBalances(cpEid, { as_of: asOf }) || [];
  let net = 0; const legs = [];
  for (const r of rows) {
    if (!/county line rail fund|clrf/i.test(String(r.name))) continue;
    const ty = String(r.type || '');
    if (ty === 'Liability') { net = r2(net + (r.balance || 0)); legs.push({ code: String(r.code), name: r.name, amt: r2(r.balance) }); }
    else if (ty === 'Asset') { net = r2(net - (r.balance || 0)); legs.push({ code: String(r.code), name: r.name, amt: r2(-(r.balance || 0)) }); }
  }
  return { net: r2(net), legs };
}

// When a leg does not mirror, look INTO the counterparty's GL and list the
// fund-side transaction(s) that have no matching entry on the other ledger.
function traceDueMismatch(ctx, quarter, prop, cpEnt, cpLegCodes) {
  const { db } = ctx;
  const fundRows = glDetail(db, FUND_EID, { from: quarter.qs, to: quarter.end, match: (c) => ACCT.dueFrom(c) || ACCT.dueToPort(c) })
    .filter((r) => matchProp(r.location_name, [prop]) === prop);
  let cpAmts = [];
  if (cpEnt && cpLegCodes && cpLegCodes.size) {
    cpAmts = glDetail(db, cpEnt.id, { from: quarter.qs, to: quarter.end, match: (c) => cpLegCodes.has(String(c)) })
      .map((r) => Math.round(Math.abs(r.signed) * 100));
  }
  const out = [];
  for (const r of fundRows) {
    if (Math.abs(r.signed) < 0.005) continue;
    const key = Math.round(Math.abs(r.signed) * 100);
    const i = cpAmts.indexOf(key);
    if (i >= 0) { cpAmts[i] = -1; continue; } // has a mirror on the other side
    out.push({
      date: r.date, num: r.entry_num || r.doc_number, amount: r.signed, memo: r.memo || r.description || '',
      note: (r.entry_num || r.doc_number || '') + ' ' + fmt(Math.abs(r.signed)) + ' "' + String(r.memo || r.description || '').slice(0, 44)
        + '" recorded on CLRF with no matching entry on ' + (cpEnt ? cpEnt.name : 'the counterparty') + ' as of ' + short(quarter.end)
        + ' — likely timing (cash cleared in the subsequent period)',
    });
  }
  return out;
}

// Build the two-sided due-from/(to) reconciliation by property, with flags.
function buildDueRecon(ctx, quarter, dueData) {
  const { db, computeBalances } = ctx;
  const portMap = resolvePortfolioEntities(db);
  const rows = [], flags = [];
  for (const prop of dueData.PROPS) {
    const fundDue = r2(dueData.end[prop] || 0);
    const cpEnt = portMap[prop] || null;
    const cp = cpEnt ? counterpartyOwedToFund(computeBalances, cpEnt.id, quarter.end) : { net: 0, legs: [] };
    if (Math.abs(fundDue) < 0.005 && Math.abs(cp.net) < 0.005) continue;
    const diff = r2(fundDue - cp.net);
    let status;
    if (Math.abs(diff) < 0.005) status = 'matched';
    else if (!cpEnt || Math.abs(cp.net) < 0.005) status = 'one_sided';
    else status = 'mismatch';
    let trace = [];
    if (status !== 'matched') {
      const codes = new Set((cp.legs || []).map((l) => String(l.code)));
      trace = traceDueMismatch(ctx, quarter, prop, cpEnt, codes);
    }
    rows.push({ prop, cpName: cpEnt ? cpEnt.name : '(no CL entity)', cpId: cpEnt ? cpEnt.id : null, fundDue, cpNet: cp.net, diff, status, legs: cp.legs, trace });
    if (status !== 'matched') {
      flags.push({
        severity: 'exception', wp: 'Due From/To Port Co',
        message: prop + ': CLRF records ' + fmt(fundDue) + ' but ' + (cpEnt ? cpEnt.name : 'the counterparty') + ' shows ' + fmt(cp.net)
          + ' (off by ' + fmt(diff) + ')' + (trace[0] ? ' — ' + trace[0].note : ''),
      });
    }
  }
  return { rows, portMap, flags };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const NAVY = 'FF1F3864';
const HDR = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const SMALLI = F({ size: 9, italic: true });
const THIN = { style: 'thin' };
const BLUE = F({ color: { argb: 'FF0000FF' } });
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const spellDate = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };
const short = (end) => { const [y, m, d] = String(end).split('-').map(Number); return m + '/' + d + '/' + String(y).slice(2); };

function hdrRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true };
  });
  if (widths) widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
}
function titleBlock(ws, name, subtitle, dateLine) {
  ws.getCell('A1').value = name; ws.getCell('A1').font = F({ size: 12, bold: true });
  ws.getCell('A2').value = subtitle; ws.getCell('A2').font = F({ bold: true });
  ws.getCell('A3').value = dateLine; ws.getCell('A3').font = SMALLI;
}
const setMoney = (ws, ref, v, o = {}) => { const c = ws.getCell(ref); c.value = v; c.numFmt = MONEY; c.font = F(o); return c; };

function buildWorkbook(data) {
  const q = data.quarter, en = data.entity_name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  // ── Summary / index ──────────────────────────────────────────────────────
  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  su.getColumn(1).width = 34; su.getColumn(2).width = 40; su.getColumn(3).width = 16; su.getColumn(4).width = 18;
  titleBlock(su, en, 'Other Workpapers — Balance Sheet Account Support', 'As of ' + spellDate(q.end));
  hdrRow(su, 5, ['Workpaper', 'Balance Sheet line', 'GL account(s)', 'CL balance'], [34, 40, 16, 18]);
  { const fl0 = data.flags || []; su.getCell('A4').value = fl0.length ? ('⚠ ' + fl0.length + ' item(s) require attention — see Exceptions below') : '✓ All balances tie'; su.getCell('A4').font = F({ bold: true, color: { argb: fl0.length ? 'FFC00000' : 'FF008000' } }); su.mergeCells('A4:D4'); }
  const idx = [
    ['Due Fr (To) Port Co', 'Due from (to) portfolio investments (net)', '101100 / 211100', data.ties.due_net],
    ['Interest Receivable', 'Interest receivable', '120010', data.ties.interest_receivable],
    ['Prepaid Expenses', 'Prepaid insurance / advisory', '150300 / 150200', r2(data.ties.prepaid_insurance + data.ties.prepaid_advisory)],
    ['Other Assets', 'Other assets', '180100', data.ties.other_assets],
    ['AP Recon', 'Accounts payable and accrued expenses', '202000 + 210000', r2(data.ties.accounts_payable + data.ties.accrued_expenses)],
    ['Accrual & Sub Cash Disbursement', 'Management fees payable', '210600', data.ties.management_fees_payable],
    ['Distributions Payable', 'Due to members', '230100', data.ties.distributions_payable],
  ];
  let sr = 6;
  for (const [wpn, bsl, acc] of idx) {
    su.getCell('A' + sr).value = wpn; su.getCell('A' + sr).font = F();
    su.getCell('B' + sr).value = bsl; su.getCell('B' + sr).font = F();
    su.getCell('C' + sr).value = acc; su.getCell('C' + sr).font = F();
    sr++;
  }
  // The D-column balances are written as live formulas linking to each supporting
  // tab's total cell, after those tabs are built (see end of buildWorkbook).
  su.getCell('A' + (sr + 1)).value = 'Each Summary balance is a live formula linked to the total on its supporting tab; every tab is GL-derived and ties to the Statement of Assets, Liabilities and Partners’ Capital.';
  su.getCell('A' + (sr + 1)).font = SMALLI; su.mergeCells('A' + (sr + 1) + ':D' + (sr + 1));
  {
    const fl = data.flags || [];
    let exRow = sr + 3;
    su.getCell('A' + exRow).value = fl.length ? ('⚠ Exceptions — Review Required (' + fl.length + ')') : '✓ No exceptions — all balances tie and mirror the counterparty ledgers.';
    su.getCell('A' + exRow).font = F({ bold: true, color: { argb: fl.length ? 'FFC00000' : 'FF008000' } });
    su.mergeCells('A' + exRow + ':D' + exRow);
    let er = exRow + 1;
    for (const f of fl) {
      su.getCell('A' + er).value = (f.severity === 'exception' ? '✖ ' : '△ ') + f.wp;
      su.getCell('A' + er).font = F({ bold: true, color: { argb: f.severity === 'exception' ? 'FFC00000' : 'FFB8860B' } });
      su.getCell('B' + er).value = f.message; su.getCell('B' + er).font = F(); su.getCell('B' + er).alignment = { wrapText: true };
      su.mergeCells('B' + er + ':D' + er);
      er++;
    }
  }

  // ── 1. Due Fr (To) Port Co ────────────────────────────────────────────────
  const du = wb.addWorksheet('Due Fr (To) Port Co', { views: [{ showGridLines: false }] });
  du.getColumn(2).width = 40; du.getColumn(3).width = 14; [5, 7, 9, 11].forEach((c) => (du.getColumn(c).width = 14));
  titleBlock(du, en, 'Due From/To Port Co reconciliation', spellDate(q.end));
  const P = data.due.PROPS;
  du.getCell('C5').value = 'JE#'; du.getCell('C5').font = F({ bold: true });
  P.forEach((p, i) => { const c = du.getCell(String.fromCharCode(69 + i * 2) + '5'); c.value = p; c.font = F({ bold: true }); c.alignment = { horizontal: 'center' }; });
  const propCol = (i) => String.fromCharCode(69 + i * 2); // E,G,I,K
  let R = 6;
  du.getCell('B' + R).value = 'Due From (To) Port Co at ' + short(q.prior_ye); du.getCell('B' + R).font = F();
  P.forEach((p, i) => setMoney(du, propCol(i) + R, data.due.beg[p]));
  R++;
  for (const j of data.due.jes) {
    du.getCell('C' + R).value = j.je; du.getCell('C' + R).font = F();
    P.forEach((p, i) => setMoney(du, propCol(i) + R, j[p] || 0));
    R++;
  }
  du.getCell('B' + R).value = 'Due From (To) Port Co at ' + short(q.end); du.getCell('B' + R).font = F({ bold: true });
  P.forEach((p, i) => { const c = setMoney(du, propCol(i) + R, data.due.end[p], { bold: true }); c.border = { top: THIN }; });
  const dueEndRow = R; // ending row — Summary links to -SUM(E:K) here
  du.getCell('B' + (R + 2)).value = 'Net Due From (To) Portfolio Company ties to the Balance Sheet: due-from asset (101100) less due-to liability (211100).';
  du.getCell('B' + (R + 2)).font = SMALLI; du.mergeCells('B' + (R + 2) + ':K' + (R + 2));
  {
    const rec = data.due.recon;
    if (rec && rec.rows && rec.rows.length) {
      let rr = R + 4;
      du.getCell('B' + rr).value = 'Intercompany reconciliation to portfolio company ledgers'; du.getCell('B' + rr).font = F({ bold: true }); rr += 1;
      du.getCell('B' + rr).value = 'Property / Counterparty'; du.getCell('B' + rr).font = F({ bold: true });
      du.getCell('E' + rr).value = 'CLRF'; du.getCell('E' + rr).font = F({ bold: true }); du.getCell('E' + rr).alignment = { horizontal: 'right' };
      du.getCell('G' + rr).value = 'Per counterparty'; du.getCell('G' + rr).font = F({ bold: true }); du.getCell('G' + rr).alignment = { horizontal: 'right' };
      du.getCell('I' + rr).value = 'Difference'; du.getCell('I' + rr).font = F({ bold: true }); du.getCell('I' + rr).alignment = { horizontal: 'right' };
      du.getCell('K' + rr).value = 'Status'; du.getCell('K' + rr).font = F({ bold: true }); rr += 1;
      for (const row of rec.rows) {
        du.getCell('B' + rr).value = row.prop + ' — ' + row.cpName; du.getCell('B' + rr).font = F();
        setMoney(du, 'E' + rr, row.fundDue); setMoney(du, 'G' + rr, row.cpNet); setMoney(du, 'I' + rr, row.diff);
        const st = row.status === 'matched' ? 'Tied' : (row.status === 'one_sided' ? 'One-sided' : 'Mismatch');
        const cc = du.getCell('K' + rr); cc.value = st; cc.font = F({ bold: row.status !== 'matched', color: { argb: row.status === 'matched' ? 'FF008000' : 'FFC00000' } });
        rr += 1;
        for (const tr of (row.trace || [])) { du.getCell('C' + rr).value = '↳ ' + tr.note; du.getCell('C' + rr).font = SMALLI; du.mergeCells('C' + rr + ':K' + rr); rr += 1; }
      }
      du.getCell('B' + (rr + 1)).value = 'Each property CLRF balance is agreed to the portfolio company own ledger (its loan payable to / due from CLRF). A difference is a real exception — most often a timing item where one side has posted and the other has not.';
      du.getCell('B' + (rr + 1)).font = SMALLI; du.mergeCells('B' + (rr + 1) + ':K' + (rr + 1));
    }
  }

  // ── 2. Interest Receivable ────────────────────────────────────────────────
  const intTotalRow = detailSheet(wb, 'Interest Receivable', en, 'Interest Receivable (120010)', q, data.interest.rows, data.interest.balance);

  // ── 3. Prepaid Expenses ───────────────────────────────────────────────────
  const pp = wb.addWorksheet('Prepaid Expenses', { views: [{ showGridLines: false }] });
  pp.getColumn(1).width = 34; [2, 3, 4].forEach((c) => (pp.getColumn(c).width = 18));
  titleBlock(pp, en, 'Schedule for Prepaid Expenses & Prepaid Insurance', 'As of ' + short(q.end));
  hdrRow(pp, 5, ['', 'Prepaid Advisory (150200)', 'Prepaid Insurance (150300)', 'Total'], [34, 20, 22, 16]);
  pp.getCell('A6').value = 'Balance at ' + short(q.prior_ye); pp.getCell('A6').font = F();
  setMoney(pp, 'B6', data.prepaid.advBeg); setMoney(pp, 'C6', data.prepaid.insBeg); setMoney(pp, 'D6', r2(data.prepaid.advBeg + data.prepaid.insBeg));
  pp.getCell('A7').value = 'Additions / (amortization) in ' + q.label; pp.getCell('A7').font = F();
  setMoney(pp, 'B7', data.prepaid.advAmort); setMoney(pp, 'C7', data.prepaid.insAmort); setMoney(pp, 'D7', r2(data.prepaid.advAmort + data.prepaid.insAmort));
  pp.getCell('A8').value = 'Balance at ' + short(q.end); pp.getCell('A8').font = F({ bold: true });
  ['B', 'C', 'D'].forEach((col, i) => { const vals = [data.prepaid.advEnd, data.prepaid.insEnd, r2(data.prepaid.advEnd + data.prepaid.insEnd)]; const c = setMoney(pp, col + '8', vals[i], { bold: true }); c.border = { top: THIN }; });
  pp.getCell('A10').value = 'Tied to Balance Sheet.'; pp.getCell('A10').font = SMALLI;

  // ── 4. Other Assets (grouped detail) ──────────────────────────────────────
  const oa = wb.addWorksheet('Other Assets', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  titleBlock(oa, en, 'Other Assets (180100) — transaction detail', 'As of ' + short(q.end));
  hdrRow(oa, 5, ['Group', 'Date', 'Type', 'Num', 'Description', 'Amount', 'Balance'], [22, 12, 14, 16, 52, 16, 16]);
  let orow = 6;
  for (const g of data.otherAssets.groups) {
    oa.getCell('A' + orow).value = g.name; oa.getCell('A' + orow).font = F({ bold: true }); orow++;
    let run = 0;
    for (const r of g.rows) {
      run = r2(run + r.signed);
      oa.getCell('B' + orow).value = r.date; oa.getCell('B' + orow).font = F();
      oa.getCell('C' + orow).value = ''; // type
      oa.getCell('D' + orow).value = r.entry_num || r.doc_number; oa.getCell('D' + orow).font = F();
      oa.getCell('E' + orow).value = r.description || r.memo; oa.getCell('E' + orow).font = F();
      setMoney(oa, 'F' + orow, r.signed); setMoney(oa, 'G' + orow, run);
      orow++;
    }
    oa.getCell('E' + orow).value = 'Total for ' + g.name; oa.getCell('E' + orow).font = F({ bold: true });
    const c = setMoney(oa, 'F' + orow, g.total, { bold: true }); c.border = { top: THIN }; orow += 1;
  }
  oa.getCell('E' + orow).value = 'TOTAL Other Assets'; oa.getCell('E' + orow).font = F({ bold: true });
  const oc = setMoney(oa, 'F' + orow, data.otherAssets.total, { bold: true }); oc.border = { top: THIN, bottom: THIN };
  const oaTotalRow = orow; // Summary links to F here

  // ── 5. AP Recon (summary + supporting schedules) ──────────────────────────
  // Purpose: show that the open-invoice report per Bill.com ties to the GL A/P
  // balance (202000). The summary reconciles Bill.com vs GL by vendor; supporting
  // tabs give the Bill.com A/P detail, the CL open-invoice list per GL, and the
  // full 202000 GL detail with offset-letter tagging (matched debits/credits
  // cancel; items tagged X relieve the beginning balance; the unlettered credits
  // that remain are the open invoices and equal the ending A/P balance).
  const AP = data.ap;
  const ap = wb.addWorksheet('AP Recon', { views: [{ showGridLines: false }] });
  ap.getColumn(1).width = 44; [2, 3, 4].forEach((c) => (ap.getColumn(c).width = 16));
  titleBlock(ap, en, 'Accounts Payable Reconciliation — Bill.com to GL (202000)', 'As of ' + short(q.end));
  hdrRow(ap, 6, ['Vendor', 'Per Bill.com', 'Per GL', 'Difference'], [44, 16, 16, 16]);
  let ar = 7;
  const apFirst = ar;
  for (const v of (AP.byVendor || [])) {
    ap.getCell('A' + ar).value = v.vendor; ap.getCell('A' + ar).font = F();
    setMoney(ap, 'B' + ar, v.billcom);
    setMoney(ap, 'C' + ar, v.gl);
    { const c = ap.getCell('D' + ar); c.value = { formula: 'B' + ar + '-C' + ar, result: v.diff }; c.numFmt = MONEY; c.font = F(); }
    ar++;
  }
  const apLast = ar - 1;
  ap.getCell('A' + ar).value = 'Total accounts payable'; ap.getCell('A' + ar).font = F({ bold: true });
  const sumOrVal = (col, tot) => (apLast >= apFirst ? { formula: 'SUM(' + col + apFirst + ':' + col + apLast + ')', result: tot } : tot);
  { const c = ap.getCell('B' + ar); c.value = sumOrVal('B', AP.billcomTotal); c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  { const c = ap.getCell('C' + ar); c.value = sumOrVal('C', AP.openTotal); c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  { const c = ap.getCell('D' + ar); c.value = { formula: 'B' + ar + '-C' + ar, result: r2(AP.billcomTotal - AP.openTotal) }; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  const apTotalRow = ar; // Summary links to C here (Per GL open invoices = 202000)
  ar += 2;
  ap.getCell('A' + ar).value = 'Accounts payable per general ledger (202000)'; ap.getCell('A' + ar).font = F();
  setMoney(ap, 'C' + ar, AP.gl); const apGlRow = ar; ar++;
  ap.getCell('A' + ar).value = 'Open invoices remaining per GL detail (offsets applied)'; ap.getCell('A' + ar).font = F();
  setMoney(ap, 'C' + ar, AP.openTotal); const apOpenRow = ar; ar++;
  ap.getCell('A' + ar).value = 'Difference'; ap.getCell('A' + ar).font = F({ bold: true });
  { const c = ap.getCell('C' + ar); c.value = { formula: 'C' + apGlRow + '-C' + apOpenRow, result: r2(AP.gl - AP.openTotal) }; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  ar += 2;
  ap.getCell('A' + ar).value = AP.billcomSource === 'aging'
    ? ('Per Bill.com = uploaded Bill.com A/P aging as of ' + (AP.billcomAsOf || short(q.end)) + '. Per GL = open invoices remaining on account 202000 after offsets, which ties to the Balance Sheet.')
    : ('Per Bill.com = open invoices per the Bill.com A/P sync (no A/P aging file uploaded for this entity; CLRF A/P is synced from Bill.com, so the open GL bills are the Bill.com open invoices). Per GL = open invoices remaining on account 202000 after offsets, which ties to the Balance Sheet.');
  ap.getCell('A' + ar).font = SMALLI; ap.mergeCells('A' + ar + ':D' + ar);

  // 5a. Bill.com A/P Detail — open invoices.
  const bc = wb.addWorksheet('Bill.com AP Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  bc.getColumn(1).width = 40; bc.getColumn(2).width = 18; bc.getColumn(3).width = 14; bc.getColumn(4).width = 16;
  titleBlock(bc, en, 'Bill.com A/P Detail — Open Invoices', 'As of ' + short(AP.billcomAsOf || q.end));
  hdrRow(bc, 6, ['Vendor', 'Invoice #', 'Bill Date', 'Amount'], [40, 18, 14, 16]);
  let bcr = 7; const bcFirst = bcr;
  for (const b of (AP.billcomLines || [])) {
    bc.getCell('A' + bcr).value = b.vendor; bc.getCell('A' + bcr).font = F();
    bc.getCell('B' + bcr).value = b.invoice || ''; bc.getCell('B' + bcr).font = F();
    bc.getCell('C' + bcr).value = b.date || ''; bc.getCell('C' + bcr).font = F();
    setMoney(bc, 'D' + bcr, b.amount); bcr++;
  }
  const bcLast = bcr - 1;
  bc.getCell('A' + bcr).value = 'Total open invoices per Bill.com'; bc.getCell('A' + bcr).font = F({ bold: true });
  { const c = bc.getCell('D' + bcr); c.value = bcLast >= bcFirst ? { formula: 'SUM(D' + bcFirst + ':D' + bcLast + ')', result: AP.billcomTotal } : AP.billcomTotal; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  bc.getCell('A' + (bcr + 2)).value = AP.billcomSource === 'aging'
    ? 'Source: uploaded Bill.com A/P aging detail. Ties to account 202000 and to the CL AP Detail tab.'
    : 'Source: open bills per the Bill.com sync into CL (no A/P aging file uploaded). Ties to account 202000 and to the CL AP Detail tab.';
  bc.getCell('A' + (bcr + 2)).font = SMALLI; bc.mergeCells('A' + (bcr + 2) + ':D' + (bcr + 2));

  // 5b. CL A/P Detail — open invoices per GL.
  const cld = wb.addWorksheet('CL AP Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  cld.getColumn(1).width = 40; cld.getColumn(2).width = 18; cld.getColumn(3).width = 14; cld.getColumn(4).width = 10; cld.getColumn(5).width = 16;
  titleBlock(cld, en, 'Accounts Payable Detail per GL — Open Invoices', 'As of ' + short(q.end));
  hdrRow(cld, 6, ['Vendor', 'Invoice #', 'Bill Date', 'JE #', 'Amount'], [40, 18, 14, 10, 16]);
  let clr = 7; const clFirst = clr;
  for (const b of (AP.openBills || [])) {
    cld.getCell('A' + clr).value = b.vendor; cld.getCell('A' + clr).font = F();
    cld.getCell('B' + clr).value = b.invoice || ''; cld.getCell('B' + clr).font = F();
    cld.getCell('C' + clr).value = b.date || ''; cld.getCell('C' + clr).font = F();
    cld.getCell('D' + clr).value = b.num || ''; cld.getCell('D' + clr).font = F();
    setMoney(cld, 'E' + clr, b.amount); clr++;
  }
  const clLast = clr - 1;
  cld.getCell('A' + clr).value = 'Total open invoices per GL'; cld.getCell('A' + clr).font = F({ bold: true });
  { const c = cld.getCell('E' + clr); c.value = clLast >= clFirst ? { formula: 'SUM(E' + clFirst + ':E' + clLast + ')', result: AP.openTotal } : AP.openTotal; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  cld.getCell('A' + (clr + 2)).value = 'These are the credits on account 202000 left uncancelled after the offset analysis (AP GL Detail tab). Ties to the A/P balance and to the Bill.com A/P Detail.';
  cld.getCell('A' + (clr + 2)).font = SMALLI; cld.mergeCells('A' + (clr + 2) + ':E' + (clr + 2));

  // 5c. AP GL Detail — 202000 detail with offset-letter tagging + bottom recon.
  const gld = wb.addWorksheet('AP GL Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  gld.getColumn(1).width = 11; gld.getColumn(2).width = 8; gld.getColumn(3).width = 26; gld.getColumn(4).width = 26;
  gld.getColumn(5).width = 46; gld.getColumn(6).width = 14; gld.getColumn(7).width = 14; gld.getColumn(8).width = 15; gld.getColumn(9).width = 8;
  titleBlock(gld, en, 'Accounts Payable (202000) — GL Detail with Offset', q.qs + ' to ' + short(q.end));
  hdrRow(gld, 6, ['Date', 'Num', 'Vendor', 'Offset Account', 'Description', 'Debit', 'Credit', 'Balance', 'Offset'], [11, 8, 26, 26, 46, 14, 14, 15, 8]);
  let gr = 7;
  gld.getCell('A' + gr).value = 'Beginning balance ' + short(q.prior_ye); gld.getCell('A' + gr).font = F({ bold: true }); gld.mergeCells('A' + gr + ':E' + gr);
  setMoney(gld, 'H' + gr, AP.begin, { bold: true });
  { const c = gld.getCell('I' + gr); c.value = 'X'; c.font = F({ bold: true }); c.alignment = { horizontal: 'center' }; }
  gr++;
  for (const x of AP.glRows) {
    gld.getCell('A' + gr).value = x.date; gld.getCell('A' + gr).font = F();
    gld.getCell('B' + gr).value = x.num; gld.getCell('B' + gr).font = F();
    gld.getCell('C' + gr).value = x.vendor; gld.getCell('C' + gr).font = F();
    gld.getCell('D' + gr).value = x.offset; gld.getCell('D' + gr).font = F();
    gld.getCell('E' + gr).value = x.memo; gld.getCell('E' + gr).font = F();
    if (x.debit) setMoney(gld, 'F' + gr, x.debit);
    if (x.credit) setMoney(gld, 'G' + gr, x.credit);
    setMoney(gld, 'H' + gr, x.balance);
    { const c = gld.getCell('I' + gr); c.value = x.letter || ''; c.font = (x.letter === 'X') ? F({ bold: true }) : F(); c.alignment = { horizontal: 'center' }; }
    gr++;
  }
  gld.getCell('E' + gr).value = 'TOTAL'; gld.getCell('E' + gr).font = F({ bold: true });
  setMoney(gld, 'F' + gr, AP.totalDebit, { bold: true }).border = { top: THIN };
  setMoney(gld, 'G' + gr, AP.totalCredit, { bold: true }).border = { top: THIN };
  setMoney(gld, 'H' + gr, AP.ending, { bold: true }).border = { top: THIN };
  gr += 2;
  gld.getCell('E' + gr).value = 'Reconciliation'; gld.getCell('E' + gr).font = F({ bold: true }); gr++;
  const rBeg = gr;  gld.getCell('E' + gr).value = 'Beginning A/P balance (tagged X)'; gld.getCell('E' + gr).font = F(); setMoney(gld, 'H' + gr, AP.begin); gr++;
  const rCr = gr;   gld.getCell('E' + gr).value = 'Add: credits (bills booked) in period'; gld.getCell('E' + gr).font = F(); setMoney(gld, 'H' + gr, AP.totalCredit); gr++;
  const rDr = gr;   gld.getCell('E' + gr).value = 'Less: debits (payments / reversals) in period'; gld.getCell('E' + gr).font = F(); setMoney(gld, 'H' + gr, r2(-AP.totalDebit)); gr++;
  const rEnd = gr;  gld.getCell('E' + gr).value = 'Ending A/P balance per GL (202000)'; gld.getCell('E' + gr).font = F({ bold: true });
  { const c = gld.getCell('H' + gr); c.value = { formula: 'H' + rBeg + '+H' + rCr + '+H' + rDr, result: AP.gl }; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  gr += 2;
  const rOpen = gr; gld.getCell('E' + gr).value = 'Open invoices remaining after offsets'; gld.getCell('E' + gr).font = F(); setMoney(gld, 'H' + gr, AP.openTotal); gr++;
  gld.getCell('E' + gr).value = 'Difference (should be zero)'; gld.getCell('E' + gr).font = F({ bold: true });
  { const c = gld.getCell('H' + gr); c.value = { formula: 'H' + rEnd + '-H' + rOpen, result: r2(AP.gl - AP.openTotal) }; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  gr += 2;
  gld.getCell('A' + gr).value = 'Matched debits and credits carry the same offset letter and net to zero. Items tagged X relieve the beginning A/P balance (payments of prior-period bills, total ' + fmt(AP.xTotal) + '). The unlettered credits that remain are the open invoices, which equal the ending A/P balance.';
  gld.getCell('A' + gr).font = SMALLI; gld.mergeCells('A' + gr + ':I' + gr);

  // ── 6. Accrual & Subsequent Cash Disbursement ─────────────────────────────
  const ac = wb.addWorksheet('Accrual & Sub Cash Disb', { views: [{ showGridLines: false }] });
  titleBlock(ac, en, 'Accrual & Subsequent Cash Disbursement', 'As of ' + short(q.end));
  hdrRow(ac, 6, ['Date', 'Vendor', 'JE#', 'Description', 'Accrual Account', 'Amount'], [12, 22, 16, 46, 26, 16]);
  let cr = 7;
  for (const r of data.accrual.rows) {
    ac.getCell('A' + cr).value = r.date; ac.getCell('A' + cr).font = F();
    ac.getCell('B' + cr).value = r.vendor; ac.getCell('B' + cr).font = F();
    ac.getCell('C' + cr).value = r.entry_num || r.doc_number; ac.getCell('C' + cr).font = F();
    ac.getCell('D' + cr).value = r.description || r.memo; ac.getCell('D' + cr).font = F();
    ac.getCell('E' + cr).value = r.account_code + ' ' + r.account_name; ac.getCell('E' + cr).font = F();
    setMoney(ac, 'F' + cr, r.credit); cr++;
  }
  ac.getCell('D' + cr).value = 'Total accruals booked in period'; ac.getCell('D' + cr).font = F({ bold: true });
  setMoney(ac, 'F' + cr, r2(data.accrual.rows.reduce((a, x) => a + x.credit, 0)), { bold: true }).border = { top: THIN };
  // Ending balances (the Summary tab links to the total of these two cells).
  const accEnd1 = cr + 2, accEnd2 = cr + 3, accEndTot = cr + 4;
  ac.getCell('D' + accEnd1).value = 'Management fees payable, end of quarter (210600)'; ac.getCell('D' + accEnd1).font = F();
  setMoney(ac, 'F' + accEnd1, data.accrual.mgmtPay);
  ac.getCell('D' + accEnd2).value = 'Accrued expenses, end of quarter (210000)'; ac.getCell('D' + accEnd2).font = F();
  setMoney(ac, 'F' + accEnd2, data.accrual.accrued);
  ac.getCell('D' + accEndTot).value = 'Total per Balance Sheet'; ac.getCell('D' + accEndTot).font = F({ bold: true });
  { const c = ac.getCell('F' + accEndTot); c.value = { formula: 'F' + accEnd1 + '+F' + accEnd2, result: r2(data.accrual.mgmtPay + data.accrual.accrued) }; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  ac.getCell('A' + (accEndTot + 2)).value = 'Both ending balances tie to the Balance Sheet.'; ac.getCell('A' + (accEndTot + 2)).font = SMALLI;

  // ── 7. Distributions Payable ──────────────────────────────────────────────
  const distTotalRow = detailSheet(wb, 'Distributions Payable', en, 'Distributions Payable (230100)', q, data.dist.rows, data.dist.balance, true);

  // ── Link each Summary balance to the total cell on its supporting schedule ──
  const sq = (t) => "'" + t + "'!";
  const sumRefs = [
    // Net Due From (To): SUM of the property columns on the ending row. A positive
    // result is a net due-FROM (asset), shown positive to match the Balance Sheet.
    { f: 'SUM(' + sq('Due Fr (To) Port Co') + 'E' + dueEndRow + ':K' + dueEndRow + ')', v: data.ties.due_net },
    { f: sq('Interest Receivable') + 'G' + intTotalRow, v: data.ties.interest_receivable },
    { f: sq('Prepaid Expenses') + 'D8', v: r2(data.ties.prepaid_insurance + data.ties.prepaid_advisory) },
    { f: sq('Other Assets') + 'F' + oaTotalRow, v: data.ties.other_assets },
    // Accounts payable & accrued expenses = trade AP (202000, AP Recon) + accrued
    // (210000, Accrual tab), tying to the single Balance Sheet line.
    { f: sq('AP Recon') + 'C' + apTotalRow + '+' + sq('Accrual & Sub Cash Disb') + 'F' + accEnd2, v: r2(data.ties.accounts_payable + data.ties.accrued_expenses) },
    { f: sq('Accrual & Sub Cash Disb') + 'F' + accEnd1, v: data.ties.management_fees_payable },
    { f: sq('Distributions Payable') + 'G' + distTotalRow, v: data.ties.distributions_payable },
  ];
  sumRefs.forEach((x, i) => { const c = su.getCell('D' + (6 + i)); c.value = { formula: x.f, result: x.v }; c.numFmt = MONEY; c.font = F(); });

  return wb;
}

const fmt = (n) => '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Generic "Account QuickReport" transaction-detail sheet (Weaver's GL layout).
function detailSheet(wb, tabName, en, subtitle, q, rows, balance, itd) {
  const ws = wb.addWorksheet(tabName, { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  titleBlock(ws, en, 'Account QuickReport — ' + subtitle, (itd ? 'Inception to ' : q.ys + ' to ') + short(q.end));
  hdrRow(ws, 5, ['Location', 'Date', 'Type', 'Num', 'Description', 'Amount', 'Balance'], [16, 12, 14, 16, 52, 16, 16]);
  let r = 6;
  for (const x of rows) {
    ws.getCell('A' + r).value = x.location_name || x.class_name || ''; ws.getCell('A' + r).font = F();
    ws.getCell('B' + r).value = x.date; ws.getCell('B' + r).font = F();
    ws.getCell('C' + r).value = ''; // transaction type not stored
    ws.getCell('D' + r).value = x.entry_num || x.doc_number; ws.getCell('D' + r).font = F();
    ws.getCell('E' + r).value = x.description || x.memo; ws.getCell('E' + r).font = F();
    setMoney(ws, 'F' + r, x.signed); setMoney(ws, 'G' + r, x.balance);
    r++;
  }
  ws.getCell('E' + r).value = 'TOTAL'; ws.getCell('E' + r).font = F({ bold: true });
  setMoney(ws, 'F' + r, r2(rows.reduce((a, x) => a + x.signed, 0)), { bold: true }).border = { top: THIN };
  setMoney(ws, 'G' + r, balance, { bold: true }).border = { top: THIN };
  ws.getCell('E' + (r + 2)).value = 'Ties to the Balance Sheet.'; ws.getCell('E' + (r + 2)).font = SMALLI;
  return r; // TOTAL row — Summary links to G here
}

// ── Persistence (same shape as the other CLRF workpapers). ────────────────────
const folderFor = (q) => 'Workpapers/Other Workpapers/' + q.year + '/' + q.quarter;
const fileNameFor = (q) => 'CLRF_Other_Workpapers_' + q.label + '.xlsx';

function saveToWorkpapers(ctx, eid, q, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(q), original = fileNameFor(q);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid)); fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function findWorkpaper(ctx, eid, quarterEnd) {
  const q = resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ? ORDER BY id DESC LIMIT 1')
    .get(eid, folderFor(q), fileNameFor(q));
  if (!row) return null;
  return Object.assign({}, row, { quarter: q, abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerOtherWorkpapersRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/other/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const q = resolveQuarter((req.body && req.body.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, q, { entity_id: eid });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, q, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Other-Summary', JSON.stringify({
          quarter: q.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          ties: data.ties, exceptions: (data.flags || []).length, flags: (data.flags || []),
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerOtherWorkpapersRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, FUND_EID };
