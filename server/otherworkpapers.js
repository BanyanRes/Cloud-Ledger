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
const { HAIKU_MODEL } = require('./aiModel');

const FUND_EID = 40;

// Account groupings — identical to the fund Balance Sheet (server/financials.js).
const ACCT = {
  dueFrom:    (c) => /^1011/.test(c) || /^1012/.test(c),   // 101100 Due From Port Co
  dueToPort:  (c) => /^2111/.test(c),                       // 211100 Due to Portfolio Company
  intRecv:    (c) => /^12001/.test(c),                      // 120010 Interest Receivable
  prepaidAdv: (c) => /^1502/.test(c),                       // 150200 Prepaid Advisory Fees
  prepaidIns: (c) => /^1503/.test(c),                       // 150300 Prepaid Insurance
  prepaidSub: (c) => /^1504/.test(c),                       // 150400 Prepaid Subscription
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
    SELECT je.id AS entry_id, je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number,
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
      date: r.date, entry_id: r.entry_id, entry_num: r.entry_num || '', doc_number: r.doc_number || '',
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

async function buildData(ctx, quarter, opts = {}) {
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
    if (!dueJEs[key].desc) dueJEs[key].desc = String(r.memo || r.description || '').trim();
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

  // 3. Prepaid Expenses — item-level roll-forward (advisory + insurance + subscription).
  const prepaid = {
    advBeg: sumWhere(bBeg, ACCT.prepaidAdv), advEnd: sumWhere(bEnd, ACCT.prepaidAdv),
    insBeg: sumWhere(bBeg, ACCT.prepaidIns), insEnd: sumWhere(bEnd, ACCT.prepaidIns),
    subBeg: sumWhere(bBeg, ACCT.prepaidSub), subEnd: sumWhere(bEnd, ACCT.prepaidSub),
  };
  prepaid.advAmort = r2(prepaid.advEnd - prepaid.advBeg);
  prepaid.insAmort = r2(prepaid.insEnd - prepaid.insBeg);
  prepaid.subAmort = r2(prepaid.subEnd - prepaid.subBeg);
  const _ppi = await buildPrepaidItems(ctx, eid, quarter); prepaid.items = _ppi.items; prepaid.itemsByAcct = _ppi.byAcct; for (const f of _ppi.flags) flags.push(f);

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
  const apRecon = await buildApRecon(ctx, eid, quarter, apGl);
  for (const f of (apRecon.flags || [])) flags.push(f);

  // 6. Accrual & Subsequent Cash Disbursement — JEs that credited the accrual
  // accounts (mgmt-fee payable, accrued expenses) during the period.
  const accrRows = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: (c) => ACCT.mgmtPay(c) || ACCT.accrued(c) || /^2110/.test(c) })
    .filter((r) => r.credit > 0);
  const mgmtPayBal = sumWhere(bEnd, ACCT.mgmtPay);
  const accruedBal = sumWhere(bEnd, ACCT.accrued);
  const affiliatesBal = sumWhere(bEnd, (c) => /^2110/.test(c));
  const mgmtItd = buildMgmtItd(db, quarter);
  const accrualItems = buildAccrualItems(db, eid, quarter, bEnd);
  for (const g of accrualItems.groups) if (Math.abs(g.schedule - g.gl) >= 0.01) flags.push({ severity: 'exception', wp: 'Accrual & Sub Cash Disb', message: g.label + ': accrual schedule net outstanding ' + fmt(g.schedule) + ' does not tie to the GL balance ' + fmt(g.gl) + '.' });
  const subDisb = buildSubsequentDisb(db, quarter);
  const flux = buildFlux(db, quarter);
  for (const f of flux.flags) flags.push(f);
  if (!subDisb.posted) flags.push({ severity: 'review', wp: 'Accrual & Sub Cash Disb', message: 'No subsequent-period cash disbursements are posted yet (' + subDisb.from + ' to ' + subDisb.to + ') — the subsequent-payment support is incomplete until the next month is booked.' });

  // 7. Distributions Payable — detail for 230x.
  const distRows = glDetail(db, eid, { to: quarter.end, match: ACCT.distPay });
  const distBal = sumWhere(bEnd, ACCT.distPay);
  const contrib = buildContribRecv(db, quarter, bEnd);
  const dueMgmtBal = sumWhere(bEnd, (c) => /^2108/.test(c));
  const distSplit = buildDistributions(db, quarter, bEnd, bBeg);

  return {
    quarter, entity_name: ent ? ent.name : ('entity ' + eid),
    due: dueData,
    flags,
    interest: { rows: intRows, balance: intBal, tie: interestTie },
    prepaid,
    otherAssets: { groups: oaGroups, total: oaTotal, begin: sumWhere(bBeg, ACCT.otherAsset) },
    ap: apRecon,
    accrual: { rows: accrRows, items: accrualItems.items, groups: accrualItems.groups, mgmtPay: mgmtPayBal, accrued: accruedBal, affiliates: affiliatesBal, mgmtItd, subDisb, flux },
    contrib,
    distSplit,
    dist: { rows: distRows, balance: distBal },
    ties: {
      due_to_portfolio: sumWhere(bEnd, ACCT.dueToPort), due_from_portfolio: sumWhere(bEnd, ACCT.dueFrom),
      due_net: r2(sumWhere(bEnd, ACCT.dueFrom) - sumWhere(bEnd, ACCT.dueToPort)),
      interest_receivable: intBal, prepaid_advisory: prepaid.advEnd, prepaid_insurance: prepaid.insEnd, prepaid_subscription: prepaid.subEnd,
      other_assets: oaTotal, accounts_payable: apGl, accrued_expenses: accruedBal,
      management_fees_payable: mgmtPayBal, distributions_payable: distBal,
      contribution_receivable: contrib.balance, due_to_mgmt: dueMgmtBal, due_to_affiliates: affiliatesBal,
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

async function buildApRecon(ctx, eid, q, apGlBal) {
  const db = ctx.db;
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
      line_id: l.line_id, date: l.date, num: (l.entry_num != null ? String(l.entry_num) : '') || String(l.doc_number || ''),
      vendor: l.vendor || '', offset: offsetFor(l.entry_id), memo: l.description || l.memo || '',
      debit: dr, credit: cr, balance: bal, letter: tag[i],
    };
  });

  // Open invoices per GL (unmatched credits) with vendor/invoice detail.
  let openBills = openBillIdx.map((i) => {
    const l = lines[i], amt = r2(l.credit || 0), v = vlookup(amt, l);
    return { vendor: v.vendor, invoice: v.invoice, date: v.date || l.date, num: v.num, amount: amt, memo: v.memo };
  });
  openBills.sort((a, b) => b.amount - a.amount);
  let openTotal = r2(openBills.reduce((a, x) => a + x.amount, 0));
  let apSource = 'matcher';

  // Open A/P is sourced from the GL-native A/P aging (the same report shown in the
  // app under A/P Aging): it nets each Bill.com payment against the SPECIFIC bill it
  // settled and carries unapplied debits forward, so it ties to the 202000 GL
  // balance by construction. The quarter-window amount-matcher above can overstate
  // open A/P when a payment does not exact-amount-match a single bill inside the
  // quarter (a combined or cross-period payment). When the aging is available and
  // agrees to this account's GL balance, use it so the workpaper matches the CL A/P
  // Aging report and the Balance Sheet.
  try {
    const aging = (typeof ctx.buildApAging === 'function') ? ctx.buildApAging(eid, q.end, '202000') : null;
    if (aging && aging.grand_total && Math.abs(r2(aging.gl_balance) - r2(apGlBal)) < 0.01) {
      const ob = [];
      for (const v of (aging.vendors || [])) for (const r of (v.rows || [])) ob.push({ vendor: v.vendor, invoice: r.num || '', date: r.date, num: String(r.entry_num || ''), amount: r2(r.amount), memo: '' });
      for (const g of (aging.gl_rows || [])) ob.push({ vendor: '', invoice: '', date: g.date, num: String(g.entry_num || ''), amount: r2(g.amount), memo: g.memo || '' });
      ob.sort((a, b) => b.amount - a.amount);
      openBills = ob;
      openTotal = r2(aging.grand_total.total);
      apSource = 'aging';
      const openSet = new Set((aging.open_items || []).map((o) => o.line_id).filter((x) => x != null));
      for (const r of glRows) r.letter = openSet.has(r.line_id) ? 'O' : '';
    }
  } catch (e) { /* fall back to the amount-matcher result */ }

  // Independent Bill.com A/P detail. The whole point of this recon is to agree
  // the GL A/P to Bill.com's OWN open-invoice report, so the Bill.com side must
  // come from Bill.com — live from the API as of the date, or an uploaded
  // Bill.com A/P aging. It is NEVER sourced from the GL (that would be GL vs GL).
  const apFlags = [];
  let billcom = null, billcomSource = 'none';
  try {
    if (false && typeof ctx.billcomOpenAsOf === 'function') { // live org-wide reconstruction is not reliable enough for the recon; the uploaded Bill.com A/P Detail report is the source
      const live = await ctx.billcomOpenAsOf(eid, q.end);
      if (Array.isArray(live) && live.length) {
        const norm = live.map((x) => ({ vendor: x.vendor || '', invoice: x.invoice_number || x.invoice || '', date: String(x.bill_date || x.date || '').slice(0, 10), amount: r2(x.amount || 0) })).filter((x) => Math.abs(x.amount) >= 0.005);
        if (norm.length) { billcom = { asOf: q.end, lines: norm, total: r2(norm.reduce((a, x) => a + x.amount, 0)) }; billcomSource = 'live'; }
      }
    }
  } catch (e) { billcom = null; }
  if (!billcom) {
    try {
      const cfg = db.prepare('SELECT ap_aging_lines_json, ap_aging_as_of FROM billcom_config WHERE entity_id = ?').get(eid);
      if (cfg && cfg.ap_aging_lines_json) {
        const arr = JSON.parse(cfg.ap_aging_lines_json) || [];
        const norm = arr.map((x) => ({ vendor: x.vendor || '', invoice: x.invoice_number || parseInv(x) || '', date: x.bill_date || '', amount: r2(x.amount || 0) })).filter((x) => Math.abs(x.amount) >= 0.005);
        if (norm.length) { billcom = { asOf: cfg.ap_aging_as_of || '', lines: norm, total: r2(norm.reduce((a, x) => a + x.amount, 0)) }; billcomSource = 'aging'; }
      }
    } catch (e) { billcom = null; }
  }
  const billcomLines = billcom ? billcom.lines : [];
  const billcomTotal = billcom ? billcom.total : null;
  const billcomAsOf = billcom ? billcom.asOf : q.end;
  if (!billcom) {
    apFlags.push({ severity: 'exception', wp: 'AP Recon', message: 'No Bill.com A/P Detail report has been uploaded, so the GL A/P of ' + fmt(apGlBal) + ' is not yet independently verified against Bill.com. Export the Bill.com A/P Detail (Open Items) report as of ' + short(q.end) + ' and upload it (Bill.com settings → A/P aging), then regenerate.' });
  } else if (Math.abs((billcomTotal || 0) - openTotal) >= 0.01) {
    apFlags.push({ severity: 'exception', wp: 'AP Recon', message: 'Bill.com open A/P ' + fmt(billcomTotal) + ' (per ' + billcomSource + ') does not agree to GL open A/P ' + fmt(openTotal) + ' (off by ' + fmt((billcomTotal || 0) - openTotal) + ')' });
  }

  // Reconciliation by vendor: Bill.com open vs GL open.
  const glByV = groupSum(openBills, (x) => String(x.vendor || '').trim());
  const bcByV = groupSum(billcomLines, (x) => String(x.vendor || '').trim());
  const vnames = new Set([...glByV.keys(), ...bcByV.keys()]);
  const byVendor = [...vnames].map((k) => {
    const gl = glByV.get(k) || 0, bc = bcByV.get(k) || 0;
    return { vendor: k, gl, billcom: bc, diff: r2(bc - gl) };
  }).sort((a, b) => b.gl - a.gl);
  if (billcom) for (const v of byVendor) { if (Math.abs(v.diff) >= 0.01) apFlags.push({ severity: 'exception', wp: 'AP Recon', message: v.vendor + ': Bill.com ' + fmt(v.billcom) + ' vs GL ' + fmt(v.gl) + ' (off by ' + fmt(v.diff) + ')' }); }

  return {
    gl: apGlBal, begin, flags: apFlags,
    ending: r2(begin + glRows.reduce((a, r) => a + r.credit - r.debit, 0)),
    totalDebit: r2(glRows.reduce((a, r) => a + r.debit, 0)),
    totalCredit: r2(glRows.reduce((a, r) => a + r.credit, 0)),
    glRows, openBills, openTotal, xTotal, apSource,
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
  // The counterparty's SUBSEQUENT-period CLRF-facing activity, to tell a genuine
  // timing difference (it posts on the other ledger next period) from a real
  // exception (it never does).
  let cpSub = [];
  if (cpEnt) {
    const subTo = addMonths(quarter.end, 3);
    try {
      cpSub = db.prepare(
        'SELECT je.date date, je.entry_num num, jl.debit debit, jl.credit credit '
        + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
        + 'LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code '
        + "WHERE je.entity_id = ? AND je.date > ? AND je.date <= ? "
        + "AND (a.name LIKE '%County Line Rail Fund%' OR a.name LIKE '%CLRF%')"
      ).all(cpEnt.id, quarter.end, subTo);
    } catch (e) { cpSub = []; }
  }
  const out = [];
  for (const r of fundRows) {
    if (Math.abs(r.signed) < 0.005) continue;
    const key = Math.round(Math.abs(r.signed) * 100);
    const i = cpAmts.indexOf(key);
    if (i >= 0) { cpAmts[i] = -1; continue; } // already mirrored in-period
    const head = fmt(Math.abs(r.signed)) + ' — "' + String(r.memo || r.description || '').slice(0, 60) + '"';
    const sub = cpSub.find((x) => Math.round((x.debit || 0) * 100) === key || Math.round((x.credit || 0) * 100) === key);
    let tail;
    if (sub) tail = ' recorded on CLRF at ' + short(quarter.end) + '; CONFIRMED TIMING — posts on ' + cpEnt.name + ' ' + sub.date + ' (' + (sub.num || '') + ').';
    else tail = ' recorded on CLRF at ' + short(quarter.end) + ' with NO matching entry on ' + (cpEnt ? cpEnt.name : 'the counterparty') + ' through ' + addMonths(quarter.end, 3) + ' — REAL EXCEPTION, investigate.';
    out.push({ date: r.date, num: r.entry_num || r.doc_number, amount: r.signed, memo: r.memo || r.description || '', confirmed: !!sub, note: head + tail });
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
      const allTiming = trace.length > 0 && trace.every((t) => t.confirmed);
      flags.push({
        severity: allTiming ? 'review' : 'exception', wp: 'Due From/To Port Co',
        message: prop + ': CLRF records ' + fmt(fundDue) + ' but ' + (cpEnt ? cpEnt.name : 'the counterparty') + ' shows ' + fmt(cp.net)
          + ' (off by ' + fmt(diff) + ')' + (trace[0] ? ' — ' + trace[0].note : ''),
      });
    }
  }
  return { rows, portMap, flags };
}

// ═══ Phase 2-4 substantive-support helpers (inserted into otherworkpapers.js) ═══
function addMonths(dateStr, n) { const [y, m, d] = String(dateStr).split('-').map(Number); const dt = new Date(Date.UTC(y, m - 1 + n, d)); return dt.toISOString().slice(0, 10); }
function nextDay(dateStr) { const [y, m, d] = String(dateStr).split('-').map(Number); const dt = new Date(Date.UTC(y, m - 1, d + 1)); return dt.toISOString().slice(0, 10); }

// ── Prepaid: verify amortization against the booking memo formula + invoice ──
function parseAmortMemo(memo) {
  const m = String(memo || '').match(/\[\s*\$?([\d,]+(?:\.\d+)?)\s*\/\s*(\d+)\s*\*\s*(\d+)\s*\]/);
  if (!m) return null;
  return { premium: Number(m[1].replace(/,/g, '')), basis: Number(m[2]), days: Number(m[3]) };
}
function normDate(s) { const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (!m) return null; let y = m[3]; if (y.length === 2) y = '20' + y; return y + '-' + String(m[1]).padStart(2, '0') + '-' + String(m[2]).padStart(2, '0'); }
function extractCoverage(text) {
  const t = String(text || '').replace(/\s+/g, ' ');
  const dr = t.match(/(?:policy period|coverage period|policy term|effective(?:\s*date)?)\D{0,20}(\d{1,2}\/\d{1,2}\/\d{2,4})\s*(?:to|through|-|–|—)\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/i)
    || t.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s*(?:to|through|–|—)\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  const pr = t.match(/(?:total premium|premium|invoice total|total due|amount due|total)\D{0,10}\$?\s*([\d,]+\.\d{2})/i);
  if (!dr && !pr) return null;
  return { start: dr ? normDate(dr[1]) : null, end: dr ? normDate(dr[2]) : null, premium: pr ? Number(pr[1].replace(/,/g, '')) : null };
}
async function llmExtractCoverage(text) {
  const body = { model: HAIKU_MODEL, max_tokens: 200, messages: [{ role: 'user', content: 'From this insurance invoice text, reply with ONLY compact JSON {"start":"YYYY-MM-DD","end":"YYYY-MM-DD","premium":number} for the policy coverage period and total premium. Text:\n' + String(text).slice(0, 6000) }] };
  const r = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers: { 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) { try { console.warn('[otherworkpapers] Claude coverage extract failed:', r.status, (await r.text()).slice(0, 300)); } catch (_e) {} return null; }
  const j = await r.json(); const txt = (j.content && j.content[0] && j.content[0].text) || ''; const m = txt.match(/\{[\s\S]*\}/); if (!m) return null;
  const o = JSON.parse(m[0]); if (!o.start || !o.end) return null; return o;
}
async function readInvoiceCoverage(uploadDir, db, entryId) {
  if (!uploadDir || !entryId) return null;
  let text = '';
  try {
    const atts = db.prepare('SELECT filename, mime_type, original_name FROM journal_attachments WHERE entry_id = ? ORDER BY id').all(entryId);
    const pdf = atts.find((a) => /pdf/i.test(a.mime_type || '') || /\.pdf$/i.test(a.original_name || ''));
    if (!pdf) return null;
    const fp = path.resolve(uploadDir, pdf.filename);
    if (!fs.existsSync(fp)) return null;
    const pdfParse = require('pdf-parse');
    const parsed = await pdfParse(fs.readFileSync(fp));
    text = String(parsed.text || '');
  } catch (e) { return null; }
  const cov = extractCoverage(text);
  if (cov) return Object.assign({ source: 'invoice' }, cov);
  if (process.env.ANTHROPIC_API_KEY) { try { const llm = await llmExtractCoverage(text); if (llm) return Object.assign({ source: 'invoice+llm' }, llm); } catch (e) { /* ignore */ } }
  return { unparsed: true, source: 'invoice' };
}
async function buildPrepaidSchedule(ctx, quarter, prepaid) {
  const { db, uploadDir } = ctx;
  const flags = [];
  const accts = [
    { code: '150300', label: 'Prepaid Insurance', match: ACCT.prepaidIns, endBal: prepaid.insEnd },
    { code: '150200', label: 'Prepaid Advisory', match: ACCT.prepaidAdv, endBal: prepaid.advEnd },
  ];
  const policies = [];
  for (const a of accts) {
    const rows = glDetail(db, FUND_EID, { to: quarter.end, match: a.match });
    if (!rows.length && Math.abs(a.endBal) < 0.005) continue;
    const amorts = rows.filter((r) => r.credit > 0).map((r) => Object.assign({}, r, { f: parseAmortMemo(r.memo) || parseAmortMemo(r.description) }));
    const additions = rows.filter((r) => r.debit > 0);
    const withF = amorts.find((x) => x.f);
    const premium = withF ? withF.f.premium : null;
    const checkRows = [];
    for (const r of amorts) {
      const booked = r2(r.credit);
      const expected = r.f ? r2(r.f.premium / r.f.basis * r.f.days) : null;
      const ok = expected == null ? null : Math.abs(expected - booked) < 0.01;
      checkRows.push({ date: r.date, num: r.entry_num || r.doc_number, memo: r.memo || r.description, booked, expected, days: r.f ? r.f.days : null, premium: r.f ? r.f.premium : null, basis: r.f ? r.f.basis : null, ok });
      if (r.date >= quarter.qs && r.date <= quarter.end && expected != null && !ok) {
        flags.push({ severity: 'exception', wp: 'Prepaid Expenses', message: a.label + ': quarter amortization booked ' + fmt(booked) + ' but the memo formula computes ' + fmt(expected) });
      }
    }
    let invoice = null;
    for (const add of additions.filter((r) => r.date >= quarter.qs && r.date <= quarter.end)) {
      invoice = await readInvoiceCoverage(uploadDir, db, add.entry_id);
      if (invoice && !invoice.unparsed) {
        if (premium != null && invoice.premium != null && Math.abs(invoice.premium - premium) >= 0.01) flags.push({ severity: 'exception', wp: 'Prepaid Expenses', message: a.label + ': invoice premium ' + fmt(invoice.premium) + ' does not match the amount being amortized ' + fmt(premium) });
      } else if (invoice && invoice.unparsed) {
        flags.push({ severity: 'review', wp: 'Prepaid Expenses', message: a.label + ': invoice attached to ' + (add.entry_num || add.entry_id) + ' but the policy period/premium could not be read — confirm manually' });
      } else {
        flags.push({ severity: 'review', wp: 'Prepaid Expenses', message: a.label + ': new prepaid ' + fmt(add.signed) + ' (' + (add.entry_num || '') + ') has no invoice attached — attach the policy invoice to support the amortization' });
      }
    }
    policies.push({ code: a.code, label: a.label, premium, checkRows, endBal: r2(a.endBal), invoice });
  }
  return { policies, flags };
}

// ── Item-level prepaid schedule (Weaver request): each prepaid item with its
//    premium, coverage period, expected straight-line amortization and remaining
//    balance, so a reviewer can independently recompute amortization and tie the
//    total back to the GL. Items are user-maintained on the Prepaid Register
//    (reuses cla_prepaid_items); when an account has no register items we derive
//    them from the GL (prior-year opening balance + current-year additions) with
//    best-effort coverage from the amortization memo or the attached invoice.
const PP_ACCTS = [
  { code: '150200', label: 'Prepaid Advisory Fees', match: ACCT.prepaidAdv },
  { code: '150300', label: 'Prepaid Insurance', match: ACCT.prepaidIns },
  { code: '150400', label: 'Prepaid Subscription', match: ACCT.prepaidSub },
];
function daysBetween(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 86400000); }
function addDays(dateStr, n) { const [y, m, d] = String(dateStr).split('-').map(Number); const dt = new Date(Date.UTC(y, m - 1, d + n)); return dt.toISOString().slice(0, 10); }
function straightLineAmort(premium, start, end, asOf) {
  if (!(premium > 0) || !isDate(start) || !isDate(end)) return null;
  const total = daysBetween(start, end); if (total <= 0) return null;
  const elapsed = Math.max(0, Math.min(total, daysBetween(start, asOf)));
  return r2(premium * elapsed / total);
}
async function buildPrepaidItems(ctx, eid, quarter) {
  const { db, uploadDir } = ctx;
  const flags = [];
  let reg = [];
  try { reg = db.prepare('SELECT * FROM cla_prepaid_items WHERE entity_id = ? ORDER BY account_code, sort_order, id').all(eid) || []; } catch (e) { reg = []; }
  const regByAcct = new Map();
  for (const it of reg) { const k = String(it.account_code); if (!regByAcct.has(k)) regByAcct.set(k, []); regByAcct.get(k).push(it); }
  const qPriorDay = addDays(quarter.qs, -1);
  const byAcct = [];
  const allItems = [];
  for (const a of PP_ACCTS) {
    const rows = glDetail(db, eid, { to: quarter.end, match: a.match });
    const regItems = regByAcct.get(a.code) || [];
    const preRows = rows.filter((r) => r.date < quarter.ys);
    const begBal = preRows.length ? preRows[preRows.length - 1].balance : 0;
    const endBal = rows.length ? rows[rows.length - 1].balance : 0;
    if (!regItems.length && Math.abs(endBal) < 0.005 && Math.abs(begBal) < 0.005) continue;
    const amortLines = rows.filter((r) => r.credit > 0 && r.date >= quarter.ys);
    const additions = rows.filter((r) => r.debit > 0 && r.date >= quarter.ys);
    let memoF = null;
    for (const r of amortLines) { const f = parseAmortMemo(r.memo) || parseAmortMemo(r.description); if (f) { memoF = f; break; } }
    let items = [];
    if (regItems.length) {
      items = regItems.map((it) => ({
        account_code: a.code, account_label: a.label,
        vendor: it.vendor || '', description: it.description || '',
        date_paid: it.date_paid || '', premium: (it.premium != null ? r2(it.premium) : r2(it.opening_balance || 0)),
        start: it.start_date || '', end: it.end_date || '', source: 'register',
      }));
    } else {
      if (begBal > 0.005) items.push({ account_code: a.code, account_label: a.label, vendor: '', description: 'Prior-year prepaid balance', date_paid: quarter.prior_ye, premium: r2(begBal), start: '', end: '', source: 'opening' });
      for (const add of additions) {
        let start = add.date, end = '';
        let cov = null;
        try { cov = await readInvoiceCoverage(uploadDir, db, add.entry_id); } catch (e) { cov = null; }
        if (cov && cov.start && cov.end) { start = cov.start; end = cov.end; }
        else if (memoF && memoF.basis) { end = addDays(start, memoF.basis); }
        else { end = addDays(start, 365); }
        const desc = String(add.memo || add.description || 'Prepaid addition').replace(/^Bill\s*-\s*/i, '').slice(0, 80);
        items.push({ account_code: a.code, account_label: a.label, vendor: add.vendor || '', description: desc, date_paid: add.date, premium: r2(add.debit), start, end, source: 'gl' });
      }
    }
    let modelRemaining = 0, needCoverage = 0;
    for (const it of items) {
      const asOfEnd = straightLineAmort(it.premium, it.start, it.end, quarter.end);
      const asOfPrior = straightLineAmort(it.premium, it.start, it.end, qPriorDay);
      if (asOfEnd == null) { it.coverageKnown = false; it.accumAmort = null; it.periodAmort = null; it.remaining = null; needCoverage += 1; }
      else { it.coverageKnown = true; it.accumAmort = asOfEnd; it.periodAmort = r2(asOfEnd - (asOfPrior || 0)); it.remaining = r2(it.premium - asOfEnd); }
      modelRemaining = r2(modelRemaining + (it.remaining || 0));
      allItems.push(it);
    }
    const diff = r2(endBal - modelRemaining);
    if (needCoverage > 0) flags.push({ severity: 'review', wp: 'Prepaid Expenses', message: a.label + ': ' + needCoverage + ' item(s) have no coverage period \u2014 enter start/end dates on the Prepaid Register so amortization can be scheduled and the balance verified.' });
    else if (Math.abs(diff) >= 0.01) flags.push({ severity: 'exception', wp: 'Prepaid Expenses', message: a.label + ': modeled remaining prepaid ' + fmt(modelRemaining) + ' does not tie to the GL balance ' + fmt(endBal) + ' (off by ' + fmt(diff) + ') \u2014 the amortization booked differs from the straight-line schedule (e.g. a period not yet amortized).' });
    byAcct.push({ code: a.code, label: a.label, items, begBal: r2(begBal), endBal: r2(endBal), modelRemaining, diff, needCoverage });
  }
  return { items: allItems, byAcct, flags };
}

// ── Accrual open-item schedule (Weaver request 10/2026): every accrual credit to
//    the accrued-liability accounts with the amount since reversed / paid, the
//    reversal date and JE, and the net amount still outstanding at quarter end.
//    Debits are matched to the accrual they relieve in this order: (1) a reversing
//    JE that points at the accrual (journal_entries.reverses_entry_id), (2) same
//    amount and vendor, (3) same amount, (4) oldest open accrual first (FIFO).
//    Debits left over after matching are listed as unapplied, so the schedule's
//    net outstanding equals the GL balance by construction.
const ACCR_GROUPS = [
  { code: '210600', prefix: '2106', label: 'Management fees payable', bsLine: 'Management fees payable', match: (c) => ACCT.mgmtPay(c) },
  { code: '210000', prefix: '2100', label: 'Accrued expenses', bsLine: 'within Accounts payable & accrued expenses', match: (c) => ACCT.accrued(c) },
  { code: '211000', prefix: '2110', label: 'Due to affiliates', bsLine: 'Due to affiliates', match: (c) => /^2110/.test(c) },
];
function buildAccrualItems(db, eid, quarter, bEnd) {
  const normV = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const revOf = new Map();
  for (const r of db.prepare('SELECT id, reverses_entry_id FROM journal_entries WHERE entity_id = ? AND reverses_entry_id IS NOT NULL').all(eid)) revOf.set(r.id, r.reverses_entry_id);
  const items = [], groups = [];
  for (const g of ACCR_GROUPS) {
    const rows = glDetail(db, eid, { to: quarter.end, match: g.match });
    const credits = rows.filter((r) => r.credit > 0).map((r) => ({ ...r, open: r2(r.credit), allocs: [] }));
    const unapplied = [];
    for (const d of rows.filter((r) => r.debit > 0)) {
      let rem = r2(d.debit);
      const take = (c) => { if (!c || rem <= 0.004) return; const a = r2(Math.min(c.open, rem)); if (a <= 0.004) return; c.open = r2(c.open - a); rem = r2(rem - a); c.allocs.push({ amount: a, date: d.date, je: d.entry_num }); };
      const open = (c) => c.open > 0.004;
      const src = revOf.get(d.entry_id);
      if (src) take(credits.find((c) => open(c) && c.entry_id === src && c.account_code === d.account_code));
      if (rem > 0.004 && d.vendor) take(credits.find((c) => open(c) && c.date <= d.date && Math.abs(c.open - rem) < 0.005 && normV(c.vendor) === normV(d.vendor)));
      if (rem > 0.004) take(credits.find((c) => open(c) && c.date <= d.date && Math.abs(c.open - rem) < 0.005));
      for (const c of credits) { if (rem <= 0.004) break; if (open(c) && c.date <= d.date) take(c); }
      for (const c of credits) { if (rem <= 0.004) break; if (open(c)) take(c); }
      if (rem > 0.004) unapplied.push({ ...d, amount: rem });
    }
    const shown = credits.filter((c) => c.date >= quarter.qs || c.open > 0.004 || c.allocs.some((a) => a.date >= quarter.qs));
    for (const c of shown) {
      const reversed = r2(c.credit - c.open);
      const last = c.allocs.length ? c.allocs[c.allocs.length - 1] : null;
      items.push({
        kind: 'accrual', group: g.code, date: c.date, vendor: c.vendor, entry_num: c.entry_num, doc_number: c.doc_number,
        description: c.description || c.memo, account_code: c.account_code, account_name: c.account_name,
        accrued: r2(c.credit), reversed, reversal_date: last ? last.date : '', reversal_je: [...new Set(c.allocs.map((a) => a.je))].join(', '), net: c.open,
      });
    }
    for (const u of unapplied) {
      items.push({
        kind: 'unapplied', group: g.code, date: u.date, vendor: u.vendor, entry_num: u.entry_num, doc_number: u.doc_number,
        description: 'Debit not matched to a specific accrual — ' + (u.description || u.memo), account_code: u.account_code, account_name: u.account_name,
        accrued: 0, reversed: u.amount, reversal_date: u.date, reversal_je: String(u.entry_num || ''), net: r2(-u.amount),
      });
    }
    const schedule = r2(credits.reduce((a, c) => a + c.open, 0) - unapplied.reduce((a, u) => a + u.amount, 0));
    groups.push({ ...g, match: undefined, schedule, gl: sumWhere(bEnd, g.match), unappliedCount: unapplied.length });
  }
  const order = new Map(ACCR_GROUPS.map((g, i) => [g.code, i]));
  items.sort((a, b) => (order.get(a.group) - order.get(b.group)) || String(a.date).localeCompare(String(b.date)) || (Number(a.entry_num) || 0) - (Number(b.entry_num) || 0));
  return { items, groups };
}

// ── Accrual: mgmt-fee ITD roll, subsequent cash disbursement, flux analysis ──
function buildMgmtItd(db, quarter) {
  const rows = glDetail(db, FUND_EID, { to: quarter.end, match: ACCT.mgmtPay });
  return { rows, end: rows.length ? rows[rows.length - 1].balance : 0 };
}
function buildSubsequentDisb(db, quarter) {
  const from = nextDay(quarter.end), to = addMonths(quarter.end, 3);
  const rows = glDetail(db, FUND_EID, { from, to, match: (c) => ACCT.accrued(c) || ACCT.mgmtPay(c) || /^2110/.test(c) }).filter((r) => r.debit > 0);
  return { from, to, rows, posted: rows.length > 0 };
}
function pnlByAccount(db, eid, from, to) {
  const rows = db.prepare(
    'SELECT jl.account_code code, a.name name, a.type type, COALESCE(SUM(jl.debit),0) td, COALESCE(SUM(jl.credit),0) tc '
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + 'LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code '
    + "WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? "
    + "AND a.type IN ('Revenue','Income','Other Income','Expense','Other Expense','Cost of Goods Sold') "
    + 'GROUP BY jl.account_code'
  ).all(eid, from, to);
  const m = new Map();
  for (const r of rows) {
    const rev = /Revenue|Income/i.test(String(r.type));
    const amt = rev ? r2((r.tc || 0) - (r.td || 0)) : r2((r.td || 0) - (r.tc || 0));
    m.set(String(r.code), { code: String(r.code), name: r.name || '', type: r.type || '', amt });
  }
  return m;
}
function buildFlux(db, quarter) {
  const curFrom = quarter.qs, curTo = quarter.end;
  const priFrom = addMonths(quarter.qs, -3), priTo = nextDay(addMonths(quarter.qs, 0));
  const priEnd = (function () { const d = new Date(Date.UTC(...quarter.qs.split('-').map((x, i) => i === 2 ? Number(x) - 1 : Number(x) - (i === 1 ? 1 : 0)))); return d.toISOString().slice(0, 10); })();
  const cur = pnlByAccount(db, FUND_EID, curFrom, curTo);
  const pri = pnlByAccount(db, FUND_EID, priFrom, priEnd);
  const codes = new Set([...cur.keys(), ...pri.keys()]);
  const rows = [], flags = [];
  for (const c of codes) {
    const cv = cur.get(c) || { name: (pri.get(c) || {}).name || '', amt: 0, type: (pri.get(c) || {}).type || '' };
    const pv = pri.get(c) || { amt: 0 };
    const diff = r2(cv.amt - pv.amt);
    const pct = Math.abs(pv.amt) > 0.005 ? diff / Math.abs(pv.amt) : null;
    if (Math.abs(cv.amt) < 0.005 && Math.abs(pv.amt) < 0.005) continue;
    rows.push({ code: c, name: cv.name, prior: r2(pv.amt), cur: r2(cv.amt), diff, pct });
    if (Math.abs(diff) >= 50000) flags.push({ severity: 'review', wp: 'Flux Analysis', message: 'Large quarter-over-quarter change in ' + (cv.name || c) + ': ' + fmt(pv.amt) + ' → ' + fmt(cv.amt) + ' (' + fmt(diff) + ') — confirm the driver' });
  }
  rows.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
  return { priFrom, priTo: priEnd, curFrom, curTo, rows, flags };
}

// ── Contribution receivable (120020) — by investor ──
function buildContribRecv(db, quarter, bEnd) {
  const rows = glDetail(db, FUND_EID, { to: quarter.end, match: (c) => /^12002/.test(c) });
  const balance = sumWhere(bEnd, (c) => /^12002/.test(c));
  const byInv = new Map();
  for (const r of rows) { const k = r.class_name || '(untagged)'; byInv.set(k, r2((byInv.get(k) || 0) + r.signed)); }
  const investors = [...byInv.entries()].map(([name, amt]) => ({ name, amt: r2(amt) })).filter((x) => Math.abs(x.amt) >= 0.005).sort((a, b) => Math.abs(b.amt) - Math.abs(a.amt));
  return { rows, balance, investors };
}

// ── Distributions payable / due to management company (by investor) ──
function buildDistributions(db, quarter, bEnd, bBeg) {
  const groups = [
    { code: '230100', label: 'Distributions Payable (Due to Members)', match: (c) => /^2301/.test(c) },
    { code: '210800', label: 'Due to Management Company', match: (c) => /^2108/.test(c) },
  ];
  const out = [];
  for (const g of groups) {
    const rows = glDetail(db, FUND_EID, { to: quarter.end, match: g.match });
    const balance = sumWhere(bEnd, g.match);
    const byInv = new Map();
    for (const r of rows) { const k = r.class_name || '(untagged)'; byInv.set(k, r2((byInv.get(k) || 0) + r.signed)); }
    const investors = [...byInv.entries()].map(([name, amt]) => ({ name, amt: r2(amt) })).filter((x) => Math.abs(x.amt) >= 0.005).sort((a, b) => Math.abs(b.amt) - Math.abs(a.amt));
    const begin = bBeg ? sumWhere(bBeg, g.match) : 0;
    out.push({ code: g.code, label: g.label, rows, balance: r2(balance), begin: r2(begin), investors });
  }
  return out;
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
// Formula cell with its cached result (so the file shows values before Excel recalculates).
const fx = (ws, ref, formula, result, o = {}) => { const c = ws.getCell(ref); c.value = { formula, result: typeof result === 'number' ? r2(result) : result }; c.numFmt = MONEY; c.font = F(o); return c; };
// Real Excel date cell (so coverage-day and aging formulas can do date math).
const setDate = (ws, ref, s, o = {}) => { const c = ws.getCell(ref); if (isDate(s)) { const [y, m, dd] = String(s).split('-').map(Number); c.value = new Date(Date.UTC(y, m - 1, dd)); c.numFmt = 'm/d/yyyy'; } else c.value = s || ''; c.font = F(o); return c; };

function buildWorkbook(data) {
  const q = data.quarter, en = data.entity_name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };

  // ── Summary / index ──────────────────────────────────────────────────────
  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  su.getColumn(1).width = 34; su.getColumn(2).width = 40; su.getColumn(3).width = 16; su.getColumn(4).width = 18;
  titleBlock(su, en, 'Other Workpapers — Balance Sheet Account Support', 'As of ' + spellDate(q.end));
  hdrRow(su, 5, ['Workpaper', 'Balance Sheet line', 'GL account(s)', 'CL balance'], [34, 40, 16, 18]);
  { const fl0 = data.flags || []; su.getCell('A4').value = fl0.length ? ('⚠ ' + fl0.length + ' item(s) require attention — see Exceptions below') : '✓ All balances tie'; su.getCell('A4').font = F({ bold: true, color: { argb: fl0.length ? 'FFC00000' : 'FF008000' } }); su.mergeCells('A4:D4'); }
  const idx = [
    ['Due Fr (To) Port Co', 'Due from (to) portfolio investments (net)', '101100 / 211100', data.ties.due_net],
    ['Interest Receivable', 'Interest receivable', '120010', data.ties.interest_receivable],
    ['Prepaid Expenses', 'Prepaid advisory / insurance / subscription', '150200 / 150300 / 150400', r2(data.ties.prepaid_insurance + data.ties.prepaid_advisory + data.ties.prepaid_subscription)],
    ['Other Assets', 'Other assets', '180100', data.ties.other_assets],
    ['AP Recon', 'Accounts payable and accrued expenses', '202000 + 210000', r2(data.ties.accounts_payable + data.ties.accrued_expenses)],
    ['Accrual & Sub Cash Disbursement', 'Management fees payable', '210600', data.ties.management_fees_payable],
    ['Distributions Payable', 'Due to members', '230100', data.ties.distributions_payable],
    ['Contribution Receivable', 'Capital contributions receivable', '120020', data.ties.contribution_receivable],
    ['Distributions Payable', 'Due to management company', '210800', data.ties.due_to_mgmt],
    ['Accrual & Sub Cash Disbursement', 'Due to affiliates', '211000', data.ties.due_to_affiliates],
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
      er++;
      const _m = String(f.message || '');
      su.getCell('A' + er).value = _m; su.getCell('A' + er).font = F({ color: { argb: 'FF333333' } });
      su.getCell('A' + er).alignment = { wrapText: true, vertical: 'top' };
      su.mergeCells('A' + er + ':D' + er);
      su.getRow(er).height = Math.min(220, 14 * Math.max(1, Math.ceil(_m.length / 108)) + 4);
      er += 2;
    }
  }

  // ── 1. Due Fr (To) Port Co ────────────────────────────────────────────────
  // Beginning balances and JE activity are GL inputs; the ending row SUMs each
  // property column, and the intercompany recon links to that ending row.
  const du = wb.addWorksheet('Due Fr (To) Port Co', { views: [{ showGridLines: false }] });
  du.getColumn(2).width = 40; du.getColumn(3).width = 14; [5, 7, 9, 11].forEach((c) => (du.getColumn(c).width = 14)); du.getColumn(12).width = 54;
  titleBlock(du, en, 'Due From/To Port Co reconciliation', spellDate(q.end));
  const P = data.due.PROPS;
  du.getCell('C5').value = 'JE#'; du.getCell('C5').font = F({ bold: true });
  du.getCell('L5').value = 'Journal entry description'; du.getCell('L5').font = F({ bold: true });
  P.forEach((p, i) => { const c = du.getCell(String.fromCharCode(69 + i * 2) + '5'); c.value = p; c.font = F({ bold: true }); c.alignment = { horizontal: 'center' }; });
  const propCol = (i) => String.fromCharCode(69 + i * 2); // E,G,I,K
  let R = 6; const dueFirst = R;
  du.getCell('B' + R).value = 'Due From (To) Port Co at ' + short(q.prior_ye); du.getCell('B' + R).font = F();
  P.forEach((p, i) => setMoney(du, propCol(i) + R, data.due.beg[p]));
  R++;
  for (const j of data.due.jes) {
    du.getCell('C' + R).value = j.je; du.getCell('C' + R).font = F();
    P.forEach((p, i) => setMoney(du, propCol(i) + R, j[p] || 0));
    du.getCell('L' + R).value = j.desc || ''; du.getCell('L' + R).font = F(); du.getCell('L' + R).alignment = { wrapText: true, vertical: 'top' };
    R++;
  }
  du.getCell('B' + R).value = 'Due From (To) Port Co at ' + short(q.end); du.getCell('B' + R).font = F({ bold: true });
  P.forEach((p, i) => { const col = propCol(i); const c = fx(du, col + R, 'SUM(' + col + dueFirst + ':' + col + (R - 1) + ')', data.due.end[p], { bold: true }); c.border = { top: THIN }; });
  const dueEndRow = R; // ending row — Summary links to SUM(E:K) here
  du.getCell('B' + (R + 2)).value = 'Net Due From (To) Portfolio Company ties to the Balance Sheet: due-from asset (101100) less due-to liability (211100). Ending row = beginning balance + each JE (SUM formula).';
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
        const pi = P.indexOf(row.prop);
        if (pi >= 0) fx(du, 'E' + rr, propCol(pi) + dueEndRow, row.fundDue); else setMoney(du, 'E' + rr, row.fundDue);
        setMoney(du, 'G' + rr, row.cpNet);
        fx(du, 'I' + rr, 'E' + rr + '-G' + rr, row.diff);
        const st = row.status === 'matched' ? 'Tied' : (row.status === 'one_sided' ? 'One-sided' : 'Mismatch');
        const cc = du.getCell('K' + rr); cc.value = st; cc.font = F({ bold: row.status !== 'matched', color: { argb: row.status === 'matched' ? 'FF008000' : 'FFC00000' } });
        rr += 1;
        for (const tr of (row.trace || [])) { const _n = '↳ ' + tr.note; du.getCell('C' + rr).value = _n; du.getCell('C' + rr).font = SMALLI; du.getCell('C' + rr).alignment = { wrapText: true, vertical: 'top' }; du.mergeCells('C' + rr + ':K' + rr); du.getRow(rr).height = Math.min(200, 14 * Math.max(1, Math.ceil(_n.length / 95)) + 4); rr += 1; }
      }
    }
  }

  // ── 2. Interest Receivable ────────────────────────────────────────────────
  const intTotalRow = detailSheet(wb, 'Interest Receivable', en, 'Interest Receivable (120010)', q, data.interest.rows, data.interest.balance);

  // ── 3. Prepaid Expenses ───────────────────────────────────────────────────
  // Roll-forward: beginning + YTD activity (GL inputs) = ending (formula); the
  // Total column SUMs the three accounts. Item schedule below recomputes the
  // straight-line amortization from premium and coverage dates with formulas.
  const pp = wb.addWorksheet('Prepaid Expenses', { views: [{ showGridLines: false }] });
  titleBlock(pp, en, 'Schedule for Prepaid Expenses', 'As of ' + short(q.end));
  hdrRow(pp, 5, ['', 'Prepaid Advisory (150200)', 'Prepaid Insurance (150300)', 'Prepaid Subscription (150400)', 'Total'], null);
  pp.getRow(5).height = 30;
  pp.getCell('A6').value = 'Balance at ' + short(q.prior_ye); pp.getCell('A6').font = F();
  setMoney(pp, 'B6', data.prepaid.advBeg); setMoney(pp, 'C6', data.prepaid.insBeg); setMoney(pp, 'D6', data.prepaid.subBeg);
  pp.getCell('A7').value = 'Net additions / (amortization), year to date'; pp.getCell('A7').font = F();
  setMoney(pp, 'B7', data.prepaid.advAmort); setMoney(pp, 'C7', data.prepaid.insAmort); setMoney(pp, 'D7', data.prepaid.subAmort);
  pp.getCell('A8').value = 'Balance at ' + short(q.end); pp.getCell('A8').font = F({ bold: true });
  const ppEnd = { B: data.prepaid.advEnd, C: data.prepaid.insEnd, D: data.prepaid.subEnd };
  ['B', 'C', 'D'].forEach((col) => { fx(pp, col + '8', col + '6+' + col + '7', ppEnd[col], { bold: true }).border = { top: THIN }; });
  fx(pp, 'E6', 'SUM(B6:D6)', r2(data.prepaid.advBeg + data.prepaid.insBeg + data.prepaid.subBeg));
  fx(pp, 'E7', 'SUM(B7:D7)', r2(data.prepaid.advAmort + data.prepaid.insAmort + data.prepaid.subAmort));
  fx(pp, 'E8', 'SUM(B8:D8)', r2(data.prepaid.advEnd + data.prepaid.insEnd + data.prepaid.subEnd), { bold: true }).border = { top: THIN };
  pp.getCell('A10').value = 'Tied to Balance Sheet. Item-level detail with coverage periods and remaining balances is below.'; pp.getCell('A10').font = SMALLI;

  // ── 4. Other Assets (grouped detail) ──────────────────────────────────────
  const oa = wb.addWorksheet('Other Assets', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  titleBlock(oa, en, 'Other Assets (180100) — transaction detail', 'As of ' + short(q.end));
  hdrRow(oa, 5, ['Group', 'Date', 'Type', 'Num', 'Description', 'Amount', 'Balance'], [22, 12, 14, 16, 52, 16, 16]);
  let orow = 6; const oaGroupTotals = [];
  for (const g of data.otherAssets.groups) {
    oa.getCell('A' + orow).value = g.name; oa.getCell('A' + orow).font = F({ bold: true }); orow++;
    const gFirst = orow; let run = 0;
    for (const r of g.rows) {
      run = r2(run + r.signed);
      setDate(oa, 'B' + orow, r.date);
      oa.getCell('C' + orow).value = ''; // type
      oa.getCell('D' + orow).value = r.entry_num || r.doc_number; oa.getCell('D' + orow).font = F();
      oa.getCell('E' + orow).value = r.description || r.memo; oa.getCell('E' + orow).font = F();
      setMoney(oa, 'F' + orow, r.signed);
      fx(oa, 'G' + orow, orow === gFirst ? 'F' + orow : 'G' + (orow - 1) + '+F' + orow, run);
      orow++;
    }
    oa.getCell('E' + orow).value = 'Total for ' + g.name; oa.getCell('E' + orow).font = F({ bold: true });
    const c = orow > gFirst ? fx(oa, 'F' + orow, 'SUM(F' + gFirst + ':F' + (orow - 1) + ')', g.total, { bold: true }) : setMoney(oa, 'F' + orow, 0, { bold: true });
    c.border = { top: THIN }; oaGroupTotals.push(orow); orow += 1;
  }
  oa.getCell('E' + orow).value = 'TOTAL Other Assets'; oa.getCell('E' + orow).font = F({ bold: true });
  const oc = oaGroupTotals.length ? fx(oa, 'F' + orow, 'SUM(' + oaGroupTotals.map((r) => 'F' + r).join(',') + ')', data.otherAssets.total, { bold: true }) : setMoney(oa, 'F' + orow, 0, { bold: true });
  oc.border = { top: THIN, bottom: THIN };
  const oaTotalRow = orow; // Summary links to F here
  orow += 2;
  const oaBeg = orow; oa.getCell('D' + orow).value = 'Beginning balance at ' + short(q.prior_ye); oa.getCell('D' + orow).font = F(); setMoney(oa, 'F' + orow, data.otherAssets.begin); orow += 1;
  const oaAct = orow; oa.getCell('D' + orow).value = 'Net activity in ' + q.label; oa.getCell('D' + orow).font = F(); orow += 1;
  const oaEnd = orow; oa.getCell('D' + orow).value = 'Ending balance at ' + short(q.end); oa.getCell('D' + orow).font = F({ bold: true });
  fx(oa, 'F' + oaEnd, 'F' + oaTotalRow, data.otherAssets.total, { bold: true }).border = { top: THIN };
  fx(oa, 'F' + oaAct, 'F' + oaEnd + '-F' + oaBeg, r2(data.otherAssets.total - data.otherAssets.begin));
  orow += 1;
  const oaGl = orow; oa.getCell('D' + orow).value = 'Balance per general ledger (180100)'; oa.getCell('D' + orow).font = F(); setMoney(oa, 'F' + orow, data.ties.other_assets); orow += 1;
  oa.getCell('D' + orow).value = 'Difference (should be zero)'; oa.getCell('D' + orow).font = F({ bold: true }); fx(oa, 'F' + orow, 'F' + oaEnd + '-F' + oaGl, r2(data.otherAssets.total - data.ties.other_assets), { bold: true }).border = { top: THIN }; orow += 2;
  oa.getCell('A' + orow).value = 'Running balance and totals are formulas over the GL transaction lines; the ending balance agrees to the Balance Sheet other assets line (180100).'; oa.getCell('A' + orow).font = SMALLI; oa.mergeCells('A' + orow + ':G' + orow);

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
  // Per Bill.com links to the Bill.com AP Detail tab; Per GL links to the CL AP
  // Detail tab, so each figure traces to its supporting schedule.
  const bcN = (AP.billcomLines || []).length, clN = (AP.openBills || []).length;
  const bcLastRow = 6 + bcN, bcTotRow = 7 + bcN, clLastRow = 6 + clN, clTotRow = 7 + clN;
  const BC = "'Bill.com AP Detail'!", CL = "'CL AP Detail'!";
  const money = (c) => { c.numFmt = MONEY; return c; };
  for (const v of (AP.byVendor || [])) {
    ap.getCell('A' + ar).value = v.vendor; ap.getCell('A' + ar).font = F();
    if (AP.billcomSource !== 'none' && bcN > 0) { const c = ap.getCell('B' + ar); c.value = { formula: 'SUMIF(' + BC + '$A$7:$A$' + bcLastRow + ',A' + ar + ',' + BC + '$D$7:$D$' + bcLastRow + ')', result: v.billcom }; money(c).font = F(); }
    else if (AP.billcomSource !== 'none') setMoney(ap, 'B' + ar, v.billcom);
    if (clN > 0) { const c = ap.getCell('C' + ar); c.value = { formula: 'SUMIF(' + CL + '$A$7:$A$' + clLastRow + ',A' + ar + ',' + CL + '$E$7:$E$' + clLastRow + ')', result: v.gl }; money(c).font = F(); }
    else setMoney(ap, 'C' + ar, v.gl);
    if (AP.billcomSource !== 'none') { const c = ap.getCell('D' + ar); c.value = { formula: 'B' + ar + '-C' + ar, result: v.diff }; money(c).font = F(); }
    ar++;
  }
  ap.getCell('A' + ar).value = 'Total accounts payable'; ap.getCell('A' + ar).font = F({ bold: true });
  if (AP.billcomSource !== 'none') { const c = ap.getCell('B' + ar); c.value = (bcN > 0 ? { formula: BC + 'D' + bcTotRow, result: AP.billcomTotal } : AP.billcomTotal); money(c).font = F({ bold: true }); c.border = { top: THIN }; }
  { const c = ap.getCell('C' + ar); c.value = (clN > 0 ? { formula: CL + 'E' + clTotRow, result: AP.openTotal } : AP.openTotal); money(c).font = F({ bold: true }); c.border = { top: THIN }; }
  { const c = ap.getCell('D' + ar); if (AP.billcomSource !== 'none') c.value = { formula: 'B' + ar + '-C' + ar, result: r2((AP.billcomTotal || 0) - AP.openTotal) }; money(c).font = F({ bold: true }); c.border = { top: THIN }; }
  const apTotalRow = ar; // Summary links to C here (Per GL open invoices = 202000)
  ar += 2;
  // The two rows below are linked to the AP GL Detail / CL AP Detail tabs once
  // those tabs are built (cells filled in after section 5c).
  ap.getCell('A' + ar).value = 'Accounts payable per general ledger (202000)'; ap.getCell('A' + ar).font = F();
  const apGlRow = ar; ar++;
  ap.getCell('A' + ar).value = 'Open invoices per A/P aging (ties to CL A/P Aging report)'; ap.getCell('A' + ar).font = F();
  const apOpenRow = ar; ar++;
  ap.getCell('A' + ar).value = 'Difference'; ap.getCell('A' + ar).font = F({ bold: true });
  fx(ap, 'C' + ar, 'C' + apGlRow + '-C' + apOpenRow, r2(AP.gl - AP.openTotal), { bold: true }).border = { top: THIN };
  ar += 2;
  ap.getCell('A' + ar).value = AP.billcomSource === 'live'
    ? ('Per Bill.com = the Bill.com open-invoice report pulled live from the Bill.com API as of ' + (AP.billcomAsOf || short(q.end)) + ' (bills less payments applied by that date). Per GL = open invoices on account 202000 per the A/P aging (payments netted against their bills), which ties to the Balance Sheet and the CL A/P Aging report.')
    : AP.billcomSource === 'aging'
    ? ('Per Bill.com = the uploaded Bill.com A/P aging as of ' + (AP.billcomAsOf || short(q.end)) + '. Per GL = open invoices on account 202000 per the A/P aging (payments netted against their bills), which ties to the Balance Sheet and the CL A/P Aging report.')
    : ('Per Bill.com is blank because no Bill.com A/P Detail report has been uploaded, so the GL A/P is NOT yet independently verified. Export the Bill.com A/P Detail (Open Items) report as of ' + short(q.end) + ' and upload it, then regenerate. See the Exceptions block on the Summary tab.');
  ap.getCell('A' + ar).font = SMALLI; ap.mergeCells('A' + ar + ':D' + ar);

  // 5a. Bill.com A/P Detail — open invoices.
  const bc = wb.addWorksheet('Bill.com AP Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  titleBlock(bc, en, 'Bill.com A/P Detail — Open Invoices', 'As of ' + short(AP.billcomAsOf || q.end));
  hdrRow(bc, 6, ['Vendor', 'Invoice #', 'Bill Date', 'Amount'], [40, 18, 14, 16]);
  let bcr = 7; const bcFirst = bcr;
  for (const b of (AP.billcomLines || [])) {
    bc.getCell('A' + bcr).value = b.vendor; bc.getCell('A' + bcr).font = F();
    bc.getCell('B' + bcr).value = b.invoice || ''; bc.getCell('B' + bcr).font = F();
    setDate(bc, 'C' + bcr, b.date || '');
    setMoney(bc, 'D' + bcr, b.amount); bcr++;
  }
  const bcLast = bcr - 1;
  if (AP.billcomSource === 'none') {
    bc.getCell('A' + bcr).value = 'No Bill.com A/P Detail report uploaded. Export the Bill.com A/P Detail (Open Items) report as of ' + short(AP.billcomAsOf || q.end) + ' and upload it (Bill.com settings → A/P aging) to complete this reconciliation.'; bc.getCell('A' + bcr).font = SMALLI; bc.mergeCells('A' + bcr + ':D' + bcr);
  } else {
    bc.getCell('A' + bcr).value = 'Total open invoices per Bill.com'; bc.getCell('A' + bcr).font = F({ bold: true });
    { const c = bcLast >= bcFirst ? fx(bc, 'D' + bcr, 'SUM(D' + bcFirst + ':D' + bcLast + ')', AP.billcomTotal, { bold: true }) : setMoney(bc, 'D' + bcr, AP.billcomTotal, { bold: true }); c.border = { top: THIN }; }
    bc.getCell('A' + (bcr + 2)).value = AP.billcomSource === 'live'
      ? ('Source: pulled live from the Bill.com API as of ' + short(AP.billcomAsOf || q.end) + ' (bills less payments applied by that date). Ties to account 202000 and to the CL AP Detail tab.')
      : 'Source: uploaded Bill.com A/P aging detail. Ties to account 202000 and to the CL AP Detail tab.';
    bc.getCell('A' + (bcr + 2)).font = SMALLI; bc.mergeCells('A' + (bcr + 2) + ':D' + (bcr + 2));
  }

  // 5b. CL A/P Detail — open invoices per GL.
  const cld = wb.addWorksheet('CL AP Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  titleBlock(cld, en, 'Accounts Payable Detail per GL — Open Invoices', 'As of ' + short(q.end));
  hdrRow(cld, 6, ['Vendor', 'Invoice #', 'Bill Date', 'JE #', 'Amount'], [40, 18, 14, 10, 16]);
  let clr = 7; const clFirst = clr;
  for (const b of (AP.openBills || [])) {
    cld.getCell('A' + clr).value = b.vendor; cld.getCell('A' + clr).font = F();
    cld.getCell('B' + clr).value = b.invoice || ''; cld.getCell('B' + clr).font = F();
    setDate(cld, 'C' + clr, b.date || '');
    cld.getCell('D' + clr).value = b.num || ''; cld.getCell('D' + clr).font = F();
    setMoney(cld, 'E' + clr, b.amount); clr++;
  }
  const clLast = clr - 1;
  cld.getCell('A' + clr).value = 'Total open invoices per GL'; cld.getCell('A' + clr).font = F({ bold: true });
  { const c = clLast >= clFirst ? fx(cld, 'E' + clr, 'SUM(E' + clFirst + ':E' + clLast + ')', AP.openTotal, { bold: true }) : setMoney(cld, 'E' + clr, AP.openTotal, { bold: true }); c.border = { top: THIN }; }
  const clTotalRow = clr;
  cld.getCell('A' + (clr + 2)).value = 'Open invoices per the A/P aging as of ' + short(q.end) + ' (each bill net of the Bill.com payments applied against it). Ties to the 202000 A/P balance, the CL A/P Aging report, and the Bill.com A/P Detail.';
  cld.getCell('A' + (clr + 2)).font = SMALLI; cld.mergeCells('A' + (clr + 2) + ':E' + (clr + 2));

  // 5c. AP GL Detail — 202000 detail with offset-letter tagging + bottom recon.
  // Running balance = prior balance + credit − debit (formula); totals are SUMs;
  // the reconciliation block links to those totals.
  const gld = wb.addWorksheet('AP GL Detail', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  titleBlock(gld, en, 'Accounts Payable (202000) — GL Detail with Offset', q.qs + ' to ' + short(q.end));
  hdrRow(gld, 6, ['Date', 'Num', 'Vendor', 'Offset Account', 'Description', 'Debit', 'Credit', 'Balance', 'Open'], [11, 8, 26, 26, 46, 14, 14, 15, 8]);
  let gr = 7;
  const gBegRow = gr;
  gld.getCell('A' + gr).value = 'Beginning balance ' + short(q.prior_ye); gld.getCell('A' + gr).font = F({ bold: true }); gld.mergeCells('A' + gr + ':E' + gr);
  setMoney(gld, 'H' + gr, AP.begin, { bold: true });
  { const c = gld.getCell('I' + gr); c.value = ''; c.alignment = { horizontal: 'center' }; }
  gr++;
  const gFirst = gr; let gRun = r2(AP.begin);
  for (const x of AP.glRows) {
    setDate(gld, 'A' + gr, x.date);
    gld.getCell('B' + gr).value = x.num; gld.getCell('B' + gr).font = F();
    gld.getCell('C' + gr).value = x.vendor; gld.getCell('C' + gr).font = F();
    gld.getCell('D' + gr).value = x.offset; gld.getCell('D' + gr).font = F();
    gld.getCell('E' + gr).value = x.memo; gld.getCell('E' + gr).font = F();
    if (x.debit) setMoney(gld, 'F' + gr, x.debit);
    if (x.credit) setMoney(gld, 'G' + gr, x.credit);
    gRun = r2(gRun + (x.credit || 0) - (x.debit || 0));
    fx(gld, 'H' + gr, 'H' + (gr - 1) + '+G' + gr + '-F' + gr, gRun);
    { const c = gld.getCell('I' + gr); c.value = x.letter || ''; c.font = (x.letter === 'X') ? F({ bold: true }) : F(); c.alignment = { horizontal: 'center' }; }
    gr++;
  }
  const gLast = gr - 1, gTot = gr;
  gld.getCell('E' + gr).value = 'TOTAL'; gld.getCell('E' + gr).font = F({ bold: true });
  (gLast >= gFirst ? fx(gld, 'F' + gr, 'SUM(F' + gFirst + ':F' + gLast + ')', AP.totalDebit, { bold: true }) : setMoney(gld, 'F' + gr, 0, { bold: true })).border = { top: THIN };
  (gLast >= gFirst ? fx(gld, 'G' + gr, 'SUM(G' + gFirst + ':G' + gLast + ')', AP.totalCredit, { bold: true }) : setMoney(gld, 'G' + gr, 0, { bold: true })).border = { top: THIN };
  fx(gld, 'H' + gr, 'H' + gBegRow + '+G' + gr + '-F' + gr, AP.ending, { bold: true }).border = { top: THIN };
  gr += 2;
  gld.getCell('E' + gr).value = 'Reconciliation'; gld.getCell('E' + gr).font = F({ bold: true }); gr++;
  const rBeg = gr;  gld.getCell('E' + gr).value = 'Beginning A/P balance (tagged X)'; gld.getCell('E' + gr).font = F(); fx(gld, 'H' + gr, 'H' + gBegRow, AP.begin); gr++;
  const rCr = gr;   gld.getCell('E' + gr).value = 'Add: credits (bills booked) in period'; gld.getCell('E' + gr).font = F(); fx(gld, 'H' + gr, 'G' + gTot, AP.totalCredit); gr++;
  const rDr = gr;   gld.getCell('E' + gr).value = 'Less: debits (payments / reversals) in period'; gld.getCell('E' + gr).font = F(); fx(gld, 'H' + gr, '-F' + gTot, r2(-AP.totalDebit)); gr++;
  const rEnd = gr;  gld.getCell('E' + gr).value = 'Ending A/P balance per GL detail'; gld.getCell('E' + gr).font = F({ bold: true });
  fx(gld, 'H' + gr, 'H' + rBeg + '+H' + rCr + '+H' + rDr, AP.ending, { bold: true }).border = { top: THIN };
  gr += 1;
  const rTb = gr;   gld.getCell('E' + gr).value = 'Balance per trial balance (202000)'; gld.getCell('E' + gr).font = F(); setMoney(gld, 'H' + gr, AP.gl); gr++;
  gld.getCell('E' + gr).value = 'Difference (should be zero)'; gld.getCell('E' + gr).font = F(); fx(gld, 'H' + gr, 'H' + rEnd + '-H' + rTb, r2(AP.ending - AP.gl)).border = { top: THIN };
  gr += 2;
  const rOpen = gr; gld.getCell('E' + gr).value = 'Open invoices per A/P aging (CL AP Detail tab)'; gld.getCell('E' + gr).font = F(); fx(gld, 'H' + gr, CL + 'E' + clTotalRow, AP.openTotal); gr++;
  gld.getCell('E' + gr).value = 'Difference vs ending balance (should be zero)'; gld.getCell('E' + gr).font = F({ bold: true });
  fx(gld, 'H' + gr, 'H' + rEnd + '-H' + rOpen, r2(AP.ending - AP.openTotal), { bold: true }).border = { top: THIN };
  gr += 2;
  gld.getCell('A' + gr).value = 'Lines marked O are the open invoices per the A/P aging as of ' + short(q.end) + '. The aging nets each Bill.com payment against the specific bill it settled and carries unapplied payments forward, so open A/P (' + fmt(AP.openTotal) + ') equals the ending 202000 balance and ties to the CL A/P Aging report and Bill.com. Full open-invoice detail is on the CL A/P Detail tab.';
  gld.getCell('A' + gr).font = SMALLI; gld.mergeCells('A' + gr + ':I' + gr);
  // Back-fill the AP Recon tie rows now that the supporting rows are known.
  fx(ap, 'C' + apGlRow, "'AP GL Detail'!H" + rEnd, AP.ending);
  fx(ap, 'C' + apOpenRow, CL + 'E' + clTotalRow, AP.openTotal);

  // ── 6. Accrual & Subsequent Cash Disbursement ─────────────────────────────
  // Open-item accrual schedule: each accrual with the amount reversed / paid,
  // the reversal date and JE, and the net outstanding (formula). The net column
  // is SUMIF'd by account and tied to each account's GL balance.
  const ac = wb.addWorksheet('Accrual & Sub Cash Disb', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
  titleBlock(ac, en, 'Accrual & Subsequent Cash Disbursement', 'As of ' + short(q.end));
  hdrRow(ac, 6, ['Date', 'Vendor', 'JE#', 'Description', 'Accrual Account', 'Accrued Amount', 'Reversed / Paid', 'Reversal Date', 'Reversal JE#', 'Net Outstanding'], [12, 22, 9, 46, 28, 15, 15, 12, 12, 16]);
  ac.getRow(6).height = 30; ac.getColumn(11).width = 40;
  ac.getCell('A4').value = 'Accruals booked this quarter, plus earlier accruals still open or reversed during the quarter. Reversed / Paid is matched to each accrual (reversing JE first, then same amount, then oldest first); Net Outstanding = Accrued − Reversed and ties to the GL by account below.'; ac.getCell('A4').font = SMALLI; ac.mergeCells('A4:J4');
  let cr = 7; const acFirst = cr;
  const AI = data.accrual.items || [];
  for (const it of AI) {
    setDate(ac, 'A' + cr, it.date);
    ac.getCell('B' + cr).value = it.vendor || ''; ac.getCell('B' + cr).font = F();
    ac.getCell('C' + cr).value = it.entry_num || it.doc_number || ''; ac.getCell('C' + cr).font = F();
    ac.getCell('D' + cr).value = it.description || ''; ac.getCell('D' + cr).font = it.kind === 'unapplied' ? F({ italic: true }) : F();
    ac.getCell('E' + cr).value = String(it.account_code) + ' ' + (it.account_name || ''); ac.getCell('E' + cr).font = F();
    if (it.kind !== 'unapplied') setMoney(ac, 'F' + cr, it.accrued);
    setMoney(ac, 'G' + cr, it.reversed);
    if (it.reversal_date) setDate(ac, 'H' + cr, it.reversal_date);
    ac.getCell('I' + cr).value = it.reversal_je || ''; ac.getCell('I' + cr).font = F();
    fx(ac, 'J' + cr, 'F' + cr + '-G' + cr, it.net);
    cr++;
  }
  const acLast = cr - 1;
  if (!AI.length) { ac.getCell('D' + cr).value = 'No accruals booked or open for this period.'; ac.getCell('D' + cr).font = SMALLI; cr++; }
  ac.getCell('D' + cr).value = 'Total'; ac.getCell('D' + cr).font = F({ bold: true });
  ['F', 'G', 'J'].forEach((col) => {
    const tot = r2(AI.reduce((a, x) => a + (col === 'F' ? x.accrued : col === 'G' ? x.reversed : x.net), 0));
    (acLast >= acFirst ? fx(ac, col + cr, 'SUM(' + col + acFirst + ':' + col + acLast + ')', tot, { bold: true }) : setMoney(ac, col + cr, 0, { bold: true })).border = { top: THIN };
  });
  cr += 2;
  ac.getCell('D' + cr).value = 'Net outstanding by account — ties to the general ledger'; ac.getCell('D' + cr).font = F({ bold: true });
  [['F', 'Per schedule'], ['G', 'Per GL'], ['H', 'Difference']].forEach(([col, t]) => { const c = ac.getCell(col + cr); c.value = t; c.font = F({ bold: true }); c.alignment = { horizontal: 'right' }; });
  cr += 1;
  const accRowFor = {};
  for (const g of (data.accrual.groups || [])) {
    ac.getCell('D' + cr).value = g.label + ' (' + g.code + ')'; ac.getCell('D' + cr).font = F();
    if (acLast >= acFirst) fx(ac, 'F' + cr, 'SUMIF($E$' + acFirst + ':$E$' + acLast + ',"' + g.prefix + '*",$J$' + acFirst + ':$J$' + acLast + ')', g.schedule);
    else setMoney(ac, 'F' + cr, 0);
    setMoney(ac, 'G' + cr, g.gl);
    const dc = fx(ac, 'H' + cr, 'F' + cr + '-G' + cr, r2(g.schedule - g.gl)); dc.font = F({ bold: Math.abs(g.schedule - g.gl) >= 0.01, color: { argb: Math.abs(g.schedule - g.gl) >= 0.01 ? 'FFC00000' : 'FF008000' } });
    ac.getCell('K' + cr).value = 'agrees to BS: ' + g.bsLine; ac.getCell('K' + cr).font = SMALLI;
    accRowFor[g.code] = cr; cr += 1;
  }
  // Summary links: mgmt fees (accEnd1), accrued (accEnd2), due to affiliates (accEnd3).
  const accEnd1 = accRowFor['210600'], accEnd2 = accRowFor['210000'], accEnd3 = accRowFor['211000'];
  cr += 1;
  ac.getCell('A' + cr).value = 'Subsequent payment support for the open accruals is on the Subsequent Cash Disb tab; the management-fee roll-forward is on the Mgmt Fee Accrual tab.'; ac.getCell('A' + cr).font = SMALLI; ac.mergeCells('A' + cr + ':J' + cr);

  // ── 7. Distributions Payable & Due to Management Company ──────────────────
  // Roll-forward by account: beginning + net activity (GL inputs) = ending
  // (formula); open balance by investor is SUMmed and tied to the ending.
  let distTotalRow, dueMgmtRow;
  {
    const dd = wb.addWorksheet('Distributions Payable', { views: [{ showGridLines: false }] });
    dd.getColumn(1).width = 48; [2, 3, 4, 5, 6].forEach((c) => (dd.getColumn(c).width = 15)); dd.getColumn(7).width = 16;
    titleBlock(dd, en, 'Distributions Payable & Due to Management Company', 'As of ' + short(q.end));
    dd.getCell('A4').value = 'Amounts payable to members (230100) and to the management company (210800), rolled forward from the prior year-end with the open balance shown by investor. Each ending balance agrees to the Balance Sheet.'; dd.getCell('A4').font = SMALLI; dd.mergeCells('A4:G4');
    let r = 6;
    for (const grp of (data.distSplit || [])) {
      dd.getCell('A' + r).value = grp.label + ' (' + grp.code + ')'; dd.getCell('A' + r).font = F({ bold: true }); r += 1;
      const bRow = r; dd.getCell('A' + r).value = 'Beginning balance at ' + short(q.prior_ye); dd.getCell('A' + r).font = F(); setMoney(dd, 'G' + r, grp.begin); r += 1;
      const aRow = r; dd.getCell('A' + r).value = 'Net activity in ' + q.label; dd.getCell('A' + r).font = F(); setMoney(dd, 'G' + r, r2(grp.balance - grp.begin)); r += 1;
      const eRow = r; dd.getCell('A' + r).value = 'Ending balance at ' + short(q.end); dd.getCell('A' + r).font = F({ bold: true });
      fx(dd, 'G' + r, 'G' + bRow + '+G' + aRow, grp.balance, { bold: true }).border = { top: THIN };
      if (grp.code === '230100') distTotalRow = r;
      if (grp.code === '210800') dueMgmtRow = r;
      r += 1;
      dd.getCell('A' + r).value = 'Open balance by investor:'; dd.getCell('A' + r).font = F({ italic: true }); r += 1;
      if (grp.investors.length) {
        const iFirst = r;
        for (const inv of grp.investors) { dd.getCell('A' + r).value = '   ' + inv.name; dd.getCell('A' + r).font = F(); setMoney(dd, 'G' + r, inv.amt); r += 1; }
        const iLast = r - 1;
        dd.getCell('A' + r).value = '   Total by investor'; dd.getCell('A' + r).font = F({ bold: true });
        fx(dd, 'G' + r, 'SUM(G' + iFirst + ':G' + iLast + ')', r2(grp.investors.reduce((a, x) => a + x.amt, 0)), { bold: true }).border = { top: THIN }; r += 1;
        dd.getCell('A' + r).value = '   Difference vs ending balance (should be zero)'; dd.getCell('A' + r).font = F();
        fx(dd, 'G' + r, 'G' + (r - 1) + '-G' + eRow, r2(grp.investors.reduce((a, x) => a + x.amt, 0) - grp.balance)); r += 1;
      }
      else { dd.getCell('A' + r).value = '   (GL lines carry no investor/class tag; the ending balance still ties to the account)'; dd.getCell('A' + r).font = SMALLI; r += 1; }
      r += 1;
    }
    dd.getCell('A' + r).value = 'Due to Affiliates (211000) ' + fmt(data.ties.due_to_affiliates) + ' is a separate Balance Sheet line, supported on the Accrual & Sub Cash Disb tab.'; dd.getCell('A' + r).font = SMALLI; dd.mergeCells('A' + r + ':G' + r); r += 1;
    dd.getCell('A' + r).value = 'Each ending balance above agrees to the Balance Sheet.'; dd.getCell('A' + r).font = SMALLI;
  }

  // ═══ Phase 2-4 supporting schedules ═══════════════════════════════════════

  // Prepaid items — item-level amortization schedule (appended to the Prepaid tab).
  // Every amortization figure is a formula: days in coverage = end − start; days
  // elapsed = MAX(0, MIN(days, as-of − start)); amort to date = ROUND(premium ×
  // elapsed ÷ days, 2); quarter amort = to-date less to-date at the prior quarter
  // end; remaining = premium − amort to date.
  {
    [34, 22, 22, 24, 16, 11, 11, 14, 14, 14].forEach((w, i) => (pp.getColumn(i + 1).width = w));
    let pr = 12;
    pp.getCell('A' + pr).value = 'Prepaid items — amortization schedule'; pp.getCell('A' + pr).font = F({ bold: true }); pr += 1;
    pp.getCell('A' + pr).value = 'Schedule as of'; pp.getCell('A' + pr).font = F(); setDate(pp, 'B' + pr, q.end, { bold: true }); const asOfRef = '$B$' + pr; pr += 1;
    const qPrior = addDays(q.qs, -1);
    pp.getCell('A' + pr).value = 'Prior quarter end'; pp.getCell('A' + pr).font = F(); setDate(pp, 'B' + pr, qPrior); const priorRef = '$B$' + pr; pr += 2;
    const HH = ['Vendor / Description', 'Date Paid', 'Premium', 'Coverage start', 'Coverage end', 'Days in coverage', 'Days elapsed', 'Amort to date', 'Amort ' + q.quarter, 'Remaining'];
    for (const acct of (data.prepaid.itemsByAcct || [])) {
      pp.getCell('A' + pr).value = acct.code + ' — ' + acct.label; pp.getCell('A' + pr).font = F({ bold: true }); pr += 1;
      hdrRow(pp, pr, HH, null); pp.getRow(pr).height = 30; pr += 1;
      const first = pr;
      for (const it of acct.items) {
        pp.getCell('A' + pr).value = (it.vendor ? it.vendor + ' — ' : '') + (it.description || ''); pp.getCell('A' + pr).font = F();
        setDate(pp, 'B' + pr, it.date_paid || '');
        setMoney(pp, 'C' + pr, it.premium);
        if (it.coverageKnown) {
          setDate(pp, 'D' + pr, it.start); setDate(pp, 'E' + pr, it.end);
          const days = daysBetween(it.start, it.end), el = Math.max(0, Math.min(days, daysBetween(it.start, q.end)));
          { const c = pp.getCell('F' + pr); c.value = { formula: 'E' + pr + '-D' + pr, result: days }; c.numFmt = '0'; c.font = F(); }
          { const c = pp.getCell('G' + pr); c.value = { formula: 'MAX(0,MIN(F' + pr + ',' + asOfRef + '-D' + pr + '))', result: el }; c.numFmt = '0'; c.font = F(); }
          fx(pp, 'H' + pr, 'ROUND(C' + pr + '*G' + pr + '/F' + pr + ',2)', it.accumAmort);
          fx(pp, 'I' + pr, 'H' + pr + '-ROUND(C' + pr + '*MAX(0,MIN(F' + pr + ',' + priorRef + '-D' + pr + '))/F' + pr + ',2)', it.periodAmort);
          fx(pp, 'J' + pr, 'C' + pr + '-H' + pr, it.remaining);
        } else {
          pp.getCell('D' + pr).value = '(enter coverage on Prepaid Register)'; pp.getCell('D' + pr).font = SMALLI; pp.mergeCells('D' + pr + ':E' + pr);
          const _rc = pp.getCell('J' + pr); _rc.value = '—'; _rc.alignment = { horizontal: 'right' }; _rc.font = SMALLI;
        }
        pr += 1;
      }
      const last = pr - 1;
      pp.getCell('A' + pr).value = 'Total ' + acct.label + (acct.needCoverage > 0 ? ' (modeled, covered items)' : ''); pp.getCell('A' + pr).font = F({ bold: true });
      const sumCol = (col, v) => { const c = last >= first ? fx(pp, col + pr, 'SUM(' + col + first + ':' + col + last + ')', v, { bold: true }) : setMoney(pp, col + pr, 0, { bold: true }); c.border = { top: THIN }; };
      sumCol('C', r2(acct.items.reduce((s, x) => s + (x.premium || 0), 0)));
      sumCol('H', r2(acct.items.reduce((s, x) => s + (x.accumAmort || 0), 0)));
      sumCol('I', r2(acct.items.reduce((s, x) => s + (x.periodAmort || 0), 0)));
      sumCol('J', acct.modelRemaining);
      const totR = pr; pr += 1;
      pp.getCell('A' + pr).value = 'Balance per general ledger (' + acct.code + ')'; pp.getCell('A' + pr).font = F(); setMoney(pp, 'J' + pr, acct.endBal); const glR = pr; pr += 1;
      if (acct.needCoverage > 0) {
        pp.getCell('A' + pr).value = acct.needCoverage + ' item(s) need coverage dates; enter them on the Prepaid Register to complete the reconciliation.'; pp.getCell('A' + pr).font = F({ color: { argb: 'FFB45309' } }); pp.mergeCells('A' + pr + ':I' + pr);
      } else {
        pp.getCell('A' + pr).value = 'Difference (GL vs schedule)'; pp.getCell('A' + pr).font = F({ bold: Math.abs(acct.diff) >= 0.01 });
        const c = fx(pp, 'J' + pr, 'J' + glR + '-J' + totR, r2(acct.endBal - acct.modelRemaining)); c.font = F({ bold: Math.abs(acct.diff) >= 0.01, color: { argb: Math.abs(acct.diff) >= 0.01 ? 'FFC00000' : 'FF008000' } }); c.border = { top: THIN };
      }
      pr += 2;
    }
    pp.getCell('A' + pr).value = 'Amortization is straight-line over each item’s coverage period and every figure above is a live formula (change a date or premium and the schedule recalculates). Where the schedule does not tie to the GL, the amortization booked differs from the straight-line schedule (for example a period not yet amortized) — see the Summary exceptions. Items and coverage dates are maintained on the Prepaid Register.'; pp.getCell('A' + pr).font = SMALLI; pp.mergeCells('A' + pr + ':J' + pr); pp.getRow(pr).height = 42; pp.getCell('A' + pr).alignment = { wrapText: true, vertical: 'top' };
  }

  // Mgmt Fee Accrual — inception-to-date roll-forward (running balance formula).
  {
    const mi = data.accrual.mgmtItd;
    const ws = wb.addWorksheet('Mgmt Fee Accrual', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
    titleBlock(ws, en, 'Management Fees Payable (210600) — inception-to-date roll-forward', 'As of ' + short(q.end));
    hdrRow(ws, 5, ['Date', 'JE#', 'Description', 'Accrual (+)', 'Payment (-)', 'Balance'], [12, 10, 46, 16, 16, 16]);
    let r = 6; const first = r; let run = 0;
    for (const x of mi.rows) {
      setDate(ws, 'A' + r, x.date);
      ws.getCell('B' + r).value = x.entry_num || x.doc_number; ws.getCell('B' + r).font = F();
      ws.getCell('C' + r).value = x.memo || x.description; ws.getCell('C' + r).font = F();
      if (x.credit) setMoney(ws, 'D' + r, x.credit); if (x.debit) setMoney(ws, 'E' + r, x.debit);
      run = r2(run + (x.credit || 0) - (x.debit || 0));
      fx(ws, 'F' + r, (r === first ? '' : 'F' + (r - 1) + '+') + 'D' + r + '-E' + r, run); r += 1;
    }
    const endR = r;
    ws.getCell('C' + r).value = 'Ending management fees payable'; ws.getCell('C' + r).font = F({ bold: true });
    (r > first ? fx(ws, 'F' + r, 'F' + (r - 1), run, { bold: true }) : setMoney(ws, 'F' + r, 0, { bold: true })).border = { top: THIN }; r += 1;
    ws.getCell('C' + r).value = 'Balance per general ledger (210600)'; ws.getCell('C' + r).font = F(); setMoney(ws, 'F' + r, data.ties.management_fees_payable); r += 1;
    ws.getCell('C' + r).value = 'Difference (should be zero)'; ws.getCell('C' + r).font = F({ bold: true }); fx(ws, 'F' + r, 'F' + endR + '-F' + (r - 1), r2(run - data.ties.management_fees_payable), { bold: true }).border = { top: THIN };
    ws.getCell('A' + (r + 2)).value = 'Balance = prior balance + accrual − payment (formula). Ties to the Balance Sheet management fees payable line.'; ws.getCell('A' + (r + 2)).font = SMALLI;
  }

  // Subsequent Cash Disbursement — accruals paid in the following period.
  {
    const sd = data.accrual.subDisb;
    const ws = wb.addWorksheet('Subsequent Cash Disb', { views: [{ showGridLines: false }] });
    titleBlock(ws, en, 'Subsequent Cash Disbursement — accruals paid after quarter end', sd.from + ' to ' + sd.to);
    hdrRow(ws, 5, ['Date', 'JE#', 'Vendor', 'Description', 'Accrual account', 'Amount'], [12, 10, 22, 40, 24, 16]);
    let r = 6; const first = r;
    if (!sd.rows.length) { ws.getCell('A' + r).value = 'No subsequent-period disbursements are posted yet for this window — subsequent-payment support is incomplete until the next month is booked.'; ws.getCell('A' + r).font = SMALLI; ws.mergeCells('A' + r + ':F' + r); r += 1; }
    for (const x of sd.rows) {
      setDate(ws, 'A' + r, x.date);
      ws.getCell('B' + r).value = x.entry_num || x.doc_number; ws.getCell('B' + r).font = F();
      ws.getCell('C' + r).value = x.vendor; ws.getCell('C' + r).font = F();
      ws.getCell('D' + r).value = x.memo || x.description; ws.getCell('D' + r).font = F();
      ws.getCell('E' + r).value = x.account_code + ' ' + x.account_name; ws.getCell('E' + r).font = F();
      setMoney(ws, 'F' + r, x.debit); r += 1;
    }
    if (sd.rows.length) { ws.getCell('E' + r).value = 'Total subsequent disbursements'; ws.getCell('E' + r).font = F({ bold: true }); fx(ws, 'F' + r, 'SUM(F' + first + ':F' + (r - 1) + ')', r2(sd.rows.reduce((a, x) => a + x.debit, 0)), { bold: true }).border = { top: THIN }; }
    ws.getCell('A' + (r + 2)).value = 'Confirms the period-end accruals (accrued expenses, management fees payable, due to affiliates) were relieved by actual payments in the following period.'; ws.getCell('A' + (r + 2)).font = SMALLI; ws.mergeCells('A' + (r + 2) + ':F' + (r + 2));
  }

  // Flux Analysis — quarter-over-quarter income statement variance.
  {
    const fxa = data.accrual.flux;
    const ws = wb.addWorksheet('Flux Analysis', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
    titleBlock(ws, en, 'Flux Analysis — quarter-over-quarter income statement', 'Prior ' + fxa.priFrom + ' to ' + fxa.priTo + '   vs   current ' + fxa.curFrom + ' to ' + fxa.curTo);
    hdrRow(ws, 5, ['Account', 'Prior quarter', 'Current quarter', 'Difference', '% change'], [40, 16, 16, 16, 12]);
    let r = 6;
    for (const x of fxa.rows) {
      ws.getCell('A' + r).value = x.name || x.code; ws.getCell('A' + r).font = F();
      setMoney(ws, 'B' + r, x.prior); setMoney(ws, 'C' + r, x.cur); fx(ws, 'D' + r, 'C' + r + '-B' + r, x.diff);
      const pc = ws.getCell('E' + r); pc.value = { formula: 'IF(B' + r + '=0,"",D' + r + '/ABS(B' + r + '))', result: x.pct == null ? '' : x.pct }; pc.numFmt = '0.0%';
      pc.font = F({ bold: Math.abs(x.diff) >= 50000, color: { argb: Math.abs(x.diff) >= 50000 ? 'FFC00000' : 'FF000000' } });
      r += 1;
    }
    ws.getCell('A' + (r + 1)).value = 'Large quarter-over-quarter movements (>= $50,000) are flagged for review on the Summary tab.'; ws.getCell('A' + (r + 1)).font = SMALLI; ws.mergeCells('A' + (r + 1) + ':E' + (r + 1));
  }

  // Contribution Receivable (120020) — by investor.
  let contribTotRow;
  {
    const crv = data.contrib;
    const ws = wb.addWorksheet('Contribution Receivable', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
    titleBlock(ws, en, 'Contribution Receivable (120020)', 'As of ' + short(q.end));
    hdrRow(ws, 5, ['Date', 'JE#', 'Investor', 'Description', 'Amount', 'Balance'], [12, 10, 26, 40, 16, 16]);
    let r = 6; const first = r; let run = 0;
    for (const x of crv.rows) {
      setDate(ws, 'A' + r, x.date);
      ws.getCell('B' + r).value = x.entry_num || x.doc_number; ws.getCell('B' + r).font = F();
      ws.getCell('C' + r).value = x.class_name || ''; ws.getCell('C' + r).font = F();
      ws.getCell('D' + r).value = x.memo || x.description; ws.getCell('D' + r).font = F();
      setMoney(ws, 'E' + r, x.signed); run = r2(run + x.signed);
      fx(ws, 'F' + r, r === first ? 'E' + r : 'F' + (r - 1) + '+E' + r, run); r += 1;
    }
    ws.getCell('D' + r).value = 'TOTAL'; ws.getCell('D' + r).font = F({ bold: true });
    (r > first ? fx(ws, 'F' + r, 'SUM(E' + first + ':E' + (r - 1) + ')', run, { bold: true }) : setMoney(ws, 'F' + r, 0, { bold: true })).border = { top: THIN };
    contribTotRow = r; r += 2;
    if (crv.investors.length) {
      ws.getCell('A' + r).value = 'Open receivable by investor'; ws.getCell('A' + r).font = F({ bold: true }); r += 1;
      const iFirst = r;
      for (const inv of crv.investors) { ws.getCell('A' + r).value = inv.name; ws.getCell('A' + r).font = F(); setMoney(ws, 'E' + r, inv.amt); r += 1; }
      ws.getCell('A' + r).value = 'Total by investor'; ws.getCell('A' + r).font = F({ bold: true });
      fx(ws, 'E' + r, 'SUM(E' + iFirst + ':E' + (r - 1) + ')', r2(crv.investors.reduce((a, x) => a + x.amt, 0)), { bold: true }).border = { top: THIN }; r += 2;
    }
    ws.getCell('A' + r).value = 'Ties to the Balance Sheet capital contributions receivable line.'; ws.getCell('A' + r).font = SMALLI;
  }


  // ── Link each Summary balance to the total cell on its supporting schedule ──
  const sq = (t) => "'" + t + "'!";
  const ACS = sq('Accrual & Sub Cash Disb');
  const sumRefs = [
    // Net Due From (To): SUM of the property columns on the ending row. A positive
    // result is a net due-FROM (asset), shown positive to match the Balance Sheet.
    { f: 'SUM(' + sq('Due Fr (To) Port Co') + 'E' + dueEndRow + ':K' + dueEndRow + ')', v: data.ties.due_net },
    { f: sq('Interest Receivable') + 'G' + intTotalRow, v: data.ties.interest_receivable },
    { f: sq('Prepaid Expenses') + 'E8', v: r2(data.ties.prepaid_insurance + data.ties.prepaid_advisory + data.ties.prepaid_subscription) },
    { f: sq('Other Assets') + 'F' + oaTotalRow, v: data.ties.other_assets },
    // Accounts payable & accrued expenses = trade AP (202000, AP Recon) + accrued
    // (210000, Accrual tab), tying to the single Balance Sheet line.
    { f: accEnd2 ? sq('AP Recon') + 'C' + apTotalRow + '+' + ACS + 'F' + accEnd2 : null, v: r2(data.ties.accounts_payable + data.ties.accrued_expenses) },
    { f: accEnd1 ? ACS + 'F' + accEnd1 : null, v: data.ties.management_fees_payable },
    { f: sq('Distributions Payable') + 'G' + distTotalRow, v: data.ties.distributions_payable },
    { f: contribTotRow ? sq('Contribution Receivable') + 'F' + contribTotRow : null, v: data.ties.contribution_receivable },
    { f: dueMgmtRow ? sq('Distributions Payable') + 'G' + dueMgmtRow : null, v: data.ties.due_to_mgmt },
    { f: accEnd3 ? ACS + 'F' + accEnd3 : null, v: data.ties.due_to_affiliates },
  ];
  if (!distTotalRow) sumRefs[6].f = null;
  sumRefs.forEach((x, i) => { const c = su.getCell('D' + (6 + i)); c.value = x.f ? { formula: x.f, result: x.v } : x.v; c.numFmt = MONEY; c.font = F(); });

  return wb;
}

const fmt = (n) => '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Generic "Account QuickReport" transaction-detail sheet (Weaver's GL layout).
// Beginning balance (GL input) + each transaction = running balance (formula);
// TOTAL SUMs the amounts, and the ending balance is checked against the GL.
function detailSheet(wb, tabName, en, subtitle, q, rows, balance, itd) {
  const ws = wb.addWorksheet(tabName, { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  titleBlock(ws, en, 'Account QuickReport — ' + subtitle, (itd ? 'Inception to ' : q.ys + ' to ') + short(q.end));
  hdrRow(ws, 5, ['Location', 'Date', 'Type', 'Num', 'Description', 'Amount', 'Balance'], [16, 12, 14, 16, 52, 16, 16]);
  const act = r2(rows.reduce((a, x) => a + x.signed, 0));
  const begin = itd ? 0 : r2(balance - act);
  let r = 6; const begRow = r;
  ws.getCell('E' + r).value = 'Beginning balance' + (itd ? '' : ' at ' + short(q.prior_ye)); ws.getCell('E' + r).font = F({ italic: true });
  setMoney(ws, 'G' + r, begin); r++;
  const first = r; let run = begin;
  for (const x of rows) {
    ws.getCell('A' + r).value = x.location_name || x.class_name || ''; ws.getCell('A' + r).font = F();
    setDate(ws, 'B' + r, x.date);
    ws.getCell('C' + r).value = ''; // transaction type not stored
    ws.getCell('D' + r).value = x.entry_num || x.doc_number; ws.getCell('D' + r).font = F();
    ws.getCell('E' + r).value = x.description || x.memo; ws.getCell('E' + r).font = F();
    setMoney(ws, 'F' + r, x.signed); run = r2(run + x.signed);
    fx(ws, 'G' + r, 'G' + (r - 1) + '+F' + r, run);
    r++;
  }
  const totR = r;
  ws.getCell('E' + r).value = 'TOTAL'; ws.getCell('E' + r).font = F({ bold: true });
  (r > first ? fx(ws, 'F' + r, 'SUM(F' + first + ':F' + (r - 1) + ')', act, { bold: true }) : setMoney(ws, 'F' + r, 0, { bold: true })).border = { top: THIN };
  fx(ws, 'G' + r, 'G' + begRow + '+F' + r, r2(begin + act), { bold: true }).border = { top: THIN };
  ws.getCell('E' + (r + 1)).value = 'Balance per general ledger'; ws.getCell('E' + (r + 1)).font = F(); setMoney(ws, 'G' + (r + 1), balance);
  ws.getCell('E' + (r + 2)).value = 'Difference (should be zero)'; ws.getCell('E' + (r + 2)).font = F({ bold: true }); fx(ws, 'G' + (r + 2), 'G' + totR + '-G' + (r + 1), r2(begin + act - balance), { bold: true }).border = { top: THIN };
  ws.getCell('E' + (r + 4)).value = 'Ties to the Balance Sheet.'; ws.getCell('E' + (r + 4)).font = SMALLI;
  return totR; // TOTAL row — Summary links to G here
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
        const data = await buildData(ctx, q, { entity_id: eid });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, q, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Other-Summary', JSON.stringify({
          quarter: q.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          ties: data.ties, exceptions: (data.flags || []).length, flags: (data.flags || []),
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
  // ── Bill.com A/P Detail (open items) for the AP Recon ─────────────────────
  // The user exports the Bill.com A/P Detail report (filtered to this fund) and
  // uploads it here; the AP Recon reconciles it to the GL. Stored as the aging
  // lines on billcom_config; the sync cutoff is never touched.
  app.get('/api/workpapers/other/:entity_id/ap-detail', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const cfg = ctx.db.prepare('SELECT ap_aging_lines_json, ap_aging_as_of, ap_aging_uploaded_at FROM billcom_config WHERE entity_id = ?').get(eid);
        let lines = [];
        try { lines = cfg && cfg.ap_aging_lines_json ? JSON.parse(cfg.ap_aging_lines_json) : []; } catch (e) { lines = []; }
        const total = Math.round(lines.reduce((a, x) => a + (Number(x.amount) || 0), 0) * 100) / 100;
        res.json({ configured: !!cfg, as_of: cfg ? cfg.ap_aging_as_of : null, uploaded_at: cfg ? cfg.ap_aging_uploaded_at : null, count: lines.length, total, lines });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });
  app.post('/api/workpapers/other/:entity_id/ap-detail', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const cfg = ctx.db.prepare('SELECT entity_id FROM billcom_config WHERE entity_id = ?').get(eid);
        // The A/P Detail upload only needs somewhere to store the aging lines; it does
        // not require a live Bill.com connection. If the entity has no billcom_config
        // row yet (e.g. a GL-native A/P like Banyan), create a minimal placeholder so
        // the upload persists rather than being rejected.
        if (!cfg) ctx.db.prepare("INSERT INTO billcom_config (entity_id, api_base_url, username, password_enc, org_id, dev_key_enc) VALUES (?, '', '', '', '', '')").run(eid);
        const asOf = (req.body && /^\d{4}-\d{2}-\d{2}$/.test(String(req.body.as_of || ''))) ? String(req.body.as_of) : null;
        const arr = Array.isArray(req.body && req.body.lines) ? req.body.lines : null;
        if (!arr) return res.status(400).json({ error: 'lines array is required' });
        const clean = arr.map((l) => {
          const invRaw = l && (l.invoice_number != null ? l.invoice_number : (l.invoice != null ? l.invoice : l.num));
          const dtRaw = l && (l.bill_date != null ? l.bill_date : l.date);
          return {
            vendor: l && l.vendor != null ? String(l.vendor) : '',
            invoice_number: invRaw != null ? String(invRaw) : '',
            bill_date: dtRaw ? String(dtRaw).slice(0, 10) : null,
            amount: l && l.amount != null && !isNaN(Number(l.amount)) ? Math.round(Number(l.amount) * 100) / 100 : null,
          };
        }).filter((l) => l.amount != null && Math.abs(l.amount) >= 0.005);
        const now = new Date().toISOString();
        const who = (req.user && (req.user.name || req.user.email)) || 'system';
        ctx.db.prepare('UPDATE billcom_config SET ap_aging_lines_json=?, ap_aging_as_of=?, ap_aging_uploaded_at=?, updated_by=?, updated_at=? WHERE entity_id=?')
          .run(JSON.stringify(clean), asOf, now, who, now, eid);
        const total = Math.round(clean.reduce((a, x) => a + x.amount, 0) * 100) / 100;
        res.json({ success: true, count: clean.length, total, as_of: asOf });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });
  app.delete('/api/workpapers/other/:entity_id/ap-detail', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        ctx.db.prepare('UPDATE billcom_config SET ap_aging_lines_json=NULL, ap_aging_as_of=NULL, ap_aging_uploaded_at=NULL WHERE entity_id=?').run(eid);
        res.json({ success: true });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });

}

module.exports = { registerOtherWorkpapersRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, FUND_EID };
