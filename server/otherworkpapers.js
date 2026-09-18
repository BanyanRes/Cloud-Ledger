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
    ys: y + '-01-01', end: quarterEnd, prior_ye: (y - 1) + '-12-31',
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

  // 2. Interest Receivable — detail.
  const intRows = glDetail(db, eid, { from: quarter.ys, to: quarter.end, match: ACCT.intRecv });
  const intBal = sumWhere(bEnd, ACCT.intRecv);

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

  // 5. AP Recon — account 202000 balance vs Bill.com. CL GL carries no per-vendor
  // tag on these lines, so the recon is presented at the ledger level with the
  // Bill.com open-bill total (when available) alongside.
  const apGl = sumWhere(bEnd, ACCT.tradeAP);
  let apBillcom = null, apByVendor = [];
  try {
    const rows = db.prepare(`
      SELECT vendor_name AS vendor, SUM(amount_due) AS amt FROM billcom_bills
      WHERE entity_id = ? AND status != 'paid' AND due_date IS NOT NULL
      GROUP BY vendor_name`).all(eid);
    if (rows && rows.length) { apByVendor = rows.map((x) => ({ vendor: x.vendor, amt: r2(x.amt) })); apBillcom = r2(apByVendor.reduce((a, x) => a + x.amt, 0)); }
  } catch (e) { apBillcom = null; }

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
    interest: { rows: intRows, balance: intBal },
    prepaid,
    otherAssets: { groups: oaGroups, total: oaTotal },
    ap: { gl: apGl, billcom: apBillcom, byVendor: apByVendor },
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

  // ── 5. AP Recon ───────────────────────────────────────────────────────────
  const ap = wb.addWorksheet('AP Recon', { views: [{ showGridLines: false }] });
  ap.getColumn(1).width = 40; [2, 3, 4].forEach((c) => (ap.getColumn(c).width = 16));
  titleBlock(ap, en, 'AP Recon', 'As of ' + short(q.end));
  hdrRow(ap, 6, ['Vendor Name', 'Per Bill.com', 'Per GL', 'Difference'], [40, 16, 16, 16]);
  let ar = 7;
  if (data.ap.byVendor && data.ap.byVendor.length) {
    for (const v of data.ap.byVendor) {
      ap.getCell('A' + ar).value = v.vendor; ap.getCell('A' + ar).font = F();
      setMoney(ap, 'B' + ar, v.amt); ap.getCell('C' + ar).value = ''; setMoney(ap, 'D' + ar, 0); ar++;
    }
  } else {
    ap.getCell('A' + ar).value = 'Accounts Payable (202000) per general ledger'; ap.getCell('A' + ar).font = F();
    setMoney(ap, 'B' + ar, data.ap.billcom == null ? '' : data.ap.billcom);
    setMoney(ap, 'C' + ar, data.ap.gl); setMoney(ap, 'D' + ar, data.ap.billcom == null ? 0 : r2(data.ap.billcom - data.ap.gl)); ar++;
  }
  ap.getCell('A' + ar).value = 'Total:'; ap.getCell('A' + ar).font = F({ bold: true });
  const apTot = data.ap.byVendor && data.ap.byVendor.length ? r2(data.ap.byVendor.reduce((a, x) => a + x.amt, 0)) : data.ap.gl;
  setMoney(ap, 'B' + ar, data.ap.billcom == null ? apTot : data.ap.billcom, { bold: true }).border = { top: THIN };
  setMoney(ap, 'C' + ar, data.ap.gl, { bold: true }).border = { top: THIN };
  setMoney(ap, 'D' + ar, data.ap.billcom == null ? 0 : r2(data.ap.billcom - data.ap.gl), { bold: true }).border = { top: THIN };
  ap.getCell('A' + (ar + 2)).value = 'Per GL = Accounts Payable account 202000, which ties to the Balance Sheet. CL carries no per-vendor tag on GL AP lines; the Bill.com column shows open bills by vendor when available.';
  ap.getCell('A' + (ar + 2)).font = SMALLI; ap.mergeCells('A' + (ar + 2) + ':D' + (ar + 2));
  const apTotalRow = ar; // Summary links to C here (Per GL)

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
          ties: data.ties,
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerOtherWorkpapersRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, FUND_EID };
