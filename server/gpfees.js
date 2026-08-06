// ─── CLRF workpaper: Schedule of Fees Paid to the General Partner & Affiliates ──
//
// A quarterly workpaper for County Line Rail Fund I, LP. CLRF's own ledger holds
// none of these fees - they sit in the four portfolio-company entities - so the
// report reads across those four and presents the result as a CLRF deliverable.
//
// Three fee types:
//   Management fee            GL 63041, expensed monthly by the property
//   Comp reimbursement        GL 63030 / 63034 / 63042
//   Development fee           GL 12913, a CAPITALIZED asset, not a P&L account
//
// The development fee needs care. The balance in 12913 moves for reasons that are
// not fees: on completion the capitalized cost is allocated out to fixed assets
// ("Completed Dev Allocations"), and amounts are occasionally reclassed elsewhere.
// So the fee is NOT the change in the balance - measuring it that way reports
// -477,221.26 for CLIP in Q4 2025 against an actual fee of 210,424.79. The rule
// used here, validated against Q4 2025, Q1 2026 and Q2 2026 for all four
// properties, is:
//
//     fee = debits to 12913, LESS credits whose description begins "Reversed --"
//
// A debit is a fee billed. A "Reversed --" credit reverses a fee accrual (its
// offsetting entry is Accounts Payable) and must net. Every other credit moves the
// fee out of 12913 to somewhere else and is never a fee.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// The four CLRF portfolio companies. CloudLedger has no parent/child or
// consolidation relationship on `entities`, so membership is declared here. Add a
// property when the fund acquires one - the report prints the entities it included
// on the Summary tab and in the Notes, so an omission shows on the face of the
// deliverable rather than silently understating the total.
const PROPERTIES = [
  { label: 'CLIP', entity_id: 54 },
  { label: 'Silsbee', entity_id: 39 },
  { label: 'SRN', entity_id: 37 },
  { label: 'Buna', entity_id: 38 },
];

const DEV_ACCT = '12913';
const FEE_ACCOUNTS = {
  '63041': { name: 'CLRO Management Fees', category: 'Management Fee' },
  '63030': { name: 'Transload Salaries - Reimbursable', category: 'Comp Reimbursement' },
  '63034': { name: 'Switchmen Salaries - Reimbursable', category: 'Comp Reimbursement' },
  '63042': { name: 'Offsite Staff', category: 'Comp Reimbursement' },
};
// Related-party accounts deliberately shown as a memo, outside the totals.
const MEMO_ACCOUNTS = { '63038': 'Management Fees - Other', '63047': 'FRA Compliance Staff' };

const AFFILIATE = {
  'Management Fee': 'County Line Rail Operations',
  'Comp Reimbursement': 'County Line Rail Operations',
  'Development Fee': 'County Line Railroad Interests',
};
const DESCRIPTION = {
  'Management Fee': 'Compensation to County Line Rail Operations, LLC for providing exclusive rail services at each '
    + 'Fund facility — including rail car handling, switching, spotting, transloading, and coordinating with the '
    + 'Class I railroads.',
  'Comp Reimbursement': 'Reimbursement to County Line Rail Operations, LLC for compensation paid to its employees '
    + "for personnel costs directly benefiting the Fund's rail assets, including rail operations, sales and "
    + 'marketing, construction management, and accounting.',
  'Development Fee': 'Development fee paid to County Line Railroad Interests in connection with the development of '
    + 'each Property following its acquisition by the Fund or its Subsidiary, in accordance with a Budget approved '
    + 'by the Manager and the Members holding a Majority Interest.',
};
const BASIS = {
  'Management Fee': 'Tiered percentage of Effective Gross Income (EGI) per facility, applied separately to rail vs. '
    + 'non-rail activities. Rail Services: 10% on first $5M of annualized EGI, 8.5% from $5M–$10M, 7.5% above $10M. '
    + 'Non-Rail Management: flat 3% of EGI. EGI = Net Effective Rent (gross rent less concessions, refunds, '
    + 'vacancies) plus other operating revenues; excludes equity contributions, loan/insurance/condemnation '
    + 'proceeds, and capital asset sales.',
  'Comp Reimbursement': 'Pass-through of actual, documented compensation costs paid by Operator to its employees, '
    + 'allocated by rail assets',
  'Development Fee': '4.0% of the total hard and soft costs actually incurred by the rail assets for the development '
    + 'of the Property, excluding acquisition costs of the Property.',
};

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

// ── Quarter arithmetic. A quarterly workpaper only accepts a real quarter end; a
// typo like 2026-03-13 would otherwise silently produce a nonsense period.
function resolveQuarter(quarterEnd) {
  if (!isDate(quarterEnd)) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const [y, m, d] = quarterEnd.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) {
    throw new Error('quarter_end must be a quarter end date: 03-31, 06-30, 09-30 or 12-31. Received ' + quarterEnd);
  }
  const q = m / 3;
  const priorEnd = new Date(Date.UTC(y, m - 3, 1));
  priorEnd.setUTCDate(0); // last day of the month before the quarter start
  return {
    label: y + '-Q' + q,
    year: String(y),
    quarter: 'Q' + q,
    start: y + '-' + String(m - 2).padStart(2, '0') + '-01',
    end: quarterEnd,
    prior_end: priorEnd.toISOString().slice(0, 10),
  };
}

// ── Development fee for the quarter, per the rule documented at the top.
function devFeeLines(db, eid, from, to) {
  const rows = db.prepare(
    'SELECT je.date AS date, je.entry_num AS entry_num, jl.debit AS debit, jl.credit AS credit, '
    + "COALESCE(NULLIF(jl.description, ''), je.memo) AS description "
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + 'WHERE je.entity_id = ? AND jl.account_code = ? AND je.date >= ? AND je.date <= ? '
    + 'ORDER BY je.date, je.entry_num'
  ).all(eid, DEV_ACCT, from, to);
  let fee = 0;
  const lines = rows.map((x) => {
    const dr = r2(x.debit), cr = r2(x.credit);
    const isReversal = /^\s*Reversed\s*--/i.test(String(x.description || ''));
    let effect = 0, treatment;
    if (dr > 0.004) { effect = dr; treatment = 'Fee billed'; }
    else if (isReversal) { effect = -cr; treatment = 'Reversal of a fee accrual'; }
    else { effect = 0; treatment = 'Excluded - transfer out of ' + DEV_ACCT + ', not a fee'; }
    fee = r2(fee + effect);
    return { date: x.date, entry_num: x.entry_num, debit: dr, credit: cr,
      description: String(x.description || ''), treatment, effect };
  });
  return { lines, fee: r2(fee) };
}

// ── Gather everything the workbook needs.
function buildData(ctx, quarter) {
  const { db, computeBalances } = ctx;
  const out = { quarter, properties: [] };
  for (const p of PROPERTIES) {
    const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(p.entity_id);
    // Quarterly trial balance: balance-sheet accounts at their closing balance on
    // the quarter end, P&L accounts as activity from the quarter start.
    const bal = computeBalances(p.entity_id, { as_of: quarter.end, close_pl_before: quarter.start });
    const dpOf = (code) => {
      const x = bal.find((b) => String(b.code) === code);
      return x ? r2((x.total_debit || 0) - (x.total_credit || 0)) : 0;
    };
    const fees = {}, memo = {};
    for (const code of Object.keys(FEE_ACCOUNTS)) { const v = dpOf(code); if (Math.abs(v) > 0.004) fees[code] = v; }
    for (const code of Object.keys(MEMO_ACCOUNTS)) { const v = dpOf(code); if (Math.abs(v) > 0.004) memo[code] = v; }
    const dev = devFeeLines(db, p.entity_id, quarter.start, quarter.end);
    const priorBal = computeBalances(p.entity_id, { as_of: quarter.prior_end });
    const px = priorBal.find((y) => String(y.code) === DEV_ACCT);
    // Operating revenue as an EGI proxy for the rate check. 401xx/411xx only:
    // interest income and unrealized gains are outside EGI per the agreement.
    const revenue = bal
      .filter((b) => /^(40|41)/.test(String(b.code)) && Math.abs(b.balance) > 0.004)
      .map((b) => ({ code: String(b.code), name: b.name, amount: r2((b.total_credit || 0) - (b.total_debit || 0)) }))
      .sort((a, b) => a.code.localeCompare(b.code));
    let tdr = 0, tcr = 0;
    bal.forEach((b) => { const v = r2((b.total_debit || 0) - (b.total_credit || 0)); if (v > 0) tdr += v; else tcr += -v; });
    out.properties.push({
      label: p.label, entity_id: p.entity_id, entity_name: ent ? ent.name : ('entity ' + p.entity_id),
      fees, memo, dev_lines: dev.lines, dev_fee: dev.fee,
      dev_open: px ? r2((px.total_debit || 0) - (px.total_credit || 0)) : 0, dev_close: dpOf(DEV_ACCT),
      revenue, tb_debits: r2(tdr), tb_credits: r2(tcr), account_count: bal.length,
    });
  }
  return out;
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const PCT = '0.00%';
const NAVY = 'FF1F3864';
const HDR_FONT = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const BLUE = F({ color: { argb: 'FF0000FF' } });   // value read from the source system
const GREEN = F({ color: { argb: 'FF008000' } });  // cross-sheet formula
const SMALL = F({ size: 9 });
const SMALLI = F({ size: 9, italic: true });
const AMBER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

// ── Summary-of-Fees tab styling. This tab alone mirrors the client's
// "Fee & Expense to GP" deliverable: Times New Roman throughout, a plain
// thin-bordered header (no navy fill), and accounting-format amounts. The other
// tabs keep the Arial/navy CloudLedger theme above.
const SUM_MONEY = '_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)';
const SF = (o = {}) => Object.assign({ name: 'Times New Roman', size: 10 }, o);
const SF_SMALL = SF();
const SF_HDR = SF({ bold: true });
const THIN = { style: 'thin' };
// Spell the quarter end out as "MARCH 31, 2026" to match the deliverable header.
const MONTHS = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
  'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
const spellQuarterEnd = (end) => {
  const [y, m, d] = String(end).split('-').map(Number);
  return MONTHS[m - 1] + ' ' + d + ', ' + y;
};

function headerRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR_FONT;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true };
  });
  if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

function buildWorkbook(data) {
  const q = data.quarter;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  wb.created = new Date();

  const s = wb.addWorksheet('Summary of Fees');
  const gl = wb.addWorksheet('GL Data', { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
  const dv = wb.addWorksheet('Dev Fee Detail');
  const fr = wb.addWorksheet('Fee Reasonableness');
  const nt = wb.addWorksheet('Notes & Sources');

  // ── GL Data ───────────────────────────────────────────────────────────────
  gl.getCell('A1').value = 'GL DATA - QUARTERLY ACTIVITY BY ACCOUNT';
  gl.getCell('A1').font = F({ size: 12, bold: true });
  gl.getCell('A2').value = 'Quarterly trial balance per entity: balance-sheet accounts at their ' + q.end
    + ' closing balance, P&L accounts as activity from ' + q.start + '.';
  gl.getCell('A2').font = SMALLI;
  headerRow(gl, 4, ['Entity', 'Entity ID', 'Period start', 'Period end', 'Account', 'Account name',
    'Fee category', 'Amount'], [12, 10, 13, 13, 10, 34, 22, 17]);
  let r = 5;
  const glFirst = r;
  const emit = (p, code, name, cat, amt) => {
    const row = gl.getRow(r);
    row.getCell(1).value = p.label; row.getCell(1).font = F();
    row.getCell(2).value = p.entity_id; row.getCell(2).font = F();
    row.getCell(3).value = q.start; row.getCell(3).font = F();
    row.getCell(4).value = q.end; row.getCell(4).font = F();
    row.getCell(5).value = code; row.getCell(5).font = F();
    row.getCell(6).value = name; row.getCell(6).font = F();
    row.getCell(7).value = cat; row.getCell(7).font = F();
    const c = row.getCell(8); c.value = amt; c.font = BLUE; c.numFmt = MONEY;
    r += 1;
  };
  for (const p of data.properties) {
    for (const code of Object.keys(p.fees).sort()) {
      emit(p, code, FEE_ACCOUNTS[code].name, FEE_ACCOUNTS[code].category, p.fees[code]);
    }
  }
  const glLast = r - 1;
  const memoFirst = r;
  for (const p of data.properties) {
    for (const code of Object.keys(p.memo).sort()) {
      emit(p, code, MEMO_ACCOUNTS[code], 'Memo - not in totals', p.memo[code]);
    }
  }
  const memoLast = r - 1;
  const hasMemo = memoLast >= memoFirst;
  const glRow = gl.getRow(r);
  glRow.getCell(7).value = hasMemo ? 'Total (incl. memo)' : 'Total';
  glRow.getCell(7).font = F({ bold: true });
  const glTot = glRow.getCell(8);
  glTot.value = { formula: 'SUM(H' + glFirst + ':H' + (hasMemo ? memoLast : glLast) + ')' };
  glTot.font = F({ bold: true }); glTot.numFmt = MONEY;
  glTot.border = { top: { style: 'thin' }, bottom: { style: 'double' } };

  // ── Dev Fee Detail ────────────────────────────────────────────────────────
  dv.getCell('A1').value = 'DEVELOPMENT FEE - ENTRY DETAIL, GL ' + DEV_ACCT;
  dv.getCell('A1').font = F({ size: 12, bold: true });
  dv.getCell('A2').value = 'The fee is NOT the change in the capitalized balance. A debit is a fee billed; a '
    + '"Reversed --" credit reverses a fee accrual and nets against it; any other credit transfers the fee out of '
    + DEV_ACCT + ' - a capitalization allocation or a reclass - and is excluded.';
  dv.getCell('A2').font = SMALLI;
  dv.getCell('A2').alignment = { wrapText: true, vertical: 'top' };
  dv.getRow(2).height = 30;
  headerRow(dv, 4, ['Entity', 'Date', 'Entry #', 'Debit', 'Credit', 'Effect on fee', 'Treatment',
    'Description'], [12, 12, 9, 15, 15, 15, 34, 66]);
  r = 5;
  const devRows = {};
  for (const p of data.properties) {
    const first = r;
    for (const l of p.dev_lines) {
      const row = dv.getRow(r);
      row.getCell(1).value = p.label; row.getCell(1).font = F();
      row.getCell(2).value = l.date; row.getCell(2).font = F();
      row.getCell(3).value = l.entry_num; row.getCell(3).font = F();
      let c = row.getCell(4); c.value = l.debit || null; c.font = BLUE; c.numFmt = MONEY;
      c = row.getCell(5); c.value = l.credit || null; c.font = BLUE; c.numFmt = MONEY;
      c = row.getCell(6); c.value = l.effect; c.font = F(); c.numFmt = MONEY;
      if (l.effect === 0) c.fill = AMBER_FILL;
      row.getCell(7).value = l.treatment; row.getCell(7).font = SMALL;
      row.getCell(7).alignment = { wrapText: true, vertical: 'top' };
      row.getCell(8).value = l.description; row.getCell(8).font = SMALL;
      row.getCell(8).alignment = { wrapText: true, vertical: 'top' };
      r += 1;
    }
    const tr = dv.getRow(r);
    tr.getCell(1).value = p.label + ' ' + q.quarter + ' development fee';
    tr.getCell(1).font = F({ bold: true });
    const tc = tr.getCell(6);
    tc.value = r > first ? { formula: 'SUM(F' + first + ':F' + (r - 1) + ')' } : 0;
    tc.font = F({ bold: true }); tc.numFmt = MONEY; tc.border = { top: { style: 'thin' } };
    devRows[p.label] = r;
    const ctxRow = dv.getRow(r + 1);
    ctxRow.getCell(1).value = '   capitalized balance ' + q.prior_end + ' ' + p.dev_open.toFixed(2) + ' -> '
      + q.end + ' ' + p.dev_close.toFixed(2) + '   (raw movement '
      + r2(p.dev_close - p.dev_open).toFixed(2) + ', which is not the fee)';
    ctxRow.getCell(1).font = SMALLI;
    r += 3;
  }

  // ── Summary of Fees ───────────────────────────────────────────────────────
  // This tab mirrors the client's "Fee & Expense to GP" deliverable exactly:
  // Times New Roman, a merged/centered title block, a plain thin-bordered header,
  // subtotal labels that read "Total management fees" / "Total reimbursements" /
  // "Total development fees", accounting-format amounts, and no memo block. The
  // memo accounts still surface on the GL Data tab; they are simply excluded from
  // the face of this schedule.
  s.getColumn(1).width = 24; s.getColumn(2).width = 45; s.getColumn(3).width = 55;
  s.getColumn(4).width = 28; s.getColumn(5).width = 26; s.getColumn(6).width = 16;
  const centerTitle = (rowNum, text, opts) => {
    s.mergeCells('A' + rowNum + ':F' + rowNum);
    const c = s.getCell('A' + rowNum);
    c.value = text; c.font = opts.font;
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    s.getRow(rowNum).height = opts.height;
  };
  centerTitle(1, 'COUNTY LINE RAIL FUND I, LP', { font: SF({ size: 12, bold: true }), height: 18 });
  centerTitle(2, 'SCHEDULE OF FEES PAID TO THE GENERAL PARTNER AND AFFILIATES',
    { font: SF({ size: 11, bold: true }), height: 18 });
  centerTitle(3, 'FOR THE QUARTER ENDED ' + spellQuarterEnd(q.end),
    { font: SF({ size: 11, bold: true }), height: 15.75 });
  centerTitle(4, '(Amounts in USD)', { font: SF({ size: 10, italic: true }), height: 13.5 });

  // Header row (row 6) - black bold text, no fill, thin rule above and below.
  const HDR_LABELS = ['Fee Type', 'Description', 'Basis of Calculation', 'GP Affiliate',
    'Portfolio Company', 'Amount'];
  const hRow = s.getRow(6); hRow.height = 21.75;
  HDR_LABELS.forEach((t, i) => {
    const c = hRow.getCell(i + 1);
    c.value = t; c.font = SF_HDR;
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    c.border = { top: THIN, bottom: THIN };
  });

  // Subtotal labels match the deliverable's plural wording exactly.
  const LABELS = { 'Management Fee': 'Management Fee',
    'Comp Reimbursement': 'Employee Compensation Reimbursement', 'Development Fee': 'Development Fee' };
  const TOTAL_LABELS = { 'Management Fee': 'Total management fees',
    'Comp Reimbursement': 'Total reimbursements', 'Development Fee': 'Total development fees' };

  r = 8;
  const subtotals = [];
  for (const cat of ['Management Fee', 'Comp Reimbursement', 'Development Fee']) {
    const first = r;
    const head = s.getRow(r);
    // Fee-type label is bold + italic on the deliverable.
    head.getCell(1).value = LABELS[cat]; head.getCell(1).font = SF({ bold: true, italic: true });
    head.getCell(1).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    [[2, DESCRIPTION[cat]], [3, BASIS[cat]], [4, AFFILIATE[cat]]].forEach((pair) => {
      const c = head.getCell(pair[0]); c.value = pair[1]; c.font = SF_SMALL;
      c.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    });
    s.getRow(r).height = cat === 'Management Fee' ? 91 : 65;
    for (const p of data.properties) {
      const row = s.getRow(r);
      const eCell = row.getCell(5); eCell.value = p.label; eCell.font = SF_SMALL;
      eCell.alignment = { horizontal: 'left', vertical: r === first ? 'top' : 'center' };
      const c = row.getCell(6);
      if (cat === 'Development Fee') {
        c.value = { formula: "'Dev Fee Detail'!F" + devRows[p.label] };
      } else {
        c.value = { formula: "SUMIFS('GL Data'!$H$" + glFirst + ':$H$' + glLast
          + ",'GL Data'!$A$" + glFirst + ':$A$' + glLast + ',$E' + r
          + ",'GL Data'!$G$" + glFirst + ':$G$' + glLast + ',"' + cat + '")' };
      }
      c.font = SF_SMALL; c.numFmt = SUM_MONEY;
      c.alignment = { horizontal: 'right', vertical: r === first ? 'top' : 'center' };
      r += 1;
    }
    const tr = s.getRow(r);
    const lc = tr.getCell(5);
    lc.value = TOTAL_LABELS[cat]; lc.font = SF_HDR;
    lc.alignment = { horizontal: 'left', vertical: 'center' };
    lc.border = { bottom: THIN };
    const tc = tr.getCell(6);
    tc.value = { formula: 'SUM(F' + first + ':F' + (r - 1) + ')' };
    tc.font = SF_HDR; tc.numFmt = SUM_MONEY;
    tc.alignment = { horizontal: 'right', vertical: 'center' };
    tc.border = { top: THIN, bottom: THIN };
    subtotals.push(r);
    r += 2;
  }

  // Grand total: label merged A:E, amount in F with a double bottom rule.
  r += 1;
  s.mergeCells('A' + r + ':E' + r);
  const gr = s.getCell('A' + r);
  gr.value = 'Total fees and expenses to GP/Affiliates'; gr.font = SF_HDR;
  gr.alignment = { horizontal: 'left', vertical: 'center' };
  s.getRow(r).height = 15;
  const gc = s.getCell('F' + r);
  gc.value = { formula: subtotals.map((x) => 'F' + x).join('+') };
  gc.font = SF_HDR; gc.numFmt = SUM_MONEY;
  gc.alignment = { horizontal: 'right', vertical: 'center' };
  gc.border = { bottom: { style: 'double' } };

  r += 2;
  s.mergeCells('A' + r + ':F' + r);
  const noteCell = s.getCell('A' + r);
  noteCell.value = "Note: None of the fees presented above is applied to offset the Fund's Management "
    + 'Fee payable to the General Partner.';
  noteCell.font = SF({ size: 9, italic: true });
  noteCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  s.getRow(r).height = 18;

  // ── Fee Reasonableness ────────────────────────────────────────────────────
  fr.getCell('A1').value = 'MANAGEMENT FEE - EFFECTIVE RATE vs STATED BASIS';
  fr.getCell('A1').font = F({ size: 12, bold: true });
  fr.getCell('A2').value = 'Operating revenue (GL 401xx / 411xx) as an EGI proxy. Excludes interest income and '
    + 'unrealized gains, which the stated basis excludes. This is NOT a recomputation of the fee - the '
    + 'annualized-EGI tiering and the rail / non-rail split do not exist in the GL.';
  fr.getCell('A2').font = SMALLI;
  fr.getCell('A2').alignment = { wrapText: true, vertical: 'top' };
  fr.getRow(2).height = 30;
  headerRow(fr, 4, ['Entity', 'Account', 'Account name', 'Revenue'], [16, 12, 40, 18]);
  r = 5;
  const revTot = {};
  for (const p of data.properties) {
    const first = r;
    for (const x of p.revenue) {
      const row = fr.getRow(r);
      row.getCell(1).value = p.label; row.getCell(1).font = F();
      row.getCell(2).value = x.code; row.getCell(2).font = F();
      row.getCell(3).value = x.name; row.getCell(3).font = F();
      const c = row.getCell(4); c.value = x.amount; c.font = BLUE; c.numFmt = MONEY;
      r += 1;
    }
    const tr = fr.getRow(r);
    tr.getCell(3).value = 'Total ' + p.label + ' operating revenue';
    tr.getCell(3).font = F({ bold: true });
    const tc = tr.getCell(4);
    tc.value = r > first ? { formula: 'SUM(D' + first + ':D' + (r - 1) + ')' } : 0;
    tc.font = F({ bold: true }); tc.numFmt = MONEY; tc.border = { top: { style: 'thin' } };
    revTot[p.label] = r;
    r += 1;
  }
  r += 1;
  headerRow(fr, r, ['Entity', 'Operating revenue', 'Management fee', 'Effective rate'], null);
  r += 1;
  for (const p of data.properties) {
    const row = fr.getRow(r);
    row.getCell(1).value = p.label; row.getCell(1).font = F();
    let c = row.getCell(2); c.value = { formula: 'D' + revTot[p.label] }; c.font = F(); c.numFmt = MONEY;
    c = row.getCell(3);
    c.value = { formula: "SUMIFS('GL Data'!$H$" + glFirst + ':$H$' + glLast
      + ",'GL Data'!$A$" + glFirst + ':$A$' + glLast + ',$A' + r
      + ",'GL Data'!$G$" + glFirst + ':$G$' + glLast + ',"Management Fee")' };
    c.font = GREEN; c.numFmt = MONEY;
    c = row.getCell(4);
    c.value = { formula: 'IF(B' + r + '=0,"n/a",C' + r + '/B' + r + ')' };
    c.font = F(); c.numFmt = PCT;
    r += 1;
  }

  // ── Notes & Sources ───────────────────────────────────────────────────────
  nt.getColumn(1).width = 16; nt.getColumn(2).width = 116;
  [3, 4, 5].forEach((i) => { nt.getColumn(i).width = 18; });
  nt.getCell('A1').value = 'NOTES, SOURCES AND CONTROLS';
  nt.getCell('A1').font = F({ size: 12, bold: true });
  const notes = [
    ['SOURCE', ''],
    ['', 'Generated from CloudLedger on ' + new Date().toISOString().slice(0, 10)
      + '. No pasted exports and no manual re-keying.'],
    ['', 'One quarterly trial balance per entity: as_of ' + q.end + ' with the P&L cut off at ' + q.start
      + '. Balance-sheet accounts show their closing balance on the quarter end; P&L accounts show activity from '
      + 'the quarter start.'],
    ['', 'Portfolio companies: ' + data.properties.map((p) => p.entity_name + ' (' + p.entity_id + ')').join(' | ')],
    ['', 'CLRF holds none of these fees in its own ledger. They are expensed or capitalized by the portfolio '
      + 'companies, so the schedule reads across those entities and presents the result as a CLRF deliverable.'],
    ['', ''],
    ['DEV FEE', ''],
    ['', 'The development fee is NOT the change in capitalized GL ' + DEV_ACCT + '. That balance also moves when '
      + 'completed development is allocated out to fixed assets, and when amounts are reclassed elsewhere. '
      + 'Measuring the fee as the balance movement reports (477,221.26) for CLIP in Q4 2025 against an actual fee '
      + 'of 210,424.79.'],
    ['', 'Rule applied: the fee is the debits to ' + DEV_ACCT + ', less credits whose description begins '
      + '"Reversed --". A debit is a fee billed. A "Reversed --" credit reverses a fee accrual - its offsetting '
      + 'entry is Accounts Payable - and nets against it. Every other credit transfers the fee out of the account '
      + 'and is never a fee. Every entry and its treatment is listed on the Dev Fee Detail tab.'],
    ['', 'Validated against Q4 2025, Q1 2026 and Q2 2026 for all four properties.'],
    ['', 'The rule relies on the "Reversed --" description prefix, which is a source-system convention rather than '
      + 'one CloudLedger enforces. A reversal described differently would be treated as a transfer out and '
      + 'excluded, understating the fee. The Dev Fee Detail tab makes every treatment visible so that is checkable.'],
    ['', ''],
    ['METHOD', ''],
    ['', 'The period is stated explicitly in every query, so no figure can silently be year-to-date.'],
    ['', 'The summary aggregates by ACCOUNT CODE and FEE CATEGORY, not by cell address, so an added or reordered '
      + 'account cannot repoint a reference.'],
    ['', 'Accounts in scope: 63041 management fee; 63030, 63034 and 63042 compensation reimbursement; '
      + DEV_ACCT + ' development fee. 63038 and 63047 are shown as a memo outside the totals. 63027 Sponsor '
      + 'Management Fee is excluded from this schedule.'],
    ['', 'The tiered EGI management fee and the 4% development fee are NOT recomputed. This schedule presents '
      + 'amounts as booked. Independent recomputation needs the annualized EGI by facility with a rail / non-rail '
      + 'split, and the qualifying development cost base.'],
    ['', ''],
    ['CONTROLS', ''],
    ['', 'Each quarterly trial balance was confirmed to foot:'],
  ];
  r = 3;
  for (const pair of notes) {
    if (pair[0]) { nt.getCell('A' + r).value = pair[0]; nt.getCell('A' + r).font = F({ bold: true }); }
    const c = nt.getCell('B' + r);
    c.value = pair[1]; c.font = SMALL; c.alignment = { wrapText: true, vertical: 'top' };
    if (pair[1].length > 150) nt.getRow(r).height = 42;
    r += 1;
  }
  ['Total debits', 'Total credits', 'Difference'].forEach((t, i) => {
    const c = nt.getCell(r, 3 + i); c.value = t; c.font = F({ bold: true });
  });
  r += 1;
  for (const p of data.properties) {
    nt.getCell('B' + r).value = p.entity_name + ' (' + p.entity_id + '): ' + p.account_count + ' accounts';
    nt.getCell('B' + r).font = SMALL;
    let c = nt.getCell(r, 3); c.value = p.tb_debits; c.font = BLUE; c.numFmt = MONEY;
    c = nt.getCell(r, 4); c.value = p.tb_credits; c.font = BLUE; c.numFmt = MONEY;
    c = nt.getCell(r, 5); c.value = { formula: 'C' + r + '-D' + r }; c.font = F(); c.numFmt = MONEY;
    r += 1;
  }

  for (const ws of wb.worksheets) {
    ws.views = [Object.assign({ showGridLines: false }, (ws.views && ws.views[0]) || {})];
  }
  return wb;
}

// ── Save into the entity's workpaper folder, one copy per period. Rerunning a
// quarter replaces that quarter's file and leaves other quarters alone.
const folderFor = (quarter) => 'Workpapers/GP Fees & Expenses/' + quarter.year + '/' + quarter.quarter;
const fileNameFor = (quarter) => 'CLRF_GP_Fees_' + quarter.label + '.xlsx';

function saveToWorkpapers(ctx, eid, quarter, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(quarter);
  const original = fileNameFor(quarter);
  // Register the folder and each ancestor so the tree shows them.
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  // Replace any prior copy for this same period - DB row and file on disk together,
  // so nothing is orphaned.
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_'
    + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, '
    + "uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

// Resolver so other features - the financial-statements package - can locate the
// workpaper for a period without hardcoding the path in a second place.
function findWorkpaper(ctx, eid, quarterEnd) {
  const quarter = resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ? ORDER BY id DESC LIMIT 1')
    .get(eid, folderFor(quarter), fileNameFor(quarter));
  if (!row) return null;
  return Object.assign({}, row, { quarter,
    abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerGpFeesRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/gp-fees/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const quarter = resolveQuarter((req.body && req.body.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, quarter);
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        const sumFees = (codes) => r2(data.properties.reduce((s, p) =>
          s + codes.reduce((t, c) => t + (p.fees[c] || 0), 0), 0));
        const totals = {
          management_fee: sumFees(['63041']),
          comp_reimbursement: sumFees(['63030', '63034', '63042']),
          development_fee: r2(data.properties.reduce((s, p) => s + p.dev_fee, 0)),
        };
        totals.total = r2(totals.management_fee + totals.comp_reimbursement + totals.development_fee);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-GP-Fees-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name,
          replaced: saved.replaced, totals,
          entities: data.properties.map((p) => p.entity_name + ' (' + p.entity_id + ')'),
          by_property: data.properties.map((p) => ({ label: p.label,
            mgmt: p.fees['63041'] || 0,
            comp: r2(['63030', '63034', '63042'].reduce((t, c) => t + (p.fees[c] || 0), 0)),
            dev: p.dev_fee })),
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerGpFeesRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, PROPERTIES };
