// ═══════════════════════════════════════════════════════════════════════════
// financials_xlsx — render the SAME structured statement object that the PDF
// renderer (financials.js › renderStatementsPdf) consumes into a styled .xlsx
// whose formatting MIRRORS the PDF: one worksheet per statement (Balance Sheet,
// Statements of Operations, Statement of Cash Flows, Statement of Changes in
// Members' Equity), the same titles (with "– Tax Basis" on the banyan profile),
// the same date/period heading lines, the same column headers, the same
// indentation levels, "$" on the first figure line and on totals, thin rule-
// above / rule-below on subtotals, and a double underline under grand totals.
//
// It is deliberately built from `buildStatements(...)`'s output — NOT scraped
// from the PDF — so the numbers are identical to the PDF by construction and
// every profile branch (banyan / bsfrgp / clip / silsbee / srn) renders the
// same shape the PDF does.
//
// Design notes:
//  - Amounts are written as real NUMBERS with an accounting numFmt so the sheet
//    stays analyzable (sort/sum), while displaying "(1,234.56)" like the PDF's
//    parenthesized negatives. Zero shows as "-" via the numFmt's zero section,
//    matching the PDF's dash for a zero (acct(..., {dash:true})).
//  - Indentation is emulated with Excel's cell `indent` on the label column,
//    scaled from the PDF's point indents (~6/12/16/26 pt → 0/1/2/3 indent steps).
//  - "$" is not a real prefix character (that would break the number); instead
//    a leading "$" is placed in a thin spacer column to the LEFT of the first
//    amount column on exactly the rows the PDF prefixes, so it reads the same
//    without corrupting the value.
//  - Rules (borders) are drawn only under amount columns, exactly like the PDF's
//    rule-above/rule-below/double-below semantics and like xlsxStyledReport.js.
// ═══════════════════════════════════════════════════════════════════════════
const ExcelJS = require('exceljs');

// Accounting number format: positive ; negative-in-parens ; zero-as-dash.
// Mirrors acct(v,{dash:true}) used throughout the PDF.
const MONEY_FMT = '#,##0.00;(#,##0.00);"-"';

// Indent steps from PDF point indents. The PDF uses 6/12/16/20/26/28 pt; Excel
// indent is an integer count of ~3-space stops, so map to 0..4.
function indentStep(pt) {
  if (pt >= 28) return 5;
  if (pt >= 26) return 4;
  if (pt >= 20) return 3;
  if (pt >= 16) return 2;
  if (pt >= 12) return 1;
  return 0;
}

// A row-builder that accumulates rows and per-row styling flags, so each
// statement is described declaratively and rendered once at the end. Column 0
// is the label; column 1 is the thin "$" spacer; amount columns start at 2.
function makeSheet(sheetName) {
  const rows = [];         // any[][]
  const meta = [];         // parallel: per-row { indent, bold, ruleAbove, ruleBelow, double, dollar, title, sub, header }
  const AMT0 = 2;          // first amount column index
  let amountColCount = 1;  // widened as rows are pushed

  function push(label, amounts, o = {}) {
    const a = Array.isArray(amounts) ? amounts : [];
    amountColCount = Math.max(amountColCount, a.length);
    const row = [label == null ? '' : label, o.dollar ? '$' : ''];
    for (let i = 0; i < a.length; i++) row[AMT0 + i] = a[i];
    rows.push(row);
    meta.push({
      indent: indentStep(o.indent || 0),
      bold: !!o.bold, ruleAbove: !!o.ruleAbove, ruleBelow: !!o.ruleBelow,
      double: !!o.double, dollar: !!o.dollar, title: !!o.title, sub: !!o.sub,
      header: !!o.header, gapAfter: o.gapAfter || 0,
    });
    if (o.gapAfter) { rows.push([]); meta.push({}); }
  }

  return {
    // A centered statement title block (entity name / statement / date line).
    titleBlock(lines) { lines.forEach(t => { rows.push([t]); meta.push({ center: true, titleLine: true }); }); rows.push([]); meta.push({}); },
    // A column-header row across the amount columns (underlined, right-aligned).
    colHeaders(labels) {
      const row = ['', ''];
      labels.forEach((l, i) => { row[AMT0 + i] = l; });
      amountColCount = Math.max(amountColCount, labels.length);
      rows.push(row); meta.push({ header: true });
    },
    sectionTitle(t) { rows.push([t]); meta.push({ bold: true, section: true }); },
    row: push,
    blank() { rows.push([]); meta.push({}); },
    _finish() { return { sheetName, rows, meta, AMT0, amountColCount }; },
  };
}

// Render an accumulated sheet description onto an ExcelJS worksheet.
function renderSheet(ws, built) {
  const { rows, meta, AMT0, amountColCount } = built;
  const amountCols = [];
  for (let i = 0; i < amountColCount; i++) amountCols.push(AMT0 + i);

  rows.forEach((row, r) => {
    const m = meta[r] || {};
    (row || []).forEach((v, c) => {
      if (v === '' || v == null) return;
      const cell = ws.getCell(r + 1, c + 1);
      cell.value = v;
      if (typeof v === 'number') cell.numFmt = MONEY_FMT;
    });
    // Label cell: font + indent.
    const label = ws.getCell(r + 1, 1);
    if (m.titleLine) { label.font = { bold: true, size: 12 }; label.alignment = { horizontal: 'center' }; }
    if (m.center) { label.alignment = { horizontal: 'center' }; }
    if (m.section) label.font = { bold: true };
    if (m.bold) label.font = { ...(label.font || {}), bold: true };
    if (m.indent) label.alignment = { ...(label.alignment || {}), indent: m.indent };
    // "$" spacer cell.
    if (m.dollar) { const d = ws.getCell(r + 1, 2); d.value = '$'; d.alignment = { horizontal: 'left' }; if (m.bold) d.font = { bold: true }; }
    // Amount cells: bold + right align + rules.
    for (const c of amountCols) {
      const cell = ws.getCell(r + 1, c + 1);
      cell.alignment = { ...(cell.alignment || {}), horizontal: 'right' };
      if (m.bold || m.header) cell.font = { ...(cell.font || {}), bold: true };
      const b = { ...(cell.border || {}) };
      if (m.ruleAbove) b.top = { style: 'thin' };
      if (m.header) b.bottom = { style: 'thin' };
      if (m.double) b.bottom = { style: 'double' };
      else if (m.ruleBelow) b.bottom = { style: 'thin' };
      if (b.top || b.bottom) cell.border = b;
    }
  });

  // Merge the centered title lines across the used width so they truly center.
  const width = AMT0 + amountColCount;
  rows.forEach((row, r) => {
    const m = meta[r] || {};
    if (m.titleLine || (m.center && (row && row.length === 1))) {
      try { ws.mergeCells(r + 1, 1, r + 1, width); } catch (_) {}
      ws.getCell(r + 1, 1).alignment = { horizontal: 'center' };
    }
  });

  // Column widths: wide label column, thin "$" column, roomy amount columns.
  ws.getColumn(1).width = 46;
  ws.getColumn(2).width = 2.5;
  for (const c of amountCols) ws.getColumn(c + 1).width = 16;
}

// ── money helpers mirroring financials.js acct() semantics ──────────────────
// The structured object already carries rounded numbers; write them as-is.
const num = (v) => (v == null || v === '' ? 0 : Number(v));
const chg = (cur, pri) => num(cur) - num(pri);

// ── Balance Sheet sheet ─────────────────────────────────────────────────────
function buildBalanceSheet(s) {
  const m = s.meta;
  const bs = s.balanceSheet;
  const sh = makeSheet('Balance Sheet');
  const title = m.profile === 'banyan'
    ? 'Statements of Assets, Liabilities, and Members\u2019 Equity \u2013 Tax Basis'
    : 'Balance Sheets';
  sh.titleBlock([m.entityName, title, m.longDate + ' and ' + m.priorLongDate]);
  sh.colHeaders([m.longDate, m.priorLongDate, 'Change']);
  const cells = (cur, pri) => [num(cur), num(pri), chg(cur, pri)];

  const renderSection = (sec, totalLabel, ruleBelowTotal) => {
    sh.row(sec.title, [], { indent: 6, bold: true });
    const showSub = sec.subs.length > 1 || sec.subs.some(su => su.contra) || m.profile === 'bsfrgp' || m.profile === 'banyan';
    for (const su of sec.subs) {
      if (showSub) sh.row(su.title, [], { indent: 16 });
      const rowIndent = showSub ? 26 : 16;
      for (const r of su.rows) { sh.row(r.name, cells(r.cur, r.pri), { indent: rowIndent, dollar: bsFirst.armed }); bsFirst.armed = false; }
      if (showSub && su.rows.length > 1) sh.row('Total ' + su.title, cells(su.subtotal.cur, su.subtotal.pri), { indent: 20, ruleAbove: true });
    }
    sh.row(totalLabel, cells(sec.total.cur, sec.total.pri), { indent: 6, bold: true, ruleAbove: true, ruleBelow: ruleBelowTotal, gapAfter: 1 });
  };
  const bsFirst = { armed: false };
  const RULE_BELOW = /^Total (Current Assets|Fixed Assets, Net)$/;

  sh.sectionTitle('ASSETS');
  bsFirst.armed = true;
  for (const sec of bs.assetSections) renderSection(sec, 'Total ' + sec.title, RULE_BELOW.test('Total ' + sec.title));
  sh.row('Total Assets', cells(bs.totalAssets.cur, bs.totalAssets.pri), { indent: 6, bold: true, ruleAbove: true, double: true, gapAfter: 1, dollar: true });

  sh.sectionTitle('LIABILITIES AND MEMBERS\u2019 EQUITY');
  bsFirst.armed = true;
  for (const sec of bs.liabSections) renderSection(sec, 'Total ' + sec.title, false);
  sh.row('Total Liabilities', cells(bs.totalLiab.cur, bs.totalLiab.pri), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
  sh.row('Members\u2019 Equity', [], { indent: 6, bold: true });
  for (const r of bs.equityRows) sh.row(r.name, cells(r.cur, r.pri), { indent: 16 });
  for (const r of (bs.retainedRows || [])) sh.row(r.name, cells(r.cur, r.pri), { indent: 16 });
  sh.row('Net Income (Loss)', cells(bs.niLine.cur, bs.niLine.pri), { indent: 16 });
  sh.row('Total Members\u2019 Equity', cells(bs.totalEquity.cur, bs.totalEquity.pri), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1 });
  sh.row('Total Liabilities and Members\u2019 Equity', cells(bs.totalLiabEquity.cur, bs.totalLiabEquity.pri), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  return sh._finish();
}

// ── Statements of Operations sheet ──────────────────────────────────────────
function buildOperations(s) {
  const m = s.meta;
  const op = s.operations;
  const sh = makeSheet('Statements of Operations');
  const periodWord = (m.colLabel || 'Month Ended').replace(/ Ended$/, '');
  const title = m.profile === 'banyan'
    ? 'Statements of Revenues and Expenses \u2013 Tax Basis'
    : 'Statements of Operations';
  sh.titleBlock([m.entityName, title, m.opsDateLine || ('For the ' + periodWord + 's Ended ' + m.longDate + ' and ' + m.priorLongDate)]);
  sh.colHeaders([m.longDate, m.priorLongDate, 'Change', 'Year to Date']);
  const cell4 = (t) => [num(t.cur), num(t.pri), chg(t.cur, t.pri), num(t.ytd)];
  const line = (r, o = {}) => sh.row(r.name, cell4(r), { indent: 16, ...o });

  if (op.banyan && op.banyan.structured) {
    const bo = op.banyan;
    const first = { armed: false };
    const renderTree = (groups, { showGroupTotal } = {}) => {
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        for (const su of g.subs) {
          const echo = su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          const li = echo ? 26 : 30;
          su.lines.forEach(r => { sh.row(r.name, cell4(r), { indent: li, dollar: first.armed }); first.armed = false; });
          if (su.lines.length > 1 && !echo) sh.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true });
        }
        if (showGroupTotal !== false) sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true });
      }
    };
    sh.sectionTitle('Revenue');
    first.armed = true;
    renderTree(bo.revenueTree, { showGroupTotal: true });
    first.armed = false;
    sh.row('Total Revenue', cell4(bo.totRev), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true });
    sh.row('Gross Profit', cell4(bo.grossProfit), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    sh.sectionTitle('Operating Expenses');
    renderTree(bo.opexTree, { showGroupTotal: true });
    sh.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    if (bo.otherIncomeTree.length || bo.otherExpenseTree.length) {
      sh.sectionTitle('Other Income (Expense)');
      renderTree(bo.otherIncomeTree, { showGroupTotal: true });
      renderTree(bo.otherExpenseTree, { showGroupTotal: true });
      sh.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    if (bo.incomeTaxTree.length) {
      sh.sectionTitle('Income Taxes');
      renderTree(bo.incomeTaxTree, { showGroupTotal: true });
      sh.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    sh.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  } else if (op.bsfrgp && op.bsfrgp.structured) {
    const bo = op.bsfrgp;
    const first = { armed: false };
    const renderTree = (groups, { showGroupTotal }) => {
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        for (const su of g.subs) {
          const echo = su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          su.lines.forEach(r => { sh.row(r.name, cell4(r), { indent: echo ? 26 : 30, dollar: first.armed }); first.armed = false; });
          if (su.lines.length > 1 && !echo) sh.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true });
        }
        if (showGroupTotal && (g.subs.length > 1 || g.subs.some(su => su.title === g.title))) {
          sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true });
        }
      }
    };
    sh.sectionTitle('Operating Expenses');
    first.armed = true;
    renderTree(bo.opexTree, { showGroupTotal: true });
    sh.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    sh.sectionTitle('Other Income (Expense)');
    renderTree(bo.otherIncomeTree, { showGroupTotal: true });
    renderTree(bo.otherExpenseTree, { showGroupTotal: true });
    sh.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    sh.sectionTitle('Income Taxes');
    renderTree(bo.incomeTaxTree, { showGroupTotal: true });
    sh.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    sh.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  } else {
    // The $ is armed once and spent by whichever section draws first: a
    // development entity whose only revenue account was interest income now
    // has no Revenue section at all, that account having moved into Other
    // Income (Expense) (Jimmy, 2026-08-28).
    const firstFig = { armed: true };
    const spendDollar = () => { const d = firstFig.armed; firstFig.armed = false; return d; };
    if (op.revenue.length) {
      sh.sectionTitle('Revenue');
      op.revenue.forEach(r => line(r, { dollar: spendDollar() }));
      sh.row('Total Revenue', cell4(op.totRev), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    if (op.cogs.length) {
      sh.sectionTitle('Cost of Revenue');
      op.cogs.forEach(r => line(r));
      sh.row('Total Cost of Revenue', cell4(op.totCogs), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1 });
      sh.row('Gross Profit', cell4(op.grossProfit), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    sh.sectionTitle('Operating Expenses');
    const groups = op.opexGroups && op.opexGroups.length ? op.opexGroups : null;
    if (groups) {
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        g.lines.forEach(r => sh.row(r.name, cell4(r), { indent: 26, dollar: spendDollar() }));
        if (g.lines.length > 1) sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 20, ruleAbove: true });
      }
    } else {
      op.opex.forEach(r => line(r, { dollar: spendDollar() }));
    }
    sh.row('Total Operating Expenses', cell4(op.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    // Other Income (Expense) / Income Taxes, from the shared classifier
    // (otherIeRoute in financials.js). Same nesting as the PDF; expense lines
    // print parenthesised as reductions of income.
    const oiTree = op.otherIncomeTree || [];
    const oeTree = op.otherExpenseTree || [];
    const itTree = op.incomeTaxTree || [];
    const negT = (t) => ({ cur: -num(t.cur), pri: -num(t.pri), ytd: -num(t.ytd) });
    const renderOie = (groups, opts) => {
      const o = opts || {};
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        for (const su of g.subs) {
          const echo = !!o.echoSub && su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          const li = echo ? 26 : 30;
          su.lines.forEach(r => sh.row(r.name, cell4(o.negate ? negT(r) : r), { indent: li }));
          if (!echo) sh.row('Total ' + su.title, cell4(o.negate ? negT(su.subtotal) : su.subtotal), { indent: 24, ruleAbove: true });
        }
        sh.row('Total ' + g.title, cell4(o.negate ? negT(g.subtotal) : g.subtotal), { indent: 16, ruleAbove: true });
      }
    };
    if (oiTree.length || oeTree.length) {
      sh.sectionTitle('Other Income (Expense)');
      renderOie(oiTree, { echoSub: true });
      renderOie(oeTree, { negate: true, echoSub: true });
      sh.row('Total Other Income (Expense)', cell4(op.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    if (itTree.length) {
      sh.sectionTitle('Income Taxes');
      renderOie(itTree, { echoSub: true });
      sh.row('Total Income Taxes', cell4(op.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1 });
    }
    sh.row('Net Income (Loss)', cell4(op.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  }
  return sh._finish();
}

// ── Statement of Cash Flows sheet ───────────────────────────────────────────
function buildCashFlow(s) {
  const m = s.meta;
  const cf = s.cashFlow;
  const sh = makeSheet('Cash Flows');
  const isZero = (v) => Math.abs(num(v)) < 0.005;
  const title = m.profile === 'banyan' ? 'Statement of Cash Flows \u2013 Tax Basis' : 'Statement of Cash Flows';
  sh.titleBlock([m.entityName, title, m.monthsEnded]);
  const one = (v) => [num(v)];

  sh.sectionTitle('Cash Flows from Operating Activities');
  sh.row('Net Income (Loss)', one(cf.netIncome), { indent: 16, dollar: true });
  sh.row('Adjustments to reconcile net income to net cash:', [], { indent: 16 });
  if (!isZero(cf.amortization)) sh.row('Amortization and depreciation', one(cf.amortization), { indent: 28 });
  sh.blank();
  sh.row('Changes in Operating Assets and Liabilities:', [], { indent: 16 });
  if (!isZero(cf.changeAR)) sh.row('(Increase) decrease in accounts receivable', one(cf.changeAR), { indent: 28 });
  if (!isZero(cf.changePrepaidOther)) sh.row('(Increase) decrease in prepaid and other current assets', one(cf.changePrepaidOther), { indent: 28 });
  if (!isZero(cf.changeIntercompany)) sh.row('(Increase) decrease in intercompany balances', one(cf.changeIntercompany), { indent: 28 });
  const changeApOther = num(cf.changeAP) + num(cf.changeAccrued);
  if (!isZero(changeApOther)) sh.row('Increase (decrease) in accounts payable and other current liabilities', one(changeApOther), { indent: 28 });
  sh.row('Net Cash Provided (Used) by Operating Activities', one(cf.netOperating), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1 });

  sh.sectionTitle('Cash Flows from Investing Activities');
  if (!isZero(cf.capex)) sh.row('Acquisition of fixed assets', one(cf.capex), { indent: 28 });
  if (!isZero(cf.ltInvest)) sh.row('(Increase) decrease in Other Assets', one(cf.ltInvest), { indent: 28 });
  sh.row('Net Cash Provided (Used) by Investing Activities', one(cf.netInvesting), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1 });

  sh.sectionTitle('Cash Flows from Financing Activities');
  if (!isZero(cf.equityContrib)) sh.row('Member contributions (distributions), net', one(cf.equityContrib), { indent: 28 });
  if (!isZero(cf.debtChange)) sh.row('Proceeds from (repayment of) long-term debt', one(cf.debtChange), { indent: 28 });
  sh.row('Net Cash Provided (Used) by Financing Activities', one(cf.netFinancing), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1 });

  sh.row('Net Increase (Decrease) in Cash', one(cf.netChange), { indent: 6, bold: true, ruleAbove: true });
  sh.row('Cash, Beginning of Period', one(cf.cashBeg), { indent: 6 });
  sh.row('Cash, End of Period', one(cf.cashEnd), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  if (!isZero(cf.tieOut)) {
    sh.blank();
    sh.row('Note: reconciled change differs from cash movement (see notes).', [], { indent: 6 });
  }
  return sh._finish();
}

// ── Statement of Changes in Members' Equity sheet ───────────────────────────
function buildEquity(s) {
  const m = s.meta;
  const eq = s.equity;
  const sh = makeSheet('Members Equity');
  const title = m.profile === 'banyan'
    ? 'Statement of Changes in Members\u2019 Equity \u2013 Tax Basis'
    : 'Statement of Changes in Members\u2019 Equity';
  sh.titleBlock([m.entityName, title, m.monthsEnded]);
  const shortMD = (long) => {
    const map = { January: 1, February: 2, March: 3, April: 4, May: 5, June: 6, July: 7, August: 8, September: 9, October: 10, November: 11, December: 12 };
    const mm = String(long).match(/^(\w+)\s+(\d+),\s+(\d+)$/);
    return mm ? (map[mm[1]] + '/' + mm[2] + '/' + mm[3]) : long;
  };
  const begDate = '1/1/' + String(m.asOf).slice(0, 4);
  const endDate = shortMD(m.longDate);
  sh.colHeaders(['Equity Balances at ' + begDate, 'Contributions', 'Distributions', 'Net Income (Loss)', 'Equity Balances at ' + endDate]);
  const cells = (r) => [num(r.beginning), num(r.contributions), num(r.distributions), num(r.netIncome), num(r.ending)];
  sh.row('Member', [], { indent: 6, bold: true });
  eq.rows.forEach((r, i) => sh.row(r.name, cells(r), { indent: 16, dollar: i === 0 }));
  const t = eq.totals;
  sh.row('Total', [num(t.beginning), num(t.contributions), num(t.distributions), num(t.netIncome), num(t.ending)], { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true });
  return sh._finish();
}

// Build the whole workbook (Buffer promise) from a built statements object.
async function buildStatementsWorkbook(s) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;
  const built = [buildBalanceSheet(s), buildOperations(s), buildCashFlow(s), buildEquity(s)];
  for (const b of built) {
    let name = b.sheetName.replace(/[\[\]:*?/\\]/g, ' ').slice(0, 31).trim();
    let k = 2; const base = name;
    while (wb.getWorksheet(name)) { name = base.slice(0, 28) + ' ' + k; k++; }
    const ws = wb.addWorksheet(name);
    ws.properties.defaultRowHeight = 15;
    renderSheet(ws, b);
  }
  return wb.xlsx.writeBuffer();
}

module.exports = { buildStatementsWorkbook, MONEY_FMT };
