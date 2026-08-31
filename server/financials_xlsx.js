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

  // push returns the 0-based row index it wrote, so a later total can name the
  // exact rows it sums (o.sumOf = [rowIdx, ...]). That makes every SUM correct by
  // construction — the builder knows its own tree — rather than guessed from row
  // geometry in the render layer.
  function push(label, amounts, o = {}) {
    const a = Array.isArray(amounts) ? amounts : [];
    amountColCount = Math.max(amountColCount, a.length);
    const row = [label == null ? '' : label, o.dollar ? '$' : ''];
    for (let i = 0; i < a.length; i++) row[AMT0 + i] = a[i];
    const idx = rows.length;
    rows.push(row);
    meta.push({
      indent: indentStep(o.indent || 0),
      bold: !!o.bold, ruleAbove: !!o.ruleAbove, ruleBelow: !!o.ruleBelow,
      double: !!o.double, dollar: !!o.dollar, title: !!o.title, sub: !!o.sub,
      header: !!o.header, gapAfter: o.gapAfter || 0,
      // Rows this total sums (0-based indices), if the builder declared them.
      sumOf: Array.isArray(o.sumOf) ? o.sumOf.slice() : null,
      // Horizontal SUM within this row: { fromCol, toCol, atCol } as 0-based
      // amount-column offsets (0 = first amount column). Used for the equity
      // "ending = beginning + contributions + distributions + net income" column.
      rowSum: o.rowSum || null,
    });
    if (o.gapAfter) { rows.push([]); meta.push({}); }
    return idx;
  }

  return {
    // A centered statement title block (entity name / statement / date line).
    titleBlock(lines) { lines.forEach(t => { rows.push([t]); meta.push({ center: true, titleLine: true }); }); rows.push([]); meta.push({}); },
    // A column-header row across the amount columns (underlined, right-aligned).
    colHeaders(labels) {
      const row = ['', ''];
      labels.forEach((l, i) => { row[AMT0 + i] = l; });
      amountColCount = Math.max(amountColCount, labels.length);
      rows.push(row); meta.push({ header: true, headerLabels: labels.slice() });
    },
    sectionTitle(t) { rows.push([t]); meta.push({ bold: true, section: true }); },
    row: push,
    blank() { rows.push([]); meta.push({}); },
    _finish() { return { sheetName, rows, meta, AMT0, amountColCount }; },
  };
}

// A1-style column letter for a 1-based column index.
function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Build live formulas for the amount cells so the workbook is analyzable, not a
// dead grid of numbers. Two kinds, both VERIFIED against the value the builder
// already computed so a formula can never display a figure that disagrees with
// the PDF:
//   • Change column  → =<cur> - <prior>  (exact by definition of chg()).
//   • Subtotal/total → =SUM(...) over the EXACT rows the builder declared feed
//     that total (meta.sumOf, a list of 0-based row indices). Contiguous rows
//     collapse to SUM(top:bottom); a gapped set becomes SUM(a,b,c). Because the
//     builder walks its own tree, the summands are always right; the value is
//     still re-checked to the cent as a guard, and on the rare miss the static
//     number is kept.
// Returns a Map keyed "r:c" (0-based, same basis as rows/meta) → { formula, result }.
function buildFormulaMap(rows, meta, AMT0, amountColCount) {
  const map = new Map();
  const key = (r, c) => r + ':' + c;
  const near = (a, b) => Math.abs(num(a) - num(b)) < 0.01;

  // Locate the Change column (0-based amount index) from the header row, if any.
  let changeCol = -1;
  const hdr = meta.find(m => m && m.header && Array.isArray(m.headerLabels));
  if (hdr) {
    const i = hdr.headerLabels.findIndex(l => /^change$/i.test(String(l || '').trim()));
    if (i >= 0) changeCol = AMT0 + i;
  }

  // Render a list of 0-based row indices as an Excel range argument for column
  // c: collapse maximal contiguous runs into A:B, join the rest with commas.
  // e.g. rows [8,9,10,12] in col C -> "C9:C11,C13".
  const rangeArg = (rowIdxs, c) => {
    const L = colLetter(c + 1);
    const rs = rowIdxs.slice().sort((a, b) => a - b);
    const parts = [];
    let i = 0;
    while (i < rs.length) {
      let j = i;
      while (j + 1 < rs.length && rs[j + 1] === rs[j] + 1) j++;
      parts.push(rs[i] === rs[j] ? (L + (rs[i] + 1)) : (L + (rs[i] + 1) + ':' + L + (rs[j] + 1)));
      i = j + 1;
    }
    return parts.join(',');
  };

  for (let r = 0; r < rows.length; r++) {
    const m = meta[r] || {};
    const row = rows[r] || [];

    // Change = current - prior (first two amount columns).
    if (changeCol >= 0 && typeof row[changeCol] === 'number'
        && typeof row[AMT0] === 'number' && typeof row[AMT0 + 1] === 'number') {
      const curRef = colLetter(AMT0 + 1) + (r + 1);
      const priRef = colLetter(AMT0 + 2) + (r + 1);
      map.set(key(r, changeCol), { formula: curRef + '-' + priRef, result: num(row[changeCol]) });
    }

    // Subtotal/total → SUM over the builder-declared summand rows, per amount
    // column (the Change column stays a cur-prior formula, handled above).
    if (Array.isArray(m.sumOf) && m.sumOf.length) {
      for (let c = AMT0; c < AMT0 + amountColCount; c++) {
        if (c === changeCol) continue;
        if (typeof row[c] !== 'number') continue;
        // Only include summand rows that actually carry a number in this column.
        const feed = m.sumOf.filter(ri => typeof (rows[ri] || [])[c] === 'number');
        if (!feed.length) continue;
        const acc = feed.reduce((s, ri) => s + num(rows[ri][c]), 0);
        if (!near(acc, row[c])) continue; // guard: never show a wrong figure
        map.set(key(r, c), { formula: 'SUM(' + rangeArg(feed, c) + ')', result: num(row[c]) });
      }
    }

    // Horizontal SUM within the row (e.g. equity ending = beginning + contrib +
    // distrib + net income). Columns are 0-based amount offsets.
    if (m.rowSum) {
      const from = AMT0 + m.rowSum.fromCol, to = AMT0 + m.rowSum.toCol, at = AMT0 + m.rowSum.atCol;
      if (typeof row[at] === 'number') {
        let acc = 0;
        for (let c = from; c <= to; c++) acc += num(row[c]);
        if (near(acc, row[at])) {
          const L1 = colLetter(from + 1), L2 = colLetter(to + 1);
          map.set(key(r, at), { formula: 'SUM(' + L1 + (r + 1) + ':' + L2 + (r + 1) + ')', result: num(row[at]) });
        }
      }
    }
  }
  return map;
}

// Render an accumulated sheet description onto an ExcelJS worksheet.
function renderSheet(ws, built) {
  const { rows, meta, AMT0, amountColCount } = built;
  const amountCols = [];
  for (let i = 0; i < amountColCount; i++) amountCols.push(AMT0 + i);

  const formulaMap = buildFormulaMap(rows, meta, AMT0, amountColCount);

  rows.forEach((row, r) => {
    const m = meta[r] || {};
    (row || []).forEach((v, c) => {
      if (v === '' || v == null) return;
      const cell = ws.getCell(r + 1, c + 1);
      const f = formulaMap.get(r + ':' + c);
      if (f && typeof v === 'number') { cell.value = { formula: f.formula, result: f.result }; }
      else { cell.value = v; }
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
// Single-member LLCs (SRN / SABINERI and Buna / CLRBUNAP) read "Member\u2019s
// Equity" (singular possessive). Pinned by code and raw name.
function isSingleMember(m) {
  return ['SABINERI', 'CLRBUNAP'].includes(String(m.entityCode || '').toUpperCase())
    || /sabine|county\s*line\s*srn|\bbuna\b/i.test(m.rawEntityName || '');
}
function meEquityLabel(m) { return isSingleMember(m) ? 'Member\u2019s Equity' : 'Members\u2019 Equity'; }

function buildBalanceSheet(s) {
  const m = s.meta;
  const meEquity = meEquityLabel(m);
  const bs = s.balanceSheet;
  const sh = makeSheet('Balance Sheet');
  const title = m.profile === 'banyan'
    ? 'Statements of Assets, Liabilities, and ' + meEquity + ' \u2013 Tax Basis'
    : 'Balance Sheets';
  sh.titleBlock([m.entityName, title, m.longDate + ' and ' + m.priorLongDate]);
  sh.colHeaders([m.longDate, m.priorLongDate, 'Change']);
  const cells = (cur, pri) => [num(cur), num(pri), chg(cur, pri)];

  // renderSection returns the row index of the section-total line so the grand
  // total (Total Assets / Total Liabilities) can SUM the section totals.
  const renderSection = (sec, totalLabel, ruleBelowTotal) => {
    sh.row(sec.title, [], { indent: 6, bold: true });
    const showSub = sec.subs.length > 1 || sec.subs.some(su => su.contra) || m.profile === 'bsfrgp' || m.profile === 'banyan';
    const feedTotal = []; // rows the section total sums: sub-totals, or bare detail rows
    for (const su of sec.subs) {
      if (showSub) sh.row(su.title, [], { indent: 16 });
      const rowIndent = showSub ? 26 : 16;
      const detailRows = [];
      for (const r of su.rows) { detailRows.push(sh.row(r.name, cells(r.cur, r.pri), { indent: rowIndent, dollar: bsFirst.armed })); bsFirst.armed = false; }
      if (showSub && su.rows.length > 1) {
        feedTotal.push(sh.row('Total ' + su.title, cells(su.subtotal.cur, su.subtotal.pri), { indent: 20, ruleAbove: true, sumOf: detailRows }));
      } else {
        // No printed sub-total line: the section total sums these detail rows directly.
        feedTotal.push(...detailRows);
      }
    }
    return sh.row(totalLabel, cells(sec.total.cur, sec.total.pri), { indent: 6, bold: true, ruleAbove: true, ruleBelow: ruleBelowTotal, gapAfter: 1, sumOf: feedTotal });
  };
  const bsFirst = { armed: false };
  const RULE_BELOW = /^Total (Current Assets|Fixed Assets, Net)$/;

  sh.sectionTitle('ASSETS');
  bsFirst.armed = true;
  const assetTotalRows = [];
  for (const sec of bs.assetSections) assetTotalRows.push(renderSection(sec, 'Total ' + sec.title, RULE_BELOW.test('Total ' + sec.title)));
  // No rule above Total Assets when the last asset section total already carries
  // one below its figures — the two would stack and read as a stray double rule
  // (Jimmy, 2026-08-30; County Line Rail Operations, a one-section balance sheet).
  const lastAsset = bs.assetSections[bs.assetSections.length - 1];
  const assetsRuledBelow = !!lastAsset && RULE_BELOW.test('Total ' + lastAsset.title);
  sh.row('Total Assets', cells(bs.totalAssets.cur, bs.totalAssets.pri), { indent: 6, bold: true, ruleAbove: !assetsRuledBelow, double: true, gapAfter: 1, dollar: true, sumOf: assetTotalRows });

  sh.sectionTitle('LIABILITIES AND ' + meEquity.toUpperCase());
  bsFirst.armed = true;
  const liabTotalRows = [];
  for (const sec of bs.liabSections) liabTotalRows.push(renderSection(sec, 'Total ' + sec.title, false));
  const totalLiabRow = sh.row('Total Liabilities', cells(bs.totalLiab.cur, bs.totalLiab.pri), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: liabTotalRows });
  sh.row(meEquity, [], { indent: 6, bold: true });
  const equityFeed = [];
  for (const r of bs.equityRows) equityFeed.push(sh.row(r.name, cells(r.cur, r.pri), { indent: 16 }));
  for (const r of (bs.retainedRows || [])) equityFeed.push(sh.row(r.name, cells(r.cur, r.pri), { indent: 16 }));
  equityFeed.push(sh.row('Net Income (Loss)', cells(bs.niLine.cur, bs.niLine.pri), { indent: 16 }));
  const totalEquityRow = sh.row('Total ' + meEquity, cells(bs.totalEquity.cur, bs.totalEquity.pri), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1, sumOf: equityFeed });
  // Total Liabilities and Members' Equity = Total Liabilities + Total Members' Equity.
  sh.row('Total Liabilities and ' + meEquity, cells(bs.totalLiabEquity.cur, bs.totalLiabEquity.pri), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: [totalLiabRow, totalEquityRow] });
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
    // renderTree returns the row indices of the rows a section grand total sums:
    // each group's total row (or, when a group prints no total, that group's
    // feed rows). Sub-totals sum their own lines; group totals sum their subs.
    const renderTree = (groups, { showGroupTotal } = {}) => {
      const sectionFeed = [];
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        const groupFeed = [];
        for (const su of g.subs) {
          const echo = su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          const li = echo ? 26 : 30;
          const lineRows = [];
          su.lines.forEach(r => { lineRows.push(sh.row(r.name, cell4(r), { indent: li, dollar: first.armed })); first.armed = false; });
          if (su.lines.length > 1 && !echo) {
            groupFeed.push(sh.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true, sumOf: lineRows }));
          } else {
            groupFeed.push(...lineRows);
          }
        }
        if (showGroupTotal !== false) {
          sectionFeed.push(sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true, sumOf: groupFeed }));
        } else {
          sectionFeed.push(...groupFeed);
        }
      }
      return sectionFeed;
    };
    sh.sectionTitle('Revenue');
    first.armed = true;
    const revFeed = renderTree(bo.revenueTree, { showGroupTotal: true });
    first.armed = false;
    const totRevRow = sh.row('Total Revenue', cell4(bo.totRev), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, sumOf: revFeed });
    // Gross Profit = Total Revenue (no COGS section on this profile).
    const grossRow = sh.row('Gross Profit', cell4(bo.grossProfit), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: [totRevRow] });
    sh.sectionTitle('Operating Expenses');
    const opexFeed = renderTree(bo.opexTree, { showGroupTotal: true });
    const totOpexRow = sh.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: opexFeed });
    const netFeed = [grossRow, totOpexRow]; // NI = Gross Profit - Opex + OtherIE + Taxes (signs already in the figures)
    if (bo.otherIncomeTree.length || bo.otherExpenseTree.length) {
      sh.sectionTitle('Other Income (Expense)');
      const oiFeed = renderTree(bo.otherIncomeTree, { showGroupTotal: true });
      const oeFeed = renderTree(bo.otherExpenseTree, { showGroupTotal: true });
      const totOtherRow = sh.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: oiFeed.concat(oeFeed) });
      netFeed.push(totOtherRow);
    }
    if (bo.incomeTaxTree.length) {
      sh.sectionTitle('Income Taxes');
      const itFeed = renderTree(bo.incomeTaxTree, { showGroupTotal: true });
      const totTaxRow = sh.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: itFeed });
      netFeed.push(totTaxRow);
    }
    // Net Income ties out as Gross Profit − Operating Expenses (+ Other, − Taxes),
    // which is exactly the sum of those subtotal rows because expenses/taxes are
    // carried as reductions in the figures. The value guard keeps it honest.
    sh.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: netFeed });
  } else if (op.bsfrgp && op.bsfrgp.structured) {
    const bo = op.bsfrgp;
    const first = { armed: false };
    const renderTree = (groups, { showGroupTotal }) => {
      const sectionFeed = [];
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        const groupFeed = [];
        for (const su of g.subs) {
          const echo = su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          const lineRows = [];
          su.lines.forEach(r => { lineRows.push(sh.row(r.name, cell4(r), { indent: echo ? 26 : 30, dollar: first.armed })); first.armed = false; });
          if (su.lines.length > 1 && !echo) {
            groupFeed.push(sh.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true, sumOf: lineRows }));
          } else {
            groupFeed.push(...lineRows);
          }
        }
        if (showGroupTotal && (g.subs.length > 1 || g.subs.some(su => su.title === g.title))) {
          sectionFeed.push(sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true, sumOf: groupFeed }));
        } else {
          sectionFeed.push(...groupFeed);
        }
      }
      return sectionFeed;
    };
    sh.sectionTitle('Operating Expenses');
    first.armed = true;
    const opexFeed = renderTree(bo.opexTree, { showGroupTotal: true });
    const totOpexRow = sh.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: opexFeed });
    sh.sectionTitle('Other Income (Expense)');
    const oiFeed = renderTree(bo.otherIncomeTree, { showGroupTotal: true });
    const oeFeed = renderTree(bo.otherExpenseTree, { showGroupTotal: true });
    const totOtherRow = sh.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: oiFeed.concat(oeFeed) });
    sh.sectionTitle('Income Taxes');
    const itFeed = renderTree(bo.incomeTaxTree, { showGroupTotal: true });
    const totTaxRow = sh.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: itFeed });
    sh.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: [totOpexRow, totOtherRow, totTaxRow] });
  } else {
    // The $ is armed once and spent by whichever section draws first: a
    // development entity whose only revenue account was interest income now
    // has no Revenue section at all, that account having moved into Other
    // Income (Expense) (Jimmy, 2026-08-28).
    const firstFig = { armed: true };
    const spendDollar = () => { const d = firstFig.armed; firstFig.armed = false; return d; };
    const netFeed = []; // subtotal rows Net Income sums
    if (op.revenue.length) {
      sh.sectionTitle('Revenue');
      const revRows = op.revenue.map(r => line(r, { dollar: spendDollar() }));
      netFeed.push(sh.row('Total Revenue', cell4(op.totRev), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: revRows }));
    }
    if (op.cogs.length) {
      sh.sectionTitle('Cost of Revenue');
      const cogsRows = op.cogs.map(r => line(r));
      const totCogsRow = sh.row('Total Cost of Revenue', cell4(op.totCogs), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1, sumOf: cogsRows });
      // Gross Profit = Total Revenue − Total COGS. Its summand rows are those two
      // totals (the COGS figure is a reduction, so a straight SUM ties out only
      // when totCogs is carried negative; the value guard drops the formula if not,
      // leaving the correct static number).
      const grossFeed = netFeed.length ? [netFeed[netFeed.length - 1], totCogsRow] : [totCogsRow];
      netFeed.length = 0;
      netFeed.push(sh.row('Gross Profit', cell4(op.grossProfit), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: grossFeed }));
    }
    sh.sectionTitle('Operating Expenses');
    const groups = op.opexGroups && op.opexGroups.length ? op.opexGroups : null;
    const opexFeed = [];
    if (groups) {
      for (const g of groups) {
        sh.row(g.title, [], { indent: 12, bold: true });
        const gLines = g.lines.map(r => sh.row(r.name, cell4(r), { indent: 26, dollar: spendDollar() }));
        if (g.lines.length > 1) opexFeed.push(sh.row('Total ' + g.title, cell4(g.subtotal), { indent: 20, ruleAbove: true, sumOf: gLines }));
        else opexFeed.push(...gLines);
      }
    } else {
      op.opex.forEach(r => opexFeed.push(line(r, { dollar: spendDollar() })));
    }
    netFeed.push(sh.row('Total Operating Expenses', cell4(op.totOpex), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: opexFeed }));
    // Other Income (Expense) / Income Taxes, from the shared classifier
    // (otherIeRoute in financials.js). Same nesting as the PDF; expense lines
    // print parenthesised as reductions of income.
    const oiTree = op.otherIncomeTree || [];
    const oeTree = op.otherExpenseTree || [];
    const itTree = op.incomeTaxTree || [];
    const negT = (t) => ({ cur: -num(t.cur), pri: -num(t.pri), ytd: -num(t.ytd) });
    const renderOie = (grps, opts) => {
      const o = opts || {};
      const sectionFeed = [];
      for (const g of grps) {
        sh.row(g.title, [], { indent: 12, bold: true });
        const groupFeed = [];
        for (const su of g.subs) {
          const echo = !!o.echoSub && su.title === g.title;
          if (!echo) sh.row(su.title, [], { indent: 20 });
          const li = echo ? 26 : 30;
          const lineRows = su.lines.map(r => sh.row(r.name, cell4(o.negate ? negT(r) : r), { indent: li }));
          if (!echo) groupFeed.push(sh.row('Total ' + su.title, cell4(o.negate ? negT(su.subtotal) : su.subtotal), { indent: 24, ruleAbove: true, sumOf: lineRows }));
          else groupFeed.push(...lineRows);
        }
        sectionFeed.push(sh.row('Total ' + g.title, cell4(o.negate ? negT(g.subtotal) : g.subtotal), { indent: 16, ruleAbove: true, sumOf: groupFeed }));
      }
      return sectionFeed;
    };
    if (oiTree.length || oeTree.length) {
      sh.sectionTitle('Other Income (Expense)');
      const oiFeed = renderOie(oiTree, { echoSub: true });
      const oeFeed = renderOie(oeTree, { negate: true, echoSub: true });
      netFeed.push(sh.row('Total Other Income (Expense)', cell4(op.totOtherIE), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: oiFeed.concat(oeFeed) }));
    }
    if (itTree.length) {
      sh.sectionTitle('Income Taxes');
      const itFeed = renderOie(itTree, { echoSub: true });
      netFeed.push(sh.row('Total Income Taxes', cell4(op.totIncomeTax), { indent: 6, bold: true, ruleAbove: true, ruleBelow: true, gapAfter: 1, sumOf: itFeed }));
    }
    sh.row('Net Income (Loss)', cell4(op.netIncome), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: netFeed });
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
  const opFeed = [];
  opFeed.push(sh.row('Net Income (Loss)', one(cf.netIncome), { indent: 16, dollar: true }));
  sh.row('Adjustments to reconcile net income to net cash:', [], { indent: 16 });
  if (!isZero(cf.amortization)) opFeed.push(sh.row('Amortization and depreciation', one(cf.amortization), { indent: 28 }));
  sh.blank();
  sh.row('Changes in Operating Assets and Liabilities:', [], { indent: 16 });
  if (!isZero(cf.changeAR)) opFeed.push(sh.row('(Increase) decrease in accounts receivable', one(cf.changeAR), { indent: 28 }));
  if (!isZero(cf.changePrepaidOther)) opFeed.push(sh.row('(Increase) decrease in prepaid and other current assets', one(cf.changePrepaidOther), { indent: 28 }));
  if (!isZero(cf.changeIntercompany)) opFeed.push(sh.row('(Increase) decrease in intercompany balances', one(cf.changeIntercompany), { indent: 28 }));
  const changeApOther = num(cf.changeAP) + num(cf.changeAccrued);
  if (!isZero(changeApOther)) opFeed.push(sh.row('Increase (decrease) in accounts payable and other current liabilities', one(changeApOther), { indent: 28 }));
  const netOpRow = sh.row('Net Cash Provided (Used) by Operating Activities', one(cf.netOperating), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1, sumOf: opFeed });

  sh.sectionTitle('Cash Flows from Investing Activities');
  const invFeed = [];
  if (!isZero(cf.capex)) invFeed.push(sh.row('Acquisition of fixed assets', one(cf.capex), { indent: 28 }));
  if (!isZero(cf.ltInvest)) invFeed.push(sh.row('(Increase) decrease in Other Assets', one(cf.ltInvest), { indent: 28 }));
  const netInvRow = sh.row('Net Cash Provided (Used) by Investing Activities', one(cf.netInvesting), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1, sumOf: invFeed });

  sh.sectionTitle('Cash Flows from Financing Activities');
  const finFeed = [];
  if (!isZero(cf.equityContrib)) finFeed.push(sh.row('Member contributions (distributions), net', one(cf.equityContrib), { indent: 28 }));
  if (!isZero(cf.debtChange)) finFeed.push(sh.row('Proceeds from (repayment of) long-term debt', one(cf.debtChange), { indent: 28 }));
  const netFinRow = sh.row('Net Cash Provided (Used) by Financing Activities', one(cf.netFinancing), { indent: 6, bold: true, ruleAbove: true, gapAfter: 1, sumOf: finFeed });

  const netChangeRow = sh.row('Net Increase (Decrease) in Cash', one(cf.netChange), { indent: 6, bold: true, ruleAbove: true, sumOf: [netOpRow, netInvRow, netFinRow] });
  const cashBegRow = sh.row('Cash, Beginning of Period', one(cf.cashBeg), { indent: 6 });
  sh.row('Cash, End of Period', one(cf.cashEnd), { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: [netChangeRow, cashBegRow] });
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
  const meEquity = meEquityLabel(m);
  const sh = makeSheet(isSingleMember(m) ? "Member's Equity" : 'Members Equity');
  const title = m.profile === 'banyan'
    ? 'Statement of Changes in ' + meEquity + ' \u2013 Tax Basis'
    : 'Statement of Changes in ' + meEquity;
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
  // Each member's ending balance = beginning + contributions + distributions +
  // net income (contributions/distributions carry their own sign), so the ending
  // column is a horizontal SUM of the four columns to its left (rowSum). The
  // Total row is the vertical SUM of the member rows.
  const memberRows = eq.rows.map((r, i) => sh.row(r.name, cells(r), { indent: 16, dollar: i === 0, rowSum: { fromCol: 0, toCol: 3, atCol: 4 } }));
  const t = eq.totals;
  sh.row('Total', [num(t.beginning), num(t.contributions), num(t.distributions), num(t.netIncome), num(t.ending)], { indent: 6, bold: true, ruleAbove: true, double: true, dollar: true, sumOf: memberRows, rowSum: { fromCol: 0, toCol: 3, atCol: 4 } });
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
