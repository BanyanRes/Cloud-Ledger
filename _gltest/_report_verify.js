// Fixture test: build SRN statements from mock GL, render full package, verify.
const fin = require('../server/financials.js');
const fs = require('fs');

// Mock GL rows keyed to the SRN chart. balance is the natural signed balance.
// We craft a tiny but tying balance sheet + a little P&L so all statements render.
const ACCTS = {
  // Assets
  '10162': { name: 'SRNR Operating x3505', type: 'Asset' },
  '10163': { name: 'SRNR Money Market x9021', type: 'Asset' },
  '12000': { name: 'Accounts Receivable', type: 'Asset' },
  '18311': { name: 'Due from County Line RR', type: 'Asset' },
  '13001': { name: 'Prepaid Insurance', type: 'Asset' },
  '15100': { name: 'Land', type: 'Asset' },
  '15165': { name: 'Railroad Track', type: 'Asset' },
  '11009': { name: 'Track Rights Intangible', type: 'Asset' },
  '16600': { name: 'Accumulated Amortization', type: 'Asset' }, // contra
  '11011': { name: 'Long Term Investment A', type: 'Asset' },
  '12600': { name: 'Capitalized Dev Costs', type: 'Asset' },
  // Liabilities
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '21006': { name: 'Accrued Expenses', type: 'Liability' },
  '25063': { name: 'Note Payable - BOT', type: 'Liability' },
  // Equity
  '34006': { name: 'Contributed Capital - Class A', type: 'Equity' },
  '34165': { name: 'Contributed Capital - Class B', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  // P&L
  '40000': { name: 'Rail Revenue', type: 'Revenue' },
  '63000': { name: 'Accounting', type: 'Expense' },
  '60000': { name: 'Salaries & Wages', type: 'Expense' },
  '61150': { name: 'Utilities', type: 'Expense' },
  '63041': { name: 'CLRO Management Fees', type: 'Expense' },
};

// Balance-sheet balances (current period). Must tie: assets = liab + equity.
// Assets: cash 300 + AR 50 + due 40 + prepaid 10 + land 500 + track 400
//         + intangible 100 - accum amort 20 + LT invest 200 + dev 1000 = 2580
// Liab: AP 80 + accrued 20 + note 900 = 1000
// Equity: contributed 34006 900 + 34165 500 + RE 200 = 1600; NI line adds -20
//   → total equity = 1580; liab+equity = 2580  ✓
const CUR = {
  '10162': 200, '10163': 100, '12000': 50, '18311': 40, '13001': 10,
  '15100': 500, '15165': 400, '11009': 100, '16600': -20, '11011': 200, '12600': 1000,
  '20000': 80, '21006': 20, '25063': 900,
  '34006': 900, '34165': 500, '39000': 200,
};
// Prior period (slightly different, still ties): reduce cash & AP by 10.
const PRI = Object.assign({}, CUR, { '10162': 210, '20000': 90 });
// Opening (year start): contributed capital lower so there's a financing inflow.
const OPEN = Object.assign({}, CUR, { '34006': 800, '10162': 100 });

// P&L YTD: revenue 100, expenses 120 → NI -20 (matches niLine on BS).
const PL_YTD = { '40000': 100, '63000': 30, '60000': 60, '61150': 10, '63041': 20 };
const PL_CUR = { '40000': 25, '63000': 8, '60000': 15, '61150': 3, '63041': 5 };
const PL_PRI = { '40000': 20, '63000': 6, '60000': 15, '61150': 2, '63041': 5 };

function rowsFrom(balMap) {
  return Object.keys(balMap).map(code => ({
    code, name: ACCTS[code].name, type: ACCTS[code].type,
    subtype: '', balance: balMap[code],
  }));
}

// getBalances mock: dispatch on the args shape buildStatements uses.
function getBalances(o) {
  // P&L windows: from/to present
  if (o.from && o.to) {
    // YTD window is ys..asOf (2026-01-01..2026-04-30); current month Apr; prior Mar.
    if (o.from === '2026-01-01' && o.to === '2026-04-30') return Promise.resolve(rowsFrom(PL_YTD));
    if (o.from === '2026-04-01') return Promise.resolve(rowsFrom(PL_CUR));
    if (o.from === '2026-03-01') return Promise.resolve(rowsFrom(PL_PRI));
    // prior-year YTD for BS prior NI line, and any other → small P&L
    return Promise.resolve(rowsFrom(PL_PRI));
  }
  // Balance-sheet snapshots: as_of present
  if (o.as_of === '2026-04-30') return Promise.resolve(rowsFrom(CUR).concat(rowsFrom(PL_YTD)));
  if (o.as_of === '2026-03-31') return Promise.resolve(rowsFrom(PRI).concat(rowsFrom(PL_PRI)));
  // opening (year start - 1 day = 2025-12-31)
  if (o.as_of === '2025-12-31') return Promise.resolve(rowsFrom(OPEN));
  return Promise.resolve(rowsFrom(CUR));
}

(async () => {
  const s = await fin.buildStatements(getBalances, {
    asOf: '2026-04-30', period: 'monthly',
    entityName: 'Sabine River & Northern Railroad',
  });

  // ── Assertions ───────────────────────────────────────────────────────────
  const A = [];
  const check = (name, cond) => { A.push({ name, ok: !!cond }); };

  check('entityName normalized to County Line SRN', s.meta.entityName === 'County Line SRN');
  check('longDate is April 30, 2026', s.meta.longDate === 'April 30, 2026');
  check('priorLongDate is March 31, 2026', s.meta.priorLongDate === 'March 31, 2026');
  check('balance sheet ties (assets = liab+equity)', s.checks.balanceSheetTies);
  check('BS diff is 0', Math.abs(s.checks.balanceSheetDiff) < 0.005);

  // Grouping: no account appears in more than one (section,sub). Reconstruct
  // codes from the built sections and ensure the reference groups exist and
  // there's no overlap.
  const seen = new Map();
  let overlap = false;
  const sections = [...s.balanceSheet.assetSections, ...s.balanceSheet.liabSections];
  const sectionTitles = sections.map(x => x.title);
  for (const sec of sections) for (const su of sec.subs) for (const r of su.rows) {
    if (seen.has(r.code)) overlap = true;
    seen.set(r.code, sec.title + ' / ' + su.title);
  }
  check('no account overlaps across BS groups', !overlap);
  check('has Current Assets section', sectionTitles.includes('Current Assets'));
  check('has Fixed Assets, Net section', sectionTitles.includes('Fixed Assets, Net'));
  check('has Intangible Assets, Net section', sectionTitles.includes('Intangible Assets, Net'));
  check('has Investments section', sectionTitles.includes('Investments'));
  check('has Other Assets section', sectionTitles.includes('Other Assets'));
  check('has Current Liabilities section', sectionTitles.includes('Current Liabilities'));
  check('has Long Term Liabilities section', sectionTitles.includes('Long Term Liabilities'));
  // Cash grouping: 10162 & 10163 under Cash and Cash Equivalents
  check('10162 in Cash and Cash Equivalents', seen.get('10162') === 'Current Assets / Cash and Cash Equivalents');
  check('12600 in Other Assets', seen.get('12600') === 'Other Assets / Other Assets');
  check('25063 in Long Term Liabilities / Loans', seen.get('25063') === 'Long Term Liabilities / Loans');

  // Equity data carries a distributions field for the 5-column render.
  check('equity rows have distributions field', s.equity.rows.every(r => 'distributions' in r));
  check('equity totals have distributions field', 'distributions' in s.equity.totals);
  check('equity Net Income (Loss) present on RE row',
    s.equity.rows.some(r => /retained earning/i.test(r.name) && Math.abs(r.netIncome + 20) < 0.005));

  // Render statements PDF with offsets (per-statement start pages)
  const { PDFDocument } = require('pdf-lib');
  const offsets = [];
  const stmtBytes = await fin.renderStatementsPdf(s, offsets);
  check('statements PDF non-empty', stmtBytes && stmtBytes.length > 1000);
  check('offsets captured for all 4 statements', offsets.length === 4);
  check('offsets include Balance Sheets', offsets.some(o => o.label === 'Balance Sheets'));

  // Inspect the statements PDF: the Members' Equity page must be landscape
  // (width > height). Find it by its offset page index.
  const stmtDoc = await PDFDocument.load(stmtBytes);
  const eqOff = offsets.find(o => /Members/.test(o.label));
  check('equity offset found', !!eqOff);
  if (eqOff) {
    const eqPage = stmtDoc.getPage(eqOff.page);
    const { width, height } = eqPage.getSize();
    check('Members Equity page is landscape (w>h)', width > height);
    check('Members Equity landscape width ~792', Math.abs(width - 792) < 1);
  }
  // All other statement pages should be portrait.
  let portraitOK = true;
  for (let i = 0; i < stmtDoc.getPageCount(); i++) {
    if (eqOff && i === eqOff.page) continue;
    const { width, height } = stmtDoc.getPage(i).getSize();
    if (width > height) portraitOK = false;
  }
  check('non-equity statement pages are portrait', portraitOK);

  // Build a tiny mock executive-summary PDF (2 pages) and a mock B2A xlsx so the
  // full package exercises the requisition path and TOC page references.
  const XLSX = require('xlsx');
  const execDoc = await PDFDocument.create();
  execDoc.addPage([612, 792]); execDoc.addPage([612, 792]);
  const execBytes = await execDoc.save();
  const wb = XLSX.utils.book_new();
  const wsData = [['Budget to Actual', '', ''], ['Cost Code', 'Budget', 'Actual'],
                  ['Sitework', '100,000.00', '92,500.00'], ['Total', '100,000.00', '92,500.00']];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, 'Budget to Actual');
  const xlsxBuf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  // Full package with exec summary + B2A workbook.
  const pkg = await fin.generatePackage({
    statements: s, execSummaryBytes: execBytes,
    reqReportBytes: xlsxBuf, reqReportName: 'requisition.xlsx',
  });
  check('package produced bytes', pkg.bytes && pkg.bytes.length > 1000);

  // TOC entries with page numbers
  const toc = pkg.info.tocEntries || [];
  check('TOC has entries with page numbers', toc.length >= 5 && toc.every(e => Number.isInteger(e.page)));
  check('TOC includes Executive Summary at page 3', toc.some(e => e.label === 'Executive Summary' && e.page === 3));
  check('TOC includes Balance Sheets', toc.some(e => e.label === 'Balance Sheets'));
  check('TOC includes Budget to Actual', toc.some(e => e.label === 'Budget to Actual'));
  check('TOC page numbers non-decreasing',
    toc.every((e, i) => i === 0 || e.page >= toc[i - 1].page));

  // Load merged PDF, count pages
  const merged = await PDFDocument.load(pkg.bytes);
  const nPages = merged.getPageCount();
  check('package has cover + TOC + exec + statements + B2A (>= 8 pages)', nPages >= 8);
  check('cover page portrait', (() => { const p = merged.getPage(0).getSize(); return p.width < p.height; })());

  // Save the package for manual inspection
  const OUT = 'C:/Users/JimmyYun/Downloads/_srn_report_verify.pdf';
  fs.writeFileSync(OUT, pkg.bytes);

  console.log('\\n=== RESULTS ===');
  let pass = 0;
  for (const a of A) { console.log((a.ok ? 'PASS' : 'FAIL') + '  ' + a.name); if (a.ok) pass++; }
  console.log('\\n' + pass + '/' + A.length + ' passed');
  console.log('pages in package:', nPages);
  console.log('sections:', JSON.stringify(pkg.info.sections));
  console.log('saved:', OUT);
  if (pass !== A.length) process.exit(1);
})().catch(e => { console.error('ERROR', e); process.exit(1); });
