// Fixture test for the CLA (Dennis Arada) 8/17/2026 review changes on the
// Banyan Residential LLC package. Builds statements from a mock Banyan GL,
// renders the FULL merged package, then reads the PDF back and asserts on what
// actually landed on the page — TOC labels, "$" placement, the cash-flow AP
// label, and page numbers (bottom-right, cover + TOC unnumbered).
//
// "$" is asserted by COORDINATE, not by text order: layout.row() draws the value
// before the "$", so in extraction order a row's "$" trails its own figures and
// would look like it belongs to the next line. We match the "$" to the row whose
// label shares its baseline (y), and require it to sit left of the first figure.
const fin = require('../server/financials.js');
const RM = require('./_rowmodel.js');
const { PDFDocument } = require('pdf-lib');
const fs = require('fs');

const ACCTS = {
  // Current Assets → Cash and Cash Equivalents  (first ASSETS figure line)
  '10143': { name: 'UBS Entity 100 Banyan - 2472', type: 'Asset' },
  '10144': { name: 'MapleMark Entity 100 Banyan - 8868', type: 'Asset' },
  '10300': { name: 'Bill.Com Clearing Out Banyan Residential Entity 100', type: 'Asset' },
  '12000': { name: 'Accounts Receivable', type: 'Asset' },
  '13002': { name: 'Prepaid Insurance', type: 'Asset' },
  '13003': { name: 'Prepaid Rent', type: 'Asset' },
  '15500': { name: 'Furniture and Fixtures', type: 'Asset' },
  '16500': { name: 'Accumulated Depreciation', type: 'Asset' }, // contra
  '11760': { name: 'Soft Costs - Architecture', type: 'Asset' },
  '11030': { name: 'Land', type: 'Asset' },
  '12720': { name: 'Other Development Costs', type: 'Asset' },
  '18311': { name: 'Due from Banyan SFR GP', type: 'Asset' },
  // Current Liabilities → Accounts Payable  (first LIABILITIES figure line)
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '20500': { name: 'Credit Card Payable', type: 'Liability' },
  '21112': { name: 'Accrued Liabilities', type: 'Liability' },
  // Equity ('39000' Retained Earnings is solved for, see balanceRE)
  '33104': { name: 'Distribution - Brosseau Children’s Trust', type: 'Equity' },
  '34004': { name: 'Contributed Capital - Brosseau Children’s Trust', type: 'Equity' },
  '34117': { name: 'Contributed Capital - Odyssey Holdings LLC', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  // P&L — 42xxx/43xxx drive the banyan "Revenue - Services" group
  '42000': { name: 'Development Fee Revenue', type: 'Revenue' },
  '42200': { name: 'Owner’s Representation Fee Revenue', type: 'Revenue' },
  '43000': { name: 'Asset Management Fee Revenue', type: 'Revenue' },
  '60000': { name: 'Salaries and Wages', type: 'Expense' },
  '63000': { name: 'Legal and Accounting', type: 'Expense' },
  '69100': { name: 'Depreciation Expense', type: 'Expense' },
  '68061': { name: 'State Franchise Tax', type: 'Expense' },
};

// P&L windows. Depreciation (69100) is matched by a real accumulated-
// depreciation movement between the opening and closing snapshots below, so the
// add-back has a corresponding asset leg and the cash flow ties.
const PL_YTD = { '42000': 963058, '42200': 80184, '43000': 321216, '60000': 1200000, '63000': 500000, '69100': 1346, '68061': 16000 };
const PL_CUR = { '42000': 47055, '42200': 13364, '43000': 18797, '60000': 300000, '63000': 120000, '69100': 192, '68061': 2000 };
const PL_PRI = { '42000': 45524, '43000': 45924, '60000': 310000, '63000': 125000, '69100': 192, '68061': 2000 };

const sumBy = (m, type) => Object.keys(m)
  .filter(c => ACCTS[c].type === type)
  .reduce((t, c) => t + m[c], 0);

// Solve Retained Earnings so the snapshot ties: A = L + E + NI. Balances are the
// natural signed values the app reads (bal = Number(row.balance)), so assets are
// positive-debit and liabilities/equity positive-credit, per the SRN fixture.
function balanceRE(bsMap, plMap) {
  const A = sumBy(bsMap, 'Asset');
  const L = sumBy(bsMap, 'Liability');
  const E = sumBy(bsMap, 'Equity');            // 39000 not yet set
  const NI = plMap ? (sumBy(plMap, 'Revenue') - sumBy(plMap, 'Expense')) : 0;
  return Object.assign({}, bsMap, { '39000': +(A - L - E - NI).toFixed(2) });
}

// Closing balance sheet (7/31/2026).
const CUR = balanceRE({
  '10143': 2460, '10144': 100481, '10300': 7408,
  '12000': 50000, '13002': 10000, '13003': 5000,
  '15500': 8000, '16500': -1346,
  '11760': 20000, '11030': 746370, '12720': 16755,
  '18311': 40000,
  '20000': 491888, '20500': -19478, '21112': 30000,
  '33104': -13064, '34004': 20277, '34117': 746370,
}, PL_YTD);

// Prior month (6/30/2026): cash, AP and accrued all move.
const PRI = balanceRE(Object.assign({}, CUR, {
  '39000': 0,
  '10144': 145017, '20000': 236756, '20500': 118131, '21112': 146791,
}), PL_PRI);

// Opening (12/31/2025): less contributed capital (a financing inflow), higher
// AR (an operating inflow), and zero accumulated depreciation so the 1,346
// add-back has a real asset leg.
const OPEN = balanceRE(Object.assign({}, CUR, {
  '39000': 0,
  '34117': 700000, '10144': 145017, '16500': 0,
  '20000': 236756, '20500': 118131, '21112': 146791,
  '12000': 787114, '15500': 5342,
}), null);

const rowsFrom = (m) => Object.keys(m).map(code => ({
  code, name: ACCTS[code].name, type: ACCTS[code].type, subtype: '', balance: m[code],
}));

function getBalances(o) {
  if (o.from && o.to) {
    if (o.from === '2026-01-01' && o.to === '2026-07-31') return Promise.resolve(rowsFrom(PL_YTD));
    if (o.from === '2026-07-01') return Promise.resolve(rowsFrom(PL_CUR));
    if (o.from === '2026-06-01') return Promise.resolve(rowsFrom(PL_PRI));
    return Promise.resolve(rowsFrom(PL_PRI));
  }
  if (o.as_of === '2026-07-31') return Promise.resolve(rowsFrom(CUR).concat(rowsFrom(PL_YTD)));
  if (o.as_of === '2026-06-30') return Promise.resolve(rowsFrom(PRI).concat(rowsFrom(PL_PRI)));
  if (o.as_of === '2025-12-31') return Promise.resolve(rowsFrom(OPEN));
  return Promise.resolve(rowsFrom(CUR));
}

// Page-number probe: a bare integer near the bottom-right corner.
function pageNumberAt(rows, want, pageW) {
  return rows.some(r => r.items.some(i => i.s === String(want) && i.y < 50 && i.x > pageW - 90));
}

(async () => {
  const A = [];
  const check = (name, cond, detail) => { A.push({ name, ok: !!cond, detail: detail == null ? '' : String(detail) }); };

  const s = await fin.buildStatements(getBalances, {
    asOf: '2026-07-31', period: 'monthly',
    entityName: 'Banyan Residential LLC', entityCode: 'BANYANRE1',
  });

  check('profile resolved to banyan', s.meta.profile === 'banyan', s.meta.profile);
  check('fixture: balance sheet ties', s.checks.balanceSheetTies, 'diff=' + s.checks.balanceSheetDiff);
  check('fixture: cash flow ties', s.checks.cashFlowTies, 'diff=' + s.checks.cashFlowDiff);

  // ── Cash flow bucketing: the two liability lines are a genuine split ──────
  const cf = s.cashFlow;
  const expectAP = (491888 - 236756) + (-19478 - 118131);
  const expectAccrued = 30000 - 146791;
  check('AP bucket = payables only (20000 + 20500)',
    Math.abs(cf.changeAP - expectAP) < 0.005, 'changeAP=' + cf.changeAP + ' want=' + expectAP);
  check('accrued bucket = the other current liabilities (21112)',
    Math.abs(cf.changeAccrued - expectAccrued) < 0.005, 'changeAccrued=' + cf.changeAccrued);
  check('no long-term debt movement leaked into the AP line',
    Math.abs(cf.debtChange) < 0.005, 'debtChange=' + cf.debtChange);

  // ── Render the FULL merged package ───────────────────────────────────────
  const { bytes, info } = await fin.generatePackage({ statements: s });
  fs.writeFileSync(__dirname + '/_banyan_cla_aug17.pdf', Buffer.from(bytes));
  const doc = await PDFDocument.load(bytes);
  const nPages = doc.getPageCount();
  const widths = doc.getPages().map(p => p.getSize().width);
  check('package rendered', nPages >= 6, nPages + ' pages');

  const P = await RM.readRows(bytes);
  const all = P.map(RM.pageText).join('\n');

  // A statement can span multiple pages (the balance sheet runs to two) and its
  // heading repeats on each continuation page. Collect EVERY page carrying a
  // given title, from page 3 on, so the TOC — which lists the same titles —
  // is never mistaken for a statement page.
  const pagesTitled = (re) => P
    .map((rows, i) => ({ rows, i }))
    .filter(o => o.i >= 2 && re.test(RM.pageText(o.rows)))
    .map(o => o.rows);
  const bs = pagesTitled(/Statements of Assets, Liabilities/);
  const pl = pagesTitled(/Statements of Revenues and Expenses/);
  const cfp = pagesTitled(/Statement of Cash Flows/);
  const eq = pagesTitled(/Statement of Changes in Members’ Equity/);
  check('found all four statements', bs.length && pl.length && cfp.length && eq.length,
    'bs=' + bs.length + 'p pl=' + pl.length + 'p cf=' + cfp.length + 'p eq=' + eq.length + 'p');
  check('balance sheet grand totals are in scope',
    bs.some(rows => /Total Liabilities and Members’ Equity/.test(RM.pageText(rows))));

  // ── 1. Table of Contents labels ──────────────────────────────────────────
  const toc = RM.pageText(P[1] || []);
  check('TOC: plural "Statements of Assets, Liabilities, and Members’ Equity – Tax Basis"',
    /Statements of Assets, Liabilities, and Members’ Equity – Tax Basis/.test(toc));
  check('TOC: stale singular "Statement of Assets" gone',
    !/(^|[^s])Statement of Assets, Liabilities/.test(toc));
  check('TOC: "Statement of Cash Flows – Tax Basis"',
    /Statement of Cash Flows – Tax Basis/.test(toc));
  check('TOC: "Statement of Changes in Members’ Equity – Tax Basis"',
    /Statement of Changes in Members’ Equity – Tax Basis/.test(toc));
  check('TOC: "Statements of Revenues and Expenses – Tax Basis" intact',
    /Statements of Revenues and Expenses – Tax Basis/.test(toc));

  // ── 2. TOC page references land on the pages they claim ──────────────────
  const bad = [];
  for (const e of (info.tocEntries || [])) {
    if (e.label === 'Executive Summary' || e.label === 'Budget to Actual') continue;
    const t = RM.pageText(P[e.page - 1] || []);
    if (!t.includes(e.label.slice(0, 28))) bad.push(e.label + ' @ p' + e.page);
  }
  check('TOC page references resolve correctly', bad.length === 0, bad.join('; '));

  // ── 3. Page numbers: bottom-right, cover + TOC unnumbered ────────────────
  check('cover page has NO page number', !pageNumberAt(P[0], 1, widths[0]));
  check('TOC page has NO page number', !pageNumberAt(P[1], 2, widths[1]));
  const missing = [];
  for (let i = 2; i < nPages; i++) {
    if (!pageNumberAt(P[i], i + 1, widths[i])) missing.push(i + 1);
  }
  check('every page from 3 on carries its own number, bottom-right',
    missing.length === 0, 'missing on p' + missing.join(',p') + ' of ' + nPages);
  check('landscape equity page numbered against its long edge',
    (() => {
      const i = P.findIndex((p, k) => k >= 2 && /Statement of Changes in Members’ Equity/.test(RM.pageText(p)));
      return i > 0 && widths[i] > 700 && pageNumberAt(P[i], i + 1, widths[i]);
    })());
  check('info.pageNumbersFrom === 3', info.pageNumbersFrom === 3, info.pageNumbersFrom);

  // ── 4. Balance sheet: NO "As of" (Jimmy declined that one) ───────────────
  check('no "As of" prefix anywhere', !/As of July 31, 2026/.test(all));
  check('balance sheet date line still reads both dates',
    bs.some(rows => /July 31, 2026 and June 30, 2026/.test(RM.pageText(rows))));

  // ── 5. Dollar signs — the five first-line / total-line pairs
  // Matched on the row's OWN baseline and required to sit left of that row's
  // first figure. "Accounts Payable" appears twice on the balance sheet (once as
  // the subsection header, once as the account line); only the account line
  // carries figures, and figureRow() skips label-only rows.
  const d = (name, pages, label, want, which) => {
    const r = RM.dollarOn(pages, label, want, which);
    check(name, r.ok, r.why || r.detail);
  };
  d('BS: $ on first ASSETS figure line', bs, 'UBS Entity 100 Banyan - 2472', true);
  d('BS: $ on Total Assets', bs, 'Total Assets', true);
  d('BS: $ on first LIABILITIES figure line', bs, 'Accounts Payable', true);
  d('BS: $ on Total Liabilities and Members’ Equity', bs, 'Total Liabilities and Members’ Equity', true);
  d('BS: no $ on the 2nd cash line', bs, 'MapleMark Entity 100 Banyan', false);
  d('BS: no $ on Total Cash and Cash Equivalents', bs, 'Total Cash and Cash Equivalents', false);

  d('P&L: $ on first revenue line', pl, 'Development Fee Revenue', true);
  d('P&L: $ on Net Income (Loss)', pl, 'Net Income (Loss)', true);
  d('P&L: no $ on 2nd revenue line', pl, 'Owner’s Representation Fee Revenue', false);
  d('P&L: no $ on Total Revenue', pl, 'Total Revenue', false);

  d('CF: $ on Net Income (Loss)', cfp, 'Net Income (Loss)', true);
  d('CF: $ on Cash, End of Period', cfp, 'Cash, End of Period', true);
  d('CF: no $ on Cash, Beginning of Period', cfp, 'Cash, Beginning of Period', false);
  d('CF: no $ on Net Increase (Decrease) in Cash', cfp, 'Net Increase (Decrease) in Cash', false);

  d('Equity: $ on first member row', eq, 'Distribution - Brosseau', true);
  d('Equity: $ on Total row', eq, 'Total', true, 'last');
  d('Equity: no $ on 2nd member row', eq, 'Contributed Capital - Brosseau', false);

  // 6. Cash flow: ONE combined liability line (Dennis's first option, 8/17).
  //    The combined figure must equal the two buckets added together, and neither
  //    of the old split labels may survive anywhere in the package.
  const cfTxt2 = cfp.map(RM.pageText).join(' ');
  check('CF: single line reads "accounts payable and other current liabilities"',
    /Increase \(decrease\) in accounts payable and other current liabilities/.test(cfTxt2));
  check('CF: the separate accrued line is gone',
    !/Increase \(decrease\) in accrued and other liabilities/.test(all));
  check('CF: no bare "in accounts payable" line left behind',
    !/in accounts payable(?! and other current liabilities)/.test(all));
  {
    const want = cf.changeAP + cf.changeAccrued;
    const row = RM.figureRow(cfp, 'Increase (decrease) in accounts payable and other current liabilities');
    const shown = row ? Number(row.money[0].s.replace(/[(),]/g, '').replace(/^-/, '')) : null;
    const signed = row && /^\(/.test(row.money[0].s) ? -shown : shown;
    check('CF: combined figure equals changeAP + changeAccrued',
      row && Math.abs(signed - want) < 0.005,
      'shown=' + signed + ' want=' + want.toFixed(2));
  }
  // ── 7. Equity heading ────────────────────────────────────────────────────
  check('Equity heading includes "– Tax Basis"',
    eq.some(rows => /Statement of Changes in Members’ Equity – Tax Basis/.test(RM.pageText(rows))));

  // ── Gap report: how much clear space each "$" has, and the longest label
  //    on each statement page (the thing that would collide first).
  console.log('\n  GAP REPORT ($ clearance from the label on its own row)');
  const gapRows = [];
  [['BS', bs], ['P&L', pl], ['CF', cfp], ['EQ', eq]].forEach(pair => {
    pair[1].forEach(rows => rows.forEach(r => {
      if (r.dollar && r.labelRight != null) {
        gapRows.push('    ' + pair[0] + '  gap=' + String(Math.round(r.dollar.x - r.labelRight)).padStart(4) +
          'pt   ' + r.label.slice(0, 46));
      }
    }));
  });
  gapRows.forEach(l => console.log(l));
  [['BS', bs], ['P&L', pl]].forEach(pair => {
    let worst = null;
    pair[1].forEach(rows => rows.forEach(r => {
      if (r.labelRight != null && (!worst || r.labelRight > worst.labelRight)) worst = r;
    }));
    if (worst) console.log('    ' + pair[0] + '  longest label ends at x=' +
      Math.round(worst.labelRight) + ':  ' + worst.label.slice(0, 52));
  });

  // ── Report ───────────────────────────────────────────────────────────────
  let fail = 0;
  for (const a of A) {
    if (!a.ok) fail++;
    console.log((a.ok ? '  PASS  ' : '  FAIL  ') + a.name + (a.ok ? '' : '   [' + a.detail + ']'));
  }
  console.log('\n' + (A.length - fail) + '/' + A.length + ' checks passed');
  console.log('netOperating=' + cf.netOperating + '  netInvesting=' + cf.netInvesting +
              '  netFinancing=' + cf.netFinancing + '  netChange=' + cf.netChange +
              '  cashBeg=' + cf.cashBeg + '  cashEnd=' + cf.cashEnd);
  console.log('pages=' + nPages + '  →  _gltest/_banyan_cla_aug17.pdf');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(2); });
