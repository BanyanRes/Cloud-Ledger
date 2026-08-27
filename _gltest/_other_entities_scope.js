// Scope check after Jimmy's 8/17 "everything except tax basis is global" call.
// For every NON-Banyan-Residential profile, the new behavior must be:
//   * dollar signs now PRESENT (first line + totals) -- was absent before
//   * the combined "accounts payable and other current liabilities" line PRESENT,
//     the separate "accrued and other liabilities" line GONE
//   * NO "- Tax Basis" anywhere, NO Banyan-only statement titles
//   * every reconciliation subtotal numerically IDENTICAL to the pre-change
//     module (the combine is presentation-only; the bucketing was NOT touched)
//
// BEFORE = server/_financials_before.js, extracted from 1e443e5.
const BEFORE = require('../server/_financials_before.js');
const AFTER  = require('../server/financials.js');
const RM     = require('./_rowmodel.js');

const ACCTS = {
  '10162': { name: 'Operating x3505', type: 'Asset' },
  '12000': { name: 'Accounts Receivable', type: 'Asset' },
  '13001': { name: 'Prepaid Insurance', type: 'Asset' },
  '15100': { name: 'Land', type: 'Asset' },
  '16600': { name: 'Accumulated Amortization', type: 'Asset' },
  '12600': { name: 'Capitalized Dev Costs', type: 'Asset' },
  '18311': { name: 'Due from County Line RR', type: 'Asset' },
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '21006': { name: 'Accrued Expenses', type: 'Liability' },
  '25063': { name: 'Note Payable - BOT', type: 'Liability' },
  '34006': { name: 'Contributed Capital - Class A', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  '40000': { name: 'Rail Revenue', type: 'Revenue' },
  '60000': { name: 'Salaries & Wages', type: 'Expense' },
  '63000': { name: 'Accounting', type: 'Expense' },
  '69100': { name: 'Depreciation Expense', type: 'Expense' },
};
const sum = (m, t) => Object.keys(m).filter(c => ACCTS[c].type === t).reduce((a, c) => a + m[c], 0);
const withRE = (bs, pl) => Object.assign({}, bs, {
  '39000': +(sum(bs, 'Asset') - sum(bs, 'Liability') - sum(bs, 'Equity')
             - (pl ? sum(pl, 'Revenue') - sum(pl, 'Expense') : 0)).toFixed(2),
});
// AP AND accrued both move, so the combined line renders and the split-accrued
// line would render if it (wrongly) survived. Note payable is unchanged, so the
// bucketing question does not even arise here (measured separately).
const PL  = { '40000': 220000, '60000': 140000, '63000': 30000, '69100': 1500 };
const CUR  = withRE({ '10162': 300000, '12000': 50000, '13001': 10000, '15100': 500000,
                      '16600': -1500, '12600': 90000, '18311': 40000,
                      '20000': 80000, '21006': 20000, '25063': 400000, '34006': 300000 }, PL);
const PRI  = withRE(Object.assign({}, CUR, { '39000': 0, '10162': 310000, '20000': 90000 }), PL);
const OPEN = withRE(Object.assign({}, CUR, { '39000': 0, '10162': 250000, '16600': 0,
                      '12000': 120000, '34006': 250000, '20000': 55000, '21006': 42000 }), null);

const rows = (m) => Object.keys(m).map(c => ({ code: c, name: ACCTS[c].name, type: ACCTS[c].type, subtype: '', balance: m[c] }));
const getBalances = (o) => Promise.resolve(
  o.from ? rows(PL)
  : o.as_of === '2026-07-31' ? rows(CUR).concat(rows(PL))
  : o.as_of === '2026-06-30' ? rows(PRI).concat(rows(PL))
  : rows(OPEN));

const ENTITIES = [
  { entityName: 'Sabine River & Northern Railroad', entityCode: '',         label: 'srn (default)' },
  { entityName: 'Banyan SFR GP Investors',          entityCode: 'BANYANSF', label: 'bsfrgp' },
  { entityName: 'County Line Industrial Park LLC',  entityCode: '',         label: 'srn (CLIP)' },
];

const build = async (mod, opts) => mod.buildStatements(getBalances,
  Object.assign({ asOf: '2026-07-31', period: 'monthly' }, opts));

const render = async (mod, s) => {
  const { bytes } = await mod.generatePackage({ statements: s });
  return bytes;
};

// Every reconciliation subtotal that must not move.
const recon = (s) => ({
  netOperating: s.cashFlow.netOperating, netInvesting: s.cashFlow.netInvesting,
  netFinancing: s.cashFlow.netFinancing, netChange: s.cashFlow.netChange,
  cashEnd: s.cashFlow.cashEnd,
  totalAssets: s.balanceSheet.totalAssets.cur,
  totalLiabEquity: s.balanceSheet.totalLiabEquity.cur,
  netIncome: s.operations.netIncome ? s.operations.netIncome.ytd : null,
});

(async () => {
  let fail = 0;
  const check = (n, ok, d) => { if (!ok) fail++; console.log((ok ? '  PASS  ' : '  FAIL  ') + n + (ok ? '' : '   [' + (d || '') + ']')); };

  for (const e of ENTITIES) {
    console.log('\n== ' + e.label + ' ==');
    const sB = await build(BEFORE, e);
    const sA = await build(AFTER, e);

    // 1. Reconciliation numerically identical -> presentation-only, bucketing untouched.
    const rB = recon(sB), rA = recon(sA);
    const keys = Object.keys(rB);
    const drift = keys.filter(k => Math.abs((rB[k] || 0) - (rA[k] || 0)) > 0.005);
    check('reconciliation subtotals unchanged vs 1e443e5', drift.length === 0,
      drift.map(k => k + ': ' + rB[k] + ' -> ' + rA[k]).join(', '));

    const bytesB = await render(BEFORE, sB);
    const bytesA = await render(AFTER, sA);
    const PB = await RM.readRows(bytesB);
    const PA = await RM.readRows(bytesA);
    const textA = PA.map(RM.pageText).join('\n');

    // 2. Dollar signs now present (were absent in this render path before).
    const dB = PB.reduce((a, p) => a + p.filter(r => r.dollar).length, 0);
    const dA = PA.reduce((a, p) => a + p.filter(r => r.dollar).length, 0);
    check('dollar signs now present (was ' + dB + ', now ' + dA + ')', dB === 0 && dA > 0);

    // Specifically: BS first line + Total Assets, CF Net Income + Cash End, and
    // the P&L Net Income. "Net Income (Loss)" appears on the balance sheet too
    // (no $, correctly), so scope each check to its own statement's page(s).
    const pagesTitled = (re) => PA.map((r, i) => ({ r, i }))
      .filter(o => o.i >= 2 && re.test(RM.pageText(o.r))).map(o => o.r);
    const plPages = pagesTitled(/Statements of Operations/);
    const cfPages = pagesTitled(/Statement of Cash Flows/);
    check('  BS $ on Total Assets', RM.dollarOn(PA, 'Total Assets', true).ok);
    check('  CF $ on Cash, End of Period', RM.dollarOn(cfPages, 'Cash, End of Period', true).ok);
    check('  CF $ on Net Income (Loss)', RM.dollarOn(cfPages, 'Net Income (Loss)', true).ok);
    check('  P&L $ on Net Income (Loss)', RM.dollarOn(plPages, 'Net Income (Loss)', true).ok);
    check('  BS Net Income (Loss) has NO $', !RM.dollarOn(PA, 'Net Income (Loss)', true).ok);

    // 3. Combined liability line present; split accrued line gone.
    check('combined "accounts payable and other current liabilities" line present',
      /Increase \(decrease\) in accounts payable and other current liabilities/.test(textA));
    check('separate "accrued and other liabilities" line gone',
      !/Increase \(decrease\) in accrued and other liabilities/.test(textA));

    // 4. NO tax-basis anywhere, NO Banyan-only titles.
    check('no "- Tax Basis" anywhere', !/– Tax Basis/.test(textA));
    check('no Banyan balance-sheet title', !/Statements of Assets, Liabilities/.test(textA));
    check('no Banyan P&L title', !/Statements of Revenues and Expenses/.test(textA));
    check('keeps its own title (Balance Sheets / Statements of Operations)',
      /Balance Sheets/.test(textA) && /Statements of Operations/.test(textA));
  }

  console.log('\n' + (fail ? fail + ' FAILURE(S)' : 'ALL PASS — non-Banyan entities get $ + combined line, no tax basis, no subtotal drift'));
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(2); });
