// Guard for the P&L "$" latch: with NO revenue lines in the period, the latch
// must NOT leak onto the first operating-expense line. Reuses the main fixture's
// account map by requiring nothing from it — a minimal Banyan GL is enough.
const fin = require('../server/financials.js');
const RM = require('./_rowmodel.js');

const A = {
  '10143': { name: 'UBS Entity 100 Banyan - 2472', type: 'Asset' },
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '34117': { name: 'Contributed Capital - Odyssey Holdings LLC', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  '60000': { name: 'Salaries and Wages', type: 'Expense' },
  '63000': { name: 'Legal and Accounting', type: 'Expense' },
};
const sum = (m, t) => Object.keys(m).filter(c => A[c].type === t).reduce((a, c) => a + m[c], 0);
const withRE = (bs, pl) => Object.assign({}, bs, {
  '39000': +(sum(bs, 'Asset') - sum(bs, 'Liability') - sum(bs, 'Equity')
             - (pl ? (sum(pl, 'Revenue') - sum(pl, 'Expense')) : 0)).toFixed(2),
});

// NO revenue accounts at all — expenses only.
const PL = { '60000': 40000, '63000': 12000 };
const CUR = withRE({ '10143': 60000, '20000': 15000, '34117': 500000 }, PL);
const PRI = withRE({ '10143': 70000, '20000': 22000, '34117': 500000, '39000': 0 }, PL);
const OPEN = withRE({ '10143': 90000, '20000': 22000, '34117': 450000, '39000': 0 }, null);

const rows = (m) => Object.keys(m).map(c => ({ code: c, name: A[c].name, type: A[c].type, subtype: '', balance: m[c] }));
const getBalances = (o) => Promise.resolve(
  o.from ? rows(PL)
  : o.as_of === '2026-07-31' ? rows(CUR).concat(rows(PL))
  : o.as_of === '2026-06-30' ? rows(PRI).concat(rows(PL))
  : rows(OPEN));

(async () => {
  const s = await fin.buildStatements(getBalances, {
    asOf: '2026-07-31', period: 'monthly',
    entityName: 'Banyan Residential LLC', entityCode: 'BANYANRE1',
  });
  const { bytes } = await fin.generatePackage({ statements: s });
  const P = await RM.readRows(bytes);
  const pl = P.map((r, i) => ({ r, i }))
    .filter(o => o.i >= 2 && /Statements of Revenues and Expenses/.test(RM.pageText(o.r)))
    .map(o => o.r);

  let fail = 0;
  const check = (n, ok, d) => { if (!ok) fail++; console.log((ok ? '  PASS  ' : '  FAIL  ') + n + (ok ? '' : '   [' + (d || '') + ']')); };
  check('P&L page found', pl.length > 0);
  // No revenue → no row should carry a "$" except Net Income (Loss).
  const dollarRows = [];
  pl.forEach(rowsOnPage => rowsOnPage.forEach(r => { if (r.dollar) dollarRows.push(r.label); }));
  check('only Net Income (Loss) carries a "$" when there is no revenue',
    dollarRows.length === 1 && /^Net Income \(Loss\)/.test(dollarRows[0]),
    'rows with $: ' + JSON.stringify(dollarRows));
  const opex = RM.dollarOn(pl, 'Salaries and Wages', false);
  check('no "$" leaked onto the first operating-expense line', opex.ok, opex.why);
  console.log('\n' + (fail ? 'FAILED' : 'ALL PASS'));
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(2); });
