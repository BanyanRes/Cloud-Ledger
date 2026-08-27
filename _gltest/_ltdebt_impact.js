// Quantifies the OTHER global change: making the cash-flow liability bucket test
// section-aware (long-term before name match). Banyan has no long-term payable,
// so it is a no-op there — but SRN-style entities do. This measures the swing.
const BEFORE = require('../server/_financials_before.js');
const AFTER  = require('../server/financials.js');

const ACCTS = {
  '10162': { name: 'Operating x3505', type: 'Asset' },
  '15100': { name: 'Land', type: 'Asset' },
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '25063': { name: 'Note Payable - BOT', type: 'Liability' },   // LONG TERM, and it moves
  '34006': { name: 'Contributed Capital - Class A', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  '40000': { name: 'Rail Revenue', type: 'Revenue' },
  '60000': { name: 'Salaries & Wages', type: 'Expense' },
};
const sum = (m, t) => Object.keys(m).filter(c => ACCTS[c].type === t).reduce((a, c) => a + m[c], 0);
const withRE = (bs, pl) => Object.assign({}, bs, {
  '39000': +(sum(bs, 'Asset') - sum(bs, 'Liability') - sum(bs, 'Equity')
             - (pl ? sum(pl, 'Revenue') - sum(pl, 'Expense') : 0)).toFixed(2),
});
const PL = { '40000': 200000, '60000': 150000 };
// Note payable DRAWN during the year: 400,000 -> 650,000.
const CUR  = withRE({ '10162': 500000, '15100': 500000, '20000': 80000, '25063': 650000, '34006': 300000 }, PL);
const PRI  = withRE(Object.assign({}, CUR, { '39000': 0, '25063': 650000 }), PL);
const OPEN = withRE(Object.assign({}, CUR, { '39000': 0, '10162': 250000, '25063': 400000 }), null);

const rows = (m) => Object.keys(m).map(c => ({ code: c, name: ACCTS[c].name, type: ACCTS[c].type, subtype: '', balance: m[c] }));
const getBalances = (o) => Promise.resolve(
  o.from ? rows(PL)
  : o.as_of === '2026-07-31' ? rows(CUR).concat(rows(PL))
  : o.as_of === '2026-06-30' ? rows(PRI).concat(rows(PL))
  : rows(OPEN));

(async () => {
  const opts = { asOf: '2026-07-31', period: 'monthly', entityName: 'Sabine River & Northern Railroad' };
  const b = await BEFORE.buildStatements(getBalances, opts);
  const a = await AFTER.buildStatements(getBalances, opts);
  const f = (n) => (n < 0 ? '(' + Math.abs(n).toLocaleString() + ')' : n.toLocaleString());
  console.log('Note Payable - BOT drawn 400,000 -> 650,000 during the year.\n');
  console.log('                                        BEFORE          AFTER');
  console.log('  changeAP (operating AP line)   ' + f(b.cashFlow.changeAP).padStart(14) + f(a.cashFlow.changeAP).padStart(15));
  console.log('  debtChange (financing)         ' + f(b.cashFlow.debtChange).padStart(14) + f(a.cashFlow.debtChange).padStart(15));
  console.log('  Net Cash from OPERATING        ' + f(b.cashFlow.netOperating).padStart(14) + f(a.cashFlow.netOperating).padStart(15));
  console.log('  Net Cash from FINANCING        ' + f(b.cashFlow.netFinancing).padStart(14) + f(a.cashFlow.netFinancing).padStart(15));
  console.log('  Net change in cash             ' + f(b.cashFlow.netChange).padStart(14) + f(a.cashFlow.netChange).padStart(15));
  const swing = a.cashFlow.netOperating - b.cashFlow.netOperating;
  console.log('\n  => Operating subtotal moves by ' + f(swing) + ', offset in financing.');
  console.log('  => Net change in cash unaffected: ' + (Math.abs(a.cashFlow.netChange - b.cashFlow.netChange) < 0.005));
})().catch(e => { console.error(e); process.exit(2); });
