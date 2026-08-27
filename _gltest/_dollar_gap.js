// Measures, for every "$" drawn on a statement, the horizontal gap to (a) the
// PREVIOUS column's number and (b) this row's OWN number. Uses fund-scale
// figures (millions, negatives in parens) so the widest realistic numbers are
// exercised. Fails if any "$" overlaps its own number or sits closer than
// MIN_PREV_GAP to the previous column's amount.
const AFTER = require('../server/financials.js');
const RM    = require('./_rowmodel.js');

const MIN_PREV_GAP = 6;   // pt of clear space required after the previous column
const MIN_OWN_GAP  = 0;   // "$" must not overlap its own number

// Big balances so every column carries a wide (10-12 digit) figure.
const A = {
  '10143': { name: 'UBS Entity 100 Banyan - 2472', type: 'Asset' },
  '10144': { name: 'MapleMark Entity 100 Banyan - 8868', type: 'Asset' },
  '12000': { name: 'Accounts Receivable', type: 'Asset' },
  '11030': { name: 'Land', type: 'Asset' },
  '15500': { name: 'Furniture and Fixtures', type: 'Asset' },
  '16500': { name: 'Accumulated Depreciation', type: 'Asset' },
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '21112': { name: 'Accrued Liabilities', type: 'Liability' },
  '34117': { name: 'Contributed Capital - Odyssey Holdings LLC', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  '42000': { name: 'Development Fee Revenue', type: 'Revenue' },
  '60000': { name: 'Salaries and Wages', type: 'Expense' },
};
const sum = (m, t) => Object.keys(m).filter(c => A[c].type === t).reduce((a, c) => a + m[c], 0);
const withRE = (bs, pl) => Object.assign({}, bs, {
  '39000': +(sum(bs, 'Asset') - sum(bs, 'Liability') - sum(bs, 'Equity')
             - (pl ? sum(pl, 'Revenue') - sum(pl, 'Expense') : 0)).toFixed(2),
});
const PL  = { '42000': 9963058.15, '60000': 4200000.55 };
const CUR  = withRE({ '10143': 2460.56, '10144': 9145017.67, '12000': 5787114.18,
                      '11030': 8746370.65, '15500': 3308000, '16500': -1346.77,
                      '20000': 7491888.68, '21112': 6146791.08, '34117': 9746370.65 }, PL);
const PRI  = withRE(Object.assign({}, CUR, { '39000': 0, '10144': 4145017.67, '20000': 2236756.01 }), PL);
const OPEN = withRE(Object.assign({}, CUR, { '39000': 0, '10143': 1624010.53, '16500': 0,
                      '12000': 787114.18, '34117': 7000000, '20000': 3236756.01, '21112': 46791.08 }), null);

const rows = (m) => Object.keys(m).map(c => ({ code: c, name: A[c].name, type: A[c].type, subtype: '', balance: m[c] }));
const getBalances = (o) => Promise.resolve(
  o.from ? rows(PL)
  : o.as_of === '2026-07-31' ? rows(CUR).concat(rows(PL))
  : o.as_of === '2026-06-30' ? rows(PRI).concat(rows(PL))
  : rows(OPEN));

const ENTITIES = [
  { entityName: 'Banyan Residential LLC', entityCode: 'BANYANRE1', label: 'banyan' },
  { entityName: 'Sabine River & Northern Railroad', entityCode: '', label: 'srn' },
];

(async () => {
  let fail = 0, worstPrev = Infinity, worstOwn = Infinity, worstPrevWhere = '', worstOwnWhere = '';
  for (const e of ENTITIES) {
    const s = await AFTER.buildStatements(getBalances, { asOf: '2026-07-31', period: 'monthly', entityName: e.entityName, entityCode: e.entityCode });
    const { bytes } = await AFTER.generatePackage({ statements: s });
    const P = await RM.readRows(bytes);
    P.forEach((rowsOnPage, pi) => rowsOnPage.forEach(r => {
      // EVERY "$" on the baseline (these rows carry one per column), not just r.dollar.
      const dollars = r.items.filter(i => i.s === '$').sort((a, b) => a.x - b.x);
      if (!dollars.length) return;
      const nums = r.money.slice().sort((a, b) => a.x - b.x);
      if (!nums.length) return;
      for (const d of dollars) {
        const dL = d.x, dR = d.x + (d.w || 5);
        const own = nums.find(nm => nm.x >= dL - 0.5);
        if (own) {
          const g = own.x - dR;
          if (g < worstOwn) { worstOwn = g; worstOwnWhere = e.label + ' p' + (pi+1) + ' "' + r.label.slice(0,30) + '"'; }
        }
        const prevs = nums.filter(nm => (nm.x + (nm.w || 0)) <= dL + 0.5);
        if (prevs.length) {
          const prev = prevs[prevs.length - 1];
          const g = dL - (prev.x + (prev.w || 0));
          if (g < worstPrev) { worstPrev = g; worstPrevWhere = e.label + ' p' + (pi+1) + ' "' + r.label.slice(0,30) + '"'; }
        }
      }
    }));
  }
  const f = (n) => (n === Infinity ? 'n/a' : Math.round(n * 10) / 10 + 'pt');
  console.log('worst gap to PREVIOUS column amount: ' + f(worstPrev) + '   at ' + worstPrevWhere);
  console.log('worst gap to OWN number            : ' + f(worstOwn) + '   at ' + worstOwnWhere);
  const okPrev = worstPrev === Infinity || worstPrev >= MIN_PREV_GAP;
  const okOwn  = worstOwn  === Infinity || worstOwn  >= MIN_OWN_GAP;
  console.log((okPrev ? '  PASS' : '  FAIL') + '  previous-column gap >= ' + MIN_PREV_GAP + 'pt');
  console.log((okOwn  ? '  PASS' : '  FAIL') + '  no overlap with own number (>= ' + MIN_OWN_GAP + 'pt)');
  process.exit(okPrev && okOwn ? 0 : 1);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(2); });
