// Regression: the rule above "Total Assets" is dropped ONLY when the last asset
// section total already carries a rule below (Total Current Assets / Total Fixed
// Assets, Net). A balance sheet whose last section is Other Assets keeps it.
const fs = require('fs');
const fitz = null;
const financials = require('./server/financials.js');

const A = (code, name, type, balance, subtype = '', bank_acct = 0) =>
  ({ code, name, type, subtype, bank_acct, balance, total_debit: 0, total_credit: 0 });

const base = (k) => [
  A('10001', 'MapleMark Operating - 4817', 'Asset', 900000 * k, '', 1),
  A('11000', 'Accounts Receivable', 'Asset', 200000 * k),
  A('20000', 'Accounts Payable', 'Liability', 150000 * k),
  A('34000', 'Members Contributions', 'Equity', 800000),
  A('40000', 'Rail Operations Revenue', 'Revenue', 300000 * k),
  A('63041', 'Management Fees', 'Expense', 120000 * k),
];

const variants = {
  current_only: (k) => base(k),
  with_other_assets: (k) => base(k).concat([
    A('12002', 'Construction in Progress', 'Asset', 500000 * k),
    A('15000', 'Land', 'Asset', 250000 * k),
  ]),
};

(async () => {
  for (const [name, rows] of Object.entries(variants)) {
    const getBalances = (o) => Promise.resolve(rows(o && o.as_of && o.as_of < '2026-07-01' ? 0.94 : 1));
    const s = await financials.buildStatements(getBalances, {
      asOf: '2026-07-31', period: 'monthly', entityName: 'Test Entity', entityCode: 'TEST',
    });
    const bytes = await financials.renderStatementsPdf(s);
    fs.writeFileSync('_bs_' + name + '.pdf', Buffer.from(bytes));
    console.log(name, '| sections:', s.balanceSheet.assetSections.map(x => x.title).join(' / '));
  }
})().catch(e => { console.error(e); process.exit(1); });
