// Standalone fixture test for the P&L operating-expense grouping.
// Verifies: (1) category subtotals sum to totOpex; (2) fixed presentation order;
// (3) no expense line dropped; (4) heuristic fallback for unmapped codes;
// (5) full PDF renders without error.
const { buildStatements, renderStatementsPdf } = require('../server/financials.js');

// Minimal chart: a few mapped expense codes across categories, one COGS, one
// unmapped-but-name-classifiable code (99999 "Legal Fees" -> Professional
// Services via heuristic), one revenue, plus BS accounts so the statement builds.
const EXP = [
  ['63000', 'Accounting', 1000],
  ['63025', 'Professional Fees', 500],
  ['99999', 'Legal Fees', 250],            // unmapped -> heuristic Professional Services
  ['67300', 'Telephone & Internet', 300],  // Technology & Software
  ['60000', 'Salaries & Wages', 8000],     // Personnel / Payroll
  ['60002', 'Payroll Taxes', 700],         // Personnel / Payroll
  ['61150', 'Utilities', 400],             // Fuel & Utilities
  ['68050', 'Property Tax', 1200],         // Taxes & Assessments
  ['65000', 'Insurance - Liability', 900], // Insurance
  ['63041', 'CLRO Management Fees', 5000], // Management Fees (last)
  ['77777', 'Random Unknown Expense', 111],// unmapped, no heuristic -> Administrative & Other
];
const COGS = [['50000', 'Car Hire', 2000]];
const REV = [['41000', 'Freight Revenue', 30000]];
const BS = [
  ['10161', 'Operating Cash', 'Asset', 50000],
  ['20000', 'Accounts Payable', 'Liability', 5000],
  ['34006', 'Member Capital', 'Equity', 45000],
];

function rowsFor({ from, to, as_of }) {
  // For any window we just return the same magnitudes (fixture); the grouping
  // logic is window-agnostic. BS accounts always present; P&L present for period
  // and YTD queries.
  const out = [];
  for (const [code, name, typ, balc] of BS) out.push({ code, name, type: typ, subtype: '', balance: balc });
  // include P&L on both period and ytd queries (from/to present)
  if (from || to) {
    for (const [code, name, amt] of REV) out.push({ code, name, type: 'Revenue', subtype: '', balance: amt });
    for (const [code, name, amt] of COGS) out.push({ code, name, type: 'Expense', subtype: 'COGS', balance: amt });
    for (const [code, name, amt] of EXP) out.push({ code, name, type: 'Expense', subtype: '', balance: amt });
  }
  return out;
}
const getBalances = async (q) => rowsFor(q);

(async () => {
  const s = await buildStatements(getBalances, { asOf: '2026-04-30', entityName: 'Test SRN', period: 'monthly' });
  const ops = s.operations;
  const groups = ops.opexGroups;
  let fail = 0;
  const check = (cond, msg) => { if (!cond) { console.log('FAIL:', msg); fail++; } else console.log('ok:', msg); };

  // 1. Every expense line accounted for exactly once across groups.
  const inGroups = groups.reduce((n, g) => n + g.lines.length, 0);
  check(inGroups === ops.opex.length, `all opex lines grouped (${inGroups} == ${ops.opex.length})`);

  // 2. Sum of category subtotals == totOpex on all three columns.
  const sum = k => Math.round(groups.reduce((a, g) => a + g.subtotal[k], 0) * 100) / 100;
  check(sum('cur') === ops.totOpex.cur, `subtotals sum cur (${sum('cur')} == ${ops.totOpex.cur})`);
  check(sum('ytd') === ops.totOpex.ytd, `subtotals sum ytd (${sum('ytd')} == ${ops.totOpex.ytd})`);

  // 3. Fixed presentation order (subsequence of the canonical order), Management Fees last.
  const ORDER = ['Professional Services','Technology & Software','Administrative & Other','Personnel / Payroll','Track & Infrastructure','Equipment & Rolling Stock','Fuel & Utilities','Contracted Services','Insurance','Taxes & Assessments','Regulatory & Compliance','Management Fees'];
  const titles = groups.map(g => g.title);
  let pos = -1, ordered = true;
  for (const t of titles) { const i = ORDER.indexOf(t); if (i < pos) ordered = false; pos = i; }
  check(ordered, 'categories in canonical order: ' + titles.join(' | '));
  check(titles[titles.length - 1] === 'Management Fees', 'Management Fees rendered last');

  // 4. Heuristic classifications landed correctly.
  const catOf = code => { for (const g of groups) if (g.lines.some(l => l.code === code)) return g.title; return null; };
  check(catOf('99999') === 'Professional Services', '99999 Legal Fees -> Professional Services (heuristic)');
  check(catOf('77777') === 'Administrative & Other', '77777 unknown -> Administrative & Other (catch-all)');
  check(catOf('63041') === 'Management Fees', '63041 -> Management Fees');

  // 5. Net income unaffected by grouping (grand total identity).
  check(ops.netIncome.ytd === Math.round((ops.grossProfit.ytd - ops.totOpex.ytd) * 100) / 100, 'net income = gross profit - totOpex');

  // 6. Full render produces a non-trivial PDF.
  const bytes = await renderStatementsPdf(s);
  check(bytes && bytes.length > 1000, `PDF rendered (${bytes ? bytes.length : 0} bytes)`);

  console.log(fail === 0 ? '\nALL PASS' : `\n${fail} FAILURE(S)`);
  process.exit(fail === 0 ? 0 : 1);
})().catch(e => { console.error('THREW:', e); process.exit(2); });
