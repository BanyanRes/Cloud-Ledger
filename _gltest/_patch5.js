// Patch 5: per-statement render blocks (BS/Operations/CashFlow/Members Equity).
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + JSON.stringify(o.slice(0, 220)));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, () => n); applied++; console.log('ok:', label);
}

// ── 5.1 Balance Sheets: repeat the "<cur> and <prior>" date-line on every page
//        (via makeLayout dateLine), no underline on the date headers, keep space.
replace('BS block heading',
`    const L = makeLayout(pdf, fonts, m, 'Balance Sheets');
    L.start();
    // Centered "<current date> and <prior date>" with only the dates underlined,
    // then a blank space before the columns/first line.
    L.drawCenteredDates(m.longDate, m.priorLongDate);
    L.space(4);
    L.setCols(twoCols);
    L.colHeaders([m.longDate, m.priorLongDate]);
    L.sectionTitle('ASSETS');`,
`    // Heading date-line repeats on every page (incl. continuation pages) and a
    // blank space follows it before the first row. Dates are NOT underlined.
    const L = makeLayout(pdf, fonts, m, 'Balance Sheets', { dateLine: m.longDate + ' and ' + m.priorLongDate });
    L.start();
    L.setCols(twoCols);
    L.colHeaders([m.longDate, m.priorLongDate]);
    L.sectionTitle('ASSETS');`);

// ── 5.2 Statements of Operations: heading "For the Months Ended <cur> and
//        <prior>", column headers are the actual dates, spacing via makeLayout.
replace('Operations block heading',
`    const L = makeLayout(pdf, fonts, m, 'Statements of Operations');
    L.setSubline(m.periodLabel + '   (with year-to-date)');
    L.start();
    L.setCols(threeCols);
    const curHdr = (m.period === 'monthly' ? 'Current Month' : (m.period === 'quarterly' ? 'Current Quarter' : 'Current Year'));
    const priHdr = (m.period === 'monthly' ? 'Prior Month' : (m.period === 'quarterly' ? 'Prior Quarter' : 'Prior Year'));
    L.colHeaders([curHdr, priHdr, 'Year to Date']);`,
`    const L = makeLayout(pdf, fonts, m, 'Statements of Operations', { dateLine: 'For the Months Ended ' + m.longDate + ' and ' + m.priorLongDate });
    L.start();
    L.setCols(threeCols);
    // Current-month column shows the current period-end date; prior-month column
    // shows the prior period-end date (per round-2 feedback).
    L.colHeaders([m.longDate, m.priorLongDate, 'Year to Date']);`);

// ── 5.3 Statement of Cash Flows: drop the "Year to Date" column heading.
replace('Cash Flow drop YTD header',
`    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows');
    L.setSubline(m.monthsEnded);
    L.start();
    L.setCols([RIGHT]);
    L.colHeaders(['Year to Date']);
    const cf = s.cashFlow;`,
`    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows', { dateLine: m.monthsEnded });
    L.start();
    L.setCols([RIGHT]);
    // No "Year to Date" column heading (single YTD column, per round-2 feedback);
    // add a little space where the header row would have been.
    L.space(6);
    const cf = s.cashFlow;`);

// ── 5.4 Statement of Changes in Members' Equity: landscape, 5 columns mirroring
//        the reference (Equity Balances at 01/01, Contributions, Distributions,
//        Net Income (Loss), Equity Balances at <asof>), each $-prefixed, wider
//        Net Income column so it stays on one row, spacing before first row.
replace('Members Equity block',
`    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\\u2019 Equity');
    L.setSubline(m.monthsEnded);
    L.start();
    const eCols = [RIGHT - 285, RIGHT - 190, RIGHT - 95, RIGHT];
    L.setCols(eCols);
    L.colHeaders(['Beginning', 'Contributions', 'Net Income\\n(Loss)', 'Ending']);
    for (const r of s.equity.rows) {
      L.row(r.name, [acct(r.beginning), acct(r.contributions), acct(r.netIncome), acct(r.ending)], { indent: 10 });
    }
    const t = s.equity.totals;
    L.row('Total Members\\u2019 Equity', [acct(t.beginning), acct(t.contributions), acct(t.netIncome), acct(t.ending)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }`,
`    // Landscape page mirroring the CPA reference: five columns, each money value
    // prefixed with "$", a Distributions column shown even when all zero, and a
    // Net Income (Loss) column wide enough to keep the value on one row.
    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\\u2019 Equity',
      { landscape: true, dateLine: m.monthsEnded });
    const LRIGHT = PAGE.h - PAGE.mR; // landscape printable right edge (PAGE.h is the long side)
    L.start();
    // Column right-edges across the landscape width. Two-line headers, dates
    // shown as m/d/yyyy short form to match the reference.
    const shortMD = (long) => {
      // "April 30, 2026" -> "4/30/2026"
      const map = { January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12 };
      const mm = long.match(/^(\\w+)\\s+(\\d+),\\s+(\\d+)$/);
      if (!mm) return long;
      return map[mm[1]] + '/' + mm[2] + '/' + mm[3];
    };
    const begDate = '1/1/' + String(m.asOf).slice(0, 4);
    const endDate = shortMD(m.longDate);
    const c1 = LRIGHT - 620, c2 = LRIGHT - 465, c3 = LRIGHT - 310, c4 = LRIGHT - 155, c5 = LRIGHT;
    const eCols = [c1, c2, c3, c4, c5];
    L.setCols(eCols);
    L.colHeaders([
      'Equity\\nBalances at\\n' + begDate,
      'Contributions',
      'Distributions',
      'Net Income\\n(Loss)',
      'Equity\\nBalances at\\n' + endDate,
    ]);
    // Money cell with a "$" prefix at the column's left and the value right-aligned.
    const dollarRow = (label, vals, o = {}) => {
      L.row(label, vals.map(v => acct(v)), Object.assign({ indent: 10, dollarPrefix: true }, o));
    };
    L.row('Member', [], { indent: 6, boldRow: true });
    for (const r of s.equity.rows) {
      dollarRow(r.name, [r.beginning, r.contributions, r.distributions, r.netIncome, r.ending], { indent: 16 });
    }
    const t = s.equity.totals;
    dollarRow('Total', [t.beginning, t.contributions, t.distributions, t.netIncome, t.ending],
      { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nPATCH5 APPLIED', applied, 'edits.');
