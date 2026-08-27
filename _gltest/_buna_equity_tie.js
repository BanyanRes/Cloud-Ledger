// _gltest/_buna_equity_tie.js — untracked harness.
// Drives buildStatements() against LIVE production balances through the `cl`
// helper and asserts that the Statement of Changes in Members' Equity
//   (a) OPENS on the prior period's closing equity, and
//   (b) CLOSES on the balance sheet's total equity.
// Regression under test: CLR Buna Property Owner's 2026 statement opened at
// 1,310,478.41 against a 12/31/2025 close of 1,162,361.34 — the 148,117.07 in
// 33011 Distribution - Ben was dropped because the account is zero in both
// balance-sheet columns.
const { execFileSync } = require('child_process');
const financials = require('../server/financials.js');

function cl(path) {
  const out = execFileSync('C:\\Program Files\\Git\\bin\\bash.exe', ['-lc', `cl GET '${path}'`],
    { maxBuffer: 64 * 1024 * 1024, encoding: 'utf8' });
  return JSON.parse(out);
}
const qs = (o) => Object.entries(o).filter(([, v]) => v != null).map(([k, v]) => `${k}=${v}`).join('&');
const makeGet = (eid) => (o) => cl(`/api/entities/${eid}/balances?${qs(o)}`);

const money = (n) => (n < 0 ? '(' : ' ') + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + (n < 0 ? ')' : ' ');

const TARGETS = [
  { eid: 38, name: 'CLR Buna Property Owner', code: 'CLRBUNAP', asOf: '2026-06-30',
    expectBeginning: 1162361.34, expectEnding: 146758.87 },
  { eid: 41, name: 'Banyan Residential', code: 'BANYANRE1', asOf: '2026-06-30' },
  { eid: 49, name: 'Banyan SFR GP Investors', code: 'BANYANSF', asOf: '2026-06-30' },
  { eid: 37, name: 'entity 37', code: '', asOf: '2026-06-30' },
  { eid: 39, name: 'entity 39', code: '', asOf: '2026-06-30' },
  { eid: 54, name: 'entity 54', code: '', asOf: '2026-06-30' },
  { eid: 70, name: 'entity 70', code: '', asOf: '2026-06-30' },
];

const r2 = (n) => Math.round(n * 100) / 100;
let pass = 0, fail = 0;

(async () => {
  for (const T of TARGETS) {
    let s;
    try {
      s = await financials.buildStatements(makeGet(T.eid), {
        asOf: T.asOf, period: 'monthly', entityName: T.name, entityCode: T.code,
      });
    } catch (e) { console.log(`\n${T.name} (${T.eid}) — SKIP: ${e.message}`); continue; }

    const t = s.equity.totals;
    const bsEquity = s.balanceSheet.totalEquity.cur;
    const ys = financials._helpers.yearStart(T.asOf);
    const open = await makeGet(T.eid)({ as_of: financials._helpers.priorMonthEnd(ys), close_pl_before: ys });
    const oA = open.filter(r => r.type === 'Asset').reduce((a, r) => a + r.balance, 0);
    const oL = open.filter(r => r.type === 'Liability').reduce((a, r) => a + r.balance, 0);
    const openEquity = r2(oA - oL);

    console.log(`\n=== ${s.meta.entityName} (entity ${T.eid}) @ ${T.asOf} ===`);
    for (const r of s.equity.rows) {
      console.log(`  ${String(r.name).slice(0, 40).padEnd(42)} beg ${money(r.beginning).padStart(16)}  contrib ${money(r.contributions).padStart(16)}  NI ${money(r.netIncome).padStart(15)}  end ${money(r.ending).padStart(16)}`);
    }
    console.log(`  ${'TOTAL'.padEnd(42)} beg ${money(t.beginning).padStart(16)}  contrib ${money(t.contributions).padStart(16)}  NI ${money(t.netIncome).padStart(15)}  end ${money(t.ending).padStart(16)}`);

    const chk = (label, a, b) => {
      const ok = Math.abs(r2(a - b)) < 0.005;
      console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${label}: ${money(a)} vs ${money(b)}${ok ? '' : '   diff ' + money(r2(a - b))}`);
      ok ? pass++ : fail++;
    };
    chk('statement beginning ties to opening A-L', t.beginning, openEquity);
    chk('statement ending ties to balance-sheet equity', t.ending, bsEquity);
    chk('balance sheet balances (A = L + E)', s.balanceSheet.totalAssets.cur, s.balanceSheet.totalLiabEquity.cur);
    console.log(`  checks.equityTies = ${s.checks.equityTies}  equityDiff = ${money(s.checks.equityDiff)}`);
    if (T.expectBeginning != null) chk('Buna beginning == 12/31/2025 close', t.beginning, T.expectBeginning);
    if (T.expectEnding != null) chk('Buna ending == 6/30/2026 A-L', t.ending, T.expectEnding);
  }
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
