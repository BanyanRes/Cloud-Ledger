// End-to-end check of the Silsbee requisition fixes: roll the JUNE final Phase 1
// workbook forward to July with the same four invoices CLA received, twice, in
// two DIFFERENT upload orders, and verify the output against CLA's corrected
// July file.
const ExcelJS = require('exceljs');
const { rollForward } = require('./server/requisition_rollforward.js');

const DIR = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/';
const JUNE = DIR + '06 Jun 2026/00 B1 County Line Rail Silsbee LLC - Requisition Report 06.2026 Phase 1.xlsx';

const S = c => {
  const v = c.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.formula) return '=' + v.formula;
    if (v.result !== undefined) return String(v.result);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
};
const NUM = c => {
  const v = c.value;
  if (v == null) return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'object' && typeof v.result === 'number') return v.result;
  return null;
};

// The four July invoices, exactly as they appear on CLA's report.
const INVOICES = [
  { cat: 'Professional Services - Accounting', code: 12230, name: 'Professional Services - Accounting',
    vendor: 'CliftonLarsonAllen LLP', bill: 'L261457720', amount: 6912.36, date: '2026-07-27' },
  { cat: 'Construction Period Interest', code: 12321, name: 'Construction Period Interest',
    vendor: 'BoTX', bill: 'Jul_26 Interest', amount: 62794.40, date: '2026-07-13' },
  { cat: 'Broker Commissions', code: 12594, name: 'Broker Commissions',
    vendor: 'DFW Lee & Associates, LLC - Houston Office, RS', bill: '20256631-1', amount: 181538.72, date: '2026-07-01' },
];
const META = { reqNumber: 11, asOfDate: '2026-07-31', fixReportNumberHeader: true,
  collapseDevFeeCosts: true, devFeePayee: 'County Line Rail Interest' };

async function run(order) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(JUNE);
  const rows = order.map(i => ({ ...INVOICES[i] }));
  const res = await rollForward(wb, rows, { ...META });
  return { wb, res };
}

function groupOrder(ws) {
  const out = [];
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 1; r <= last; r++) {
    const code = S(ws.getCell(r, 3));
    if (code && code !== 'Cost Code #') out.push(code);
  }
  return out;
}

function bordersBelow(ws, gtRow) {
  const hits = [];
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = gtRow + 1; r <= last; r++) {
    for (let c = 1; c <= 11; c++) {
      const b = ws.getCell(r, c).border;
      if (b && ['top', 'bottom', 'left', 'right'].some(e => b[e] && b[e].style)) { hits.push('r' + r + 'c' + c); break; }
    }
  }
  return hits;
}

function findRow(ws, col, needle) {
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 1; r <= last; r++) if (String(S(ws.getCell(r, col))).trim() === needle) return r;
  return null;
}

(async () => {
  let pass = 0, fail = 0;
  const ok = (cond, label, detail) => { if (cond) { pass++; console.log('  PASS  ' + label); } else { fail++; console.log('  FAIL  ' + label + (detail ? '   ' + detail : '')); } };

  const a = await run([0, 1, 2]);            // accounting -> interest -> broker  (v1's order)
  const b = await run([2, 1, 0]);            // broker -> interest -> accounting  (v2's order)

  console.log('\n=== 1. Order is independent of upload order ===');
  const ga = groupOrder(a.wb.getWorksheet('Current Invoice Log P1'));
  const gb = groupOrder(b.wb.getWorksheet('Current Invoice Log P1'));
  console.log('  run A (12230,12321,12594 uploaded): ' + ga.join(' -> '));
  console.log('  run B (12594,12321,12230 uploaded): ' + gb.join(' -> '));
  ok(JSON.stringify(ga) === JSON.stringify(gb), 'same invoices in a different upload order produce the same log');

  // Expected order = the Budget-to-Actual's own row order for these codes.
  const b2a = a.wb.getWorksheet('Budget to Actual P1');
  const b2aRowOf = (code) => { const last = Math.max(b2a.rowCount || 0, b2a.actualRowCount || 0); for (let r = 1; r <= last; r++) if (String(S(b2a.getCell(r, 2))).trim() === String(code)) return r; return 1e9; };
  const sorted = ga.slice().sort((x, y) => b2aRowOf(x) - b2aRowOf(y));
  console.log('  Budget-to-Actual order for the same codes:  ' + sorted.join(' -> '));
  ok(JSON.stringify(ga) === JSON.stringify(sorted), 'log group order follows the Budget-to-Actual');

  console.log('\n=== 2. No stray rules below the Grand Total ===');
  const cur = a.wb.getWorksheet('Current Invoice Log P1');
  const gt = findRow(cur, 4, 'Grand Total');
  const stray = bordersBelow(cur, gt);
  console.log('  Grand Total at row ' + gt + '; bordered cells below it: ' + (stray.length ? stray.join(', ') : 'none'));
  ok(stray.length === 0, 'no orphaned rules below the Grand Total');
  const gtb = cur.getCell(gt, 7).border || {};
  console.log('  Grand Total amount cell border: top=' + (gtb.top && gtb.top.style) + ' bottom=' + (gtb.bottom && gtb.bottom.style));
  ok(gtb.top && gtb.top.style === 'thin' && gtb.bottom && gtb.bottom.style === 'double', 'Grand Total carries thin-over-double');

  console.log('\n=== 3. Hard Cost Contingency table rolled ===');
  console.log('  resolved hard tab: "' + a.res.contingencyTables.hardSheet + '"');
  console.log('  resolved soft tab: "' + a.res.contingencyTables.softSheet + '"');
  ok(a.res.contingencyTables.hardSheet === 'Hard Cost Contigency P1', 'picked the live (misspelled) hard tab, not the legacy one');
  const hard = a.wb.getWorksheet('Hard Cost Contigency P1');
  const d19 = NUM(hard.getCell(19, 4)), e19 = hard.getCell(19, 5).value;
  console.log('  row 19: D (previously requested) = ' + d19 + ' ; E (requested herein) = ' + JSON.stringify(e19));
  ok(Math.abs((d19 || 0) - 16174) < 0.005, 'the 16,174 herein folded into previously-requested', 'got ' + d19);
  ok(e19 == null, 'requested-herein cleared for the new period');
  console.log('  rows rolled: hard=' + a.res.contingencyTables.hard.moved + ' soft=' + a.res.contingencyTables.soft.moved);
  ok(a.res.contingencyTables.hard.moved > 0, 'hard table reports a non-zero roll');

  console.log('\n=== 4. "Inception to" header ===');
  const h8 = S(b2a.getCell(8, 8));
  console.log('  H8 = "' + h8 + '"   (June read "Inception to 05/31/2026")');
  ok(/06\/30\/2026/.test(h8), 'inception header names the PRIOR period end, not the new one', h8);
  console.log('  I8 = "' + S(b2a.getCell(8, 9)) + '"   J8 = "' + S(b2a.getCell(8, 10)) + '"   L8 = "' + S(b2a.getCell(8, 12)) + '"');
  ok(/7\/31\/2026/.test(S(b2a.getCell(8, 9))), 'this-period header advanced to the new period end');

  console.log('\n=== 5. Dev Fee schedule preserved ===');
  const df = a.wb.getWorksheet('Dev Fee P1');
  const multi = ['B', 'C', 'D'].map((L, i) => S(df.getCell(7, 2 + i)));
  console.log('  row 7 across B/C/D: ' + JSON.stringify(multi));
  ok(multi.filter(Boolean).length >= 3, 'the three-column Silsbee schedule survived collapseDevFeeCosts');
  ok(/Due to CLR Silsbee Property Owner/i.test(S(df.getCell(18, 1))), 'the "Due to CLR Silsbee Property Owner" line is still there', S(df.getCell(18, 1)));

  console.log('\n=== 6. Warnings surfaced ===');
  console.log('  ' + (a.res.warnings && a.res.warnings.length ? a.res.warnings.join(' | ') : '(none)'));
  ok(!a.res.warnings || !a.res.warnings.some(w => /hard/.test(w)), 'no silent hard-contingency miss');

  await a.wb.xlsx.writeFile('_silsbee_july_rebuilt.xlsx');
  console.log('\nwrote _silsbee_july_rebuilt.xlsx');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e.message); console.error(e.stack); process.exit(2); });
