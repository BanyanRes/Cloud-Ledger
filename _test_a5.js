// Validate the A5 cost-code continuity check against every real requisition
// report we have: roll each one forward with a synthetic invoice on a code that
// already carries a Prior balance, then reconcile. A5 must pass on a correct
// roll (no false positives) and must fail once a Prior-log amount is nudged.
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { rollForward, findSheet } = require('./server/requisition_rollforward.js');
const { reconcile, readLog, COL, cellNum, applyInvoiceCols } = require('./server/requisition_reconcile.js');

const ROOT = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents';

function newestReport(dir) {
  let best = null, bestT = 0;
  const walk = (d, depth) => {
    if (depth > 4) return;
    let entries = [];
    try { entries = fs.readdirSync(d, { withFileTypes: true }); } catch (e) { return; }
    for (const e of entries) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) walk(p, depth + 1);
      else if (/\.xlsx$/i.test(e.name) && !/^~\$/.test(e.name) && /requisition|req\b/i.test(e.name)) {
        const t = fs.statSync(p).mtimeMs;
        if (t > bestT) { bestT = t; best = p; }
      }
    }
  };
  walk(dir, 0);
  return best;
}
const sheets = wb => ({
  prior: findSheet(wb, 'Prior Invoice Log'),
  current: findSheet(wb, 'Current Invoice Log'),
  b2a: findSheet(wb, 'Budget to Actual'),
  devFee: findSheet(wb, 'Dev Fee'),
});
const pick = c => (c ? (c.id === 'A5' ? c : null) : null);

(async () => {
  const props = fs.readdirSync(ROOT, { withFileTypes: true })
    .filter(e => e.isDirectory() && !/^(AP Invoices|General|Insurance|_|z)/.test(e.name))
    .map(e => path.join(ROOT, e.name));

  let checked = 0, falsePos = 0, negOk = 0, negMiss = 0;
  for (const p of props) {
    const reqDir = fs.readdirSync(p, { withFileTypes: true })
      .filter(e => e.isDirectory() && /requisition/i.test(e.name))
      .map(e => path.join(p, e.name))[0];
    if (!reqDir) continue;
    const file = newestReport(reqDir);
    if (!file) continue;
    const label = path.basename(p);
    try {
      // Untouched copy = the "prior period" side of the reconciliation.
      const prevWb = new ExcelJS.Workbook(); await prevWb.xlsx.readFile(file);
      const prevS = sheets(prevWb);
      if (!prevS.prior || !prevS.current) { console.log('  skip   ' + label + ' (missing logs)'); continue; }
      // Calibrate the column map exactly the way production does (rollForward
      // calls applyInvoiceCols on the workbook it rolls before reconcile runs).
      applyInvoiceCols(prevS.current);
      const oPrior = readLog(prevS.prior);
      // Choose a code that carries a balance in BOTH last month's Prior log and
      // last month's Current log, so the expected carry-forward is non-zero --
      // the 10,000 prior + 1,000 current = 11,000 case exactly.
      const oCurrL = readLog(prevS.current);
      const both = Object.keys(oCurrL.byCode)
        .filter(k => k !== '__none__' && oPrior.byCode[k] && Math.abs(oCurrL.byCode[k]) > 0.005);
      const codeKey = both[0] || Object.keys(oPrior.byCode).filter(k => k !== '__none__')[0];
      if (!codeKey) { console.log('  skip   ' + label + ' (no coded prior rows)'); continue; }
      const priorRow = oPrior.rows.find(r => String(r.code) === codeKey);
      const inv = [{ cat: 'A5 test', code: Number(codeKey), name: priorRow ? priorRow.name : 'Test', vendor: 'A5 Vendor', bill: 'A5-1', amount: 1000, date: '2026-07-05' }];

      const nextWb = new ExcelJS.Workbook(); await nextWb.xlsx.readFile(file);
      await rollForward(nextWb, inv.map(x => ({ ...x })), { reqNumber: 99, asOfDate: '2026-07-31', fixReportNumberHeader: true });
      const nextS = sheets(nextWb);
      const r = reconcile(prevS, nextS);
      const a1 = r.checks.find(c => c.id === 'A1');
      const a5 = r.checks.find(c => c.id === 'A5');
      checked++;
      const before = oPrior.byCode[codeKey];
      const after = readLog(nextS.prior).byCode[codeKey];
      const carried = (Math.round(((after || 0) - (before || 0)) * 100) / 100);
      const wanted = Math.round((oCurrL.byCode[codeKey] || 0) * 100) / 100;
      if (!a5.pass) falsePos++;
      console.log((a5.pass ? '  ok     ' : '  FALSE+ ') + label.padEnd(30)
        + ' A1=' + (a1.pass ? 'pass' : 'FAIL') + ' A5=' + (a5.pass ? 'pass' : 'FAIL')
        + ' code ' + codeKey + ' ' + (before || 0).toFixed(2) + ' -> ' + (after || 0).toFixed(2) + ' (+' + carried.toFixed(2) + ', current log had ' + wanted.toFixed(2) + ')');
      if (!a5.pass) console.log('           ' + a5.detail.slice(0, 300));

      // Negative control: move $250 of that code onto a different Prior-log row.
      const nPriorLog = readLog(nextS.prior);
      const victim = nPriorLog.rows.find(x => String(x.code) === codeKey);
      const other = nPriorLog.rows.find(x => x.code != null && String(x.code) !== codeKey);
      if (victim && other) {
        nextS.prior.getCell(victim.row, COL.amount).value = victim.amount - 250;
        nextS.prior.getCell(other.row, COL.amount).value = other.amount + 250;
        const r2 = reconcile(prevS, sheets(nextWb));
        const a5b = r2.checks.find(c => c.id === 'A5');
        const a1b = r2.checks.find(c => c.id === 'A1');
        if (!a5b.pass) { negOk++; }
        else { negMiss++; console.log('           NEGATIVE CONTROL MISSED on ' + label); }
        if (!a5b.pass && negOk === 1) console.log('           sample failure text: ' + a5b.help.slice(0, 260) + ' | A1 still ' + (a1b.pass ? 'PASSES' : 'fails'));
      }
    } catch (e) {
      console.log('  ERROR  ' + label + ' :: ' + e.message);
    }
  }
  console.log('\nreports checked ' + checked + ' | A5 false positives ' + falsePos
    + ' | mis-code caught ' + negOk + ' | mis-code missed ' + negMiss);
})();
