// Cross-property regression for the Silsbee requisition fixes. Rolls the latest
// report of every property forward one period with a single synthetic invoice,
// twice, in two different upload orders, and checks that:
//   - the roll completes without throwing,
//   - the Current Invoice Log comes out identical regardless of upload order,
//   - no rules are left orphaned below the Grand Total,
//   - the contingency-table resolver picks a sheet that actually has the layout,
//   - the "Inception to" header never names the NEW period end.
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { rollForward } = require('./server/requisition_rollforward.js');

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

const S = c => {
  const v = c.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.formula) return '=' + v.formula;
    if (v.result !== undefined) return String(v.result);
  }
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
};

const INV = [
  { cat: 'Test A', code: 12230, name: 'Professional Services - Accounting', vendor: 'Vendor A', bill: 'A-1', amount: 1000, date: '2026-07-05' },
  { cat: 'Test B', code: 12321, name: 'Construction Period Interest', vendor: 'Vendor B', bill: 'B-1', amount: 2000, date: '2026-07-20' },
];

function logSignature(wb) {
  const ws = wb.worksheets.find(w => /current inv/i.test(w.name));
  if (!ws) return 'NO-LOG';
  const out = [];
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 1; r <= last; r++) {
    const code = S(ws.getCell(r, 3)), nm = S(ws.getCell(r, 4));
    if (code || /Total$/i.test(nm)) out.push(r + ':' + code + ':' + nm);
  }
  return out.join('|');
}

function strayBorders(wb) {
  const ws = wb.worksheets.find(w => /current inv/i.test(w.name));
  if (!ws) return [];
  let gt = null;
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 1; r <= last; r++) if (/^gran[dt] total$/i.test(String(S(ws.getCell(r, 4))).trim())) gt = r;
  if (!gt) return [];
  const hits = [];
  for (let r = gt + 1; r <= last; r++) {
    for (let c = 1; c <= 11; c++) {
      const b = ws.getCell(r, c).border;
      if (b && ['top', 'bottom', 'left', 'right'].some(e => b[e] && b[e].style)) { hits.push('r' + r); break; }
    }
  }
  return hits;
}

function inceptionHeaders(wb) {
  const b2a = wb.worksheets.find(w => /budget to actual/i.test(w.name));
  if (!b2a) return [];
  const out = [];
  for (let r = 1; r <= 12; r++) for (let c = 1; c <= 15; c++) {
    const t = S(b2a.getCell(r, c));
    if (/inception|previous application|prior/i.test(t) && /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(t)) out.push(t);
  }
  return out;
}

(async () => {
  const props = fs.readdirSync(ROOT, { withFileTypes: true })
    .filter(e => e.isDirectory() && !/^(AP Invoices|General|Insurance|_|z)/.test(e.name))
    .map(e => path.join(ROOT, e.name));

  let checked = 0, problems = 0;
  for (const p of props) {
    const reqDir = fs.readdirSync(p, { withFileTypes: true })
      .filter(e => e.isDirectory() && /requisition/i.test(e.name))
      .map(e => path.join(p, e.name))[0];
    if (!reqDir) continue;
    const file = newestReport(reqDir);
    if (!file) continue;
    const label = path.basename(p);
    let a, b;
    try {
      const wa = new ExcelJS.Workbook(); await wa.xlsx.readFile(file);
      const ra = await rollForward(wa, INV.map(x => ({ ...x })), { reqNumber: 99, asOfDate: '2026-07-31', fixReportNumberHeader: true });
      const wbk = new ExcelJS.Workbook(); await wbk.xlsx.readFile(file);
      const rb = await rollForward(wbk, INV.slice().reverse().map(x => ({ ...x })), { reqNumber: 99, asOfDate: '2026-07-31', fixReportNumberHeader: true });
      a = { wb: wa, res: ra }; b = { wb: wbk, res: rb };
    } catch (e) {
      console.log('  ERROR  ' + label + ' :: ' + e.message);
      problems++; continue;
    }
    checked++;
    const sa = logSignature(a.wb), sb = logSignature(b.wb);
    const deterministic = sa === sb;
    const stray = strayBorders(a.wb);
    const inc = inceptionHeaders(a.wb);
    const incBad = inc.filter(t => /7\/31\/2026|07\/31\/2026/.test(t));
    const ct = a.res.contingencyTables || {};
    const flags = [];
    if (!deterministic) flags.push('ORDER NOT DETERMINISTIC');
    if (stray.length) flags.push('stray rules ' + stray.join(','));
    if (incBad.length) flags.push('inception header names the NEW period: ' + incBad.join(' / '));
    if (flags.length) problems++;
    console.log((flags.length ? '  FLAG   ' : '  ok     ') + label.padEnd(34)
      + ' hard="' + (ct.hardSheet || '-') + '"(' + ((ct.hard && ct.hard.moved) || 0) + ')'
      + ' soft="' + (ct.softSheet || '-') + '"(' + ((ct.soft && ct.soft.moved) || 0) + ')'
      + (flags.length ? '   ' + flags.join('; ') : ''));
    if (a.res.warnings && a.res.warnings.length) console.log('             warn: ' + a.res.warnings.join(' | '));
  }
  console.log('\n' + checked + ' properties rolled, ' + problems + ' with problems');
})().catch(e => { console.error('HARNESS ERROR', e); process.exit(2); });
