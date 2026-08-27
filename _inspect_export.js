// Read the exported Custom Detail workbook back and report what CLA asked about:
// the Doc # column, the rule under each account's last amount, the rule under each
// subtotal, the double rule on the grand total, and whether the totals are still
// live formulas over the right ranges after the column shift.
const ExcelJS = require('exceljs');
const path = process.argv[2];

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const ws = wb.worksheets[0];
  const txt = (r, c) => { const v = ws.getCell(r, c).value; return v == null ? '' : (typeof v === 'object' ? (v.formula ? '=' + v.formula : (v.result != null ? v.result : JSON.stringify(v))) : v); };
  const bot = (r, c) => { const b = ws.getCell(r, c).border; return b && b.bottom ? b.bottom.style : ''; };

  console.log('sheet:', ws.name, '| rows:', ws.rowCount, '| cols:', ws.columnCount);
  console.log('');
  console.log('--- header row and the shape of the sheet -----------------------');
  for (let r = 1; r <= Math.min(9, ws.rowCount); r++) {
    const cells = [];
    for (let c = 1; c <= ws.columnCount; c++) cells.push(String(txt(r, c)).slice(0, 34));
    console.log(String(r).padStart(3) + ' | ' + cells.join(' | '));
  }
  console.log('');
  console.log('--- every row carrying a rule, and every total row --------------');
  const amtCol = ws.columnCount; // last column holds the amount on this layout
  for (let r = 1; r <= ws.rowCount; r++) {
    const label = String(txt(r, 1));
    const b = bot(r, amtCol);
    const isTotal = /^Total|^PERIOD ACTIVITY|^ENDING BALANCE/.test(label);
    if (!b && !isTotal) continue;
    console.log(String(r).padStart(3) + ' ' + (b || 'none').padEnd(7) + ' | ' + label.slice(0, 52).padEnd(52) + ' | ' + String(txt(r, amtCol)).slice(0, 34));
  }
  console.log('');
  console.log('--- the July/August bills CLA said were missing -----------------');
  for (let r = 1; r <= ws.rowCount; r++) {
    const d = String(txt(r, 2));
    if (!/^2026-0[78]/.test(d)) continue;
    console.log('  ' + d + '  doc=' + String(txt(r, 3)).padEnd(20) + ' ' + String(txt(r, 4)).padEnd(9) + ' ' + String(txt(r, 5)).slice(0, 44).padEnd(44) + ' ' + txt(r, amtCol));
  }
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
