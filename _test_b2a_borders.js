// Scratch: render the real SRN Requisition Report #14 "Budget to Actual" sheet
// through xlsxSheetToPdf to verify the workbook's own cell borders (main-table
// gridlines + reconciliation underlines) are now reproduced.
const fs = require('fs');
const { xlsxSheetToPdf } = require('./server/xlsxToPdf');
const { readSheetBorders } = require('./server/xlsxBorders');

const SRC = 'C:/Users/JimmyYun/OneDrive - banyanres.com/Desktop/0005 B1 County Line SRN Requisition Report #14 04.30.2026.xlsx';

(async () => {
  const buf = fs.readFileSync(SRC);
  // First, sanity-check how many border cells we extracted, for the B2A sheet.
  const XLSX = require('xlsx');
  const wb = XLSX.read(buf, { type: 'buffer', bookSheets: true });
  console.log('sheets:', wb.SheetNames);
  // Pick the Budget to Actual sheet (case-insensitive), else first.
  let target = wb.SheetNames.find(n => /budget\s*to\s*actual/i.test(n)) || wb.SheetNames[0];
  console.log('target sheet:', target);
  const b = await readSheetBorders(buf, target);
  console.log('border cells found:', b.size);
  // Show a small sample of what edges were detected.
  let shown = 0;
  for (const [k, v] of b) {
    if (shown++ >= 8) break;
    console.log('  ', k, JSON.stringify(v));
  }
  const conv = await xlsxSheetToPdf(buf, target, {});
  fs.writeFileSync('C:/Users/JimmyYun/_tmp_mgmt/_b2a_borders.pdf', Buffer.from(conv.bytes));
  console.log('wrote _b2a_borders.pdf; sheetUsed=', conv.sheetUsed);
})();
