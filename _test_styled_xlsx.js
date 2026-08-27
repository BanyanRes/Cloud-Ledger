// Verify xlsxStyledReport actually writes the three underline styles CLA asked
// for, and that live formulas survive. Reads the produced workbook back with
// ExcelJS and asserts on real cell borders rather than trusting the write.
const ExcelJS = require('exceljs');
const S = require('./server/xlsxStyledReport.js');

const rows = [
  ['Banyan Residential'],                                          // 0 title
  ['Custom Detail Report'],                                        // 1 title
  ['Project: P-10100.902 - Van Buren'],                            // 2 meta
  [],                                                              // 3
  ['Account', 'Date', 'Doc #', 'JE', 'Description', 'Amount'],     // 4 header
  ['11760 Other Due Diligence', '2026-06-30', '0708137', 'JE-1595', 'Bill - Novogradac', 7000],   // 5
  ['11760 Other Due Diligence', '2026-07-17', '3092', 'JE-1605', 'Bill - Tatum', 4000],           // 6 last detail
  ['Total 11760 Other Due Diligence', '', '', '', '', 11000],       // 7 subtotal
  ['PERIOD ACTIVITY', '', '', '', '', 11000],                       // 8 grand total
];
const formulas = [
  { r: 7, c: 5, f: 'SUM(F6:F7)' },
  { r: 8, c: 5, f: 'F8' },
];
const style = {
  titleRows: [0, 1], metaRows: [2], headerRows: [4],
  underlineRows: [6, 7], doubleUnderlineRows: [8], amountCols: [5],
};

(async () => {
  const buf = await S.buildStyledWorkbookBuffer({ rows, formulas, style, filename: 'test.xlsx' });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet('Report');
  const bottom = (r, c) => { const b = ws.getCell(r + 1, c + 1).border; return b && b.bottom ? b.bottom.style : null; };
  const checks = [
    ['header row 4 col5 thin', bottom(4, 5), 'thin'],
    ['last detail row 6 col5 thin', bottom(6, 5), 'thin'],
    ['subtotal row 7 col5 thin', bottom(7, 5), 'thin'],
    ['grand total row 8 col5 double', bottom(8, 5), 'double'],
    ['detail row 5 col5 no rule', bottom(5, 5), null],
    ['label col of subtotal has no rule', bottom(7, 0), null],
    ['subtotal is a formula', ws.getCell(8, 6).value && ws.getCell(8, 6).value.formula, 'SUM(F6:F7)'],
    ['subtotal keeps cached result', ws.getCell(8, 6).value && ws.getCell(8, 6).value.result, 11000],
    ['detail amount stays a value', ws.getCell(6, 6).value, 7000],
    ['subtotal bold', !!(ws.getCell(8, 1).font || {}).bold, true],
    ['title bold 13', ((ws.getCell(1, 1).font || {}).size), 13],
    ['money format on amount', ws.getCell(6, 6).numFmt, '#,##0.00;(#,##0.00)'],
  ];
  let bad = 0;
  for (const [label, got, want] of checks) {
    const ok = String(got) === String(want);
    if (!ok) bad++;
    console.log((ok ? 'PASS  ' : 'FAIL  ') + label + '  got=' + JSON.stringify(got) + ' want=' + JSON.stringify(want));
  }
  console.log(bad === 0 ? '\nALL PASS' : '\n' + bad + ' FAILED');
  process.exit(bad === 0 ? 0 : 1);
})();
