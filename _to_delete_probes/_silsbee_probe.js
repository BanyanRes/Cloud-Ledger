const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const files = {
  v1: D+'CL Items/Requisition_Report Silsbee July v1.xlsx',
  v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx',
  P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx',
  P2: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 2.xlsx',
};
(async () => {
  for (const [k,f] of Object.entries(files)) {
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.readFile(f); } catch(e){ console.log(k,'ERR',e.message); continue; }
    console.log('==== '+k+' ====');
    wb.worksheets.forEach(ws=>console.log('   ['+ws.name+'] rows='+ws.rowCount+' cols='+ws.columnCount));
  }
})();
