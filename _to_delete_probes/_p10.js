const ExcelJS = require('exceljs');
const J='C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/06 Jun 2026/00 B1 County Line Rail Silsbee LLC - Requisition Report 06.2026 Phase 1.xlsx';
const num = c => { const v=c.value; if(v==null)return null; if(typeof v==='number')return v; if(typeof v==='object'&&typeof v.result==='number')return v.result; return null; };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '='+v.formula;} return String(v); };
(async()=>{ const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(J); const b=wb.getWorksheet('Budget to Actual P1');
 for(const r of [25,26,27,65]) console.log('JUNE r'+r+' code='+S(b.getCell(r,2))+' name='+S(b.getCell(r,3))+' D='+num(b.getCell(r,4))+' E='+num(b.getCell(r,5))+' F='+num(b.getCell(r,6)));
})();
