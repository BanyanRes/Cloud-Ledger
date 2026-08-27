const ExcelJS = require('exceljs');
const f='C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/06 Jun 2026/00 B1 County Line Rail Silsbee LLC - Requisition Report 06.2026 Phase 1.xlsx';
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '='+v.formula; if(v.result!==undefined) return String(v.result);} return String(v); };
(async()=>{ const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f); const b=wb.getWorksheet('Budget to Actual P1');
 for(let r=1;r<=9;r++){const o=[];for(let c=1;c<=13;c++){const t=S(b.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log(o.join(' | '));}
})();
