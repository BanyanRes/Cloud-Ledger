const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '='+v.formula; if(v.result!==undefined) return String(v.result);} return String(v); };
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  const ci=wb.getWorksheet('Current Invoice Log P1');
  console.log('##### '+k+' Current log cols L..Y, rows 1..25');
  for(let r=1;r<=25;r++){const o=[];for(let c=12;c<=25;c++){const t=S(ci.getCell(r,c)); if(t)o.push(c+':'+t);} if(o.length)console.log('  r'+r+' '+o.join(' | '));}
  const df=wb.getWorksheet('Dev Fee P1');
  console.log('##### '+k+' Dev Fee P1 rows 1..40 cols A..H');
  for(let r=1;r<=40;r++){const o=[];for(let c=1;c<=8;c++){const t=S(df.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log('  '+o.join(' | '));}
 }
})();
