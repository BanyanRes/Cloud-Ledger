const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.formula) return '='+v.formula; if(v.result!==undefined) return JSON.stringify(v.result); if(v.richText) return v.richText.map(t=>t.text).join(''); if(v instanceof Date) return v.toISOString().slice(0,10); return JSON.stringify(v);} return String(v); };
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  console.log('########## '+k+' ##########');
  const b=wb.getWorksheet('Budget to Actual P1');
  console.log('--- B2A header rows 5..12, cols A..N');
  for(let r=5;r<=12;r++){const o=[];for(let c=1;c<=14;c++){const t=S(b.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log('  '+o.join(' | '));}
  console.log('--- B2A sample data rows 13..30, cols A..N');
  for(let r=13;r<=30;r++){const o=[];for(let c=1;c<=14;c++){const t=S(b.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log('  '+o.join(' | '));}
  console.log('--- Current Invoice Log P1: group order (subtotal labels)');
  const ci=wb.getWorksheet('Current Invoice Log P1');
  let n=0;
  for(let r=1;r<=ci.rowCount;r++){ for(let c=1;c<=12;c++){ const t=S(ci.getCell(r,c)); if(/Total$/i.test(t)){ console.log('   r'+r+' c'+c+' '+t); n++; break; } } if(n>60)break; }
 }
})();
