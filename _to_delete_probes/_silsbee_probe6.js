const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const num = c => { const v=c.value; if(v==null)return null; if(typeof v==='number')return v; if(typeof v==='object'&&typeof v.result==='number')return v.result; return null; };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '='+v.formula; if(v.result!==undefined) return String(v.result);} return String(v); };
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  const b=wb.getWorksheet('Budget to Actual P1');
  console.log('##### '+k+' — B2A rows with F (Current Contingency) != 0');
  for(let r=9;r<=120;r++){ const v=num(b.getCell(r,6)); if(v!=null&&Math.abs(v)>0.004) console.log('   r'+r+' '+S(b.getCell(r,2))+' '+S(b.getCell(r,3))+'  E='+num(b.getCell(r,5))+' F='+v+' G='+num(b.getCell(r,7))+' J='+num(b.getCell(r,10))+' L='+num(b.getCell(r,12))); }
  console.log('##### '+k+' — accounting line 12230');
  for(let r=9;r<=120;r++){ if(S(b.getCell(r,2))==='12230'){ console.log('   r'+r+' D='+num(b.getCell(r,4))+' E='+num(b.getCell(r,5))+' F='+num(b.getCell(r,6))+' G='+S(b.getCell(r,7))+'->'+num(b.getCell(r,7))+' H='+num(b.getCell(r,8))+' I='+num(b.getCell(r,9))+' J='+num(b.getCell(r,10))+' L='+S(b.getCell(r,12))+'->'+num(b.getCell(r,12))); } }
 }
})();
