const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v1: D+'CL Items/Requisition_Report Silsbee July v1.xlsx', v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.formula) return '{f}'; if(v.result!==undefined) return JSON.stringify(v.result); if(v.richText) return v.richText.map(t=>t.text).join(''); if(v instanceof Date) return v.toISOString().slice(0,10); return JSON.stringify(v);} if(v instanceof Date) return v.toISOString().slice(0,10); return String(v); };
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  const ci=wb.getWorksheet('Current Invoice Log P1');
  console.log('##### '+k+' Current Invoice Log P1 (rows 1..'+Math.min(ci.rowCount,95)+') cols A..K');
  for(let r=1;r<=Math.min(ci.rowCount,95);r++){
    const o=[]; for(let c=1;c<=11;c++){const t=S(ci.getCell(r,c)); o.push(t);} 
    if(o.some(x=>x!=='')) console.log(String(r).padStart(3)+': '+o.map((t,i)=>t?String.fromCharCode(65+i)+':'+t:'').filter(Boolean).join(' | '));
  }
 }
})();
