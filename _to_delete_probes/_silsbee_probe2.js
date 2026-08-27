const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.formula) return '='+v.formula+' ->'+JSON.stringify(v.result); if(v.result!==undefined) return JSON.stringify(v.result); if(v.richText) return v.richText.map(t=>t.text).join(''); return JSON.stringify(v);} return String(v); };
function dump(ws, r1, r2, c1, c2){ for(let r=r1;r<=r2;r++){ const cells=[]; for(let c=c1;c<=c2;c++){ const t=S(ws.getCell(r,c)); if(t!=='') cells.push(String.fromCharCode(64+c)+r+'='+t);} if(cells.length) console.log('  '+cells.join(' | ')); } }
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  console.log('########## '+k+' ##########');
  for (const nm of ['Hard cost contingency','Hard Cost Contigency P1','Soft Cost Contingency P1']){
    const ws=wb.getWorksheet(nm); if(!ws){console.log('-- missing '+nm);continue;}
    console.log('--- ['+nm+'] rows 1..'+Math.min(ws.rowCount,30));
    dump(ws,1,Math.min(ws.rowCount,30),1,8);
  }
 }
})();
