const ExcelJS = require('exceljs');
const D = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/07 Jul 2026/';
const F = { v2: D+'CL Items/Requisition_Report Silsbee July v2.xlsx', P1: D+'00 B1 County Line Rail Silsbee LLC - Requisition Report 07.2026 Phase 1.xlsx' };
const num = c => { const v=c.value; if(v==null)return null; if(typeof v==='number')return v; if(typeof v==='object'&&typeof v.result==='number')return v.result; return null; };
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '=' +v.formula; if(v.result!==undefined) return String(v.result);} return String(v); };
const bdesc = b => !b?'':['top','bottom','left','right'].filter(k=>b[k]&&b[k].style).map(k=>k[0]+':'+b[k].style).join(',');
(async()=>{
 for (const [k,f] of Object.entries(F)){
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  const b=wb.getWorksheet('Budget to Actual P1');
  console.log('##### '+k+' — B2A rows where Balance Remaining (L) < 0');
  for(let r=9;r<=120;r++){ const L=num(b.getCell(r,12)); if(L!=null && L<-0.005){ console.log('   r'+r+' code='+S(b.getCell(r,2))+' name='+S(b.getCell(r,3))+' D='+num(b.getCell(r,4))+' E='+num(b.getCell(r,5))+' F='+num(b.getCell(r,6))+' G='+num(b.getCell(r,7))+' J='+num(b.getCell(r,10))+' L='+L); } }
  console.log('##### '+k+' — Current Invoice Log P1 borders rows 4..30, cols A..K');
  const ci=wb.getWorksheet('Current Invoice Log P1');
  for(let r=4;r<=Math.min(ci.rowCount,30);r++){ const o=[]; for(let c=1;c<=11;c++){ const d=bdesc(ci.getCell(r,c).border); if(d)o.push(String.fromCharCode(64+c)+':'+d);} if(o.length)console.log('   r'+r+'  '+o.join('  ')); }
 }
})();
