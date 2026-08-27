const ExcelJS = require('exceljs');
const f='C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/01 Silsbee/02 Requisition Report/2026/06 Jun 2026/00 B1 County Line Rail Silsbee LLC - Requisition Report 06.2026 Phase 1.xlsx';
const S = c => { const v=c.value; if(v==null) return ''; if(typeof v==='object'){ if(v.richText) return v.richText.map(t=>t.text).join(''); if(v.formula) return '='+v.formula; if(v.result!==undefined) return String(v.result);} return String(v); };
const bdesc = b => !b?'':['top','bottom'].filter(k=>b[k]&&b[k].style).map(k=>k[0]+':'+b[k].style).join(',');
(async()=>{
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(f);
  console.log('SHEETS: '+wb.worksheets.map(w=>w.name+'('+w.rowCount+'r/'+w.columnCount+'c)').join(', '));
  const ci=wb.getWorksheet('Current Invoice Log P1');
  console.log('--- JUNE Current Invoice Log P1 group order + borders');
  for(let r=1;r<=Math.min(ci.rowCount,100);r++){ const code=S(ci.getCell(r,3)), nm=S(ci.getCell(r,4)), amt=S(ci.getCell(r,7)); const bd=bdesc(ci.getCell(r,7).border);
    if(code||/Total/i.test(nm)||bd) console.log('  r'+r+' code='+code+' | '+nm+' | amt='+amt+(bd?'  [G '+bd+']':'')); }
  const df=wb.getWorksheet('Dev Fee P1');
  console.log('--- JUNE Dev Fee P1');
  for(let r=1;r<=Math.min(df.rowCount,30);r++){const o=[];for(let c=1;c<=8;c++){const t=S(df.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log('  '+o.join(' | '));}
  const h=wb.getWorksheet('Hard Cost Contigency P1');
  console.log('--- JUNE Hard Cost Contigency P1');
  for(let r=1;r<=Math.min(h.rowCount,32);r++){const o=[];for(let c=2;c<=8;c++){const t=S(h.getCell(r,c)); if(t)o.push(String.fromCharCode(64+c)+r+'='+t);} if(o.length)console.log('  '+o.join(' | '));}
})();
