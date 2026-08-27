const ExcelJS=require('exceljs'), fs=require('fs');
(async()=>{
  const wb=new ExcelJS.Workbook(); await wb.xlsx.load(fs.readFileSync('_p1.xlsx'));
  const buf=await wb.xlsx.writeBuffer(); fs.writeFileSync('_plain.xlsx', Buffer.from(buf));
  console.log('plain round-trip wrote', buf.byteLength);
})().catch(e=>console.log('ERR',e.message));
