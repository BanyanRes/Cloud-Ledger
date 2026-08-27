const fs=require('fs');
const setup=JSON.parse(fs.readFileSync('C:/Users/JimmyYun/Downloads/_iv_setup.json','utf8'));
const TOK=setup.token, BASE='https://cloud-ledger.up.railway.app';
const H={Authorization:'Bearer '+TOK};
(async()=>{
  // 1) upload Q1 investment template
  const p='C:/Users/JimmyYun/OneDrive - banyanres.com/Desktop/CLRF Investment Balance 3-31-26_updated by JY.xlsx';
  const buf=fs.readFileSync(p);
  const fd=new FormData();
  fd.append('folder_path','Workpapers/Investment & Valuation/Q1 2026');
  fd.append('files',new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),'CLRF Investment Balance 3-31-26.xlsx');
  const up=await fetch(BASE+'/api/entities/40/files',{method:'POST',headers:H,body:fd});
  console.log('UPLOAD',up.status,JSON.stringify(await up.json()));
  // 2) poll for new deploy: new endpoint returns non-404 when live
  for(let i=0;i<40;i++){
    const r=await fetch(BASE+'/api/workpapers/investment-valuation/40/generate',{method:'POST',headers:{...H,'Content-Type':'application/json'},body:JSON.stringify({})});
    if(r.status!==404){console.log('DEPLOY LIVE after ~'+(i*15)+'s, probe status',r.status,JSON.stringify(await r.json()));break;}
    if(i===39){console.log('TIMED OUT waiting for deploy');process.exit(1);}
    await new Promise(s=>setTimeout(s,15000));
  }
})().catch(e=>{console.error('FAIL',e);process.exit(1);});
