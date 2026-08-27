const fs=require('fs');
const setup=JSON.parse(fs.readFileSync('C:/Users/JimmyYun/Downloads/_iv_setup.json','utf8'));
(async()=>{
  const t0=Date.now();
  const r=await fetch('https://cloud-ledger.up.railway.app/api/workpapers/investment-valuation/40/generate',{
    method:'POST',
    headers:{Authorization:'Bearer '+setup.token,'Content-Type':'application/json'},
    body:JSON.stringify({quarter_end:'2026-06-30'})});
  const d=await r.json();
  console.log('STATUS',r.status,'in',Math.round((Date.now()-t0)/1000)+'s');
  console.log(JSON.stringify(d,null,1));
  // verify against the locally validated solve
  const exp={clip:151343839.64,silsbee:28550000,buna:11750000,srn:79230000};
  let ok=r.status===200;
  if(d.solve)for(const k of Object.keys(exp)){
    if(Math.abs(d.solve[k].valuation-exp[k])>0.01){console.log('MISMATCH',k,d.solve[k].valuation,'vs',exp[k]);ok=false;}}
  else ok=false;
  if(d.unrealized&&!(d.unrealized.clip===141167&&d.unrealized.silsbee===0&&d.unrealized.buna===0&&d.unrealized.srn===0)){console.log('UNREALIZED MISMATCH');ok=false;}
  console.log(ok?'LIVE SOLVE MATCHES LOCAL':'VERIFY FAILED');
  process.exit(ok?0:1);
})().catch(e=>{console.error('FAIL',e);process.exit(1);});
