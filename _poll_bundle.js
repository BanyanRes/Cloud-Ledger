(async()=>{
  const get=async()=>{const html=await fetch('https://cloud-ledger.up.railway.app/',{cache:'no-store'}).then(r=>r.text());const m=html.match(/assets\/index-[A-Za-z0-9_-]+\.js/);return m?m[0]:null;};
  const start=await get();console.log('bundle now:',start);
  for(let i=0;i<40;i++){
    await new Promise(s=>setTimeout(s,15000));
    const cur=await get();
    if(cur&&cur!==start){console.log('NEW BUNDLE after ~'+((i+1)*15)+'s:',cur);process.exit(0);}
  }
  console.log('timeout');process.exit(1);
})();
