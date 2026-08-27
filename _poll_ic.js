const BASE='https://cloud-ledger.up.railway.app';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
(async()=>{
  const before=(await (await fetch(BASE)).text()).match(/assets\/index-[A-Za-z0-9_-]+\.js/)[0];
  console.log('bundle before:', before);
  for(let i=0;i<60;i++){
    await sleep(10000);
    const html=await (await fetch(BASE)).text();
    const now=(html.match(/assets\/index-[A-Za-z0-9_-]+\.js/)||[''])[0];
    if(now&&now!==before){
      console.log('bundle now:  ', now, '(after', (i+1)*10, 's)');
      const js=await (await fetch(BASE+'/'+now)).text();
      const marks=['Show unmapped accounts','What the name reads as','only names that read as intercompany',
        'Counterparty, account code or name','Save mapping','not an intercompany name'];
      for(const m of marks) console.log(js.includes(m)?'  OK   ':'  MISS ', m);
      for(const p of ['/api/intercompany/accounts/unmapped?entity_id=1','/api/intercompany/reconcile/entity?entity_id=1']){
        const r=await fetch(BASE+p); console.log('  ',r.status,p);
      }
      return;
    }
    process.stdout.write('.');
  }
  console.log('\ntimed out waiting for a new bundle');
})();
