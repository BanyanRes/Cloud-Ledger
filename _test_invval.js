const fs=require('fs');
const JSZip=require('jszip');
const iv=require('./server/invval');
(async()=>{
  const tpl=fs.readFileSync('_q1_tpl.xlsx');
  const zip=await JSZip.loadAsync(tpl);
  const P=await (async()=>{const wb=await zip.file('xl/workbook.xml').async('string');const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');const rid={};for(const m of rels.matchAll(/<Relationship\b[^>]*>/g)){const id=(m[0].match(/Id="(rId\d+)"/)||[])[1];const tg=(m[0].match(/Target="([^"]+)"/)||[])[1];if(id&&tg)rid[id]=tg;}const p={};for(const m of wb.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g))if(rid[m[2]])p[m[1].replace(/&amp;/g,'&').trim()]='xl/'+rid[m[2]].replace(/^\//,'');return p;})();
  const params=await iv.parseModelParams(zip,P);
  console.log('PARAMS ownPct',params.clipOwnPct,params.silOwnPct,'uscPct',params.uscPct);
  console.log('PARAMS em',JSON.stringify(params.em),'sil',JSON.stringify(params.sil));
  console.log('PARAMS usc',JSON.stringify(params.usc));
  console.log('PRIOR VALS',JSON.stringify(params.priorVals));
  // Q2 data
  const tb=JSON.parse(fs.readFileSync('C:/Users/JimmyYun/Downloads/_tb_63026.json','utf8'));
  const dp=r=>(r.type==='Asset'||r.type==='Expense')?r.balance:-r.balance;
  const r2=n=>Math.round(n*100)/100;
  const bal=(k,c)=>{const r=tb[k].closing.find(x=>x.code===c);return r?r2(dp(r)):0;};
  const port={
    clip:{nwc:7337275.42,loanBal:bal('clip','25063')},
    silsbee:{nwc:243613.24,loanBal:bal('silsbee','25063')},
    buna:{nwc:-6157382.07,loanBal:bal('buna','25063')},
    srn:{nwc:2139585.83,loanBal:bal('srn','25063')},
  };
  const books={clip:65319338,silsbee:8787580,buna:922600,srn:60408356};
  const fvAdj={clip:141167,silsbee:0,buna:0,srn:0};
  const serial=(()=>{const [y,m,d]=[2026,6,30];return Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/864e5);})();
  const model=iv.makeModel(params,serial);
  const solve=iv.solveValuations(model,port,books,fvAdj,params.priorVals);
  console.log('SOLVE clip',solve.clip.valuation,'proceeds',solve.clip.proceeds,'promote',solve.clip.promote);
  console.log('SOLVE silsbee',solve.silsbee.valuation,solve.silsbee.proceeds,'changed',solve.silsbee.changed);
  console.log('SOLVE buna',solve.buna.valuation,solve.buna.proceeds,'changed',solve.buna.changed);
  console.log('SOLVE srn',solve.srn.valuation,solve.srn.proceeds,'changed',solve.srn.changed);
  const expect={clip:151343839.64,silsbee:28550000,buna:11750000,srn:79230000};
  let ok=true;
  for(const k of Object.keys(expect)){if(Math.abs(solve[k].valuation-expect[k])>0.01){console.log('MISMATCH',k,solve[k].valuation,'vs',expect[k]);ok=false;}}
  // full portfolio account maps for the builder
  const build=(k)=>{const acc={};const t=c=>(acc[c]=acc[c]||{open:0,close:0,td:0,tc:0,name:'',type:''});
    for(const r of tb[k].opening){const a=t(r.code);a.open=r2(dp(r));a.name=r.name;a.type=r.type;}
    for(const r of tb[k].closing){const a=t(r.code);a.close=r2(dp(r));a.name=r.name;a.type=r.type;}
    for(const r of tb[k].activity){const a=t(r.code);a.td=r2(r.total_debit);a.tc=r2(r.total_credit);if(!a.name){a.name=r.name;a.type=r.type;}}
    const codes=Object.keys(acc).filter(c=>{const a=acc[c];return Math.abs(a.open)>0.005||Math.abs(a.close)>0.005||a.td>0.005||a.tc>0.005;}).sort((x,y)=>Number(x)-Number(y));
    return {acc,codes};};
  for(const k of Object.keys(port))Object.assign(port[k],build(k));
  const inv=await iv.buildInvestmentWorkbook(tpl,{qtr:{end:'2026-06-30',year:'2026',quarter:'Q2',label:'Q2 2026',q:2},
    port,books,booksExact:{clip:65319338.46,silsbee:8787579.67,buna:922599.55,srn:60408356.37},fvAdj,solve,devTotal:54241570.57,params});
  fs.writeFileSync('_test_inv_out.xlsx',inv.buf);
  console.log('BUILT',inv.buf.length,'bytes; L=',JSON.stringify(inv.invBalance.L));
  // verify: orphans + refs + no calcChain
  const z2=await JSZip.loadAsync(inv.buf);
  let orphans=0,refs=0;
  const wb2=await z2.file('xl/workbook.xml').async('string');
  const rels2=await z2.file('xl/_rels/workbook.xml.rels').async('string');
  const rid2={};for(const m of rels2.matchAll(/<Relationship\b[^>]*>/g)){const id=(m[0].match(/Id="(rId\d+)"/)||[])[1];const tg=(m[0].match(/Target="([^"]+)"/)||[])[1];if(id&&tg)rid2[id]=tg;}
  for(const m of wb2.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)){
    const x=await z2.file('xl/'+rid2[m[2]].replace(/^\//,'')).async('string');
    const hosts=new Set([...x.matchAll(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"/g)].map(mm=>mm[1]));
    for(const mm of x.matchAll(/<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"/g))if(!hosts.has(mm[1]))orphans++;
    refs+=(x.match(/#REF!/g)||[]).length;
  }
  console.log('calcChain gone:',!z2.file('xl/calcChain.xml'),'orphans:',orphans,'refs:',refs);
  if(!(!z2.file('xl/calcChain.xml')&&orphans===0&&refs===2))ok=false;
  console.log(ok?'ALL TESTS PASS':'TEST FAILURES');
  process.exit(ok?0:1);
})().catch(e=>{console.error('FAIL',e.stack);process.exit(1);});
