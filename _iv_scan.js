const fs=require('fs');const JSZip=require('jszip');
const setup=JSON.parse(fs.readFileSync('C:/Users/JimmyYun/Downloads/_iv_setup.json','utf8'));
const H={Authorization:'Bearer '+setup.token};
(async()=>{
  const files=await fetch('https://cloud-ledger.up.railway.app/api/entities/40/files',{headers:H}).then(r=>r.json());
  const list=Array.isArray(files)?files:(files.files||[]);
  const q2=list.filter(f=>f.folder_path==='Workpapers/Investment & Valuation/Q2 2026');
  let allOk=true;
  for(const f of q2){
    const buf=Buffer.from(await fetch('https://cloud-ledger.up.railway.app/api/entity-files/'+f.id+'/download?token='+encodeURIComponent(setup.token)).then(r=>r.arrayBuffer()));
    const zip=await JSZip.loadAsync(buf);
    const wb=await zip.file('xl/workbook.xml').async('string');
    const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rid={};for(const m of rels.matchAll(/<Relationship\b[^>]*>/g)){const id=(m[0].match(/Id="(rId\d+)"/)||[])[1];const tg=(m[0].match(/Target="([^"]+)"/)||[])[1];if(id&&tg)rid[id]=tg;}
    let orphans=0,refs=0,badName=0;
    for(const m of wb.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)){
      const p='xl/'+rid[m[2]].replace(/^\//,'');const x=await zip.file(p).async('string');
      const hosts=new Set([...x.matchAll(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"/g)].map(mm=>mm[1]));
      for(const mm of x.matchAll(/<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"/g))if(!hosts.has(mm[1]))orphans++;
      refs+=(x.match(/#REF!/g)||[]).length;
    }
    for(const m of wb.matchAll(/<definedName[^>]*>([^<]*)<\/definedName>/g))if(/(?<![A-Za-z0-9_'])SOI!/.test(m[1]))badName++;
    const cc=!!zip.file('xl/calcChain.xml');
    console.log(f.original_name+' ('+(buf.length/1e6).toFixed(1)+'MB): calcChain='+cc+' orphans='+orphans+' #REF tokens='+refs+' staleSOInames='+badName);
    // investment: 2 pre-existing USC K44 tokens allowed; valuation: known pre-existing #REF!s in appraiser tabs are expected, orphans/calcChain are the corruption signals
    if(cc||orphans>0||badName>0)allOk=false;
    if(/Investment/.test(f.original_name)&&refs!==2)allOk=false;
  }
  console.log(allOk?'INTEGRITY OK':'INTEGRITY ISSUES');
  process.exit(allOk?0:1);
})().catch(e=>{console.error('FAIL',e);process.exit(1);});
