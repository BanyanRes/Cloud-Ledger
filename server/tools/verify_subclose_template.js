// Tie-out verification for the rebuilt subclose_template.xlsx.
const path=require('path'), fs=require('fs'), JSZip=require('jszip'), ExcelJS=require('exceljs');
const decode=s=>String(s==null?'':s).replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&apos;/g,"'").replace(/&quot;/g,'"').replace(/&amp;/g,'&');
const A=p=>path.join(__dirname,'..','assets',p);
const rowOf=a=>+a.match(/\d+/)[0];

function numVal(cell){const v=cell.value; if(v&&typeof v==='object')return ('result'in v)?v.result:(typeof v.formula==='undefined'?null:null); return typeof v==='number'?v:null;}
function txtVal(cell){const v=cell.value; if(v==null)return ''; if(typeof v==='object'){if(v.richText)return v.richText.map(t=>t.text).join(''); if('result'in v)return v.result; if('text'in v)return v.text; return '';} return v;}

(async()=>{
  // original external cache
  const zipSrc=await JSZip.loadAsync(fs.readFileSync(A('subclose_template_src.xlsx')));
  const el1=await zipSrc.file('xl/externalLinks/externalLink1.xml').async('string');
  const names=[...el1.matchAll(/<sheetName val="([^"]*)"/g)].map(m=>decode(m[1]));
  const cache={};
  for(const sb of el1.matchAll(/<sheetData sheetId="(\d+)"[^>]*>([\s\S]*?)<\/sheetData>/g)){
    const nm=names[+sb[1]]; const map={};
    for(const c of sb[2].matchAll(/<cell r="([^"]*)"([^>]*)>(?:<v>([\s\S]*?)<\/v>)?<\/cell>/g)){
      const t=(c[2].match(/t="([^"]*)"/)||[])[1]; const raw=c[3];
      map[c[1].replace(/\$/g,'')]= raw==null||raw===''?null : (t==='str'||t==='e'?decode(raw):(t==='b'?raw==='1':Number(raw)));
    }
    cache[nm]=map;
  }

  // A. no [1] remains in output
  const zipOut=await JSZip.loadAsync(fs.readFileSync(A('subclose_template.xlsx')));
  let ext1=0,sheetXmls=Object.keys(zipOut.files).filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  for(const n of sheetXmls){ const x=await zipOut.file(n).async('string'); ext1+=(x.match(/\[1\]/g)||[]).length; }
  const hasExtLinks=Object.keys(zipOut.files).some(n=>/externalLink/i.test(n));
  console.log('A. remaining [1] refs in output:',ext1,' | externalLink parts present:',hasExtLinks);

  // load output workbook
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(A('subclose_template.xlsx'));
  console.log('   sheets:',wb.worksheets.map(w=>w.name).join(' | '));
  const sum=wb.getWorksheet('Sub Close Summary');

  // src workbook (for original stored results)
  const wbSrc=new ExcelJS.Workbook(); await wbSrc.xlsx.readFile(A('subclose_template_src.xlsx'));
  const sumSrc=wbSrc.getWorksheet('Sub Close Summary');

  const close=(a,b,tol=0.01)=>{if(typeof a!=='number'||typeof b!=='number')return a===b; return Math.abs(a-b)<=tol;};

  // B. direct single-cell external refs (5 cached sheets): supporting tab value == cache value
  let bChk=0,bBad=0; const bEx=[];
  for(const sheet of ['Subsequent Close','Closing Interest','Management Fees - New Investor','Subsequent Close - OLD B4 25k','True-Up']){
    const ws=wb.getWorksheet(sheet); const cc=cache[sheet]||{};
    for(const [addr,cv] of Object.entries(cc)){
      if(cv==null)continue;
      const got = typeof cv==='number' ? numVal(ws.getCell(addr)) : txtVal(ws.getCell(addr));
      bChk++;
      if(!close(got,cv,0.005)){ bBad++; if(bEx.length<10)bEx.push(`${sheet}!${addr} exp=${cv} got=${got}`);}
    }
  }
  console.log(`B. cached-sheet cell fidelity: checked=${bChk} mismatches=${bBad}`, bEx.length?bEx:'');

  // Build lookup maps from output supporting tabs
  function lookupMap(sheet, valCol){
    const ws=wb.getWorksheet(sheet); const m={};
    ws.eachRow(row=>{ const key=txtVal(row.getCell('A')); const val=numVal(row.getCell(valCol)); if(key&&key!=='Investor Name'&&typeof val==='number') m[key]=val; });
    return m;
  }
  const TU_S=lookupMap('True-Up','S'), TU_U=lookupMap('True-Up','U'), TU_V=lookupMap('True-Up','V');
  const DC_Y=lookupMap('May 26 Distribution Checking','Y'), DC_AA=lookupMap('May 26 Distribution Checking','AA');

  // C+D. XLOOKUP tie-out against original stored results
  let xChk=0,xBad=0; const xEx=[];
  const r2=x=>Math.round(x*100)/100;
  sumSrc.eachRow((row)=>{
    row.eachCell((cell)=>{
      const v=cell.value; if(!v||typeof v!=='object'||typeof v.formula!=='string')return;
      const f=v.formula; const orig=('result'in v)?v.result:null;
      if(typeof orig!=='number')return;
      const r=rowOf(cell.address); const name=txtVal(sumSrc.getCell('E'+r));
      const sign=f.trim().startsWith('-')?-1:1; // formula may negate the ROUND(XLOOKUP())
      let computed=null;
      if(f.includes('True-Up')){
        const col=/!\$?S:\$?S/.test(f)?TU_S:/!\$?U:\$?U/.test(f)?TU_U:/!\$?V:\$?V/.test(f)?TU_V:null;
        if(col){ const raw=col[name]; if(typeof raw==='number') computed=sign*r2(raw); }
      } else if(f.includes('May 26 Distribution Checking')){
        // DC tab was reconstructed so that sign*ROUND(lookup) reproduces the cached result
        if(/!\$?Y:\$?Y/.test(f)){ const raw=DC_Y[name]; if(typeof raw==='number') computed=sign*r2(raw); }
        else if(/!\$?AA:\$?AA/.test(f)){ const raw=DC_AA[name]; if(typeof raw==='number') computed=sign*r2(raw); }
      } else return;
      xChk++;
      if(!close(computed,orig,0.02)){ xBad++; if(xEx.length<12)xEx.push(`${cell.address} name="${name}" exp=${orig} got=${computed}`);}
    });
  });
  console.log(`C/D. XLOOKUP (True-Up + Dist Checking) tie-out: checked=${xChk} mismatches=${xBad}`, xEx.length?xEx:'');

  // E. spot-check the visible summary direct-arith cells vs original result
  console.log('\nE. visible summary spot-check (recompute from Subsequent Close tab):');
  const sc=wb.getWorksheet('Subsequent Close');
  const scn=a=>numVal(sc.getCell(a));
  const b4=scn('T82')-scn('M82')+scn('W82')-scn('P82')+scn('X82')-scn('Q82');
  const b5=scn('T85')+scn('W85')+scn('X85');
  const b12=scn('R86'); const b13=-scn('N86');
  const chk=(lbl,got,exp)=>console.log(`   ${lbl}: computed=${got}  original=${exp}  ${close(got,exp,0.02)?'OK':'*** MISMATCH'}`);
  chk('B4',b4,numVal(sumSrc.getCell('B4')));
  chk('B5',b5,numVal(sumSrc.getCell('B5')));
  chk('B12',b12,numVal(sumSrc.getCell('B12')));
  chk('B13',b13,numVal(sumSrc.getCell('B13')));
})().catch(e=>{console.error(e);process.exit(1);});
