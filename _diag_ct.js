const fs=require('fs'),path=require('path'),ExcelJS=require('exceljs');
const ROOT='C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents';
const src=fs.readFileSync('server/requisition_rollforward.js','utf8').replace(/\r\n/g,'\n');
const S=c=>{const v=c.value;if(v==null)return'';if(typeof v==='object'){if(v.richText)return v.richText.map(t=>t.text).join('');if(v.formula)return'='+v.formula;if(v.result!==undefined)return String(v.result);}return String(v);};
const layout=ws=>{let p=false,h=false;for(let r=1;r<=12;r++){if(/previous/i.test(S(ws.getCell(r,4))))p=true;if(/herein/i.test(S(ws.getCell(r,5))))h=true;}return p&&h;};
function newest(dir){let b=null,t=0;const walk=(d,dep)=>{if(dep>4)return;let es=[];try{es=fs.readdirSync(d,{withFileTypes:true})}catch(e){return}
 for(const e of es){const p=path.join(d,e.name);if(e.isDirectory())walk(p,dep+1);else if(/\.xlsx$/i.test(e.name)&&!/^~\$/.test(e.name)&&/requisition|req\b/i.test(e.name)){const m=fs.statSync(p).mtimeMs;if(m>t){t=m;b=p}}}};walk(dir,0);return b;}
(async()=>{
for(const d of fs.readdirSync(ROOT,{withFileTypes:true}).filter(e=>e.isDirectory()&&!/^(AP|General|Insurance|_|z)/.test(e.name))){
  const p=path.join(ROOT,d.name);
  const rd=fs.readdirSync(p,{withFileTypes:true}).filter(e=>e.isDirectory()&&/requisition/i.test(e.name)).map(e=>path.join(p,e.name))[0];
  if(!rd)continue; const f=newest(rd); if(!f)continue;
  const wb=new ExcelJS.Workbook(); try{await wb.xlsx.readFile(f)}catch(e){continue}
  const cands=wb.worksheets.filter(w=>/contin?gency/i.test(w.name));
  console.log('== '+d.name+'  <- '+path.basename(f));
  for(const w of cands) console.log('     "'+w.name+'"  D/E layout: '+layout(w));
}
})();
