// scratch patch — do not commit
const fs = require('fs');
function patch(P, edits) {
  let s = fs.readFileSync(P, 'utf8');
  const NL = s.includes('\r\n') ? '\r\n' : '\n';
  const lf = (str) => NL === '\r\n' ? str.replace(/\n/g, '\r\n') : str;
  for (const [find, replace, label] of edits) {
    const f = find.replace(/\n/g, NL);
    const n = s.split(f).length - 1;
    if (n !== 1) throw new Error(`[${P}] anchor "${label}" found ${n} times (want 1)`);
    s = s.replace(f, lf(replace));
  }
  fs.writeFileSync(P, s);
  console.log(P, 'patched OK; NL =', JSON.stringify(NL));
}

patch('client/src/api.js', [
  [
`  getArOpenInvoices: (eid, amount) => request('/entities/' + eid + '/ar/open-invoices' + (amount != null ? '?amount=' + amount : '')),`,
`  getArOpenInvoices: (eid, amount) => request('/entities/' + eid + '/ar/open-invoices' + (amount != null ? '?amount=' + amount : '')),
  arOpeningImport: (eid, items, force) => request('/entities/' + eid + '/ar/opening-import', { method: 'POST', body: { items, force: !!force } }),`,
'api-openimport'],
]);

const modal = `function ArOpeningUploadModal({entityId,asOf,onClose,onImported}){
  const [parsing,setParsing]=useState(false);
  const [preview,setPreview]=useState(null);
  const [err,setErr]=useState('');
  const [force,setForce]=useState(false);
  const [saving,setSaving]=useState(false);
  const [fileName,setFileName]=useState('');
  const fmtDate=(v)=>{ if(v===null||v===undefined||v==='')return null; if(v instanceof Date&&!isNaN(v)){return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');} const s=String(v).trim(); if(!s)return null; if(s.indexOf('/')>=0){const p=s.split(' ')[0].split('/'); if(p.length===3){let y=p[2]; if(y.length===2)y='20'+y; return y+'-'+p[0].padStart(2,'0')+'-'+p[1].padStart(2,'0');}} if(s.length>=10&&s.charAt(4)==='-')return s.slice(0,10); return null; };
  const num=(v)=>{ if(v===null||v===undefined||v==='')return null; if(typeof v==='number')return v; const n=parseFloat(String(v).split(',').join('').split('$').join('')); return isNaN(n)?null:n; };
  const onFile=async(file)=>{
    setErr('');setPreview(null);setParsing(true);setFileName(file.name);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array',cellDates:true});
      let rows=null,headerIdx=-1;
      for(const name of wb.SheetNames){
        const rws=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''});
        for(let i=0;i<Math.min(rws.length,20);i++){
          const low=rws[i].map(c=>String(c).toLowerCase());
          if(low.some(c=>c.indexOf('document no')>=0)&&low.some(c=>c.indexOf('total')>=0)){ rows=rws;headerIdx=i;break; }
        }
        if(rows)break;
      }
      if(!rows)throw new Error('Could not find an aging detail sheet (needs a header row with Document no. and Total).');
      const H=rows[headerIdx].map(c=>String(c).toLowerCase().trim());
      const findCol=(...names)=>{ for(const nm of names){ for(let i=0;i<H.length;i++){ if(H[i].indexOf(nm)>=0)return i; } } return -1; };
      const ci={ id:findCol('customer id'), name:findCol('customer name'), doc:findCol('document no'), posting:findCol('gl posting','posting date'), inv:findCol('invoice date'), due:findCol('due date'), total:findCol('total') };
      if(ci.doc<0||ci.total<0)throw new Error('Missing a Document no. or Total column in the header row.');
      let fileAsOf=null;
      for(let i=0;i<headerIdx;i++){ const r=rows[i]; for(let c=0;c<r.length-1;c++){ if(String(r[c]).toLowerCase().indexOf('as of date')>=0){ const f=fmtDate(r[c+1]); if(f)fileAsOf=f; } } }
      const items=[]; let cust=null;
      for(let i=headerIdx+1;i<rows.length;i++){
        const r=rows[i];
        const idCell=String(ci.id>=0?r[ci.id]:'').trim();
        const nameCell=String(ci.name>=0?r[ci.name]:'').trim();
        const docCell=String(r[ci.doc]||'').trim();
        const low=idCell.toLowerCase();
        if(low.indexOf('total for')>=0)continue;
        if(low.indexOf('grand total')>=0)break;
        if(nameCell)cust=nameCell;
        if(!docCell)continue;
        const amt=num(r[ci.total]); if(amt===null)continue;
        items.push({customer_name:cust||'(no customer)',document_no:docCell,posting_date:ci.posting>=0?fmtDate(r[ci.posting]):null,invoice_date:ci.inv>=0?fmtDate(r[ci.inv]):null,due_date:ci.due>=0?fmtDate(r[ci.due]):null,amount:Math.round(amt*100)/100});
      }
      if(!items.length)throw new Error('No invoice rows found below the header row.');
      const total=Math.round(items.reduce((s,x)=>s+x.amount,0)*100)/100;
      const useAsOf=fileAsOf||asOf;
      let glBalance=null,recon=null;
      try{ const ag=await api.getArAging(entityId,useAsOf); glBalance=ag.gl_ar_balance; recon=Math.round((glBalance-total)*100)/100; }catch(e){}
      const custCount=new Set(items.map(x=>x.customer_name)).size;
      setPreview({items,count:items.length,total,asOf:useAsOf,glBalance,recon,custCount});
    }catch(e){ setErr(e.message||String(e)); } finally{ setParsing(false); }
  };
  const doImport=async()=>{ if(!preview)return; setSaving(true);setErr('');
    try{ const res=await api.arOpeningImport(entityId,preview.items,force); onImported(res); }
    catch(e){ setErr(e.message||String(e)); } finally{ setSaving(false); }
  };
  const tie=preview&&preview.recon!==null?Math.abs(preview.recon)<0.005:null;
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:640}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Upload A/R aging detail</div>
    <div style={{fontSize:12,color:T.textMuted,marginBottom:16}}>Import a prior-system aging detail (for example the 6/30/26 transition report) as opening A/R. Items age by GL posting date and tie to the GL control account; go-forward receipts are applied from bank coding. Re-uploading replaces the current opening set.</div>
    <div style={{...S.card,background:T.bgElevated,padding:16,marginBottom:14,textAlign:'center'}}>
      <input id="ar-open-file" type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>{const f=e.target.files[0];if(f)onFile(f); e.target.value='';}}/>
      <label htmlFor="ar-open-file" style={{...S.btnP,display:'inline-block',cursor:'pointer'}}>{parsing?'Reading file…':'Choose .xlsx file'}</label>
      {fileName&&<div style={{fontSize:12,color:T.textMuted,marginTop:8}}>{fileName}</div>}
    </div>
    {err&&<div style={{...S.err,marginBottom:12}}>{err}</div>}
    {preview&&<div style={{...S.card,padding:14,marginBottom:14}}>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Invoices</span><span style={{fontWeight:600}}>{preview.count} across {preview.custCount} customers</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>As of</span><span style={{fontWeight:600}}>{preview.asOf}</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Grand total</span><span style={{fontWeight:600}}>{'$'+fmt(preview.total)}</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'6px 0 3px',marginTop:6}}><span style={{color:T.textMuted}}>GL control balance</span><span>{preview.glBalance!==null?'$'+fmt(preview.glBalance):'-'}</span></div>
      {preview.recon!==null&&<div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Difference</span><span style={{fontWeight:700,color:tie?T.green:T.orange}}>{tie?'ties out':'$'+fmt(preview.recon)+' off'}</span></div>}
      {preview.recon!==null&&!tie&&<div style={{fontSize:12,color:T.orange,marginTop:6}}>This detail does not tie to the GL control balance as of {preview.asOf}. You can still import; the aging will carry the difference as an un-itemized residual.</div>}
    </div>}
    {preview&&<label style={{display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.textMuted,marginBottom:12}}><input type="checkbox" style={S.checkbox} checked={force} onChange={e=>setForce(e.target.checked)}/>Replace even if cash receipts were already applied to opening items</label>}
    <div style={{display:'flex',gap:10,justifyContent:'flex-end'}}>
      <button style={S.btnS} onClick={onClose} disabled={saving}>Cancel</button>
      <button style={{...S.btnP,opacity:(!preview||saving)?0.5:1}} onClick={doImport} disabled={!preview||saving}>{saving?'Importing…':(preview?'Import '+preview.count+' items':'Import')}</button>
    </div>
  </div></div>);
}

function ArAgingReport({entityId,entityName}){`;

patch('client/src/App.jsx', [
  // insert the modal component before ArAgingReport
  [`function ArAgingReport({entityId,entityName}){`, modal, 'insert-modal'],

  // state
  [`  const[showDetail,setShowDetail]=useState(true);`,
   `  const[showDetail,setShowDetail]=useState(true);const[showUpload,setShowUpload]=useState(false);`,
   'state-showupload'],

  // render modal at top of the report
  [`  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>`,
   `  return(<div>
    {showUpload&&<ArOpeningUploadModal entityId={entityId} asOf={asOf} onClose={()=>setShowUpload(false)} onImported={()=>{setShowUpload(false);run();}}/>}
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>`,
   'render-modal'],

  // header buttons
  [`      {hasAnything&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div>`,
   `      <div style={{display:'flex',gap:8,alignItems:'center'}}><button style={S.btnS} onClick={()=>setShowUpload(true)}>Upload aging detail</button>{hasAnything&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>`,
   'header-buttons'],
]);
