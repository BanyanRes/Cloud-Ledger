import { useState, useEffect, useCallback } from 'react';
import { api } from './api';
import * as XLSX from 'xlsx';

const fmt = (n) => { const v = Math.abs(n); const s = v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return n < 0 ? '(' + s + ')' : s; };
const today = () => new Date().toISOString().slice(0, 10);
const fy_start = () => new Date().getFullYear() + '-01-01';
const fmtSize = (b) => b > 1048576 ? (b/1048576).toFixed(1)+'MB' : (b/1024).toFixed(0)+'KB';
const acctLabel = (code, name) => code + ' - ' + name;
function exportToExcel(data, fileName) { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = data[0].map((_, ci) => ({ wch: Math.min(Math.max(...data.map(r => String(r[ci]||'').length), 8) + 2, 40) })); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); XLSX.writeFile(wb, fileName); }

const C = {bg:'#0c0e12',card:'#13161d',border:'#1e2230',borderLight:'#282d3e',text:'#c4cad7',textBright:'#eef0f6',textDim:'#6b7394',accent:'#4f8ff7',green:'#34d399',red:'#f87171',orange:'#fb923c',purple:'#a78bfa',teal:'#2dd4bf'};
const S = {
  app:{fontFamily:"'DM Sans',sans-serif",background:C.bg,color:C.text,minHeight:'100vh',display:'flex',flexDirection:'column',fontSize:13},
  topBar:{background:C.card,borderBottom:'1px solid '+C.border,padding:'0 20px',height:52,display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0,zIndex:10},
  body:{display:'flex',flex:1,overflow:'hidden'},sidebar:{width:210,background:C.card,borderRight:'1px solid '+C.border,padding:'12px 0',flexShrink:0,overflowY:'auto'},
  navItem:a=>({padding:'9px 18px',cursor:'pointer',fontSize:12.5,fontWeight:a?600:400,color:a?C.textBright:C.textDim,background:a?C.border:'transparent',borderLeft:a?'2px solid '+C.accent:'2px solid transparent'}),
  navSection:{padding:'14px 18px 5px',fontSize:10,fontWeight:700,color:C.textDim,textTransform:'uppercase',letterSpacing:'0.1em'},
  main:{flex:1,padding:'20px 24px',overflowY:'auto'},card:{background:C.card,border:'1px solid '+C.border,borderRadius:8,padding:20,marginBottom:16},
  h1:{fontSize:20,fontWeight:700,color:C.textBright,marginBottom:2},h2:{fontSize:14,fontWeight:600,color:C.textBright,marginBottom:14},sub:{fontSize:12,color:C.textDim,marginBottom:16},
  table:{width:'100%',borderCollapse:'collapse',fontSize:12.5},
  th:{textAlign:'left',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  thR:{textAlign:'right',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  thC:{textAlign:'center',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  td:{padding:'7px 10px',borderBottom:'1px solid '+C.border},tdR:{padding:'7px 10px',borderBottom:'1px solid '+C.border,textAlign:'right',fontVariantNumeric:'tabular-nums'},
  tdC:{padding:'7px 10px',borderBottom:'1px solid '+C.border,textAlign:'center'},tdBold:{padding:'8px 10px',borderBottom:'2px solid '+C.borderLight,color:C.textBright,fontWeight:700,fontVariantNumeric:'tabular-nums'},
  input:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'7px 10px',color:C.text,fontSize:12.5,outline:'none',width:'100%',boxSizing:'border-box'},
  inputSm:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'5px 8px',color:C.text,fontSize:12,outline:'none',boxSizing:'border-box'},
  select:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'7px 10px',color:C.text,fontSize:12.5,outline:'none',width:'100%',boxSizing:'border-box'},
  selectSm:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'5px 8px',color:C.text,fontSize:12,outline:'none',boxSizing:'border-box'},
  btnP:{background:'#1a6334',color:'#fff',border:'none',borderRadius:5,padding:'7px 16px',fontSize:12,fontWeight:600,cursor:'pointer'},
  btnS:{background:C.border,color:C.text,border:'1px solid '+C.borderLight,borderRadius:5,padding:'7px 16px',fontSize:12,fontWeight:500,cursor:'pointer'},
  btnD:{background:'#7f1d1d',color:'#fca5a5',border:'none',borderRadius:5,padding:'5px 10px',fontSize:11,fontWeight:600,cursor:'pointer'},
  btnExport:{background:C.teal+'18',color:C.teal,border:'1px solid '+C.teal+'40',borderRadius:5,padding:'6px 14px',fontSize:11,fontWeight:600,cursor:'pointer'},
  row:{display:'flex',gap:10,marginBottom:10,flexWrap:'wrap'},col:{flex:1,minWidth:120},
  label:{fontSize:11,color:C.textDim,marginBottom:3,display:'block',fontWeight:500},
  err:{color:C.red,fontSize:11,marginTop:4},success:{color:C.green,fontSize:11,marginTop:4},
  tag:t=>{const c={Asset:C.accent,Liability:C.orange,Equity:C.green,Revenue:C.purple,Expense:C.red};return{display:'inline-block',padding:'1px 7px',borderRadius:8,fontSize:10,fontWeight:600,color:c[t]||C.textDim,background:(c[t]||C.textDim)+'18'}},
  badge:{background:C.accent+'20',color:C.accent,padding:'2px 8px',borderRadius:8,fontSize:10,fontWeight:600},
  reportHeader:{borderBottom:'3px double '+C.borderLight,paddingBottom:10,marginBottom:14,textAlign:'center'},
  sectionHeader:{background:C.border,padding:'7px 10px',fontWeight:700,color:C.textBright,fontSize:12},
  indentTd:{padding:'6px 10px 6px 24px',borderBottom:'1px solid '+C.border+'10',fontSize:12.5},
  subtotalRow:{borderTop:'1px solid '+C.borderLight},grandTotalRow:{borderTop:'3px double '+C.borderLight,background:C.border},
  link:{color:C.accent,cursor:'pointer',fontSize:12,background:'none',border:'none',padding:0,textDecoration:'underline'},
  checkbox:{width:16,height:16,cursor:'pointer',accentColor:C.green},
  logoIcon:{width:30,height:30,background:'linear-gradient(135deg,'+C.accent+','+C.green+')',borderRadius:7,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:13,color:C.bg},
  modal:{position:'fixed',inset:0,background:'rgba(0,0,0,0.6)',zIndex:200,display:'flex',alignItems:'center',justifyContent:'center'},
  modalBox:{background:C.card,border:'1px solid '+C.border,borderRadius:10,width:'92%',maxWidth:920,maxHeight:'90vh',overflowY:'auto',padding:24,position:'relative'},
  modalClose:{position:'absolute',top:12,right:16,cursor:'pointer',color:C.textDim,fontSize:20,background:'none',border:'none'},
  filterBar:{display:'flex',alignItems:'flex-end',gap:12,flexWrap:'wrap',marginBottom:16},
  fileChip:{display:'inline-flex',alignItems:'center',gap:4,background:C.border,padding:'3px 8px',borderRadius:4,fontSize:11,color:C.text,marginRight:4,marginBottom:4},
  attachLink:{display:'inline-flex',alignItems:'center',gap:3,padding:'2px 8px',borderRadius:4,fontSize:10,color:C.accent,background:C.accent+'12',textDecoration:'none',marginRight:4,marginBottom:2},
};

const BLANK_JE = () => ({date:today(),memo:'',lines:[{account_code:'',debit:'',credit:''},{account_code:'',debit:'',credit:''}]});

// ─── Auth Screen ───
function AuthScreen({onLogin}){const[mode,setMode]=useState('login');const[email,setEmail]=useState('');const[pw,setPw]=useState('');const[name,setName]=useState('');const[confirmPw,setConfirmPw]=useState('');const[role,setRole]=useState('Accountant');
  const[err,setErr]=useState('');const[success,setSuccess]=useState('');const[loading,setLoading]=useState(false);const[tempPw,setTempPw]=useState('');
  const doLogin=async()=>{setLoading(true);setErr('');try{const d=await api.login(email.trim().toLowerCase(),pw);api.setToken(d.token);onLogin(d.user);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doSignup=async()=>{if(!name.trim()){setErr('Name required');return;}if(pw.length<3){setErr('Min 3 chars');return;}if(pw!==confirmPw){setErr("Don't match");return;}setLoading(true);setErr('');try{await api.signup(name.trim(),email.trim().toLowerCase(),pw,role);setSuccess('Created!');setTimeout(()=>{setMode('login');setSuccess('');},1200);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doForgot=async()=>{if(!email.trim()){setErr('Enter email');return;}setLoading(true);setErr('');try{const r=await api.forgotPassword(email.trim().toLowerCase());setTempPw(r.temp_password);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const hk=e=>{if(e.key==='Enter'){mode==='login'?doLogin():mode==='signup'?doSignup():doForgot();}};
  return(<div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100vh',background:C.bg}}><div style={{...S.card,width:400,textAlign:'center',padding:36}}>
    <div style={{...S.logoIcon,width:44,height:44,fontSize:18,margin:'0 auto 12px'}}>GL</div><div style={{fontSize:20,fontWeight:700,color:C.textBright,marginBottom:2}}>CloudLedger</div><div style={{fontSize:12,color:C.textDim,marginBottom:24}}>Multi-Entity Cloud Accounting</div>
    {mode==='forgot'?(<><div style={{fontSize:14,fontWeight:600,color:C.textBright,marginBottom:16}}>Reset Password</div>
      <div style={{marginBottom:10}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');setTempPw('');}} onKeyDown={hk}/></div>
      {err&&<div style={S.err}>{err}</div>}{tempPw&&<div style={{background:C.green+'15',border:'1px solid '+C.green+'40',borderRadius:8,padding:16,margin:'10px 0'}}><div style={{fontSize:12,color:C.green}}>Temporary password:</div><div style={{fontSize:18,fontWeight:700,color:C.textBright,fontFamily:'monospace'}}>{tempPw}</div><div style={{fontSize:11,color:C.textDim,marginTop:8}}>Sign in with this, then change in Settings.</div></div>}
      <button style={{...S.btnP,width:'100%',padding:9,marginTop:6}} onClick={doForgot} disabled={loading}>Reset Password</button>
      <div style={{marginTop:16}}><button style={S.link} onClick={()=>{setMode('login');setErr('');setTempPw('');}}>Back to Sign In</button></div>
    </>):(<><div style={{display:'flex',marginBottom:20,borderRadius:6,overflow:'hidden',border:'1px solid '+C.border}}>
      <div onClick={()=>{setMode('login');setErr('');}} style={{flex:1,padding:'8px 0',cursor:'pointer',fontSize:12,fontWeight:600,textAlign:'center',background:mode==='login'?C.accent+'20':'transparent',color:mode==='login'?C.accent:C.textDim}}>Sign In</div>
      <div onClick={()=>{setMode('signup');setErr('');}} style={{flex:1,padding:'8px 0',cursor:'pointer',fontSize:12,fontWeight:600,textAlign:'center',background:mode==='signup'?C.green+'20':'transparent',color:mode==='signup'?C.green:C.textDim}}>Create Account</div></div>
    {mode==='login'?(<><div style={{marginBottom:10}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
      <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>{err&&<div style={S.err}>{err}</div>}
      <button style={{...S.btnP,width:'100%',padding:9,marginTop:6}} onClick={doLogin} disabled={loading}>Sign In</button>
      <div style={{display:'flex',justifyContent:'space-between',marginTop:14}}><button style={S.link} onClick={()=>{setMode('forgot');setErr('');}}>Forgot password?</button><button style={S.link} onClick={()=>setMode('signup')}>Create account</button></div>
    </>):(<><div style={{marginBottom:10}}><input style={S.input} placeholder="Full Name" value={name} onChange={e=>{setName(e.target.value);setErr('');}} onKeyDown={hk}/></div>
      <div style={{marginBottom:10}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
      <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
      <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Confirm" value={confirmPw} onChange={e=>{setConfirmPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
      <div style={{marginBottom:10,textAlign:'left'}}><label style={S.label}>Role</label><select style={S.select} value={role} onChange={e=>setRole(e.target.value)}><option value="Admin">Admin</option><option value="Accountant">Accountant</option><option value="Viewer">Viewer</option></select></div>
      {err&&<div style={S.err}>{err}</div>}{success&&<div style={S.success}>{success}</div>}
      <button style={{...S.btnP,width:'100%',padding:9,marginTop:6,background:'#065f46'}} onClick={doSignup} disabled={loading}>Create Account</button>
      <div style={{marginTop:14}}><button style={S.link} onClick={()=>setMode('login')}>Back to Sign In</button></div></>)}</>)}
  </div></div>);}

// ─── Small Modals ───
function ChangePasswordModal({onClose}){const[cur,setCur]=useState('');const[nw,setNw]=useState('');const[cf,setCf]=useState('');const[err,setErr]=useState('');const[ok,setOk]=useState(false);
  return(<div style={S.modal} onClick={onClose}><div style={{...S.modalBox,maxWidth:380,textAlign:'center'}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>x</button><div style={{fontSize:16,fontWeight:700,color:C.textBright,marginBottom:16}}>Change Password</div>
    <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Current" value={cur} onChange={e=>{setCur(e.target.value);setErr('');}}/></div>
    <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="New" value={nw} onChange={e=>{setNw(e.target.value);setErr('');}}/></div>
    <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Confirm" value={cf} onChange={e=>{setCf(e.target.value);setErr('');}}/></div>
    {err&&<div style={S.err}>{err}</div>}{ok&&<div style={S.success}>Done!</div>}
    <button style={{...S.btnP,width:'100%',padding:9,marginTop:6}} onClick={async()=>{if(nw.length<3){setErr('Min 3');return;}if(nw!==cf){setErr("Don't match");return;}try{await api.changePassword(cur,nw);setOk(true);setTimeout(onClose,1500);}catch(e){setErr(e.message);}}}>Update</button>
  </div></div>);}

function QuickAddAccountModal({entityId,onClose,onCreated}){const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  return(<div style={S.modal} onClick={onClose}><div style={{...S.modalBox,maxWidth:620}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>x</button><div style={{fontSize:16,fontWeight:700,color:C.textBright,marginBottom:16}}>Quick Add Account</div>
    <div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Subtype</label><input style={S.input} value={form.subtype} onChange={e=>setForm(f=>({...f,subtype:e.target.value}))}/></div></div>
    <div style={{marginBottom:10}}><label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank/cash account</label></div>
    {err&&<div style={S.err}>{err}</div>}<div style={{display:'flex',gap:10,marginTop:8}}><button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{const a=await api.createAccount(entityId,form);onCreated(a);onClose();}catch(e){setErr(e.message);}}}>Add</button><button style={S.btnS} onClick={onClose}>Cancel</button></div>
  </div></div>);}

// ─── JE Modal (receives form state from parent so it persists!) ───
function JournalEntryModal({entityId,user,onClose,onPosted,jeForm,setJeForm,pendingFiles,setPendingFiles}){
  const[accounts,setAccounts]=useState([]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[posting,setPosting]=useState(false);const[posted,setPosted]=useState('');
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts);},[entityId]);
  const addLine=()=>setJeForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:''}]}));
  const removeLine=i=>setJeForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setJeForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const tDr=jeForm.lines.reduce((s,l)=>s+(parseFloat(l.debit)||0),0);const tCr=jeForm.lines.reduce((s,l)=>s+(parseFloat(l.credit)||0),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;

  const post=async()=>{if(!jeForm.date||!jeForm.memo.trim()){setErr('Date & memo required');return;}if(jeForm.lines.some(l=>!l.account_code)){setErr('All lines need account');return;}if(!bal){setErr('Must balance');return;}
    setPosting(true);setErr('');try{const r=await api.createEntry(entityId,{date:jeForm.date,memo:jeForm.memo.trim(),lines:jeForm.lines.map(l=>({account_code:l.account_code,debit:parseFloat(l.debit)||0,credit:parseFloat(l.credit)||0}))});
      if(pendingFiles.length>0)await api.uploadAttachments(entityId,r.id,pendingFiles);
      setJeForm(BLANK_JE());setPendingFiles([]);setPosted('JE-'+String(r.entry_num).padStart(4,'0')+' posted!');setTimeout(()=>setPosted(''),3000);if(onPosted)onPosted();}
    catch(e){setErr(e.message);}finally{setPosting(false);}};

  const hasDraft = jeForm.memo || jeForm.lines.some(l=>l.account_code||l.debit||l.credit) || pendingFiles.length>0;

  return(<div style={S.modal} onClick={onClose}><div style={{...S.modalBox,maxWidth:960}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>x</button>
    <div style={{fontSize:16,fontWeight:700,color:C.textBright,marginBottom:16}}>New Journal Entry {hasDraft&&<span style={{fontSize:11,color:C.orange,marginLeft:8}}>Draft in progress</span>}</div>
    <div style={S.row}><div style={{...S.col,maxWidth:160}}><label style={S.label}>Date</label><input style={S.input} type="date" value={jeForm.date} onChange={e=>setJeForm(f=>({...f,date:e.target.value}))}/></div>
      <div style={{...S.col,flex:4}}><label style={S.label}>Memo / Description</label><input style={S.input} placeholder="Enter description" value={jeForm.memo} onChange={e=>setJeForm(f=>({...f,memo:e.target.value}))}/></div></div>
    <table style={{...S.table,marginBottom:10}}><thead><tr><th style={S.th}>Account</th><th style={{...S.thR,width:130}}>Debit</th><th style={{...S.thR,width:130}}>Credit</th><th style={{...S.th,width:30}}></th></tr></thead>
      <tbody>{jeForm.lines.map((l,i)=><tr key={i}><td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}>
        <select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select account...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.debit} onChange={e=>updateLine(i,'debit',e.target.value)}/></td>
        <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.credit} onChange={e=>updateLine(i,'credit',e.target.value)}/></td>
        <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border,textAlign:'center'}}>{jeForm.lines.length>2&&<span style={{cursor:'pointer',color:C.red}} onClick={()=>removeLine(i)}>x</span>}</td></tr>)}
      <tr style={{background:C.border}}><td style={{...S.tdBold,textAlign:'right'}}>Total</td><td style={{...S.tdBold,textAlign:'right'}}>{'$'+fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right'}}>{'$'+fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table>
    {/* Attachments */}
    <div style={{marginBottom:12}}><label style={{...S.label,marginBottom:6}}>Attachments</label>
      <div>{pendingFiles.map((f,i)=><span key={i} style={S.fileChip}>{f.name} ({fmtSize(f.size)}) <span style={{cursor:'pointer',color:C.red,marginLeft:2}} onClick={()=>setPendingFiles(p=>p.filter((_,j)=>j!==i))}>x</span></span>)}</div>
      <input type="file" multiple accept=".pdf,.xlsx,.xls,.csv,.png,.jpg,.jpeg,.eml,.msg,.doc,.docx" style={{display:'none'}} id="je-attach" onChange={e=>{setPendingFiles(p=>[...p,...Array.from(e.target.files)]);e.target.value='';}}/>
      <label htmlFor="je-attach" style={{...S.btnS,display:'inline-block',padding:'5px 12px',fontSize:11,cursor:'pointer',marginTop:4}}>+ Attach Files</label>
      {pendingFiles.length>0&&<span style={{fontSize:11,color:C.green,marginLeft:8}}>{pendingFiles.length} file(s) ready</span>}</div>
    <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
      <button style={S.btnS} onClick={addLine}>+ Line</button>
      <button style={{...S.btnS,color:C.teal,borderColor:C.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New Account</button>
      <div style={{flex:1}}/>
      {!bal&&tDr>0&&<span style={{color:C.orange,fontSize:11}}>Diff: ${fmt(tDr-tCr)}</span>}
      {bal&&<span style={{color:C.green,fontSize:11}}>Balanced</span>}
      {err&&<span style={S.err}>{err}</span>}{posted&&<span style={S.success}>{posted}</span>}
      <button style={{...S.btnP,opacity:posting?.6:1}} onClick={post} disabled={posting}>{posting?'Posting...':'Post Entry'}</button></div>
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)))}/>}
  </div></div>);}

// ─── Entity Picker ───
function EntityPicker({entities,activeId,onSelect,onManage}){const[open,setOpen]=useState(false);const[search,setSearch]=useState('');const active=entities.find(e=>e.id===activeId);
  const filtered=entities.filter(e=>e.name.toLowerCase().includes(search.toLowerCase())||e.code.toLowerCase().includes(search.toLowerCase()));
  return(<div style={{position:'relative'}}><div style={{display:'flex',alignItems:'center',gap:6,cursor:'pointer',padding:'5px 12px',borderRadius:6,background:C.border,border:'1px solid '+C.borderLight}} onClick={()=>setOpen(!open)}>
    <span style={{fontWeight:600,color:C.textBright,fontSize:12}}>{active?.code||'-'}</span><span style={{color:C.textDim,fontSize:12,maxWidth:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{active?.name||'Select'}</span></div>
    {open&&<><div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>{setOpen(false);setSearch('');}}/>
      <div style={{position:'absolute',top:'100%',left:0,background:C.card,border:'1px solid '+C.border,borderRadius:8,maxHeight:350,overflowY:'auto',zIndex:100,boxShadow:'0 12px 40px rgba(0,0,0,0.5)',width:320}}>
        <div style={{position:'sticky',top:0,padding:10,background:C.card,borderBottom:'1px solid '+C.border}}><input style={S.input} placeholder={'Search '+entities.length+'...'} value={search} onChange={e=>setSearch(e.target.value)} autoFocus/></div>
        {filtered.map(e=><div key={e.id} style={{padding:'8px 14px',cursor:'pointer',background:e.id===activeId?C.accent+'12':'transparent',borderLeft:e.id===activeId?'2px solid '+C.accent:'2px solid transparent'}} onClick={()=>{onSelect(e.id);setOpen(false);setSearch('');}}>
          <span style={{fontWeight:600,color:C.textBright,fontSize:12}}>{e.code}</span><span style={{color:C.text,fontSize:12,marginLeft:8}}>{e.name}</span></div>)}
        <div style={{borderTop:'1px solid '+C.border,padding:10}}><button style={{...S.btnS,width:'100%',fontSize:11}} onClick={()=>{onManage();setOpen(false);}}>Manage Entities</button></div></div></>}</div>);}

// ═══ Main App - JE form state lives HERE so it persists across navigation ═══
export default function App(){
  const[user,setUser]=useState(null);const[entities,setEntities]=useState([]);const[activeEntity,setActiveEntity]=useState(null);
  const[page,setPage]=useState('dashboard');const[loading,setLoading]=useState(true);
  const[showJE,setShowJE]=useState(false);const[showChangePw,setShowChangePw]=useState(false);const[rk,setRk]=useState(0);
  // *** JE form state lifted to App so it survives page navigation ***
  const[jeForm,setJeForm]=useState(BLANK_JE());const[pendingFiles,setPendingFiles]=useState([]);
  const hasDraft=jeForm.memo||jeForm.lines.some(l=>l.account_code||l.debit||l.credit)||pendingFiles.length>0;

  useEffect(()=>{const t=api.getToken();if(t){api.me().then(u=>{if(u)setUser(u);}).catch(()=>api.clearToken()).finally(()=>setLoading(false));}else setLoading(false);},[]);
  useEffect(()=>{if(user)api.getEntities().then(e=>{setEntities(e);if(e.length>0&&!activeEntity)setActiveEntity(e[0].id);});},[user]);
  const refreshEntities=useCallback(async()=>{const e=await api.getEntities();setEntities(e);return e;},[]);
  const canAccess=s=>{if(!user)return false;if(user.role==='Admin')return true;return({Accountant:['entries','reports','coa','bankrec'],Viewer:['reports']}[user.role]||[]).includes(s);};
  if(loading)return<div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:C.textDim}}>Loading...</div></div>;
  if(!user)return<AuthScreen onLogin={setUser}/>;
  const navItems=[{id:'dashboard',label:'Dashboard',section:'reports'},{id:'d1',divider:1,label:'TRANSACTIONS'},{id:'journal',label:'Journal Entries',section:'entries'},{id:'d2',divider:1,label:'ACCOUNTS'},{id:'coa',label:'Chart of Accounts',section:'coa'},{id:'ledger',label:'General Ledger',section:'reports'},{id:'d2b',divider:1,label:'BANKING'},{id:'banktxn',label:'Bank Transactions',section:'bankrec'},{id:'bankrec',label:'Bank Reconciliation',section:'bankrec'},{id:'d3',divider:1,label:'REPORTS'},{id:'trial',label:'Trial Balance',section:'reports'},{id:'bs',label:'Balance Sheet',section:'reports'},{id:'is',label:'Income Statement',section:'reports'},{id:'d4',divider:1,label:'ADMINISTRATION'},{id:'entities',label:'Entities ('+entities.length+')',section:'all'},{id:'users',label:'Users',section:'all'}];
  return(<div style={S.app}>
    <div style={S.topBar}><div style={{display:'flex',alignItems:'center',gap:16}}>
      <div style={{display:'flex',alignItems:'center',gap:10}}><div style={S.logoIcon}>GL</div><div style={{fontSize:16,fontWeight:700,color:C.textBright}}>CloudLedger</div></div>
      <div style={{width:1,height:24,background:C.border}}/><EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')}/></div>
      <div style={{display:'flex',alignItems:'center',gap:8}}>
        {canAccess('entries')&&activeEntity&&<button style={{...S.btnP,position:'relative'}} onClick={()=>setShowJE(true)}>+ Journal Entry{hasDraft&&<span style={{position:'absolute',top:-4,right:-4,width:8,height:8,borderRadius:4,background:C.orange}}/>}</button>}
        <span style={{fontSize:12}}>{user.name}</span><span style={S.badge}>{user.role}</span>
        <button style={{...S.btnS,padding:'4px 10px',fontSize:10}} onClick={()=>setShowChangePw(true)}>Settings</button>
        <button style={{...S.btnS,padding:'4px 10px',fontSize:11}} onClick={()=>{api.clearToken();setUser(null);}}>Sign Out</button></div></div>
    <div style={S.body}><div style={S.sidebar}>{navItems.map(n=>n.divider?<div key={n.id} style={S.navSection}>{n.label}</div>
      :(n.section==='all'?user.role==='Admin':canAccess(n.section))?<div key={n.id} style={S.navItem(page===n.id)} onClick={()=>setPage(n.id)}>{n.label}</div>:null)}</div>
      <div style={S.main}>
        {page==='dashboard'&&<Dashboard entityId={activeEntity} key={rk}/>}
        {page==='journal'&&activeEntity&&<JournalList entityId={activeEntity} key={activeEntity+'-'+rk} onNewEntry={()=>setShowJE(true)}/>}
        {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} canEdit={canAccess('coa')}/>}
        {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity} key={activeEntity+'-'+rk}/>}
        {page==='banktxn'&&activeEntity&&<BankTransactions entityId={activeEntity}/>}
        {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user}/>}
        {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity} key={activeEntity+'-'+rk}/>}
        {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity}/>}
        {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity}/>}
        {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='users'&&<UserManagement currentUser={user}/>}
      </div></div>
    {showJE&&activeEntity&&<JournalEntryModal entityId={activeEntity} user={user} onClose={()=>setShowJE(false)} onPosted={()=>setRk(k=>k+1)} jeForm={jeForm} setJeForm={setJeForm} pendingFiles={pendingFiles} setPendingFiles={setPendingFiles}/>}
    {showChangePw&&<ChangePasswordModal onClose={()=>setShowChangePw(false)}/>}
  </div>);}

// ═══ Dashboard ═══
function Dashboard({entityId}){const[summary,setSummary]=useState([]);useEffect(()=>{api.getSummary().then(setSummary);},[]);const curr=summary.find(e=>e.id===entityId);
  return(<div><div style={S.h1}>Dashboard</div><div style={S.sub}>{summary.length} entities</div>
    {curr&&<div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:12,marginBottom:20}}>
      {[{l:'Total Assets',v:curr.assets,c:C.accent},{l:'Total Liabilities',v:curr.liabilities,c:C.orange},{l:'Revenue',v:curr.revenue,c:C.purple},{l:'Expenses',v:curr.expenses,c:C.red},{l:'Net Income',v:curr.net_income,c:curr.net_income>=0?C.green:C.red},{l:'Entries',v:curr.entry_count,c:C.textDim,raw:1}].map(s=>
        <div key={s.l} style={{...S.card,textAlign:'center',padding:16}}><div style={{fontSize:22,fontWeight:700,color:s.c,fontVariantNumeric:'tabular-nums'}}>{s.raw?s.v:'$'+fmt(s.v)}</div><div style={{fontSize:11,color:C.textDim,marginTop:3}}>{s.l}</div></div>)}</div>}
    <div style={S.card}><div style={S.h2}>All Entities</div><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={S.thR}>Assets</th><th style={S.thR}>Liabilities</th><th style={S.thR}>Net Income</th><th style={S.thR}>JEs</th></tr></thead>
      <tbody>{summary.sort((a,b)=>a.code.localeCompare(b.code)).map(e=><tr key={e.id} style={e.id===entityId?{background:C.accent+'08'}:{}}><td style={{...S.td,fontWeight:600,color:C.accent}}>{e.code}</td><td style={S.td}>{e.name}</td><td style={S.tdR}>{fmt(e.assets)}</td><td style={S.tdR}>{fmt(e.liabilities)}</td><td style={{...S.tdR,color:e.net_income>=0?C.green:C.red,fontWeight:600}}>{fmt(e.net_income)}</td><td style={S.tdR}>{e.entry_count}</td></tr>)}</tbody></table></div></div>);}

// ═══ Journal List ═══
function JournalList({entityId,onNewEntry}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[from,setFrom]=useState('');const[to,setTo]=useState('');
  const load=useCallback(async()=>{const[e,a]=await Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]);setEntries(e);setAccounts(a);},[entityId,from,to]);
  useEffect(()=>{load();},[load]);const del=async id=>{await api.deleteEntry(entityId,id);load();};const acctName=code=>accounts.find(a=>a.code===code)?.name||'?';
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div><div style={S.h1}>Journal Entries</div><div style={{fontSize:12,color:C.textDim}}>{entries.length} entries</div></div><button style={S.btnP} onClick={onNewEntry}>+ New Entry</button></div>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      {(from||to)&&<button style={{...S.btnS,padding:'5px 10px',fontSize:10,marginTop:14}} onClick={()=>{setFrom('');setTo('');}}>Clear</button>}</div>
    <div style={S.card}>{entries.length===0?<div style={{textAlign:'center',padding:40,color:C.textDim,fontSize:12}}>No entries</div>:
      entries.map(e=><div key={e.id} style={{borderBottom:'2px solid '+C.borderLight,paddingBottom:12,marginBottom:12}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
          <div style={{display:'flex',alignItems:'center',gap:12}}><span style={{fontWeight:700,color:C.accent,fontSize:13}}>JE-{String(e.entry_num).padStart(4,'0')}</span>
            <span style={{color:C.textDim,fontSize:12}}>{e.date}</span><span style={{color:C.text,fontSize:12}}>{e.memo}</span>
            {e.attachments&&e.attachments.length>0&&<span style={{fontSize:10,color:C.teal}}>({e.attachments.length} attachment{e.attachments.length>1?'s':''})</span>}</div>
          <div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontSize:11,color:C.textDim}}>{e.created_by}</span><span style={{cursor:'pointer',color:C.red,fontSize:12}} onClick={()=>del(e.id)}>Delete</span></div></div>
        <table style={S.table}><thead><tr><th style={S.th}>Account</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
          <tbody>{e.lines.map((l,i)=><tr key={i}><td style={{...S.td,paddingLeft:l.credit>0&&l.debit===0?24:10}}>{acctLabel(l.account_code,acctName(l.account_code))}</td>
            <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td><td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td></tr>)}</tbody></table>
        {e.attachments&&e.attachments.length>0&&<div style={{marginTop:6}}>{e.attachments.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>)}</div>}
      </div>)}</div></div>);}

// ═══ Chart of Accounts ═══
function ChartOfAccounts({entityId,canEdit}){const[accounts,setAccounts]=useState([]);const[showAdd,setShowAdd]=useState(false);
  const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  const load=useCallback(async()=>{setAccounts(await api.getAccounts(entityId));},[entityId]);useEffect(()=>{load();},[load]);
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div><div style={S.h1}>Chart of Accounts</div><div style={{fontSize:12,color:C.textDim}}>{accounts.length} accounts</div></div>
    {canEdit&&<button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add'}</button>}</div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}>
      <div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Subtype</label><input style={S.input} value={form.subtype} onChange={e=>setForm(f=>({...f,subtype:e.target.value}))}/></div></div>
      <div style={{marginBottom:10}}><label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank/cash account</label></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createAccount(entityId,form);setForm({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});setShowAdd(false);setErr('');load();}catch(e){setErr(e.message);}}}>Add</button></div>}
    <div style={S.card}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Name</th><th style={S.th}>Type</th><th style={S.th}>Subtype</th><th style={S.thC}>Bank</th>{canEdit&&<th style={{...S.th,width:30}}></th>}</tr></thead>
      <tbody>{accounts.map(a=><tr key={a.code}><td style={{...S.td,fontWeight:600}}>{a.code}</td><td style={S.td}>{a.name}</td><td style={S.td}><span style={S.tag(a.type)}>{a.type}</span></td><td style={{...S.td,color:C.textDim}}>{a.subtype}</td>
        <td style={S.tdC}>{a.bank_acct?'Yes':''}</td>{canEdit&&<td style={S.td}><span style={{cursor:'pointer',color:C.red}} onClick={async()=>{try{await api.deleteAccount(entityId,a.code);load();}catch(e){alert(e.message);}}}>x</span></td>}</tr>)}</tbody></table></div></div>);}

// ═══ General Ledger (with attachment links!) ═══
function GeneralLedger({entityId}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[filter,setFilter]=useState('');const[from,setFrom]=useState(fy_start());const[to,setTo]=useState(today());
  useEffect(()=>{Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]).then(([e,a])=>{setEntries(e);setAccounts(a);});},[entityId,from,to]);
  const filtered=accounts.filter(a=>!filter||a.code===filter).sort((a,b)=>a.code.localeCompare(b.code));
  // Build a map of entry attachments for quick lookup
  const entryAttachments = {};
  entries.forEach(e => { if(e.attachments && e.attachments.length > 0) entryAttachments[e.id] = e.attachments; });
  const doExport=()=>{const rows=[['General Ledger','','','','',''],['Period: '+(from||'Begin')+' to '+(to||today())],[]];filtered.forEach(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({date:e.date,je:'JE-'+String(e.entry_num).padStart(4,'0'),memo:e.memo,debit:l.debit,credit:l.credit});});});if(txns.length===0&&!filter)return;rows.push([acctLabel(acct.code,acct.name)]);rows.push(['Date','JE','Memo','Debit','Credit','Balance']);let run=0;const isDr=acct.type==='Asset'||acct.type==='Expense';txns.sort((a,b)=>a.date.localeCompare(b.date)).forEach(t=>{run+=isDr?(t.debit-t.credit):(t.credit-t.debit);rows.push([t.date,t.je,t.memo,t.debit||'',t.credit||'',run]);});rows.push([]);});exportToExcel(rows,'GL.xlsx');};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}><div style={S.h1}>General Ledger</div><button style={S.btnExport} onClick={doExport}>Export Excel</button></div><div style={S.sub}/>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <div style={{maxWidth:260}}><label style={S.label}>Account</label><select style={{...S.inputSm,width:'100%'}} value={filter} onChange={e=>setFilter(e.target.value)}><option value="">All</option>{accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div></div>
    {filtered.map(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({...l,date:e.date,memo:e.memo,jeNum:e.entry_num,jeId:e.id});});});
      if(txns.length===0&&!filter)return null;txns.sort((a,b)=>a.date.localeCompare(b.date));let run=0;const dr=acct.type==='Asset'||acct.type==='Expense';
      return(<div key={acct.code} style={S.card}><div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10}}><span style={{fontWeight:700,color:C.textBright}}>{acct.code}</span><span>{acct.name}</span><span style={S.tag(acct.type)}>{acct.type}</span></div>
        {txns.length===0?<div style={{color:C.textDim,fontSize:12}}>No transactions</div>:
        <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th><th style={{...S.th,width:90}}>Docs</th></tr></thead>
          <tbody>{txns.map((t,i)=>{run+=dr?(t.debit-t.credit):(t.credit-t.debit);const atts=entryAttachments[t.jeId];return<tr key={i}><td style={S.td}>{t.date}</td><td style={S.td}><span style={{color:C.accent}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td><td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600}}>{fmt(run)}</td>
            <td style={S.td}>{atts?atts.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>):''}</td></tr>;})}</tbody></table>}</div>);})}</div>);}

// ═══ Bank Transactions (with Quick Add Account) ═══
function BankTransactions({entityId}){const[accounts,setAccounts]=useState([]);const[bankAccts,setBankAccts]=useState([]);const[txns,setTxns]=useState([]);
  const[selAcct,setSelAcct]=useState('');const[statusFilter,setStatusFilter]=useState('pending');const[err,setErr]=useState('');const[msg,setMsg]=useState('');const[showAddAcct,setShowAddAcct]=useState(false);
  const load=useCallback(async()=>{const a=await api.getAccounts(entityId);setAccounts(a);setBankAccts(a.filter(x=>x.bank_acct||(['cash','bank','checking','savings'].some(w=>x.name.toLowerCase().includes(w))&&x.type==='Asset')));
    if(selAcct)setTxns(await api.getBankTransactions(entityId,selAcct,statusFilter||undefined));},[entityId,selAcct,statusFilter]);
  useEffect(()=>{load();},[load]);
  const handleUpload=async e=>{const file=e.target.files[0];if(!file||!selAcct)return;e.target.value='';setErr('');setMsg('');try{const r=await api.uploadBankTransactions(entityId,selAcct,file);setMsg(r.count+' transactions imported');load();}catch(ex){setErr(ex.message);}};
  const codeTransaction=async(id,account_code,memo)=>{await api.codeBankTransaction(entityId,id,account_code,memo);load();};
  const postCoded=async()=>{const ids=txns.filter(t=>t.status==='coded').map(t=>t.id);if(!ids.length){setErr('No coded transactions');return;}try{const r=await api.postBankTransactions(entityId,ids);setMsg(r.posted+' JEs created');load();}catch(ex){setErr(ex.message);}};
  const handleAcctCreated=a=>{setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));setBankAccts(p=>{if(a.bank_acct)return[...p,a].sort((x,y)=>x.code.localeCompare(y.code));return p;});};

  return(<div><div style={S.h1}>Bank Transactions</div><div style={S.sub}>Upload bank statements, code transactions, and post to GL</div>
    <div style={S.card}><div style={S.h2}>Select Bank Account</div><div style={S.row}>
      <div style={{...S.col,flex:2}}><select style={S.select} value={selAcct} onChange={e=>{setSelAcct(e.target.value);setTxns([]);}}><option value="">Select bank account...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
      <div style={S.col}><select style={S.select} value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}><option value="">All</option><option value="pending">Pending</option><option value="coded">Coded</option><option value="posted">Posted</option></select></div>
      {selAcct&&<div><input type="file" accept=".csv,.xlsx,.xls" id="bank-upload" style={{display:'none'}} onChange={handleUpload}/><label htmlFor="bank-upload" style={{...S.btnP,display:'inline-block',cursor:'pointer',marginTop:14}}>Upload CSV/Excel</label></div>}
    </div>{err&&<div style={S.err}>{err}</div>}{msg&&<div style={S.success}>{msg}</div>}</div>

    {selAcct&&txns.length>0&&<div style={S.card}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
        <div style={S.h2}>{txns.length} Transactions</div>
        <div style={{display:'flex',gap:8}}>
          <button style={{...S.btnS,color:C.teal,borderColor:C.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New Account</button>
          {txns.some(t=>t.status==='coded')&&<button style={S.btnP} onClick={postCoded}>Post {txns.filter(t=>t.status==='coded').length} Coded to GL</button>}</div></div>
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Description</th><th style={S.thR}>Amount</th><th style={S.th}>GL Account</th><th style={S.th}>Memo</th><th style={S.th}>Status</th><th style={{...S.th,width:30}}></th></tr></thead>
        <tbody>{txns.map(t=><tr key={t.id} style={t.status==='posted'?{opacity:0.5}:{}}>
          <td style={S.td}>{t.date}</td><td style={{...S.td,maxWidth:200,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{t.description}</td>
          <td style={{...S.tdR,color:t.amount>=0?C.green:C.red,fontWeight:600}}>{fmt(t.amount)}</td>
          <td style={{...S.td,padding:'3px 4px'}}>{t.status==='posted'?<span style={{fontSize:11,color:C.textDim}}>{t.account_code}</span>:
            <select style={S.selectSm} value={t.account_code||''} onChange={e=>codeTransaction(t.id,e.target.value,t.memo)}><option value="">- Code -</option>{accounts.filter(a=>a.code!==selAcct).sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select>}</td>
          <td style={{...S.td,padding:'3px 4px'}}>{t.status==='posted'?<span style={{fontSize:11}}>{t.memo}</span>:
            <input style={S.inputSm} placeholder="Memo" value={t.memo||''} onChange={e=>{const v=e.target.value;setTxns(prev=>prev.map(x=>x.id===t.id?{...x,memo:v}:x));}} onBlur={()=>codeTransaction(t.id,t.account_code,t.memo)}/>}</td>
          <td style={S.td}><span style={{fontSize:10,fontWeight:600,padding:'2px 6px',borderRadius:4,background:t.status==='posted'?C.green+'20':t.status==='coded'?C.accent+'20':C.orange+'20',color:t.status==='posted'?C.green:t.status==='coded'?C.accent:C.orange}}>{t.status}</span></td>
          <td style={S.td}>{t.status!=='posted'&&<span style={{cursor:'pointer',color:C.red,fontSize:11}} onClick={async()=>{await api.deleteBankTransaction(entityId,t.id);load();}}>x</span>}</td>
        </tr>)}</tbody></table></div>}
    {selAcct&&txns.length===0&&<div style={{...S.card,textAlign:'center',padding:40,color:C.textDim}}>No transactions. Upload a bank statement to get started.</div>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={handleAcctCreated}/>}
  </div>);}

// ═══ Reports ═══
function TrialBalance({entityId}){const[balances,setBalances]=useState([]);const[asOf,setAsOf]=useState(today());const fyStart=asOf.slice(0,4)+'-01-01';
  useEffect(()=>{api.getBalances(entityId,{as_of:asOf,close_pl_before:fyStart}).then(setBalances);},[entityId,asOf,fyStart]);
  let tDr=0,tCr=0;const rows=balances.filter(b=>Math.abs(b.balance)>0.005).map(b=>{const isDr=b.type==='Asset'||b.type==='Expense';const dr=(isDr&&b.balance>0)||(!isDr&&b.balance<0)?Math.abs(b.balance):0;const cr=(isDr&&b.balance<0)||(!isDr&&b.balance>0)?Math.abs(b.balance):0;tDr+=dr;tCr+=cr;return{...b,dr,cr};});
  const doExport=()=>{const d=[['Trial Balance'],['As of '+asOf],[],['Code','Account','Type','Debit','Credit']];rows.forEach(r=>d.push([r.code,r.name,r.type,r.dr||'',r.cr||'']));d.push([]);d.push(['','','Total',tDr,tCr]);exportToExcel(d,'TB_'+asOf+'.xlsx');};
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
    <div style={S.filterBar}><div><label style={S.label}>As of</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Trial Balance</div><div style={{fontSize:12,color:C.textDim}}>As of {asOf}</div></div>
    <table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={S.td}>{r.code}</td><td style={S.td}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td><td style={S.tdR}>{r.dr>0?fmt(r.dr):''}</td><td style={S.tdR}>{r.cr>0?fmt(r.cr):''}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(tCr)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:12,fontSize:12,color:Math.abs(tDr-tCr)<0.005?C.green:C.red}}>{Math.abs(tDr-tCr)<0.005?'In balance':'Off by $'+fmt(tDr-tCr)}</div></div></div>);}

function BalanceSheet({entityId}){const[balances,setBalances]=useState([]);const[asOf,setAsOf]=useState(today());const fyStart=asOf.slice(0,4)+'-01-01';
  useEffect(()=>{api.getBalances(entityId,{as_of:asOf,close_pl_before:fyStart}).then(setBalances);},[entityId,asOf,fyStart]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=t=>get(t).reduce((s,b)=>s+b.balance,0);
  const ni=sum('Revenue')-sum('Expense');const tA=sum('Asset');const tLE=sum('Liability')+sum('Equity')+ni;
  const doExport=()=>{const d=[['Balance Sheet'],['As of '+asOf],[]];[['Assets','Asset'],['Liabilities','Liability'],['Equity','Equity']].forEach(([t,ty])=>{d.push([t,'']);get(ty).forEach(b=>d.push(['  '+b.name,b.balance]));if(ty==='Equity'&&Math.abs(ni)>0.005)d.push(['  Net Income (current period)',ni]);d.push(['Total '+t,ty==='Equity'?sum(ty)+ni:sum(ty)]);d.push([]);});d.push(['Total L+E',tLE]);exportToExcel(d,'BS_'+asOf+'.xlsx');};
  const Sec=({title,type,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{get(type).map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+C.border+'10'}}>{fmt(b.balance)}</td></tr>)}
    {type==='Equity'&&Math.abs(ni)>0.005&&<tr><td style={{...S.indentTd,fontStyle:'italic'}}>Net Income (current period)</td><td style={{...S.tdR,fontStyle:'italic'}}>{fmt(ni)}</td></tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:10}}>Total {title}</td><td style={{...S.tdR,fontWeight:700}}>${fmt(total)}</td></tr></>);
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
    <div style={S.filterBar}><div><label style={S.label}>As of</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Balance Sheet</div><div style={{fontSize:12,color:C.textDim}}>As of {asOf}</div></div>
    <table style={{...S.table,maxWidth:550,margin:'0 auto'}}><tbody><Sec title="Assets" type="Asset" total={tA}/><tr><td colSpan={2} style={{padding:6}}/></tr>
      <Sec title="Liabilities" type="Liability" total={sum('Liability')}/><tr><td colSpan={2} style={{padding:3}}/></tr><Sec title="Equity" type="Equity" total={sum('Equity')+ni}/>
      <tr style={S.grandTotalRow}><td style={S.tdBold}>Total L + E</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(tLE)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:12,fontSize:12,color:Math.abs(tA-tLE)<0.005?C.green:C.red}}>{Math.abs(tA-tLE)<0.005?'A = L + E':'Off by $'+fmt(tA-tLE)}</div></div></div>);}

function IncomeStatement({entityId}){const[balances,setBalances]=useState([]);const[from,setFrom]=useState(fy_start());const[to,setTo]=useState(today());
  useEffect(()=>{api.getBalances(entityId,{from,to}).then(setBalances);},[entityId,from,to]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=arr=>arr.reduce((s,b)=>s+b.balance,0);
  const rev=get('Revenue');const cogs=get('Expense').filter(b=>b.subtype==='COGS');const opex=get('Expense').filter(b=>b.subtype==='Operating Expense');const other=get('Expense').filter(b=>b.subtype!=='COGS'&&b.subtype!=='Operating Expense');
  const tRev=sum(rev);const gp=tRev-sum(cogs);const oi=gp-sum(opex);const ni=oi-sum(other);
  const doExport=()=>{const d=[['Income Statement'],['Period: '+from+' to '+to],[]];[['Revenue',rev],['COGS',cogs],['OpEx',opex],['Other',other]].forEach(([t,items])=>{if(!items.length)return;d.push([t,'']);items.forEach(b=>d.push(['  '+b.name,b.balance]));d.push(['Total '+t,sum(items)]);d.push([]);});d.push(['Net Income',ni]);exportToExcel(d,'IS_'+from+'_'+to+'.xlsx');};
  const Sec=({title,items,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{items.map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+C.border+'10'}}>{fmt(b.balance)}</td></tr>)}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:10}}>Total {title}</td><td style={{...S.tdR,fontWeight:700}}>${fmt(total)}</td></tr></>);
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Income Statement</div><div style={{fontSize:12,color:C.textDim}}>Period: {from} to {to}</div></div>
    <table style={{...S.table,maxWidth:550,margin:'0 auto'}}><tbody><Sec title="Revenue" items={rev} total={tRev}/>
      {cogs.length>0&&<><Sec title="COGS" items={cogs} total={sum(cogs)}/><tr style={{background:C.border}}><td style={{...S.td,fontWeight:700,color:C.textBright}}>Gross Profit</td><td style={{...S.tdR,fontWeight:700,color:C.textBright}}>${fmt(gp)}</td></tr></>}
      <Sec title="Operating Expenses" items={opex} total={sum(opex)}/><tr style={{background:C.border}}><td style={{...S.td,fontWeight:700,color:C.textBright}}>Operating Income</td><td style={{...S.tdR,fontWeight:700,color:C.textBright}}>${fmt(oi)}</td></tr>
      {other.length>0&&<Sec title="Other Expenses" items={other} total={sum(other)}/>}
      <tr style={S.grandTotalRow}><td style={{...S.tdBold,fontSize:14}}>Net Income</td><td style={{...S.tdBold,textAlign:'right',fontSize:14,color:ni>=0?C.green:C.red}}>${fmt(ni)}</td></tr></tbody></table></div></div>);}

// ═══ Bank Rec ═══
function BankReconciliation({entityId,user}){const[accounts,setAccounts]=useState([]);const[entries,setEntries]=useState([]);const[recs,setRecs]=useState([]);
  const[view,setView]=useState('list');const[selAcct,setSelAcct]=useState('');const[stmtDate,setStmtDate]=useState(today());const[stmtBal,setStmtBal]=useState('');
  const[cleared,setCleared]=useState({});const[checked,setChecked]=useState({});
  const load=useCallback(async()=>{const[a,e,r]=await Promise.all([api.getAccounts(entityId),api.getEntries(entityId),api.getReconciliations(entityId)]);setAccounts(a);setEntries(e);setRecs(r);},[entityId]);
  useEffect(()=>{load();},[load]);
  const bankAccts=accounts.filter(a=>a.bank_acct||(['cash','bank','checking','savings'].some(w=>a.name.toLowerCase().includes(w))&&a.type==='Asset'));
  useEffect(()=>{if(selAcct)api.getCleared(entityId,selAcct).then(setCleared);else setCleared({});},[selAcct,entityId]);
  const getTxns=code=>{const txns=[];entries.forEach(e=>{e.lines.forEach((l,li)=>{if(l.account_code===code){const acct=accounts.find(a=>a.code===code);const isDr=acct?.type==='Asset'||acct?.type==='Expense';txns.push({jeId:e.id,jeNum:e.entry_num,lineIdx:li,date:e.date,memo:e.memo,amount:isDr?(l.debit-l.credit):(l.credit-l.debit),debit:l.debit,credit:l.credit,key:e.id+'-'+li});}});});txns.sort((a,b)=>a.date.localeCompare(b.date));return txns;};
  const txns=selAcct?getTxns(selAcct):[];const uncl=txns.filter(t=>!cleared[t.key]);const bookBal=txns.reduce((s,t)=>s+t.amount,0);const stmtNum=parseFloat(stmtBal)||0;
  const outDep=uncl.filter(t=>!checked[t.key]&&t.amount>0).reduce((s,t)=>s+t.amount,0);const outPay=uncl.filter(t=>!checked[t.key]&&t.amount<0).reduce((s,t)=>s+t.amount,0);
  const diff=bookBal-(stmtNum+outDep+outPay);const isRec=Math.abs(diff)<0.005&&stmtNum!==0;
  const finalize=async()=>{if(!isRec)return;await api.createReconciliation(entityId,{account_code:selAcct,statement_date:stmtDate,statement_balance:stmtNum,book_balance:bookBal,cleared_keys:Object.keys(checked).filter(k=>checked[k])});setChecked({});setStmtBal('');setView('list');load();};
  if(view==='new')return(<div><button style={{...S.btnS,marginBottom:16}} onClick={()=>{setView('list');setSelAcct('');setChecked({});}}>Back</button>
    <div style={S.h1}>New Bank Reconciliation</div><div style={S.card}><div style={S.row}>
      <div style={{...S.col,flex:2}}><label style={S.label}>Account</label><select style={S.select} value={selAcct} onChange={e=>{setSelAcct(e.target.value);setChecked({});}}><option value="">Select...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Statement Date</label><input style={S.input} type="date" value={stmtDate} onChange={e=>setStmtDate(e.target.value)}/></div>
      <div style={S.col}><label style={S.label}>Ending Balance</label><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={stmtBal} onChange={e=>setStmtBal(e.target.value)}/></div></div></div>
    {selAcct&&<><div style={S.card}><div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))',gap:10}}>
      {[{l:'Book Bal',v:bookBal,c:C.textBright},{l:'Statement',v:stmtNum,c:C.textBright},{l:'Out. Dep',v:outDep,c:C.green},{l:'Out. Pay',v:outPay,c:C.red},{l:'Adj Bank',v:stmtNum+outDep+outPay,c:C.accent},{l:'Diff',v:diff,c:isRec?C.green:C.red}].map(s=>
        <div key={s.l} style={{padding:12,borderRadius:8,border:'1px solid '+(s.l==='Diff'&&isRec?C.green+'40':C.border),textAlign:'center',background:s.l==='Diff'&&isRec?C.green+'08':'transparent'}}>
          <div style={{fontSize:10,color:C.textDim}}>{s.l}</div><div style={{fontSize:16,fontWeight:700,color:s.c}}>${fmt(s.v)}</div></div>)}</div></div>
      <div style={S.card}><div style={{display:'flex',justifyContent:'space-between',marginBottom:10}}><div style={S.h2}>Uncleared ({uncl.length})</div><div style={{display:'flex',gap:6}}>
        <button style={{...S.btnS,padding:'4px 8px',fontSize:10}} onClick={()=>{const nc={};uncl.forEach(t=>{nc[t.key]=true;});setChecked(nc);}}>All</button>
        <button style={{...S.btnS,padding:'4px 8px',fontSize:10}} onClick={()=>setChecked({})}>None</button></div></div>
        {uncl.length===0?<div style={{textAlign:'center',padding:20,color:C.textDim}}>All cleared</div>:
        <table style={S.table}><thead><tr><th style={S.thC} width={36}>Clr</th><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Amount</th></tr></thead>
          <tbody>{uncl.map(t=><tr key={t.key} style={checked[t.key]?{background:C.green+'08'}:{cursor:'pointer'}} onClick={()=>setChecked(p=>({...p,[t.key]:!p[t.key]}))}>
            <td style={S.tdC}><input type="checkbox" style={S.checkbox} checked={!!checked[t.key]} readOnly/></td><td style={S.td}>{t.date}</td><td style={S.td}><span style={{color:C.accent}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td>
            <td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600,color:t.amount>=0?C.green:C.red}}>{fmt(t.amount)}</td></tr>)}</tbody></table>}</div>
      <div style={{display:'flex',justifyContent:'flex-end',gap:10}}><button style={isRec?S.btnP:{...S.btnP,opacity:.5,cursor:'not-allowed'}} onClick={finalize}>{isRec?'Finalize':'Diff must be $0'}</button></div></>}</div>);
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div><div style={S.h1}>Bank Reconciliation</div><div style={{fontSize:12,color:C.textDim}}>{recs.length} completed</div></div>
    <button style={S.btnP} onClick={()=>setView('new')}>+ New</button></div>
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(220px,1fr))',gap:12,marginBottom:16}}>
      {bankAccts.map(a=>{const t=getTxns(a.code);const bal=t.reduce((s,x)=>s+x.amount,0);return<div key={a.code} style={S.card}><div style={{fontWeight:700,color:C.textBright}}>{a.name}</div><div style={{fontSize:11,color:C.textDim}}>{a.code}</div><div style={{fontSize:20,fontWeight:700,color:C.textBright,marginTop:8}}>${fmt(bal)}</div></div>;})}</div>
    <div style={S.card}><div style={S.h2}>History</div>{recs.length===0?<div style={{textAlign:'center',padding:30,color:C.textDim}}>None yet</div>:
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.thR}>Stmt</th><th style={S.thR}>Book</th><th style={S.thR}>Cleared</th><th style={S.th}>By</th></tr></thead>
        <tbody>{recs.map(r=><tr key={r.id}><td style={S.td}>{r.statement_date}</td><td style={S.td}>{r.account_code}</td><td style={S.tdR}>${fmt(r.statement_balance)}</td><td style={S.tdR}>${fmt(r.book_balance)}</td><td style={S.tdR}>{r.cleared_count}</td><td style={S.td}>{r.completed_by}</td></tr>)}</tbody></table>}</div></div>);}

// ═══ Entity & User Management ═══
function EntityManagement({refresh,entities,activeEntity,setActiveEntity}){const[showAdd,setShowAdd]=useState(false);const[bulk,setBulk]=useState(false);const[form,setForm]=useState({code:'',name:''});const[bulkText,setBulkText]=useState('');const[err,setErr]=useState('');
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}><div><div style={S.h1}>Entity Management</div><div style={{fontSize:12,color:C.textDim}}>{entities.length} entities</div></div>
    <div style={{display:'flex',gap:8}}><button style={S.btnS} onClick={()=>{setBulk(!bulk);setShowAdd(false);}}>{bulk?'Cancel':'Bulk'}</button><button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setBulk(false);}}>{showAdd?'Cancel':'+ Add'}</button></div></div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div><div style={{...S.col,flex:3}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createEntity(form.code,form.name);setForm({code:'',name:''});setShowAdd(false);setErr('');refresh();}catch(e){setErr(e.message);}}}>Create</button></div>}
    {bulk&&<div style={{...S.card,borderColor:C.accent+'60'}}><div style={S.h2}>Bulk Import</div><textarea style={{...S.input,height:140,fontFamily:'monospace',fontSize:11,resize:'vertical'}} placeholder="CODE, Entity Name" value={bulkText} onChange={e=>setBulkText(e.target.value)}/>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:8}} onClick={async()=>{const ents=bulkText.split('\n').map(l=>l.trim()).filter(Boolean).map(l=>{const[c,...r]=l.split(',').map(p=>p.trim());return{code:c,name:r.join(',')};}).filter(e=>e.code&&e.name);if(!ents.length){setErr('None');return;}try{await api.bulkCreateEntities(ents);setBulkText('');setBulk(false);refresh();}catch(e){setErr(e.message);}}}>Import</button></div>}
    <div style={S.card}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={{...S.th,width:120}}>Actions</th></tr></thead>
      <tbody>{entities.sort((a,b)=>a.code.localeCompare(b.code)).map(e=><tr key={e.id} style={e.id===activeEntity?{background:C.accent+'08'}:{}}><td style={{...S.td,fontWeight:700,color:C.accent}}>{e.code}</td><td style={{...S.td,color:C.textBright}}>{e.name}</td>
        <td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnS,padding:'3px 8px',fontSize:10}} onClick={()=>setActiveEntity(e.id)}>Select</button><button style={{...S.btnD,padding:'3px 8px',fontSize:10}} onClick={async()=>{if(!confirm('Delete?'))return;await api.deleteEntity(e.id);const r=await refresh();if(activeEntity===e.id)setActiveEntity(r[0]?.id||null);}}>Delete</button></div></td></tr>)}</tbody></table></div></div>);}

function UserManagement({currentUser}){const[users,setUsers]=useState([]);const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({name:'',email:'',password:'',role:'Viewer'});const[err,setErr]=useState('');
  const[resetId,setResetId]=useState(null);const[resetPw,setResetPw]=useState('');const[resetMsg,setResetMsg]=useState('');
  useEffect(()=>{api.getUsers().then(setUsers);},[]);
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}><div><div style={S.h1}>User Management</div><div style={{fontSize:12,color:C.textDim}}>{users.length} users</div></div>
    <button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add User'}</button></div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}><div style={S.col}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Email</label><input style={S.input} value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Password</label><input style={S.input} value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Role</label><select style={S.select} value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))}><option>Admin</option><option>Accountant</option><option>Viewer</option></select></div></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.name||!form.email||!form.password){setErr('All required');return;}try{await api.signup(form.name,form.email,form.password,form.role);setForm({name:'',email:'',password:'',role:'Viewer'});setShowAdd(false);setErr('');api.getUsers().then(setUsers);}catch(e){setErr(e.message);}}}>Add</button></div>}
    <div style={S.card}><table style={S.table}><thead><tr><th style={S.th}>Name</th><th style={S.th}>Email</th><th style={S.th}>Role</th><th style={{...S.th,width:180}}>Actions</th></tr></thead>
      <tbody>{users.map(u=><tr key={u.id}><td style={{...S.td,fontWeight:600}}>{u.name}{u.id===currentUser.id?<span style={{color:C.accent,fontSize:10,marginLeft:6}}>(you)</span>:''}</td><td style={S.td}>{u.email}</td><td style={S.td}><span style={S.badge}>{u.role}</span></td>
        <td style={S.td}><div style={{display:'flex',gap:6}}>{u.id!==currentUser.id&&<button style={{...S.btnS,padding:'3px 8px',fontSize:10}} onClick={()=>{setResetId(u.id);setResetPw('');setResetMsg('');}}>Reset PW</button>}
          {u.id!==currentUser.id&&<button style={{...S.btnD,padding:'3px 8px',fontSize:10}} onClick={async()=>{await api.deleteUser(u.id);setUsers(p=>p.filter(x=>x.id!==u.id));}}>Delete</button>}</div></td></tr>)}</tbody></table></div>
    {resetId&&<div style={S.modal} onClick={()=>setResetId(null)}><div style={{...S.modalBox,maxWidth:380,textAlign:'center'}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setResetId(null)}>x</button><div style={{fontSize:16,fontWeight:700,color:C.textBright,marginBottom:16}}>Reset Password</div>
      <div style={{fontSize:12,color:C.textDim,marginBottom:12}}>For: {users.find(u=>u.id===resetId)?.email}</div>
      <input style={S.input} type="password" placeholder="New password" value={resetPw} onChange={e=>{setResetPw(e.target.value);setResetMsg('');}}/>
      {resetMsg&&<div style={{fontSize:11,marginTop:6,color:resetMsg.includes('!')?C.green:C.red}}>{resetMsg}</div>}
      <button style={{...S.btnP,width:'100%',padding:9,marginTop:10}} onClick={async()=>{if(resetPw.length<3){setResetMsg('Min 3');return;}try{await api.adminResetPassword(resetId,resetPw);setResetMsg('Done!');setTimeout(()=>setResetId(null),1500);}catch(e){setResetMsg(e.message);}}}>Reset</button>
    </div></div>}</div>);}
