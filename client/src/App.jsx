import { useState, useEffect, useCallback, useMemo } from 'react';
import { api } from './api';

const fmt = (n) => { const v = Math.abs(n); const s = v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return n < 0 ? '(' + s + ')' : s; };
const today = () => new Date().toISOString().slice(0, 10);

// ─── Colors & Styles ───
const C = { bg:'#0c0e12',card:'#13161d',border:'#1e2230',borderLight:'#282d3e',text:'#c4cad7',textBright:'#eef0f6',textDim:'#6b7394',accent:'#4f8ff7',green:'#34d399',red:'#f87171',orange:'#fb923c',purple:'#a78bfa',teal:'#2dd4bf' };
const S = {
  app:{fontFamily:"'DM Sans',sans-serif",background:C.bg,color:C.text,minHeight:'100vh',display:'flex',flexDirection:'column',fontSize:13},
  topBar:{background:C.card,borderBottom:'1px solid '+C.border,padding:'0 20px',height:52,display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0,zIndex:10},
  body:{display:'flex',flex:1,overflow:'hidden'},
  sidebar:{width:210,background:C.card,borderRight:'1px solid '+C.border,padding:'12px 0',flexShrink:0,overflowY:'auto'},
  navItem:(a)=>({padding:'9px 18px',cursor:'pointer',fontSize:12.5,fontWeight:a?600:400,color:a?C.textBright:C.textDim,background:a?C.border:'transparent',borderLeft:a?'2px solid '+C.accent:'2px solid transparent'}),
  navSection:{padding:'14px 18px 5px',fontSize:10,fontWeight:700,color:C.textDim,textTransform:'uppercase',letterSpacing:'0.1em'},
  main:{flex:1,padding:'20px 24px',overflowY:'auto'},
  card:{background:C.card,border:'1px solid '+C.border,borderRadius:8,padding:20,marginBottom:16},
  h1:{fontSize:20,fontWeight:700,color:C.textBright,marginBottom:2},h2:{fontSize:14,fontWeight:600,color:C.textBright,marginBottom:14},
  sub:{fontSize:12,color:C.textDim,marginBottom:16},
  table:{width:'100%',borderCollapse:'collapse',fontSize:12.5},
  th:{textAlign:'left',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  thR:{textAlign:'right',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  thC:{textAlign:'center',padding:'8px 10px',borderBottom:'2px solid '+C.border,color:C.textDim,fontWeight:600,fontSize:10,textTransform:'uppercase'},
  td:{padding:'7px 10px',borderBottom:'1px solid '+C.border},
  tdR:{padding:'7px 10px',borderBottom:'1px solid '+C.border,textAlign:'right',fontVariantNumeric:'tabular-nums'},
  tdC:{padding:'7px 10px',borderBottom:'1px solid '+C.border,textAlign:'center'},
  tdBold:{padding:'8px 10px',borderBottom:'2px solid '+C.borderLight,color:C.textBright,fontWeight:700,fontVariantNumeric:'tabular-nums'},
  input:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'7px 10px',color:C.text,fontSize:12.5,outline:'none',width:'100%',boxSizing:'border-box'},
  select:{background:C.bg,border:'1px solid '+C.border,borderRadius:5,padding:'7px 10px',color:C.text,fontSize:12.5,outline:'none',width:'100%',boxSizing:'border-box'},
  btnP:{background:'#1a6334',color:'#fff',border:'none',borderRadius:5,padding:'7px 16px',fontSize:12,fontWeight:600,cursor:'pointer'},
  btnS:{background:C.border,color:C.text,border:'1px solid '+C.borderLight,borderRadius:5,padding:'7px 16px',fontSize:12,fontWeight:500,cursor:'pointer'},
  btnD:{background:'#7f1d1d',color:'#fca5a5',border:'none',borderRadius:5,padding:'5px 10px',fontSize:11,fontWeight:600,cursor:'pointer'},
  row:{display:'flex',gap:10,marginBottom:10,flexWrap:'wrap'},col:{flex:1,minWidth:120},
  label:{fontSize:11,color:C.textDim,marginBottom:3,display:'block',fontWeight:500},
  err:{color:C.red,fontSize:11,marginTop:4},success:{color:C.green,fontSize:11,marginTop:4},
  tag:(t)=>{const c={Asset:C.accent,Liability:C.orange,Equity:C.green,Revenue:C.purple,Expense:C.red};return{display:'inline-block',padding:'1px 7px',borderRadius:8,fontSize:10,fontWeight:600,color:c[t]||C.textDim,background:(c[t]||C.textDim)+'18'}},
  badge:{background:C.accent+'20',color:C.accent,padding:'2px 8px',borderRadius:8,fontSize:10,fontWeight:600},
  reportHeader:{borderBottom:'3px double '+C.borderLight,paddingBottom:10,marginBottom:14,textAlign:'center'},
  sectionHeader:{background:C.border,padding:'7px 10px',fontWeight:700,color:C.textBright,fontSize:12},
  indentTd:{padding:'6px 10px 6px 24px',borderBottom:'1px solid '+C.border+'10',fontSize:12.5},
  subtotalRow:{borderTop:'1px solid '+C.borderLight},
  grandTotalRow:{borderTop:'3px double '+C.borderLight,background:C.border},
  link:{color:C.accent,cursor:'pointer',fontSize:12,background:'none',border:'none',padding:0,textDecoration:'underline'},
  checkbox:{width:16,height:16,cursor:'pointer',accentColor:C.green},
  logoIcon:{width:30,height:30,background:'linear-gradient(135deg,'+C.accent+','+C.green+')',borderRadius:7,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:13,color:C.bg},
};

// ─── Auth Screen ───
function AuthScreen({ onLogin }) {
  const [mode,setMode]=useState('login');
  const [email,setEmail]=useState('');const [pw,setPw]=useState('');const [name,setName]=useState('');const [confirmPw,setConfirmPw]=useState('');const [role,setRole]=useState('Accountant');
  const [err,setErr]=useState('');const [success,setSuccess]=useState('');const [loading,setLoading]=useState(false);

  const handleLogin = async () => {
    setLoading(true); setErr('');
    try { const data = await api.login(email.trim().toLowerCase(), pw); api.setToken(data.token); onLogin(data.user); }
    catch (e) { setErr(e.message); } finally { setLoading(false); }
  };
  const handleSignup = async () => {
    if (!name.trim()){setErr('Name required');return;} if(pw.length<3){setErr('Password min 3 chars');return;} if(pw!==confirmPw){setErr("Passwords don't match");return;}
    setLoading(true); setErr('');
    try { await api.signup(name.trim(), email.trim().toLowerCase(), pw, role); setSuccess('Account created!'); setTimeout(()=>{setMode('login');setSuccess('');setErr('');},1200); }
    catch (e) { setErr(e.message); } finally { setLoading(false); }
  };
  const hk=(e)=>{if(e.key==='Enter'){mode==='login'?handleLogin():handleSignup();}};

  return (
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100vh',background:C.bg}}>
      <div style={{...S.card,width:380,textAlign:'center',padding:36}}>
        <div style={{...S.logoIcon,width:44,height:44,fontSize:18,margin:'0 auto 12px'}}>GL</div>
        <div style={{fontSize:20,fontWeight:700,color:C.textBright,marginBottom:2}}>CloudLedger</div>
        <div style={{fontSize:12,color:C.textDim,marginBottom:24}}>Multi-Entity Cloud Accounting</div>
        <div style={{display:'flex',marginBottom:20,borderRadius:6,overflow:'hidden',border:'1px solid '+C.border}}>
          <div onClick={()=>{setMode('login');setErr('');}} style={{flex:1,padding:'8px 0',cursor:'pointer',fontSize:12,fontWeight:600,textAlign:'center',background:mode==='login'?C.accent+'20':'transparent',color:mode==='login'?C.accent:C.textDim}}>Sign In</div>
          <div onClick={()=>{setMode('signup');setErr('');}} style={{flex:1,padding:'8px 0',cursor:'pointer',fontSize:12,fontWeight:600,textAlign:'center',background:mode==='signup'?C.green+'20':'transparent',color:mode==='signup'?C.green:C.textDim}}>Create Account</div>
        </div>
        {mode==='login' ? (<>
          <div style={{marginBottom:10}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          {err&&<div style={S.err}>{err}</div>}
          <button style={{...S.btnP,width:'100%',padding:9,marginTop:6,opacity:loading?0.7:1}} onClick={handleLogin} disabled={loading}>{loading?'Signing in...':'Sign In'}</button>
          <div style={{fontSize:11,color:C.textDim,marginTop:16}}>No account? <button style={S.link} onClick={()=>setMode('signup')}>Create one</button></div>
        </>) : (<>
          <div style={{marginBottom:10}}><input style={S.input} placeholder="Full Name" value={name} onChange={e=>{setName(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:10}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:10}}><input style={S.input} type="password" placeholder="Confirm Password" value={confirmPw} onChange={e=>{setConfirmPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:10,textAlign:'left'}}><label style={S.label}>Role</label><select style={S.select} value={role} onChange={e=>setRole(e.target.value)}><option value="Admin">Admin</option><option value="Accountant">Accountant</option><option value="Viewer">Viewer</option></select></div>
          {err&&<div style={S.err}>{err}</div>}{success&&<div style={S.success}>{success}</div>}
          <button style={{...S.btnP,width:'100%',padding:9,marginTop:6,background:'#065f46',opacity:loading?0.7:1}} onClick={handleSignup} disabled={loading}>{loading?'Creating...':'Create Account'}</button>
          <div style={{fontSize:11,color:C.textDim,marginTop:16}}>Have an account? <button style={S.link} onClick={()=>setMode('login')}>Sign in</button></div>
        </>)}
      </div>
    </div>
  );
}

// ─── Entity Picker ───
function EntityPicker({ entities, activeId, onSelect, onManage }) {
  const [open,setOpen]=useState(false);const [search,setSearch]=useState('');
  const active = entities.find(e=>e.id===activeId);
  const filtered = entities.filter(e=>e.name.toLowerCase().includes(search.toLowerCase())||e.code.toLowerCase().includes(search.toLowerCase()));
  return (
    <div style={{position:'relative'}}>
      <div style={{display:'flex',alignItems:'center',gap:6,cursor:'pointer',padding:'5px 12px',borderRadius:6,background:C.border,border:'1px solid '+C.borderLight}} onClick={()=>setOpen(!open)}>
        <span style={{fontWeight:600,color:C.textBright,fontSize:12}}>{active?.code||'\u2014'}</span>
        <span style={{color:C.textDim,fontSize:12,maxWidth:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{active?.name||'Select entity'}</span>
        <span style={{color:C.textDim,fontSize:10,marginLeft:4}}>{'\u25be'}</span>
      </div>
      {open&&<>
        <div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>{setOpen(false);setSearch('');}}/>
        <div style={{position:'absolute',top:'100%',left:0,background:C.card,border:'1px solid '+C.border,borderRadius:8,maxHeight:350,overflowY:'auto',zIndex:100,boxShadow:'0 12px 40px rgba(0,0,0,0.5)',width:320}}>
          <div style={{position:'sticky',top:0,padding:10,background:C.card,borderBottom:'1px solid '+C.border}}><input style={S.input} placeholder={'Search '+entities.length+' entities...'} value={search} onChange={e=>setSearch(e.target.value)} autoFocus/></div>
          {filtered.map(e=><div key={e.id} style={{padding:'8px 14px',cursor:'pointer',display:'flex',justifyContent:'space-between',background:e.id===activeId?C.accent+'12':'transparent',borderLeft:e.id===activeId?'2px solid '+C.accent:'2px solid transparent'}} onClick={()=>{onSelect(e.id);setOpen(false);setSearch('');}}>
            <div><span style={{fontWeight:600,color:C.textBright,fontSize:12}}>{e.code}</span><span style={{color:C.text,fontSize:12,marginLeft:8}}>{e.name}</span></div>
          </div>)}
          <div style={{borderTop:'1px solid '+C.border,padding:10}}><button style={{...S.btnS,width:'100%',fontSize:11}} onClick={()=>{onManage();setOpen(false);}}>Manage Entities</button></div>
        </div>
      </>}
    </div>
  );
}

// ─── Main App ───
export default function App() {
  const [user,setUser]=useState(null);const [entities,setEntities]=useState([]);const [activeEntity,setActiveEntity]=useState(null);const [page,setPage]=useState('dashboard');const [loading,setLoading]=useState(true);

  // Check existing token on mount
  useEffect(()=>{
    const token = api.getToken();
    if (token) { api.me().then(u=>{if(u)setUser(u);}).catch(()=>api.clearToken()).finally(()=>setLoading(false)); }
    else setLoading(false);
  },[]);

  // Load entities when user logs in
  useEffect(()=>{
    if (user) { api.getEntities().then(e=>{setEntities(e);if(e.length>0&&!activeEntity)setActiveEntity(e[0].id);}); }
  },[user]);

  const refreshEntities = useCallback(async ()=>{const e=await api.getEntities();setEntities(e);return e;},[]);
  const handleLogin = (u)=>{setUser(u);};
  const handleLogout = ()=>{api.clearToken();setUser(null);setEntities([]);setActiveEntity(null);};

  const canAccess = (s)=>{if(!user)return false;const r=user.role;if(r==='Admin')return true;const p={Accountant:['entries','reports','coa','bankrec'],Viewer:['reports']};return(p[r]||[]).includes(s);};

  if (loading) return <div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:C.textDim}}>Loading...</div></div>;
  if (!user) return <AuthScreen onLogin={handleLogin}/>;

  const entity = entities.find(e=>e.id===activeEntity);
  const navItems = [
    {id:'dashboard',label:'Dashboard',section:'reports'},
    {id:'d1',divider:true,label:'TRANSACTIONS'},
    {id:'journal',label:'Journal Entries',section:'entries'},
    {id:'d2',divider:true,label:'ACCOUNTS'},
    {id:'coa',label:'Chart of Accounts',section:'coa'},
    {id:'ledger',label:'General Ledger',section:'reports'},
    {id:'d2b',divider:true,label:'BANKING'},
    {id:'bankrec',label:'Bank Reconciliation',section:'bankrec'},
    {id:'d3',divider:true,label:'REPORTS'},
    {id:'trial',label:'Trial Balance',section:'reports'},
    {id:'bs',label:'Balance Sheet',section:'reports'},
    {id:'is',label:'Income Statement',section:'reports'},
    {id:'d4',divider:true,label:'ADMINISTRATION'},
    {id:'entities',label:'Entities ('+entities.length+')',section:'all'},
    {id:'users',label:'Users',section:'all'},
  ];

  return (
    <div style={S.app}>
      <div style={S.topBar}>
        <div style={{display:'flex',alignItems:'center',gap:16}}>
          <div style={{display:'flex',alignItems:'center',gap:10}}><div style={S.logoIcon}>GL</div><div style={{fontSize:16,fontWeight:700,color:C.textBright}}>CloudLedger</div></div>
          <div style={{width:1,height:24,background:C.border}}/>
          <EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')}/>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <span style={{fontSize:12}}>{user.name}</span><span style={S.badge}>{user.role}</span>
          <button style={{...S.btnS,padding:'4px 10px',fontSize:11}} onClick={handleLogout}>Sign Out</button>
        </div>
      </div>
      <div style={S.body}>
        <div style={S.sidebar}>{navItems.map(n=>n.divider?<div key={n.id} style={S.navSection}>{n.label}</div>
          :(n.section==='all'?user.role==='Admin':canAccess(n.section))?<div key={n.id} style={S.navItem(page===n.id)} onClick={()=>setPage(n.id)}>{n.label}</div>:null)}</div>
        <div style={S.main}>
          {page==='dashboard'&&<Dashboard entityId={activeEntity}/>}
          {page==='journal'&&activeEntity&&<JournalEntries entityId={activeEntity} user={user}/>}
          {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} canEdit={canAccess('coa')}/>}
          {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity}/>}
          {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user}/>}
          {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity}/>}
          {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity}/>}
          {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity}/>}
          {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
          {page==='users'&&<UserManagement currentUser={user}/>}
        </div>
      </div>
    </div>
  );
}

// ─── Dashboard ───
function Dashboard({ entityId }) {
  const [summary,setSummary]=useState([]);
  useEffect(()=>{api.getSummary().then(setSummary);},[entityId]);
  const curr = summary.find(e=>e.id===entityId);
  return (<div>
    <div style={S.h1}>Dashboard</div><div style={S.sub}>{summary.length} entities</div>
    {curr&&<div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:12,marginBottom:20}}>
      {[{l:'Total Assets',v:curr.assets,c:C.accent},{l:'Total Liabilities',v:curr.liabilities,c:C.orange},{l:'Revenue',v:curr.revenue,c:C.purple},{l:'Expenses',v:curr.expenses,c:C.red},{l:'Net Income',v:curr.net_income,c:curr.net_income>=0?C.green:C.red},{l:'Entries',v:curr.entry_count,c:C.textDim,raw:1}].map(s=>(
        <div key={s.l} style={{...S.card,textAlign:'center',padding:16}}><div style={{fontSize:22,fontWeight:700,color:s.c,fontVariantNumeric:'tabular-nums'}}>{s.raw?s.v:'$'+fmt(s.v)}</div><div style={{fontSize:11,color:C.textDim,marginTop:3}}>{s.l}</div></div>
      ))}
    </div>}
    <div style={S.card}><div style={S.h2}>All Entities</div><table style={S.table}>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={S.thR}>Assets</th><th style={S.thR}>Liabilities</th><th style={S.thR}>Net Income</th><th style={S.thR}>JEs</th></tr></thead>
      <tbody>{summary.sort((a,b)=>a.code.localeCompare(b.code)).map(e=>(
        <tr key={e.id} style={e.id===entityId?{background:C.accent+'08'}:{}}><td style={{...S.td,fontWeight:600,color:C.accent}}>{e.code}</td><td style={S.td}>{e.name}</td><td style={S.tdR}>{fmt(e.assets)}</td><td style={S.tdR}>{fmt(e.liabilities)}</td><td style={{...S.tdR,color:e.net_income>=0?C.green:C.red,fontWeight:600}}>{fmt(e.net_income)}</td><td style={S.tdR}>{e.entry_count}</td></tr>
      ))}</tbody></table>
    </div>
  </div>);
}

// ─── Journal Entries ───
function JournalEntries({ entityId, user }) {
  const [entries,setEntries]=useState([]);const [accounts,setAccounts]=useState([]);const [showForm,setShowForm]=useState(false);
  const blank={date:today(),memo:'',lines:[{account_code:'',debit:'',credit:''},{account_code:'',debit:'',credit:''}]};
  const [form,setForm]=useState(blank);const [err,setErr]=useState('');
  const load=useCallback(async()=>{const [e,a]=await Promise.all([api.getEntries(entityId),api.getAccounts(entityId)]);setEntries(e);setAccounts(a);},[entityId]);
  useEffect(()=>{load();},[load]);
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:''}]}));
  const removeLine=(i)=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const totalDr=form.lines.reduce((s,l)=>s+(parseFloat(l.debit)||0),0);
  const totalCr=form.lines.reduce((s,l)=>s+(parseFloat(l.credit)||0),0);
  const balanced=Math.abs(totalDr-totalCr)<0.005&&totalDr>0;

  const post=async()=>{
    if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}
    if(form.lines.some(l=>!l.account_code)){setErr('All lines need account');return;}
    if(!balanced){setErr('Must balance');return;}
    try{await api.createEntry(entityId,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseFloat(l.debit)||0,credit:parseFloat(l.credit)||0}))});setForm(blank);setShowForm(false);setErr('');load();}
    catch(e){setErr(e.message);}
  };
  const del=async(id)=>{await api.deleteEntry(entityId,id);load();};

  return (<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div><div style={S.h1}>Journal Entries</div><div style={{fontSize:12,color:C.textDim}}>{entries.length} entries</div></div>
      <button style={S.btnP} onClick={()=>setShowForm(!showForm)}>{showForm?'Cancel':'+ New Entry'}</button>
    </div>
    {showForm&&<div style={{...S.card,borderColor:C.green+'60'}}>
      <div style={S.h2}>New Journal Entry</div>
      <div style={S.row}><div style={S.col}><label style={S.label}>Date</label><input style={S.input} type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/></div>
        <div style={{...S.col,flex:3}}><label style={S.label}>Memo</label><input style={S.input} placeholder="Description" value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))}/></div></div>
      <table style={{...S.table,marginBottom:10}}><thead><tr><th style={S.th}>Account</th><th style={{...S.thR,width:120}}>Debit</th><th style={{...S.thR,width:120}}>Credit</th><th style={{...S.th,width:30}}></th></tr></thead>
        <tbody>{form.lines.map((l,i)=>(
          <tr key={i}><td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}><select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select...</option>{accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{a.code} \u2013 {a.name}</option>)}</select></td>
            <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.debit} onChange={e=>updateLine(i,'debit',e.target.value)}/></td>
            <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.credit} onChange={e=>updateLine(i,'credit',e.target.value)}/></td>
            <td style={{padding:'3px 4px',borderBottom:'1px solid '+C.border,textAlign:'center'}}>{form.lines.length>2&&<span style={{cursor:'pointer',color:C.red}} onClick={()=>removeLine(i)}>{'\u00d7'}</span>}</td></tr>
        ))}<tr style={{background:C.border}}><td style={{...S.tdBold,textAlign:'right'}}>Total</td><td style={{...S.tdBold,textAlign:'right'}}>{'$'+fmt(totalDr)}</td><td style={{...S.tdBold,textAlign:'right'}}>{'$'+fmt(totalCr)}</td><td style={S.tdBold}></td></tr></tbody></table>
      <div style={{display:'flex',gap:8,alignItems:'center'}}><button style={S.btnS} onClick={addLine}>+ Line</button><div style={{flex:1}}/>
        {!balanced&&totalDr>0&&<span style={{color:C.orange,fontSize:11}}>Diff: ${fmt(totalDr-totalCr)}</span>}
        {balanced&&<span style={{color:C.green,fontSize:11}}>{'\u2713'} Balanced</span>}
        {err&&<span style={S.err}>{err}</span>}<button style={S.btnP} onClick={post}>Post Entry</button></div>
    </div>}
    <div style={S.card}>{entries.length===0?<div style={{textAlign:'center',padding:40,color:C.textDim,fontSize:12}}>No entries yet</div>:
      <table style={S.table}><thead><tr><th style={S.th}>JE #</th><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.th}>By</th><th style={{...S.th,width:30}}></th></tr></thead>
        <tbody>{entries.map(e=>e.lines.map((l,i)=>(
          <tr key={e.id+'-'+i} style={i===0?{borderTop:'2px solid '+C.borderLight}:{}}>
            {i===0&&<td style={S.td} rowSpan={e.lines.length}><span style={{fontWeight:600,color:C.accent}}>JE-{String(e.entry_num).padStart(4,'0')}</span></td>}
            {i===0&&<td style={S.td} rowSpan={e.lines.length}>{e.date}</td>}
            <td style={{...S.td,paddingLeft:l.credit>0&&l.debit===0?24:10}}>{l.account_code} \u2013 {accounts.find(a=>a.code===l.account_code)?.name||'?'}</td>
            {i===0&&<td style={S.td} rowSpan={e.lines.length}>{e.memo}</td>}
            <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td><td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td>
            {i===0&&<td style={S.td} rowSpan={e.lines.length}>{e.created_by}</td>}
            {i===0&&<td style={S.td} rowSpan={e.lines.length}><span style={{cursor:'pointer',color:C.red,fontSize:12}} onClick={()=>del(e.id)}>{'\ud83d\uddd1'}</span></td>}
          </tr>)))}</tbody></table>}
    </div>
  </div>);
}

// ─── Chart of Accounts ───
function ChartOfAccounts({ entityId, canEdit }) {
  const [accounts,setAccounts]=useState([]);const [showAdd,setShowAdd]=useState(false);
  const [form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const [err,setErr]=useState('');
  const load=useCallback(async()=>{setAccounts(await api.getAccounts(entityId));},[entityId]);
  useEffect(()=>{load();},[load]);
  const add=async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createAccount(entityId,form);setForm({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});setShowAdd(false);setErr('');load();}catch(e){setErr(e.message);}};
  const del=async(code)=>{try{await api.deleteAccount(entityId,code);load();}catch(e){alert(e.message);}};
  return (<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div><div style={S.h1}>Chart of Accounts</div><div style={{fontSize:12,color:C.textDim}}>{accounts.length} accounts</div></div>
      {canEdit&&<button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add'}</button>}
    </div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}>
      <div style={S.col}><label style={S.label}>Code</label><input style={S.input} value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Subtype</label><input style={S.input} value={form.subtype} onChange={e=>setForm(f=>({...f,subtype:e.target.value}))}/></div>
    </div><div style={{marginBottom:10}}><label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank/cash account</label></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={add}>Add</button></div>}
    <div style={S.card}><table style={S.table}>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Name</th><th style={S.th}>Type</th><th style={S.th}>Subtype</th><th style={S.thC}>Bank</th>{canEdit&&<th style={{...S.th,width:30}}></th>}</tr></thead>
      <tbody>{accounts.map(a=>(
        <tr key={a.code}><td style={{...S.td,fontWeight:600}}>{a.code}</td><td style={S.td}>{a.name}</td><td style={S.td}><span style={S.tag(a.type)}>{a.type}</span></td><td style={{...S.td,color:C.textDim}}>{a.subtype}</td>
          <td style={S.tdC}>{a.bank_acct?<span style={{color:C.teal}}>{'\u2713'}</span>:''}</td>
          {canEdit&&<td style={S.td}><span style={{cursor:'pointer',color:C.red}} onClick={()=>del(a.code)}>{'\u00d7'}</span></td>}</tr>
      ))}</tbody></table></div>
  </div>);
}

// ─── General Ledger ───
function GeneralLedger({ entityId }) {
  const [entries,setEntries]=useState([]);const [accounts,setAccounts]=useState([]);const [filter,setFilter]=useState('');
  useEffect(()=>{Promise.all([api.getEntries(entityId),api.getAccounts(entityId)]).then(([e,a])=>{setEntries(e);setAccounts(a);});},[entityId]);
  const filtered = accounts.filter(a=>!filter||a.code===filter).sort((a,b)=>a.code.localeCompare(b.code));
  return (<div><div style={S.h1}>General Ledger</div><div style={S.sub}></div>
    <div style={{...S.card,padding:14}}><div style={{maxWidth:280}}><label style={S.label}>Filter</label>
      <select style={S.select} value={filter} onChange={e=>setFilter(e.target.value)}><option value="">All Accounts</option>{accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{a.code} \u2013 {a.name}</option>)}</select></div></div>
    {filtered.map(acct=>{
      const txns=[];entries.forEach(e=>{e.lines.forEach((l,li)=>{if(l.account_code===acct.code)txns.push({...l,date:e.date,memo:e.memo,jeNum:e.entry_num});});});
      if(txns.length===0&&!filter)return null;txns.sort((a,b)=>a.date.localeCompare(b.date));let running=0;const dr=acct.type==='Asset'||acct.type==='Expense';
      return (<div key={acct.code} style={S.card}>
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10}}><span style={{fontWeight:700,color:C.textBright}}>{acct.code}</span><span>{acct.name}</span><span style={S.tag(acct.type)}>{acct.type}</span></div>
        {txns.length===0?<div style={{color:C.textDim,fontSize:12}}>No transactions</div>:
        <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th></tr></thead>
          <tbody>{txns.map((t,i)=>{running+=dr?(t.debit-t.credit):(t.credit-t.debit);return(
            <tr key={i}><td style={S.td}>{t.date}</td><td style={S.td}><span style={{color:C.accent}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td><td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600}}>{fmt(running)}</td></tr>);})}</tbody></table>}
      </div>);
    })}
  </div>);
}

// ─── Reports: Trial Balance, Balance Sheet, Income Statement ───
function TrialBalance({ entityId }) {
  const [balances,setBalances]=useState([]);
  useEffect(()=>{api.getBalances(entityId).then(setBalances);},[entityId]);
  let totalDr=0,totalCr=0;
  const rows=balances.filter(b=>Math.abs(b.balance)>0.005).map(b=>{
    const isDr=b.type==='Asset'||b.type==='Expense';
    const dr=(isDr&&b.balance>0)||(!isDr&&b.balance<0)?Math.abs(b.balance):0;
    const cr=(isDr&&b.balance<0)||(!isDr&&b.balance>0)?Math.abs(b.balance):0;
    totalDr+=dr;totalCr+=cr;return{...b,dr,cr};
  });
  return (<div><div style={S.card}>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Trial Balance</div><div style={{fontSize:12,color:C.textDim}}>As of {today()}</div></div>
    <table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={S.td}>{r.code}</td><td style={S.td}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td><td style={S.tdR}>{r.dr>0?fmt(r.dr):''}</td><td style={S.tdR}>{r.cr>0?fmt(r.cr):''}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalDr)}</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalCr)}</td></tr>
      </tbody></table>
    <div style={{textAlign:'center',marginTop:12,fontSize:12,color:Math.abs(totalDr-totalCr)<0.005?C.green:C.red}}>{Math.abs(totalDr-totalCr)<0.005?'\u2713 In balance':'\u26a0 Difference: $'+fmt(totalDr-totalCr)}</div>
  </div></div>);
}

function BalanceSheet({ entityId }) {
  const [balances,setBalances]=useState([]);
  useEffect(()=>{api.getBalances(entityId).then(setBalances);},[entityId]);
  const get=(t)=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);
  const sum=(t)=>get(t).reduce((s,b)=>s+b.balance,0);
  const ni=sum('Revenue')-sum('Expense'); const totalA=sum('Asset'); const totalLE=sum('Liability')+sum('Equity')+ni;
  const Sec=({title,type,total})=>(<>
    <tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>
    {get(type).map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+C.border+'10'}}>{fmt(b.balance)}</td></tr>)}
    {type==='Equity'&&Math.abs(ni)>0.005&&<tr><td style={{...S.indentTd,fontStyle:'italic'}}>Net Income</td><td style={{...S.tdR,fontStyle:'italic'}}>{fmt(ni)}</td></tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:10}}>Total {title}</td><td style={{...S.tdR,fontWeight:700}}>${fmt(total)}</td></tr>
  </>);
  return (<div><div style={S.card}>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Balance Sheet</div><div style={{fontSize:12,color:C.textDim}}>As of {today()}</div></div>
    <table style={{...S.table,maxWidth:550,margin:'0 auto'}}><tbody>
      <Sec title="Assets" type="Asset" total={totalA}/><tr><td colSpan={2} style={{padding:6}}></td></tr>
      <Sec title="Liabilities" type="Liability" total={sum('Liability')}/><tr><td colSpan={2} style={{padding:3}}></td></tr>
      <Sec title="Equity" type="Equity" total={sum('Equity')+ni}/>
      <tr style={S.grandTotalRow}><td style={S.tdBold}>Total L + E</td><td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalLE)}</td></tr>
    </tbody></table>
    <div style={{textAlign:'center',marginTop:12,fontSize:12,color:Math.abs(totalA-totalLE)<0.005?C.green:C.red}}>{Math.abs(totalA-totalLE)<0.005?'\u2713 A = L + E':'\u26a0 Off by $'+fmt(totalA-totalLE)}</div>
  </div></div>);
}

function IncomeStatement({ entityId }) {
  const [balances,setBalances]=useState([]);
  useEffect(()=>{api.getBalances(entityId).then(setBalances);},[entityId]);
  const get=(t)=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);
  const sum=(arr)=>arr.reduce((s,b)=>s+b.balance,0);
  const rev=get('Revenue'); const cogs=get('Expense').filter(b=>b.subtype==='COGS'); const opex=get('Expense').filter(b=>b.subtype==='Operating Expense'); const other=get('Expense').filter(b=>b.subtype!=='COGS'&&b.subtype!=='Operating Expense');
  const totalRev=sum(rev); const gp=totalRev-sum(cogs); const oi=gp-sum(opex); const ni=oi-sum(other);
  const Sec=({title,items,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>
    {items.map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+C.border+'10'}}>{fmt(b.balance)}</td></tr>)}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:10}}>Total {title}</td><td style={{...S.tdR,fontWeight:700}}>${fmt(total)}</td></tr></>);
  return (<div><div style={S.card}>
    <div style={S.reportHeader}><div style={{fontSize:18,fontWeight:700,color:C.textBright}}>Income Statement</div><div style={{fontSize:12,color:C.textDim}}>Period Ending {today()}</div></div>
    <table style={{...S.table,maxWidth:550,margin:'0 auto'}}><tbody>
      <Sec title="Revenue" items={rev} total={totalRev}/>
      {cogs.length>0&&<><Sec title="COGS" items={cogs} total={sum(cogs)}/><tr style={{background:C.border}}><td style={{...S.td,fontWeight:700,color:C.textBright}}>Gross Profit</td><td style={{...S.tdR,fontWeight:700,color:C.textBright}}>${fmt(gp)}</td></tr></>}
      <Sec title="Operating Expenses" items={opex} total={sum(opex)}/>
      <tr style={{background:C.border}}><td style={{...S.td,fontWeight:700,color:C.textBright}}>Operating Income</td><td style={{...S.tdR,fontWeight:700,color:C.textBright}}>${fmt(oi)}</td></tr>
      {other.length>0&&<Sec title="Other Expenses" items={other} total={sum(other)}/>}
      <tr style={S.grandTotalRow}><td style={{...S.tdBold,fontSize:14}}>Net Income</td><td style={{...S.tdBold,textAlign:'right',fontSize:14,color:ni>=0?C.green:C.red}}>${fmt(ni)}</td></tr>
    </tbody></table>
  </div></div>);
}

// ─── Bank Reconciliation ───
function BankReconciliation({ entityId, user }) {
  const [accounts,setAccounts]=useState([]);const [entries,setEntries]=useState([]);const [recs,setRecs]=useState([]);
  const [view,setView]=useState('list');const [selAcct,setSelAcct]=useState('');const [stmtDate,setStmtDate]=useState(today());const [stmtBal,setStmtBal]=useState('');
  const [cleared,setCleared]=useState({});const [checked,setChecked]=useState({});const [err,setErr]=useState('');

  const load=useCallback(async()=>{const [a,e,r]=await Promise.all([api.getAccounts(entityId),api.getEntries(entityId),api.getReconciliations(entityId)]);setAccounts(a);setEntries(e);setRecs(r);},[entityId]);
  useEffect(()=>{load();},[load]);

  const bankAccts=accounts.filter(a=>a.bank_acct||(['cash','bank','checking','savings'].some(w=>a.name.toLowerCase().includes(w))&&a.type==='Asset'));

  // Load cleared items when account selected
  useEffect(()=>{if(selAcct){api.getCleared(entityId,selAcct).then(setCleared);}else setCleared({});},[selAcct,entityId]);

  const getTxns=(code)=>{const txns=[];entries.forEach(e=>{e.lines.forEach((l,li)=>{if(l.account_code===code){const acct=accounts.find(a=>a.code===code);const isDr=acct?.type==='Asset'||acct?.type==='Expense';const amt=isDr?(l.debit-l.credit):(l.credit-l.debit);txns.push({jeId:e.id,jeNum:e.entry_num,lineIdx:li,date:e.date,memo:e.memo,amount:amt,debit:l.debit,credit:l.credit,key:e.id+'-'+li});}});});txns.sort((a,b)=>a.date.localeCompare(b.date));return txns;};

  const txns=selAcct?getTxns(selAcct):[];
  const uncleared=txns.filter(t=>!cleared[t.key]);
  const bookBal=txns.reduce((s,t)=>s+t.amount,0);
  const stmtNum=parseFloat(stmtBal)||0;
  const outDep=uncleared.filter(t=>!checked[t.key]&&t.amount>0).reduce((s,t)=>s+t.amount,0);
  const outPay=uncleared.filter(t=>!checked[t.key]&&t.amount<0).reduce((s,t)=>s+t.amount,0);
  const adjBal=stmtNum+outDep+outPay;
  const diff=bookBal-adjBal;
  const isRec=Math.abs(diff)<0.005&&stmtNum!==0;

  const finalize=async()=>{
    if(!isRec){setErr('Difference must be $0.00');return;}
    const clearedKeys=Object.keys(checked).filter(k=>checked[k]);
    await api.createReconciliation(entityId,{account_code:selAcct,statement_date:stmtDate,statement_balance:stmtNum,book_balance:bookBal,cleared_keys:clearedKeys});
    setChecked({});setStmtBal('');setView('list');load();
  };

  if(view==='new')return(<div>
    <button style={{...S.btnS,marginBottom:16}} onClick={()=>{setView('list');setSelAcct('');setChecked({});setErr('');}}>{'< Back'}</button>
    <div style={S.h1}>New Bank Reconciliation</div>
    <div style={S.card}><div style={S.h2}>Setup</div><div style={S.row}>
      <div style={{...S.col,flex:2}}><label style={S.label}>Account</label><select style={S.select} value={selAcct} onChange={e=>{setSelAcct(e.target.value);setChecked({});}}><option value="">Select...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{a.code} \u2013 {a.name}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Statement Date</label><input style={S.input} type="date" value={stmtDate} onChange={e=>setStmtDate(e.target.value)}/></div>
      <div style={S.col}><label style={S.label}>Ending Balance</label><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={stmtBal} onChange={e=>setStmtBal(e.target.value)}/></div>
    </div></div>
    {selAcct&&<>
      <div style={S.card}><div style={S.h2}>Summary</div><div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:12}}>
        {[{l:'Book Balance',v:bookBal,c:C.textBright},{l:'Statement',v:stmtNum,c:C.textBright},{l:'Out. Deposits',v:outDep,c:C.green},{l:'Out. Payments',v:outPay,c:C.red},{l:'Adj. Bank Bal',v:adjBal,c:C.accent},{l:'Difference',v:diff,c:isRec?C.green:C.red}].map(s=>(
          <div key={s.l} style={{padding:14,borderRadius:8,border:'1px solid '+(s.l==='Difference'&&isRec?C.green+'40':C.border),textAlign:'center',background:s.l==='Difference'&&isRec?C.green+'08':'transparent'}}>
            <div style={{fontSize:10,color:C.textDim}}>{s.l}</div><div style={{fontSize:18,fontWeight:700,color:s.c}}>${fmt(s.v)}</div>
            {s.l==='Difference'&&isRec&&<div style={{fontSize:10,color:C.green}}>{'\u2713'} Reconciled!</div>}
          </div>))}
      </div></div>
      <div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
        <div style={S.h2}>Uncleared ({uncleared.length})</div>
        <div style={{display:'flex',gap:8}}><button style={{...S.btnS,padding:'4px 10px',fontSize:10}} onClick={()=>{const nc={};uncleared.forEach(t=>{nc[t.key]=true;});setChecked(nc);}}>All</button><button style={{...S.btnS,padding:'4px 10px',fontSize:10}} onClick={()=>setChecked({})}>None</button></div>
      </div>
        {uncleared.length===0?<div style={{textAlign:'center',padding:20,color:C.textDim,fontSize:12}}>All cleared</div>:
        <table style={S.table}><thead><tr><th style={S.thC} width={40}>Clear</th><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Amount</th></tr></thead>
          <tbody>{uncleared.map(t=>(
            <tr key={t.key} style={checked[t.key]?{background:C.green+'08'}:{cursor:'pointer'}} onClick={()=>setChecked(p=>({...p,[t.key]:!p[t.key]}))}>
              <td style={S.tdC}><input type="checkbox" style={S.checkbox} checked={!!checked[t.key]} readOnly/></td>
              <td style={S.td}>{t.date}</td><td style={S.td}><span style={{color:C.accent}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td>
              <td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td>
              <td style={{...S.tdR,fontWeight:600,color:t.amount>=0?C.green:C.red}}>{fmt(t.amount)}</td></tr>))}</tbody></table>}
      </div>
      <div style={{display:'flex',justifyContent:'flex-end',gap:10,alignItems:'center'}}>
        {err&&<span style={S.err}>{err}</span>}
        <button style={isRec?S.btnP:{...S.btnP,opacity:0.5,cursor:'not-allowed'}} onClick={finalize}>{isRec?'\u2713 Finalize':'Difference must be $0.00'}</button>
      </div>
    </>}
  </div>);

  return (<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div><div style={S.h1}>Bank Reconciliation</div><div style={{fontSize:12,color:C.textDim}}>{recs.length} completed</div></div>
      <button style={S.btnP} onClick={()=>setView('new')}>+ New Reconciliation</button>
    </div>
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(220px,1fr))',gap:12,marginBottom:16}}>
      {bankAccts.map(a=>{const t=getTxns(a.code);const bal=t.reduce((s,x)=>s+x.amount,0);const lastR=recs.filter(r=>r.account_code===a.code).sort((x,y)=>y.statement_date.localeCompare(x.statement_date))[0];
        return(<div key={a.code} style={S.card}><div style={{fontWeight:700,color:C.textBright,fontSize:13}}>{a.name}</div><div style={{fontSize:11,color:C.textDim}}>{a.code}</div>
          <div style={{fontSize:20,fontWeight:700,color:C.textBright,marginTop:8}}>${fmt(bal)}</div>
          <div style={{fontSize:10,color:C.textDim,marginTop:4}}>{lastR?'Last: '+lastR.statement_date:'Never reconciled'}</div></div>);})}
    </div>
    <div style={S.card}><div style={S.h2}>History</div>{recs.length===0?<div style={{textAlign:'center',padding:30,color:C.textDim,fontSize:12}}>No reconciliations yet</div>:
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.thR}>Stmt Bal</th><th style={S.thR}>Book Bal</th><th style={S.thR}>Cleared</th><th style={S.th}>By</th></tr></thead>
        <tbody>{recs.map(r=>(
          <tr key={r.id}><td style={S.td}>{r.statement_date}</td><td style={S.td}>{r.account_code}</td><td style={S.tdR}>${fmt(r.statement_balance)}</td><td style={S.tdR}>${fmt(r.book_balance)}</td><td style={S.tdR}>{r.cleared_count} items</td><td style={S.td}>{r.completed_by}</td></tr>
        ))}</tbody></table>}
    </div>
  </div>);
}

// ─── Entity Management ───
function EntityManagement({ refresh, entities, activeEntity, setActiveEntity }) {
  const [showAdd,setShowAdd]=useState(false);const [bulk,setBulk]=useState(false);const [form,setForm]=useState({code:'',name:''});const [bulkText,setBulkText]=useState('');const [err,setErr]=useState('');
  const addOne=async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createEntity(form.code,form.name);setForm({code:'',name:''});setShowAdd(false);setErr('');refresh();}catch(e){setErr(e.message);}};
  const addBulk=async()=>{const lines=bulkText.split('\n').map(l=>l.trim()).filter(Boolean);const ents=lines.map(l=>{const[code,...rest]=l.split(',').map(p=>p.trim());return{code,name:rest.join(',')};}).filter(e=>e.code&&e.name);
    if(ents.length===0){setErr('No valid entries');return;}try{await api.bulkCreateEntities(ents);setBulkText('');setBulk(false);setErr('');refresh();}catch(e){setErr(e.message);}};
  const del=async(id)=>{if(!confirm('Delete this entity?'))return;await api.deleteEntity(id);const e=await refresh();if(activeEntity===id)setActiveEntity(e[0]?.id||null);};
  return (<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div><div style={S.h1}>Entity Management</div><div style={{fontSize:12,color:C.textDim}}>{entities.length} entities</div></div>
      <div style={{display:'flex',gap:8}}><button style={S.btnS} onClick={()=>{setBulk(!bulk);setShowAdd(false);}}>{bulk?'Cancel':'Bulk Import'}</button><button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setBulk(false);}}>{showAdd?'Cancel':'+ Add'}</button></div>
    </div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div><div style={{...S.col,flex:3}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div></div>{err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={addOne}>Create</button></div>}
    {bulk&&<div style={{...S.card,borderColor:C.accent+'60'}}><div style={S.h2}>Bulk Import</div><div style={{fontSize:11,color:C.textDim,marginBottom:8}}>Format: CODE, Entity Name</div>
      <textarea style={{...S.input,height:140,fontFamily:'monospace',fontSize:11,resize:'vertical'}} placeholder="CLR-F1, County Line Rail Fund I" value={bulkText} onChange={e=>setBulkText(e.target.value)}/>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:8}} onClick={addBulk}>Import</button></div>}
    <div style={S.card}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={{...S.th,width:120}}>Actions</th></tr></thead>
      <tbody>{entities.sort((a,b)=>a.code.localeCompare(b.code)).map(e=>(
        <tr key={e.id} style={e.id===activeEntity?{background:C.accent+'08'}:{}}><td style={{...S.td,fontWeight:700,color:C.accent}}>{e.code}</td><td style={{...S.td,color:C.textBright}}>{e.name}</td>
          <td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnS,padding:'3px 8px',fontSize:10}} onClick={()=>setActiveEntity(e.id)}>Select</button><button style={{...S.btnD,padding:'3px 8px',fontSize:10}} onClick={()=>del(e.id)}>Delete</button></div></td></tr>
      ))}</tbody></table></div>
  </div>);
}

// ─── User Management ───
function UserManagement({ currentUser }) {
  const [users,setUsers]=useState([]);const [showAdd,setShowAdd]=useState(false);const [form,setForm]=useState({name:'',email:'',password:'',role:'Viewer'});const [err,setErr]=useState('');
  useEffect(()=>{api.getUsers().then(setUsers);},[]);
  const add=async()=>{if(!form.name||!form.email||!form.password){setErr('All fields required');return;}try{await api.signup(form.name,form.email,form.password,form.role);setForm({name:'',email:'',password:'',role:'Viewer'});setShowAdd(false);setErr('');api.getUsers().then(setUsers);}catch(e){setErr(e.message);}};
  const del=async(id)=>{await api.deleteUser(id);setUsers(u=>u.filter(x=>x.id!==id));};
  return (<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
      <div><div style={S.h1}>User Management</div><div style={{fontSize:12,color:C.textDim}}>{users.length} users</div></div>
      <button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add User'}</button>
    </div>
    {showAdd&&<div style={{...S.card,borderColor:C.green+'60'}}><div style={S.row}>
      <div style={S.col}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Email</label><input style={S.input} value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Password</label><input style={S.input} value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Role</label><select style={S.select} value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))}><option>Admin</option><option>Accountant</option><option>Viewer</option></select></div>
    </div>{err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={add}>Add</button></div>}
    <div style={S.card}><table style={S.table}><thead><tr><th style={S.th}>Name</th><th style={S.th}>Email</th><th style={S.th}>Role</th><th style={{...S.th,width:30}}></th></tr></thead>
      <tbody>{users.map(u=>(
        <tr key={u.id}><td style={{...S.td,fontWeight:600}}>{u.name}{u.id===currentUser.id?<span style={{color:C.accent,fontSize:10,marginLeft:6}}>(you)</span>:''}</td><td style={S.td}>{u.email}</td><td style={S.td}><span style={S.badge}>{u.role}</span></td>
          <td style={S.td}>{u.id!==currentUser.id&&<span style={{cursor:'pointer',color:C.red}} onClick={()=>del(u.id)}>{'\u00d7'}</span>}</td></tr>
      ))}</tbody></table></div>
  </div>);
}
