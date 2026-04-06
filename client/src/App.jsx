import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { api } from './api';
import * as XLSX from 'xlsx';

const fmt = n => { const v = Math.abs(n); const s = v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return n < 0 ? '(' + s + ')' : s; };
const today = () => new Date().toISOString().slice(0, 10);
const fy_start = () => new Date().getFullYear() + '-01-01';
const fmtSize = b => b > 1048576 ? (b/1048576).toFixed(1)+' MB' : (b/1024).toFixed(0)+' KB';
const acctLabel = (code, name) => code + ' - ' + name;
function exportToExcel(data, fn) { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = data[0].map((_, ci) => ({ wch: Math.min(Math.max(...data.map(r => String(r[ci]||'').length), 8)+2, 40) })); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); XLSX.writeFile(wb, fn); }
const BLANK_JE = () => ({date:today(),memo:'',lines:[{account_code:'',debit:'',credit:''},{account_code:'',debit:'',credit:''}]});
const SIDEBAR_KEY = 'cl_sidebar';

// ─── Cloud Ledger Logo SVG ───
function Logo({size=32}) {
  const s = size/40;
  return (<svg width={size} height={size} viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
    <defs><linearGradient id="clg" x1="4" y1="8" x2="36" y2="36" gradientUnits="userSpaceOnUse">
      <stop stopColor="#1d4ed8"/><stop offset="1" stopColor="#2563eb"/></linearGradient></defs>
    {/* Cloud shape */}
    <path d="M32 22a6 6 0 00-5.8-6 8 8 0 00-15.4-1A5 5 0 007 19.5 5 5 0 0010 24h0" fill="none" stroke="url(#clg)" strokeWidth="2.2" strokeLinecap="round"/>
    <path d="M14 22a3.5 3.5 0 013-5.8 5.5 5.5 0 0110.6.8A4.2 4.2 0 0130 21a4.2 4.2 0 01-1 3" fill="none" stroke="url(#clg)" strokeWidth="1.8" strokeLinecap="round" opacity="0.5"/>
    {/* Ledger lines */}
    <line x1="11" y1="27" x2="29" y2="27" stroke="#2563eb" strokeWidth="1.6" strokeLinecap="round"/>
    <line x1="11" y1="30.5" x2="25" y2="30.5" stroke="#2563eb" strokeWidth="1.6" strokeLinecap="round" opacity="0.6"/>
    <line x1="11" y1="34" x2="27" y2="34" stroke="#2563eb" strokeWidth="1.6" strokeLinecap="round" opacity="0.35"/>
    {/* Accent dots on ledger lines */}
    <circle cx="14" cy="27" r="1.3" fill="#059669"/>
    <circle cx="22" cy="30.5" r="1.3" fill="#2563eb" opacity="0.6"/>
  </svg>);
}

// Auto-categorization hints (deterministic, keyword-based)
const HINTS = [
  // Payroll & compensation
  { kw:['payroll','salary','wage','bonus','compensation','adp','paychex','gusto','payday','direct dep','net pay','garnish'], sub:['Operating Expense'], re:/salari|payroll/i },
  // Rent & occupancy
  { kw:['rent','lease','office space','property mgmt','landlord','realty','real estate'], sub:['Operating Expense'], re:/rent/i },
  // Utilities
  { kw:['electric','gas bill','water bill','utility','utilities','power','energy','pgande','con edison','duke energy','sewer','trash','waste mgmt'], sub:['Operating Expense'], re:/utilit/i },
  // Insurance
  { kw:['insurance','premium','policy','allstate','state farm','geico','liberty mutual','hartford','travelers','workers comp','liability ins','general ins'], sub:['Operating Expense'], re:/insurance/i },
  // Office & supplies
  { kw:['supplies','office depot','staples','amazon','walmart','target','costco','sams club','home depot','lowes','paper','toner','shipping','fedex','ups','usps','postage'], sub:['Operating Expense'], re:/supplies|office/i },
  // Marketing & advertising
  { kw:['advertising','marketing','google ads','facebook','meta ads','ad spend','linkedin','yelp','social media','print ad','billboard','promo','campaign','mailchimp','hubspot','constant contact'], sub:['Operating Expense'], re:/market/i },
  // Professional services
  { kw:['legal','attorney','law firm','accounting','cpa','consulting','professional fee','advisory','audit','tax prep','bookkeep'], sub:['Operating Expense'], re:/profession|legal|consult/i },
  // Technology & software
  { kw:['software','subscription','saas','cloud','hosting','aws','azure','google cloud','microsoft','adobe','zoom','slack','quickbooks','xero','netsuite','salesforce','dropbox','github'], sub:['Operating Expense'], re:/software|tech|computer/i },
  // Travel & meals
  { kw:['travel','airline','hotel','airbnb','uber','lyft','taxi','parking','toll','mileage','meal','restaurant','doordash','grubhub','catering'], sub:['Operating Expense'], re:/travel|meal/i },
  // Interest & bank charges
  { kw:['interest','loan payment','finance charge','bank fee','service charge','wire fee','nsf','overdraft','monthly fee','annual fee','credit card fee'], sub:['Other Expense'], re:/interest/i },
  // Depreciation
  { kw:['depreciation','amortization'], sub:['Operating Expense'], re:/deprec|amort/i },
  // Revenue / deposits
  { kw:['deposit','payment received','wire in','ach credit','revenue','sales','client payment','customer payment','invoice payment','consulting revenue','service revenue','tenant','rental income'], sub:['Operating Revenue','Other Revenue'], re:/revenue|income|sales/i },
  // Loan proceeds / financing
  { kw:['loan proceeds','draw','line of credit','loc advance','note payable'], sub:['Long-term Liability'], re:/note|loan/i },
];
function suggestAccount(desc, accounts, bankCode) {
  if (!desc) return null; const d = desc.toLowerCase();
  for (const h of HINTS) {
    if (h.kw.some(k => d.includes(k))) {
      const cs = accounts.filter(a => a.code !== bankCode && h.sub.includes(a.subtype));
      return cs.find(a => h.re.test(a.name)) || cs[0] || null;
    }
  }
  return null;
}

// ─── Light Theme Design Tokens ───
const T = {
  bg: '#f8f9fb', bgCard: '#ffffff', bgHover: '#f3f4f6', bgElevated: '#f9fafb',
  border: '#e5e7eb', borderLight: '#f3f4f6', borderFocus: '#3b82f640',
  text: '#1f2937', textBright: '#111827', textMuted: '#6b7280', textDim: '#9ca3af',
  accent: '#2563eb', accentLight: '#3b82f6', accentDim: '#2563eb10',
  green: '#059669', greenDim: '#05966910', greenBorder: '#05966930',
  red: '#dc2626', redDim: '#dc262610', orange: '#d97706', orangeDim: '#d9770610',
  purple: '#7c3aed', purpleDim: '#7c3aed10', teal: '#0d9488', tealDim: '#0d948810',
  shadow: '0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)',
  shadowLg: '0 4px 16px rgba(0,0,0,0.1), 0 2px 4px rgba(0,0,0,0.06)',
  radius: '10px', radiusSm: '8px', radiusXs: '6px',
  sidebarBg: '#1e293b', sidebarText: '#94a3b8', sidebarActive: '#e2e8f0', sidebarAccent: '#3b82f6',
};

const S = {
  app: { fontFamily: "'Inter',-apple-system,sans-serif", background: T.bg, color: T.text, minHeight: '100vh', display: 'flex', flexDirection: 'column', fontSize: 13, lineHeight: 1.5 },
  topBar: { background: T.bgCard, borderBottom: '1px solid '+T.border, padding: '0 24px', height: 56, display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0, zIndex: 10, boxShadow: T.shadow },
  body: { display: 'flex', flex: 1, overflow: 'hidden' },
  sidebar: col => ({ width: col ? 56 : 224, background: T.sidebarBg, padding: col ? '12px 4px' : '16px 0', flexShrink: 0, overflowY: 'auto', overflowX: 'hidden', transition: 'width 0.2s ease', display: 'flex', flexDirection: 'column' }),
  navItem: (a, col) => ({ padding: col ? '10px 0' : '9px 20px', cursor: 'pointer', fontSize: 12.5, fontWeight: a ? 600 : 400, color: a ? T.sidebarActive : T.sidebarText, background: a ? '#ffffff12' : 'transparent', borderRadius: col ? T.radiusXs : '0 6px 6px 0', margin: col ? '2px 6px' : '1px 8px 1px 0', borderLeft: col ? 'none' : (a ? '3px solid '+T.sidebarAccent : '3px solid transparent'), transition: 'all 0.12s', textAlign: col ? 'center' : 'left', whiteSpace: 'nowrap', overflow: 'hidden' }),
  navSection: col => ({ padding: col ? '12px 0 4px' : '18px 20px 6px', fontSize: 9, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.12em', textAlign: col ? 'center' : 'left' }),
  main: { flex: 1, padding: '28px 32px', overflowY: 'auto', background: T.bg },
  card: { background: T.bgCard, border: '1px solid '+T.border, borderRadius: T.radius, padding: 24, marginBottom: 20, boxShadow: T.shadow },
  cardFlush: { background: T.bgCard, border: '1px solid '+T.border, borderRadius: T.radius, marginBottom: 20, boxShadow: T.shadow, overflow: 'hidden' },
  h1: { fontSize: 24, fontWeight: 700, color: T.textBright, marginBottom: 4, letterSpacing: '-0.02em' },
  h2: { fontSize: 15, fontWeight: 600, color: T.textBright, marginBottom: 16 },
  sub: { fontSize: 13, color: T.textMuted, marginBottom: 20 },
  table: { width: '100%', borderCollapse: 'collapse', fontSize: 13 },
  th: { textAlign: 'left', padding: '10px 14px', borderBottom: '2px solid '+T.border, color: T.textMuted, fontWeight: 600, fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.06em', background: T.bgElevated },
  thR: { textAlign: 'right', padding: '10px 14px', borderBottom: '2px solid '+T.border, color: T.textMuted, fontWeight: 600, fontSize: 10, textTransform: 'uppercase', letterSpacing: '0.06em', background: T.bgElevated },
  thC: { textAlign: 'center', padding: '10px 14px', borderBottom: '2px solid '+T.border, color: T.textMuted, fontWeight: 600, fontSize: 10, textTransform: 'uppercase', background: T.bgElevated },
  td: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight },
  tdR: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight, textAlign: 'right', fontVariantNumeric: 'tabular-nums' },
  tdC: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight, textAlign: 'center' },
  tdBold: { padding: '10px 14px', borderBottom: '2px solid '+T.border, color: T.textBright, fontWeight: 700, fontVariantNumeric: 'tabular-nums' },
  input: { background: '#fff', border: '1px solid '+T.border, borderRadius: T.radiusXs, padding: '9px 12px', color: T.text, fontSize: 13, outline: 'none', width: '100%', boxSizing: 'border-box' },
  inputSm: { background: '#fff', border: '1px solid '+T.border, borderRadius: T.radiusXs, padding: '6px 10px', color: T.text, fontSize: 12, outline: 'none', boxSizing: 'border-box' },
  select: { background: '#fff', border: '1px solid '+T.border, borderRadius: T.radiusXs, padding: '9px 12px', color: T.text, fontSize: 13, outline: 'none', width: '100%', boxSizing: 'border-box' },
  selectSm: { background: '#fff', border: '1px solid '+T.border, borderRadius: T.radiusXs, padding: '6px 10px', color: T.text, fontSize: 12, outline: 'none', boxSizing: 'border-box' },
  btnP: { background: T.accent, color: '#fff', border: 'none', borderRadius: T.radiusXs, padding: '9px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' },
  btnS: { background: '#fff', color: T.text, border: '1px solid '+T.border, borderRadius: T.radiusXs, padding: '9px 20px', fontSize: 13, fontWeight: 500, cursor: 'pointer' },
  btnD: { background: T.redDim, color: T.red, border: '1px solid '+T.red+'30', borderRadius: T.radiusXs, padding: '6px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer' },
  btnExport: { background: T.tealDim, color: T.teal, border: '1px solid '+T.teal+'30', borderRadius: T.radiusXs, padding: '8px 16px', fontSize: 12, fontWeight: 600, cursor: 'pointer' },
  btnGhost: { background: 'transparent', color: T.textMuted, border: 'none', padding: '6px 10px', fontSize: 12, cursor: 'pointer' },
  row: { display: 'flex', gap: 12, marginBottom: 12, flexWrap: 'wrap' }, col: { flex: 1, minWidth: 130 },
  label: { fontSize: 11, color: T.textMuted, marginBottom: 4, display: 'block', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.03em' },
  err: { color: T.red, fontSize: 12, marginTop: 6 }, success: { color: T.green, fontSize: 12, marginTop: 6 },
  tag: t => { const c = { Asset:T.accent, Liability:T.orange, Equity:T.green, Revenue:T.purple, Expense:T.red }; return { display:'inline-block',padding:'2px 10px',borderRadius:20,fontSize:10,fontWeight:600,color:c[t]||T.textDim,background:(c[t]||T.textDim)+'12' }; },
  badge: { background: T.accentDim, color: T.accent, padding: '3px 10px', borderRadius: 20, fontSize: 11, fontWeight: 600 },
  link: { color: T.accent, cursor: 'pointer', fontSize: 12, background: 'none', border: 'none', padding: 0, textDecoration: 'none' },
  checkbox: { width: 16, height: 16, cursor: 'pointer', accentColor: T.green },
  logoIcon: { width: 32, height: 32, background: 'linear-gradient(135deg,'+T.accent+','+T.green+')', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, fontSize: 14, color: '#fff' },
  modal: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', backdropFilter: 'blur(4px)', zIndex: 200, display: 'flex', alignItems: 'center', justifyContent: 'center' },
  modalBox: { background: '#fff', border: '1px solid '+T.border, borderRadius: '14px', width: '94%', maxWidth: 960, maxHeight: '92vh', overflowY: 'auto', padding: 28, position: 'relative', boxShadow: T.shadowLg },
  modalClose: { position: 'absolute', top: 16, right: 20, cursor: 'pointer', color: T.textMuted, fontSize: 18, background: T.bgElevated, border: '1px solid '+T.border, borderRadius: 6, width: 28, height: 28, display: 'flex', alignItems: 'center', justifyContent: 'center' },
  filterBar: { display: 'flex', alignItems: 'flex-end', gap: 14, flexWrap: 'wrap', marginBottom: 20 },
  attachLink: { display:'inline-flex',alignItems:'center',gap:4,padding:'3px 10px',borderRadius:6,fontSize:11,color:T.accent,background:T.accentDim,textDecoration:'none',marginRight:4,marginBottom:2,fontWeight:500 },
  reportHeader: { borderBottom: '2px solid '+T.border, paddingBottom: 12, marginBottom: 16, textAlign: 'center' },
  sectionHeader: { background: T.bgElevated, padding: '8px 14px', fontWeight: 600, color: T.textBright, fontSize: 12, borderBottom: '1px solid '+T.border },
  indentTd: { padding: '8px 14px 8px 28px', borderBottom: '1px solid '+T.borderLight, fontSize: 13 },
  subtotalRow: { borderTop: '1px solid '+T.border },
  grandTotalRow: { borderTop: '2px solid '+T.border, background: T.bgElevated },
  summaryBar: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 14, marginBottom: 20 },
  summaryItem: { padding: '16px 14px', background: T.bgCard, border: '1px solid '+T.border, borderRadius: T.radiusSm, textAlign: 'center', boxShadow: T.shadow },
  statVal: { fontSize: 26, fontWeight: 700, fontVariantNumeric: 'tabular-nums', lineHeight: 1.2 },
  statLabel: { fontSize: 11, color: T.textMuted, marginTop: 6, fontWeight: 500 },
};
const NI = { dashboard:'\u25a3', journal:'\u270e', coa:'\u2630', ledger:'\u2261', banktxn:'\u21c5', bankrec:'\u2611', trial:'\u2696', bs:'\u25a6', is:'\u25a4', entities:'\u2302', users:'\u263a' };

// ─── Autocomplete ───
function AccountAutocomplete({accounts,value,onChange,placeholder,exclude}){
  const[q,setQ]=useState('');const[open,setOpen]=useState(false);const ref=useRef(null);
  const sel=accounts.find(a=>a.code===value);
  const filtered=useMemo(()=>{const s=q.toLowerCase();return accounts.filter(a=>(!exclude||a.code!==exclude)&&(a.code.toLowerCase().includes(s)||a.name.toLowerCase().includes(s))).sort((a,b)=>a.code.localeCompare(b.code));},[accounts,q,exclude]);
  useEffect(()=>{const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener('mousedown',h);return()=>document.removeEventListener('mousedown',h);},[]);
  return(<div ref={ref} style={{position:'relative'}}><input style={S.inputSm} placeholder={placeholder||'Search account...'} value={open?q:(sel?acctLabel(sel.code,sel.name):'')}
    onFocus={()=>{setOpen(true);setQ('');}} onChange={e=>{setQ(e.target.value);setOpen(true);}} onKeyDown={e=>{if(e.key==='Escape')setOpen(false);if(e.key==='Enter'&&filtered.length>0){onChange(filtered[0].code);setOpen(false);}}}/>
    {open&&filtered.length>0&&<div style={{position:'absolute',top:'100%',left:0,right:0,background:'#fff',border:'1px solid '+T.border,borderRadius:T.radiusSm,maxHeight:340,overflowY:'auto',zIndex:50,boxShadow:T.shadowLg,marginTop:4}}>
      {filtered.map(a=><div key={a.code} style={{padding:'8px 12px',cursor:'pointer',fontSize:12,display:'flex',justifyContent:'space-between',background:a.code===value?T.accentDim:'transparent'}}
        onClick={()=>{onChange(a.code);setOpen(false);}} onMouseEnter={e=>e.currentTarget.style.background=T.bgHover} onMouseLeave={e=>e.currentTarget.style.background=a.code===value?T.accentDim:'transparent'}>
        <span><b style={{color:T.textBright}}>{a.code}</b> <span style={{color:T.textMuted}}>{a.name}</span></span><span style={S.tag(a.type)}>{a.type}</span></div>)}</div>}</div>);}

// ─── Auth ───
function AuthScreen({onLogin}){const[mode,setMode]=useState('login');const[email,setEmail]=useState('');const[pw,setPw]=useState('');const[name,setName]=useState('');const[confirmPw,setConfirmPw]=useState('');const[role,setRole]=useState('Accountant');
  const[err,setErr]=useState('');const[success,setSuccess]=useState('');const[loading,setLoading]=useState(false);const[tempPw,setTempPw]=useState('');
  const doLogin=async()=>{setLoading(true);setErr('');try{const d=await api.login(email.trim().toLowerCase(),pw);api.setToken(d.token);onLogin(d.user);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doSignup=async()=>{if(!name.trim()){setErr('Name required');return;}if(pw.length<3){setErr('Min 3 chars');return;}if(pw!==confirmPw){setErr("Passwords don't match");return;}setLoading(true);setErr('');try{await api.signup(name.trim(),email.trim().toLowerCase(),pw,role);setSuccess('Account created!');setTimeout(()=>{setMode('login');setSuccess('');},1200);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doForgot=async()=>{if(!email.trim()){setErr('Enter email');return;}setLoading(true);setErr('');try{const r=await api.forgotPassword(email.trim().toLowerCase());setTempPw(r.temp_password);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const hk=e=>{if(e.key==='Enter'){mode==='login'?doLogin():mode==='signup'?doSignup():doForgot();}};
  return(<div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100vh',background:'#f1f5f9'}}>
    <div style={{background:'#fff',border:'1px solid '+T.border,borderRadius:16,width:420,padding:44,textAlign:'center',boxShadow:T.shadowLg}}>
      <div style={{margin:'0 auto 16px',width:48}}><Logo size={48}/></div>
      <div style={{fontSize:24,fontWeight:800,color:T.textBright,marginBottom:4}}>CloudLedger</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:32}}>Multi-Entity Cloud Accounting</div>
      {mode==='forgot'?(<>
        <div style={{fontSize:15,fontWeight:600,color:T.textBright,marginBottom:20}}>Reset Password</div>
        <div style={{marginBottom:12}}><input style={S.input} placeholder="Email address" value={email} onChange={e=>{setEmail(e.target.value);setErr('');setTempPw('');}} onKeyDown={hk}/></div>
        {err&&<div style={S.err}>{err}</div>}
        {tempPw&&<div style={{background:T.greenDim,border:'1px solid '+T.greenBorder,borderRadius:T.radiusSm,padding:20,margin:'12px 0'}}><div style={{fontSize:12,color:T.green}}>Temporary password:</div><div style={{fontSize:20,fontWeight:700,color:T.textBright,fontFamily:'monospace',letterSpacing:2}}>{tempPw}</div></div>}
        <button style={{...S.btnP,width:'100%',padding:11,marginTop:8}} onClick={doForgot} disabled={loading}>{loading?'...':'Reset Password'}</button>
        <div style={{marginTop:20}}><button style={S.link} onClick={()=>{setMode('login');setErr('');setTempPw('');}}>Back to Sign In</button></div>
      </>):(<>
        <div style={{display:'flex',marginBottom:24,borderRadius:T.radiusSm,overflow:'hidden',border:'1px solid '+T.border}}>
          <div onClick={()=>{setMode('login');setErr('');}} style={{flex:1,padding:'10px 0',cursor:'pointer',fontSize:13,fontWeight:600,textAlign:'center',background:mode==='login'?T.accentDim:'transparent',color:mode==='login'?T.accent:T.textMuted}}>Sign In</div>
          <div onClick={()=>{setMode('signup');setErr('');}} style={{flex:1,padding:'10px 0',cursor:'pointer',fontSize:13,fontWeight:600,textAlign:'center',background:mode==='signup'?T.greenDim:'transparent',color:mode==='signup'?T.green:T.textMuted}}>Create Account</div></div>
        {mode==='login'?(<>
          <div style={{marginBottom:12}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:12}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          {err&&<div style={S.err}>{err}</div>}
          <button style={{...S.btnP,width:'100%',padding:11,marginTop:8}} onClick={doLogin} disabled={loading}>{loading?'...':'Sign In'}</button>
          <div style={{display:'flex',justifyContent:'space-between',marginTop:18}}><button style={S.link} onClick={()=>{setMode('forgot');setErr('');}}>Forgot password?</button><button style={S.link} onClick={()=>setMode('signup')}>Create account</button></div>
        </>):(<>
          <div style={{marginBottom:12}}><input style={S.input} placeholder="Full Name" value={name} onChange={e=>{setName(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:12}}><input style={S.input} placeholder="Email" value={email} onChange={e=>{setEmail(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:12}}><input style={S.input} type="password" placeholder="Password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:12}}><input style={S.input} type="password" placeholder="Confirm Password" value={confirmPw} onChange={e=>{setConfirmPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
          <div style={{marginBottom:12,textAlign:'left'}}><label style={S.label}>Role</label><select style={S.select} value={role} onChange={e=>setRole(e.target.value)}><option value="Admin">Admin</option><option value="Accountant">Accountant</option><option value="Viewer">Viewer</option></select></div>
          {err&&<div style={S.err}>{err}</div>}{success&&<div style={S.success}>{success}</div>}
          <button style={{...S.btnP,width:'100%',padding:11,marginTop:8,background:T.green}} onClick={doSignup} disabled={loading}>Create Account</button>
          <div style={{marginTop:18}}><button style={S.link} onClick={()=>setMode('login')}>Back to Sign In</button></div></>)}</>)}
    </div></div>);}

// ─── Modals ───
function SettingsModal({onClose,user,onUserUpdate}){
  const[tab,setTab]=useState('profile');
  const[name,setName]=useState(user.name);const[email,setEmail]=useState(user.email);const[profileErr,setProfileErr]=useState('');const[profileOk,setProfileOk]=useState(false);const[saving,setSaving]=useState(false);
  const[cur,setCur]=useState('');const[nw,setNw]=useState('');const[cf,setCf]=useState('');const[pwErr,setPwErr]=useState('');const[pwOk,setPwOk]=useState(false);
  const saveProfile=async()=>{if(!name.trim()||!email.trim()){setProfileErr('Name and email required');return;}setSaving(true);setProfileErr('');
    try{const updated=await api.updateProfile(name.trim(),email.trim().toLowerCase());onUserUpdate(updated);setProfileOk(true);setTimeout(()=>setProfileOk(false),3000);}catch(e){setProfileErr(e.message);}finally{setSaving(false);}};
  const changePw=async()=>{if(nw.length<3){setPwErr('Min 3 chars');return;}if(nw!==cf){setPwErr("Passwords don't match");return;}
    try{await api.changePassword(cur,nw);setPwOk(true);setCur('');setNw('');setCf('');setTimeout(()=>setPwOk(false),3000);}catch(e){setPwErr(e.message);}};
  return(<div style={S.modal} onClick={onClose}><div style={{...S.modalBox,maxWidth:500}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Settings</div>
    <div style={{display:'flex',gap:0,marginBottom:24,borderBottom:'2px solid '+T.border}}>
      {['profile','password'].map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:'10px 20px',fontSize:13,fontWeight:tab===t?600:400,color:tab===t?T.accent:T.textMuted,background:'transparent',border:'none',borderBottom:tab===t?'2px solid '+T.accent:'2px solid transparent',marginBottom:-2,cursor:'pointer',textTransform:'capitalize'}}>{t}</button>)}</div>
    {tab==='profile'&&<div>
      <div style={{marginBottom:14}}><label style={S.label}>Name</label><input style={S.input} value={name} onChange={e=>{setName(e.target.value);setProfileErr('');}}/></div>
      <div style={{marginBottom:14}}><label style={S.label}>Login Email</label><input style={S.input} type="email" value={email} onChange={e=>{setEmail(e.target.value);setProfileErr('');}}/></div>
      <div style={{marginBottom:14}}><label style={S.label}>Role</label><div style={{padding:'9px 12px',background:T.bgElevated,borderRadius:T.radiusXs,border:'1px solid '+T.border,color:T.textMuted}}>{user.role} <span style={{fontSize:11}}>(contact an admin to change)</span></div></div>
      {profileErr&&<div style={S.err}>{profileErr}</div>}{profileOk&&<div style={S.success}>Profile updated! You may need to sign out and back in for the name to appear everywhere.</div>}
      <button style={{...S.btnP,marginTop:8,opacity:saving?.6:1}} onClick={saveProfile} disabled={saving}>{saving?'Saving...':'Save Profile'}</button></div>}
    {tab==='password'&&<div>
      <div style={{marginBottom:14}}><label style={S.label}>Current Password</label><input style={S.input} type="password" value={cur} onChange={e=>{setCur(e.target.value);setPwErr('');}}/></div>
      <div style={{marginBottom:14}}><label style={S.label}>New Password</label><input style={S.input} type="password" value={nw} onChange={e=>{setNw(e.target.value);setPwErr('');}}/></div>
      <div style={{marginBottom:14}}><label style={S.label}>Confirm New Password</label><input style={S.input} type="password" value={cf} onChange={e=>{setCf(e.target.value);setPwErr('');}}/></div>
      {pwErr&&<div style={S.err}>{pwErr}</div>}{pwOk&&<div style={S.success}>Password changed!</div>}
      <button style={{...S.btnP,marginTop:8}} onClick={changePw}>Change Password</button></div>}
  </div></div>);}

function QuickAddAccountModal({entityId,onClose,onCreated}){const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  return(<div style={S.modal} onClick={onClose}><div style={{...S.modalBox,maxWidth:640}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button><div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Add New Account</div>
    <div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
    <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
    {err&&<div style={S.err}>{err}</div>}<div style={{display:'flex',gap:10,marginTop:12}}><button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Code and name required');return;}try{const a=await api.createAccount(entityId,form);onCreated(a);onClose();}catch(e){setErr(e.message);}}}>Add Account</button><button style={S.btnS} onClick={onClose}>Cancel</button></div>
  </div></div>);}

// ─── JE Modal — form state received from App (persists across open/close) ───
function JournalEntryModal({entityId,user,onClose,onPosted,form,setForm,pendingFiles,setPendingFiles}){
  const[accounts,setAccounts]=useState([]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[posting,setPosting]=useState(false);const[posted,setPosted]=useState('');
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts);},[entityId]);
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const tDr=form.lines.reduce((s,l)=>s+(parseFloat(l.debit)||0),0);const tCr=form.lines.reduce((s,l)=>s+(parseFloat(l.credit)||0),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;
  const discard=()=>{setForm(BLANK_JE());setPendingFiles([]);};
  const onFilesSelected=e=>{const files=Array.from(e.target.files);if(files.length>0)setPendingFiles(p=>[...p,...files]);e.target.value='';};
  const post=async()=>{if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}if(form.lines.some(l=>!l.account_code)){setErr('All lines need an account');return;}if(!bal){setErr('Entry must balance');return;}
    setPosting(true);setErr('');try{const r=await api.createEntry(entityId,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseFloat(l.debit)||0,credit:parseFloat(l.credit)||0}))});
      let msg='JE-'+String(r.entry_num).padStart(4,'0')+' posted';
      if(pendingFiles.length>0){try{const u=await api.uploadAttachments(entityId,r.id,pendingFiles);msg+=' with '+u.length+' attachment(s)';}catch(ue){msg+=' (attachments failed: '+ue.message+')';}}
      setForm(BLANK_JE());setPendingFiles([]);setPosted(msg+'!');setTimeout(()=>setPosted(''),5000);if(onPosted)onPosted();}
    catch(e){setErr(e.message);}finally{setPosting(false);}};
  const hasContent=form.memo||form.lines.some(l=>l.account_code||l.debit||l.credit)||pendingFiles.length>0;

  return(<div style={S.modal}><div style={{...S.modalBox,maxWidth:980}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright}}>New Journal Entry</div>
      <div style={{display:'flex',alignItems:'center',gap:12}}>
        {hasContent&&<span style={{fontSize:11,color:T.orange,fontWeight:500}}>In progress</span>}
        {hasContent&&<button style={{...S.btnGhost,color:T.red,fontSize:12}} onClick={discard}>Discard</button>}
      </div></div>
    <div style={{background:T.bgElevated,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:18,marginBottom:16}}>
      <div style={S.row}><div style={{...S.col,maxWidth:170}}><label style={S.label}>Date</label><input style={S.input} type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/></div>
        <div style={{...S.col,flex:4}}><label style={S.label}>Memo / Description</label><input style={S.input} placeholder="What is this entry for?" value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))}/></div></div></div>
    <div style={{...S.cardFlush,marginBottom:16}}><table style={S.table}><thead><tr><th style={S.th}>Account</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select account...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.debit} onChange={e=>updateLine(i,'debit',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.credit} onChange={e=>updateLine(i,'credit',e.target.value)}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
    <div style={{border:'1px solid '+(pendingFiles.length>0?T.teal+'40':T.border),borderRadius:T.radiusSm,padding:16,marginBottom:16,background:pendingFiles.length>0?T.tealDim:'#fff'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:pendingFiles.length>0?12:0}}>
        <span style={{fontSize:12,fontWeight:600,color:pendingFiles.length>0?T.teal:T.textMuted}}>{pendingFiles.length>0?pendingFiles.length+' file(s) attached':'No attachments'}</span>
        <div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
          <button style={{...S.btnS,padding:'7px 16px',pointerEvents:'none'}}>Attach files</button>
          <input type="file" multiple accept=".pdf,.xlsx,.xls,.csv,.png,.jpg,.jpeg,.eml,.msg,.doc,.docx"
            style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:'pointer'}} onChange={onFilesSelected}/></div></div>
      {pendingFiles.length>0&&<div style={{display:'flex',flexWrap:'wrap',gap:6}}>{pendingFiles.map((f,i)=><div key={i} style={{display:'flex',alignItems:'center',gap:6,background:T.bgElevated,padding:'6px 12px',borderRadius:6,fontSize:12,border:'1px solid '+T.border}}>
        <span style={{fontWeight:500}}>{f.name}</span><span style={{color:T.textDim}}>({fmtSize(f.size)})</span>
        <button style={{...S.btnGhost,color:T.red,padding:0,fontSize:14}} onClick={()=>setPendingFiles(p=>p.filter((_,j)=>j!==i))}>&times;</button></div>)}</div>}</div>
    <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
      <button style={S.btnS} onClick={addLine}>+ Add line</button>
      <button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New account</button>
      <div style={{flex:1}}/>
      {!bal&&tDr>0&&<span style={{fontSize:12,color:T.orange,fontWeight:600}}>Off by ${fmt(tDr-tCr)}</span>}
      {bal&&<span style={{fontSize:12,color:T.green,fontWeight:600}}>Balanced</span>}
      {err&&<span style={S.err}>{err}</span>}{posted&&<span style={S.success}>{posted}</span>}
      <button style={{...S.btnP,padding:'10px 28px',fontSize:14,opacity:posting?.6:1}} onClick={post} disabled={posting}>{posting?'Posting...':'Post Entry'}</button></div>
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)))}/>}
  </div></div>);}

// ─── Entity Picker ───
function EntityPicker({entities,activeId,onSelect,onManage}){const[open,setOpen]=useState(false);const[search,setSearch]=useState('');const active=entities.find(e=>e.id===activeId);
  const filtered=entities.filter(e=>e.name.toLowerCase().includes(search.toLowerCase())||e.code.toLowerCase().includes(search.toLowerCase()));
  return(<div style={{position:'relative'}}><div style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',padding:'6px 14px',borderRadius:T.radiusSm,background:T.bgElevated,border:'1px solid '+T.border}} onClick={()=>setOpen(!open)}>
    <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>{active?.code||'-'}</span><span style={{color:T.textMuted,fontSize:12,maxWidth:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{active?.name}</span>
    <span style={{color:T.textDim,fontSize:9}}>{'\u25bc'}</span></div>
    {open&&<><div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>{setOpen(false);setSearch('');}}/>
      <div style={{position:'absolute',top:'100%',left:0,background:'#fff',border:'1px solid '+T.border,borderRadius:T.radius,maxHeight:380,overflowY:'auto',zIndex:100,boxShadow:T.shadowLg,width:340,marginTop:6}}>
        <div style={{position:'sticky',top:0,padding:12,background:'#fff',borderBottom:'1px solid '+T.border}}><input style={S.input} placeholder={'Search '+entities.length+' entities...'} value={search} onChange={e=>setSearch(e.target.value)} autoFocus/></div>
        {filtered.map(e=><div key={e.id} style={{padding:'10px 16px',cursor:'pointer',background:e.id===activeId?T.accentDim:'transparent',borderLeft:e.id===activeId?'3px solid '+T.accent:'3px solid transparent'}} onClick={()=>{onSelect(e.id);setOpen(false);setSearch('');}}>
          <span style={{fontWeight:600,color:T.textBright,fontSize:13}}>{e.code}</span><span style={{color:T.text,fontSize:12,marginLeft:10}}>{e.name}</span></div>)}
        <div style={{borderTop:'1px solid '+T.border,padding:12}}><button style={{...S.btnS,width:'100%'}} onClick={()=>{onManage();setOpen(false);}}>Manage Entities</button></div></div></>}</div>);}

// ═══ Main App — JE form state lives here so it persists across modal open/close ═══
export default function App(){
  const[user,setUser]=useState(null);const[entities,setEntities]=useState([]);const[activeEntity,setActiveEntity]=useState(null);
  const[page,setPage]=useState('dashboard');const[loading,setLoading]=useState(true);
  const[showJE,setShowJE]=useState(false);const[showChangePw,setShowChangePw]=useState(false);const[rk,setRk]=useState(0);
  const[sidebarCol,setSidebarCol]=useState(()=>{try{return localStorage.getItem(SIDEBAR_KEY)==='true';}catch{return false;}});
  // JE form state lives in App — survives modal close, cleared only on post/discard
  const[jeForm,setJeForm]=useState(BLANK_JE());const[jePendingFiles,setJePendingFiles]=useState([]);
  // Bank transaction state lifted to App so it persists across page navigation
  const[bankSelAcct,setBankSelAcct]=useState('');const[bankTxns,setBankTxns]=useState([]);const[bankUploading,setBankUploading]=useState(false);const[bankStatusFilter,setBankStatusFilter]=useState('');

  useEffect(()=>{try{localStorage.setItem(SIDEBAR_KEY,String(sidebarCol));}catch{}},[sidebarCol]);
  useEffect(()=>{const t=api.getToken();if(t){api.me().then(u=>{if(u)setUser(u);}).catch(()=>api.clearToken()).finally(()=>setLoading(false));}else setLoading(false);},[]);
  useEffect(()=>{if(user)api.getEntities().then(e=>{setEntities(e);if(e.length>0&&!activeEntity)setActiveEntity(e[0].id);});},[user]);
  const refreshEntities=useCallback(async()=>{const e=await api.getEntities();setEntities(e);return e;},[]);
  const canAccess=s=>{if(!user)return false;if(user.role==='Admin')return true;return({Accountant:['entries','reports','coa','bankrec'],Viewer:['reports']}[user.role]||[]).includes(s);};
  if(loading)return<div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:T.textMuted}}>Loading...</div></div>;
  if(!user)return<AuthScreen onLogin={setUser}/>;
  const jeHasContent=jeForm.memo||jeForm.lines.some(l=>l.account_code||l.debit||l.credit)||jePendingFiles.length>0;

  const navItems=[
    {id:'dashboard',label:'Dashboard',icon:NI.dashboard,section:'reports'},
    {id:'d1',divider:1,label:'TRANSACTIONS'},{id:'journal',label:'Journal Entries',icon:NI.journal,section:'entries'},
    {id:'d2',divider:1,label:'ACCOUNTS'},{id:'coa',label:'Chart of Accounts',icon:NI.coa,section:'coa'},{id:'ledger',label:'General Ledger',icon:NI.ledger,section:'reports'},
    {id:'d2b',divider:1,label:'BANKING'},{id:'banktxn',label:'Bank Transactions',icon:NI.banktxn,section:'bankrec'},{id:'bankrec',label:'Bank Reconciliation',icon:NI.bankrec,section:'bankrec'},
    {id:'d3',divider:1,label:'REPORTS'},{id:'trial',label:'Trial Balance',icon:NI.trial,section:'reports'},{id:'bs',label:'Balance Sheet',icon:NI.bs,section:'reports'},{id:'is',label:'Income Statement',icon:NI.is,section:'reports'},
    {id:'d4',divider:1,label:'ADMIN'},{id:'entities',label:'Entities ('+entities.length+')',icon:NI.entities,section:'all'},{id:'users',label:'Users',icon:NI.users,section:'all'},
  ];

  return(<div style={S.app}>
    <div style={S.topBar}><div style={{display:'flex',alignItems:'center',gap:16}}>
      <button style={{...S.btnGhost,fontSize:18,padding:'4px 6px',color:T.textMuted}} onClick={()=>setSidebarCol(c=>!c)}>{sidebarCol?'\u2630':'\u2190'}</button>
      <div style={{display:'flex',alignItems:'center',gap:10}}><Logo size={32}/>{!sidebarCol&&<div style={{fontSize:17,fontWeight:800,color:T.textBright}}>CloudLedger</div>}</div>
      <div style={{width:1,height:28,background:T.border}}/><EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')}/></div>
      <div style={{display:'flex',alignItems:'center',gap:10}}>
        {canAccess('entries')&&activeEntity&&<button style={{...S.btnP,position:'relative'}} onClick={()=>setShowJE(true)}>+ Journal Entry{jeHasContent&&<span style={{position:'absolute',top:-3,right:-3,width:8,height:8,borderRadius:4,background:T.orange,border:'2px solid #fff'}}/>}</button>}
        <span style={{fontSize:13,fontWeight:500}}>{user.name}</span><span style={S.badge}>{user.role}</span>
        <button style={S.btnS} onClick={()=>setShowChangePw(true)}>Settings</button>
        <button style={S.btnS} onClick={()=>{api.clearToken();setUser(null);}}>Sign Out</button></div></div>
    <div style={S.body}><div style={S.sidebar(sidebarCol)}>
      {navItems.map(n=>n.divider?(!sidebarCol?<div key={n.id} style={S.navSection(sidebarCol)}>{n.label}</div>:<div key={n.id} style={{height:8}}/>)
        :(n.section==='all'?user.role==='Admin':canAccess(n.section))?<div key={n.id} style={S.navItem(page===n.id,sidebarCol)} onClick={()=>setPage(n.id)} title={n.label}>
          {sidebarCol?<span style={{fontSize:15}}>{n.icon}</span>:<span>{n.icon}  {n.label}</span>}</div>:null)}</div>
      <div style={S.main}>{(()=>{const en=entities.find(e=>e.id===activeEntity);const entityName=en?en.name:'';return<>
        {page==='dashboard'&&<Dashboard entityId={activeEntity} key={rk}/>}
        {page==='journal'&&activeEntity&&<JournalList entityId={activeEntity} entityName={entityName} key={activeEntity+'-'+rk} onNewEntry={()=>setShowJE(true)}/>}
        {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} canEdit={canAccess('coa')}/>}
        {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity} entityName={entityName} key={activeEntity+'-'+rk}/>}
        {page==='banktxn'&&activeEntity&&<BankTransactions entityId={activeEntity} bankSelAcct={bankSelAcct} setBankSelAcct={setBankSelAcct} bankTxns={bankTxns} setBankTxns={setBankTxns} bankUploading={bankUploading} setBankUploading={setBankUploading} bankStatusFilter={bankStatusFilter} setBankStatusFilter={setBankStatusFilter}/>}
        {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user}/>}
        {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity} entityName={entityName} key={activeEntity+'-'+rk}/>}
        {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity} entityName={entityName}/>}
        {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity} entityName={entityName}/>}
        {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='users'&&<UserManagement currentUser={user}/>}
      </>})()}</div></div>
    {showJE&&activeEntity&&<JournalEntryModal entityId={activeEntity} user={user} onClose={()=>setShowJE(false)} onPosted={()=>setRk(k=>k+1)} form={jeForm} setForm={setJeForm} pendingFiles={jePendingFiles} setPendingFiles={setJePendingFiles}/>}
    {showChangePw&&<SettingsModal onClose={()=>setShowChangePw(false)} user={user} onUserUpdate={u=>setUser(u)}/>}
  </div>);}

// ═══ Dashboard ═══
function Dashboard({entityId}){const[summary,setSummary]=useState([]);useEffect(()=>{api.getSummary().then(setSummary);},[]);const curr=summary.find(e=>e.id===entityId);
  return(<div><div style={S.h1}>Dashboard</div><div style={S.sub}>{summary.length} entities under management</div>
    {curr&&<div style={S.summaryBar}>{[{l:'Total Assets',v:curr.assets,c:T.accent},{l:'Total Liabilities',v:curr.liabilities,c:T.orange},{l:'Revenue',v:curr.revenue,c:T.purple},{l:'Expenses',v:curr.expenses,c:T.red},{l:'Net Income',v:curr.net_income,c:curr.net_income>=0?T.green:T.red},{l:'Entries',v:curr.entry_count,c:T.textMuted,raw:1}].map(s=>
      <div key={s.l} style={S.summaryItem}><div style={{...S.statVal,color:s.c}}>{s.raw?s.v:'$'+fmt(s.v)}</div><div style={S.statLabel}>{s.l}</div></div>)}</div>}
    <div style={S.cardFlush}><div style={{padding:'18px 20px'}}><div style={S.h2}>All Entities</div></div><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={S.thR}>Assets</th><th style={S.thR}>Liabilities</th><th style={S.thR}>Net Income</th><th style={S.thR}>JEs</th></tr></thead>
      <tbody>{summary.sort((a,b)=>a.code.localeCompare(b.code)).map(e=><tr key={e.id} style={e.id===entityId?{background:T.accentDim}:{}}><td style={{...S.td,fontWeight:600,color:T.accent}}>{e.code}</td><td style={S.td}>{e.name}</td><td style={S.tdR}>{fmt(e.assets)}</td><td style={S.tdR}>{fmt(e.liabilities)}</td><td style={{...S.tdR,color:e.net_income>=0?T.green:T.red,fontWeight:600}}>{fmt(e.net_income)}</td><td style={S.tdR}>{e.entry_count}</td></tr>)}</tbody></table></div></div>);}

// ═══ Edit JE Modal ═══
function EditJEModal({entityId,entry,accounts:initAccounts,onClose,onSaved}){
  const[accounts,setAccounts]=useState(initAccounts||[]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[saving,setSaving]=useState(false);
  const[form,setForm]=useState({date:entry.date,memo:entry.memo,lines:entry.lines.map(l=>({account_code:l.account_code,debit:l.debit>0?String(l.debit):'',credit:l.credit>0?String(l.credit):''}))});
  useEffect(()=>{if(!initAccounts?.length)api.getAccounts(entityId).then(setAccounts);},[entityId,initAccounts]);
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const tDr=form.lines.reduce((s,l)=>s+(parseFloat(l.debit)||0),0);const tCr=form.lines.reduce((s,l)=>s+(parseFloat(l.credit)||0),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;
  const save=async()=>{if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}if(form.lines.some(l=>!l.account_code)){setErr('All lines need an account');return;}if(!bal){setErr('Must balance');return;}
    setSaving(true);setErr('');try{await api.updateEntry(entityId,entry.id,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseFloat(l.debit)||0,credit:parseFloat(l.credit)||0}))});
      onSaved();onClose();}catch(e){setErr(e.message);}finally{setSaving(false);}};
  return(<div style={S.modal}><div style={{...S.modalBox,maxWidth:960}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Edit JE-{String(entry.entry_num).padStart(4,'0')}</div>
    <div style={{background:T.bgElevated,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:18,marginBottom:16}}>
      <div style={S.row}><div style={{...S.col,maxWidth:170}}><label style={S.label}>Date</label><input style={S.input} type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/></div>
        <div style={{...S.col,flex:4}}><label style={S.label}>Memo</label><input style={S.input} value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))}/></div></div></div>
    <div style={{...S.cardFlush,marginBottom:16}}><table style={S.table}><thead><tr><th style={S.th}>Account</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.debit} onChange={e=>updateLine(i,'debit',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.credit} onChange={e=>updateLine(i,'credit',e.target.value)}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
    <div style={{display:'flex',gap:10,alignItems:'center'}}>
      <button style={S.btnS} onClick={addLine}>+ Add line</button>
      <button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New account</button>
      <div style={{flex:1}}/>
      {!bal&&tDr>0&&<span style={{fontSize:12,color:T.orange,fontWeight:600}}>Off by ${fmt(tDr-tCr)}</span>}
      {bal&&<span style={{fontSize:12,color:T.green,fontWeight:600}}>Balanced</span>}
      {err&&<span style={S.err}>{err}</span>}
      <button style={S.btnS} onClick={onClose}>Cancel</button>
      <button style={{...S.btnP,padding:'10px 28px',fontSize:14,opacity:saving?.6:1}} onClick={save} disabled={saving}>{saving?'Saving...':'Save Changes'}</button></div>
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)))}/>}
  </div></div>);}

// ═══ Journal List ═══
function JournalList({entityId,entityName,onNewEntry}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[from,setFrom]=useState('');const[to,setTo]=useState('');
  const[editEntry,setEditEntry]=useState(null);
  const load=useCallback(async()=>{const[e,a]=await Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]);setEntries(e);setAccounts(a);},[entityId,from,to]);
  useEffect(()=>{load();},[load]);const del=async id=>{if(!confirm('Delete this journal entry?'))return;await api.deleteEntry(entityId,id);load();};const acctName=code=>accounts.find(a=>a.code===code)?.name||'?';
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}><div><div style={S.h1}>Journal Entries</div><div style={S.sub}>{entityName} &middot; {entries.length} entries</div></div><button style={S.btnP} onClick={onNewEntry}>+ New Entry</button></div>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      {(from||to)&&<button style={{...S.btnGhost,marginTop:14,color:T.red}} onClick={()=>{setFrom('');setTo('');}}>Clear</button>}</div>
    {entries.length===0?<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No entries found</div>:
      entries.map(e=><div key={e.id} style={{...S.card,padding:18}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
          <div style={{display:'flex',alignItems:'center',gap:14}}><span style={{fontWeight:700,color:T.accent,fontSize:14}}>JE-{String(e.entry_num).padStart(4,'0')}</span>
            <span style={{color:T.textMuted}}>{e.date}</span><span style={{fontWeight:500}}>{e.memo}</span>
            {e.attachments?.length>0&&<span style={{fontSize:11,color:T.teal,fontWeight:500}}>({e.attachments.length} file{e.attachments.length>1?'s':''})</span>}</div>
          <div style={{display:'flex',alignItems:'center',gap:8}}>
            <span style={{fontSize:12,color:T.textDim}}>{e.created_by}</span>
            <button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>setEditEntry(e)}>Edit</button>
            <button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={()=>del(e.id)}>Delete</button></div></div>
        <table style={S.table}><thead><tr><th style={S.th}>Account</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
          <tbody>{e.lines.map((l,i)=><tr key={i}><td style={{...S.td,paddingLeft:l.credit>0&&l.debit===0?28:14}}>{acctLabel(l.account_code,acctName(l.account_code))}</td>
            <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td><td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td></tr>)}</tbody></table>
        {e.attachments?.length>0&&<div style={{marginTop:10,display:'flex',flexWrap:'wrap',gap:4}}>{e.attachments.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>)}</div>}
      </div>)}
    {editEntry&&<EditJEModal entityId={entityId} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={load}/>}
  </div>);}

// ═══ Chart of Accounts ═══
function ChartOfAccounts({entityId,canEdit}){const[accounts,setAccounts]=useState([]);const[showAdd,setShowAdd]=useState(false);
  const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState({});const[editErr,setEditErr]=useState('');
  const load=useCallback(async()=>{setAccounts(await api.getAccounts(entityId));},[entityId]);useEffect(()=>{load();},[load]);
  const startEdit=a=>{setEditing(a.code);setEditForm({new_code:a.code,name:a.name,type:a.type,subtype:a.subtype||'',bank_acct:!!a.bank_acct});setEditErr('');};
  const saveEdit=async()=>{if(!editForm.new_code||!editForm.name){setEditErr('Code and name required');return;}
    try{await api.updateAccount(entityId,editing,editForm);setEditing(null);load();}catch(e){setEditErr(e.message);}};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}><div><div style={S.h1}>Chart of Accounts</div><div style={S.sub}>{accounts.length} accounts</div></div>
    {canEdit&&<button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add Account'}</button>}</div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}><div style={S.row}>
      <div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
      <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createAccount(entityId,form);setForm({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});setShowAdd(false);setErr('');load();}catch(e){setErr(e.message);}}}>Add Account</button></div>}
    {editing&&<div style={{...S.card,borderColor:T.accent+'40',marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:12}}>Edit Account: {editing}</div>
      <div style={S.row}>
        <div style={S.col}><label style={S.label}>Account Code</label><input style={S.input} value={editForm.new_code} onChange={e=>setEditForm(f=>({...f,new_code:e.target.value}))}/></div>
        <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))}/></div>
        <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={editForm.type} onChange={e=>setEditForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
      <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={editForm.bank_acct} onChange={e=>setEditForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
      {editErr&&<div style={S.err}>{editErr}</div>}
      {editForm.new_code!==editing&&<div style={{fontSize:11,color:T.orange,marginBottom:8}}>Changing code from {editing} to {editForm.new_code} will update all journal entries, bank transactions, and reconciliations.</div>}
      <div style={{display:'flex',gap:10}}><button style={S.btnP} onClick={saveEdit}>Save Changes</button><button style={S.btnS} onClick={()=>setEditing(null)}>Cancel</button></div></div>}
    <div style={S.cardFlush}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Name</th><th style={S.th}>Type</th><th style={S.thC}>Bank</th>{canEdit&&<th style={{...S.th,width:80}}>Actions</th>}</tr></thead>
      <tbody>{accounts.map(a=><tr key={a.code} style={editing===a.code?{background:T.accentDim}:{}}>
        <td style={{...S.td,fontWeight:600,color:T.textBright}}>{a.code}</td><td style={S.td}>{a.name}</td><td style={S.td}><span style={S.tag(a.type)}>{a.type}</span></td>
        <td style={S.tdC}>{a.bank_acct?<span style={{color:T.green}}>Yes</span>:''}</td>
        {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
          <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(a)}>Edit</button>
          <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={async()=>{try{await api.deleteAccount(entityId,a.code);load();}catch(e){alert(e.message);}}}>x</button></div></td>}</tr>)}</tbody></table></div></div>);}

// ═══ General Ledger ═══
function GeneralLedger({entityId,entityName}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[filter,setFilter]=useState('');const[from,setFrom]=useState(fy_start());const[to,setTo]=useState(today());
  useEffect(()=>{Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]).then(([e,a])=>{setEntries(e);setAccounts(a);});},[entityId,from,to]);
  const filtered=accounts.filter(a=>!filter||a.code===filter).sort((a,b)=>a.code.localeCompare(b.code));
  const entryAtts={};entries.forEach(e=>{if(e.attachments?.length>0)entryAtts[e.id]=e.attachments;});
  const doExport=()=>{const rows=[[entityName||'General Ledger'],['General Ledger'],['Period: '+(from||'Begin')+' to '+(to||today())],[]];filtered.forEach(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({date:e.date,je:'JE-'+String(e.entry_num).padStart(4,'0'),memo:e.memo,debit:l.debit,credit:l.credit});});});if(txns.length===0&&!filter)return;rows.push([acctLabel(acct.code,acct.name)]);rows.push(['Date','JE','Memo','Debit','Credit','Balance']);let run=0;const isDr=acct.type==='Asset'||acct.type==='Expense';txns.sort((a,b)=>a.date.localeCompare(b.date)).forEach(t=>{run+=isDr?(t.debit-t.credit):(t.credit-t.debit);rows.push([t.date,t.je,t.memo,t.debit||'',t.credit||'',run]);});rows.push([]);});exportToExcel(rows,'GL.xlsx');};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}><div><div style={S.h1}>General Ledger</div>{entityName&&<div style={{fontSize:13,color:T.textMuted}}>{entityName}</div>}</div><button style={S.btnExport} onClick={doExport}>Export Excel</button></div><div style={S.sub}/>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <div style={{maxWidth:280}}><label style={S.label}>Account</label><select style={{...S.inputSm,width:'100%'}} value={filter} onChange={e=>setFilter(e.target.value)}><option value="">All accounts</option>{accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div></div>
    {filtered.map(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({...l,date:e.date,memo:e.memo,jeNum:e.entry_num,jeId:e.id});});});
      if(txns.length===0&&!filter)return null;txns.sort((a,b)=>a.date.localeCompare(b.date));let run=0;const dr=acct.type==='Asset'||acct.type==='Expense';
      return(<div key={acct.code} style={S.cardFlush}><div style={{padding:'14px 20px',display:'flex',alignItems:'center',gap:10,borderBottom:'1px solid '+T.border}}>
        <span style={{fontWeight:700,color:T.textBright,fontSize:14}}>{acct.code}</span><span>{acct.name}</span><span style={S.tag(acct.type)}>{acct.type}</span></div>
        {txns.length===0?<div style={{padding:20,color:T.textDim}}>No transactions</div>:
        <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th><th style={{...S.th,width:100}}>Docs</th></tr></thead>
          <tbody>{txns.map((t,i)=>{run+=dr?(t.debit-t.credit):(t.credit-t.debit);const atts=entryAtts[t.jeId];return<tr key={i}><td style={{...S.td,color:T.textMuted}}>{t.date}</td><td style={S.td}><span style={{color:T.accent,fontWeight:600}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td><td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(run)}</td>
            <td style={S.td}>{atts?atts.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>):''}</td></tr>;})}</tbody></table>}</div>);})}</div>);}

// ═══ Bank Transactions (state lifted to App for navigation persistence) ═══
function BankTransactions({entityId,bankSelAcct:selAcct,setBankSelAcct:setSelAcct,bankTxns:txns,setBankTxns:setTxns,bankUploading:uploading,setBankUploading:setUploading,bankStatusFilter:statusFilter,setBankStatusFilter:setStatusFilter}){
  const[accounts,setAccounts]=useState([]);const[bankAccts,setBankAccts]=useState([]);
  const[err,setErr]=useState('');const[msg,setMsg]=useState('');const[showAddAcct,setShowAddAcct]=useState(false);
  const[uploadProgress,setUploadProgress]=useState('');const[lastBatchId,setLastBatchId]=useState(null);

  const loadAccounts=useCallback(async()=>{const a=await api.getAccounts(entityId);setAccounts(a);setBankAccts(a.filter(x=>x.bank_acct||(['cash','bank','checking','savings'].some(w=>x.name.toLowerCase().includes(w))&&x.type==='Asset')));return a;},[entityId]);
  const loadTxns=useCallback(async(acct,status)=>{if(!acct)return;const t=await api.getBankTransactions(entityId,acct,status||undefined);setTxns(t);},[entityId,setTxns]);
  useEffect(()=>{loadAccounts();},[loadAccounts]);
  useEffect(()=>{if(selAcct&&txns.length===0)loadTxns(selAcct,statusFilter);},[selAcct]);
  const reload=()=>loadTxns(selAcct,statusFilter);

  const onFileSelected=async e=>{const file=e.target.files[0];if(!file||!selAcct)return;e.target.value='';setErr('');setMsg('');setUploading(true);setUploadProgress('Uploading file...');
    try{const r=await api.uploadBankTransactions(entityId,selAcct,file);
      setLastBatchId(r.batch_id);
      setUploadProgress('Auto-categorizing '+r.count+' transactions...');
      const imported=await api.getBankTransactions(entityId,selAcct,'pending');let auto=0;
      for(const t of imported){if(!t.account_code){const sg=suggestAccount(t.description,accounts,selAcct);if(sg){await api.codeBankTransaction(entityId,t.id,sg.code,t.memo||t.description);auto++;}}}
      setMsg(r.count+' imported'+(auto>0?', '+auto+' auto-categorized':''));loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}finally{setUploading(false);setUploadProgress('');}};
  const cancelUpload=()=>{setUploading(false);setUploadProgress('');setMsg('Upload cancelled');};
  const deleteBatch=async()=>{if(!lastBatchId)return;if(!confirm('Delete all unposted transactions from the last upload?'))return;
    try{const r=await api.deleteBankBatch(entityId,lastBatchId);setMsg(r.deleted+' transactions removed');setLastBatchId(null);loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}};
  const codeTransaction=async(id,acct_code,memo)=>{await api.codeBankTransaction(entityId,id,acct_code,memo);
    setTxns(prev=>prev.map(t=>t.id===id?{...t,account_code:acct_code,memo:memo,status:acct_code?'coded':'pending'}:t));};
  const postCoded=async()=>{const ids=txns.filter(t=>t.status==='coded').map(t=>t.id);if(!ids.length){setErr('Nothing coded');return;}try{const r=await api.postBankTransactions(entityId,ids);setMsg(r.posted+' JEs created');loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}};
  const changeAcct=v=>{setSelAcct(v);setTxns([]);setLastBatchId(null);if(v)loadTxns(v,statusFilter);};
  const changeStatus=v=>{setStatusFilter(v);if(selAcct)loadTxns(selAcct,v);};

  const filteredTxns=txns;
  const totalIn=filteredTxns.filter(t=>t.amount>0).reduce((s,t)=>s+t.amount,0);const totalOut=filteredTxns.filter(t=>t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const uncat=filteredTxns.filter(t=>t.status==='pending').length;

  return(<div><div style={S.h1}>Bank Transactions</div><div style={S.sub}>Upload, categorize, and post bank activity to the general ledger</div>
    <div style={S.card}><div style={S.row}>
      <div style={{...S.col,flex:2}}><label style={S.label}>Bank Account</label><select style={S.select} value={selAcct} onChange={e=>changeAcct(e.target.value)}><option value="">Select bank account...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Status</label><select style={S.select} value={statusFilter} onChange={e=>changeStatus(e.target.value)}><option value="">All</option><option value="pending">Pending</option><option value="coded">Coded</option><option value="posted">Posted</option></select></div>
      {selAcct&&<div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
        {uploading
          ?<div style={{display:'flex',alignItems:'center',gap:10,marginBottom:4}}>
            <div style={{width:16,height:16,border:'2px solid '+T.accent,borderTopColor:'transparent',borderRadius:'50%',animation:'spin 1s linear infinite'}}/>
            <span style={{fontSize:12,color:T.accent,fontWeight:500}}>{uploadProgress||'Processing...'}</span>
            <button style={{...S.btnD,padding:'7px 16px',fontSize:12}} onClick={cancelUpload}>Cancel</button></div>
          :<div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
            <button style={{...S.btnP,pointerEvents:'none'}}>Upload CSV / Excel</button>
            <input type="file" accept=".csv,.xlsx,.xls" style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:'pointer'}} onChange={onFileSelected}/></div>}
      </div>}
    </div>
    {err&&<div style={S.err}>{err}</div>}{msg&&<div style={S.success}>{msg}</div>}
    {lastBatchId&&<div style={{marginTop:10,display:'flex',alignItems:'center',gap:10}}>
      <span style={{fontSize:12,color:T.textMuted}}>Last upload batch loaded</span>
      <button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={deleteBatch}>Delete uploaded batch</button></div>}
    </div>
    <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    {selAcct&&filteredTxns.length>0&&<div style={S.summaryBar}>
      <div style={S.summaryItem}><div style={{...S.statVal,fontSize:20,color:T.textBright}}>{filteredTxns.length}</div><div style={S.statLabel}>Transactions</div></div>
      <div style={S.summaryItem}><div style={{...S.statVal,fontSize:20,color:T.orange}}>{uncat}</div><div style={S.statLabel}>Uncategorized</div></div>
      <div style={S.summaryItem}><div style={{...S.statVal,fontSize:20,color:T.green}}>${fmt(totalIn)}</div><div style={S.statLabel}>Total Inflows</div></div>
      <div style={S.summaryItem}><div style={{...S.statVal,fontSize:20,color:T.red}}>${fmt(totalOut)}</div><div style={S.statLabel}>Total Outflows</div></div>
      <div style={S.summaryItem}><div style={{...S.statVal,fontSize:20,color:T.textBright}}>${fmt(totalIn-totalOut)}</div><div style={S.statLabel}>Net</div></div></div>}
    {selAcct&&filteredTxns.length>0&&<div style={S.cardFlush}>
      <div style={{padding:'16px 20px',display:'flex',justifyContent:'space-between',alignItems:'center',borderBottom:'1px solid '+T.border}}>
        <div style={S.h2}>{filteredTxns.length} Transactions</div>
        <div style={{display:'flex',gap:10}}>
          <button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New Account</button>
          {filteredTxns.some(t=>t.status==='coded')&&<button style={S.btnP} onClick={postCoded}>Post {filteredTxns.filter(t=>t.status==='coded').length} to GL</button>}</div></div>
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Description</th><th style={S.thR}>Amount</th><th style={{...S.th,width:240}}>GL Account</th><th style={{...S.th,width:160}}>Memo</th><th style={{...S.th,width:70}}>Status</th><th style={{...S.th,width:36}}></th></tr></thead>
        <tbody>{filteredTxns.map(t=><tr key={t.id} style={t.status==='posted'?{opacity:0.45}:{}}>
          <td style={{...S.td,color:T.textMuted,fontSize:12,whiteSpace:'nowrap'}}>{t.date}</td>
          <td style={{...S.td,maxWidth:240,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontWeight:500}}>{t.description}</td>
          <td style={{...S.tdR,fontSize:15,fontWeight:700,color:t.amount>=0?T.green:T.red}}>{t.amount>=0?'+':''}{fmt(t.amount)}</td>
          <td style={{...S.td,padding:'4px 6px'}}>{t.status==='posted'?<span style={{fontSize:12,color:T.textDim}}>{t.account_code}</span>:
            <AccountAutocomplete accounts={accounts} value={t.account_code||''} exclude={selAcct} onChange={v=>codeTransaction(t.id,v,t.memo)} placeholder="Search GL account..."/>}</td>
          <td style={{...S.td,padding:'4px 6px'}}>{t.status==='posted'?<span style={{fontSize:12,color:T.textDim}}>{t.memo}</span>:
            <input style={S.inputSm} placeholder="Memo" value={t.memo||''} onChange={e=>{const v=e.target.value;setTxns(prev=>prev.map(x=>x.id===t.id?{...x,memo:v}:x));}} onBlur={()=>codeTransaction(t.id,t.account_code,t.memo)}/>}</td>
          <td style={S.td}><span style={{fontSize:10,fontWeight:600,padding:'3px 8px',borderRadius:20,background:t.status==='posted'?T.greenDim:t.status==='coded'?T.accentDim:T.orangeDim,color:t.status==='posted'?T.green:t.status==='coded'?T.accent:T.orange}}>{t.status}</span></td>
          <td style={S.td}>{t.status!=='posted'&&<button style={S.btnGhost} onClick={async()=>{await api.deleteBankTransaction(entityId,t.id);setTxns(prev=>prev.filter(x=>x.id!==t.id));}}>x</button>}</td>
        </tr>)}</tbody></table></div>}
    {selAcct&&filteredTxns.length===0&&!uploading&&<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No transactions yet. Upload a bank statement above.</div>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>{setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));if(a.bank_acct)setBankAccts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));}}/>}
  </div>);}

// ═══ Reports ═══
function TrialBalance({entityId,entityName}){const[balances,setBalances]=useState([]);const[asOf,setAsOf]=useState(today());const fyS=asOf.slice(0,4)+'-01-01';
  useEffect(()=>{api.getBalances(entityId,{as_of:asOf,close_pl_before:fyS}).then(setBalances);},[entityId,asOf,fyS]);
  let tDr=0,tCr=0;const rows=balances.filter(b=>Math.abs(b.balance)>0.005).map(b=>{const isDr=b.type==='Asset'||b.type==='Expense';const dr=(isDr&&b.balance>0)||(!isDr&&b.balance<0)?Math.abs(b.balance):0;const cr=(isDr&&b.balance<0)||(!isDr&&b.balance>0)?Math.abs(b.balance):0;tDr+=dr;tCr+=cr;return{...b,dr,cr};});
  const doExport=()=>{const d=[[entityName||'Trial Balance'],['Trial Balance'],['As of '+asOf],[],['Code','Account','Type','Debit','Credit']];rows.forEach(r=>d.push([r.code,r.name,r.type,r.dr||'',r.cr||'']));d.push([]);d.push(['','','Total',tDr,tCr]);exportToExcel(d,'TB_'+asOf+'.xlsx');};
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Trial Balance</div><div style={{fontSize:13,color:T.textMuted}}>As of {asOf}</div></div>
    <table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={{...S.td,fontWeight:600,color:T.textBright}}>{r.code}</td><td style={S.td}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td><td style={S.tdR}>{r.dr>0?fmt(r.dr):''}</td><td style={S.tdR}>{r.cr>0?fmt(r.cr):''}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(tDr-tCr)<0.005?T.green:T.red}}>{Math.abs(tDr-tCr)<0.005?'In balance':'Off by $'+fmt(tDr-tCr)}</div></div></div>);}

function BalanceSheet({entityId,entityName}){const[balances,setBalances]=useState([]);const[asOf,setAsOf]=useState(today());const fyS=asOf.slice(0,4)+'-01-01';
  useEffect(()=>{api.getBalances(entityId,{as_of:asOf,close_pl_before:fyS}).then(setBalances);},[entityId,asOf,fyS]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=t=>get(t).reduce((s,b)=>s+b.balance,0);
  const ni=sum('Revenue')-sum('Expense');const tA=sum('Asset');const tLE=sum('Liability')+sum('Equity')+ni;
  const doExport=()=>{const d=[[entityName||'Balance Sheet'],['Balance Sheet'],['As of '+asOf],[]];[['Assets','Asset'],['Liabilities','Liability'],['Equity','Equity']].forEach(([t,ty])=>{d.push([t,'']);get(ty).forEach(b=>d.push(['  '+b.name,b.balance]));if(ty==='Equity'&&Math.abs(ni)>0.005)d.push(['  Net Income (current period)',ni]);d.push(['Total '+t,ty==='Equity'?sum(ty)+ni:sum(ty)]);d.push([]);});d.push(['Total L+E',tLE]);exportToExcel(d,'BS_'+asOf+'.xlsx');};
  const Sec=({title,type,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{get(type).map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+T.borderLight}}>{fmt(b.balance)}</td></tr>)}
    {type==='Equity'&&Math.abs(ni)>0.005&&<tr><td style={{...S.indentTd,fontStyle:'italic',color:T.textMuted}}>Net Income (current period)</td><td style={{...S.tdR,fontStyle:'italic'}}>{fmt(ni)}</td></tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>${fmt(total)}</td></tr></>);
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Balance Sheet</div><div style={{fontSize:13,color:T.textMuted}}>As of {asOf}</div></div>
    <table style={{...S.table,maxWidth:580,margin:'0 auto'}}><tbody><Sec title="Assets" type="Asset" total={tA}/><tr><td colSpan={2} style={{padding:8}}/></tr>
      <Sec title="Liabilities" type="Liability" total={sum('Liability')}/><tr><td colSpan={2} style={{padding:4}}/></tr><Sec title="Equity" type="Equity" total={sum('Equity')+ni}/>
      <tr style={S.grandTotalRow}><td style={S.tdBold}>Total Liabilities + Equity</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tLE)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(tA-tLE)<0.005?T.green:T.red}}>{Math.abs(tA-tLE)<0.005?'A = L + E':'Off by $'+fmt(tA-tLE)}</div></div></div>);}

function IncomeStatement({entityId,entityName}){const[balances,setBalances]=useState([]);const[from,setFrom]=useState(fy_start());const[to,setTo]=useState(today());
  useEffect(()=>{api.getBalances(entityId,{from,to}).then(setBalances);},[entityId,from,to]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=arr=>arr.reduce((s,b)=>s+b.balance,0);
  const rev=get('Revenue');const cogs=get('Expense').filter(b=>b.subtype==='COGS');const opex=get('Expense').filter(b=>b.subtype==='Operating Expense');const other=get('Expense').filter(b=>b.subtype!=='COGS'&&b.subtype!=='Operating Expense');
  const tRev=sum(rev);const gp=tRev-sum(cogs);const oi=gp-sum(opex);const ni=oi-sum(other);
  const doExport=()=>{const d=[[entityName||'Income Statement'],['Income Statement'],['Period: '+from+' to '+to],[]];[['Revenue',rev],['COGS',cogs],['Operating Expenses',opex],['Other',other]].forEach(([t,items])=>{if(!items.length)return;d.push([t,'']);items.forEach(b=>d.push(['  '+b.name,b.balance]));d.push(['Total '+t,sum(items)]);d.push([]);});d.push(['Net Income',ni]);exportToExcel(d,'IS_'+from+'_'+to+'.xlsx');};
  const Sec=({title,items,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{items.map(b=><tr key={b.code}><td style={S.indentTd}>{b.name}</td><td style={{...S.tdR,borderBottom:'1px solid '+T.borderLight}}>{fmt(b.balance)}</td></tr>)}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>${fmt(total)}</td></tr></>);
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Income Statement</div><div style={{fontSize:13,color:T.textMuted}}>Period: {from} to {to}</div></div>
    <table style={{...S.table,maxWidth:580,margin:'0 auto'}}><tbody><Sec title="Revenue" items={rev} total={tRev}/>
      {cogs.length>0&&<><Sec title="Cost of Goods Sold" items={cogs} total={sum(cogs)}/><tr style={{background:T.bgElevated}}><td style={{...S.td,fontWeight:700,color:T.textBright}}>Gross Profit</td><td style={{...S.tdR,fontWeight:700,color:T.textBright,fontSize:15}}>${fmt(gp)}</td></tr></>}
      <Sec title="Operating Expenses" items={opex} total={sum(opex)}/>
      <tr style={{background:T.bgElevated}}><td style={{...S.td,fontWeight:700,color:T.textBright}}>Operating Income</td><td style={{...S.tdR,fontWeight:700,color:T.textBright,fontSize:15}}>${fmt(oi)}</td></tr>
      {other.length>0&&<Sec title="Other Expenses" items={other} total={sum(other)}/>}
      <tr style={S.grandTotalRow}><td style={{...S.tdBold,fontSize:15}}>Net Income</td><td style={{...S.tdBold,textAlign:'right',fontSize:18,color:ni>=0?T.green:T.red}}>${fmt(ni)}</td></tr></tbody></table></div></div>);}

// ═══ Bank Reconciliation ═══
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
  if(view==='new')return(<div><button style={{...S.btnS,marginBottom:20}} onClick={()=>{setView('list');setSelAcct('');setChecked({});}}>&larr; Back</button>
    <div style={S.h1}>New Bank Reconciliation</div><div style={S.card}><div style={S.row}>
      <div style={{...S.col,flex:2}}><label style={S.label}>Account</label><select style={S.select} value={selAcct} onChange={e=>{setSelAcct(e.target.value);setChecked({});}}><option value="">Select...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Statement Date</label><input style={S.input} type="date" value={stmtDate} onChange={e=>setStmtDate(e.target.value)}/></div>
      <div style={S.col}><label style={S.label}>Ending Balance</label><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={stmtBal} onChange={e=>setStmtBal(e.target.value)}/></div></div></div>
    {selAcct&&<><div style={S.summaryBar}>{[{l:'Book Balance',v:bookBal,c:T.textBright},{l:'Statement',v:stmtNum,c:T.textBright},{l:'Out. Deposits',v:outDep,c:T.green},{l:'Out. Payments',v:outPay,c:T.red},{l:'Adjusted Bank',v:stmtNum+outDep+outPay,c:T.accent},{l:'Difference',v:diff,c:isRec?T.green:T.red}].map(s=>
      <div key={s.l} style={{...S.summaryItem,border:s.l==='Difference'&&isRec?'1px solid '+T.greenBorder:undefined,background:s.l==='Difference'&&isRec?T.greenDim:undefined}}>
        <div style={{...S.statVal,fontSize:18,color:s.c}}>${fmt(s.v)}</div><div style={S.statLabel}>{s.l}</div></div>)}</div>
      <div style={S.cardFlush}><div style={{padding:'14px 20px',display:'flex',justifyContent:'space-between',borderBottom:'1px solid '+T.border}}>
        <div style={S.h2}>Uncleared ({uncl.length})</div><div style={{display:'flex',gap:8}}>
          <button style={{...S.btnS,padding:'6px 14px',fontSize:11}} onClick={()=>{const nc={};uncl.forEach(t=>{nc[t.key]=true;});setChecked(nc);}}>All</button>
          <button style={{...S.btnS,padding:'6px 14px',fontSize:11}} onClick={()=>setChecked({})}>None</button></div></div>
        {uncl.length===0?<div style={{padding:30,textAlign:'center',color:T.textDim}}>All cleared</div>:
        <table style={S.table}><thead><tr><th style={S.thC} width={40}>Clr</th><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Amount</th></tr></thead>
          <tbody>{uncl.map(t=><tr key={t.key} style={{cursor:'pointer',background:checked[t.key]?T.greenDim:'transparent'}} onClick={()=>setChecked(p=>({...p,[t.key]:!p[t.key]}))}>
            <td style={S.tdC}><input type="checkbox" style={S.checkbox} checked={!!checked[t.key]} readOnly/></td><td style={{...S.td,color:T.textMuted}}>{t.date}</td><td style={S.td}><span style={{color:T.accent,fontWeight:600}}>JE-{String(t.jeNum).padStart(4,'0')}</span></td><td style={S.td}>{t.memo}</td>
            <td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:700,color:t.amount>=0?T.green:T.red}}>{fmt(t.amount)}</td></tr>)}</tbody></table>}</div>
      <div style={{display:'flex',justifyContent:'flex-end',marginTop:16}}><button style={{...S.btnP,padding:'10px 28px',fontSize:14,opacity:isRec?1:.5,cursor:isRec?'pointer':'not-allowed'}} onClick={finalize}>{isRec?'Finalize Reconciliation':'Difference must be $0.00'}</button></div></>}</div>);
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
    <div><div style={S.h1}>Bank Reconciliation</div><div style={S.sub}>{recs.length} completed</div></div><button style={S.btnP} onClick={()=>setView('new')}>+ New Reconciliation</button></div>
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(240px,1fr))',gap:16,marginBottom:20}}>
      {bankAccts.map(a=>{const t=getTxns(a.code);const bal=t.reduce((s,x)=>s+x.amount,0);return<div key={a.code} style={{...S.card,padding:20}}>
        <div style={{fontWeight:700,color:T.textBright,fontSize:14,marginBottom:4}}>{a.name}</div><div style={{fontSize:12,color:T.textDim,marginBottom:12}}>{a.code}</div>
        <div style={{fontSize:24,fontWeight:700,color:T.textBright}}>${fmt(bal)}</div></div>;})}</div>
    <div style={S.cardFlush}><div style={{padding:'16px 20px'}}><div style={S.h2}>History</div></div>{recs.length===0?<div style={{padding:40,textAlign:'center',color:T.textDim}}>No reconciliations yet</div>:
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.thR}>Statement</th><th style={S.thR}>Book</th><th style={S.thR}>Cleared</th><th style={S.th}>By</th></tr></thead>
        <tbody>{recs.map(r=><tr key={r.id}><td style={S.td}>{r.statement_date}</td><td style={S.td}>{r.account_code}</td><td style={S.tdR}>${fmt(r.statement_balance)}</td><td style={S.tdR}>${fmt(r.book_balance)}</td><td style={S.tdR}>{r.cleared_count}</td><td style={S.td}>{r.completed_by}</td></tr>)}</tbody></table>}</div></div>);}

// ═══ Entity Management ═══
function EntityManagement({refresh,entities,activeEntity,setActiveEntity}){const[showAdd,setShowAdd]=useState(false);const[bulk,setBulk]=useState(false);const[form,setForm]=useState({code:'',name:''});const[bulkText,setBulkText]=useState('');const[err,setErr]=useState('');
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
    <div><div style={S.h1}>Entity Management</div><div style={S.sub}>{entities.length} entities</div></div>
    <div style={{display:'flex',gap:10}}><button style={S.btnS} onClick={()=>{setBulk(!bulk);setShowAdd(false);}}>{bulk?'Cancel':'Bulk Import'}</button><button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setBulk(false);}}>{showAdd?'Cancel':'+ Add Entity'}</button></div></div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}><div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div><div style={{...S.col,flex:3}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createEntity(form.code,form.name);setForm({code:'',name:''});setShowAdd(false);setErr('');refresh();}catch(e){setErr(e.message);}}}>Create Entity</button></div>}
    {bulk&&<div style={{...S.card,borderColor:T.accent+'40'}}><div style={{...S.h2,marginBottom:8}}>Bulk Import</div><div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>One per line: CODE, Entity Name</div>
      <textarea style={{...S.input,height:160,fontFamily:'monospace',fontSize:12,resize:'vertical'}} value={bulkText} onChange={e=>setBulkText(e.target.value)}/>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:10}} onClick={async()=>{const ents=bulkText.split('\n').map(l=>l.trim()).filter(Boolean).map(l=>{const[c,...r]=l.split(',').map(p=>p.trim());return{code:c,name:r.join(',')};}).filter(e=>e.code&&e.name);if(!ents.length){setErr('None');return;}try{await api.bulkCreateEntities(ents);setBulkText('');setBulk(false);refresh();}catch(e){setErr(e.message);}}}>Import</button></div>}
    <div style={S.cardFlush}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Entity</th><th style={{...S.th,width:140}}>Actions</th></tr></thead>
      <tbody>{entities.sort((a,b)=>a.code.localeCompare(b.code)).map(e=><tr key={e.id} style={e.id===activeEntity?{background:T.accentDim}:{}}><td style={{...S.td,fontWeight:700,color:T.accent}}>{e.code}</td><td style={{...S.td,fontWeight:500}}>{e.name}</td>
        <td style={S.td}><div style={{display:'flex',gap:8}}><button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>setActiveEntity(e.id)}>Select</button><button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={async()=>{if(!confirm('Delete?'))return;await api.deleteEntity(e.id);const r=await refresh();if(activeEntity===e.id)setActiveEntity(r[0]?.id||null);}}>Delete</button></div></td></tr>)}</tbody></table></div></div>);}

// ═══ User Management (with role editing) ═══
function UserManagement({currentUser}){
  const[users,setUsers]=useState([]);const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({name:'',email:'',password:'',role:'Viewer'});const[err,setErr]=useState('');const[loadErr,setLoadErr]=useState('');
  const[resetId,setResetId]=useState(null);const[resetPw,setResetPw]=useState('');const[resetMsg,setResetMsg]=useState('');
  const[editingRole,setEditingRole]=useState(null);
  const loadUsers=useCallback(()=>{api.getUsers().then(setUsers).catch(e=>setLoadErr(e.message));},[]);
  useEffect(()=>{loadUsers();},[loadUsers]);
  const changeRole=async(userId,newRole)=>{try{await api.updateUser(userId,{name:users.find(u=>u.id===userId)?.name,role:newRole});setEditingRole(null);loadUsers();}catch(e){alert(e.message);}};

  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
    <div><div style={S.h1}>User Management</div><div style={S.sub}>{users.length} registered users</div></div>
    <button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add User'}</button></div>
    {loadErr&&<div style={{...S.card,borderColor:T.red+'40',color:T.red}}>Failed to load users: {loadErr}</div>}
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:14}}>Create New User</div>
      <div style={S.row}>
        <div style={S.col}><label style={S.label}>Full Name</label><input style={S.input} placeholder="e.g. Jane Smith" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
        <div style={S.col}><label style={S.label}>Login Email</label><input style={S.input} type="email" placeholder="e.g. jane@company.com" value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></div>
        <div style={S.col}><label style={S.label}>Password</label><input style={S.input} type="password" placeholder="Min 3 characters" value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))}/></div>
        <div style={S.col}><label style={S.label}>Role</label><select style={S.select} value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))}><option>Admin</option><option>Accountant</option><option>Viewer</option></select></div></div>
      <div style={{fontSize:11,color:T.textMuted,marginBottom:10}}>This email and password will be used to sign in from any device.</div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.name||!form.email||!form.password){setErr('All fields required');return;}try{await api.signup(form.name,form.email,form.password,form.role);setForm({name:'',email:'',password:'',role:'Viewer'});setShowAdd(false);setErr('');loadUsers();}catch(e){setErr(e.message);}}}>Create User</button></div>}
    <div style={S.cardFlush}>
      <table style={S.table}><thead><tr>
        <th style={S.th}>Name</th>
        <th style={S.th}>Login Email</th>
        <th style={S.th}>Role</th>
        <th style={{...S.th,width:240}}>Actions</th></tr></thead>
      <tbody>{users.length===0&&!loadErr?<tr><td colSpan={4} style={{...S.td,textAlign:'center',padding:40,color:T.textDim}}>No users found</td></tr>:
        users.map(u=><tr key={u.id}>
          <td style={{...S.td,fontWeight:600,color:T.textBright}}>{u.name}{u.id===currentUser.id?<span style={{color:T.accent,fontSize:10,marginLeft:8,fontWeight:500}}>(you)</span>:''}</td>
          <td style={{...S.td,fontFamily:'monospace',fontSize:12,color:T.textMuted}}>{u.email}</td>
          <td style={S.td}>{editingRole===u.id?
            <select style={S.selectSm} value={u.role} onChange={e=>changeRole(u.id,e.target.value)} onBlur={()=>setEditingRole(null)} autoFocus><option>Admin</option><option>Accountant</option><option>Viewer</option></select>
            :<div style={{display:'flex',alignItems:'center',gap:6}}><span style={S.badge}>{u.role}</span>
              {u.id!==currentUser.id&&<button style={{...S.btnGhost,fontSize:10,color:T.accent}} onClick={()=>setEditingRole(u.id)}>Edit</button>}</div>}</td>
          <td style={S.td}><div style={{display:'flex',gap:8}}>
            {u.id!==currentUser.id&&<button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>{setResetId(u.id);setResetPw('');setResetMsg('');}}>Reset PW</button>}
            {u.id!==currentUser.id&&<button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={async()=>{if(!confirm('Delete user '+u.name+'?'))return;await api.deleteUser(u.id);loadUsers();}}>Delete</button>}</div></td>
        </tr>)}</tbody></table></div>
    {resetId&&<div style={S.modal} onClick={()=>setResetId(null)}><div style={{...S.modalBox,maxWidth:400,textAlign:'center'}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setResetId(null)}>&times;</button><div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Reset Password</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:6}}>User: <strong style={{color:T.textBright}}>{users.find(u=>u.id===resetId)?.name}</strong></div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:16,fontFamily:'monospace'}}>{users.find(u=>u.id===resetId)?.email}</div>
      <input style={S.input} type="password" placeholder="New password" value={resetPw} onChange={e=>{setResetPw(e.target.value);setResetMsg('');}}/>
      {resetMsg&&<div style={{fontSize:12,marginTop:8,color:resetMsg.includes('!')?T.green:T.red}}>{resetMsg}</div>}
      <button style={{...S.btnP,width:'100%',padding:11,marginTop:12}} onClick={async()=>{if(resetPw.length<3){setResetMsg('Min 3 chars');return;}try{await api.adminResetPassword(resetId,resetPw);setResetMsg('Password reset!');setTimeout(()=>setResetId(null),1500);}catch(e){setResetMsg(e.message);}}}>Reset Password</button>
    </div></div>}</div>);}
