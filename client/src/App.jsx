import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { api } from './api';
import * as XLSX from 'xlsx';

const fmt = n => { const v = Math.abs(n); const s = v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return n < 0 ? '(' + s + ')' : s; };
const fmtAmt = (raw) => {
  if (raw === '' || raw == null) return '';
  const cleaned = String(raw).replace(/,/g, '');
  if (!/^\d*\.?\d{0,2}$/.test(cleaned)) return null;
  if (cleaned === '' || cleaned === '.') return cleaned;
  const [intPart, decPart] = cleaned.split('.');
  const intFmt = intPart === '' ? '' : Number(intPart).toLocaleString('en-US');
  return decPart === undefined ? intFmt : intFmt + '.' + decPart;
};
const parseAmt = v => { const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? 0 : n; };
const blurAmt = v => { const t = String(v).trim(); return (t && t !== '.') ? parseAmt(t).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) : t; };
const today = () => new Date().toISOString().slice(0, 10);
const fy_start = () => new Date().getFullYear() + '-01-01';
const fmtSize = b => b > 1048576 ? (b/1048576).toFixed(1)+' MB' : (b/1024).toFixed(0)+' KB';
const acctLabel = (code, name) => code + ' - ' + name;
function exportToExcel(data, fn) { const ws = XLSX.utils.aoa_to_sheet(data); ws['!cols'] = data[0].map((_, ci) => ({ wch: Math.min(Math.max(...data.map(r => String(r[ci]||'').length), 8)+2, 40) })); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); XLSX.writeFile(wb, fn); }
const BLANK_JE = () => ({date:today(),memo:'',lines:[{account_code:'',debit:'',credit:'',description:''},{account_code:'',debit:'',credit:'',description:''}]});
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
  td: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight, fontSize: 13, lineHeight: 1.4, verticalAlign: 'middle', height: 42, boxSizing: 'border-box', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', fontVariantNumeric: 'tabular-nums' },
  tdR: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight, textAlign: 'right', fontVariantNumeric: 'tabular-nums', fontSize: 13, lineHeight: 1.4, verticalAlign: 'middle', height: 42, boxSizing: 'border-box', whiteSpace: 'nowrap' },
  tdC: { padding: '10px 14px', borderBottom: '1px solid '+T.borderLight, textAlign: 'center', fontSize: 13, lineHeight: 1.4, verticalAlign: 'middle', height: 42, boxSizing: 'border-box', whiteSpace: 'nowrap' },
  tdBold: { padding: '10px 14px', borderBottom: '2px solid '+T.border, color: T.textBright, fontWeight: 700, fontVariantNumeric: 'tabular-nums', fontSize: 13, lineHeight: 1.4, verticalAlign: 'middle', height: 42, boxSizing: 'border-box', whiteSpace: 'nowrap' },
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
  modal: { position: 'fixed', inset: 0, zIndex: 200, display: 'flex', alignItems: 'center', justifyContent: 'center' },
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
const NI = { dashboard:'\u25a3', journal:'\u270e', coa:'\u2630', ledger:'\u2261', banktxn:'\u21c5', bankrec:'\u2611', trial:'\u2696', bs:'\u25a6', is:'\u25a4', wip:'▧', entities:'\u2302', users:'\u263a' };

// ─── Autocomplete ───
function AccountAutocomplete({accounts,value,onChange,placeholder,exclude}){
  const[q,setQ]=useState('');const[open,setOpen]=useState(false);const[placement,setPlacement]=useState('down');const ref=useRef(null);const inputRef=useRef(null);
  const sel=accounts.find(a=>a.code===value);
  const filtered=useMemo(()=>{const s=q.toLowerCase();return accounts.filter(a=>(!exclude||a.code!==exclude)&&(a.code.toLowerCase().includes(s)||a.name.toLowerCase().includes(s))).sort((a,b)=>a.code.localeCompare(b.code));},[accounts,q,exclude]);
  useEffect(()=>{const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener('mousedown',h);return()=>document.removeEventListener('mousedown',h);},[]);
  // Decide whether to open the dropdown upward or downward based on available space
  const computePlacement=()=>{if(!inputRef.current)return;const r=inputRef.current.getBoundingClientRect();const below=window.innerHeight-r.bottom;const above=r.top;const desired=340;setPlacement(below<desired&&above>below?'up':'down');};
  return(<div ref={ref} style={{position:'relative'}}><input ref={inputRef} style={S.inputSm} placeholder={placeholder||'Search account...'} value={open?q:(sel?acctLabel(sel.code,sel.name):'')}
    onFocus={()=>{computePlacement();setOpen(true);setQ('');}} onChange={e=>{setQ(e.target.value);setOpen(true);}} onKeyDown={e=>{if(e.key==='Escape')setOpen(false);if(e.key==='Enter'&&filtered.length>0){onChange(filtered[0].code);setOpen(false);}}}/>
    {open&&filtered.length>0&&<div style={{position:'absolute',...(placement==='up'?{bottom:'100%',marginBottom:4}:{top:'100%',marginTop:4}),left:0,right:0,background:'#fff',border:'1px solid '+T.border,borderRadius:T.radiusSm,maxHeight:340,overflowY:'auto',zIndex:50,boxShadow:T.shadowLg}}>
      {filtered.map(a=><div key={a.code} style={{padding:'8px 12px',cursor:'pointer',fontSize:12,display:'flex',justifyContent:'space-between',background:a.code===value?T.accentDim:'transparent'}}
        onClick={()=>{onChange(a.code);setOpen(false);}} onMouseEnter={e=>e.currentTarget.style.background=T.bgHover} onMouseLeave={e=>e.currentTarget.style.background=a.code===value?T.accentDim:'transparent'}>
        <span><b style={{color:T.textBright}}>{a.code}</b> <span style={{color:T.textMuted}}>{a.name}</span></span><span style={S.tag(a.type)}>{a.type}</span></div>)}</div>}</div>);}

// ─── Auth ───
function AuthScreen({onLogin}){const[mode,setMode]=useState('login');const[email,setEmail]=useState('');const[pw,setPw]=useState('');const[name,setName]=useState('');const[confirmPw,setConfirmPw]=useState('');const[role,setRole]=useState('Accountant');
  const[err,setErr]=useState('');const[success,setSuccess]=useState('');const[loading,setLoading]=useState(false);const[tempPw,setTempPw]=useState('');
  const doLogin=async()=>{setLoading(true);setErr('');try{const d=await api.login(email.trim().toLowerCase(),pw);api.setToken(d.token);onLogin(d.user);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doSignup=async()=>{if(!name.trim()){setErr('Name required');return;}if(pw.length<3){setErr('Min 3 chars');return;}if(pw!==confirmPw){setErr("Passwords don't match");return;}setLoading(true);setErr('');try{await api.signup(name.trim(),email.trim().toLowerCase(),pw,role);setSuccess('Account created!');setTimeout(()=>{setMode('login');setSuccess('');},1200);}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const doForgot=async()=>{if(!email.trim()){setErr('Enter email');return;}setLoading(true);setErr('');try{await api.forgotPassword(email.trim().toLowerCase());setSuccess('If an account exists for that email, a reset link has been sent. Check your inbox.');}catch(e){setErr(e.message);}finally{setLoading(false);}};
  const hk=e=>{if(e.key==='Enter'){mode==='login'?doLogin():mode==='signup'?doSignup():doForgot();}};
  return(<div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100vh',background:'#f1f5f9'}}>
    <div style={{background:'#fff',border:'1px solid '+T.border,borderRadius:16,width:420,padding:44,textAlign:'center',boxShadow:T.shadowLg}}>
      <div style={{margin:'0 auto 16px',width:48}}><Logo size={48}/></div>
      <div style={{fontSize:24,fontWeight:800,color:T.textBright,marginBottom:4}}>CloudLedger</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:32}}>Multi-Entity Cloud Accounting</div>
      {mode==='forgot'?(<>
        <div style={{fontSize:15,fontWeight:600,color:T.textBright,marginBottom:20}}>Reset Password</div>
        <div style={{marginBottom:12}}><input style={S.input} placeholder="Email address" value={email} onChange={e=>{setEmail(e.target.value);setErr('');setSuccess('');}} onKeyDown={hk}/></div>
        <div style={{fontSize:12,color:T.textMuted,marginBottom:12,lineHeight:1.5}}>Enter your email and we will send you a link to reset your password.</div>
        {err&&<div style={S.err}>{err}</div>}
        {success&&<div style={{background:T.greenDim,border:'1px solid '+T.greenBorder,borderRadius:T.radiusSm,padding:16,margin:'12px 0',fontSize:13,color:T.green,lineHeight:1.5}}>{success}</div>}
        <button style={{...S.btnP,width:'100%',padding:11,marginTop:8}} onClick={doForgot} disabled={loading}>{loading?'...':'Reset Password'}</button>
        <div style={{marginTop:20}}><button style={S.link} onClick={()=>{setMode('login');setErr('');setSuccess('');}}>Back to Sign In</button></div>
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
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:500}} onClick={e=>e.stopPropagation()}>
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
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:640}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button><div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Add New Account</div>
    <div style={S.row}><div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
    <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
    {err&&<div style={S.err}>{err}</div>}<div style={{display:'flex',gap:10,marginTop:12}}><button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Code and name required');return;}try{const a=await api.createAccount(entityId,form);onCreated(a);onClose();}catch(e){setErr(e.message);}}}>Add Account</button><button style={S.btnS} onClick={onClose}>Cancel</button></div>
  </div></div>);}

// ─── JE Modal — form state received from App (persists across open/close) ───
function JournalEntryModal({entityId,isTurnkeyEntity,dimsEnabled,user,onClose,onPosted,form,setForm,pendingFiles,setPendingFiles}){
  const[accounts,setAccounts]=useState([]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[posting,setPosting]=useState(false);const[posted,setPosted]=useState('');
  const[projects,setProjects]=useState([]);
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts);api.getTurnkeyProjects().then(setProjects).catch(()=>setProjects([]));api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  const showProject=isTurnkeyEntity||projects.length>0;
  const showLocation=dimsEnabled&&locations.length>0;const showClass=dimsEnabled&&classes.length>0;
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:'',description:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const tDr=form.lines.reduce((s,l)=>s+parseAmt(l.debit),0);const tCr=form.lines.reduce((s,l)=>s+parseAmt(l.credit),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;
  const discard=()=>{setForm(BLANK_JE());setPendingFiles([]);};
  const onFilesSelected=e=>{const files=Array.from(e.target.files);if(files.length>0)setPendingFiles(p=>[...p,...files]);e.target.value='';};
  const post=async()=>{if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}if(form.lines.some(l=>!l.account_code)){setErr('All lines need an account');return;}if(!bal){setErr('Entry must balance');return;}
    setPosting(true);setErr('');try{const r=await api.createEntry(entityId,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseAmt(l.debit),credit:parseAmt(l.credit),description:l.description||'',project_id:l.project_id||null,location_id:l.location_id||null,class_id:l.class_id||null}))});
      let msg='JE-'+String(r.entry_num).padStart(4,'0')+' posted';
      if(pendingFiles.length>0){try{const u=await api.uploadAttachments(entityId,r.id,pendingFiles);msg+=' with '+u.length+' attachment(s)';}catch(ue){msg+=' (attachments failed: '+ue.message+')';}}
      setForm(BLANK_JE());setPendingFiles([]);setPosted(msg+'!');setTimeout(()=>setPosted(''),5000);if(onPosted)onPosted();}
    catch(e){setErr(e.message);}finally{setPosting(false);}};
  const hasContent=form.memo||form.lines.some(l=>l.account_code||l.debit||l.credit)||pendingFiles.length>0;

  return(<div style={S.modal}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:980}} onClick={e=>e.stopPropagation()}>
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
    <div style={{...S.cardFlush,marginBottom:16}}><table style={S.table}><thead><tr><th style={S.th}>Account</th>{showProject&&<th style={{...S.th,width:170}}>Project</th>}{showLocation&&<th style={{...S.th,width:150}}>Location</th>}{showClass&&<th style={{...S.th,width:150}}>Class</th>}<th style={S.th}>Description</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select account...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        {showProject&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.project_id||''} onChange={e=>updateLine(i,'project_id',e.target.value)}><option value="">— none —</option>{projects.map(pr=><option key={pr.turnkey_project_id} value={pr.turnkey_project_id}>{pr.project_code} — {pr.project_name}</option>)}</select></td>}
        {showLocation&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.location_id||''} onChange={e=>updateLine(i,'location_id',e.target.value)}><option value="">— none —</option>{locations.map(loc=><option key={loc.id} value={loc.id}>{loc.code?loc.code+" — ":""}{loc.name}</option>)}</select></td>}
        {showClass&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.class_id||''} onChange={e=>updateLine(i,'class_id',e.target.value)}><option value="">— none —</option>{classes.map(c=><option key={c.id} value={c.id}>{c.code?c.code+" — ":""}{c.name}</option>)}</select></td>}
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={S.input} placeholder="(optional)" value={l.description||''} onChange={e=>updateLine(i,'description',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.debit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'debit',f);}} onBlur={e=>updateLine(i,'debit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.credit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'credit',f);}} onBlur={e=>updateLine(i,'credit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td colSpan={2+(showProject?1:0)+(showLocation?1:0)+(showClass?1:0)} style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
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
  const filtered=entities.filter(e=>e.name.toLowerCase().includes(search.toLowerCase()));
  return(<div style={{position:'relative'}}><div style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',padding:'6px 14px',borderRadius:T.radiusSm,background:T.bgElevated,border:'1px solid '+T.border}} onClick={()=>setOpen(!open)}>
    <span style={{fontWeight:600,color:T.textBright,fontSize:13,maxWidth:240,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{active?.name||'Select entity'}</span>
    <span style={{color:T.textDim,fontSize:9}}>{'\u25bc'}</span></div>
    {open&&<><div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>{setOpen(false);setSearch('');}}/>
      <div style={{position:'absolute',top:'100%',left:0,background:'#fff',border:'1px solid '+T.border,borderRadius:T.radius,maxHeight:380,overflowY:'auto',zIndex:100,boxShadow:T.shadowLg,width:340,marginTop:6}}>
        <div style={{position:'sticky',top:0,padding:12,background:'#fff',borderBottom:'1px solid '+T.border}}><input style={S.input} placeholder={'Search '+entities.length+' entities...'} value={search} onChange={e=>setSearch(e.target.value)} autoFocus/></div>
        {filtered.map(e=><div key={e.id} style={{padding:'10px 16px',cursor:'pointer',background:e.id===activeId?T.accentDim:'transparent',borderLeft:e.id===activeId?'3px solid '+T.accent:'3px solid transparent'}} onClick={()=>{onSelect(e.id);setOpen(false);setSearch('');}}>
          <span style={{fontWeight:600,color:T.textBright,fontSize:13}}>{e.name}</span></div>)}
        <div style={{borderTop:'1px solid '+T.border,padding:12}}><button style={{...S.btnS,width:'100%'}} onClick={()=>{onManage();setOpen(false);}}>Manage Entities</button></div></div></>}</div>);}

// ═══ Main App — JE form state lives here so it persists across modal open/close ═══
function ResetPasswordScreen({token}){
  const[pw,setPw]=useState('');const[confirm,setConfirm]=useState('');
  const[err,setErr]=useState('');const[done,setDone]=useState(false);const[loading,setLoading]=useState(false);
  const submit=async()=>{
    if(pw.length<6){setErr('Password must be at least 6 characters');return;}
    if(pw!==confirm){setErr("Passwords don't match");return;}
    setLoading(true);setErr('');
    try{await api.resetPassword(token,pw);setDone(true);setTimeout(()=>{window.location.href='/';},2000);}
    catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const hk=e=>{if(e.key==='Enter')submit();};
  return(<div style={{display:'flex',alignItems:'center',justifyContent:'center',minHeight:'100vh',background:'#f1f5f9'}}>
    <div style={{background:'#fff',border:'1px solid '+T.border,borderRadius:16,width:420,padding:44,textAlign:'center',boxShadow:T.shadowLg}}>
      <div style={{margin:'0 auto 16px',width:48}}><Logo size={48}/></div>
      <div style={{fontSize:24,fontWeight:800,color:T.textBright,marginBottom:4}}>CloudLedger</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:32}}>Set a new password</div>
      {done?(<div style={{background:T.greenDim,border:'1px solid '+T.greenBorder,borderRadius:T.radiusSm,padding:20,fontSize:14,color:T.green,lineHeight:1.5}}>Password updated. Redirecting to sign in…</div>):(<>
        <div style={{marginBottom:12}}><input style={S.input} type="password" placeholder="New password" value={pw} onChange={e=>{setPw(e.target.value);setErr('');}} onKeyDown={hk}/></div>
        <div style={{marginBottom:12}}><input style={S.input} type="password" placeholder="Confirm new password" value={confirm} onChange={e=>{setConfirm(e.target.value);setErr('');}} onKeyDown={hk}/></div>
        {err&&<div style={S.err}>{err}</div>}
        <button style={{...S.btnP,width:'100%',padding:11,marginTop:8}} onClick={submit} disabled={loading}>{loading?'...':'Update password'}</button>
        <div style={{marginTop:20}}><button style={S.link} onClick={()=>{window.location.href='/';}}>Back to Sign In</button></div>
      </>)}
    </div>
  </div>);
}

export default function App(){
  const[user,setUser]=useState(null);const[entities,setEntities]=useState([]);const[activeEntity,setActiveEntity]=useState(null);
  const[page,setPage]=useState('dashboard');const[loading,setLoading]=useState(true);
  // Back-button trap with diagnostic logging
  useEffect(()=>{
    console.log('[CL-BACK] mount, user=', user?.email || 'null');
    if (!user) return;
    let leavingApp = false;
    try {
      window.history.pushState({cl_app: 1}, '');
      window.history.pushState({cl_app: 2}, '');
      console.log('[CL-BACK] 2 sentinels pushed. length=', window.history.length, 'state=', window.history.state);
    } catch (e) { console.error('[CL-BACK] push failed:', e); }

    const onPop = (e) => {
      console.log('[CL-BACK] popstate event.state=', e.state, 'history.state=', window.history.state, 'length=', window.history.length);
      if (leavingApp) { console.log('[CL-BACK] leavingApp=true'); return; }
      try {
        window.history.pushState({cl_app: 1}, '');
        window.history.pushState({cl_app: 2}, '');
        console.log('[CL-BACK] re-pushed 2 sentinels. length=', window.history.length);
      } catch (err) { console.error('[CL-BACK] re-push failed:', err); }
      setPage(curPage => {
        console.log('[CL-BACK] curPage=', curPage);
        if (curPage !== 'dashboard') { console.log('[CL-BACK] -> dashboard'); return 'dashboard'; }
        setTimeout(() => {
          console.log('[CL-BACK] on dashboard, asking confirm');
          if (window.confirm('Leave CloudLedger?')) {
            console.log('[CL-BACK] user confirmed, exiting');
            leavingApp = true;
            window.history.go(-3);
          } else {
            console.log('[CL-BACK] user cancelled');
          }
        }, 0);
        return curPage;
      });
    };
    window.addEventListener('popstate', onPop);
    console.log('[CL-BACK] popstate listener attached');
    return () => {
      console.log('[CL-BACK] cleanup');
      window.removeEventListener('popstate', onPop);
    };
  }, [user]);
  // Make all modal windows draggable by their top header strip
  useEffect(() => {
    const styleEl = document.createElement('style');
    styleEl.textContent = '.cl-modal-box::before{content:"";position:absolute;top:0;left:0;right:56px;height:44px;cursor:move;border-top-left-radius:14px;border-top-right-radius:14px;z-index:1;}';
    document.head.appendChild(styleEl);
    const onDown = (e) => {
      const box = e.target.closest && e.target.closest('.cl-modal-box');
      if (!box) return;
      const rect = box.getBoundingClientRect();
      if (e.clientY > rect.top + 44) return;
      if (e.target.closest('input,textarea,select,button,a,[contenteditable]')) return;
      e.preventDefault();
      const m = (box.style.transform || '').match(/translate\(([-\d.]+)px,\s*([-\d.]+)px\)/);
      const ox = m ? parseFloat(m[1]) : 0;
      const oy = m ? parseFloat(m[2]) : 0;
      const sx = e.clientX, sy = e.clientY;
      const onMove = (ev) => { box.style.transform = 'translate(' + (ox + ev.clientX - sx) + 'px,' + (oy + ev.clientY - sy) + 'px)'; };
      const onUp = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    };
    document.addEventListener('mousedown', onDown);
    return () => { document.removeEventListener('mousedown', onDown); if (styleEl.parentNode) styleEl.parentNode.removeChild(styleEl); };
  }, []);
  const[showJE,setShowJE]=useState(false);const[showChangePw,setShowChangePw]=useState(false);const[rk,setRk]=useState(0);
  const[sidebarCol,setSidebarCol]=useState(()=>{try{return localStorage.getItem(SIDEBAR_KEY)==='true';}catch{return false;}});
  // JE form state lives in App — survives modal close, cleared only on post/discard
  const[jeForm,setJeForm]=useState(BLANK_JE());const[jePendingFiles,setJePendingFiles]=useState([]);
  // Bank transaction state lifted to App so it persists across page navigation
  const BANK_SEL_KEY='cl_bank_sel_by_entity';
  const BANK_STATUS_KEY='cl_bank_status_by_entity';
  const loadPerEntity=k=>{try{return JSON.parse(localStorage.getItem(k)||'{}');}catch{return {};}};
  const[bankSelByEntity,setBankSelByEntity]=useState(()=>loadPerEntity(BANK_SEL_KEY));
  const[bankStatusByEntity,setBankStatusByEntity]=useState(()=>loadPerEntity(BANK_STATUS_KEY));
  const bankSelAcct=activeEntity?(bankSelByEntity[activeEntity]||''):'';
  const bankStatusFilter=activeEntity?(bankStatusByEntity[activeEntity]||''):'';
  const setBankSelAcct=v=>{if(!activeEntity)return;const next={...bankSelByEntity,[activeEntity]:v||''};setBankSelByEntity(next);try{localStorage.setItem(BANK_SEL_KEY,JSON.stringify(next));}catch{}};
  const setBankStatusFilter=v=>{if(!activeEntity)return;const next={...bankStatusByEntity,[activeEntity]:v||''};setBankStatusByEntity(next);try{localStorage.setItem(BANK_STATUS_KEY,JSON.stringify(next));}catch{}};
  const[bankTxns,setBankTxns]=useState([]);const[bankUploading,setBankUploading]=useState(false);
  // Report filter state lifted so they persist across page navigation
  const[tbAsOf,setTbAsOf]=useState(today());
  const[wipAsOf,setWipAsOf]=useState(today());
  const[bsAsOf,setBsAsOf]=useState(today());
  const[isFrom,setIsFrom]=useState(fy_start());const[isTo,setIsTo]=useState(today());
  const[glFrom,setGlFrom]=useState(fy_start());const[glTo,setGlTo]=useState(today());const[glFilter,setGlFilter]=useState('');
  // Requisition working set — lifted to App so uploaded invoices/coding survive
  // navigation. Kept per-entity so switching entities shows the right set.
  const[reqStateByEntity,setReqStateByEntity]=useState({});
  const reqState=(activeEntity&&reqStateByEntity[activeEntity])||null;
  const setReqState=updater=>{if(!activeEntity)return;setReqStateByEntity(prev=>{const cur=prev[activeEntity]||{cards:[],reqNum:'',asOf:today(),result:null,detail:null};const next=typeof updater==='function'?updater(cur):updater;return{...prev,[activeEntity]:next};});};

  useEffect(()=>{try{localStorage.setItem(SIDEBAR_KEY,String(sidebarCol));}catch{}},[sidebarCol]);
  useEffect(()=>{const t=api.getToken();if(t){api.me().then(u=>{if(u)setUser(u);}).catch(()=>api.clearToken()).finally(()=>setLoading(false));}else setLoading(false);},[]);
  useEffect(()=>{if(user)api.getEntities().then(e=>{setEntities(e);if(e.length>0&&!activeEntity)setActiveEntity(e[0].id);});},[user]);
  const refreshEntities=useCallback(async()=>{const e=await api.getEntities();setEntities(e);return e;},[]);
  const canAccess=s=>{if(!user)return false;if(user.role==='Admin')return true;return({Accountant:['entries','reports','coa','bankrec'],Viewer:['entries','reports','coa','bankrec']}[user.role]||[]).includes(s);};
  // Read-only users (Viewer) SEE the same sections as an Accountant but cannot edit.
  // canEdit gates every write control; it must never be derived from mere visibility.
  const canEdit = !!user && (user.role==='Admin' || user.role==='Accountant');
  if(loading)return<div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:T.textMuted}}>Loading...</div></div>;
  const _resetToken=(()=>{try{return new URLSearchParams(window.location.search).get('reset_token');}catch{return null;}})();
  if(_resetToken)return<ResetPasswordScreen token={_resetToken}/>;
  if(!user)return<AuthScreen onLogin={setUser}/>;
  const jeHasContent=jeForm.memo||jeForm.lines.some(l=>l.account_code||l.debit||l.credit)||jePendingFiles.length>0;

  const _activeEnt = entities.find(e=>e.id===activeEntity);
  const isTurnkeyEntity = !!(_activeEnt && (_activeEnt.code==='TURNKEYR' || /turnkey\s*rail/i.test(_activeEnt.name||'')));
  const isDevEntity = !!(_activeEnt && _activeEnt.entity_type==='development');
  const isShellEntity = !!(_activeEnt && _activeEnt.entity_type==='shell');
  const dimsEnabled = !!_activeEnt && !isShellEntity;// location/class dimensions available on every entity EXCEPT shell
  const navItems=[
    {id:'dashboard',label:'Dashboard',icon:NI.dashboard,section:'reports'},
    {id:'d1',divider:1,label:'TRANSACTIONS'},{id:'journal',label:'Journal Entries',icon:NI.journal,section:'entries'},
    {id:'d2',divider:1,label:'ACCOUNTS'},{id:'coa',label:'Chart of Accounts',icon:NI.coa,section:'coa'},...(dimsEnabled?[{id:'dimensions',label:'Locations & Classes',icon:'🏷️',section:'coa'}]:[]),{id:'ledger',label:'General Ledger',icon:NI.ledger,section:'reports'},
    {id:'d2b',divider:1,label:'BANKING'},{id:'banktxn',label:'Bank Transactions',icon:NI.banktxn,section:'bankrec'},{id:'bankrec',label:'Bank Reconciliation',icon:NI.bankrec,section:'bankrec'},
    {id:'d3',divider:1,label:'REPORTS'},{id:'trial',label:'Trial Balance',icon:NI.trial,section:'reports'},{id:'bs',label:'Balance Sheet',icon:NI.bs,section:'reports'},{id:'is',label:'Income Statement',icon:NI.is,section:'reports'},
    ...(isTurnkeyEntity?[{id:'wip',label:'WIP Schedule',icon:NI.wip,section:'reports'}]:[]),
    ...(isDevEntity?[{id:'d3b',divider:1,label:'DEVELOPMENT'},{id:'requisitions',label:'Requisitions',icon:'🏗️',section:'reports'}]:[]),
    {id:'d4',divider:1,label:'ADMIN'},{id:'entities',label:'Entities ('+entities.length+')',icon:NI.entities,section:'all'},{id:'users',label:'Users',icon:NI.users,section:'all'},
    {id:'d5',divider:1,label:'INTEGRATIONS'},{id:'billcom',label:'Bill.com Setup',icon:'💳',section:'all'},
  ];

  return(<div style={S.app}>
    <div style={S.topBar}><div style={{display:'flex',alignItems:'center',gap:16}}>
      <button style={{...S.btnGhost,fontSize:18,padding:'4px 6px',color:T.textMuted}} onClick={()=>setSidebarCol(c=>!c)}>{sidebarCol?'\u2630':'\u2190'}</button>
      <div style={{display:'flex',alignItems:'center',gap:10}}><Logo size={32}/>{!sidebarCol&&<div style={{fontSize:17,fontWeight:800,color:T.textBright}}>CloudLedger</div>}</div>
      <div style={{width:1,height:28,background:T.border}}/><EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')}/></div>
      <div style={{display:'flex',alignItems:'center',gap:10}}>
        {canEdit&&activeEntity&&<button style={{...S.btnP,position:'relative'}} onClick={()=>setShowJE(true)}>+ Journal Entry{jeHasContent&&<span style={{position:'absolute',top:-3,right:-3,width:8,height:8,borderRadius:4,background:T.orange,border:'2px solid #fff'}}/>}</button>}
        <span style={{fontSize:13,fontWeight:500}}>{user.name}</span><span style={S.badge}>{user.role}</span>
        <button style={S.btnS} onClick={()=>setShowChangePw(true)}>Settings</button>
        <button style={S.btnS} onClick={()=>{api.clearToken();setUser(null);}}>Sign Out</button></div></div>
    <div style={S.body}><div style={S.sidebar(sidebarCol)}>
      {navItems.map(n=>n.divider?(!sidebarCol?<div key={n.id} style={S.navSection(sidebarCol)}>{n.label}</div>:<div key={n.id} style={{height:8}}/>)
        :(n.section==='all'?user.role==='Admin':canAccess(n.section))?<div key={n.id} style={S.navItem(page===n.id,sidebarCol)} onClick={()=>setPage(n.id)} title={n.label}>
          {sidebarCol?<span style={{fontSize:15}}>{n.icon}</span>:<span>{n.icon}  {n.label}</span>}</div>:null)}</div>
      <div style={S.main}>{(()=>{const en=entities.find(e=>e.id===activeEntity);const entityName=en?en.name:'';return<>
        {page==='dashboard'&&<Dashboard entityId={activeEntity} setActiveEntity={setActiveEntity} setPage={setPage} user={user} key={rk}/>}
        {page==='journal'&&activeEntity&&<JournalList entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} canEdit={canEdit} key={activeEntity+'-'+rk} onNewEntry={()=>setShowJE(true)}/>}
        {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} canEdit={canEdit}/>}
        {page==='dimensions'&&activeEntity&&dimsEnabled&&<DimensionsManager entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} key={activeEntity+'-'+rk} from={glFrom} setFrom={setGlFrom} to={glTo} setTo={setGlTo} filter={glFilter} setFilter={setGlFilter}/>}
        {page==='banktxn'&&activeEntity&&<BankTransactions entityId={activeEntity} canEdit={canEdit} bankSelAcct={bankSelAcct} setBankSelAcct={setBankSelAcct} bankTxns={bankTxns} setBankTxns={setBankTxns} bankUploading={bankUploading} setBankUploading={setBankUploading} bankStatusFilter={bankStatusFilter} setBankStatusFilter={setBankStatusFilter}/>}
        {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user} canEdit={canEdit}/>}
        {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} key={activeEntity+'-'+rk} asOf={tbAsOf} setAsOf={setTbAsOf}/>}
        {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity} entityName={entityName} asOf={bsAsOf} setAsOf={setBsAsOf}/>}
        {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity} entityName={entityName} from={isFrom} setFrom={setIsFrom} to={isTo} setTo={setIsTo}/>}
        {page==='wip'&&activeEntity&&<WipSchedule entityName={entityName} asOf={wipAsOf} setAsOf={setWipAsOf}/>}
        {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='users'&&<UserManagement currentUser={user}/>}
        {page==='billcom'&&<BillcomSetup entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='requisitions'&&activeEntity&&isDevEntity&&<Requisitions entityId={activeEntity} entityName={entityName} canEdit={canEdit} reqState={reqState} setReqState={setReqState}/>}
      </>})()}</div></div>
    {showJE&&activeEntity&&<JournalEntryModal entityId={activeEntity} isTurnkeyEntity={isTurnkeyEntity} dimsEnabled={dimsEnabled} user={user} onClose={()=>setShowJE(false)} onPosted={()=>setRk(k=>k+1)} form={jeForm} setForm={setJeForm} pendingFiles={jePendingFiles} setPendingFiles={setJePendingFiles}/>}
    {showChangePw&&<SettingsModal onClose={()=>setShowChangePw(false)} user={user} onUserUpdate={u=>setUser(u)}/>}
  </div>);}

// ═══ Spreadsheet Editor Modal ═══
function SpreadsheetEditorModal({ file, onClose, onSaved }) {
  const COL_W = 110;
  const ROW_H = 24;
  const ROW_NUM_W = 46;

  const [workbook, setWorkbook]       = useState(null);
  const [sheetDataMap, setSheetDataMap] = useState({});
  const [activeSheet, setActiveSheet] = useState('');
  const [sel, setSel]                 = useState({ r: 0, c: 0 });
  const [editing, setEditing]         = useState(false);
  const [editVal, setEditVal]         = useState('');
  const [loading, setLoading]         = useState(true);
  const [saving, setSaving]           = useState(false);
  const [err, setErr]                 = useState('');
  const [msg, setMsg]                 = useState('');
  const [dirty, setDirty]             = useState(false);

  const gridRef      = useRef(null);
  const editInputRef = useRef(null);

  const colLetter = useCallback(n => {
    let s = ''; let x = n + 1;
    while (x > 0) { const rem = (x - 1) % 26; s = String.fromCharCode(65 + rem) + s; x = Math.floor((x - 1) / 26); }
    return s;
  }, []);

  const parseWorksheet = useCallback(ws => {
    if (!ws || !ws['!ref']) return Array.from({ length: 20 }, () => Array(10).fill({ v: '', w: '', t: 's' }));
    const range = XLSX.utils.decode_range(ws['!ref']);
    const numR = Math.max(range.e.r - range.s.r + 1, 20);
    const numC = Math.max(range.e.c - range.s.c + 1, 10);
    const data = [];
    for (let r = 0; r < numR; r++) {
      const row = [];
      for (let c = 0; c < numC; c++) {
        const addr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
        const cell = ws[addr];
        row.push(cell
          ? { v: cell.v ?? '', w: cell.w || (cell.v != null ? String(cell.v) : ''), t: cell.t || 's', f: cell.f || null }
          : { v: '', w: '', t: 's', f: null });
      }
      data.push(row);
    }
    return data;
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const resp = await fetch(api.downloadEntityFile(file.id));
        if (!resp.ok) throw new Error('Failed to load file');
        const ab = await resp.arrayBuffer();
        const wb = XLSX.read(ab, { type: 'array', cellStyles: true, cellNF: true });
        const map = {};
        wb.SheetNames.forEach(n => { map[n] = parseWorksheet(wb.Sheets[n]); });
        setWorkbook(wb); setSheetDataMap(map); setActiveSheet(wb.SheetNames[0]);
      } catch (ex) { setErr(ex.message); }
      finally { setLoading(false); }
    })();
  }, [file.id, parseWorksheet]);

  useEffect(() => { if (!loading && !editing && gridRef.current) gridRef.current.focus(); }, [loading, editing]);
  useEffect(() => { if (editing && editInputRef.current) { editInputRef.current.focus(); editInputRef.current.select(); } }, [editing]);

  const currentData = sheetDataMap[activeSheet] || [];
  const numRows = currentData.length;
  const numCols = currentData[0]?.length || 0;

  const getCell   = (r, c) => currentData[r]?.[c] || { v: '', w: '', t: 's', f: null };
  const dispVal   = (r, c) => { const cell = getCell(r, c); return cell.w !== undefined ? cell.w : (cell.v != null ? String(cell.v) : ''); };
  const rawVal    = (r, c) => { const cell = getCell(r, c); return cell.v != null ? String(cell.v) : ''; };
  const isNumCell = (r, c) => getCell(r, c).t === 'n';

  const updateCell = (r, c, val) => {
    setSheetDataMap(prev => {
      const data = prev[activeSheet].map(row => [...row]);
      if (r < data.length && c < (data[0]?.length || 0)) {
        const num = Number(val); const isNum = val.trim() !== '' && !isNaN(num);
        data[r][c] = { v: isNum ? num : val, w: val, t: isNum ? 'n' : 's', f: null };
      }
      return { ...prev, [activeSheet]: data };
    });
    setDirty(true);
  };

  const startEdit  = (r, c, init = null) => { setSel({ r, c }); setEditVal(init !== null ? init : rawVal(r, c)); setEditing(true); };
  const commitEdit = ()        => { updateCell(sel.r, sel.c, editVal); setEditing(false); };
  const cancelEdit = ()        => { setEditing(false); setEditVal(''); };

  const moveSel = (dr, dc) => setSel(prev => ({
    r: Math.max(0, Math.min(numRows - 1, prev.r + dr)),
    c: Math.max(0, Math.min(numCols - 1, prev.c + dc)),
  }));

  const switchSheet = name => { if (editing) commitEdit(); setActiveSheet(name); setSel({ r: 0, c: 0 }); setEditing(false); };

  const handleGridKeyDown = e => {
    if (editing) return;
    if (e.key === 'ArrowUp')    { e.preventDefault(); moveSel(-1,  0); }
    else if (e.key === 'ArrowDown')  { e.preventDefault(); moveSel( 1,  0); }
    else if (e.key === 'ArrowLeft')  { e.preventDefault(); moveSel( 0, -1); }
    else if (e.key === 'ArrowRight') { e.preventDefault(); moveSel( 0,  1); }
    else if (e.key === 'Tab')        { e.preventDefault(); moveSel(0, 1); }
    else if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEdit(sel.r, sel.c); }
    else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); updateCell(sel.r, sel.c, ''); }
    else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) { startEdit(sel.r, sel.c, e.key); }
  };

  const handleEditKeyDown = e => {
    if (e.key === 'Enter')  { e.preventDefault(); commitEdit(); moveSel(1, 0); }
    else if (e.key === 'Tab')    { e.preventDefault(); commitEdit(); moveSel(0, 1); }
    else if (e.key === 'Escape') { cancelEdit(); }
  };

  const save = async () => {
    if (!workbook) return;
    setSaving(true); setErr(''); setMsg('');
    try {
      Object.entries(sheetDataMap).forEach(([sheetName, data]) => {
        const ws = workbook.Sheets[sheetName]; if (!ws) return;
        const wsRange = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
        data.forEach((row, r) => {
          row.forEach((cell, c) => {
            const addr = XLSX.utils.encode_cell({ r: r + wsRange.s.r, c: c + wsRange.s.c });
            if (cell.v === '' || cell.v == null) { if (ws[addr]) { ws[addr].v = ''; ws[addr].t = 's'; delete ws[addr].w; } return; }
            if (!ws[addr]) ws[addr] = {};
            ws[addr].v = cell.v; ws[addr].t = cell.t || 's';
            if (cell.f) ws[addr].f = cell.f; else delete ws[addr].f;
            delete ws[addr].w;
          });
        });
        ws['!ref'] = XLSX.utils.encode_range({ s: wsRange.s, e: { r: wsRange.s.r + data.length - 1, c: wsRange.s.c + (data[0]?.length || 1) - 1 } });
      });
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob  = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const newFile = new File([blob], file.original_name, { type: blob.type });
      await api.replaceEntityFile(file.id, newFile);
      setDirty(false); setMsg('Saved!');
      if (onSaved) onSaved();
      setTimeout(() => setMsg(''), 3000);
    } catch (ex) { setErr('Save failed: ' + ex.message); }
    finally { setSaving(false); }
  };

  const handleClose = () => {
    if (dirty && !confirm('You have unsaved changes. Close without saving?')) return;
    onClose();
  };

  const formulaBarCell = `${colLetter(sel.c)}${sel.r + 1}`;
  const formulaBarVal  = editing ? editVal : (getCell(sel.r, sel.c).f ? '=' + getCell(sel.r, sel.c).f : rawVal(sel.r, sel.c));

  const grd = { background: T.sidebarBg };
  const hdr = { position: 'sticky', background: '#2d3748', color: '#a0aec0', fontSize: 11, fontWeight: 600, textAlign: 'center', border: '1px solid #4a5568', userSelect: 'none', padding: '3px 0', whiteSpace: 'nowrap', overflow: 'hidden' };

  if (loading) return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', zIndex: 500, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: '#fff', borderRadius: 12, padding: 32, fontSize: 14, color: T.textMuted }}>Loading spreadsheet…</div>
    </div>
  );

  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 500, display: 'flex', flexDirection: 'column', background: '#1a202c', fontFamily: "'Inter',-apple-system,sans-serif" }}>

      {/* ── Toolbar ── */}
      <div style={{ height: 48, background: '#2d3748', borderBottom: '1px solid #4a5568', display: 'flex', alignItems: 'center', padding: '0 16px', gap: 12, flexShrink: 0 }}>
        <span style={{ fontSize: 16 }}>📊</span>
        <span style={{ color: '#e2e8f0', fontWeight: 600, fontSize: 14, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.original_name}</span>
        {dirty && <span style={{ fontSize: 11, color: '#f6ad55', background: '#7c2d0020', border: '1px solid #f6ad5540', borderRadius: 4, padding: '2px 8px' }}>Unsaved changes</span>}
        {msg   && <span style={{ fontSize: 11, color: '#68d391', background: '#276749', borderRadius: 4, padding: '2px 8px' }}>{msg}</span>}
        {err   && <span style={{ fontSize: 11, color: '#fc8181', background: '#742a2a', borderRadius: 4, padding: '2px 8px', maxWidth: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{err}</span>}
        <button
          onClick={save} disabled={saving || !dirty}
          style={{ background: dirty ? T.accent : '#4a5568', color: '#fff', border: 'none', borderRadius: 6, padding: '7px 18px', fontSize: 13, fontWeight: 600, cursor: dirty ? 'pointer' : 'not-allowed', opacity: saving ? 0.7 : 1 }}>
          {saving ? 'Saving…' : '💾 Save'}
        </button>
        <button onClick={handleClose} style={{ background: '#4a5568', color: '#e2e8f0', border: 'none', borderRadius: 6, padding: '7px 14px', fontSize: 13, cursor: 'pointer' }}>✕ Close</button>
      </div>

      {/* ── Formula Bar ── */}
      <div style={{ height: 32, background: '#2d3748', borderBottom: '1px solid #4a5568', display: 'flex', alignItems: 'center', padding: '0 12px', gap: 10, flexShrink: 0 }}>
        <span style={{ fontSize: 11, color: '#718096', fontWeight: 600, minWidth: 28, background: '#1a202c', border: '1px solid #4a5568', borderRadius: 4, padding: '2px 6px', textAlign: 'center' }}>{formulaBarCell}</span>
        <span style={{ color: '#718096', fontSize: 13 }}>ƒx</span>
        <div style={{ flex: 1, fontSize: 12, color: '#e2e8f0', background: '#1a202c', border: '1px solid #4a5568', borderRadius: 4, padding: '3px 8px', fontFamily: 'monospace', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {formulaBarVal}
        </div>
      </div>

      {/* ── Grid ── */}
      <div ref={gridRef} tabIndex={0} onKeyDown={handleGridKeyDown}
        style={{ flex: 1, overflow: 'auto', outline: 'none', background: '#1a202c' }}>
        <table style={{ borderCollapse: 'collapse', tableLayout: 'fixed', fontSize: 12 }}>
          <thead>
            <tr>
              <th style={{ ...hdr, position: 'sticky', top: 0, left: 0, zIndex: 4, width: ROW_NUM_W, minWidth: ROW_NUM_W }}></th>
              {Array.from({ length: numCols }, (_, c) => (
                <th key={c} style={{ ...hdr, position: 'sticky', top: 0, zIndex: 2, width: COL_W, minWidth: COL_W }}>{colLetter(c)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {currentData.map((row, r) => (
              <tr key={r}>
                <td style={{ ...hdr, position: 'sticky', left: 0, zIndex: 1, width: ROW_NUM_W, minWidth: ROW_NUM_W, height: ROW_H }}>{r + 1}</td>
                {row.map((cell, c) => {
                  const isSel = sel.r === r && sel.c === c;
                  const isEd  = isSel && editing;
                  return (
                    <td key={c}
                      onClick={() => { if (editing) commitEdit(); setSel({ r, c }); }}
                      onDoubleClick={() => startEdit(r, c)}
                      style={{
                        width: COL_W, minWidth: COL_W, height: ROW_H, maxHeight: ROW_H,
                        border: isSel ? '2px solid ' + T.accent : '1px solid #2d3748',
                        background: isSel ? '#1e3a5f' : (r % 2 === 0 ? '#1e2532' : '#1a202c'),
                        padding: isEd ? 0 : '0 6px',
                        color: '#e2e8f0', boxSizing: 'border-box',
                        overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                        textAlign: isNumCell(r, c) ? 'right' : 'left',
                        cursor: 'default', userSelect: 'none',
                      }}>
                      {isEd
                        ? <input ref={editInputRef} value={editVal} onChange={e => setEditVal(e.target.value)} onKeyDown={handleEditKeyDown} onBlur={commitEdit}
                            style={{ width: '100%', height: '100%', border: 'none', outline: 'none', padding: '0 6px', fontSize: 12, background: '#1e3a5f', color: '#fff', boxSizing: 'border-box', fontFamily: 'inherit', textAlign: isNumCell(r, c) ? 'right' : 'left' }} />
                        : dispVal(r, c)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* ── Sheet Tabs ── */}
      {workbook && workbook.SheetNames.length > 1 && (
        <div style={{ height: 34, background: '#2d3748', borderTop: '1px solid #4a5568', display: 'flex', alignItems: 'flex-end', padding: '0 8px', gap: 2, flexShrink: 0, overflowX: 'auto' }}>
          {workbook.SheetNames.map(name => (
            <button key={name} onClick={() => switchSheet(name)}
              style={{ background: name === activeSheet ? '#fff' : '#3d4a5c', color: name === activeSheet ? T.text : '#a0aec0', border: '1px solid #4a5568', borderBottom: name === activeSheet ? '1px solid #fff' : '1px solid #4a5568', borderRadius: '4px 4px 0 0', padding: '5px 16px', fontSize: 12, cursor: 'pointer', fontFamily: 'inherit', whiteSpace: 'nowrap' }}>
              {name}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}


function BillcomSetup({entities,activeEntity,setActiveEntity}) {
  const[selectedEntity,setSelectedEntity]=useState(activeEntity||(entities[0]?entities[0].id:null));
  const[cfg,setCfg]=useState(null);
  const[loading,setLoading]=useState(false);
  const[saving,setSaving]=useState(false);
  const[testing,setTesting]=useState(false);
  const[msg,setMsg]=useState('');const[err,setErr]=useState('');
  const[env,setEnv]=useState('sandbox');
  const[username,setUsername]=useState('');
  const[password,setPassword]=useState('');
  const[orgId,setOrgId]=useState('');
  const[devKey,setDevKey]=useState('');
  const[defaultApAcct,setDefaultApAcct]=useState('');
  const[defaultCashAcct,setDefaultCashAcct]=useState('');

  // Phase 2: account mapping state
  const[tab,setTab]=useState('config'); // 'config' | 'mapping' | 'sync'
  const[bcAccounts,setBcAccounts]=useState([]);
  const[clAccounts,setClAccounts]=useState([]);
  const[mappings,setMappings]=useState({}); // keyed by billcom_account_id
  const[bcMeta,setBcMeta]=useState(null);
  const[mapLoading,setMapLoading]=useState(false);
  const[mapSaving,setMapSaving]=useState(false);
  const[mapPushing,setMapPushing]=useState(false);
  const[mapMsg,setMapMsg]=useState('');const[mapErr,setMapErr]=useState('');
  // Phase 3: sync state
  const[syncing,setSyncing]=useState(false);
  const[syncResult,setSyncResult]=useState(null);
  const[syncLogs,setSyncLogs]=useState([]);
  const[syncLogsLoading,setSyncLogsLoading]=useState(false);
  const[syncMsg,setSyncMsg]=useState('');const[syncErr,setSyncErr]=useState('');

  const loadMapping=useCallback(async()=>{
    if(!selectedEntity)return;
    setMapLoading(true);setMapMsg('');setMapErr('');setBcMeta(null);
    try{
      const [cl,saved]=await Promise.all([
        api.getAccounts(selectedEntity),
        api.getBillcomMappings(selectedEntity),
      ]);
      setClAccounts(Array.isArray(cl)?cl:[]);
      const savedList = (saved && Array.isArray(saved.mappings)) ? saved.mappings : (Array.isArray(saved) ? saved : []);
      const m={};
      savedList.forEach(r=>{m[r.billcom_account_id]=r.cl_account_code;});
      setMappings(m);
      try{
        const r=await api.getBillcomAccounts(selectedEntity);
        setBcAccounts(Array.isArray(r.accounts)?r.accounts:[]);
        setBcMeta({count:r.count});
      }catch(e){setMapErr('Bill.com fetch failed: '+e.message);setBcAccounts([]);}
    }catch(e){setMapErr(e.message);}
    setMapLoading(false);
  },[selectedEntity]);

  const saveMappings=async()=>{
    if(!selectedEntity)return;
    setMapSaving(true);setMapMsg('');setMapErr('');
    try{
      const payload=bcAccounts
        .filter(a=>mappings[a.id])
        .map(a=>({billcom_account_id:a.id,billcom_account_name:a.name,cl_account_code:mappings[a.id]}));
      await api.saveBillcomMappings(selectedEntity,payload);
      setMapMsg('Saved '+payload.length+' mapping(s).');
    }catch(e){setMapErr(e.message);}
    setMapSaving(false);
  };

  const pushCoaToBillcom=async()=>{
    if(!selectedEntity)return;
    if(!window.confirm('Push all CloudLedger Expense accounts to Bill.com and auto-create mappings? Accounts already in Bill.com (by number or name) will be skipped.'))return;
    setMapPushing(true);setMapMsg('');setMapErr('');
    try{
      const r=await api.pushBillcomCoa(selectedEntity,{all_expenses:true});
      const pushed=(r.pushed||[]).length;
      const mappedOnly=(r.mapped_only||[]).length;
      const skipped=(r.skipped_existing||[]).length;
      const errs=(r.errors||[]).length;
      const parts=[];
      if(pushed>0)parts.push('Created '+pushed+' in Bill.com');
      if(mappedOnly>0)parts.push('Mapped '+mappedOnly+' existing');
      if(skipped>0)parts.push('Skipped '+skipped+' already mapped');
      if(errs>0)parts.push(errs+' error(s)');
      setMapMsg(parts.join(' | ')||'No changes.');
      if(errs>0)setMapErr('Errors: '+r.errors.map(e=>e.code+' '+e.name+': '+(e.error||e.status)).join('; '));
      await loadMapping();
    }catch(e){setMapErr('Push failed: '+e.message);}
    setMapPushing(false);
  };

  const loadSyncLogs=useCallback(async()=>{
    if(!selectedEntity)return;
    setSyncLogsLoading(true);
    try{
      const r=await api.getBillcomSyncLog(selectedEntity,50);
      setSyncLogs(Array.isArray(r.logs)?r.logs:[]);
    }catch(e){setSyncErr('Failed to load logs: '+e.message);}
    setSyncLogsLoading(false);
  },[selectedEntity]);

  const runSync=async()=>{
    if(!selectedEntity)return;
    setSyncing(true);setSyncMsg('');setSyncErr('');setSyncResult(null);
    try{
      const r=await api.syncBillcom(selectedEntity);
      setSyncResult(r);
      const b=r.bills||{},py=r.payments||{};
      setSyncMsg('Done. Bills: '+(b.synced||0)+' synced, '+(b.skipped||0)+' skipped, '+(b.errors||0)+' errors. Payments: '+(py.synced||0)+' synced, '+(py.skipped||0)+' skipped, '+(py.errors||0)+' errors.');
      loadSyncLogs();
    }catch(e){setSyncErr('Sync failed: '+e.message);}
    setSyncing(false);
  };

  const load=useCallback(async()=>{
    if(!selectedEntity)return;
    setLoading(true);setMsg('');setErr('');
    try{
      const r=await api.getBillcomConfig(selectedEntity);
      setCfg(r);
      if(r.configured){
        setEnv(r.environment||'sandbox');
        setUsername(r.username||'');
        setOrgId(r.org_id||'');
        setDefaultApAcct(r.default_ap_account||'');
        setDefaultCashAcct(r.default_cash_account||'');
        setPassword('');setDevKey('');
      }else{
        setEnv('sandbox');setUsername('');setOrgId('');setDefaultApAcct('');setDefaultCashAcct('');
        setPassword('');setDevKey('');
      }
    }catch(e){setErr(e.message);}
    setLoading(false);
  },[selectedEntity]);
  useEffect(()=>{load();},[load]);

  // Phase 2+3: reset mapping & sync tab when entity changes
  useEffect(()=>{
    setTab('config');
    setBcAccounts([]);setClAccounts([]);setMappings({});setBcMeta(null);
    setMapMsg('');setMapErr('');
    setSyncResult(null);setSyncLogs([]);setSyncMsg('');setSyncErr('');
  },[selectedEntity]);

  const save=async()=>{
    if(!selectedEntity){setErr('Select an entity first');return;}
    setSaving(true);setMsg('');setErr('');
    try{
      const body={environment:env,username,org_id:orgId,default_ap_account:defaultApAcct||null,default_cash_account:defaultCashAcct||null};
      if(password)body.password=password;
      if(devKey)body.dev_key=devKey;
      await api.saveBillcomConfig(selectedEntity,body);
      setMsg('Configuration saved.');
      setPassword('');setDevKey('');
      load();
    }catch(e){setErr(e.message);}
    setSaving(false);
  };

  const test=async()=>{
    if(!selectedEntity){setErr('Select an entity first');return;}
    setTesting(true);setMsg('');setErr('');
    try{
      const r=await api.testBillcomConnection(selectedEntity);
      setMsg('Connection successful. '+(r.message||''));
      load();
    }catch(e){setErr('Connection failed: '+e.message);load();}
    setTesting(false);
  };

  const remove=async()=>{
    if(!selectedEntity)return;
    if(!confirm('Delete Bill.com configuration for this entity?'))return;
    setMsg('');setErr('');
    try{await api.deleteBillcomConfig(selectedEntity);setMsg('Configuration deleted.');load();}
    catch(e){setErr(e.message);}
  };

  const en=entities.find(e=>e.id===selectedEntity);

  return(<div>
    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:20}}>
      <div>
        <div style={{fontSize:22,fontWeight:700,color:T.textBright}}>Bill.com Setup</div>
        <div style={{fontSize:13,color:T.textMuted,marginTop:4}}>Connect a CloudLedger entity to a Bill.com Organization for AP integration.</div>
      </div>
    </div>

    <div style={{...S.card,padding:20,marginBottom:20}}>
      <div style={{fontSize:12,fontWeight:600,color:T.textMuted,marginBottom:6}}>ENTITY</div>
      <select value={selectedEntity||''} onChange={e=>setSelectedEntity(parseInt(e.target.value)||null)} style={{...S.input,maxWidth:400}}>
        <option value="">-- Select entity --</option>
        {entities.map(e=><option key={e.id} value={e.id}>{e.name}</option>)}
      </select>
    </div>

    {selectedEntity&&(loading?<div style={{color:T.textMuted}}>Loading...</div>:<>
      {cfg&&cfg.configured&&<div style={{...S.card,padding:16,marginBottom:16,background:'#f0fdf4',border:'1px solid #86efac'}}>
        <div style={{fontSize:13,fontWeight:600,color:'#15803d'}}>Configured</div>
        <div style={{fontSize:12,color:T.textMuted,marginTop:6}}>
          Last tested: {cfg.last_tested_at?new Date(cfg.last_tested_at).toLocaleString():'never'}
          {cfg.last_test_status&&<> · Status: <b style={{color:cfg.last_test_status==='success'?'#15803d':T.red}}>{cfg.last_test_status}</b></>}
        </div>
        {cfg.last_test_message&&<div style={{fontSize:11,color:T.textMuted,marginTop:4,fontFamily:'monospace'}}>{cfg.last_test_message}</div>}
      </div>}

      {/* Phase 2: Tab bar */}
      <div style={{display:'flex',gap:0,borderBottom:'1px solid '+T.border,marginBottom:16}}>
        <button onClick={()=>setTab('config')} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='config'?'2px solid '+T.accent:'2px solid transparent',color:tab==='config'?T.textBright:T.textMuted,cursor:'pointer'}}>Config</button>
        <button onClick={()=>{setTab('mapping');if(cfg&&cfg.configured)loadMapping();}} disabled={!cfg||!cfg.configured} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='mapping'?'2px solid '+T.accent:'2px solid transparent',color:tab==='mapping'?T.textBright:T.textMuted,cursor:cfg&&cfg.configured?'pointer':'not-allowed',opacity:cfg&&cfg.configured?1:0.5}}>Account Mapping</button>
        <button onClick={()=>{setTab('sync');if(cfg&&cfg.configured)loadSyncLogs();}} disabled={!cfg||!cfg.configured} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='sync'?'2px solid '+T.accent:'2px solid transparent',color:tab==='sync'?T.textBright:T.textMuted,cursor:cfg&&cfg.configured?'pointer':'not-allowed',opacity:cfg&&cfg.configured?1:0.5}}>Sync</button>
      </div>

      {tab==='config'?<>
      <div style={{...S.card,padding:20}}>
        <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:16}}>Credentials</div>

        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
          <div>
            <label style={S.label}>Environment</label>
            <select value={env} onChange={e=>setEnv(e.target.value)} style={S.input}>
              <option value="sandbox">Sandbox (testing)</option>
              <option value="production">Production (live)</option>
            </select>
          </div>
          <div>
            <label style={S.label}>Bill.com Username (email)</label>
            <input type="text" value={username} onChange={e=>setUsername(e.target.value)} style={S.input} placeholder="user@example.com" autoComplete="new-password"/>
          </div>
          <div>
            <label style={S.label}>Password {cfg&&cfg.configured&&<span style={{fontWeight:400,color:T.textMuted}}>(stored: {cfg.password_masked||'***'} — leave blank to keep)</span>}</label>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} style={S.input} placeholder={cfg&&cfg.configured?'(unchanged)':'Bill.com password'} autoComplete="new-password"/>
          </div>
          <div>
            <label style={S.label}>Organization ID</label>
            <input type="text" value={orgId} onChange={e=>setOrgId(e.target.value)} style={S.input} placeholder="008..." autoComplete="new-password"/>
          </div>
          <div>
            <label style={S.label}>Developer Key {cfg&&cfg.configured&&<span style={{fontWeight:400,color:T.textMuted}}>(stored: {cfg.dev_key_masked||'***'} — leave blank to keep)</span>}</label>
            <input type="password" value={devKey} onChange={e=>setDevKey(e.target.value)} style={S.input} placeholder={cfg&&cfg.configured?'(unchanged)':'Developer key'} autoComplete="new-password"/>
          </div>
          <div>
            <label style={S.label}>Default AP Account</label>
            <input type="text" value={defaultApAcct} onChange={e=>setDefaultApAcct(e.target.value)} style={S.input} placeholder="e.g. 21000" autoComplete="new-password"/>
          </div>
          <div>
            <label style={S.label}>Default Cash Account</label>
            <input type="text" value={defaultCashAcct} onChange={e=>setDefaultCashAcct(e.target.value)} style={S.input} placeholder="e.g. 10000" autoComplete="new-password"/>
          </div>
        </div>

        <div style={{display:'flex',gap:10,marginTop:20,alignItems:'center'}}>
          <button style={S.btnP} onClick={save} disabled={saving}>{saving?'Saving...':'Save'}</button>
          {cfg&&cfg.configured&&<button style={S.btnS} onClick={test} disabled={testing}>{testing?'Testing...':'Test Connection'}</button>}
          {cfg&&cfg.configured&&<button style={{...S.btnS,color:T.red}} onClick={remove}>Delete Config</button>}
          {msg&&<span style={{color:'#15803d',fontSize:13}}>{msg}</span>}
          {err&&<span style={{color:T.red,fontSize:13}}>{err}</span>}
        </div>

        <div style={{fontSize:11,color:T.textMuted,marginTop:16,padding:12,background:T.bgElevated,borderRadius:T.radiusSm}}>
          <b>Where to find these:</b><br/>
          Username: your Bill.com login email<br/>
          Organization ID: Bill.com → Settings → Sync and Integrations → Manage Developer Keys (starts with 008)<br/>
          Developer Key: same page; click "Generate developer key" (Admin role required)<br/>
          Default AP Account: GL code where bills post and payments debit (e.g. 21000 Accounts Payable)<br/>
          Default Cash Account: GL code where payments credit (e.g. 10000 Cash)
        </div>
      </div>
      </>:tab==='mapping'?<>
      {/* Phase 2: Account Mapping branch */}
      <div style={{...S.card,padding:20}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:14}}>
          <div>
            <div style={{fontSize:14,fontWeight:600,color:T.textBright}}>Account Mapping</div>
            <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>Map Bill.com chart of accounts to CloudLedger GL accounts.{bcMeta?' '+bcMeta.count+' Bill.com account(s) loaded.':''}</div>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button style={S.btnS} onClick={loadMapping} disabled={mapLoading||mapPushing}>{mapLoading?'Loading...':'Refresh from Bill.com'}</button>
            <button style={S.btnS} onClick={pushCoaToBillcom} disabled={mapPushing||mapLoading||mapSaving} title="Create every CloudLedger Expense account in Bill.com and auto-map them">{mapPushing?'Pushing...':'Push CL COA to Bill.com'}</button>
            <button style={S.btnP} onClick={saveMappings} disabled={mapSaving||mapLoading}>{mapSaving?'Saving...':'Save Mappings'}</button>
          </div>
        </div>

        {mapErr&&<div style={{padding:10,marginBottom:12,background:'#fef2f2',border:'1px solid #fecaca',borderRadius:T.radiusSm,color:T.red,fontSize:12}}>{mapErr}</div>}
        {mapMsg&&<div style={{padding:10,marginBottom:12,background:'#f0fdf4',border:'1px solid #86efac',borderRadius:T.radiusSm,color:'#15803d',fontSize:12}}>{mapMsg}</div>}

        {mapLoading?<div style={{color:T.textMuted,padding:20,textAlign:'center'}}>Loading accounts...</div>:
         bcAccounts.length===0?<div style={{color:T.textMuted,padding:20,textAlign:'center',fontSize:13}}>No Bill.com accounts loaded. Click "Refresh from Bill.com" to fetch.</div>:
         <div style={{border:'1px solid '+T.border,borderRadius:T.radiusSm,overflow:'hidden'}}>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}>
            <thead>
              <tr style={{background:T.bgElevated}}>
                <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Bill.com Account</th>
                <th style={{textAlign:'left',padding:'10px 12px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase',width:'40%'}}>CloudLedger GL Account</th>
              </tr>
            </thead>
            <tbody>
              {bcAccounts.map(a=>(
                <tr key={a.id} style={{borderTop:'1px solid '+T.border}}>
                  <td style={{padding:'10px 12px',verticalAlign:'top'}}>
                    <div style={{fontWeight:600,color:T.textBright}}>{a.accountNumber?a.accountNumber+' — ':''}{a.name}</div>
                    {a.description&&<div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{a.description}</div>}
                  </td>
                  <td style={{padding:'10px 12px'}}>
                    <select value={mappings[a.id]||''} onChange={e=>setMappings({...mappings,[a.id]:e.target.value})} style={{...S.input,fontSize:13}}>
                      <option value="">-- Not mapped --</option>
                      {clAccounts.map(c=><option key={c.code} value={c.code}>{c.code} — {c.name}</option>)}
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
         </div>
        }
      </div>
      </>:<>
      {/* Phase 3: Sync branch */}
      <div style={{...S.card,padding:20}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:14}}>
          <div>
            <div style={{fontSize:14,fontWeight:600,color:T.textBright}}>Sync</div>
            <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>Pull approved bills and payments from Bill.com and create journal entries. Already-synced items are skipped.</div>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button style={S.btnS} onClick={loadSyncLogs} disabled={syncLogsLoading||syncing}>{syncLogsLoading?'Loading...':'Refresh Log'}</button>
            <button style={S.btnP} onClick={runSync} disabled={syncing}>{syncing?'Syncing...':'Sync Now'}</button>
          </div>
        </div>

        {syncErr&&<div style={{padding:10,marginBottom:12,background:'#fef2f2',border:'1px solid #fecaca',borderRadius:T.radiusSm,color:T.red,fontSize:12}}>{syncErr}</div>}
        {syncMsg&&<div style={{padding:10,marginBottom:12,background:'#f0fdf4',border:'1px solid #86efac',borderRadius:T.radiusSm,color:'#15803d',fontSize:12}}>{syncMsg}</div>}

        {syncResult&&syncResult.missing_mappings&&syncResult.missing_mappings.length>0&&<div style={{padding:12,marginBottom:14,background:'#fffbeb',border:'1px solid #fcd34d',borderRadius:T.radiusSm}}>
          <div style={{fontSize:13,fontWeight:600,color:'#92400e',marginBottom:6}}>Missing GL Mappings ({syncResult.missing_mappings.length})</div>
          <div style={{fontSize:12,color:'#92400e',marginBottom:8}}>These Bill.com accounts appeared in bills but aren't mapped to a CloudLedger GL account. Map them in the Account Mapping tab, then sync again.</div>
          <ul style={{margin:0,paddingLeft:20,fontSize:12,color:'#78350f'}}>
            {syncResult.missing_mappings.map(m=><li key={m.billcom_account_id}>{m.name} ({m.affected_bills} bill{m.affected_bills===1?'':'s'})</li>)}
          </ul>
        </div>}

        <div style={{fontSize:12,fontWeight:600,color:T.textMuted,textTransform:'uppercase',marginBottom:8}}>Recent Sync Log</div>
        {syncLogs.length===0?<div style={{color:T.textMuted,padding:20,textAlign:'center',fontSize:13}}>No sync activity yet. Click "Sync Now" to start.</div>:
         <div style={{border:'1px solid '+T.border,borderRadius:T.radiusSm,overflow:'hidden',maxHeight:400,overflowY:'auto'}}>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
            <thead style={{position:'sticky',top:0,background:T.bgElevated}}>
              <tr>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>When</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Type</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Bill.com ID</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Status</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Message</th>
              </tr>
            </thead>
            <tbody>
              {syncLogs.map(l=>(
                <tr key={l.id} style={{borderTop:'1px solid '+T.border}}>
                  <td style={{padding:'8px 10px',color:T.textMuted,whiteSpace:'nowrap'}}>{l.created_at?new Date(l.created_at).toLocaleString():''}</td>
                  <td style={{padding:'8px 10px'}}>{l.sync_type}</td>
                  <td style={{padding:'8px 10px',fontFamily:'monospace',fontSize:11}}>{l.billcom_id}</td>
                  <td style={{padding:'8px 10px'}}><span style={{color:l.status==='success'?'#15803d':T.red,fontWeight:600}}>{l.status}</span></td>
                  <td style={{padding:'8px 10px',color:T.textMuted}}>{l.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
         </div>}
      </div>
      </>}
    </>)}
  </div>);
}


// ═══ Workpapers Modal (entity file storage with folders) ═══
function WorkpapersModal({entity, user, onClose}){
  const[files,setFiles]=useState([]);
  const[folders,setFolders]=useState([]);
  const[curPath,setCurPath]=useState('');
  const[loading,setLoading]=useState(true);
  const[err,setErr]=useState('');
  const[msg,setMsg]=useState('');
  const[uploading,setUploading]=useState(false);
  const[newFolderMode,setNewFolderMode]=useState(false);
  const[newFolderName,setNewFolderName]=useState('');
  const[uploadTarget,setUploadTarget]=useState('');
  const[renamingFolder,setRenamingFolder]=useState(null);
  const[renameValue,setRenameValue]=useState('');
  const[replacingFileId,setReplacingFileId]=useState(null);
  const[editingFile,setEditingFile]=useState(null);
  const fileInputRef = useRef(null);
  const replaceInputRef = useRef(null);
  const canEdit = user.role === 'Admin' || user.role === 'Accountant';
  const isEditable = f => /\.(xlsx|xls|csv)$/i.test(f.original_name);

  // Keep upload target in sync with the folder the user is browsing
  useEffect(() => { setUploadTarget(curPath); }, [curPath]);

  const load = useCallback(async () => {
    setLoading(true); setErr('');
    try { const r = await api.getEntityFiles(entity.id); setFiles(r.files); setFolders(r.folders); }
    catch (e) { setErr(e.message); } finally { setLoading(false); }
  }, [entity.id]);
  useEffect(() => { load(); }, [load]);

  // Files in the current directory
  const currentFiles = files.filter(f => f.folder_path === curPath);
  // Direct child folders of the current directory
  const childFolders = useMemo(() => {
    const set = new Set();
    folders.forEach(fp => {
      if (curPath === '') { if (!fp.includes('/')) set.add(fp); }
      else if (fp.startsWith(curPath + '/')) { const rest = fp.slice(curPath.length + 1); if (!rest.includes('/')) set.add(fp); }
    });
    return Array.from(set).sort();
  }, [folders, curPath]);

  const rootLabel = entity.name + ' Workpapers';
  const breadcrumbs = useMemo(() => {
    if (!curPath) return [{ label: rootLabel, path: '' }];
    const parts = curPath.split('/');
    return [{ label: rootLabel, path: '' }, ...parts.map((_, i) => ({ label: parts[i], path: parts.slice(0, i + 1).join('/') }))];
  }, [curPath, rootLabel]);

  // All folder paths for the upload-target dropdown (root + every known folder)
  const allFolderPaths = useMemo(() => ['', ...folders], [folders]);

  const doUpload = async e => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    setErr(''); setMsg(''); setUploading(true);
    try {
      const r = await api.uploadEntityFiles(entity.id, fileList, uploadTarget);
      setMsg(r.uploaded + ' file(s) uploaded to ' + (uploadTarget || 'root'));
      load();
    } catch (ex) {
      setErr(ex.message || 'Upload failed');
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const renameFolder = async fp => {
    const trimmed = renameValue.trim();
    if (!trimmed || trimmed.includes('/')) { setErr('Folder name cannot be empty or contain slashes'); return; }
    const parentPath = fp.includes('/') ? fp.slice(0, fp.lastIndexOf('/')) : '';
    const newPath = parentPath ? parentPath + '/' + trimmed : trimmed;
    if (newPath === fp) { setRenamingFolder(null); return; }
    setErr(''); setMsg('');
    try {
      await api.renameEntityFolder(entity.id, fp, newPath);
      // If we're currently inside the renamed folder or a descendant, update curPath
      if (curPath === fp) setCurPath(newPath);
      else if (curPath.startsWith(fp + '/')) setCurPath(newPath + curPath.slice(fp.length));
      setRenamingFolder(null); setRenameValue('');
      setMsg('Folder renamed');
      load();
    } catch (ex) { setErr(ex.message); }
  };

  const createFolder = async () => {
    if (!newFolderName.trim()) { setErr('Folder name required'); return; }
    const fullPath = curPath ? curPath + '/' + newFolderName.trim() : newFolderName.trim();
    setErr(''); setMsg('');
    try { await api.createEntityFolder(entity.id, fullPath); setNewFolderMode(false); setNewFolderName(''); setMsg('Folder created'); load(); }
    catch (ex) { setErr(ex.message); }
  };

  const deleteFile = async f => {
    if (!confirm('Delete "' + f.original_name + '"?')) return;
    setErr(''); setMsg('');
    try { await api.deleteEntityFile(f.id); load(); }
    catch (ex) { setErr(ex.message); }
  };

  const deleteFolder = async fp => {
    if (!confirm('Delete folder "' + fp.split('/').pop() + '"? The folder must be empty.')) return;
    setErr(''); setMsg('');
    try { await api.deleteEntityFolder(entity.id, fp); load(); }
    catch (ex) { setErr(ex.message); }
  };

  const doReplace = async e => {
    const file = e.target.files && e.target.files[0];
    if (!file || !replacingFileId) return;
    setErr(''); setMsg(''); setUploading(true);
    try {
      await api.replaceEntityFile(replacingFileId, file);
      setMsg('File replaced successfully.');
      load();
    } catch (ex) {
      setErr(ex.message || 'Replace failed');
    } finally {
      setUploading(false);
      setReplacingFileId(null);
      if (replaceInputRef.current) replaceInputRef.current.value = '';
    }
  };

  const fmtPstDate = ts => {
    if (!ts) return '';
    return new Date(ts + (ts.includes('Z') || ts.includes('+') ? '' : 'Z')).toLocaleString('en-US', { timeZone: 'America/Los_Angeles', year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true });
  };

  return (<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox, maxWidth: 944, maxHeight: '92vh', display: 'flex', flexDirection: 'column'}} onClick={e => e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{marginBottom: 16}}>
      <div style={{fontSize: 18, fontWeight: 700, color: T.textBright, display: 'flex', alignItems: 'center', gap: 10}}>
        <span style={{fontSize: 22}}>📁</span> {entity.name} Workpapers
      </div>
    </div>
    {/* Breadcrumbs */}
    <div style={{display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap', marginBottom: 14, padding: '10px 14px', background: T.bgElevated, borderRadius: T.radiusSm, fontSize: 13}}>
      {breadcrumbs.map((b, i) => <span key={i} style={{display: 'flex', alignItems: 'center', gap: 6}}>
        {i > 0 && <span style={{color: T.textDim}}>&rsaquo;</span>}
        {i === breadcrumbs.length - 1
          ? <span style={{fontWeight: 600, color: T.textBright}}>{b.label}</span>
          : <button style={{background: 'none', border: 0, padding: 0, color: T.accent, cursor: 'pointer', fontSize: 13, fontFamily: 'inherit'}} onClick={() => setCurPath(b.path)}>{b.label}</button>}
      </span>)}
    </div>
    {/* Actions */}
    {canEdit && <div style={{display: 'flex', gap: 10, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center'}}>
      <input ref={fileInputRef} type="file" multiple style={{display: 'none'}} onChange={doUpload}/>
      <input ref={replaceInputRef} type="file" style={{display: 'none'}} onChange={doReplace}/>
      <button style={{...S.btnP, opacity: uploading ? 0.6 : 1, cursor: uploading ? 'not-allowed' : 'pointer'}} disabled={uploading} onClick={() => fileInputRef.current && fileInputRef.current.click()}>{uploading ? 'Uploading...' : '+ Upload Files'}</button>
      <div style={{display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: T.textMuted}}>
        <span>to:</span>
        <select style={{...S.inputSm, minWidth: 200, maxWidth: 320}} value={uploadTarget} onChange={e => setUploadTarget(e.target.value)} disabled={uploading}>
          {allFolderPaths.map(fp => <option key={fp || '__root__'} value={fp}>{fp ? rootLabel + ' / ' + fp : rootLabel + ' (root)'}</option>)}
        </select>
      </div>
      {!newFolderMode
        ? <button style={S.btnS} onClick={() => { setNewFolderMode(true); setNewFolderName(''); }}>+ New Folder</button>
        : <div style={{display: 'flex', gap: 6, alignItems: 'center'}}>
            <input style={{...S.inputSm, minWidth: 180}} placeholder="Folder name" value={newFolderName} onChange={e => setNewFolderName(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') createFolder(); if (e.key === 'Escape') setNewFolderMode(false); }} autoFocus/>
            <button style={{...S.btnS, padding: '6px 12px', fontSize: 11}} onClick={createFolder}>Create</button>
            <button style={{...S.btnGhost, fontSize: 11}} onClick={() => setNewFolderMode(false)}>Cancel</button>
          </div>}
    </div>}
    {err && <div style={{...S.err, marginBottom: 10}}>{err}</div>}
    {msg && <div style={{...S.success, marginBottom: 10}}>{msg}</div>}
    {/* File & folder list */}
    <div style={{flex: '0 1 auto', minHeight: 160, overflowY: 'auto', border: '1px solid ' + T.border, borderRadius: T.radiusSm}}>
      {loading ? <div style={{padding: 60, textAlign: 'center', color: T.textMuted}}>Loading...</div>
       : (childFolders.length === 0 && currentFiles.length === 0)
         ? <div style={{padding: 60, textAlign: 'center', color: T.textDim}}>This folder is empty{canEdit ? '. Upload files or create a subfolder to get started.' : '.'}</div>
         : <table style={S.table}>
            <thead style={{position: 'sticky', top: 0, background: T.bgCard, zIndex: 1}}>
              <tr><th style={S.th}>Name</th><th style={S.thR} width={110}>Size</th><th style={S.th} width={180}>Uploaded By</th><th style={S.th} width={180}>Date (PST)</th><th style={S.th} width={120}></th></tr>
            </thead>
            <tbody>
              {childFolders.map(fp => <tr key={'d-' + fp} style={{background: T.bgElevated}}>
                <td style={S.td}>
                  {renamingFolder === fp
                    ? <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
                        <span style={{fontSize: 16}}>📁</span>
                        <input style={{...S.inputSm, minWidth: 220}} value={renameValue} onChange={e => setRenameValue(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') renameFolder(fp); if (e.key === 'Escape') { setRenamingFolder(null); setRenameValue(''); } }} autoFocus/>
                        <button style={{...S.btnS, padding: '4px 10px', fontSize: 11}} onClick={() => renameFolder(fp)}>Save</button>
                        <button style={{...S.btnGhost, fontSize: 11}} onClick={() => { setRenamingFolder(null); setRenameValue(''); }}>Cancel</button>
                      </div>
                    : <button style={{background: 'none', border: 0, padding: 0, color: T.accent, fontWeight: 600, cursor: 'pointer', fontSize: 13, fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 8}} onClick={() => setCurPath(fp)}>
                        <span style={{fontSize: 16}}>📁</span> {fp.split('/').pop()}
                      </button>}
                </td>
                <td style={S.tdR}></td><td style={S.td}></td><td style={S.td}></td>
                <td style={S.td}>{canEdit && renamingFolder !== fp && <div style={{display: 'flex', gap: 6}}>
                  <button style={{...S.btnGhost, color: T.accent, fontSize: 11}} onClick={() => { setRenamingFolder(fp); setRenameValue(fp.split('/').pop()); setErr(''); }}>Rename</button>
                  <button style={{...S.btnGhost, color: T.red, fontSize: 11}} onClick={() => deleteFolder(fp)}>Delete</button>
                </div>}</td>
              </tr>)}
              {currentFiles.map(f => <tr key={'f-' + f.id}>
                <td style={S.td}>
                  {canEdit && isEditable(f)
                    ? <button onClick={() => setEditingFile(f)} title="Click to edit in browser"
                        style={{background: 'none', border: 0, padding: 0, color: T.textBright, cursor: 'pointer', fontFamily: 'inherit', fontSize: 13, textAlign: 'left', display: 'flex', alignItems: 'center', gap: 8}}
                        onMouseEnter={e => e.currentTarget.style.color = T.accent}
                        onMouseLeave={e => e.currentTarget.style.color = T.textBright}>
                        <span style={{fontSize: 14}}>📄</span> {f.original_name}
                      </button>
                    : <a href={api.downloadEntityFile(f.id)} target="_blank" rel="noreferrer" style={{color: T.textBright, textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 8}} onMouseEnter={e => e.currentTarget.style.color = T.accent} onMouseLeave={e => e.currentTarget.style.color = T.textBright}>
                        <span style={{fontSize: 14}}>📄</span> {f.original_name}
                      </a>}
                </td>
                <td style={{...S.tdR, color: T.textMuted, fontSize: 12}}>{fmtSize(f.size)}</td>
                <td style={{...S.td, color: T.textMuted, fontSize: 12}}>{f.uploaded_by}</td>
                <td style={{...S.td, color: T.textMuted, fontSize: 12}}>{fmtPstDate(f.created_at)}</td>
                <td style={S.td}><div style={{display: 'flex', gap: 6}}>
                  <a href={api.downloadEntityFile(f.id)} target="_blank" rel="noreferrer" style={{...S.btnGhost, color: T.accent, fontSize: 11, textDecoration: 'none'}}>Download</a>
                  {canEdit && <button style={{...S.btnGhost, color: T.textMuted, fontSize: 11}} disabled={uploading} onClick={() => { setReplacingFileId(f.id); replaceInputRef.current && replaceInputRef.current.click(); }}>Replace</button>}
                  {canEdit && <button style={{...S.btnGhost, color: T.red, fontSize: 11}} onClick={() => deleteFile(f)}>Delete</button>}
                </div></td>
              </tr>)}
            </tbody>
          </table>}
    </div>
  {editingFile && <SpreadsheetEditorModal file={editingFile} onClose={() => setEditingFile(null)} onSaved={() => { setEditingFile(null); load(); }} />}
  </div></div>);
}

// ═══ Dashboard ═══
function Dashboard({entityId,setActiveEntity,setPage,user}){const[summary,setSummary]=useState([]);useEffect(()=>{api.getSummary().then(setSummary);},[]);const curr=summary.find(e=>e.id===entityId);
  const[wpEntity,setWpEntity]=useState(null);
  const go=id=>{setActiveEntity(id);setPage('journal');};
  return(<div><div style={S.h1}>Dashboard</div><div style={S.sub}>{summary.length} entities under management &middot; click a row to open, click the folder icon for workpapers</div>
    <div style={S.cardFlush}><div style={{padding:'18px 20px'}}><div style={S.h2}>All Entities</div></div><table style={S.table}><thead><tr><th style={{...S.th,width:40}}></th><th style={S.th}>Entity</th><th style={S.thR}>Assets</th><th style={S.thR}>Liabilities</th><th style={S.thR}>Net Income</th><th style={S.thR}>JEs</th></tr></thead>
      <tbody>{summary.sort((a,b)=>a.name.localeCompare(b.name)).map(e=><tr key={e.id} style={{cursor:'pointer',background:e.id===entityId?T.accentDim:'transparent',transition:'background 0.1s'}} onMouseEnter={ev=>{if(e.id!==entityId)ev.currentTarget.style.background=T.bgHover;}} onMouseLeave={ev=>{if(e.id!==entityId)ev.currentTarget.style.background='transparent';}}>
        <td style={{...S.td,textAlign:'center',padding:'8px 6px'}} onClick={ev=>{ev.stopPropagation();setWpEntity(e);}} title="Open workpapers folder"><span style={{fontSize:18,cursor:'pointer',display:'inline-block',lineHeight:1}}>📁</span></td>
        <td style={{...S.td,fontWeight:600,color:T.accent,textDecoration:'underline'}} onClick={()=>go(e.id)}>{e.name}</td>
        <td style={S.tdR} onClick={()=>go(e.id)}>{fmt(e.assets)}</td>
        <td style={S.tdR} onClick={()=>go(e.id)}>{fmt(e.liabilities)}</td>
        <td style={{...S.tdR,color:e.net_income>=0?T.green:T.red,fontWeight:600}} onClick={()=>go(e.id)}>{fmt(e.net_income)}</td>
        <td style={S.tdR} onClick={()=>go(e.id)}>{e.entry_count}</td>
      </tr>)}</tbody></table></div>
    {wpEntity&&<WorkpapersModal entity={wpEntity} user={user} onClose={()=>setWpEntity(null)}/>}
  </div>);}

// ═══ Edit JE Modal ═══
function EditJEModal({entityId,dimsEnabled,entry,accounts:initAccounts,onClose,onSaved}){
  const[accounts,setAccounts]=useState(initAccounts||[]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[saving,setSaving]=useState(false);
  const[projects,setProjects]=useState([]);useEffect(()=>{api.getTurnkeyProjects().then(setProjects).catch(()=>setProjects([]));},[entityId]);const showProject=projects.length>0||(entry.lines||[]).some(l=>l.project_id);
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);
  useEffect(()=>{api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  const showLocation=(dimsEnabled&&locations.length>0)||(entry.lines||[]).some(l=>l.location_id);const showClass=(dimsEnabled&&classes.length>0)||(entry.lines||[]).some(l=>l.class_id);
  const[form,setForm]=useState({date:entry.date,memo:entry.memo,lines:entry.lines.map(l=>({account_code:l.account_code,project_id:l.project_id||'',location_id:l.location_id||'',class_id:l.class_id||'',description:l.description||'',debit:l.debit>0?l.debit.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):'',credit:l.credit>0?l.credit.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):''}))});
  const[attachments,setAttachments]=useState(entry.attachments||[]);
  const[attUploading,setAttUploading]=useState(false);
  const attInputRef=useRef(null);
  useEffect(()=>{if(!initAccounts?.length)api.getAccounts(entityId).then(setAccounts);},[entityId,initAccounts]);
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:'',description:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  const tDr=form.lines.reduce((s,l)=>s+parseAmt(l.debit),0);const tCr=form.lines.reduce((s,l)=>s+parseAmt(l.credit),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;
  const save=async()=>{if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}if(form.lines.some(l=>!l.account_code)){setErr('All lines need an account');return;}if(!bal){setErr('Must balance');return;}
    setSaving(true);setErr('');try{await api.updateEntry(entityId,entry.id,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseAmt(l.debit),credit:parseAmt(l.credit),description:l.description||'',project_id:l.project_id||null,location_id:l.location_id||null,class_id:l.class_id||null}))});
      onSaved();onClose();}catch(e){setErr(e.message);}finally{setSaving(false);}};
  const uploadAtt=async e=>{const fl=e.target.files;if(!fl||fl.length===0)return;setErr('');setAttUploading(true);
    try{const r=await api.uploadAttachments(entityId,entry.id,fl);setAttachments(p=>[...p,...(r.attachments||r.files||r||[])]);}
    catch(ex){setErr(ex.message);}finally{setAttUploading(false);if(attInputRef.current)attInputRef.current.value='';}};
  const deleteAtt=async a=>{if(!confirm('Delete '+a.original_name+'?'))return;try{await api.deleteAttachment(a.id);setAttachments(p=>p.filter(x=>x.id!==a.id));}catch(ex){setErr(ex.message);}};
  const fmtPst=ts=>ts?new Date(ts+(ts.includes('Z')||ts.includes('+')?'':'Z')).toLocaleString('en-US',{timeZone:'America/Los_Angeles',year:'numeric',month:'short',day:'numeric',hour:'numeric',minute:'2-digit',hour12:true,timeZoneName:'short'}):'';
  return(<div style={S.modal}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:960}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Edit JE-{String(entry.entry_num).padStart(4,'0')}</div>
    {(entry.created_by||entry.created_at)&&<div style={{fontSize:11,color:T.textMuted,marginBottom:2}}>
      Posted{entry.created_by?' by '+entry.created_by:''}{entry.created_at?' on '+fmtPst(entry.created_at):''}
    </div>}
    {entry.updated_by&&entry.updated_at&&<div style={{fontSize:11,color:T.orange,marginBottom:16,fontStyle:'italic'}}>
      Last edited by {entry.updated_by} on {fmtPst(entry.updated_at)}
    </div>}
    {!(entry.updated_by&&entry.updated_at)&&<div style={{marginBottom:16}}/>}
    <div style={{background:T.bgElevated,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:18,marginBottom:16}}>
      <div style={S.row}><div style={{...S.col,maxWidth:170}}><label style={S.label}>Date</label><input style={S.input} type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/></div>
        <div style={{...S.col,flex:4}}><label style={S.label}>Memo</label><input style={S.input} value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))}/></div></div></div>
    <div style={{...S.cardFlush,marginBottom:16}}><table style={S.table}><thead><tr><th style={S.th}>Account</th>{showProject&&<th style={{...S.th,width:170}}>Project</th>}{showLocation&&<th style={{...S.th,width:150}}>Location</th>}{showClass&&<th style={{...S.th,width:150}}>Class</th>}<th style={S.th}>Description</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        {showProject&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.project_id||''} onChange={e=>updateLine(i,'project_id',e.target.value)}><option value="">— none —</option>{projects.map(pr=><option key={pr.turnkey_project_id} value={pr.turnkey_project_id}>{pr.project_code} — {pr.project_name}</option>)}</select></td>}
        {showLocation&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.location_id||''} onChange={e=>updateLine(i,'location_id',e.target.value)}><option value="">— none —</option>{locations.map(loc=><option key={loc.id} value={loc.id}>{loc.code?loc.code+" — ":""}{loc.name}</option>)}</select></td>}
        {showClass&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={l.class_id||''} onChange={e=>updateLine(i,'class_id',e.target.value)}><option value="">— none —</option>{classes.map(c=><option key={c.id} value={c.id}>{c.code?c.code+" — ":""}{c.name}</option>)}</select></td>}
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={S.input} value={l.description||''} placeholder="(optional)" onChange={e=>updateLine(i,'description',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.debit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'debit',f);}} onBlur={e=>updateLine(i,'debit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.credit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'credit',f);}} onBlur={e=>updateLine(i,'credit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td colSpan={2+(showProject?1:0)+(showLocation?1:0)+(showClass?1:0)} style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
    {/* Attachments */}
    <div style={{marginBottom:16}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:8}}>
        <div style={{fontSize:12,fontWeight:600,color:T.textBright,textTransform:'uppercase',letterSpacing:0.4}}>Attachments {attachments.length>0&&<span style={{color:T.textMuted,fontWeight:500}}>({attachments.length})</span>}</div>
        <div>
          <input ref={attInputRef} type="file" multiple style={{display:'none'}} onChange={uploadAtt}/>
          <button style={{...S.btnS,fontSize:11,padding:'6px 12px',color:T.accent,borderColor:T.accent+'40'}} disabled={attUploading} onClick={()=>attInputRef.current&&attInputRef.current.click()}>{attUploading?'Uploading...':'+ Attach Files'}</button>
        </div>
      </div>
      {attachments.length===0?<div style={{fontSize:12,color:T.textDim,padding:'10px 14px',background:T.bgElevated,borderRadius:T.radiusSm,border:'1px dashed '+T.border,textAlign:'center'}}>No attachments</div>:
      <div style={{display:'flex',flexDirection:'column',gap:6}}>{attachments.map(a=><div key={a.id} style={{display:'flex',alignItems:'center',gap:10,padding:'8px 12px',background:T.bgElevated,borderRadius:T.radiusSm,border:'1px solid '+T.border}}>
        <span style={{fontSize:14}}>📎</span>
        <a href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={{color:T.accent,fontSize:12,fontWeight:500,textDecoration:'none',flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{a.original_name}</a>
        <span style={{fontSize:11,color:T.textMuted}}>{fmtSize(a.size||0)}</span>
        <button style={{...S.btnGhost,color:T.red,fontSize:10}} onClick={()=>deleteAtt(a)}>Delete</button>
      </div>)}</div>}
    </div>
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

// ═══ Bulk Journal Entry Upload ═══
// Upload an .xlsx/.csv where each row is one balanced entry
// (Date, Memo, Debit Account #, Credit Account #, Amount [, Line Description, Location, Class]),
// preview + validate, then post the valid rows.
function BulkJEModal({entityId,onClose,onPosted}){
  const[file,setFile]=useState(null);
  const[preview,setPreview]=useState(null);// {entries, mapped, total, valid, invalid, line_count}
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[posted,setPosted]=useState(null);
  const doPreview=async(f)=>{
    setErr('');setPreview(null);setBusy(true);
    try{const r=await api.bulkEntriesPreview(entityId,f);setPreview(r);}
    catch(e){setErr(e.message);}
    finally{setBusy(false);}
  };
  const onPick=e=>{const f=e.target.files[0];e.target.value='';if(f){setFile(f);doPreview(f);}};
  const commit=async()=>{
    if(!preview)return;
    const valid=preview.entries.filter(en=>en.valid).map(en=>({date:en.date,memo:en.memo,lines:en.lines.map(l=>({account_code:l.account_code,debit:l.debit,credit:l.credit,location_id:l.location_id,class_id:l.class_id}))}));
    if(!valid.length){setErr('No valid entries to post.');return;}
    setBusy(true);setErr('');
    try{const r=await api.bulkEntriesCommit(entityId,valid);setPosted(r.posted);setTimeout(()=>onPosted(),900);}
    catch(e){setErr(e.message);}
    finally{setBusy(false);}
  };
  const validCount=preview?preview.entries.filter(e=>e.valid).length:0;
  return(<div style={S.modal} onClick={()=>{if(!busy)onClose();}}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:920}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:6}}>Bulk Upload Journal Entries</div>
    <div style={{fontSize:13,color:T.textMuted,marginBottom:16,lineHeight:1.6}}>One row per journal line. Lines sharing the same <strong>Date</strong> are grouped into one entry, which must balance. Required columns: <strong>Date</strong>, <strong>Account #</strong>, <strong>Debit</strong>, <strong>Credit</strong>. Optional: Account Description, Memo, Location, Class. Accepts .xlsx or .csv.</div>

    <div style={{position:'relative',border:'1.5px dashed '+T.border,borderRadius:8,padding:'20px 16px',textAlign:'center',background:T.bgElevated,marginBottom:14}}>
      <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:2}}>{file?file.name:'Click to choose a spreadsheet'}</div>
      <div style={{fontSize:11,color:T.textMuted}}>{file?'Click again to choose a different file':'.xlsx or .csv'}</div>
      <input type="file" accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" disabled={busy} style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:busy?'not-allowed':'pointer'}} onChange={onPick}/>
    </div>

    {busy&&!posted&&<div style={{fontSize:12,color:T.accent,margin:'8px 0'}}>Working&hellip;</div>}
    {err&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:12}}>{err}</div>}
    {posted!=null&&<div style={{fontSize:14,fontWeight:600,color:T.green,padding:12,background:T.greenDim,borderRadius:6,border:'1px solid '+T.greenBorder}}>Posted {posted} journal {posted===1?'entry':'entries'}. ✓</div>}

    {preview&&posted==null&&<div>
      <div style={{display:'flex',gap:16,alignItems:'center',marginBottom:12,fontSize:13,flexWrap:'wrap'}}>
        <span style={{fontWeight:700,color:T.textBright}}>{preview.line_count} line{preview.line_count===1?'':'s'} &rarr; {preview.total} entr{preview.total===1?'y':'ies'}</span>
        <span style={{color:T.green,fontWeight:600}}>{preview.valid} valid</span>
        {preview.invalid>0&&<span style={{color:T.red,fontWeight:600}}>{preview.invalid} with errors (skipped)</span>}
      </div>
      <div style={{maxHeight:360,overflow:'auto',border:'1px solid '+T.border,borderRadius:8}}>
        {preview.entries.map((en,ei)=><div key={ei} style={{borderBottom:'1px solid '+T.borderLight,padding:'10px 12px',background:en.valid?'#fff':T.redDim}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6,gap:10}}>
            <div style={{display:'flex',alignItems:'center',gap:10,minWidth:0}}>
              <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>{en.date||<span style={{color:T.red}}>no date</span>}</span>
              <span style={{fontSize:12,color:T.textMuted,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={en.memo}>{en.memo}</span>
              <span style={{fontSize:11,color:T.textDim}}>rows {en.rows.join(', ')}</span>
            </div>
            {en.valid?<span style={{color:T.green,fontWeight:600,fontSize:11,whiteSpace:'nowrap'}}>OK</span>
              :<span style={{color:T.red,fontSize:11,maxWidth:300,textAlign:'right'}} title={en.errors.join('; ')}>{en.errors.join('; ')}</span>}
          </div>
          <table style={{...S.table,width:'100%',tableLayout:'fixed'}}>
            <colgroup><col style={{width:'110px'}}/><col/><col style={{width:'120px'}}/><col style={{width:'120px'}}/></colgroup>
            <tbody>{en.lines.map((l,li)=><tr key={li}>
              <td style={{...S.td,fontSize:12}}>{l.account_code}</td>
              <td style={{...S.td,fontSize:12,color:T.textMuted,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={l.account_name+(l.errors&&l.errors.length?(' — '+l.errors.join('; ')):'')}>{l.account_name||''}{l.errors&&l.errors.length>0&&<span style={{color:T.red}}> ⚠ {l.errors.join('; ')}</span>}</td>
              <td style={{...S.tdR,fontSize:12}}>{l.debit>0?fmt(l.debit):''}</td>
              <td style={{...S.tdR,fontSize:12}}>{l.credit>0?fmt(l.credit):''}</td>
            </tr>)}
            <tr style={{borderTop:'1px solid '+T.border}}>
              <td style={{...S.td,fontSize:11,fontWeight:700}} colSpan={2}>Totals</td>
              <td style={{...S.tdR,fontSize:11,fontWeight:700,color:Math.abs(en.total_debit-en.total_credit)<=0.005?T.green:T.red}}>{fmt(en.total_debit)}</td>
              <td style={{...S.tdR,fontSize:11,fontWeight:700,color:Math.abs(en.total_debit-en.total_credit)<=0.005?T.green:T.red}}>{fmt(en.total_credit)}</td>
            </tr></tbody>
          </table>
        </div>)}
      </div>
      <div style={{display:'flex',gap:10,marginTop:16,justifyContent:'flex-end'}}>
        <button style={S.btnS} onClick={onClose} disabled={busy}>Cancel</button>
        <button style={{...S.btnP,opacity:validCount&&!busy?1:0.5}} disabled={!validCount||busy} onClick={commit}>{busy?'Posting…':'Post '+validCount+' '+(validCount===1?'Entry':'Entries')}</button>
      </div>
    </div>}
  </div></div>);
}

// ═══ Journal List ═══
function JournalList({entityId,entityName,dimsEnabled,canEdit=true,onNewEntry}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[from,setFrom]=useState('');const[to,setTo]=useState('');
  const[editEntry,setEditEntry]=useState(null);const[showBulk,setShowBulk]=useState(false);
  const load=useCallback(async()=>{const[e,a]=await Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]);setEntries(e);setAccounts(a);},[entityId,from,to]);
  useEffect(()=>{load();},[load]);const del=async id=>{if(!confirm('Delete this journal entry?'))return;await api.deleteEntry(entityId,id);load();};const acctName=code=>accounts.find(a=>a.code===code)?.name||'?';
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}><div><div style={S.h1}>Journal Entries</div><div style={S.sub}>{entityName} &middot; {entries.length} entries{!canEdit&&' · read-only'}</div></div>{canEdit&&<div style={{display:'flex',gap:8}}><button style={S.btnS} onClick={()=>setShowBulk(true)}>Bulk Upload</button><button style={S.btnP} onClick={onNewEntry}>+ New Entry</button></div>}</div>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      {(from||to)&&<button style={{...S.btnGhost,marginTop:14,color:T.red}} onClick={()=>{setFrom('');setTo('');}}>Clear</button>}</div>
    {entries.length===0?<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No entries found</div>:
      <div style={{display:'flex',flexDirection:'column',gap:8}}>{entries.map(e=><div key={e.id} style={{...S.card,padding:14,marginBottom:0}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
          <div style={{display:'flex',alignItems:'center',gap:14}}><span style={{fontWeight:700,color:T.accent,fontSize:14}}>JE-{String(e.entry_num).padStart(4,'0')}</span>
            <span style={{color:T.textMuted}}>{e.date}</span><span style={{fontWeight:500}}>{e.memo}</span>
            {e.attachments?.length>0&&<span style={{fontSize:11,color:T.teal,fontWeight:500}}>({e.attachments.length} file{e.attachments.length>1?'s':''})</span>}</div>
          <div style={{display:'flex',alignItems:'center',gap:8}}>
            <span style={{fontSize:12,color:T.textDim}}>{e.created_by}</span>
            {canEdit&&<button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>setEditEntry(e)}>Edit</button>}
            {canEdit&&<button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={()=>del(e.id)}>Delete</button>}</div></div>
        <table style={{...S.table,tableLayout:'fixed',width:'100%'}}>
          <colgroup><col style={{width:'280px'}}/><col/><col style={{width:'140px'}}/><col style={{width:'140px'}}/></colgroup>
          <thead><tr><th style={S.th}>Account</th><th style={S.th}>Description</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
          <tbody>{e.lines.map((l,i)=><tr key={i}><td style={S.td} title={acctLabel(l.account_code,acctName(l.account_code))}>{acctLabel(l.account_code,acctName(l.account_code))}</td>
            <td style={{...S.td,color:T.textMuted,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={l.description||''}>{l.description||''}</td>
            <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td><td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td></tr>)}</tbody></table>
        {e.attachments?.length>0&&<div style={{marginTop:10,display:'flex',flexWrap:'wrap',gap:4}}>{e.attachments.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>)}</div>}
      </div>)}</div>}
    {editEntry&&<EditJEModal entityId={entityId} dimsEnabled={dimsEnabled} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={load}/>}
    {showBulk&&<BulkJEModal entityId={entityId} onClose={()=>setShowBulk(false)} onPosted={()=>{setShowBulk(false);load();}}/>}
  </div>);}

// ═══ Dimensions (Locations & Classes) manager ═══
function DimList({title,subtitle,items,canEdit,onCreate,onUpdate,onDelete}){
  const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({code:'',name:''});const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState({code:'',name:''});const[editErr,setEditErr]=useState('');
  const startEdit=it=>{setEditing(it.id);setEditForm({code:it.code||'',name:it.name||''});setEditErr('');};
  const add=async()=>{if(!form.name.trim()){setErr('Name required');return;}try{await onCreate({name:form.name.trim(),code:form.code.trim()||null});setForm({code:'',name:''});setShowAdd(false);setErr('');}catch(e){setErr(e.message);}};
  const save=async()=>{if(!editForm.name.trim()){setEditErr('Name required');return;}try{await onUpdate(editing,{name:editForm.name.trim(),code:editForm.code.trim()||null});setEditing(null);}catch(e){setEditErr(e.message);}};
  const del=async it=>{if(!confirm('Delete "'+it.name+'"?'))return;try{await onDelete(it.id);}catch(e){alert(e.message);}};
  return(<div style={{flex:1,minWidth:340}}>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
      <div><div style={{fontSize:15,fontWeight:700,color:T.textBright}}>{title}</div><div style={{fontSize:12,color:T.textMuted}}>{subtitle||(items.length+' total')}</div></div>
      {canEdit&&<button style={{...S.btnP,padding:'6px 12px',fontSize:12}} onClick={()=>{setShowAdd(!showAdd);setErr('');}}>{showAdd?'Cancel':'+ Add'}</button>}</div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40',padding:14,marginBottom:12}}><div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
      <div style={{flex:1}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')add();}}/></div>
      <button style={S.btnP} onClick={add}>Add</button></div>{err&&<div style={{...S.err,marginTop:8,marginBottom:0}}>{err}</div>}</div>}
    <div style={S.cardFlush}><table style={{...S.table,tableLayout:'fixed'}}><thead><tr><th style={S.th}>Name</th>{canEdit&&<th style={{...S.th,width:84}}>Actions</th>}</tr></thead>
      <tbody>{items.length===0&&<tr><td colSpan={canEdit?2:1} style={{...S.td,color:T.textMuted,textAlign:'center',padding:'18px'}}>None yet</td></tr>}
      {items.map(it=>editing===it.id?
        <tr key={it.id} style={{background:T.accentDim}}>
          <td style={{padding:'6px 8px'}}><input style={S.input} value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')save();}}/></td>
          {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnGhost,color:T.green,fontSize:11}} onClick={save}>Save</button><button style={{...S.btnGhost,fontSize:11}} onClick={()=>setEditing(null)}>Cancel</button></div></td>}
        </tr>
        :<tr key={it.id}>
          <td style={S.td} title={it.name}>{it.name}</td>
          {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
            <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(it)}>Edit</button>
            <button style={{...S.btnGhost,color:it.line_count>0?T.textMuted:T.red,fontSize:11}} title={it.line_count>0?'Used on '+it.line_count+' line(s) — cannot delete':'Delete'} onClick={()=>del(it)}>x</button></div></td>}
        </tr>)}
      </tbody></table></div>
    {editErr&&<div style={{...S.err,marginTop:8}}>{editErr}</div>}</div>);
}
function DimensionsManager({entityId,entityName,canEdit}){
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);
  const load=useCallback(async()=>{const[l,c]=await Promise.all([api.getLocations(entityId),api.getClasses(entityId)]);setLocations(l||[]);setClasses(c||[]);},[entityId]);
  useEffect(()=>{load();},[load]);
  return(<div><div style={{marginBottom:20}}><div style={S.h1}>Locations & Classes</div><div style={S.sub}>{entityName} — dimensions you can tag on journal-entry lines and filter reports by</div></div>
    <div style={{display:'flex',gap:24,flexWrap:'wrap',alignItems:'flex-start'}}>
      <DimList title="Locations" subtitle={(locations.length)+' location'+(locations.length===1?'':'s')+' (deals / properties)'} items={locations} canEdit={canEdit}
        onCreate={async d=>{await api.createLocation(entityId,d);await load();}}
        onUpdate={async(id,d)=>{await api.updateLocation(entityId,id,d);await load();}}
        onDelete={async id=>{await api.deleteLocation(entityId,id);await load();}}/>
      <DimList title="Investor Classes" subtitle={(classes.length)+' class'+(classes.length===1?'':'es')+' (investors / capital classes)'} items={classes} canEdit={canEdit}
        onCreate={async d=>{await api.createClass(entityId,d);await load();}}
        onUpdate={async(id,d)=>{await api.updateClass(entityId,id,d);await load();}}
        onDelete={async id=>{await api.deleteClass(entityId,id);await load();}}/>
    </div></div>);
}

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
        <td style={{...S.td,color:T.textBright}}>{a.code}</td><td style={S.td}>{a.name}</td><td style={S.td}><span style={S.tag(a.type)}>{a.type}</span></td>
        <td style={S.tdC}>{a.bank_acct?<span style={{color:T.green}}>Yes</span>:''}</td>
        {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
          <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(a)}>Edit</button>
          <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={async()=>{try{await api.deleteAccount(entityId,a.code);load();}catch(e){alert(e.message);}}}>x</button></div></td>}</tr>)}</tbody></table></div></div>);}

// ═══ General Ledger ═══
function GeneralLedger({entityId,entityName,dimsEnabled,from,setFrom,to,setTo,filter,setFilter}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);
  const[editEntry,setEditEntry]=useState(null);
  const reload=useCallback(()=>{Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]).then(([e,a])=>{setEntries(e);setAccounts(a);});},[entityId,from,to]);
  useEffect(()=>{reload();},[reload]);
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
        <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:900}}><thead><tr><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th><th style={{...S.th,width:100}}>Docs</th></tr></thead>
          <tbody>{txns.map((t,i)=>{run+=dr?(t.debit-t.credit):(t.credit-t.debit);const atts=entryAtts[t.jeId];return<tr key={i}><td style={{...S.td,color:T.textMuted}}>{t.date}</td><td style={S.td}><button style={{background:'none',border:0,padding:0,color:T.accent,fontWeight:600,cursor:'pointer',fontSize:'inherit',fontFamily:'inherit'}} onClick={()=>{const e=entries.find(x=>x.id===t.jeId);if(e)setEditEntry(e);}}>JE-{String(t.jeNum).padStart(4,'0')}</button></td><td style={S.td} title={t.description||t.memo}>{t.description||t.memo}</td><td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(run)}</td>
            <td style={S.td}>{atts?atts.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>):''}</td></tr>;})}</tbody></table></div>}</div>);})}
    {editEntry&&<EditJEModal entityId={entityId} dimsEnabled={dimsEnabled} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={()=>{setEditEntry(null);reload();}}/>}
    </div>);}

// ═══ Bank Transactions (state lifted to App for navigation persistence) ═══
// ═══ Bank Transaction Split Modal ═══
function SplitBankTransactionModal({txn, accounts, excludeCode, entityId, onClose, onSaved}){
  const target = Math.abs(txn.amount);
  const initialLines = (txn.splits && txn.splits.length > 0)
    ? txn.splits.map(s => ({ account_code: s.account_code, amount: String(s.amount), memo: s.memo || '' }))
    : (txn.account_code
        ? [{ account_code: txn.account_code, amount: target.toFixed(2), memo: txn.memo || '' }, { account_code: '', amount: '', memo: '' }]
        : [{ account_code: '', amount: '', memo: '' }, { account_code: '', amount: '', memo: '' }]);
  const [lines, setLines] = useState(initialLines);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);

  const parseAmt = v => { const n = parseFloat(String(v).replace(/[,$]/g,'')); return isNaN(n) ? 0 : n; };
  const total = lines.reduce((s, l) => s + parseAmt(l.amount), 0);
  const remaining = +(target - total).toFixed(2);
  const balanced = Math.abs(remaining) < 0.005;

  const updateLine = (i, field, val) => setLines(prev => prev.map((l, idx) => idx === i ? { ...l, [field]: val } : l));
  const addLine = () => setLines(prev => [...prev, { account_code: '', amount: remaining > 0 ? remaining.toFixed(2) : '', memo: '' }]);
  const removeLine = i => setLines(prev => prev.length > 1 ? prev.filter((_, idx) => idx !== i) : prev);
  const autoFillLast = () => { const lastIdx = lines.length - 1; if (remaining !== 0) updateLine(lastIdx, 'amount', (parseAmt(lines[lastIdx].amount) + remaining).toFixed(2)); };

  const save = async () => {
    setErr('');
    const valid = lines.filter(l => l.account_code && parseAmt(l.amount) > 0);
    if (valid.length === 0) { setErr('Add at least one account with an amount'); return; }
    if (!balanced) { setErr('Splits must total ' + fmt(target) + ' (currently off by ' + fmt(remaining) + ')'); return; }
    setSaving(true);
    try {
      await api.splitBankTransaction(entityId, txn.id, valid.map(l => ({ account_code: l.account_code, amount: parseAmt(l.amount), memo: l.memo || null })));
      onSaved();
    } catch (e) { setErr(e.message); } finally { setSaving(false); }
  };

  const clearSplits = async () => {
    if (!confirm('Remove all splits and revert to a single-account coding?')) return;
    setSaving(true); setErr('');
    try { await api.codeBankTransaction(entityId, txn.id, null, txn.memo || null); onSaved(); }
    catch (e) { setErr(e.message); } finally { setSaving(false); }
  };

  return (<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox, maxWidth: 820}} onClick={e => e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Split Transaction</div>
    <div style={{fontSize:12,color:T.textMuted,marginBottom:16}}>{txn.date} &middot; {txn.description}</div>
    <div style={{...S.card,background:T.bgElevated,padding:12,marginBottom:14,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <div><div style={{fontSize:11,color:T.textMuted,fontWeight:600,textTransform:'uppercase',letterSpacing:0.4}}>Transaction Amount</div>
        <div style={{fontSize:22,fontWeight:700,color:txn.amount>=0?T.green:T.red,marginTop:2}}>{txn.amount>=0?'+':'-'}${fmt(target)}</div></div>
      <div style={{textAlign:'right'}}><div style={{fontSize:11,color:T.textMuted,fontWeight:600,textTransform:'uppercase',letterSpacing:0.4}}>Remaining</div>
        <div style={{fontSize:22,fontWeight:700,color:balanced?T.green:T.orange,marginTop:2}}>${fmt(remaining)}</div></div>
    </div>
    <table style={{...S.table,marginBottom:10}}>
      <thead><tr><th style={S.th}>GL Account</th><th style={{...S.thR,width:140}}>Amount</th><th style={{...S.th,width:180}}>Memo</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{lines.map((l, i) => <tr key={i}>
        <td style={{...S.td,padding:'4px 6px'}}><AccountAutocomplete accounts={accounts} value={l.account_code} exclude={excludeCode} onChange={v => updateLine(i, 'account_code', v)} placeholder="Search GL account..."/></td>
        <td style={{...S.td,padding:'4px 6px'}}><input style={{...S.inputSm,textAlign:'right',fontFamily:'monospace'}} value={l.amount} onChange={e => updateLine(i, 'amount', e.target.value)} placeholder="0.00"/></td>
        <td style={{...S.td,padding:'4px 6px'}}><input style={S.inputSm} value={l.memo} onChange={e => updateLine(i, 'memo', e.target.value)} placeholder="Memo"/></td>
        <td style={{...S.td,padding:'4px 6px',textAlign:'center'}}>{lines.length > 1 && <button style={S.btnGhost} onClick={() => removeLine(i)}>x</button>}</td>
      </tr>)}</tbody>
    </table>
    <div style={{display:'flex',gap:8,marginBottom:14}}>
      <button style={{...S.btnS,fontSize:11,padding:'6px 12px'}} onClick={addLine}>+ Add line</button>
      {!balanced && lines.length > 0 && <button style={{...S.btnS,fontSize:11,padding:'6px 12px',color:T.accent,borderColor:T.accent+'40'}} onClick={autoFillLast}>Auto-fill remaining to last line</button>}
    </div>
    {err && <div style={{...S.err,marginBottom:12}}>{err}</div>}
    <div style={{display:'flex',gap:10,justifyContent:'space-between',alignItems:'center'}}>
      <div>{txn.splits && txn.splits.length > 0 && <button style={{...S.btnS,fontSize:11,color:T.red,borderColor:T.red+'40'}} onClick={clearSplits} disabled={saving}>Clear splits</button>}</div>
      <div style={{display:'flex',gap:10}}>
        <button style={S.btnS} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={{...S.btnP,opacity:(!balanced||saving)?0.5:1}} onClick={save} disabled={!balanced||saving}>{saving?'Saving...':'Save Splits'}</button>
      </div>
    </div>
  </div></div>);
}

function BankTransactions({entityId,canEdit=true,bankSelAcct:selAcct,setBankSelAcct:setSelAcct,bankTxns:txns,setBankTxns:setTxns,bankUploading:uploading,setBankUploading:setUploading,bankStatusFilter:statusFilter,setBankStatusFilter:setStatusFilter}){
  const[accounts,setAccounts]=useState([]);const[bankAccts,setBankAccts]=useState([]);
  const[err,setErr]=useState('');const[msg,setMsg]=useState('');const[showAddAcct,setShowAddAcct]=useState(false);
  const[uploadProgress,setUploadProgress]=useState('');const[discarding,setDiscarding]=useState(false);
  const[splitTxn,setSplitTxn]=useState(null);
  // Resizable column widths — persisted per-user in localStorage
  const BT_COLS_KEY='cl_bt_col_widths';
  const BT_DEFAULT_W={date:110,desc:260,amount:130,gl:280,memo:200,status:90};
  const[colW,setColW]=useState(()=>{try{return{...BT_DEFAULT_W,...(JSON.parse(localStorage.getItem(BT_COLS_KEY)||'{}'))};}catch{return BT_DEFAULT_W;}});
  const colWRef=useRef(colW);colWRef.current=colW;
  useEffect(()=>{try{localStorage.setItem(BT_COLS_KEY,JSON.stringify(colW));}catch{}},[colW]);
  const startResize=(key,ev)=>{ev.preventDefault();const startX=ev.clientX;const startW=colWRef.current[key];const onMove=e=>setColW(p=>({...p,[key]:Math.max(60,Math.min(800,startW+(e.clientX-startX)))}));const onUp=()=>{document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);};document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);};
  const resizeHandle=key=><span onMouseDown={ev=>startResize(key,ev)} style={{position:'absolute',right:0,top:6,bottom:6,width:6,cursor:'col-resize',userSelect:'none',borderRight:'2px solid '+T.border,transition:'border-color 0.15s'}} onMouseEnter={e=>{e.currentTarget.style.borderRightColor=T.accent;e.currentTarget.style.borderRightWidth='3px';}} onMouseLeave={e=>{e.currentTarget.style.borderRightColor=T.border;e.currentTarget.style.borderRightWidth='2px';}} title="Drag to resize column"/>;

  const loadAccounts=useCallback(async()=>{const a=await api.getAccounts(entityId);setAccounts(a);setBankAccts(a.filter(x=>x.bank_acct||(['cash','bank','checking','savings'].some(w=>x.name.toLowerCase().includes(w))&&x.type==='Asset')));return a;},[entityId]);
  const loadTxns=useCallback(async(acct,status)=>{if(!acct)return;const t=await api.getBankTransactions(entityId,acct,status||undefined);setTxns(t);},[entityId,setTxns]);
  useEffect(()=>{loadAccounts();},[loadAccounts]);
  useEffect(()=>{if(selAcct)loadTxns(selAcct,statusFilter);else setTxns([]);},[selAcct,entityId]);
  const reload=()=>loadTxns(selAcct,statusFilter);

  const onFileSelected=async e=>{const file=e.target.files[0];if(!file||!selAcct)return;e.target.value='';setErr('');setMsg('');setUploading(true);setUploadProgress('Uploading file...');
    try{const r=await api.uploadBankTransactions(entityId,selAcct,file);
      setUploadProgress('Auto-categorizing '+r.count+' transactions...');
      const imported=await api.getBankTransactions(entityId,selAcct,'pending');let auto=0;
      for(const t of imported){if(!t.account_code){const sg=suggestAccount(t.description,accounts,selAcct);if(sg){await api.codeBankTransaction(entityId,t.id,sg.code,t.memo||t.description);auto++;}}}
      setMsg(r.count+' imported'+(auto>0?', '+auto+' auto-categorized':''));loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}finally{setUploading(false);setUploadProgress('');}};
  const cancelUpload=()=>{setUploading(false);setUploadProgress('');setMsg('Upload cancelled');};
  const discardAllUnposted=async()=>{const unposted=txns.filter(t=>t.status!=='posted');if(!unposted.length){setErr('Nothing to discard');return;}
    const batchIds=[...new Set(unposted.map(t=>t.batch_id).filter(Boolean))];
    if(!confirm('Discard all '+unposted.length+' unposted transaction(s) from this account? This cannot be undone. Posted transactions will be kept.'))return;
    setErr('');setMsg('');setDiscarding(true);
    try{let total=0;for(const bid of batchIds){const r=await api.deleteBankBatch(entityId,bid);total+=(r.deleted||0);}
      setMsg(total+' transaction(s) discarded');loadTxns(selAcct,statusFilter);}
    catch(ex){setErr(ex.message);}finally{setDiscarding(false);}};
  const codeTransaction=async(id,acct_code,memo)=>{await api.codeBankTransaction(entityId,id,acct_code,memo);
    setTxns(prev=>prev.map(t=>t.id===id?{...t,account_code:acct_code,memo:memo,status:acct_code?'coded':'pending'}:t));};
  const postCoded=async()=>{const ids=txns.filter(t=>t.status==='coded').map(t=>t.id);if(!ids.length){setErr('Nothing coded');return;}try{const r=await api.postBankTransactions(entityId,ids);setMsg(r.posted+' JEs created');loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}};
  const changeAcct=v=>{setSelAcct(v);setTxns([]);if(v)loadTxns(v,statusFilter);};
  const changeStatus=v=>{setStatusFilter(v);if(selAcct)loadTxns(selAcct,v);};

  const filteredTxns=txns;
  const totalIn=filteredTxns.filter(t=>t.amount>0).reduce((s,t)=>s+t.amount,0);const totalOut=filteredTxns.filter(t=>t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const uncat=filteredTxns.filter(t=>t.status==='pending').length;

  return(<div><div style={S.h1}>Bank Transactions</div><div style={S.sub}>Upload, categorize, and post bank activity to the general ledger</div>
    <div style={S.card}><div style={S.row}>
      <div style={{...S.col,flex:2}}><label style={S.label}>Bank Account</label><select style={S.select} value={selAcct} onChange={e=>changeAcct(e.target.value)}><option value="">Select bank account...</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
      <div style={S.col}><label style={S.label}>Status</label><select style={S.select} value={statusFilter} onChange={e=>changeStatus(e.target.value)}><option value="">All</option><option value="pending">Pending</option><option value="coded">Coded</option><option value="posted">Posted</option></select></div>
      {canEdit&&selAcct&&<div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
        {uploading
          ?<div style={{display:'flex',alignItems:'center',gap:10,marginBottom:4}}>
            <div style={{width:16,height:16,border:'2px solid '+T.accent,borderTopColor:'transparent',borderRadius:'50%',animation:'spin 1s linear infinite'}}/>
            <span style={{fontSize:12,color:T.accent,fontWeight:500}}>{uploadProgress||'Processing...'}</span>
            <button style={{...S.btnD,padding:'7px 16px',fontSize:12}} onClick={cancelUpload}>Cancel</button></div>
          :<div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
            <button style={{...S.btnP,pointerEvents:'none'}}>Upload CSV / Excel / PDF</button>
            <input type="file" accept=".csv,.xlsx,.xls,.pdf" style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:'pointer'}} onChange={onFileSelected}/></div>}
      </div>}
    </div>
    {err&&<div style={S.err}>{err}</div>}{msg&&<div style={S.success}>{msg}</div>}
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
          {canEdit&&<button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New Account</button>}
          {canEdit&&filteredTxns.some(t=>t.status!=='posted')&&<button style={{...S.btnD,padding:'8px 14px',fontSize:12}} disabled={discarding} onClick={discardAllUnposted}>{discarding?'Discarding...':'Discard '+filteredTxns.filter(t=>t.status!=='posted').length+' unposted'}</button>}
          {canEdit&&filteredTxns.some(t=>t.status==='coded')&&<button style={S.btnP} onClick={postCoded}>Post {filteredTxns.filter(t=>t.status==='coded').length} to GL</button>}</div></div>
      <table style={{...S.table,tableLayout:'fixed',width:'100%'}}>
        <colgroup><col style={{width:colW.date}}/><col style={{width:colW.desc}}/><col style={{width:colW.amount}}/><col style={{width:colW.gl}}/><col style={{width:colW.memo}}/><col style={{width:colW.status}}/><col style={{width:36}}/></colgroup>
        <thead><tr>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Date{resizeHandle('date')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Description{resizeHandle('desc')}</th>
          <th style={{...S.thR,position:'relative',borderRight:'1px solid '+T.borderLight}}>Amount{resizeHandle('amount')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>GL Account{resizeHandle('gl')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Memo{resizeHandle('memo')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Status{resizeHandle('status')}</th>
          <th style={{...S.th,width:36}}></th></tr></thead>
        <tbody>{filteredTxns.map(t=><tr key={t.id} style={t.status==='posted'?{opacity:0.45}:{}}>
          <td style={{...S.td,color:T.textMuted,fontSize:12,borderRight:'1px solid '+T.borderLight}} title={t.date}>{t.date}</td>
          <td style={{...S.td,fontWeight:500,borderRight:'1px solid '+T.borderLight}} title={t.description}>{t.description}</td>
          <td style={{...S.tdR,fontSize:15,fontWeight:700,color:t.amount>=0?T.green:T.red,borderRight:'1px solid '+T.borderLight}}>{t.amount>=0?'+':''}{fmt(t.amount)}</td>
          <td style={{...S.td,padding:'4px 6px',overflow:'visible',borderRight:'1px solid '+T.borderLight}}>{(t.status==='posted'||!canEdit)
            ? (t.splits && t.splits.length>0
                ? <span style={{fontSize:11,color:T.textDim}}>Split: {t.splits.length} accts</span>
                : <span style={{fontSize:12,color:T.textDim}}>{t.account_code}</span>)
            : (t.splits && t.splits.length>0
                ? <button style={{...S.btnS,padding:'5px 10px',fontSize:11,color:T.purple,borderColor:T.purple+'40',width:'100%',textAlign:'left'}} onClick={()=>setSplitTxn(t)} title={t.splits.map(s=>s.account_code+' $'+fmt(s.amount)).join(' | ')}>Split: {t.splits.length} accts &middot; ${fmt(t.splits.reduce((s,x)=>s+x.amount,0))}</button>
                : <div style={{display:'flex',gap:4,alignItems:'center'}}>
                    <div style={{flex:1,minWidth:0}}><AccountAutocomplete accounts={accounts} value={t.account_code||''} exclude={selAcct} onChange={v=>codeTransaction(t.id,v,t.memo)} placeholder="Search GL account..."/></div>
                    <button style={{...S.btnGhost,fontSize:10,color:T.purple,padding:'4px 6px',whiteSpace:'nowrap'}} onClick={()=>setSplitTxn(t)} title="Split across multiple accounts">Split</button>
                  </div>)}</td>
          <td style={{...S.td,padding:'4px 6px',overflow:'visible',borderRight:'1px solid '+T.borderLight}}>{(t.status==='posted'||!canEdit)?<span style={{fontSize:12,color:T.textDim}}>{t.memo}</span>:
            (t.splits && t.splits.length>0
              ? <span style={{fontSize:11,color:T.textDim,fontStyle:'italic'}}>(per split)</span>
              : <input style={S.inputSm} placeholder="Memo" value={t.memo||''} onChange={e=>{const v=e.target.value;setTxns(prev=>prev.map(x=>x.id===t.id?{...x,memo:v}:x));}} onBlur={()=>codeTransaction(t.id,t.account_code,t.memo)}/>)}</td>
          <td style={{...S.td,borderRight:'1px solid '+T.borderLight}}><span style={{fontSize:10,fontWeight:600,padding:'3px 8px',borderRadius:20,background:t.status==='posted'?T.greenDim:t.status==='coded'?T.accentDim:T.orangeDim,color:t.status==='posted'?T.green:t.status==='coded'?T.accent:T.orange}}>{t.status}</span></td>
          <td style={S.td}>{canEdit&&t.status!=='posted'&&<button style={S.btnGhost} onClick={async()=>{await api.deleteBankTransaction(entityId,t.id);setTxns(prev=>prev.filter(x=>x.id!==t.id));}}>x</button>}</td>
        </tr>)}</tbody></table></div>}
    {selAcct&&filteredTxns.length===0&&!uploading&&<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No transactions yet. Upload a bank statement above.</div>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>{setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));if(a.bank_acct)setBankAccts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));}}/>}
    {splitTxn&&<SplitBankTransactionModal txn={splitTxn} accounts={accounts} excludeCode={selAcct} entityId={entityId} onClose={()=>setSplitTxn(null)} onSaved={()=>{setSplitTxn(null);loadTxns(selAcct,statusFilter);}}/>}
  </div>);}

// ═══ Reports ═══
function WipSchedule({entityName,asOf,setAsOf}){
  const[data,setData]=useState(null);
  const[err,setErr]=useState('');
  const[loading,setLoading]=useState(true);
  const validAsOf=/^\d{4}-\d{2}-\d{2}$/.test(asOf)?asOf:today();
  useEffect(()=>{let alive=true;setLoading(true);setErr('');
    api.getTurnkeyWip(validAsOf).then(d=>{if(alive){setData(d);setLoading(false);}}).catch(e=>{if(alive){setErr(e.message);setLoading(false);}});
    return()=>{alive=false;};},[validAsOf]);
  const rows=(data&&data.rows)||[];
  const tot=(data&&data.total)||null;
  const doExport=()=>{
    const d=[[entityName||'WIP Schedule'],['Work-in-Progress Schedule'],['As of '+validAsOf],[],
      ['Job #','Job Name','Contract','Revised Contract','Costs to Date','Est Cost to Complete','Est Total Cost','Est Gross Profit','% Complete','Earned Revenue','Billed to Date','Over/(Under) Billing']];
    rows.forEach(r=>d.push([r.project_code||r.turnkey_project_id,r.project_name||'',r.contract_amount,r.revised_contract,r.costs_to_date,r.estimated_cost_to_complete,r.estimated_total_cost,r.estimated_gross_profit,(r.percent_complete||0)/100,r.earned_revenue,r.billed_to_date,r.over_under_billing]));
    if(tot){d.push([]);d.push(['','Total',tot.contract_amount,tot.revised_contract,tot.costs_to_date,tot.estimated_cost_to_complete,tot.estimated_total_cost,tot.estimated_gross_profit,'',tot.earned_revenue,tot.billed_to_date,tot.over_under_billing]);}
    exportToExcel(d,'WIP_'+validAsOf+'.xlsx');
  };
  const numCell=(n)=><td style={S.tdR}>{fmt(n)}</td>;
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport} disabled={!rows.length}>Export Excel</button></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>WIP Schedule</div><div style={{fontSize:13,color:T.textMuted}}>As of {validAsOf}</div></div>
    {err&&<div style={S.err}>{err}</div>}
    {loading?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Loading…</div>:
     rows.length===0?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>No projects linked yet. Projects appear here once they are linked from Turnkey Rail.</div>:
    <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:1100}}>
      <thead><tr><th style={S.th}>Job #</th><th style={S.th}>Job Name</th><th style={S.thR}>Contract</th><th style={S.thR}>Revised</th><th style={S.thR}>Costs to Date</th><th style={S.thR}>Est Cost to Compl.</th><th style={S.thR}>Est Total Cost</th><th style={S.thR}>Est Gross Profit</th><th style={S.thR}>% Compl.</th><th style={S.thR}>Earned Rev.</th><th style={S.thR}>Billed</th><th style={S.thR}>Over/(Under)</th></tr></thead>
      <tbody>{rows.map(r=><tr key={r.turnkey_project_id}>
        <td style={{...S.td,color:T.textBright}}>{r.project_code||r.turnkey_project_id}</td>
        <td style={S.td} title={r.project_name}>{r.project_name||''}</td>
        {numCell(r.contract_amount)}{numCell(r.revised_contract)}{numCell(r.costs_to_date)}{numCell(r.estimated_cost_to_complete)}{numCell(r.estimated_total_cost)}{numCell(r.estimated_gross_profit)}
        <td style={S.tdR}>{(r.percent_complete||0).toFixed(1)}%</td>
        {numCell(r.earned_revenue)}{numCell(r.billed_to_date)}
        <td style={{...S.tdR,color:r.over_under_label==='under'?T.orange:T.textBright}}>{fmt(r.over_under_billing)} {r.over_under_label}</td></tr>)}
        {tot&&<tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={2}>Total</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.contract_amount)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.revised_contract)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.costs_to_date)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.estimated_cost_to_complete)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.estimated_total_cost)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.estimated_gross_profit)}</td>
          <td style={S.tdBold}></td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.earned_revenue)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.billed_to_date)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>${fmt(tot.over_under_billing)}</td></tr>}
      </tbody></table></div>}
  </div></div>);
}

function TrialBalance({entityId,entityName,dimsEnabled,asOf,setAsOf}){
  const[balances,setBalances]=useState([]);
  const[drillAcct,setDrillAcct]=useState(null);
  const[locations,setLocations]=useState([]);
  const[locId,setLocId]=useState('');// '' = all (whole-entity TB); otherwise a location_id
  const[classes,setClasses]=useState([]);
  const[classId,setClassId]=useState('');// '' = all investors; otherwise a class_id (investor)
  // Guard: while the user is editing the date input, asOf can briefly be '' or a partial string like '2026-'.
  // Avoid crashing the page on Invalid Date — fall back to today() until a complete YYYY-MM-DD is entered.
  const validAsOf=/^\d{4}-\d{2}-\d{2}$/.test(asOf)&&!isNaN(new Date(asOf+'T00:00:00').getTime())?asOf:today();
  const fyS=validAsOf.slice(0,4)+'-01-01';
  const locName=locId?(locations.find(l=>String(l.id)===String(locId))?.name||''):'';
  const className=classId?(classes.find(c=>String(c.id)===String(classId))?.name||''):'';
  const dimmed=!!(locId||classId); // any dimension selected → activity view
  const scopeLabel=[locName,className].filter(Boolean).join(' · ');
  // 12-month window ending at asOf. If asOf is 2026-04-10, window is 2025-04-11 → 2026-04-10.
  const drillFrom=useMemo(()=>{const d=new Date(validAsOf+'T00:00:00');d.setFullYear(d.getFullYear()-1);d.setDate(d.getDate()+1);return d.toISOString().slice(0,10);},[validAsOf]);
  // Whole-entity TB uses the soft-close (close_pl_before) path. A dimension-scoped TB
  // (location and/or class/investor) is activity-based: it sums only lines carrying
  // the selected tag(s), so there is no period-close/RE roll — pass the dimension
  // id(s) with the as_of date only.
  useEffect(()=>{
    if(dimmed) api.getBalances(entityId,{as_of:validAsOf,...(locId?{location_id:locId}:{}),...(classId?{class_id:classId}:{})}).then(setBalances);
    else api.getBalances(entityId,{as_of:validAsOf,close_pl_before:fyS}).then(setBalances);
  },[entityId,validAsOf,fyS,locId,classId,dimmed]);
  useEffect(()=>{api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));},[entityId]);
  useEffect(()=>{api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  let tDr=0,tCr=0;const rows=balances.filter(b=>Math.abs(b.balance)>0.005).map(b=>{const isDr=b.type==='Asset'||b.type==='Expense';const dr=(isDr&&b.balance>0)||(!isDr&&b.balance<0)?Math.abs(b.balance):0;const cr=(isDr&&b.balance<0)||(!isDr&&b.balance>0)?Math.abs(b.balance):0;tDr+=dr;tCr+=cr;return{...b,dr,cr};});
  const fnameTag=[locName,className].filter(Boolean).map(s=>s.replace(/[^A-Za-z0-9]+/g,'_')).join('_');
  const doExport=()=>{const lbl=scopeLabel?(' — '+scopeLabel):'';const d=[[entityName||'Trial Balance'],['Trial Balance'+lbl],['As of '+asOf],[],['Code','Account','Type','Debit','Credit']];rows.forEach(r=>d.push([r.code,r.name,r.type,r.dr||'',r.cr||'']));d.push([]);d.push(['','','Total',tDr,tCr]);exportToExcel(d,'TB'+(fnameTag?'_'+fnameTag:'')+'_'+asOf+'.xlsx');};
  const amtStyle={...S.tdR,cursor:'pointer'};
  // GL detail export (optionally scoped to the selected location and/or investor).
  // Pulls flat lines with running balance from /gl-detail through the as-of date;
  // dimension-tagged lines only when a dimension is selected.
  const doExportGL=async()=>{
    try{
      const r=await api.getGLDetail(entityId,{to:validAsOf,...(locId?{location_id:locId}:{}),...(classId?{class_id:classId}:{})});
      const lbl=scopeLabel?(' — '+scopeLabel):'';
      const d=[[entityName||'General Ledger'],['GL Detail'+lbl],['Through '+asOf],[],['Date','Entry #','Account','Account Name','Memo / Description','Location','Class','Debit','Credit','Running Bal']];
      (r.lines||[]).forEach(l=>d.push([l.date,l.entry_num,l.account_code,l.account_name,l.description||l.memo||'',l.location_name,l.class_name,l.debit||'',l.credit||'',l.running_balance]));
      d.push([]);d.push(['','','','','','','','Total Dr','Total Cr','']);
      d.push(['','','','','','','',r.total_debit,r.total_credit,'']);
      exportToExcel(d,'GL'+(fnameTag?'_'+fnameTag:'')+'_'+asOf+'.xlsx');
    }catch(e){alert('GL export failed: '+e.message);}
  };
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
      {dimsEnabled&&<div><label style={S.label}>Location</label><select style={S.inputSm} value={locId} onChange={e=>setLocId(e.target.value)}><option value="">All (whole entity)</option>{locations.map(l=><option key={l.id} value={l.id}>{l.name}{l.line_count!=null?(' ('+l.line_count+')'):''}</option>)}</select></div>}
      {dimsEnabled&&<div><label style={S.label}>Investor (Class)</label><select style={S.inputSm} value={classId} onChange={e=>setClassId(e.target.value)}><option value="">All investors</option>{classes.map(c=><option key={c.id} value={c.id}>{c.name}{c.line_count!=null?(' ('+c.line_count+')'):''}</option>)}</select></div>}</div>
    <div style={{display:'flex',gap:8}}><button style={S.btnExport} onClick={doExportGL} title="Export flat GL detail (dimension-tagged only when a location/investor is selected)">Export GL Detail</button><button style={S.btnExport} onClick={doExport}>Export TB</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Trial Balance{scopeLabel?(' — '+scopeLabel):''}</div><div style={{fontSize:13,color:T.textMuted}}>As of {asOf}{dimmed?' · dimension-tagged activity only':''}</div></div>
    <table style={{...S.table,tableLayout:'fixed',width:'100%'}}>
      <colgroup><col style={{width:'90px'}}/><col/><col style={{width:'120px'}}/><col style={{width:'160px'}}/><col style={{width:'160px'}}/></colgroup>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th></tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={{...S.td,color:T.textBright}}>{r.code}</td><td style={{...S.td,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={r.name}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td>
        <td style={amtStyle} onClick={()=>r.dr>0&&setDrillAcct(r)} title={r.dr>0?'Click for 12-month detail':''}>{r.dr>0?<span style={{color:T.accent,borderBottom:'1px dotted '+T.accent+'80'}}>{fmt(r.dr)}</span>:''}</td>
        <td style={amtStyle} onClick={()=>r.cr>0&&setDrillAcct(r)} title={r.cr>0?'Click for 12-month detail':''}>{r.cr>0?<span style={{color:T.accent,borderBottom:'1px dotted '+T.accent+'80'}}>{fmt(r.cr)}</span>:''}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(tDr-tCr)<0.005?T.green:T.red}}>{Math.abs(tDr-tCr)<0.005?'In balance':'Off by $'+fmt(tDr-tCr)}</div></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillFrom} to={asOf} onClose={()=>setDrillAcct(null)}/>}
  </div>);
}

// ═══ Account Drill-Down Modal (12-month GL detail from TB) ═══
function AccountDrillDownModal({entityId,entityName,acct,from,to,onClose}){
  const[lines,setLines]=useState([]);
  const[begBal,setBegBal]=useState(0);
  const[loading,setLoading]=useState(true);
  const[err,setErr]=useState('');
  const[allEntries,setAllEntries]=useState([]);
  const[allAccounts,setAllAccounts]=useState([]);
  const[viewEntry,setViewEntry]=useState(null);
  useEffect(()=>{
    (async()=>{
      setLoading(true);setErr('');
      try{
        const[entries,accts]=await Promise.all([api.getEntries(entityId,from,to),api.getAccounts(entityId)]);
        setAllEntries(entries);setAllAccounts(accts);
        const txns=[];
        entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({date:e.date,entry_num:e.entry_num,jeId:e.id,memo:e.memo,debit:l.debit||0,credit:l.credit||0,created_by:e.created_by,created_at:e.created_at});});});
        txns.sort((a,b)=>a.date.localeCompare(b.date)||a.entry_num-b.entry_num);
        const windowDr=txns.reduce((s,t)=>s+t.debit,0);
        const windowCr=txns.reduce((s,t)=>s+t.credit,0);
        const isDr=acct.type==='Asset'||acct.type==='Expense';
        const netWindow=isDr?(windowDr-windowCr):(windowCr-windowDr);
        setBegBal(acct.balance-netWindow);
        setLines(txns);
      }catch(e){setErr(e.message);}finally{setLoading(false);}
    })();
  },[entityId,acct.code,acct.balance,acct.type,from,to]);
  const openJE=jeId=>{const e=allEntries.find(x=>x.id===jeId);if(e)setViewEntry(e);};
  const isDr=acct.type==='Asset'||acct.type==='Expense';
  let running=begBal;
  const totalDr=lines.reduce((s,l)=>s+l.debit,0);
  const totalCr=lines.reduce((s,l)=>s+l.credit,0);
  const doExport=()=>{const d=[[entityName||'Account Detail'],[acct.code+' - '+acct.name],['Period: '+from+' to '+to],[],['Date','JE','Memo','Debit','Credit','Balance']];
    d.push(['','','Beginning Balance','','',begBal]);
    let r=begBal;lines.forEach(l=>{r+=isDr?(l.debit-l.credit):(l.credit-l.debit);d.push([l.date,'JE-'+String(l.entry_num).padStart(4,'0'),l.memo,l.debit||'',l.credit||'',r]);});
    d.push(['','','Totals',totalDr,totalCr,r]);
    exportToExcel(d,'GL_'+acct.code+'_'+to+'.xlsx');};
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:1100,maxHeight:'90vh',display:'flex',flexDirection:'column'}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14,gap:16}}>
      <div><div style={{fontSize:18,fontWeight:700,color:T.textBright}}>{acct.code} &mdash; {acct.name}</div>
        <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>Trailing 12 months &middot; {from} to {to}</div>
        <div style={{marginTop:6}}><span style={S.tag(acct.type)}>{acct.type}</span></div></div>
      <button style={S.btnExport} onClick={doExport}>Export Excel</button>
    </div>
    {loading?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Loading...</div>:
     err?<div style={S.err}>{err}</div>:
     <div style={{flex:1,overflowY:'auto',border:'1px solid '+T.border,borderRadius:T.radiusSm}}>
       <table style={S.table}>
         <thead style={{position:'sticky',top:0,background:T.bgCard,zIndex:1}}><tr>
           <th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th>
           <th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th></tr></thead>
         <tbody>
           <tr style={{background:T.bgElevated}}>
             <td style={{...S.td,color:T.textMuted,fontStyle:'italic'}} colSpan={3}>Beginning balance as of {from}</td>
             <td style={S.tdR}></td><td style={S.tdR}></td>
             <td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(begBal)}</td></tr>
           {lines.length===0?
             <tr><td colSpan={6} style={{...S.td,textAlign:'center',padding:30,color:T.textDim}}>No activity in this period</td></tr>
             :lines.map((l,i)=>{running+=isDr?(l.debit-l.credit):(l.credit-l.debit);
               const tip=(l.created_by?'Posted by '+l.created_by:'')+(l.created_at?(l.created_by?' on ':'Posted on ')+new Date(l.created_at+(l.created_at.includes('Z')||l.created_at.includes('+')?'':'Z')).toLocaleString('en-US',{timeZone:'America/Los_Angeles',year:'numeric',month:'short',day:'numeric',hour:'numeric',minute:'2-digit',hour12:true,timeZoneName:'short'}):'');
               return<tr key={i}>
                 <td style={{...S.td,color:T.textMuted,whiteSpace:'nowrap'}}>{l.date}</td>
                 <td style={S.td} title={tip}><button style={{background:'none',border:0,padding:0,color:T.accent,fontWeight:600,cursor:'pointer',fontSize:'inherit',fontFamily:'inherit'}} onClick={()=>openJE(l.jeId)}>JE-{String(l.entry_num).padStart(4,'0')}</button></td>
                 <td style={S.td}>{l.memo}</td>
                 <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td>
                 <td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td>
                 <td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(running)}</td></tr>;})}
           <tr style={S.grandTotalRow}>
             <td style={{...S.tdBold}} colSpan={3}>Period Totals</td>
             <td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalDr)}</td>
             <td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalCr)}</td>
             <td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>${fmt(running)}</td></tr>
         </tbody></table></div>}
    {viewEntry&&<EditJEModal entityId={entityId} entry={viewEntry} accounts={allAccounts} onClose={()=>setViewEntry(null)} onSaved={()=>setViewEntry(null)}/>}
  </div></div>);
}

function BalanceSheet({entityId,entityName,asOf,setAsOf}){const[balances,setBalances]=useState([]);const[drillAcct,setDrillAcct]=useState(null);
  // Guard: while the user is editing the date input, asOf can briefly be '' or a partial string like '2026-'.
  const validAsOf=/^\d{4}-\d{2}-\d{2}$/.test(asOf)&&!isNaN(new Date(asOf+'T00:00:00').getTime())?asOf:today();
  const fyS=validAsOf.slice(0,4)+'-01-01';
  const drillFrom=useMemo(()=>{const d=new Date(validAsOf+'T00:00:00');d.setFullYear(d.getFullYear()-1);d.setDate(d.getDate()+1);return d.toISOString().slice(0,10);},[validAsOf]);
  useEffect(()=>{api.getBalances(entityId,{as_of:validAsOf,close_pl_before:fyS}).then(setBalances);},[entityId,validAsOf,fyS]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=t=>get(t).reduce((s,b)=>s+b.balance,0);
  const ni=sum('Revenue')-sum('Expense');const tA=sum('Asset');const tLE=sum('Liability')+sum('Equity')+ni;
  const doExport=()=>{const d=[[entityName||'Balance Sheet'],['Balance Sheet'],['As of '+asOf],[]];[['Assets','Asset'],['Liabilities','Liability'],['Equity','Equity']].forEach(([t,ty])=>{d.push([t,'']);get(ty).forEach(b=>d.push(['  '+b.name,b.balance]));if(ty==='Equity'&&Math.abs(ni)>0.005)d.push(['  Net Income (current period)',ni]);d.push(['Total '+t,ty==='Equity'?sum(ty)+ni:sum(ty)]);d.push([]);});d.push(['Total L+E',tLE]);exportToExcel(d,'BS_'+asOf+'.xlsx');};
  const Sec=({title,type,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{get(type).map(b=><tr key={b.code}><td style={{...S.indentTd,cursor:'pointer'}} onClick={()=>setDrillAcct(b)}><span style={{borderBottom:'1px dotted '+T.accent+'80',color:T.accent}}>{b.name}</span></td><td style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct(b)}>{fmt(b.balance)}</td></tr>)}
    {type==='Equity'&&Math.abs(ni)>0.005&&<tr><td style={{...S.indentTd,fontStyle:'italic',color:T.textMuted}}>Net Income (current period)</td><td style={{...S.tdR,fontStyle:'italic'}}>{fmt(ni)}</td></tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>${fmt(total)}</td></tr></>);
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></div>
    <button style={S.btnExport} onClick={doExport}>Export Excel</button></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Balance Sheet</div><div style={{fontSize:13,color:T.textMuted}}>As of {asOf}</div></div>
    <table style={{...S.table,maxWidth:580,margin:'0 auto'}}><tbody><Sec title="Assets" type="Asset" total={tA}/><tr><td colSpan={2} style={{padding:8}}/></tr>
      <Sec title="Liabilities" type="Liability" total={sum('Liability')}/><tr><td colSpan={2} style={{padding:4}}/></tr><Sec title="Equity" type="Equity" total={sum('Equity')+ni}/>
      <tr style={S.grandTotalRow}><td style={S.tdBold}>Total Liabilities + Equity</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tLE)}</td></tr></tbody></table>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(tA-tLE)<0.005?T.green:T.red}}>{Math.abs(tA-tLE)<0.005?'A = L + E':'Off by $'+fmt(tA-tLE)}</div></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillFrom} to={asOf} onClose={()=>setDrillAcct(null)}/>}
    </div>);}

function IncomeStatement({entityId,entityName,from,setFrom,to,setTo}){const[balances,setBalances]=useState([]);const[drillAcct,setDrillAcct]=useState(null);
  useEffect(()=>{api.getBalances(entityId,{from,to}).then(setBalances);},[entityId,from,to]);
  const get=t=>balances.filter(b=>b.type===t&&Math.abs(b.balance)>0.005);const sum=arr=>arr.reduce((s,b)=>s+b.balance,0);
  const rev=get('Revenue');const cogs=get('Expense').filter(b=>b.subtype==='COGS');const opex=get('Expense').filter(b=>b.subtype==='Operating Expense');const other=get('Expense').filter(b=>b.subtype!=='COGS'&&b.subtype!=='Operating Expense');
  const tRev=sum(rev);const gp=tRev-sum(cogs);const oi=gp-sum(opex);const ni=oi-sum(other);
  const doExport=()=>{const d=[[entityName||'Income Statement'],['Income Statement'],['Period: '+from+' to '+to],[]];[['Revenue',rev],['COGS',cogs],['Operating Expenses',opex],['Other',other]].forEach(([t,items])=>{if(!items.length)return;d.push([t,'']);items.forEach(b=>d.push(['  '+b.name,b.balance]));d.push(['Total '+t,sum(items)]);d.push([]);});d.push(['Net Income',ni]);exportToExcel(d,'IS_'+from+'_'+to+'.xlsx');};
  const Sec=({title,items,total})=>(<><tr><td style={S.sectionHeader} colSpan={2}>{title}</td></tr>{items.map(b=><tr key={b.code}><td style={{...S.indentTd,cursor:'pointer'}} onClick={()=>setDrillAcct(b)}><span style={{borderBottom:'1px dotted '+T.accent+'80',color:T.accent}}>{b.name}</span></td><td style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct(b)}>{fmt(b.balance)}</td></tr>)}
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
      <tr style={S.grandTotalRow}><td style={{...S.tdBold,fontSize:15}}>Net Income</td><td style={{...S.tdBold,textAlign:'right',fontSize:18,color:ni>=0?T.green:T.red}}>${fmt(ni)}</td></tr></tbody></table></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={from} to={to} onClose={()=>setDrillAcct(null)}/>}
    </div>);}

// ═══ Bank Reconciliation ═══
function BankReconciliation({entityId,user,canEdit=true}){const[accounts,setAccounts]=useState([]);const[entries,setEntries]=useState([]);const[recs,setRecs]=useState([]);
  const[view,setView]=useState('list');const[selAcct,setSelAcct]=useState('');const[stmtDate,setStmtDate]=useState(today());const[stmtBal,setStmtBal]=useState('');
  const[cleared,setCleared]=useState({});const[checked,setChecked]=useState({});
  const[viewEntry,setViewEntry]=useState(null);
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
            <td style={S.tdC}><input type="checkbox" style={S.checkbox} checked={!!checked[t.key]} readOnly/></td><td style={{...S.td,color:T.textMuted}}>{t.date}</td><td style={S.td} onClick={e=>{e.stopPropagation();const ent=entries.find(x=>x.id===t.jeId);if(ent)setViewEntry(ent);}}><button style={{background:'none',border:0,padding:0,color:T.accent,fontWeight:600,cursor:'pointer',fontSize:'inherit',fontFamily:'inherit'}}>JE-{String(t.jeNum).padStart(4,'0')}</button></td><td style={S.td}>{t.memo}</td>
            <td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:700,color:t.amount>=0?T.green:T.red}}>{fmt(t.amount)}</td></tr>)}</tbody></table>}</div>
      <div style={{display:'flex',justifyContent:'flex-end',marginTop:16}}><button style={{...S.btnP,padding:'10px 28px',fontSize:14,opacity:isRec?1:.5,cursor:isRec?'pointer':'not-allowed'}} onClick={finalize}>{isRec?'Finalize Reconciliation':'Difference must be $0.00'}</button></div></>}
    {viewEntry&&<EditJEModal entityId={entityId} entry={viewEntry} accounts={accounts} onClose={()=>setViewEntry(null)} onSaved={()=>{setViewEntry(null);load();}}/>}
    </div>);
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
    <div><div style={S.h1}>Bank Reconciliation</div><div style={S.sub}>{recs.length} completed</div></div>{canEdit&&<button style={S.btnP} onClick={()=>setView('new')}>+ New Reconciliation</button>}</div>
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(240px,1fr))',gap:16,marginBottom:20}}>
      {bankAccts.map(a=>{const t=getTxns(a.code);const bal=t.reduce((s,x)=>s+x.amount,0);return<div key={a.code} style={{...S.card,padding:20}}>
        <div style={{fontWeight:700,color:T.textBright,fontSize:14,marginBottom:4}}>{a.name}</div><div style={{fontSize:12,color:T.textDim,marginBottom:12}}>{a.code}</div>
        <div style={{fontSize:24,fontWeight:700,color:T.textBright}}>${fmt(bal)}</div></div>;})}</div>
    <div style={S.cardFlush}><div style={{padding:'16px 20px'}}><div style={S.h2}>History</div></div>{recs.length===0?<div style={{padding:40,textAlign:'center',color:T.textDim}}>No reconciliations yet</div>:
      <table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.thR}>Statement</th><th style={S.thR}>Book</th><th style={S.thR}>Cleared</th><th style={S.th}>By</th></tr></thead>
        <tbody>{recs.map(r=><tr key={r.id}><td style={S.td}>{r.statement_date}</td><td style={S.td}>{r.account_code}</td><td style={S.tdR}>${fmt(r.statement_balance)}</td><td style={S.tdR}>${fmt(r.book_balance)}</td><td style={S.tdR}>{r.cleared_count}</td><td style={S.td}>{r.completed_by}</td></tr>)}</tbody></table>}</div></div>);}

// ═══ Entity Management ═══
// ═══ Requisitions (development-project coding engine) ═══
function Requisitions({entityId,entityName,canEdit=true,reqState,setReqState}){
  const[err,setErr]=useState('');
  // Persistent working set lifted to App (survives navigation), kept per-entity.
  const rs=reqState||{cards:[],reqNum:'',asOf:today(),result:null,detail:null,file:null};
  const rfCards=rs.cards, rfReqNum=rs.reqNum, rfAsOf=rs.asOf, rfResult=rs.result, rfDetail=rs.detail, rfFile=rs.file||null;
  const setRfCards=updater=>setReqState(cur=>({...cur,cards:typeof updater==='function'?updater(cur.cards||[]):updater}));
  const setRfReqNum=v=>setReqState(cur=>({...cur,reqNum:v}));
  const setRfAsOf=v=>setReqState(cur=>({...cur,asOf:v}));
  const setRfResult=v=>setReqState(cur=>({...cur,result:v}));
  const setRfDetail=v=>setReqState(cur=>({...cur,detail:v}));
  const setRfFile=v=>setReqState(cur=>({...cur,file:v}));
  // Transient (mid-operation) state stays local.
  const[rfBusy,setRfBusy]=useState(false);const[rfErr,setRfErr]=useState('');
  const[rfReading,setRfReading]=useState(0);const[rfReadErr,setRfReadErr]=useState('');
  // Cost-code -> name catalog (from requisition_coa_map / invoice history) used to
  // auto-fill the Cost Code Name when a code is entered. Loaded per entity.
  const[coaMap,setCoaMap]=useState({});
  useEffect(()=>{let alive=true;(async()=>{try{const r=await api.getRequisitionCoaMap(entityId);if(alive)setCoaMap((r&&r.map)||{});}catch{if(alive)setCoaMap({});}})();return()=>{alive=false;};},[entityId]);
  // Cost-code -> name parsed straight from the uploaded prior workbook's
  // "Prior Invoice Log" (col C = Cost Code #, col F = Cost Code Name). This is
  // the most authoritative source for this requisition, so it takes precedence
  // over the server catalog when auto-filling. Built when the workbook is chosen.
  const[wbCoaMap,setWbCoaMap]=useState({});
  const parseWorkbookCoaMap=async(file)=>{
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      const ws=wb.Sheets['Prior Invoice Log'];
      if(!ws){setWbCoaMap({});return;}
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,blankrows:false});
      const m={};
      // Data starts after the 2-row header; col C=index 2 (code), col F=index 5 (name).
      for(const row of rows){
        const code=row&&row[2]!=null?String(row[2]).trim():'';
        const name=row&&row[5]!=null?String(row[5]).trim():'';
        if(!code||!/\d/.test(code))continue;          // skip headers / subtotal rows (no numeric code)
        if(/total/i.test(name))continue;               // skip "X Total" subtotal label rows
        if(name&&!m[code])m[code]=name;                // first (top-most) name for a code wins
      }
      setWbCoaMap(m);
    }catch{setWbCoaMap({});}
  };

  // Read each uploaded invoice with Claude and append an editable card.
  const onRfInvoices=async(e)=>{const files=[...e.target.files];e.target.value='';if(!files.length)return;
    setRfReadErr('');setRfReading(n=>n+files.length);
    for(const f of files){
      try{const r=await api.readRequisitionInvoice(entityId,f);
        setRfCards(cards=>[...cards,{
          _id:Date.now()+'-'+Math.random().toString(36).slice(2,7),
          filename:r.filename||f.name,
          cost_code:r.cost_code||'',
          cost_code_name:r.cost_code_name||'',
          vendor:r.vendor||'',
          bill:r.bill_number||'',
          amount:r.amount!=null?String(r.amount):'',
          date:r.invoice_date||'',
          confidence:r.confidence||'new',
          // Original bytes echoed by read-invoice; held client-side and sent at
          // roll-forward (invoices are NOT persisted server-side until then).
          file_b64:r.file_b64||null,
          original_name:r.original_name||r.filename||f.name,
          mime_type:r.mime_type||f.type||null,
        }]);
      }catch(ex){setRfReadErr(ex.message);}
      finally{setRfReading(n=>Math.max(0,n-1));}
    }};
  const updateCard=(id,field,val)=>setRfCards(cards=>cards.map(c=>{
    if(c._id!==id)return c;
    const next={...c,[field]:val};
    // When the Cost Code changes, auto-fill the Cost Code Name. The uploaded
    // prior workbook's Prior Invoice Log is the most authoritative source, so it
    // wins; the server catalog is the fallback. Only overwrite the name if it's
    // blank or still matches the previous code's auto value, so manual edits stick.
    if(field==='cost_code'){
      const nameFor=code=>{const k=String(code).trim();if(wbCoaMap[k])return wbCoaMap[k];const h=coaMap[k];return h?(h.cost_code_name||''):'';};
      const newName=nameFor(val);
      const prevName=nameFor(c.cost_code);
      const nameIsAuto=!c.cost_code_name||(prevName&&c.cost_code_name===prevName);
      if(newName&&nameIsAuto)next.cost_code_name=newName;
    }
    return next;
  }));
  const removeCard=id=>setRfCards(cards=>cards.filter(c=>c._id!==id));

  const runRollForward=async(force=false)=>{
    if(!rfFile){setRfErr('Upload the prior requisition workbook (.xlsx) first.');return;}
    const newCurrent=rfCards.map(c=>{
      const amount=c.amount!==''?parseFloat(String(c.amount).replace(/[$,]/g,'')):NaN;
      return {code:c.cost_code||undefined,name:c.cost_code_name||undefined,vendor:c.vendor||undefined,bill:c.bill||undefined,date:c.date||undefined,...(Number.isFinite(amount)?{amount}:{})};
    }).filter(x=>Number.isFinite(x.amount));
    if(!newCurrent.length){setRfErr('Add at least one invoice with an amount before rolling forward.');return;}
    // Send the kept invoices (with their original bytes) to be persisted now.
    const invoices=rfCards.map(c=>({
      vendor:c.vendor||null,bill_number:c.bill||null,
      amount:c.amount!==''?c.amount:null,invoice_date:c.date||null,
      cost_code:c.cost_code||null,cost_code_name:c.cost_code_name||null,confidence:c.confidence||null,
      original_name:c.original_name||c.filename||null,mime_type:c.mime_type||null,file_b64:c.file_b64||null,
    }));
    setRfBusy(true);setRfErr('');setRfDetail(null);setRfResult(null);
    try{
      const {blob,filename,summary,failedChecks,workpaperFolder,workpaperSaved,packetFileId,packetFileName,forced}=await api.rollForwardRequisition(entityId,rfFile,newCurrent,{reqNumber:rfReqNum,asOfDate:rfAsOf,invoices,force});
      const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
      // Also download the invoice-packet PDF into the user's Downloads folder
      // (it is retained in Workpapers too). Fetch the saved entity-file as a blob
      // and trigger a second download; a short delay avoids the browser
      // suppressing the back-to-back download.
      if(packetFileId){
        try{
          const presp=await fetch(api.downloadEntityFile(packetFileId));
          if(presp.ok){
            const pblob=await presp.blob();
            const purl=URL.createObjectURL(pblob);const pa=document.createElement('a');
            pa.href=purl;pa.download=packetFileName||'Invoice Packet.pdf';
            document.body.appendChild(pa);setTimeout(()=>{pa.click();pa.remove();URL.revokeObjectURL(purl);},400);
          }
        }catch(pe){/* packet download is best-effort; the workbook already downloaded */}
      }
      // Success: clear the working set (invoices/workbook/req#), keep the result banner.
      setReqState(cur=>({...cur,cards:[],file:null,reqNum:'',detail:null,result:{filename,summary,failedChecks,count:newCurrent.length,workpaperFolder,workpaperSaved,forced}}));
    }catch(e){setRfErr(e.message);if(e.detail)setRfDetail(e.detail);}
    finally{setRfBusy(false);}};

  // Explicit cancel/clear of the in-progress requisition working set.
  const clearReq=()=>{if(!confirm('Clear the uploaded workbook and all invoices? This cannot be undone.'))return;setReqState({cards:[],reqNum:'',asOf:today(),result:null,detail:null,file:null});setRfErr('');setRfReadErr('');};

  const tierStyle=conf=>conf==='high'?{color:T.green,background:T.greenDim,border:'1px solid '+T.greenBorder}
    :conf==='review'?{color:T.orange,background:T.orangeDim,border:'1px solid '+T.orange+'40'}
    :{color:T.textMuted,background:T.bgElevated,border:'1px solid '+T.border};

  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
      <div><div style={S.h1}>Requisitions</div><div style={S.sub}>{entityName} &mdash; roll forward to the next requisition</div></div>
    </div>
    {!canEdit&&<div style={{...S.card,textAlign:'center',padding:50,color:T.textDim}}>The requisition roll-forward tool is read-only for your account. Contact an administrator if you need to run a requisition.</div>}
    {canEdit&&<>
    {err&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:12}}>{err}</div>}

    <div>
      <div style={S.card}>
        <div style={{...S.h2,marginBottom:6}}>Roll Forward to Next Requisition</div>
        <div style={{fontSize:12,color:T.textMuted,marginBottom:14}}>Upload the <strong>prior requisition workbook</strong> (.xlsx), then add this period's invoices one at a time below &mdash; each invoice is read automatically and its fields pre-filled for you to check. The engine folds the prior Current Invoice Log into the Prior Log, replaces the Current Log with these invoices, re-points cross-sheet references, and runs a reconciliation check before producing the next workbook. The result downloads automatically on success.</div>

        <div style={{marginBottom:14}}>
          <label style={S.label}>Prior requisition workbook (.xlsx)</label>
          <div style={{display:'flex',gap:10,alignItems:'center',marginTop:4}}>
            <div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
              <button style={{...S.btnS,pointerEvents:'none'}}>{rfFile?'Change file':'Choose .xlsx'}</button>
              <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:'pointer'}} onChange={e=>{const f=e.target.files[0];e.target.value='';if(f){setRfFile(f);parseWorkbookCoaMap(f);}}}/></div>
            <span style={{fontSize:12,color:rfFile?T.textBright:T.textMuted}}>{rfFile?rfFile.name:'No file selected'}</span>
          </div>
        </div>

        <div style={S.row}>
          <div style={{...S.col,flex:1}}><label style={S.label}>New Requisition #</label><input style={S.input} type="number" placeholder="e.g. 15" value={rfReqNum} onChange={e=>setRfReqNum(e.target.value)}/></div>
          <div style={{...S.col,flex:1}}><label style={S.label}>As-of Date</label><input style={S.input} type="date" value={rfAsOf} onChange={e=>setRfAsOf(e.target.value)}/></div>
        </div>

        <label style={{...S.label,marginTop:6}}>This period's invoices</label>
        <div style={{position:'relative',border:'1.5px dashed '+T.border,borderRadius:T.radiusXs||8,padding:'22px 16px',textAlign:'center',background:T.bgElevated,marginTop:4}}>
          <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:2}}>Drop invoice PDFs here, or click to upload</div>
          <div style={{fontSize:11,color:T.textMuted}}>Multiple files at once is fine &mdash; each file is read as a separate invoice and its fields are pre-filled.</div>
          <input type="file" accept=".pdf,application/pdf,image/*" multiple disabled={rfBusy} style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:rfBusy?'not-allowed':'pointer'}} onChange={onRfInvoices}/>
        </div>
        {rfReading>0&&<div style={{fontSize:12,color:T.accent,margin:'8px 0'}}>Reading {rfReading} invoice{rfReading===1?'':'s'}&hellip;</div>}
        {rfReadErr&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',margin:'8px 0'}}>{rfReadErr}</div>}

        {rfCards.length>0&&<div style={{marginTop:14}}>
          <div style={{fontSize:12,fontWeight:600,color:T.textMuted,marginBottom:8}}>Invoices read &middot; {rfCards.length}</div>
          {rfCards.map((c,idx)=><div key={c._id} style={{border:'1px solid '+T.border,borderRadius:8,padding:'12px 14px',marginBottom:10,background:'#fff'}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontSize:11,color:T.textMuted}}>#{idx+1}</span>
                <span style={{fontSize:11,color:T.textMuted,maxWidth:240,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.filename}</span>
              </div>
              <button style={{...S.btnD,padding:'4px 10px',fontSize:11}} onClick={()=>removeCard(c._id)}>Remove</button>
            </div>
            <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
              <div style={{flex:'1 1 90px'}}><label style={S.label}>Cost Code</label><input style={S.input} value={c.cost_code} onChange={e=>updateCard(c._id,'cost_code',e.target.value)}/></div>
              <div style={{flex:'2 1 160px'}}><label style={S.label}>Cost Code Name</label><input style={S.input} value={c.cost_code_name} onChange={e=>updateCard(c._id,'cost_code_name',e.target.value)}/></div>
              <div style={{flex:'2 1 160px'}}><label style={S.label}>Vendor</label><input style={S.input} value={c.vendor} onChange={e=>updateCard(c._id,'vendor',e.target.value)}/></div>
              <div style={{flex:'1 1 120px'}}><label style={S.label}>Bill #</label><input style={S.input} value={c.bill} onChange={e=>updateCard(c._id,'bill',e.target.value)}/></div>
              <div style={{flex:'1 1 110px'}}><label style={S.label}>Amount</label><input style={S.input} value={c.amount} onChange={e=>updateCard(c._id,'amount',e.target.value)}/></div>
              <div style={{flex:'1 1 120px'}}><label style={S.label}>Invoice Date</label><input style={S.input} type="date" value={c.date} onChange={e=>updateCard(c._id,'date',e.target.value)}/></div>
            </div>
          </div>)}
        </div>}

        {rfErr&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',margin:'10px 0'}}>{rfErr}</div>}
        <div style={{display:'flex',gap:10,marginTop:12,alignItems:'center'}}>
          <button style={S.btnP} disabled={rfBusy||rfCards.length===0} onClick={runRollForward}>{rfBusy?'Rolling forward...':'Roll Forward & Download'+(rfCards.length?' ('+rfCards.length+')':'')}</button>
          {(rfCards.length>0||rfFile)&&<button style={S.btnS} disabled={rfBusy} onClick={clearReq}>Cancel</button>}
        </div>
      </div>

      {rfResult&&<div style={{...S.card,background:rfResult.forced?T.redDim:T.greenDim,borderColor:rfResult.forced?T.red+'40':T.greenBorder}}>
        <div style={{fontWeight:700,color:rfResult.forced?T.red:T.green,marginBottom:8}}>{rfResult.forced?'Forced roll-forward':'Roll-forward complete'} &mdash; {rfResult.filename} downloaded</div>
        <div style={{fontSize:12,color:T.text,marginBottom:10}}>{rfResult.count} current-period invoice line{rfResult.count===1?'':'s'} folded forward. {rfResult.forced?'This file was produced despite a failed required check \u2014 review and hand-correct the flagged lines below before relying on it.':'Reconciliation checks passed:'}</div>
        {rfResult.summary&&<div style={{display:'flex',gap:14,flexWrap:'wrap'}}>
          {[['Checks',rfResult.summary.total],['Passed',rfResult.summary.passed],['Required failed',rfResult.summary.requiredFailed],['Advisory failed',rfResult.summary.recommendedFailed]].map(([k,v])=>
            <div key={k} style={{flex:'1 1 120px',textAlign:'center'}}>
              <div style={{fontSize:22,fontWeight:700,color:k==='Required failed'&&v>0?T.red:T.textBright}}>{v!=null?v:'—'}</div>
              <div style={{fontSize:10,color:T.textMuted,marginTop:2,textTransform:'uppercase',letterSpacing:'0.05em'}}>{k}</div></div>)}
        </div>}
        {rfResult.failedChecks&&rfResult.failedChecks.length>0&&<div style={{marginTop:12,paddingTop:10,borderTop:'1px solid '+T.greenBorder}}>
          <div style={{fontSize:11,fontWeight:700,color:T.orange,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:6}}>Advisory checks not evaluated / not passed</div>
          <table style={S.table}><thead><tr><th style={S.th}>Check</th><th style={S.th}>Level</th><th style={S.thR}>Expected</th><th style={S.thR}>Actual</th><th style={S.th}>Detail</th></tr></thead>
            <tbody>{rfResult.failedChecks.map((c,i)=><tr key={i}>
              <td style={{...S.td,fontWeight:600,color:T.textBright}}>{c.id}</td>
              <td style={S.td}>{c.level}</td>
              <td style={S.tdR}>{c.expected!=null?Number(c.expected).toLocaleString(undefined,{maximumFractionDigits:2}):'—'}</td>
              <td style={S.tdR}>{c.actual!=null?Number(c.actual).toLocaleString(undefined,{maximumFractionDigits:2}):'—'}</td>
              <td style={{...S.td,fontSize:11,color:T.textMuted}}>{c.detail}</td></tr>)}</tbody></table>
        </div>}
        {rfResult.workpaperFolder&&<div style={{fontSize:12,color:T.text,marginTop:12,paddingTop:10,borderTop:'1px solid '+T.greenBorder}}>
          Saved to Workpapers: <strong>{rfResult.workpaperFolder}</strong>
          {rfResult.workpaperSaved&&<span style={{color:T.textMuted}}> &mdash; {[rfResult.workpaperSaved.workbook?'report':null,rfResult.workpaperSaved.packet?'invoice packet':null].filter(Boolean).join(' + ')||'no files'}</span>}
        </div>}
      </div>}

      {rfDetail&&rfDetail.checks&&<div style={{...S.card,background:T.redDim,borderColor:T.red+'40'}}>
        <div style={{fontWeight:700,color:T.red,marginBottom:8}}>Reconciliation failed &mdash; workbook not produced</div>
        <div style={{fontSize:12,color:T.text,marginBottom:10}}>A roll-forward only moves data, so a failure means a mechanical issue (a dropped amount, a shifted reference, or a stale subtotal range). The failing checks:</div>
        <table style={S.table}><thead><tr><th style={S.th}>Check</th><th style={S.th}>Level</th><th style={S.thR}>Expected</th><th style={S.thR}>Actual</th><th style={S.th}>Detail</th></tr></thead>
          <tbody>{rfDetail.checks.filter(c=>!c.pass).map((c,i)=><tr key={i}>
            <td style={{...S.td,fontWeight:600,color:T.textBright}}>{c.id}</td>
            <td style={S.td}>{c.level}</td>
            <td style={S.tdR}>{c.expected!=null?Number(c.expected).toLocaleString(undefined,{maximumFractionDigits:2}):'—'}</td>
            <td style={S.tdR}>{c.actual!=null?Number(c.actual).toLocaleString(undefined,{maximumFractionDigits:2}):'—'}</td>
            <td style={{...S.td,fontSize:11,color:T.textMuted}}>{c.detail}</td></tr>)}</tbody></table>
        <div style={{marginTop:12,display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
          <button style={{...S.btnP,background:T.red,borderColor:T.red}} disabled={rfBusy} onClick={()=>runRollForward(true)}>{rfBusy?'Rolling forward...':'Force roll-forward & download anyway'}</button>
          <span style={{fontSize:11,color:T.textMuted}}>Produces the file despite the failed check(s) so you can fix the flagged lines by hand. Any prepopulation beats starting from scratch.</span>
        </div>
      </div>}
    </div>
    </>}
  </div>);}

function EntityManagement({refresh,entities,activeEntity,setActiveEntity}){
  const[showAdd,setShowAdd]=useState(false);const[bulk,setBulk]=useState(false);
  const[name,setName]=useState('');const[newType,setNewType]=useState('accounting');const[newDisplayId,setNewDisplayId]=useState('');const[bulkText,setBulkText]=useState('');const[err,setErr]=useState('');
  const[typeBusy,setTypeBusy]=useState(null);// entity id whose type is being toggled
  const[importing,setImporting]=useState(null);// entity id being imported into
  const[importAsOf,setImportAsOf]=useState('2024-12-31');const[importMsg,setImportMsg]=useState('');const[importErr,setImportErr]=useState('');const[importBusy,setImportBusy]=useState(false);
  const onTBFile=async e=>{const file=e.target.files[0];if(!file||!importing)return;e.target.value='';setImportBusy(true);setImportMsg('');setImportErr('');
    try{const r=await api.importTrialBalance(importing,file,importAsOf);setImportMsg('Imported '+r.accounts_imported+' accounts. Opening JE: $'+r.total_debit.toFixed(2)+' debits / $'+r.total_credit.toFixed(2)+' credits.'+(r.plug_added?' (Difference plugged to Retained Earnings.)':''));}
    catch(ex){setImportErr(ex.message);}finally{setImportBusy(false);}};
  // ── GL detail import (two-step: preview/map → import) ──
  const[glEntity,setGlEntity]=useState(null);// entity id for GL import modal
  const[glFile,setGlFile]=useState(null);
  const[glStep,setGlStep]=useState('upload');// 'upload' | 'map' | 'done'
  const[glPreview,setGlPreview]=useState(null);// {columns,total_rows,suggested,preview}
  const[glMap,setGlMap]=useState({});// field -> column name
  const[glFused,setGlFused]=useState(false);const[glFusedDelim,setGlFusedDelim]=useState('auto');
  const[glBusy,setGlBusy]=useState(false);const[glErr,setGlErr]=useState('');const[glResult,setGlResult]=useState(null);const[glUnbalanced,setGlUnbalanced]=useState(null);
  const resetGl=()=>{setGlEntity(null);setGlFile(null);setGlStep('upload');setGlPreview(null);setGlMap({});setGlFused(false);setGlFusedDelim('auto');setGlBusy(false);setGlErr('');setGlResult(null);setGlUnbalanced(null);};
  const onGlFile=async e=>{const file=e.target.files[0];if(!file||!glEntity)return;e.target.value='';setGlBusy(true);setGlErr('');
    try{const r=await api.importGLPreview(glEntity,file);setGlFile(file);setGlPreview(r);
      const s=r.suggested||{};setGlMap({account_number:s.account_number||'',account_name:s.account_name||'',transaction_date:s.transaction_date||'',description:s.description||'',memo:s.memo||'',debit:s.debit||'',credit:s.credit||'',reference:s.reference||'',running_balance:s.running_balance||'',class:s.class||'',location:s.location||''});
      setGlFused(!!s.fused);setGlFusedDelim('auto');setGlStep('map');}
    catch(ex){setGlErr(ex.message);}finally{setGlBusy(false);}};
  const runGlImport=async()=>{if(!glFile||!glEntity)return;setGlBusy(true);setGlErr('');setGlUnbalanced(null);
    const mapping={...glMap,fused:glFused,fused_column:glFused?(glMap.account_number||glPreview?.suggested?.fused_column):null,fused_delimiter:glFusedDelim==='auto'?null:glFusedDelim};
    try{const r=await api.importGL(glEntity,glFile,mapping);setGlResult(r);setGlStep('done');}
    catch(ex){setGlErr(ex.message);if(ex.detail&&ex.detail.unbalanced_groups)setGlUnbalanced(ex.detail);}finally{setGlBusy(false);}};
  const GL_FIELDS=[{k:'account_number',label:'Account Number',req:true},{k:'account_name',label:'Account Name',req:true},{k:'transaction_date',label:'Transaction Date',req:true},{k:'description',label:'Description',req:false},{k:'memo',label:'Memo',req:false},{k:'debit',label:'Debit',req:true},{k:'credit',label:'Credit',req:true},{k:'reference',label:'Reference / Doc # (groups lines into JEs)',req:false},{k:'running_balance',label:'Running Balance (verification only, not stored)',req:false},{k:'class',label:'Class (e.g. investor — tracked per line)',req:false},{k:'location',label:'Location (e.g. deal / asset — tracked per line)',req:false}];
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}>
    <div><div style={S.h1}>Entity Management</div><div style={S.sub}>{entities.length} entities</div></div>
    <div style={{display:'flex',gap:10}}><button style={S.btnS} onClick={()=>{setBulk(!bulk);setShowAdd(false);}}>{bulk?'Cancel':'Bulk Import'}</button><button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setBulk(false);}}>{showAdd?'Cancel':'+ Add Entity'}</button></div></div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:12}}>Create New Entity</div>
      <div style={S.row}><div style={{...S.col,flex:3}}><label style={S.label}>Entity Name</label><input style={S.input} placeholder="e.g. CLR Fund I LP" value={name} onChange={e=>setName(e.target.value)}/></div>
        <div style={{...S.col,flex:1}}><label style={S.label}>Entity ID</label><input style={S.input} placeholder="e.g. 0005 B1a" value={newDisplayId} onChange={e=>setNewDisplayId(e.target.value)}/></div>
        <div style={{...S.col,flex:2}}><label style={S.label}>Entity Type</label><select style={S.input} value={newType} onChange={e=>setNewType(e.target.value)}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option></select></div></div>
      {err&&<div style={S.err}>{err}</div>}
      <div style={{fontSize:11,color:T.textMuted,marginBottom:10}}>A default chart of accounts will be created. You can replace it by importing a trial balance from the entity row. Development-project entities unlock the Requisitions coding tools.</div>
      <button style={S.btnP} onClick={async()=>{if(!name.trim()){setErr('Name required');return;}try{await api.createEntity(name.trim(),newType,newDisplayId.trim());setName('');setNewType('accounting');setNewDisplayId('');setShowAdd(false);setErr('');refresh();}catch(e){setErr(e.message);}}}>Create Entity</button></div>}
    {bulk&&<div style={{...S.card,borderColor:T.accent+'40'}}><div style={{...S.h2,marginBottom:8}}>Bulk Import Entities</div><div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>One entity name per line</div>
      <textarea style={{...S.input,height:160,fontFamily:'monospace',fontSize:12,resize:'vertical'}} value={bulkText} onChange={e=>setBulkText(e.target.value)}/>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:10}} onClick={async()=>{const names=bulkText.split('\n').map(l=>l.trim()).filter(Boolean);if(!names.length){setErr('None');return;}try{for(const n of names)await api.createEntity(n);setBulkText('');setBulk(false);refresh();}catch(e){setErr(e.message);}}}>Import</button></div>}
    <div style={{...S.cardFlush,overflowX:'auto'}}><table style={{...S.table,minWidth:980}}><thead><tr><th style={S.th}>Entity</th><th style={{...S.th,width:720,minWidth:720}}>Actions</th></tr></thead>
      <tbody>{entities.sort((a,b)=>a.name.localeCompare(b.name)).map(e=><tr key={e.id} style={e.id===activeEntity?{background:T.accentDim}:{}}>
        <td style={{...S.td,fontWeight:600,color:T.textBright}}>{e.display_id&&<span style={{marginRight:8,fontSize:11,fontWeight:700,color:T.textMuted,fontFamily:'monospace'}}>{e.display_id}</span>}{e.name}{e.entity_type==='development'&&<span style={{marginLeft:8,fontSize:9,fontWeight:700,color:T.green,background:T.greenDim,border:'1px solid '+T.greenBorder,borderRadius:4,padding:'2px 6px',textTransform:'uppercase',letterSpacing:'0.05em',verticalAlign:'middle'}}>Dev Project</span>}{e.entity_type==='shell'&&<span style={{marginLeft:8,fontSize:9,fontWeight:700,color:T.teal,background:T.tealDim,border:'1px solid '+T.teal+'40',borderRadius:4,padding:'2px 6px',textTransform:'uppercase',letterSpacing:'0.05em',verticalAlign:'middle'}}>Shell</span>}</td>
        <td style={S.td}><div style={{display:'flex',gap:8,flexWrap:'nowrap',whiteSpace:'nowrap'}}>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0}} onClick={()=>setActiveEntity(e.id)}>Select</button>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0,color:T.accent,borderColor:T.accent+'40'}} onClick={()=>{setImporting(e.id);setImportMsg('');setImportErr('');}}>Import Trial Balance</button>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0,color:T.accent,borderColor:T.accent+'40'}} onClick={()=>{resetGl();setGlEntity(e.id);}}>Import General Ledger Detail</button>
          <select style={{...S.inputSm,padding:'5px 8px',fontSize:11,flexShrink:0,width:'auto'}} disabled={typeBusy===e.id} title="Entity type" value={e.entity_type||'accounting'} onChange={async(ev)=>{const next=ev.target.value;if(next===e.entity_type)return;if(!confirm('Set "'+e.name+'" to '+({accounting:'Accounting',development:'Development Project',shell:'Shell'}[next]||next)+'?'))return;setTypeBusy(e.id);try{await api.updateEntity(e.id,{entity_type:next});await refresh();}catch(ex){alert(ex.message);}finally{setTypeBusy(null);}}}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option></select>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0}} title="Set the short Entity ID used as the invoice-packet filename prefix" onClick={async()=>{const cur=e.display_id||'';const v=prompt('Entity ID for "'+e.name+'"\n(used as the invoice-packet filename prefix; leave blank to use the entity name):',cur);if(v===null)return;try{await api.updateEntity(e.id,{display_id:v.trim()});await refresh();}catch(ex){alert(ex.message);}}}>Edit ID</button>
          <button style={{...S.btnD,padding:'5px 12px',fontSize:11,flexShrink:0}} onClick={async()=>{if(!confirm('Delete entity '+e.name+' and all its data?'))return;await api.deleteEntity(e.id);const r=await refresh();if(activeEntity===e.id)setActiveEntity(r[0]?.id||null);}}>Delete</button>
        </div></td></tr>)}</tbody></table></div>
    {importing&&<div style={S.modal} onClick={()=>{if(!importBusy)setImporting(null);}}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:560}} onClick={ev=>ev.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setImporting(null)}>&times;</button>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:6}}>Import Trial Balance</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:18}}>Entity: <strong style={{color:T.textBright}}>{entities.find(en=>en.id===importing)?.name}</strong></div>
      <div style={{...S.card,background:T.bgElevated,padding:14,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:600,color:T.textBright,marginBottom:6}}>File requirements</div>
        <div style={{fontSize:11,color:T.textMuted,lineHeight:1.6}}>Excel (.xlsx) or CSV with columns:<br/>
          &bull; <strong>Account Number</strong> (or Code, Acct, Number)<br/>
          &bull; <strong>Account Name</strong> (or Name, Description)<br/>
          &bull; <strong>Amount</strong> (or Balance) &mdash; or separate <strong>Debit</strong> and <strong>Credit</strong> columns<br/><br/>
          Account types are auto-derived from the account number:<br/>
          &bull; 10000-19999 = Asset &nbsp;&bull; 20000-29999 = Liability<br/>
          &bull; 30000-39999 = Equity &nbsp;&bull; 40000-49999 = Revenue<br/>
          &bull; 50000-69999 = Expense &nbsp;&bull; 70000+ = Revenue</div></div>
      <div style={{marginBottom:14}}><label style={S.label}>As of Date</label><input style={S.input} type="date" value={importAsOf} onChange={e=>setImportAsOf(e.target.value)}/></div>
      <div style={{fontSize:11,color:T.orange,marginBottom:14,padding:10,background:T.orangeDim,borderRadius:6,border:'1px solid '+T.orange+'30'}}>
        <strong>Warning:</strong> This will replace the entire chart of accounts for this entity and create an opening balance journal entry. Existing accounts and journal entries (if any) will not be deleted, but the COA will be rebuilt from your file.
      </div>
      {importMsg&&<div style={{...S.success,padding:10,background:T.greenDim,borderRadius:6,border:'1px solid '+T.greenBorder,marginBottom:10}}>{importMsg}</div>}
      {importErr&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:10}}>{importErr}</div>}
      <div style={{display:'flex',gap:10,alignItems:'center'}}>
        <div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
          <button style={{...S.btnP,pointerEvents:'none',opacity:importBusy?.6:1}}>{importBusy?'Importing...':'Choose File & Import'}</button>
          <input type="file" accept=".csv,.xlsx,.xls" disabled={importBusy} style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:importBusy?'not-allowed':'pointer'}} onChange={onTBFile}/></div>
        <button style={S.btnS} onClick={()=>setImporting(null)} disabled={importBusy}>Close</button>
      </div></div></div>}
    {glEntity&&<div style={S.modal} onClick={()=>{if(!glBusy)resetGl();}}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:glStep==='map'?920:600}} onClick={ev=>ev.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>{if(!glBusy)resetGl();}}>&times;</button>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:6}}>Import General Ledger Detail</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:18}}>Entity: <strong style={{color:T.textBright}}>{entities.find(en=>en.id===glEntity)?.name}</strong></div>
      {glStep==='upload'&&<>
        <div style={{...S.card,background:T.bgElevated,padding:14,marginBottom:14}}>
          <div style={{fontSize:12,fontWeight:600,color:T.textBright,marginBottom:6}}>How it works</div>
          <div style={{fontSize:11,color:T.textMuted,lineHeight:1.6}}>Upload an inception-to-date GL detail export (.xlsx or .csv) from any accounting system. We'll detect the columns and let you map them &mdash; account number, name, date, description/memo, debit and credit. If account number and name share one cell (e.g. <em>1000 &middot; Cash</em>), we'll split them.<br/><br/>Transactions are grouped into balanced journal entries by date + reference (if a reference column exists), otherwise into one journal entry per transaction date. Every entry must balance (debits = credits); if any date is out of balance, the import is halted and nothing is saved.</div></div>
        <div style={{fontSize:11,color:T.orange,marginBottom:14,padding:10,background:T.orangeDim,borderRadius:6,border:'1px solid '+T.orange+'30'}}>
          <strong>Warning:</strong> Importing GL detail rebuilds the chart of accounts from your file and replaces any prior trial-balance or GL import on this entity (latest import wins), so the two never double-count.</div>
        {glErr&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:10}}>{glErr}</div>}
        <div style={{display:'flex',gap:10,alignItems:'center'}}>
          <div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
            <button style={{...S.btnP,pointerEvents:'none',opacity:glBusy?.6:1}}>{glBusy?'Reading...':'Choose File'}</button>
            <input type="file" accept=".csv,.xlsx,.xls" disabled={glBusy} style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:glBusy?'not-allowed':'pointer'}} onChange={onGlFile}/></div>
          <button style={S.btnS} onClick={resetGl} disabled={glBusy}>Cancel</button>
        </div></>}
      {glStep==='map'&&glPreview&&<>
        <div style={{fontSize:12,color:T.textMuted,marginBottom:12}}>Detected <strong style={{color:T.textBright}}>{glPreview.total_rows}</strong> rows. Map your columns to CloudLedger fields, then import. Required fields are marked <span style={{color:T.red}}>*</span>.</div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px 16px',marginBottom:14}}>
          {GL_FIELDS.map(f=><div key={f.k}>
            <label style={{...S.label,fontSize:11}}>{f.label}{f.req&&<span style={{color:T.red}}> *</span>}</label>
            <select style={S.input} value={glMap[f.k]||''} onChange={ev=>setGlMap(m=>({...m,[f.k]:ev.target.value}))}>
              <option value="">— none —</option>
              {glPreview.columns.map(c=><option key={c} value={c}>{c}</option>)}
            </select></div>)}
        </div>
        <div style={{...S.card,background:T.bgElevated,padding:12,marginBottom:14}}>
          <label style={{display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.textBright,cursor:'pointer'}}>
            <input type="checkbox" checked={glFused} onChange={ev=>setGlFused(ev.target.checked)}/>
            Account Number column also contains the Account Name in one cell (split it)</label>
          {glFused&&<div style={{marginTop:10,display:'flex',alignItems:'center',gap:10}}>
            <span style={{fontSize:11,color:T.textMuted}}>Separator:</span>
            <select style={{...S.input,width:200,margin:0}} value={glFusedDelim} onChange={ev=>setGlFusedDelim(ev.target.value)}>
              <option value="auto">Auto-detect</option><option value=" · ">space · space</option><option value=" - ">space - space</option><option value=":">colon</option><option value=" ">first space</option></select>
            <span style={{fontSize:11,color:T.textMuted}}>e.g. "1000 · Cash" → 1000 + Cash</span></div>}
        </div>
        <div style={{fontSize:11,fontWeight:600,color:T.textBright,marginBottom:6}}>Preview (first {glPreview.preview.length} rows)</div>
        <div style={{overflowX:'auto',marginBottom:14,border:'1px solid '+T.borderLight,borderRadius:6}}>
          <table style={{...S.table,fontSize:10}}><thead><tr>{glPreview.columns.map(c=><th key={c} style={{...S.th,fontSize:10,whiteSpace:'nowrap',padding:'5px 8px'}}>{c}</th>)}</tr></thead>
            <tbody>{glPreview.preview.map((row,i)=><tr key={i}>{glPreview.columns.map(c=><td key={c} style={{...S.td,fontSize:10,whiteSpace:'nowrap',padding:'4px 8px'}}>{String(row[c]??'')}</td>)}</tr>)}</tbody></table></div>
        {glErr&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:10}}>{glErr}</div>}
        {glUnbalanced&&glUnbalanced.unbalanced_groups&&<div style={{...S.err,padding:10,background:T.redDim,borderRadius:6,border:'1px solid '+T.red+'30',marginBottom:10}}>
          <div style={{fontWeight:600,marginBottom:6}}>{glUnbalanced.unbalanced_count} {glUnbalanced.grouping==='by_reference'?'reference group(s)':'date(s)'} out of balance &mdash; nothing was imported.</div>
          <div style={{overflowX:'auto'}}><table style={{...S.table,fontSize:11}}><thead><tr>
            <th style={{...S.th,fontSize:11,padding:'4px 8px'}}>{glUnbalanced.grouping==='by_reference'?'Date / Ref':'Date'}</th>
            <th style={{...S.th,fontSize:11,padding:'4px 8px',textAlign:'right'}}>Debits</th>
            <th style={{...S.th,fontSize:11,padding:'4px 8px',textAlign:'right'}}>Credits</th>
            <th style={{...S.th,fontSize:11,padding:'4px 8px',textAlign:'right'}}>Difference</th>
            <th style={{...S.th,fontSize:11,padding:'4px 8px',textAlign:'right'}}>Lines</th></tr></thead>
            <tbody>{glUnbalanced.unbalanced_groups.map((g,i)=><tr key={i}>
              <td style={{...S.td,fontSize:11,padding:'4px 8px',whiteSpace:'nowrap'}}>{g.date}{g.reference?(' / '+g.reference):''}</td>
              <td style={{...S.td,fontSize:11,padding:'4px 8px',textAlign:'right'}}>{g.debit.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
              <td style={{...S.td,fontSize:11,padding:'4px 8px',textAlign:'right'}}>{g.credit.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
              <td style={{...S.td,fontSize:11,padding:'4px 8px',textAlign:'right',color:T.red,fontWeight:600}}>{g.difference.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
              <td style={{...S.td,fontSize:11,padding:'4px 8px',textAlign:'right'}}>{g.lines}</td></tr>)}</tbody></table></div>
          <div style={{fontSize:11,marginTop:6,opacity:.85}}>A balanced general ledger nets to zero within every transaction date. An out-of-balance date usually means a single-sided export or a line posted to the wrong date. Fix the source file and re-import.</div>
        </div>}
        <div style={{display:'flex',gap:10,alignItems:'center'}}>
          <button style={{...S.btnP,opacity:glBusy?.6:1}} disabled={glBusy} onClick={runGlImport}>{glBusy?'Importing...':'Import'}</button>
          <button style={S.btnS} onClick={()=>{setGlStep('upload');setGlErr('');}} disabled={glBusy}>Back</button>
          <button style={S.btnS} onClick={resetGl} disabled={glBusy}>Cancel</button>
        </div></>}
      {glStep==='done'&&glResult&&<>
        <div style={{...S.success,padding:12,background:T.greenDim,borderRadius:6,border:'1px solid '+T.greenBorder,marginBottom:12}}>
          Imported <strong>{glResult.lines_imported}</strong> lines into <strong>{glResult.entries_created}</strong> journal entries across <strong>{glResult.accounts_imported}</strong> accounts.<br/>
          Total debits ${glResult.total_debit.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})} / credits ${glResult.total_credit.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})} &mdash; {glResult.balanced?'balanced ✓':'NOT balanced ✗'}.<br/>
          {(glResult.classes_imported>0||glResult.locations_imported>0)&&<>Dimensions: {glResult.classes_imported>0&&<><strong>{glResult.classes_imported}</strong> classes</>}{glResult.classes_imported>0&&glResult.locations_imported>0&&', '}{glResult.locations_imported>0&&<><strong>{glResult.locations_imported}</strong> locations</>} tracked.<br/></>}
          Grouping: {glResult.grouping==='by_reference'?'by reference / document #':'one journal entry per transaction date'}.{glResult.rows_skipped>0&&<> {glResult.rows_skipped} row(s) skipped (blank/zero/unparseable).</>}
        </div>
        {glResult.persisted&&<div style={{...(glResult.persisted_ok?S.success:S.err),padding:12,borderRadius:6,marginBottom:12,background:glResult.persisted_ok?T.greenDim:T.redDim,border:'1px solid '+(glResult.persisted_ok?T.greenBorder:T.red+'30')}}>
          <strong>Persistence check{glResult.entity_id?' (entity '+glResult.entity_id+')':''}:</strong> {glResult.persisted_ok?'✓ saved':'✗ NOT SAVED'} &mdash; {glResult.persisted.entries} entries, {glResult.persisted.lines} lines, {glResult.persisted.accounts} accounts now in the database for this entity.
        </div>}
        {glResult.verification&&<div style={{...(glResult.verification.mismatches.length?S.err:S.success),padding:12,borderRadius:6,marginBottom:12,background:glResult.verification.mismatches.length?T.redDim:T.greenDim,border:'1px solid '+(glResult.verification.mismatches.length?T.red+'30':T.greenBorder)}}>
          <strong>Running-balance check:</strong> {glResult.verification.matched}/{glResult.verification.checked} accounts match.
          {glResult.verification.mismatches.length>0&&<><br/><span style={{fontSize:11}}>Mismatches (computed vs reported): {glResult.verification.mismatches.map(mm=>mm.code+' ('+mm.computed.toLocaleString()+' vs '+mm.reported.toLocaleString()+')').join(', ')}</span></>}
        </div>}
        <button style={S.btnP} onClick={()=>{const r=refresh&&refresh();resetGl();}}>Done</button>
      </>}
    </div></div>}
  </div>);}

// ═══ User Management (with role editing) ═══
function UserManagement({currentUser}){
  const[users,setUsers]=useState([]);const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({name:'',email:'',password:'',role:'Viewer'});const[err,setErr]=useState('');const[loadErr,setLoadErr]=useState('');
  const[resetId,setResetId]=useState(null);const[resetPw,setResetPw]=useState('');const[resetMsg,setResetMsg]=useState('');
  const[editingRole,setEditingRole]=useState(null);
  const loadUsers=useCallback(()=>{api.getUsers().then(setUsers).catch(e=>setLoadErr(e.message));},[]);
  const[accessUser,setAccessUser]=useState(null);
  const[accessEntities,setAccessEntities]=useState([]); // selected ids
  const[accessAllEntities,setAccessAllEntities]=useState([]); // all entities for picker
  const[accessSaving,setAccessSaving]=useState(false);
  const[accessErr,setAccessErr]=useState('');
  const openAccess=async(u)=>{
    setAccessUser(u);setAccessErr('');setAccessSaving(false);
    try{
      const[ents,acc]=await Promise.all([api.getEntities(),api.getUserEntityAccess(u.id)]);
      setAccessAllEntities(ents);
      setAccessEntities(acc.entity_ids||[]);
    }catch(e){setAccessErr(e.message);}
  };
  const toggleAccessEntity=(id)=>setAccessEntities(prev=>prev.includes(id)?prev.filter(x=>x!==id):[...prev,id]);
  const saveAccess=async()=>{
    setAccessSaving(true);setAccessErr('');
    try{await api.setUserEntityAccess(accessUser.id,accessEntities);setAccessUser(null);}
    catch(e){setAccessErr(e.message);}
    setAccessSaving(false);
  };
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
            {u.id!==currentUser.id&&u.role!=='Admin'&&<button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>openAccess(u)} title="Limit which entities this user can access">Access</button>}
            {u.id!==currentUser.id&&<button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>{setResetId(u.id);setResetPw('');setResetMsg('');}}>Reset PW</button>}
            {u.id!==currentUser.id&&<button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={async()=>{if(!confirm('Delete user '+u.name+'?'))return;await api.deleteUser(u.id);loadUsers();}}>Delete</button>}</div></td>
        </tr>)}</tbody></table></div>
    {accessUser&&<div style={S.modal} onClick={()=>setAccessUser(null)}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:520}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setAccessUser(null)}>&times;</button>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:6}}>Entity Access</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:14}}>User: <strong style={{color:T.textBright}}>{accessUser.name}</strong> ({accessUser.role})</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10,padding:'8px 12px',background:T.bgInset,borderRadius:6}}>
        {accessEntities.length===0?'No restrictions — user can access ALL entities. Check entities below to restrict access to only those.':'User can access only the '+accessEntities.length+' checked entit'+(accessEntities.length===1?'y':'ies')+' below. Uncheck all to grant access to all entities.'}
      </div>
      <div style={{maxHeight:320,overflowY:'auto',border:'1px solid '+T.border,borderRadius:6,marginBottom:12}}>
        {accessAllEntities.map(e=>(
          <label key={e.id} style={{display:'flex',alignItems:'center',padding:'8px 12px',borderBottom:'1px solid '+T.border,cursor:'pointer',gap:10}}>
            <input type="checkbox" checked={accessEntities.includes(e.id)} onChange={()=>toggleAccessEntity(e.id)}/>
            <span style={{color:T.textBright,fontSize:13}}>{e.name}</span>
            {e.code&&<span style={{color:T.textMuted,fontSize:11,fontFamily:'monospace'}}>{e.code}</span>}
          </label>
        ))}
        {accessAllEntities.length===0&&<div style={{padding:16,color:T.textMuted,textAlign:'center'}}>No entities</div>}
      </div>
      {accessErr&&<div style={S.err}>{accessErr}</div>}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
        <button style={S.btnGhost} onClick={()=>setAccessEntities([])} disabled={accessSaving}>Clear (= all access)</button>
        <div style={{display:'flex',gap:8}}>
          <button style={S.btnGhost} onClick={()=>setAccessUser(null)} disabled={accessSaving}>Cancel</button>
          <button style={S.btnP} onClick={saveAccess} disabled={accessSaving}>{accessSaving?'Saving...':'Save'}</button>
        </div>
      </div>
    </div></div>}
        {resetId&&<div style={S.modal} onClick={()=>setResetId(null)}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:400,textAlign:'center'}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setResetId(null)}>&times;</button><div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Reset Password</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:6}}>User: <strong style={{color:T.textBright}}>{users.find(u=>u.id===resetId)?.name}</strong></div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:16,fontFamily:'monospace'}}>{users.find(u=>u.id===resetId)?.email}</div>
      <input style={S.input} type="password" placeholder="New password" value={resetPw} onChange={e=>{setResetPw(e.target.value);setResetMsg('');}}/>
      {resetMsg&&<div style={{fontSize:12,marginTop:8,color:resetMsg.includes('!')?T.green:T.red}}>{resetMsg}</div>}
      <button style={{...S.btnP,width:'100%',padding:11,marginTop:12}} onClick={async()=>{if(resetPw.length<3){setResetMsg('Min 3 chars');return;}try{await api.adminResetPassword(resetId,resetPw);setResetMsg('Password reset!');setTimeout(()=>setResetId(null),1500);}catch(e){setResetMsg(e.message);}}}>Reset Password</button>
    </div></div>}</div>);}
