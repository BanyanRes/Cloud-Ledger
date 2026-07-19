import { useState, useEffect, useCallback, useRef, useMemo, Fragment } from 'react';
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
// Quick date-range presets (previous complete calendar period) for report filters.
const presetRange = (kind) => {
  const iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  const n = new Date(), y = n.getFullYear(), m = n.getMonth();
  if (kind === 'all') return { from: '2015-01-01', to: iso(n) };
  if (kind === 'month') return { from: iso(new Date(y, m - 1, 1)), to: iso(new Date(y, m, 0)) };
  if (kind === 'quarter') { const cq = Math.floor(m / 3); let sy = y, sm = (cq - 1) * 3; if (sm < 0) { sy = y - 1; sm = 9; } return { from: iso(new Date(sy, sm, 1)), to: iso(new Date(sy, sm + 3, 0)) }; }
  if (kind === 'year') return { from: iso(new Date(y - 1, 0, 1)), to: iso(new Date(y - 1, 11, 31)) };
  return { from: '', to: '' };
};
const PRESETS = [['all', 'All'], ['month', 'Last Month'], ['quarter', 'Last Quarter'], ['year', 'Last Year']];
const _iso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
// Split [from,to] into period columns for report column-display modes.
const buildPeriodCols = (from, to, mode) => {
  if (mode === 'total' || !from || !to) return [{ from: from || '', to: to || '', label: 'Amount' }];
  const cols = []; let s = new Date(from + 'T00:00:00'); const end = new Date(to + 'T00:00:00'); let guard = 0;
  while (s <= end && guard++ < 800) {
    const y = s.getFullYear(); let e, label;
    if (mode === 'monthly') { e = new Date(y, s.getMonth() + 1, 0); label = s.toLocaleString('en-US', { month: 'short', year: '2-digit' }); }
    else if (mode === 'quarterly') { const q = Math.floor(s.getMonth() / 3); e = new Date(y, q * 3 + 3, 0); label = 'Q' + (q + 1) + " '" + String(y).slice(2); }
    else { e = new Date(y, 11, 31); label = String(y); }
    cols.push({ from: _iso(s), to: _iso(e > end ? end : e), label });
    s = new Date(e); s.setDate(s.getDate() + 1);
  }
  return cols;
};
// The equal-length window immediately before [from,to] (for comparative columns).
const priorWindow = (from, to) => {
  if (!from || !to) return null;
  const s = new Date(from + 'T00:00:00'), e = new Date(to + 'T00:00:00');
  // Whole-month windows (month/quarter/year) -> previous equivalent CALENDAR
  // period. Otherwise fall back to the equal-length window immediately before.
  const monthStart = s.getDate() === 1;
  const monthEnd = (() => { const n = new Date(e); n.setDate(n.getDate() + 1); return n.getDate() === 1; })();
  if (monthStart && monthEnd) {
    const months = (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
    return { from: _iso(new Date(s.getFullYear(), s.getMonth() - months, 1)), to: _iso(new Date(s.getFullYear(), s.getMonth(), 0)) };
  }
  const days = Math.round((e - s) / 86400000) + 1;
  const pe = new Date(s); pe.setDate(pe.getDate() - 1);
  const ps = new Date(pe); ps.setDate(ps.getDate() - (days - 1));
  return { from: _iso(ps), to: _iso(pe) };
};
const COL_MODES = [['total', 'Total Only'], ['monthly', 'Monthly'], ['quarterly', 'Quarterly'], ['yearly', 'Yearly']];
const fy_start = () => new Date().getFullYear() + '-01-01';
const fmtSize = b => b > 1048576 ? (b/1048576).toFixed(1)+' MB' : (b/1024).toFixed(0)+' KB';
// Display people's names with each part's first letter capitalized (e.g.
// "omar dominguez" -> "Omar Dominguez"), regardless of how they were entered.
// Non-destructive: only upper-cases the first letter of each whitespace-
// separated part, leaving the rest of each part as typed.
const titleName = s => String(s == null ? '' : s).replace(/(^|\s)(\S)/g, (m, p, c) => p + c.toUpperCase());
// Entity-type grouping metadata, shared by Dashboard + Entity Management.
const ENTITY_TYPES = [
  { key:'accounting',  label:'Accounting',  icon:'📒' },
  { key:'development', label:'Development', icon:'🏗️' },
  { key:'shell',       label:'Shell',       icon:'🗂️' },
];
const entTypeOf = e => (e && e.entity_type) || 'accounting';
const groupByType = list => {
  const g = { accounting:[], development:[], shell:[] };
  (list||[]).forEach(e => { (g[entTypeOf(e)] || (g[entTypeOf(e)]=[])).push(e); });
  return g;
};
const acctLabel = (code, name) => code + ' - ' + name;
// Per-entity relabel of the Class dimension. Turnkey Rail (TURNKEYR) is an
// operating rail company, not a development entity, so it has no requisition /
// Req# to tie invoices to draws — it uses the Class dimension to tag each
// invoice's Pay Application (Bill.com already syncs Class). The underlying
// dimension is unchanged; only the on-screen label differs. _activeEntityCode is
// refreshed by <App> on every render (top-down), so child components read the
// current entity's term. Extend the map to relabel Class for other entities.
let _activeEntityCode = null;
const CLASS_DIM_LABELS = { TURNKEYR: 'Pay Application' };
const classTerm = () => CLASS_DIM_LABELS[_activeEntityCode] || 'Class';
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
  const[dimProjects,setDimProjects]=useState([]);
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts);api.getTurnkeyProjects().then(setProjects).catch(()=>setProjects([]));api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>setDimProjects([]));api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  // Turnkey entities use the Turnkey project picker; all other (non-shell) entities
  // use the project dimension, and the field is always shown with an inline add.
  const useDimProjects=!isTurnkeyEntity&&dimsEnabled;
  const showProject=isTurnkeyEntity||useDimProjects;
  const showLocation=dimsEnabled&&locations.length>0;const showClass=dimsEnabled&&classes.length>0;
  // Inline "+ new project" from a JE line: prompt, create, refresh, select it on that line.
  const addProjectInline=async(i)=>{
    const name=(prompt('New project name or code (e.g. P-10100.001):')||'').trim();
    if(!name) return;
    try{ const p=await api.createProject(entityId,{name,code:name});
      const list=await api.getProjects(entityId); setDimProjects(list||[]);
      updateLine(i,'project_id',p.id);
    }catch(e){ alert(e.message); }
  };
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:'',description:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  // Single "Dimensions" column: one dropdown per line that lists every applicable
  // dimension value (Project / Location / Class) and lets the user pick exactly ONE.
  // The selected option's value is a tagged string ("project:ID" | "location:ID" |
  // "class:ID"); applying it sets that one dimension and clears the other two so a
  // line never carries more than one dimension at a time.
  const showDims=showProject||showLocation||showClass;
  const projOpts=useDimProjects
    ?dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)}))
    :projects.map(pr=>({v:'project:'+pr.turnkey_project_id,label:'Project — '+pr.project_code+' — '+pr.project_name}));
  const locOpts=locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name}));
  const clsOpts=classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name}));
  const lineDimValue=l=>l.project_id?'project:'+l.project_id:l.location_id?'location:'+l.location_id:l.class_id?'class:'+l.class_id:'';
  const setLineDim=(i,val)=>{
    if(val==='__new__'){addProjectInline(i);return;}
    const[kind,id]=val?val.split(':'):['',''];
    setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,
      project_id:kind==='project'?id:null,
      location_id:kind==='location'?id:null,
      class_id:kind==='class'?id:null}:l)}));
  };
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

  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,width:'min(1200px, 96vw)',maxWidth:'96vw',height:'auto',maxHeight:'92vh',resize:'both',overflow:'auto',minWidth:'min(560px, 96vw)',minHeight:360}} onClick={e=>e.stopPropagation()}>
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
    <div style={{...S.cardFlush,marginBottom:16,maxHeight:'52vh',overflowY:'auto'}}><table className="cl-colresize" style={S.table}><thead style={{position:'sticky',top:0,zIndex:2,background:T.bgElevated}}><tr><th style={{...S.th,minWidth:300}}>Account</th>{showDims&&<th style={{...S.th,width:140}}>Dimension</th>}<th style={S.th}>Description</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} title={l.account_code?acctLabel(l.account_code,(accounts.find(a=>a.code===l.account_code)||{}).name||''):''} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select account...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code} title={acctLabel(a.code,a.name)}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        {showDims&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={lineDimValue(l)} onChange={e=>setLineDim(i,e.target.value)}><option value="">— none —</option>{showProject&&<optgroup label="Project">{projOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}{useDimProjects&&<option value="__new__">+ New project…</option>}</optgroup>}{showLocation&&<optgroup label="Location">{locOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</optgroup>}{showClass&&<optgroup label={classTerm()}>{clsOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</optgroup>}</select></td>}
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={S.input} placeholder="(optional)" value={l.description||''} onChange={e=>updateLine(i,'description',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.debit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'debit',f);}} onBlur={e=>updateLine(i,'debit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} placeholder="0.00" value={l.credit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'credit',f);}} onBlur={e=>updateLine(i,'credit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td colSpan={2+(showDims?1:0)} style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
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
function GlobalSearch({entities,activeEntity,onSelectEntity,onGo,onPickJE,onPickAccount}){
  const[open,setOpen]=useState(false);
  const[q,setQ]=useState("");
  const[accounts,setAccounts]=useState([]);
  const[entries,setEntries]=useState([]);
  const[loaded,setLoaded]=useState(false);
  // Lazy-load the active entity's accounts + recent entries the first time the box opens.
  useEffect(()=>{
    if(!open||!activeEntity||loaded)return;
    let cancelled=false;
    (async()=>{
      try{
        const[a,e]=await Promise.all([api.getAccounts(activeEntity).catch(()=>[]),api.getEntries(activeEntity).catch(()=>[])]);
        if(!cancelled){setAccounts(a||[]);setEntries(e||[]);setLoaded(true);}
      }catch{ if(!cancelled)setLoaded(true); }
    })();
    return()=>{cancelled=true;};
  },[open,activeEntity,loaded]);
  // Reset the cached data when the active entity changes.
  useEffect(()=>{setLoaded(false);setAccounts([]);setEntries([]);},[activeEntity]);
  const t=q.trim().toLowerCase();
  const entHits = t? entities.filter(e=>(e.name||"").toLowerCase().includes(t)||(e.code||"").toLowerCase().includes(t)).slice(0,6):[];
  const acctHits = t? accounts.filter(a=>(a.code||"").toLowerCase().includes(t)||(a.name||"").toLowerCase().includes(t)).slice(0,6):[];
  const jeHits = t? entries.filter(e=>{
    const enStr=String(e.entry_num);const jn="je-"+enStr.padStart(4,"0");const qn=t.replace(/^je[-\s]?/,"").replace(/^0+(?=\d)/,"");
    if(jn.includes(t)||enStr.includes(t)||(/^\d+$/.test(qn)&&(enStr===qn||enStr.includes(qn))))return true;
    if((e.date||"").toLowerCase().includes(t))return true;
    if((e.memo||"").toLowerCase().includes(t))return true;
    const amtQ=t.replace(/[$,\s]/g,"");
    if(amtQ&&/^[0-9.]+$/.test(amtQ)&&(e.lines||[]).some(l=>{const d=Number(l.debit||0),c=Number(l.credit||0);return String(d).includes(amtQ)||String(c).includes(amtQ)||d.toFixed(2).includes(amtQ)||c.toFixed(2).includes(amtQ);}))return true;
    return (e.lines||[]).some(l=>(l.account_code||"").toLowerCase().includes(t)||(l.description||"").toLowerCase().includes(t));
  }).sort((a,b)=>{const _qn=t.replace(/^je[- ]?/,"").replace(/^0+/,"");if(!/^[0-9]+$/.test(_qn))return 0;const _r=e=>{const en=String(e.entry_num);return en===_qn?0:en.indexOf(_qn)===0?1:en.includes(_qn)?2:3;};return _r(a)-_r(b);}).slice(0,6):[];
  const close=()=>{setOpen(false);setQ("");};
  const pickEntity=(id)=>{onSelectEntity(id);onGo("dashboard");close();};
  const pickAccount=(code)=>{if(onPickAccount)onPickAccount(code);else onGo("coa");close();};
  const pickJE=(id)=>{if(onPickJE)onPickJE(id);else onGo("journal");close();};
  const Row=({children,onClick})=>(<div onClick={onClick} style={{padding:"8px 14px",cursor:"pointer",fontSize:13,borderRadius:6}} onMouseEnter={e=>e.currentTarget.style.background=T.bgElevated} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{children}</div>);
  const Hdr=({children})=>(<div style={{padding:"8px 14px 4px",fontSize:10,fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",color:T.textDim}}>{children}</div>);
  const hasAny=entHits.length||acctHits.length||jeHits.length;
  return(<div style={{position:"relative"}}>
    <input value={q} onFocus={()=>setOpen(true)} onChange={e=>{setQ(e.target.value);setOpen(true);}} placeholder="Search everything…"
      style={{width:240,padding:"7px 12px",borderRadius:T.radiusSm,border:"1px solid "+T.border,background:T.bgElevated,fontSize:13,color:T.textBright}}/>
    {open&&q&&<><div style={{position:"fixed",inset:0,zIndex:50}} onClick={close}/>
      <div style={{position:"absolute",top:"100%",left:0,marginTop:6,width:360,maxHeight:420,overflowY:"auto",background:"#fff",border:"1px solid "+T.border,borderRadius:T.radius,boxShadow:T.shadowLg,zIndex:100,padding:"6px 0"}}>
        {!hasAny&&<div style={{padding:"16px 14px",fontSize:13,color:T.textDim}}>{loaded?"No matches":"Searching…"}</div>}
        {entHits.length>0&&<><Hdr>Entities</Hdr>{entHits.map(e=><Row key={"e"+e.id} onClick={()=>pickEntity(e.id)}><span style={{fontWeight:600,color:T.textBright}}>{e.name}</span>{e.code&&<span style={{color:T.textDim,marginLeft:6,fontSize:11}}>{e.code}</span>}</Row>)}</>}
        {acctHits.length>0&&<><Hdr>Accounts (current entity)</Hdr>{acctHits.map(a=><Row key={"a"+a.code} onClick={()=>pickAccount(a.code)}><span style={{color:T.accent,fontWeight:600}}>{a.code}</span><span style={{marginLeft:8}}>{a.name}</span><span style={{...S.tag(a.type),marginLeft:8,transform:"scale(0.85)"}}>{a.type}</span></Row>)}</>}
        {jeHits.length>0&&<><Hdr>Journal Entries (current entity)</Hdr>{jeHits.map(e=><Row key={"j"+e.id} onClick={()=>pickJE(e.id)}><span style={{color:T.accent,fontWeight:600}}>JE-{String(e.entry_num).padStart(4,"0")}</span><span style={{color:T.textMuted,marginLeft:8}}>{e.date}</span><span style={{marginLeft:8}}>{e.memo}</span></Row>)}</>}
      </div></>}
  </div>);
}
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
  // Workpapers modal openable from the header (any page), for the active entity.
  const[wpEntity,setWpEntity]=useState(null);
  const[page,setPage]=useState('dashboard');const[loading,setLoading]=useState(true);
  // JE to auto-open after navigating to the Journal page (used by global search).
  const[pendingJEId,setPendingJEId]=useState(null);
  // Account code to pre-filter the CoA page (used by global search).
  const[pendingAcctCode,setPendingAcctCode]=useState(null);
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
    styleEl.textContent = '.cl-modal-box::before{content:"";position:absolute;top:0;left:0;right:56px;height:44px;cursor:move;border-top-left-radius:14px;border-top-right-radius:14px;z-index:1;}'
      + '.cl-colresize th{position:relative;}'
      + '.cl-colresize th::after{content:"";position:absolute;top:0;right:0;width:7px;height:100%;cursor:col-resize;}';
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
    // Column resizing: drag the right edge of a header cell in a .cl-colresize
    // table to set that column's width.
    const onColDown = (e) => {
      const th = e.target.closest && e.target.closest('.cl-colresize th');
      if (!th) return;
      const rect = th.getBoundingClientRect();
      if (rect.right - e.clientX > 8) return; // only when grabbing the right edge
      e.preventDefault(); e.stopPropagation();
      const startX = e.clientX, startW = rect.width;
      document.body.style.cursor = 'col-resize';
      const onMove = (ev) => { const w = Math.max(40, startW + ev.clientX - startX); th.style.width = th.style.minWidth = th.style.maxWidth = w + 'px'; };
      const onUp = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); document.body.style.cursor = ''; };
      document.addEventListener('mousemove', onMove); document.addEventListener('mouseup', onUp);
    };
    document.addEventListener('mousedown', onColDown, true);
    document.addEventListener('mousedown', onDown);
    return () => { document.removeEventListener('mousedown', onColDown, true); document.removeEventListener('mousedown', onDown); if (styleEl.parentNode) styleEl.parentNode.removeChild(styleEl); };
  }, []);
  const[showJE,setShowJE]=useState(false);const[showChangePw,setShowChangePw]=useState(false);const[rk,setRk]=useState(0);const[pendingReportConfig,setPendingReportConfig]=useState(null);
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
  const canAccess=s=>{if(!user)return false;if(user.role==='Admin')return true;return({Accountant:['entries','reports','coa','bankrec','billcom','workpapers'],Viewer:['entries','reports','coa','bankrec','workpapers']}[user.role]||[]).includes(s);};
  // Read-only users (Viewer) SEE the same sections as an Accountant but cannot edit.
  // canEdit gates every write control; it must never be derived from mere visibility.
  const canEdit = !!user && (user.role==='Admin' || user.role==='Accountant');
  if(loading)return<div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:T.textMuted}}>Loading...</div></div>;
  const _resetToken=(()=>{try{return new URLSearchParams(window.location.search).get('reset_token');}catch{return null;}})();
  if(_resetToken)return<ResetPasswordScreen token={_resetToken}/>;
  if(!user)return<AuthScreen onLogin={setUser}/>;
  const jeHasContent=jeForm.memo||jeForm.lines.some(l=>l.account_code||l.debit||l.credit)||jePendingFiles.length>0;

  const _activeEnt = entities.find(e=>e.id===activeEntity);
  _activeEntityCode = _activeEnt ? _activeEnt.code : null; // drives per-entity Class relabel (classTerm)
  const isTurnkeyEntity = !!(_activeEnt && (_activeEnt.code==='TURNKEYR' || /turnkey\s*rail/i.test(_activeEnt.name||'')));
  const isDevEntity = !!(_activeEnt && _activeEnt.entity_type==='development');
  // County Line Rail Fund — the only entity with the Management Fee workpaper for now.
  const isCLRF = !!(_activeEnt && (_activeEnt.code==='CLRF' || /county\s*line\s*rail\s*fund/i.test(_activeEnt.name||'')));
  const isShellEntity = !!(_activeEnt && _activeEnt.entity_type==='shell');
  const dimsEnabled = !!_activeEnt && !isShellEntity;// location/class dimensions available on every entity EXCEPT shell
  const arEnabled = !!_activeEnt && !isShellEntity;// AR / customer invoicing available on every entity EXCEPT shell
  const navItems=[
    {id:'dashboard',label:'Dashboard',icon:NI.dashboard,section:'reports'},
    {id:'d1',divider:1,label:'TRANSACTIONS'},{id:'journal',label:'Journal Entries',icon:NI.journal,section:'entries'},
    {id:'d2',divider:1,label:'ACCOUNTS'},{id:'coa',label:'Chart of Accounts',icon:NI.coa,section:'coa'},...(dimsEnabled?[{id:'dimensions',label:'Dimensions',icon:'🏷️',section:'coa'}]:[]),{id:'ledger',label:'General Ledger',icon:NI.ledger,section:'reports'},
    {id:'d2b',divider:1,label:'BANKING'},{id:'banktxn',label:'Bank Transactions',icon:NI.banktxn,section:'bankrec'},{id:'bankrec',label:'Bank Reconciliation',icon:NI.bankrec,section:'bankrec'},
    {id:'d3',divider:1,label:'REPORTS'},{id:'wp_finstmts',label:'Financial Statements',icon:'📑',section:'reports'},{id:'ttm',label:'Trailing 12 Months',icon:'📈',section:'reports'},{id:'fundrep',label:'Fund Reporting',icon:'🏦',section:'reports'},{id:'trial',label:'Trial Balance',icon:NI.trial,section:'reports'},{id:'bs',label:'Balance Sheet',icon:NI.bs,section:'reports'},{id:'is',label:'Income Statement',icon:NI.is,section:'reports'},
    {id:'customdetail',label:'Custom Detail',icon:'📋',section:'reports'},...(dimsEnabled?[{id:'pivot',label:'Pivot Summary',icon:'📊',section:'reports'}]:[]),{id:'apaging',label:'AP Aging',icon:'⏳',section:'reports'},{id:'commitments',label:'Commitments',icon:'🤝',section:'reports'},{id:'memorized',label:'Memorized Reports',icon:'★',section:'reports'},
    ...(isTurnkeyEntity?[{id:'wip',label:'WIP Schedule',icon:NI.wip,section:'reports'}]:[]),
    ...(isDevEntity?[{id:'d3b',divider:1,label:'DEVELOPMENT'},{id:'requisitions',label:'Requisitions',icon:'🏗️',section:'reports'}]:[]),
    ...(arEnabled?[{id:'d3c',divider:1,label:'RECEIVABLES'},{id:'ar_customers',label:'Customers',icon:'👥',section:'coa'}]:[]),
    ...(isCLRF?[{id:'dwp',divider:1,label:'WORKPAPERS'},{id:'wp_mgmtfee',label:'Management Fee',icon:'📄',section:'workpapers'}]:[]),
    {id:'d4',divider:1,label:'ADMIN'},{id:'entities',label:'Entities ('+entities.length+')',icon:NI.entities,section:'all'},{id:'users',label:'Users',icon:NI.users,section:'all'},
    {id:'d5',divider:1,label:'INTEGRATIONS'},{id:'billcom',label:'Bill.com Setup',icon:'💳',section:'billcom'},
  ];

  return(<div style={S.app}>
    <div style={S.topBar}><div style={{display:'flex',alignItems:'center',gap:16}}>
      <button style={{...S.btnGhost,fontSize:18,padding:'4px 6px',color:T.textMuted}} onClick={()=>setSidebarCol(c=>!c)}>{sidebarCol?'\u2630':'\u2190'}</button>
      <div style={{display:'flex',alignItems:'center',gap:10}}><Logo size={32}/>{!sidebarCol&&<div style={{fontSize:17,fontWeight:800,color:T.textBright}}>CloudLedger</div>}</div>
      <div style={{width:1,height:28,background:T.border}}/>{_activeEnt&&<button style={{...S.btnGhost,fontSize:18,padding:'4px 8px',lineHeight:1}} title={'Open '+_activeEnt.name+' Workpapers'} onClick={()=>setWpEntity(_activeEnt)}>📁</button>}<EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')}/>{entities.length>0&&<GlobalSearch entities={entities} activeEntity={activeEntity} onSelectEntity={setActiveEntity} onGo={setPage} onPickJE={(id)=>{setPendingJEId(id);setPage('journal');}} onPickAccount={(code)=>{setPendingAcctCode(code);setPage('coa');}}/>}</div>
      <div style={{display:'flex',alignItems:'center',gap:10}}>
        {canEdit&&activeEntity&&<button style={{...S.btnP,position:'relative'}} onClick={()=>setShowJE(true)}>+ Journal Entry{jeHasContent&&<span style={{position:'absolute',top:-3,right:-3,width:8,height:8,borderRadius:4,background:T.orange,border:'2px solid #fff'}}/>}</button>}
        <span style={{fontSize:13,fontWeight:500}}>{titleName(user.name)}</span><span style={S.badge}>{user.role}</span>
        <button style={S.btnS} onClick={()=>setShowChangePw(true)}>Settings</button>
        <button style={S.btnS} onClick={()=>{api.clearToken();setUser(null);}}>Sign Out</button></div></div>
    <div style={S.body}><div style={S.sidebar(sidebarCol)}>
      {navItems.map(n=>n.divider?(!sidebarCol?<div key={n.id} style={S.navSection(sidebarCol)}>{n.label}</div>:<div key={n.id} style={{height:8}}/>)
        :(n.section==='all'?user.role==='Admin':canAccess(n.section))?<div key={n.id} style={S.navItem(page===n.id,sidebarCol)} onClick={()=>setPage(n.id)} title={n.label}>
          {sidebarCol?<span style={{display:'inline-block',width:18,textAlign:'center',fontSize:15}}>{n.icon}</span>:<span><span style={{display:'inline-block',width:22,textAlign:'center',marginRight:8}}>{n.icon}</span>{n.label}</span>}</div>:null)}</div>
      <div style={S.main}>{(()=>{const en=entities.find(e=>e.id===activeEntity);const entityName=en?en.name:'';return<>
        {page==='dashboard'&&<Dashboard entityId={activeEntity} setActiveEntity={setActiveEntity} setPage={setPage} user={user} key={rk}/>}
        {page==='journal'&&activeEntity&&<JournalList entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} canEdit={canEdit} key={activeEntity+'-'+rk} onNewEntry={()=>setShowJE(true)} openJEId={pendingJEId} clearOpenJE={()=>setPendingJEId(null)}/>}
        {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} entityName={entityName} canEdit={canEdit}/>}
        {page==='dimensions'&&activeEntity&&dimsEnabled&&<DimensionsManager entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ar_customers'&&activeEntity&&arEnabled&&<CustomersManager entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} key={activeEntity+'-'+rk} from={glFrom} setFrom={setGlFrom} to={glTo} setTo={setGlTo} filter={glFilter} setFilter={setGlFilter}/>}
        {page==='banktxn'&&activeEntity&&<BankTransactions entityId={activeEntity} canEdit={canEdit} bankSelAcct={bankSelAcct} setBankSelAcct={setBankSelAcct} bankTxns={bankTxns} setBankTxns={setBankTxns} bankUploading={bankUploading} setBankUploading={setBankUploading} bankStatusFilter={bankStatusFilter} setBankStatusFilter={setBankStatusFilter}/>}
        {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user} canEdit={canEdit}/>}
        {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} isClrf={_activeEnt?.code==='COUNTYLI1'} key={activeEntity+'-'+rk} asOf={tbAsOf} setAsOf={setTbAsOf} canEdit={canEdit}/>}
        {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity} entityName={entityName} asOf={bsAsOf} setAsOf={setBsAsOf} canEdit={canEdit}/>}
        {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity} entityName={entityName} from={isFrom} setFrom={setIsFrom} to={isTo} setTo={setIsTo} canEdit={canEdit}/>}
        {page==='customdetail'&&activeEntity&&<CustomDetailReport entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='customdetail'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='pivot'&&activeEntity&&dimsEnabled&&<PivotReport entityId={activeEntity} entityName={entityName} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='pivot'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='apaging'&&activeEntity&&<ApAgingReport entityId={activeEntity} entityName={entityName} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='apaging'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='commitments'&&activeEntity&&<CommitmentsPage entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='memorized'&&activeEntity&&<MemorizedReportsPage entityId={activeEntity} entityName={entityName} canEdit={canEdit} onOpen={(r)=>{const c=r.config||{};if(r.report_type==='trial'&&c.asOf)setTbAsOf(c.asOf);else if(r.report_type==='bs'&&c.asOf)setBsAsOf(c.asOf);else if(r.report_type==='is'){if(c.from)setIsFrom(c.from);if(c.to)setIsTo(c.to);}else setPendingReportConfig({type:r.report_type,config:c});setPage(r.report_type==='drilldown'?'coa':r.report_type);}} key={activeEntity+'-'+rk}/>}
        {page==='wip'&&activeEntity&&<WipSchedule entityName={entityName} asOf={wipAsOf} setAsOf={setWipAsOf}/>}
        {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='users'&&<UserManagement currentUser={user}/>}
        {page==='billcom'&&<BillcomSetup entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='requisitions'&&activeEntity&&isDevEntity&&<Requisitions entityId={activeEntity} entityName={entityName} canEdit={canEdit} reqState={reqState} setReqState={setReqState}/>}
        {page==='wp_mgmtfee'&&activeEntity&&isCLRF&&<MgmtFeeWorkpaper entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='wp_finstmts'&&activeEntity&&<FinancialStatements entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ttm'&&activeEntity&&<TrailingTwelveMonths entityId={activeEntity} entityName={entityName} key={activeEntity+'-'+rk}/>}
        {page==='fundrep'&&activeEntity&&<FundReporting entityId={activeEntity} entityName={entityName} key={activeEntity+'-fr-'+rk}/>}
      </>})()}</div></div>
    {showJE&&activeEntity&&<JournalEntryModal entityId={activeEntity} isTurnkeyEntity={isTurnkeyEntity} dimsEnabled={dimsEnabled} user={user} onClose={()=>setShowJE(false)} onPosted={()=>setRk(k=>k+1)} form={jeForm} setForm={setJeForm} pendingFiles={jePendingFiles} setPendingFiles={setJePendingFiles}/>}
    {showChangePw&&<SettingsModal onClose={()=>setShowChangePw(false)} user={user} onUserUpdate={u=>setUser(u)}/>}
    {wpEntity&&<WorkpapersModal entity={wpEntity} user={user} onClose={()=>setWpEntity(null)}/>}
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
  const[defaultClearingAcct,setDefaultClearingAcct]=useState('');

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

  const pushCoaToBillcom=async(scope='all')=>{
    if(!selectedEntity)return;
    const isAll=scope==='all';
    const label=isAll?'all CloudLedger accounts (every type)':'CloudLedger Expense accounts';
    if(!window.confirm('Push '+label+' to Bill.com and auto-create mappings? Accounts already in Bill.com (by number or name) will be skipped.'))return;
    setMapPushing(true);setMapMsg('');setMapErr('');
    try{
      const r=await api.pushBillcomCoa(selectedEntity, isAll?{all:true}:{all_expenses:true});
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
        setDefaultClearingAcct(r.default_clearing_account||'');
        setPassword('');setDevKey('');
      }else{
        setEnv('sandbox');setUsername('');setOrgId('');setDefaultApAcct('');setDefaultCashAcct('');setDefaultClearingAcct('');
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
      const body={environment:env,username,org_id:orgId,default_ap_account:defaultApAcct||null,default_cash_account:defaultCashAcct||null,default_clearing_account:defaultClearingAcct||null};
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
          <div>
            <label style={S.label}>Default Clearing Account (Money Out Clearing)</label>
            <input type="text" value={defaultClearingAcct} onChange={e=>setDefaultClearingAcct(e.target.value)} style={S.input} placeholder="e.g. 10072" autoComplete="new-password"/>
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
            <button style={S.btnS} onClick={()=>pushCoaToBillcom('all')} disabled={mapPushing||mapLoading||mapSaving} title="Create every CloudLedger account (all types) in Bill.com and auto-map them">{mapPushing?'Pushing...':'Push ALL CL accounts to Bill.com'}</button>
            <button style={S.btnS} onClick={()=>pushCoaToBillcom('expenses')} disabled={mapPushing||mapLoading||mapSaving} title="Create only CloudLedger Expense accounts in Bill.com and auto-map them">{mapPushing?'Pushing...':'Push Expenses only'}</button>
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

        {syncResult&&syncResult.payments&&syncResult.payments.skip_reason&&<div style={{padding:12,marginBottom:14,background:'#fffbeb',border:'1px solid #fcd34d',borderRadius:T.radiusSm}}>
          <div style={{fontSize:13,fontWeight:600,color:'#92400e',marginBottom:6}}>Payments not synced</div>
          <div style={{fontSize:12,color:'#92400e'}}>{syncResult.payments.note||syncResult.payments.skip_reason}</div>
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
function Dashboard({entityId,setActiveEntity,setPage,user}){const[summary,setSummary]=useState([]);useEffect(()=>{api.getSummary().then(setSummary);},[]);
  const[wpEntity,setWpEntity]=useState(null);
  const[open,setOpen]=useState({accounting:false,development:false,shell:false});
  const go=id=>{setActiveEntity(id);setPage('journal');};
  const grouped=groupByType(summary);
  const toggle=k=>setOpen(o=>({...o,[k]:!o[k]}));
  const colSpan=6;
  return(<div><div style={S.h1}>Dashboard</div><div style={S.sub}>{summary.length} entities under management &middot; grouped by type &middot; click a type to expand, a row to open, the folder for workpapers</div>
    <div style={S.cardFlush}><table style={S.table}><thead><tr><th style={{...S.th,width:40}}></th><th style={S.th}>Entity</th><th style={S.thR}>Assets</th><th style={S.thR}>Liabilities</th><th style={S.thR}>Net Income</th><th style={S.thR}>JEs</th></tr></thead>
      <tbody>
        {ENTITY_TYPES.map(t=>{const rows=(grouped[t.key]||[]).slice().sort((a,b)=>a.name.localeCompare(b.name));
          const agg=rows.reduce((s,e)=>({a:s.a+(e.assets||0),l:s.l+(e.liabilities||0),ni:s.ni+(e.net_income||0),je:s.je+(e.entry_count||0)}),{a:0,l:0,ni:0,je:0});
          const isOpen=open[t.key];
          return(<Fragment key={t.key}>
            <tr style={{cursor:'pointer',background:T.bgElevated,borderTop:'2px solid '+T.border}} onClick={()=>toggle(t.key)}>
              <td style={{...S.td,textAlign:'center',fontSize:12,color:T.textMuted}}>{isOpen?'▾':'▸'}</td>
              <td style={{...S.td,fontWeight:700,color:T.textBright}}><span style={{marginRight:8}}>{t.icon}</span>{t.label}<span style={{marginLeft:8,fontSize:11,fontWeight:600,color:T.textMuted}}>({rows.length})</span></td>
              <td style={{...S.tdR,color:T.textMuted,fontWeight:600}}>{fmt(agg.a)}</td>
              <td style={{...S.tdR,color:T.textMuted,fontWeight:600}}>{fmt(agg.l)}</td>
              <td style={{...S.tdR,color:T.textMuted,fontWeight:600}}>{fmt(agg.ni)}</td>
              <td style={{...S.tdR,color:T.textMuted,fontWeight:600}}>{agg.je}</td>
            </tr>
            {isOpen&&rows.length===0&&<tr><td colSpan={colSpan} style={{...S.td,color:T.textMuted,padding:'10px 20px 10px 48px',fontSize:12}}>No {t.label.toLowerCase()} entities.</td></tr>}
            {isOpen&&rows.map(e=><tr key={e.id} style={{cursor:'pointer',background:e.id===entityId?T.accentDim:'transparent',transition:'background 0.1s'}} onMouseEnter={ev=>{if(e.id!==entityId)ev.currentTarget.style.background=T.bgHover;}} onMouseLeave={ev=>{if(e.id!==entityId)ev.currentTarget.style.background='transparent';}}>
              <td style={{...S.td,textAlign:'center',padding:'8px 6px'}} onClick={ev=>{ev.stopPropagation();setWpEntity(e);}} title="Open workpapers folder"><span style={{fontSize:18,cursor:'pointer',display:'inline-block',lineHeight:1}}>📁</span></td>
              <td style={{...S.td,fontWeight:600,color:T.accent,textDecoration:'underline',paddingLeft:32}} onClick={()=>go(e.id)}>{e.display_id&&<span style={{marginRight:8,fontSize:11,fontWeight:700,color:T.textMuted,fontFamily:'monospace'}}>{e.display_id}</span>}{e.name}</td>
              <td style={S.tdR} onClick={()=>go(e.id)}>{fmt(e.assets)}</td>
              <td style={S.tdR} onClick={()=>go(e.id)}>{fmt(e.liabilities)}</td>
              <td style={{...S.tdR,color:e.net_income>=0?T.green:T.red,fontWeight:600}} onClick={()=>go(e.id)}>{fmt(e.net_income)}</td>
              <td style={S.tdR} onClick={()=>go(e.id)}>{e.entry_count}</td>
            </tr>)}
          </Fragment>);})}
      </tbody></table></div>
    {wpEntity&&<WorkpapersModal entity={wpEntity} user={user} onClose={()=>setWpEntity(null)}/>}
  </div>);}

// ═══ Edit JE Modal ═══
function EditJEModal({entityId,dimsEnabled,entry,accounts:initAccounts,onClose,onSaved}){
  const[accounts,setAccounts]=useState(initAccounts||[]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[saving,setSaving]=useState(false);
  const[projects,setProjects]=useState([]);const[dimProjects,setDimProjects]=useState([]);
  useEffect(()=>{api.getTurnkeyProjects().then(setProjects).catch(()=>setProjects([]));api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>setDimProjects([]));},[entityId]);
  // Turnkey projects take precedence; otherwise use the project dimension (always shown for non-Turnkey when dims are on).
  const isTurnkeyHere=projects.length>0;
  const useDimProjects=!isTurnkeyHere&&dimsEnabled;
  const showProject=isTurnkeyHere||useDimProjects||(entry.lines||[]).some(l=>l.project_id);
  const addProjectInline=async(i)=>{
    const name=(prompt('New project name or code (e.g. P-10100.001):')||'').trim();
    if(!name) return;
    try{ const p=await api.createProject(entityId,{name,code:name});
      const list=await api.getProjects(entityId); setDimProjects(list||[]);
      updateLine(i,'project_id',p.id);
    }catch(e){ alert(e.message); }
  };
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);
  useEffect(()=>{api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  const showLocation=(dimsEnabled&&locations.length>0)||(entry.lines||[]).some(l=>l.location_id);const showClass=(dimsEnabled&&classes.length>0)||(entry.lines||[]).some(l=>l.class_id);
  const[form,setForm]=useState({date:entry.date,memo:entry.memo,lines:(entry.lines||[]).map(l=>({account_code:l.account_code,project_id:l.project_id||'',location_id:l.location_id||'',class_id:l.class_id||'',description:l.description||'',debit:l.debit>0?l.debit.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):'',credit:l.credit>0?l.credit.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):''}))});
  const[attachments,setAttachments]=useState(entry.attachments||[]);
  const[attUploading,setAttUploading]=useState(false);
  const attInputRef=useRef(null);
  useEffect(()=>{if(!initAccounts?.length)api.getAccounts(entityId).then(setAccounts);},[entityId,initAccounts]);
  const addLine=()=>setForm(f=>({...f,lines:[...f.lines,{account_code:'',debit:'',credit:'',description:''}]}));
  const removeLine=i=>setForm(f=>({...f,lines:f.lines.filter((_,j)=>j!==i)}));
  const updateLine=(i,k,v)=>setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,[k]:v}:l)}));
  // Single "Dimensions" dropdown per line — pick exactly one of Project/Location/Class.
  const showDims=showProject||showLocation||showClass;
  const projOpts=useDimProjects
    ?dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)}))
    :projects.map(pr=>({v:'project:'+pr.turnkey_project_id,label:'Project — '+pr.project_code+' — '+pr.project_name}));
  const locOpts=locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name}));
  const clsOpts=classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name}));
  const lineDimValue=l=>l.project_id?'project:'+l.project_id:l.location_id?'location:'+l.location_id:l.class_id?'class:'+l.class_id:'';
  const setLineDim=(i,val)=>{
    if(val==='__new__'){addProjectInline(i);return;}
    const[kind,id]=val?val.split(':'):['',''];
    setForm(f=>({...f,lines:f.lines.map((l,j)=>j===i?{...l,
      project_id:kind==='project'?id:'',
      location_id:kind==='location'?id:'',
      class_id:kind==='class'?id:''}:l)}));
  };
  const tDr=form.lines.reduce((s,l)=>s+parseAmt(l.debit),0);const tCr=form.lines.reduce((s,l)=>s+parseAmt(l.credit),0);const bal=Math.abs(tDr-tCr)<0.005&&tDr>0;
  const save=async()=>{if(!form.date||!form.memo.trim()){setErr('Date and memo required');return;}if(form.lines.some(l=>!l.account_code)){setErr('All lines need an account');return;}if(!bal){setErr('Must balance');return;}
    setSaving(true);setErr('');try{await api.updateEntry(entityId,entry.id,{date:form.date,memo:form.memo.trim(),lines:form.lines.map(l=>({account_code:l.account_code,debit:parseAmt(l.debit),credit:parseAmt(l.credit),description:l.description||'',project_id:l.project_id||null,location_id:l.location_id||null,class_id:l.class_id||null}))});
      onSaved();onClose();}catch(e){setErr(e.message);}finally{setSaving(false);}};
  const del=async()=>{if(!confirm('Delete JE-'+String(entry.entry_num).padStart(4,'0')+'? This permanently removes the entry and all its lines. This cannot be undone.'))return;
    setSaving(true);setErr('');try{await api.deleteEntry(entityId,entry.id);onSaved();onClose();}catch(e){setErr(e.message);setSaving(false);}};
  const uploadAtt=async e=>{const fl=e.target.files;if(!fl||fl.length===0)return;setErr('');setAttUploading(true);
    try{const r=await api.uploadAttachments(entityId,entry.id,fl);setAttachments(p=>[...p,...(r.attachments||r.files||r||[])]);}
    catch(ex){setErr(ex.message);}finally{setAttUploading(false);if(attInputRef.current)attInputRef.current.value='';}};
  const deleteAtt=async a=>{if(!confirm('Delete '+a.original_name+'?'))return;try{await api.deleteAttachment(a.id);setAttachments(p=>p.filter(x=>x.id!==a.id));}catch(ex){setErr(ex.message);}};
  const fmtPst=ts=>ts?new Date(ts+(ts.includes('Z')||ts.includes('+')?'':'Z')).toLocaleString('en-US',{timeZone:'America/Los_Angeles',year:'numeric',month:'short',day:'numeric',hour:'numeric',minute:'2-digit',hour12:true,timeZoneName:'short'}):'';
  // Only close on a genuine backdrop click (press AND release on the overlay) —
  // otherwise releasing a resize/drag onto the backdrop would close the window.
  const jeObDown=useRef(false);
  return(<div style={S.modal} onMouseDown={e=>{jeObDown.current=(e.target===e.currentTarget);}} onClick={e=>{if(jeObDown.current&&e.target===e.currentTarget)onClose();}}><div className="cl-modal-box" style={{...S.modalBox,width:'min(1200px, 96vw)',maxWidth:'96vw',height:'auto',maxHeight:'92vh',resize:'both',overflow:'auto',minWidth:'min(560px, 96vw)',minHeight:360}} onClick={e=>e.stopPropagation()}>
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
    <div style={{...S.cardFlush,marginBottom:16,maxHeight:'52vh',overflowY:'auto'}}><table className="cl-colresize" style={S.table}><thead style={{position:'sticky',top:0,zIndex:2,background:T.bgElevated}}><tr><th style={{...S.th,minWidth:300}}>Account</th>{showDims&&<th style={{...S.th,width:140}}>Dimension</th>}<th style={S.th}>Description</th><th style={{...S.thR,width:140}}>Debit</th><th style={{...S.thR,width:140}}>Credit</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{form.lines.map((l,i)=><tr key={i}><td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}>
        <select style={S.select} title={l.account_code?acctLabel(l.account_code,(accounts.find(a=>a.code===l.account_code)||{}).name||''):''} value={l.account_code} onChange={e=>updateLine(i,'account_code',e.target.value)}><option value="">Select...</option>
          {accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code} title={acctLabel(a.code,a.name)}>{acctLabel(a.code,a.name)}</option>)}</select></td>
        {showDims&&<td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><select style={S.select} value={lineDimValue(l)} onChange={e=>setLineDim(i,e.target.value)}><option value="">— none —</option>{showProject&&<optgroup label="Project">{projOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}{useDimProjects&&<option value="__new__">+ New project…</option>}</optgroup>}{showLocation&&<optgroup label="Location">{locOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</optgroup>}{showClass&&<optgroup label={classTerm()}>{clsOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</optgroup>}</select></td>}
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={S.input} value={l.description||''} placeholder="(optional)" onChange={e=>updateLine(i,'description',e.target.value)}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.debit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'debit',f);}} onBlur={e=>updateLine(i,'debit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px 8px',borderBottom:'1px solid '+T.borderLight}}><input style={{...S.input,textAlign:'right'}} value={l.credit} onChange={e=>{const f=fmtAmt(e.target.value);if(f!==null)updateLine(i,'credit',f);}} onBlur={e=>updateLine(i,'credit',blurAmt(e.target.value))}/></td>
        <td style={{padding:'6px',borderBottom:'1px solid '+T.borderLight,textAlign:'center'}}>{form.lines.length>2&&<button style={S.btnGhost} onClick={()=>removeLine(i)}>&times;</button>}</td></tr>)}
      <tr style={{background:T.bgElevated}}><td colSpan={2+(showDims?1:0)} style={{...S.tdBold,textAlign:'right',fontSize:12}}>TOTAL</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tDr)}</td><td style={{...S.tdBold,textAlign:'right',fontSize:15}}>${fmt(tCr)}</td><td style={S.tdBold}></td></tr></tbody></table></div>
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
      <button style={{...S.btnS,color:T.red,borderColor:T.red+'40'}} onClick={del} disabled={saving} title="Permanently delete this journal entry">Delete JE</button>
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
function JournalList({entityId,entityName,dimsEnabled,canEdit=true,onNewEntry,openJEId,clearOpenJE}){const[entries,setEntries]=useState([]);const[accounts,setAccounts]=useState([]);const[from,setFrom]=useState('');const[to,setTo]=useState('');const[q,setQ]=useState('');
  const[editEntry,setEditEntry]=useState(null);const[showBulk,setShowBulk]=useState(false);
  const[colW,setColW]=useState(()=>{try{const s=JSON.parse(localStorage.getItem('cl_je_colw'));if(s&&typeof s.acct==='number')return s;}catch(e){}return{acct:380,desc:300,debit:140,credit:140};});
  useEffect(()=>{try{localStorage.setItem('cl_je_colw',JSON.stringify(colW));}catch(e){}},[colW]);
  const startColDrag=(key,ev)=>{ev.preventDefault();ev.stopPropagation();const sx=ev.clientX;const sw=colW[key];const min=key==='acct'?140:key==='desc'?120:80;document.body.style.userSelect='none';const mv=e=>setColW(p=>({...p,[key]:Math.max(min,sw+(e.clientX-sx))}));const up=()=>{document.body.style.userSelect='';window.removeEventListener('mousemove',mv);window.removeEventListener('mouseup',up);};window.addEventListener('mousemove',mv);window.addEventListener('mouseup',up);};
  const grip=key=>(<span onMouseDown={e=>startColDrag(key,e)} title="Drag to resize column" style={{position:'absolute',top:0,right:0,width:8,height:'100%',cursor:'col-resize',userSelect:'none',borderRight:'2px solid transparent'}} onMouseEnter={e=>e.currentTarget.style.borderRight='2px solid '+T.accent} onMouseLeave={e=>e.currentTarget.style.borderRight='2px solid transparent'}/>);
  const load=useCallback(async()=>{const[e,a]=await Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]);setEntries(e);setAccounts(a);},[entityId,from,to]);
  useEffect(()=>{load();},[load]);
  // When arriving from global search with a target JE, open it once entries load.
  useEffect(()=>{if(openJEId&&entries.length){const hit=entries.find(e=>e.id===openJEId);if(hit){setEditEntry(hit);}if(clearOpenJE)clearOpenJE();}},[openJEId,entries]);const del=async id=>{if(!confirm('Delete this journal entry?'))return;await api.deleteEntry(entityId,id);load();};const acctName=code=>accounts.find(a=>a.code===code)?.name||'?';
  const shown=entries.filter(e=>{const t=q.trim().toLowerCase();if(!t)return true;const enStr=String(e.entry_num);const jeNum='je-'+enStr.padStart(4,'0');const qn=t.replace(/^je[-\s]?/,'').replace(/^0+(?=\d)/,'');if(jeNum.includes(t)||enStr.includes(t)||(/^\d+$/.test(qn)&&(enStr===qn||enStr.includes(qn))))return true;if((e.date||'').toLowerCase().includes(t))return true;if((e.memo||'').toLowerCase().includes(t))return true;const amtQ=t.replace(/[$,\s]/g,'');const amtMatch=amtQ&&/^[0-9.]+$/.test(amtQ)&&(e.lines||[]).some(l=>{const d=Number(l.debit||0),c=Number(l.credit||0);return String(d).includes(amtQ)||String(c).includes(amtQ)||d.toFixed(2).includes(amtQ)||c.toFixed(2).includes(amtQ);});if(amtMatch)return true;return (e.lines||[]).some(l=>(l.account_code||'').toLowerCase().includes(t)||(acctName(l.account_code)||'').toLowerCase().includes(t)||(l.description||'').toLowerCase().includes(t)||(l.class_name||'').toLowerCase().includes(t)||(l.location_name||'').toLowerCase().includes(t)||(l.project_name||'').toLowerCase().includes(t));}).sort((a,b)=>{const _t=q.trim().toLowerCase();const _qn=_t.replace(/^je[- ]?/,'').replace(/^0+/,'');if(!/^[0-9]+$/.test(_qn))return 0;const _r=e=>{const en=String(e.entry_num);return en===_qn?0:en.indexOf(_qn)===0?1:en.includes(_qn)?2:3;};return _r(a)-_r(b);});
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}><div><div style={S.h1}>Journal Entries</div><div style={S.sub}>{entityName} &middot; {q?shown.length+' of '+entries.length:entries.length} entries{!canEdit&&' · read-only'}</div></div>{canEdit&&<div style={{display:'flex',gap:8}}><button style={S.btnS} onClick={()=>setShowBulk(true)}>Bulk Upload</button><button style={S.btnP} onClick={onNewEntry}>+ New Entry</button></div>}</div>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <div style={{flex:1,minWidth:200}}><label style={S.label}>Search</label><input style={{...S.inputSm,width:'100%'}} placeholder="JE#, memo, date, account, amount, description..." value={q} onChange={e=>setQ(e.target.value)}/></div>
      {(from||to||q)&&<button style={{...S.btnGhost,marginTop:14,color:T.red}} onClick={()=>{setFrom('');setTo('');setQ('');}}>Clear</button>}</div>
    {shown.length===0?<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No entries found</div>:
      <div style={{display:'flex',flexDirection:'column',gap:8}}>{shown.map(e=><div key={e.id} style={{...S.card,padding:14,marginBottom:0}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
          <div style={{display:'flex',alignItems:'center',gap:14}}><span style={{fontWeight:700,color:T.accent,fontSize:14}}>JE-{String(e.entry_num).padStart(4,'0')}</span>
            <span style={{color:T.textMuted}}>{e.date}</span><span style={{fontWeight:500}}>{e.memo}</span>
            {e.attachments?.length>0&&<span style={{fontSize:11,color:T.teal,fontWeight:500}}>({e.attachments.length} file{e.attachments.length>1?'s':''})</span>}</div>
          <div style={{display:'flex',alignItems:'center',gap:8}}>
            <span style={{fontSize:12,color:T.textDim}}>{e.created_by}</span>
            {canEdit&&<button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>setEditEntry(e)}>Edit</button>}
            {canEdit&&<button style={{...S.btnD,padding:'5px 12px',fontSize:11}} onClick={()=>del(e.id)}>Delete</button>}</div></div>
        <div style={{overflowX:'auto'}}><table style={{...S.table,tableLayout:'fixed',width:colW.acct+colW.desc+colW.debit+colW.credit}}>
          <colgroup><col style={{width:colW.acct}}/><col style={{width:colW.desc}}/><col style={{width:colW.debit}}/><col style={{width:colW.credit}}/></colgroup>
          <thead><tr><th style={{...S.th,position:'relative'}}>Account{grip('acct')}</th><th style={{...S.th,position:'relative'}}>Description{grip('desc')}</th><th style={{...S.thR,position:'relative'}}>Debit{grip('debit')}</th><th style={{...S.thR,position:'relative'}}>Credit{grip('credit')}</th></tr></thead>
          <tbody>{e.lines.map((l,i)=><tr key={i}><td style={S.td} title={acctLabel(l.account_code,acctName(l.account_code))}>{acctLabel(l.account_code,acctName(l.account_code))}</td>
            <td style={{...S.td,color:T.textMuted,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={[l.description||'',l.project_name?('Project: '+(l.project_code&&l.project_code!==l.project_name?l.project_code:l.project_name)):'',l.location_name?('Location: '+l.location_name):'',l.class_name?('Class: '+l.class_name):''].filter(Boolean).join('  ·  ')}>{l.description||''}{(l.project_name||l.location_name||l.class_name)&&<span style={{marginLeft:l.description?8:0,fontSize:10,color:T.accent}}>{[l.project_name?('▦ '+(l.project_code&&l.project_code!==l.project_name?l.project_code:l.project_name)):'',l.location_name,l.class_name].filter(Boolean).join(' · ')}</span>}</td>
            <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td><td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td></tr>)}</tbody></table></div>
        {e.attachments?.length>0&&<div style={{marginTop:10,display:'flex',flexWrap:'wrap',gap:4}}>{e.attachments.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>)}</div>}
      </div>)}</div>}
    {editEntry&&<EditJEModal entityId={entityId} dimsEnabled={dimsEnabled} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={load}/>}
    {showBulk&&<BulkJEModal entityId={entityId} onClose={()=>setShowBulk(false)} onPosted={()=>{setShowBulk(false);load();}}/>}
  </div>);}

// ═══ Dimensions (Locations & Classes) manager ═══
function DimList({title,subtitle,items,canEdit,onCreate,onUpdate,onDelete,onBulkUpload}){
  const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({code:'',name:''});const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState({code:'',name:''});const[editErr,setEditErr]=useState('');
  const[uploading,setUploading]=useState(false);
  const fileRef=useRef(null);
  const startEdit=it=>{setEditing(it.id);setEditForm({code:it.code||'',name:it.name||''});setEditErr('');};
  const add=async()=>{if(!form.name.trim()){setErr('Name required');return;}try{await onCreate({name:form.name.trim(),code:form.code.trim()||null});setForm({code:'',name:''});setShowAdd(false);setErr('');}catch(e){setErr(e.message);}};
  const save=async()=>{if(!editForm.name.trim()){setEditErr('Name required');return;}try{await onUpdate(editing,{name:editForm.name.trim(),code:editForm.code.trim()||null});setEditing(null);}catch(e){setEditErr(e.message);}};
  const del=async it=>{if(!confirm('Delete "'+it.name+'"?'))return;try{await onDelete(it.id);}catch(e){alert(e.message);}};
  // Parse an uploaded xlsx/csv into [{code,name}] rows: looks for "code"/"name"
  // header columns (case-insensitive); falls back to first two columns.
  const onFile=async e=>{const file=e.target.files&&e.target.files[0];e.target.value='';if(!file||!onBulkUpload)return;
    setUploading(true);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}).filter(r=>r&&r.some(c=>String(c).trim()));
      if(!rows.length){alert('No rows found in the file.');return;}
      const hdr=rows[0].map(c=>String(c).trim().toLowerCase());
      let ci=hdr.findIndex(h=>/^(project )?code$/.test(h)); let ni=hdr.findIndex(h=>/name|description/.test(h));
      let body;
      if(ci>=0&&ni>=0){body=rows.slice(1);}else{ci=0;ni=1;body=rows.slice(/code|name|project/i.test(rows[0].join(' '))?1:0);}
      const projects=body.map(r=>({code:String(r[ci]==null?'':r[ci]).trim(),name:String(r[ni]==null?'':r[ni]).trim()})).filter(r=>r.code&&r.name);
      if(!projects.length){alert('Could not find Code and Name columns in the file.');return;}
      await onBulkUpload(projects);
    }catch(ex){alert('Import failed: '+ex.message);}finally{setUploading(false);}
  };
  return(<div style={{flex:1,minWidth:340}}>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
      <div><div style={{fontSize:15,fontWeight:700,color:T.textBright}}>{title}</div><div style={{fontSize:12,color:T.textMuted}}>{subtitle||(items.length+' total')}</div></div>
      <div style={{display:'flex',gap:8}}>
        {canEdit&&onBulkUpload&&<><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:'none'}} onChange={onFile}/>
          <button style={{...S.btnS,padding:'6px 12px',fontSize:12}} disabled={uploading} onClick={()=>fileRef.current&&fileRef.current.click()}>{uploading?'Importing…':'Import code → name'}</button></>}
        {canEdit&&<button style={{...S.btnP,padding:'6px 12px',fontSize:12}} onClick={()=>{setShowAdd(!showAdd);setErr('');}}>{showAdd?'Cancel':'+ Add'}</button>}</div></div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40',padding:14,marginBottom:12}}><div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
      {onBulkUpload&&<div style={{width:120}}><label style={S.label}>Code</label><input style={S.input} value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')add();}}/></div>}
      <div style={{flex:1}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')add();}}/></div>
      <button style={S.btnP} onClick={add}>Add</button></div>{err&&<div style={{...S.err,marginTop:8,marginBottom:0}}>{err}</div>}</div>}
    <div style={S.cardFlush}><table style={{...S.table,tableLayout:'fixed'}}><thead><tr><th style={S.th}>Name</th>{canEdit&&<th style={{...S.th,width:84}}>Actions</th>}</tr></thead>
      <tbody>{items.length===0&&<tr><td colSpan={canEdit?2:1} style={{...S.td,color:T.textMuted,textAlign:'center',padding:'18px'}}>None yet</td></tr>}
      {items.map(it=>editing===it.id?
        <tr key={it.id} style={{background:T.accentDim}}>
          <td style={{padding:'6px 8px'}}><div style={{display:'flex',gap:6}}>{onBulkUpload&&<input style={{...S.input,width:110}} placeholder="Code" value={editForm.code} onChange={e=>setEditForm(f=>({...f,code:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')save();}}/>}<input style={S.input} placeholder="Name" value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))} onKeyDown={e=>{if(e.key==='Enter')save();}}/></div></td>
          {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnGhost,color:T.green,fontSize:11}} onClick={save}>Save</button><button style={{...S.btnGhost,fontSize:11}} onClick={()=>setEditing(null)}>Cancel</button></div></td>}
        </tr>
        :<tr key={it.id}>
          <td style={S.td} title={it.code&&it.code!==it.name?(it.code+' — '+it.name):it.name}>{it.code&&it.code!==it.name?<><span style={{color:T.textMuted,fontFamily:'monospace',fontSize:12}}>{it.code}</span>{'  '}{it.name}</>:it.name}</td>
          {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
            <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(it)}>Edit</button>
            <button style={{...S.btnGhost,color:it.line_count>0?T.textMuted:T.red,fontSize:11}} title={it.line_count>0?'Used on '+it.line_count+' line(s) — cannot delete':'Delete'} onClick={()=>del(it)}>x</button></div></td>}
        </tr>)}
      </tbody></table></div>
    {editErr&&<div style={{...S.err,marginTop:8}}>{editErr}</div>}</div>);
}
function DimensionsManager({entityId,entityName,canEdit}){
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);const[projects,setProjects]=useState([]);
  const load=useCallback(async()=>{const[l,c,p]=await Promise.all([api.getLocations(entityId),api.getClasses(entityId),api.getProjects(entityId)]);setLocations(l||[]);setClasses(c||[]);setProjects(p||[]);},[entityId]);
  useEffect(()=>{load();},[load]);
  return(<div><div style={{marginBottom:20}}><div style={S.h1}>Dimensions</div><div style={S.sub}>{entityName} — dimensions you can tag on journal-entry lines and filter reports by</div></div>
    <div style={{display:'flex',gap:24,flexWrap:'wrap',alignItems:'flex-start'}}>
      <DimList title="Locations" subtitle={(locations.length)+' location'+(locations.length===1?'':'s')+' (deals / properties)'} items={locations} canEdit={canEdit}
        onCreate={async d=>{await api.createLocation(entityId,d);await load();}}
        onUpdate={async(id,d)=>{await api.updateLocation(entityId,id,d);await load();}}
        onDelete={async id=>{await api.deleteLocation(entityId,id);await load();}}/>
      <DimList title={classTerm()==='Class'?'Investor Classes':classTerm()+'s'} subtitle={classTerm()==='Class'?((classes.length)+' class'+(classes.length===1?'':'es')+' (investors / capital classes)'):((classes.length)+' '+(classes.length===1?classTerm().toLowerCase():classTerm().toLowerCase()+'s'))} items={classes} canEdit={canEdit}
        onCreate={async d=>{await api.createClass(entityId,d);await load();}}
        onUpdate={async(id,d)=>{await api.updateClass(entityId,id,d);await load();}}
        onDelete={async id=>{await api.deleteClass(entityId,id);await load();}}/>
      <DimList title="Projects" subtitle={(projects.length)+' project'+(projects.length===1?'':'s')+' (Intacct project / QBO class)'} items={projects} canEdit={canEdit}
        onCreate={async d=>{await api.createProject(entityId,d);await load();}}
        onUpdate={async(id,d)=>{await api.updateProject(entityId,id,d);await load();}}
        onDelete={async id=>{await api.deleteProject(entityId,id);await load();}}
        onBulkUpload={async projects=>{
          const applyAll=confirm('Import '+projects.length+' project codes.\n\nOK = apply to ALL accounting & development entities (except County Line Rail Fund).\nCancel = this entity only.\n\nExisting codes are updated with the new name; new codes are added. Nothing is deleted.');
          const r=await api.bulkProjects(entityId,projects,applyAll);
          await load();
          alert('Done. '+r.entities+' entit'+(r.entities===1?'y':'ies')+' updated · '+r.created+' created · '+r.updated+' renamed · '+r.skipped+' unchanged'+(r.failed?(' · '+r.failed+' failed'):'')+'.');
        }}/>
    </div></div>);
}

// ═══ AR Customers manager ═══
function CustomersManager({entityId,entityName,canEdit}){
  const[customers,setCustomers]=useState([]);const[loading,setLoading]=useState(true);
  const[showAdd,setShowAdd]=useState(false);
  const blank={name:'',email:'',address:'',terms_days:30};
  const[form,setForm]=useState(blank);const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState(blank);const[editErr,setEditErr]=useState('');
  const load=useCallback(async()=>{setLoading(true);try{const c=await api.getArCustomers(entityId);setCustomers(c||[]);}finally{setLoading(false);}},[entityId]);
  useEffect(()=>{load();},[load]);
  const add=async()=>{if(!form.name.trim()){setErr('Name required');return;}try{await api.createArCustomer(entityId,{name:form.name.trim(),email:form.email.trim()||null,address:form.address.trim()||null,terms_days:+form.terms_days||30});setForm(blank);setShowAdd(false);setErr('');await load();}catch(e){setErr(e.message);}};
  const startEdit=c=>{setEditing(c.id);setEditForm({name:c.name||'',email:c.email||'',address:c.address||'',terms_days:c.terms_days??30});setEditErr('');};
  const save=async()=>{if(!editForm.name.trim()){setEditErr('Name required');return;}try{await api.updateArCustomer(entityId,editing,{name:editForm.name.trim(),email:editForm.email.trim()||null,address:editForm.address.trim()||null,terms_days:+editForm.terms_days||30});setEditing(null);await load();}catch(e){setEditErr(e.message);}};
  const toggleActive=async c=>{try{await api.updateArCustomer(entityId,c.id,{active:c.active?0:1});await load();}catch(e){alert(e.message);}};
  const del=async c=>{if(!confirm('Delete "'+c.name+'"? (Only possible if no invoices exist.)'))return;try{await api.deleteArCustomer(entityId,c.id);await load();}catch(e){alert(e.message);}};
  return(<div><div style={{marginBottom:20,display:'flex',justifyContent:'space-between',alignItems:'flex-start'}}>
    <div><div style={S.h1}>Customers</div><div style={S.sub}>{entityName} — bill-to parties for the invoices you send. Email here is where invoices go.</div></div>
    {canEdit&&<button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setErr('');}}>{showAdd?'Cancel':'+ Add Customer'}</button>}</div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40',padding:16,marginBottom:16}}>
      <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 220px'}}><label style={S.label}>Customer name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} autoFocus/></div>
        <div style={{flex:'1 1 220px'}}><label style={S.label}>Email (invoice recipient)</label><input style={S.input} type="email" placeholder="ar@customer.com" value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></div>
        <div style={{flex:'0 0 120px'}}><label style={S.label}>Terms (days)</label><input style={S.input} type="number" value={form.terms_days} onChange={e=>setForm(f=>({...f,terms_days:e.target.value}))}/></div>
      </div>
      <div style={{marginTop:10}}><label style={S.label}>Billing address (optional)</label><textarea style={{...S.input,minHeight:54,resize:'vertical'}} value={form.address} onChange={e=>setForm(f=>({...f,address:e.target.value}))}/></div>
      <div style={{marginTop:12,display:'flex',justifyContent:'flex-end'}}><button style={S.btnP} onClick={add}>Add Customer</button></div>
      {err&&<div style={{...S.err,marginTop:8,marginBottom:0}}>{err}</div>}</div>}
    <div style={S.cardFlush}><table style={S.table}><thead><tr>
      <th style={S.th}>Name</th><th style={S.th}>Email</th><th style={{...S.th,width:80,textAlign:'right'}}>Terms</th><th style={{...S.th,width:90}}>Status</th>{canEdit&&<th style={{...S.th,width:150}}>Actions</th>}</tr></thead>
      <tbody>
        {loading&&<tr><td colSpan={canEdit?5:4} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>Loading…</td></tr>}
        {!loading&&customers.length===0&&<tr><td colSpan={canEdit?5:4} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>No customers yet — add one to start invoicing.</td></tr>}
        {!loading&&customers.map(c=>editing===c.id?
          <tr key={c.id} style={{background:T.accentDim}}>
            <td style={{padding:'6px 8px'}}><input style={S.input} value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))}/></td>
            <td style={{padding:'6px 8px'}}><input style={S.input} value={editForm.email} onChange={e=>setEditForm(f=>({...f,email:e.target.value}))}/></td>
            <td style={{padding:'6px 8px'}}><input style={{...S.input,textAlign:'right'}} type="number" value={editForm.terms_days} onChange={e=>setEditForm(f=>({...f,terms_days:e.target.value}))}/></td>
            <td style={S.td}>—</td>
            {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnGhost,color:T.green,fontSize:11}} onClick={save}>Save</button><button style={{...S.btnGhost,fontSize:11}} onClick={()=>setEditing(null)}>Cancel</button></div></td>}
          </tr>
          :<tr key={c.id} style={c.active?undefined:{opacity:0.5}}>
            <td style={S.td} title={c.address||''}>{c.name}</td>
            <td style={S.td}>{c.email||<span style={{color:T.textMuted}}>— no email —</span>}</td>
            <td style={{...S.td,textAlign:'right'}}>Net {c.terms_days}</td>
            <td style={S.td}>{c.active?<span style={{color:T.green,fontSize:12}}>Active</span>:<span style={{color:T.textMuted,fontSize:12}}>Inactive</span>}</td>
            {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
              <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(c)}>Edit</button>
              <button style={{...S.btnGhost,fontSize:11}} onClick={()=>toggleActive(c)}>{c.active?'Deactivate':'Reactivate'}</button>
              <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={()=>del(c)}>x</button></div></td>}
          </tr>)}
      </tbody></table></div>
    {editErr&&<div style={{...S.err,marginTop:8}}>{editErr}</div>}
    <div style={{marginTop:14,fontSize:12,color:T.textMuted}}>Next: recurring invoice templates and one-click send are coming in the next update. For now this is where you manage who you bill.</div>
  </div>);
}

// ═══ Chart of Accounts ═══
function ChartOfAccounts({entityId,entityName,canEdit}){const[accounts,setAccounts]=useState([]);const[showAdd,setShowAdd]=useState(false);const[q,setQ]=useState('');
  const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState({});const[editErr,setEditErr]=useState('');
  const[balByCode,setBalByCode]=useState({});const[drillAcct,setDrillAcct]=useState(null);
  const asOf=today();
  const yearAgo=(()=>{const d=new Date();d.setFullYear(d.getFullYear()-1);return d.toISOString().slice(0,10);})();
  const load=useCallback(async()=>{
    const[accts,bals]=await Promise.all([api.getAccounts(entityId),api.getBalances(entityId,{as_of:asOf}).catch(()=>[])]);
    setAccounts(accts);const m={};(bals||[]).forEach(b=>{m[b.code]=b.balance;});setBalByCode(m);
  },[entityId,asOf]);useEffect(()=>{load();},[load]);
  const editPanelRef=useRef(null);
  // The edit form renders at the top of the page; when Edit is clicked on a row
  // far down a long chart of accounts, the form would open off-screen and look
  // like nothing happened. Scroll it into view so it's always visible.
  const startEdit=a=>{setEditing(a.code);setEditForm({new_code:a.code,name:a.name,type:a.type,subtype:a.subtype||'',bank_acct:!!a.bank_acct});setEditErr('');
    setTimeout(()=>{editPanelRef.current&&editPanelRef.current.scrollIntoView({behavior:'smooth',block:'center'});},50);};
  const saveEdit=async()=>{if(!editForm.new_code||!editForm.name){setEditErr('Code and name required');return;}
    try{await api.updateAccount(entityId,editing,editForm);setEditing(null);load();}catch(e){setEditErr(e.message);}};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:20}}><div><div style={S.h1}>Chart of Accounts</div><div style={S.sub}>{accounts.length} accounts</div></div>
    <div style={{display:'flex',alignItems:'center',gap:10}}>
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search code, name, or type..." style={{...S.inputSm,width:260,padding:'8px 12px'}}/>
      {canEdit&&<button style={S.btnP} onClick={()=>setShowAdd(!showAdd)}>{showAdd?'Cancel':'+ Add Account'}</button>}
    </div></div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}><div style={S.row}>
      <div style={S.col}><label style={S.label}>Code</label><input style={S.input} placeholder="e.g. 61500" value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))}/></div>
      <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
      <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
      <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={form.bank_acct} onChange={e=>setForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={async()=>{if(!form.code||!form.name){setErr('Required');return;}try{await api.createAccount(entityId,form);setForm({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});setShowAdd(false);setErr('');load();}catch(e){setErr(e.message);}}}>Add Account</button></div>}
    {editing&&<div ref={editPanelRef} style={{...S.card,borderColor:T.accent+'40',marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:12}}>Edit Account: {editing}</div>
      <div style={S.row}>
        <div style={S.col}><label style={S.label}>Account Code</label><input style={S.input} value={editForm.new_code} onChange={e=>setEditForm(f=>({...f,new_code:e.target.value}))}/></div>
        <div style={{...S.col,flex:2}}><label style={S.label}>Name</label><input style={S.input} value={editForm.name} onChange={e=>setEditForm(f=>({...f,name:e.target.value}))}/></div>
        <div style={S.col}><label style={S.label}>Type</label><select style={S.select} value={editForm.type} onChange={e=>setEditForm(f=>({...f,type:e.target.value}))}>{['Asset','Liability','Equity','Revenue','Expense'].map(t=><option key={t}>{t}</option>)}</select></div></div>
      <div style={{marginBottom:14}}><label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={editForm.bank_acct} onChange={e=>setEditForm(f=>({...f,bank_acct:e.target.checked}))}/>Bank / cash account</label></div>
      {editErr&&<div style={S.err}>{editErr}</div>}
      {editForm.new_code!==editing&&<div style={{fontSize:11,color:T.orange,marginBottom:8}}>Changing code from {editing} to {editForm.new_code} will update all journal entries, bank transactions, and reconciliations.</div>}
      <div style={{display:'flex',gap:10}}><button style={S.btnP} onClick={saveEdit}>Save Changes</button><button style={S.btnS} onClick={()=>setEditing(null)}>Cancel</button></div></div>}
    <div style={S.cardFlush}><table style={S.table}><thead><tr><th style={S.th}>Code</th><th style={S.th}>Name</th><th style={S.th}>Type</th><th style={S.thC}>Bank</th><th style={S.thR}>Balance (as of {asOf})</th>{canEdit&&<th style={{...S.th,width:80}}>Actions</th>}</tr></thead>
      <tbody>{accounts.filter(a=>{const t=q.trim().toLowerCase();if(!t)return true;return (a.code||'').toLowerCase().includes(t)||(a.name||'').toLowerCase().includes(t)||(a.type||'').toLowerCase().includes(t);}).map(a=><tr key={a.code} style={editing===a.code?{background:T.accentDim}:{cursor:'pointer'}} onClick={e=>{if(e.target.closest('button'))return;setDrillAcct({code:a.code,name:a.name,type:a.type,balance:balByCode[a.code]||0});}}>
        <td style={{...S.td,color:T.textBright}}>{a.code}</td><td style={S.td}>{a.name}</td><td style={S.td}><span style={S.tag(a.type)}>{a.type}</span></td>
        <td style={S.tdC}>{a.bank_acct?<span style={{color:T.green}}>Yes</span>:''}</td>
        <td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(balByCode[a.code]||0)}</td>
        {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}>
          <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(a)}>Edit</button>
          <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={async()=>{try{await api.deleteAccount(entityId,a.code);load();}catch(e){alert(e.message);}}}>x</button></div></td>}</tr>)}</tbody></table></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={yearAgo} to={asOf} onClose={()=>setDrillAcct(null)} onChanged={load}/>}
    </div>);}

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
function BankMatchModal({txn, entityId, onClose, onMatched}){
  const [cands, setCands] = useState(null);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);
  const [sel, setSel] = useState(null);
  useEffect(() => { (async () => {
    try { const r = await api.getBankMatchCandidates(entityId, txn.id); setCands(r.candidates || []); }
    catch (e) { setErr(e.message); setCands([]); }
  })(); }, [entityId, txn.id]);
  const confirm = async () => {
    if (!sel) { setErr('Select a journal entry to match'); return; }
    setSaving(true); setErr('');
    try { await api.matchBankTransaction(entityId, txn.id, sel); onMatched(); }
    catch (e) { setErr(e.message); } finally { setSaving(false); }
  };
  return (<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox, maxWidth: 760}} onClick={e => e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Match to Existing Journal Entry</div>
    <div style={{fontSize:12,color:T.textMuted,marginBottom:16}}>{txn.date} &middot; {txn.description}</div>
    <div style={{...S.card,background:T.bgElevated,padding:12,marginBottom:14,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
      <div><div style={{fontSize:11,color:T.textMuted,fontWeight:600,textTransform:'uppercase',letterSpacing:0.4}}>Bank Amount</div>
        <div style={{fontSize:22,fontWeight:700,color:txn.amount>=0?T.green:T.red,marginTop:2}}>{txn.amount>=0?'+':'-'}${fmt(Math.abs(txn.amount))}</div></div>
      <div style={{textAlign:'right',fontSize:11,color:T.textMuted}}>Showing posted JEs that hit this bank account<br/>for {fmt(Math.abs(txn.amount))} within &plusmn;7 days</div>
    </div>
    {cands === null && <div style={{textAlign:'center',padding:40,color:T.textDim}}>Finding candidates&hellip;</div>}
    {cands !== null && cands.length === 0 && <div style={{textAlign:'center',padding:40,color:T.textDim}}>No matching journal entries found within &plusmn;7 days.<br/>Code this transaction to an account instead, or post it to create a new JE.</div>}
    {cands !== null && cands.length > 0 && <table style={{...S.table,marginBottom:14}}>
      <thead><tr><th style={{...S.th,width:36}}></th><th style={S.th}>JE #</th><th style={S.th}>Date</th><th style={S.th}>Memo</th><th style={S.thR}>Amount</th><th style={{...S.thR,width:90}}>Day diff</th></tr></thead>
      <tbody>{cands.map(c => { const cid = c.je_id; return <tr key={cid} style={{cursor:'pointer',background:sel===cid?T.tealDim:'transparent'}} onClick={()=>setSel(cid)}>
        <td style={{...S.td,textAlign:'center'}}><input type="radio" checked={sel===cid} onChange={()=>setSel(cid)}/></td>
        <td style={{...S.td,fontWeight:600,color:T.teal}}>#{c.entry_num||cid}</td>
        <td style={{...S.td,color:T.textMuted,fontSize:12}}>{c.date}</td>
        <td style={{...S.td}} title={c.memo}>{c.memo}</td>
        <td style={{...S.tdR,fontFamily:'monospace'}}>${fmt(Math.abs(c.bank_net))}</td>
        <td style={{...S.tdR,color:T.textMuted}}>{c.date_diff!=null?Math.abs(c.date_diff)+'d':''}</td>
      </tr>; })}</tbody>
    </table>}
    {err && <div style={{...S.err,marginBottom:12}}>{err}</div>}
    <div style={{display:'flex',gap:10,justifyContent:'flex-end'}}>
      <button style={S.btnS} onClick={onClose} disabled={saving}>Cancel</button>
      <button style={{...S.btnP,opacity:(!sel||saving)?0.5:1}} onClick={confirm} disabled={!sel||saving}>{saving?'Matching...':'Confirm Match'}</button>
    </div>
  </div></div>);
}

function SplitBankTransactionModal({txn, accounts, excludeCode, entityId, onClose, onSaved}){
  const target = Math.abs(txn.amount);
  const initialLines = (txn.splits && txn.splits.length > 0)
    ? txn.splits.map(s => ({ account_code: s.account_code, amount: String(s.amount), memo: s.memo || '', project_id: s.project_id||null, class_id: s.class_id||null, location_id: s.location_id||null }))
    : (txn.account_code
        ? [{ account_code: txn.account_code, amount: target.toFixed(2), memo: txn.memo || '', project_id: txn.project_id||null, class_id: txn.class_id||null, location_id: txn.location_id||null }, { account_code: '', amount: '', memo: '' }]
        : [{ account_code: '', amount: '', memo: '' }, { account_code: '', amount: '', memo: '' }]);
  const [lines, setLines] = useState(initialLines);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);
  const [locations, setLocations] = useState([]); const [classes, setClasses] = useState([]); const [dimProjects, setDimProjects] = useState([]);
  useEffect(() => { api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>{}); api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>{}); api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>{}); }, [entityId]);
  const dimOpts = [
    ...dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)})),
    ...locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name})),
    ...classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name})),
  ];
  const showDims = dimOpts.length > 0;
  const lineDimValue = l => l.project_id?'project:'+l.project_id:l.location_id?'location:'+l.location_id:l.class_id?'class:'+l.class_id:'';
  const setLineDim = (i, val) => { const [kind,id] = val?val.split(':'):['','']; setLines(prev=>prev.map((l,idx)=>idx===i?{...l,project_id:kind==='project'?id:null,class_id:kind==='class'?id:null,location_id:kind==='location'?id:null}:l)); };

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
      await api.splitBankTransaction(entityId, txn.id, valid.map(l => ({ account_code: l.account_code, amount: parseAmt(l.amount), memo: l.memo || null, project_id: l.project_id||null, class_id: l.class_id||null, location_id: l.location_id||null })));
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
      <thead><tr><th style={S.th}>GL Account</th>{showDims&&<th style={{...S.th,width:170}}>Dimension</th>}<th style={{...S.thR,width:140}}>Amount</th><th style={{...S.th,width:180}}>Memo</th><th style={{...S.th,width:36}}></th></tr></thead>
      <tbody>{lines.map((l, i) => <tr key={i}>
        <td style={{...S.td,padding:'4px 6px'}}><AccountAutocomplete accounts={accounts} value={l.account_code} exclude={excludeCode} onChange={v => updateLine(i, 'account_code', v)} placeholder="Search GL account..."/></td>
        {showDims&&<td style={{...S.td,padding:'4px 6px'}}><select style={{...S.inputSm,width:'100%'}} value={lineDimValue(l)} onChange={e=>setLineDim(i,e.target.value)}><option value="">No dimension</option>{dimOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</select></td>}
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
  const[matchTxn,setMatchTxn]=useState(null);
  // Dimensions (Location / Class / Project) available to tag when coding a txn.
  const[locations,setLocations]=useState([]);const[classes,setClasses]=useState([]);const[dimProjects,setDimProjects]=useState([]);
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
  useEffect(()=>{api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>setDimProjects([]));},[entityId]);
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
  const codeTransaction=async(id,acct_code,memo,dims)=>{const cur=txns.find(t=>t.id===id)||{};
    const d={project_id:dims&&'project_id'in dims?dims.project_id:(cur.project_id||null),class_id:dims&&'class_id'in dims?dims.class_id:(cur.class_id||null),location_id:dims&&'location_id'in dims?dims.location_id:(cur.location_id||null)};
    await api.codeBankTransaction(entityId,id,acct_code,memo,d);
    setTxns(prev=>prev.map(t=>t.id===id?{...t,account_code:acct_code,memo:memo,...d,status:acct_code?'coded':'pending'}:t));};
  // One tagged dimension per transaction (Project / Location / Class), mirroring JEs.
  const projOpts=dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)}));
  const locOpts=locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name}));
  const clsOpts=classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name}));
  const dimOpts=[...projOpts,...locOpts,...clsOpts];const showDims=dimOpts.length>0;
  const txnDimValue=t=>t.project_id?'project:'+t.project_id:t.location_id?'location:'+t.location_id:t.class_id?'class:'+t.class_id:'';
  const setTxnDim=(t,val)=>{const[kind,id]=val?val.split(':'):['',''];codeTransaction(t.id,t.account_code||null,t.memo,{project_id:kind==='project'?id:null,class_id:kind==='class'?id:null,location_id:kind==='location'?id:null});};
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
            : (t.status==='matched'
                ? <div style={{display:'flex',gap:6,alignItems:'center'}}>
                    <span style={{flex:1,minWidth:0,fontSize:11,color:T.teal,fontWeight:600,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}} title={'Matched to JE #'+(t.matched_entry_id||t.je_id)}>&#9656; JE #{t.matched_entry_id||t.je_id}</span>
                    <button style={{...S.btnGhost,fontSize:10,color:T.red,padding:'4px 6px',whiteSpace:'nowrap'}} onClick={async()=>{try{await api.unmatchBankTransaction(entityId,t.id);reload();}catch(ex){setErr(ex.message);}}} title="Unlink from this journal entry">Unmatch</button>
                  </div>
                : t.splits && t.splits.length>0
                ? <button style={{...S.btnS,padding:'5px 10px',fontSize:11,color:T.purple,borderColor:T.purple+'40',width:'100%',textAlign:'left'}} onClick={()=>setSplitTxn(t)} title={t.splits.map(s=>s.account_code+' $'+fmt(s.amount)).join(' | ')}>Split: {t.splits.length} accts &middot; ${fmt(t.splits.reduce((s,x)=>s+x.amount,0))}</button>
                : <div>
                    <div style={{display:'flex',gap:4,alignItems:'center'}}>
                    <div style={{flex:1,minWidth:0}}><AccountAutocomplete accounts={accounts} value={t.account_code||''} exclude={selAcct} onChange={v=>codeTransaction(t.id,v,t.memo)} placeholder="Search GL account..."/></div>
                    <button style={{...S.btnGhost,fontSize:10,color:T.teal,padding:'4px 6px',whiteSpace:'nowrap'}} onClick={()=>setMatchTxn(t)} title="Match to an existing journal entry">Match</button>
                    <button style={{...S.btnGhost,fontSize:10,color:T.purple,padding:'4px 6px',whiteSpace:'nowrap'}} onClick={()=>setSplitTxn(t)} title="Split across multiple accounts">Split</button>
                    </div>
                    {showDims&&<select style={{...S.inputSm,marginTop:4,width:'100%'}} value={txnDimValue(t)} onChange={e=>setTxnDim(t,e.target.value)} title="Tag a Location / Class / Project dimension"><option value="">No dimension</option>{dimOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</select>}
                  </div>)}</td>
          <td style={{...S.td,padding:'4px 6px',overflow:'visible',borderRight:'1px solid '+T.borderLight}}>{(t.status==='posted'||!canEdit)?<span style={{fontSize:12,color:T.textDim}}>{t.memo}</span>:
            (t.splits && t.splits.length>0
              ? <span style={{fontSize:11,color:T.textDim,fontStyle:'italic'}}>(per split)</span>
              : <input style={S.inputSm} placeholder="Memo" value={t.memo||''} onChange={e=>{const v=e.target.value;setTxns(prev=>prev.map(x=>x.id===t.id?{...x,memo:v}:x));}} onBlur={()=>codeTransaction(t.id,t.account_code,t.memo)}/>)}</td>
          <td style={{...S.td,borderRight:'1px solid '+T.borderLight}}><span style={{fontSize:10,fontWeight:600,padding:'3px 8px',borderRadius:20,background:t.status==='posted'?T.greenDim:t.status==='matched'?T.tealDim:t.status==='coded'?T.accentDim:T.orangeDim,color:t.status==='posted'?T.green:t.status==='matched'?T.teal:t.status==='coded'?T.accent:T.orange}}>{t.status}</span></td>
          <td style={S.td}>{canEdit&&t.status!=='posted'&&<button style={S.btnGhost} onClick={async()=>{await api.deleteBankTransaction(entityId,t.id);setTxns(prev=>prev.filter(x=>x.id!==t.id));}}>x</button>}</td>
        </tr>)}</tbody></table></div>}
    {selAcct&&filteredTxns.length===0&&!uploading&&<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No transactions yet. Upload a bank statement above.</div>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>{setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));if(a.bank_acct)setBankAccts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));}}/>}
    {splitTxn&&<SplitBankTransactionModal txn={splitTxn} accounts={accounts} excludeCode={selAcct} entityId={entityId} onClose={()=>setSplitTxn(null)} onSaved={()=>{setSplitTxn(null);loadTxns(selAcct,statusFilter);}}/>}
    {matchTxn&&<BankMatchModal txn={matchTxn} entityId={entityId} onClose={()=>setMatchTxn(null)} onMatched={()=>{setMatchTxn(null);loadTxns(selAcct,statusFilter);}}/>}
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

// ══ Financial-report options (Liting #2): date presets, period columns, prior-period comparative ══
const RPT_DATE_FILTERS=[['all','All'],['month','Last Month'],['quarter','Last Quarter'],['year','Last Year']];
const RPT_COL_MODES=[['total','Total Only'],['monthly','Monthly'],['quarterly','Quarterly'],['yearly','Yearly']];
const _ymd=d=>new Date(d).toISOString().slice(0,10);
const _mkDate=s=>new Date((/^\d{4}-\d{2}-\d{2}$/.test(s)?s:today())+'T00:00:00');
// Overall [from,to] window for a date preset, anchored at `anchor` (YYYY-MM-DD).
function rptWindow(filter,anchor){
  const a=_mkDate(anchor);const to=_ymd(a);
  const back=(n,unit)=>{const s=new Date(a);if(unit==='m')s.setMonth(s.getMonth()-n);else s.setFullYear(s.getFullYear()-n);s.setDate(s.getDate()+1);return _ymd(s);};
  if(filter==='month')return{from:back(1,'m'),to};
  if(filter==='quarter')return{from:back(3,'m'),to};
  if(filter==='year')return{from:back(1,'y'),to};
  return{from:null,to};// 'all' = inception → anchor
}
// Split a window into calendar-aligned sub-periods. Returns [{label,from,to}].
function rptPeriods(filter,mode,anchor){
  const w=rptWindow(filter,anchor);
  if(mode==='total')return[{label:'Total',from:w.from,to:w.to}];
  const to=_mkDate(w.to);
  const cs=w.from?_mkDate(w.from):new Date(to.getFullYear(),0,1);
  cs.setDate(1);
  if(mode==='quarterly')cs.setMonth(Math.floor(cs.getMonth()/3)*3);
  if(mode==='yearly')cs.setMonth(0);
  const stepM=mode==='monthly'?1:mode==='quarterly'?3:12;
  const cols=[];let cur=new Date(cs);
  while(cur<=to&&cols.length<120){
    const segStart=new Date(cur);const next=new Date(segStart);next.setMonth(next.getMonth()+stepM);
    let segEnd=new Date(next);segEnd.setDate(segEnd.getDate()-1);if(segEnd>to)segEnd=new Date(to);
    let segFrom=segStart;if(w.from&&_mkDate(w.from)>segStart)segFrom=_mkDate(w.from);
    const label=mode==='monthly'?segStart.toLocaleString('en-US',{month:'short',year:'2-digit'})
      :mode==='quarterly'?('Q'+(Math.floor(segStart.getMonth()/3)+1)+" '"+String(segStart.getFullYear()).slice(2))
      :String(segStart.getFullYear());
    cols.push({label,from:_ymd(segFrom),to:_ymd(segEnd)});
    cur=next;
  }
  return cols;
}
// Immediately-preceding window of equal length (for the "previous period" comparative).
function rptPriorWindow(p){
  if(!p||!p.from)return null; // inception-based windows have no prior
  const from=_mkDate(p.from),to=_mkDate(p.to);
  const days=Math.round((to-from)/86400000)+1;
  const pTo=new Date(from);pTo.setDate(pTo.getDate()-1);
  const pFrom=new Date(pTo);pFrom.setDate(pFrom.getDate()-days+1);
  return{label:'Prev',from:_ymd(pFrom),to:_ymd(pTo)};
}
const rptPct=(cur,prev)=>Math.abs(prev)<0.005?null:(cur-prev)/Math.abs(prev)*100;
const rptChgCell=(cur,prev)=>{const d=cur-prev;const p=rptPct(cur,prev);return{d,p};};
function ReportControls({dateFilter,setDateFilter,colMode,setColMode,compare,setCompare,anchorLabel}){
  return(<>
    <div><label style={S.label}>Date range</label><select style={S.inputSm} value={dateFilter} onChange={e=>setDateFilter(e.target.value)}>{RPT_DATE_FILTERS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>
    <div><label style={S.label}>Columns</label><select style={S.inputSm} value={colMode} onChange={e=>setColMode(e.target.value)}>{RPT_COL_MODES.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>
    <div><label style={S.label}>Compare</label><label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,color:T.textMuted,height:34}} title="Adds a Previous Period column with $ and % change (Total Only view)"><input type="checkbox" checked={compare} onChange={e=>setCompare(e.target.checked)} disabled={colMode!=='total'}/> Prev period ($ / %)</label></div>
  </>);
}

function TrialBalance({entityId,entityName,dimsEnabled,isClrf,asOf,setAsOf,canEdit=true}){
  const[data,setData]=useState([]);
  const[dateFilter,setDateFilter]=useState('all');const[colMode,setColMode]=useState('total');const[compare,setCompare]=useState(false);
  const[rk,setRk]=useState(0);
  const[drillAcct,setDrillAcct]=useState(null);
  const[locations,setLocations]=useState([]);
  const[locId,setLocId]=useState('');// '' = all (whole-entity TB); otherwise a location_id
  const[classes,setClasses]=useState([]);
  const[classId,setClassId]=useState('');// '' = all investors; otherwise a class_id (investor)
  const[projects,setProjects]=useState([]);
  const[projId,setProjId]=useState('');// '' = all; otherwise a project_id
  // Filter rule: only County Line Rail Fund uses Location/Investor dimensions.
  // Every other entity filters by Project instead.
  const showLocInv=dimsEnabled&&isClrf;
  const showProj=dimsEnabled&&!isClrf;
  // Guard: while the user is editing the date input, asOf can briefly be '' or a partial string like '2026-'.
  // Avoid crashing the page on Invalid Date — fall back to today() until a complete YYYY-MM-DD is entered.
  const validAsOf=/^\d{4}-\d{2}-\d{2}$/.test(asOf)&&!isNaN(new Date(asOf+'T00:00:00').getTime())?asOf:today();
  const fyS=validAsOf.slice(0,4)+'-01-01';
  const locName=locId?(locations.find(l=>String(l.id)===String(locId))?.name||''):'';
  const className=classId?(classes.find(c=>String(c.id)===String(classId))?.name||''):'';
  const projName=projId?(()=>{const p=projects.find(p=>String(p.id)===String(projId));return p?(p.code&&p.code!==p.name?p.code:p.name):'';})():'';
  const dimmed=!!(locId||classId||projId); // any dimension selected → activity view
  const scopeLabel=[locName,className,projName].filter(Boolean).join(' · ');
  // 12-month window ending at asOf. If asOf is 2026-04-10, window is 2025-04-11 → 2026-04-10.
  const drillFrom=useMemo(()=>{const d=new Date(validAsOf+'T00:00:00');d.setFullYear(d.getFullYear()-1);d.setDate(d.getDate()+1);return d.toISOString().slice(0,10);},[validAsOf]);
  // Whole-entity TB uses the soft-close (close_pl_before) path. A dimension-scoped TB
  // (location and/or class/investor) is activity-based: it sums only lines carrying
  // the selected tag(s), so there is no period-close/RE roll — pass the dimension
  // id(s) with the as_of date only.
  const anchor=validAsOf;
  const periods=useMemo(()=>rptPeriods(dateFilter,colMode,anchor),[dateFilter,colMode,anchor]);
  const prior=(compare&&colMode==='total'&&periods[0]&&periods[0].from)?rptPriorWindow(periods[0]):null;
  const cols=useMemo(()=>prior?[prior,...periods]:periods,[JSON.stringify(prior),JSON.stringify(periods)]);
  const dimArgs=useMemo(()=>({...(locId?{location_id:locId}:{}),...(classId?{class_id:classId}:{}),...(projId?{project_id:projId}:{})}),[locId,classId,projId]);
  useEffect(()=>{let ok=true;Promise.all(cols.map(c=>api.getBalances(entityId,dimmed?{as_of:c.to,...dimArgs}:{as_of:c.to,close_pl_before:c.to.slice(0,4)+'-01-01'}).catch(()=>[]))).then(r=>{if(ok)setData(r);});return()=>{ok=false;};},[entityId,JSON.stringify(cols),JSON.stringify(dimArgs),dimmed,rk]);
  useEffect(()=>{api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>setLocations([]));},[entityId]);
  useEffect(()=>{api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>setClasses([]));},[entityId]);
  useEffect(()=>{api.getProjects(entityId).then(d=>setProjects(d||[])).catch(()=>setProjects([]));},[entityId]);
  const meta=useMemo(()=>{const m=new Map();data.forEach(bs=>(bs||[]).forEach(b=>{if(!m.has(b.code))m.set(b.code,{code:b.code,name:b.name,type:b.type});}));return[...m.values()].sort((a,b)=>String(a.code).localeCompare(String(b.code)));},[data]);
  const vmap=useMemo(()=>data.map(bs=>{const mm=new Map();(bs||[]).forEach(b=>mm.set(b.code,b.balance));return mm;}),[data]);
  const balAt=(code,ci)=>(vmap[ci]&&vmap[ci].get(code))||0;
  const typeOf=code=>(meta.find(m=>m.code===code)||{}).type;
  const drcr=(code,ci)=>{const b=balAt(code,ci);const isDr=typeOf(code)==='Asset'||typeOf(code)==='Expense';return{dr:(isDr&&b>0)||(!isDr&&b<0)?Math.abs(b):0,cr:(isDr&&b<0)||(!isDr&&b>0)?Math.abs(b):0};};
  const rows=meta.filter(m=>vmap.some(mm=>Math.abs((mm&&mm.get(m.code))||0)>0.005));
  const nCols=cols.length;const curI=nCols-1;const priI=prior?0:-1;
  const totDr=ci=>rows.reduce((s,r)=>s+drcr(r.code,ci).dr,0);const totCr=ci=>rows.reduce((s,r)=>s+drcr(r.code,ci).cr,0);
  const pctTxt=p=>p==null?'—':(p>=0?'+':'')+p.toFixed(1)+'%';
  const oneYrBefore=d=>{const x=new Date(d+'T00:00:00');x.setFullYear(x.getFullYear()-1);x.setDate(x.getDate()+1);return x.toISOString().slice(0,10);};
  const colHead=(c,i)=>prior&&i===0?'Prev':(c.label==='Total'?'':c.label);
  const fnameTag=[locName,className,projName].filter(Boolean).map(s=>s.replace(/[^A-Za-z0-9]+/g,'_')).join('_');
  const doExport=()=>{const lbl=scopeLabel?(' — '+scopeLabel):'';const hdr=['Code','Account','Type'];cols.forEach((c,i)=>{const h=colHead(c,i);hdr.push((h?h+' ':'')+'Debit',(h?h+' ':'')+'Credit');});const d=[[entityName||'Trial Balance'],['Trial Balance'+lbl],['As of '+anchor],[],hdr];
    rows.forEach(r=>{const row=[r.code,r.name,r.type];cols.forEach((c,i)=>{const x=drcr(r.code,i);row.push(x.dr||'',x.cr||'');});d.push(row);});
    const tot=['','','Total'];cols.forEach((c,i)=>{tot.push(totDr(i),totCr(i));});d.push([]);d.push(tot);
    exportToExcel(d,'TB'+(fnameTag?'_'+fnameTag:'')+'_'+anchor+'.xlsx');};
  const amtStyle={...S.tdR,cursor:'pointer'};
  // GL detail export (optionally scoped to the selected location and/or investor).
  // Pulls flat lines with running balance from /gl-detail through the as-of date;
  // dimension-tagged lines only when a dimension is selected.
  const doExportGL=async()=>{
    try{
      const r=await api.getGLDetail(entityId,{to:validAsOf,...(locId?{location_id:locId}:{}),...(classId?{class_id:classId}:{}),...(projId?{project_id:projId}:{})});
      const lbl=scopeLabel?(' — '+scopeLabel):'';
      const d=[[entityName||'General Ledger'],['GL Detail'+lbl],['Through '+asOf],[],['Date','Entry #','Account','Account Name','Memo / Description','Project','Location',classTerm(),'Debit','Credit','Running Bal']];
      (r.lines||[]).forEach(l=>d.push([l.date,l.entry_num,l.account_code,l.account_name,l.description||l.memo||'',l.project_code&&l.project_code!==l.project_name?l.project_code:(l.project_name||''),l.location_name,l.class_name,l.debit||'',l.credit||'',l.running_balance]));
      d.push([]);d.push(['','','','','','','','','Total Dr','Total Cr','']);
      d.push(['','','','','','','','',r.total_debit,r.total_credit,'']);
      exportToExcel(d,'GL'+(fnameTag?'_'+fnameTag:'')+'_'+asOf+'.xlsx');
    }catch(e){alert('GL export failed: '+e.message);}
  };
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
      <ReportControls dateFilter={dateFilter} setDateFilter={setDateFilter} colMode={colMode} setColMode={setColMode} compare={compare} setCompare={setCompare}/>
      {showProj&&<div><label style={S.label}>Project</label><select style={S.inputSm} value={projId} onChange={e=>setProjId(e.target.value)}><option value="">All (whole entity)</option>{projects.map(p=><option key={p.id} value={p.id}>{p.code&&p.code!==p.name?p.code+' — '+p.name:p.name}{p.line_count!=null?(' ('+p.line_count+')'):''}</option>)}</select></div>}
      {showLocInv&&<div><label style={S.label}>Location</label><select style={S.inputSm} value={locId} onChange={e=>setLocId(e.target.value)}><option value="">All (whole entity)</option>{locations.map(l=><option key={l.id} value={l.id}>{l.name}{l.line_count!=null?(' ('+l.line_count+')'):''}</option>)}</select></div>}
      {showLocInv&&<div><label style={S.label}>Investor (Class)</label><select style={S.inputSm} value={classId} onChange={e=>setClassId(e.target.value)}><option value="">All investors</option>{classes.map(c=><option key={c.id} value={c.id}>{c.name}{c.line_count!=null?(' ('+c.line_count+')'):''}</option>)}</select></div>}</div>
    <div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='trial' currentConfig={{asOf,dateFilter,colMode,compare}} onApply={(c)=>{if(c.asOf)setAsOf(c.asOf);if(c.dateFilter)setDateFilter(c.dateFilter);if(c.colMode)setColMode(c.colMode);if(typeof c.compare==='boolean')setCompare(c.compare);}} canEdit={canEdit}/><button style={S.btnExport} onClick={doExportGL} title="Export flat GL detail (dimension-tagged only when a location/investor is selected)">Export GL Detail</button><button style={S.btnExport} onClick={doExport}>Export TB</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Trial Balance{scopeLabel?(' — '+scopeLabel):''}</div><div style={{fontSize:13,color:T.textMuted}}>As of {asOf}{dimmed?' · dimension-tagged activity only':''}</div></div>
    <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:520}}>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th>
        {cols.map((c,i)=>{const h=colHead(c,i);return[<th key={'hd'+i} style={S.thR}>{(h?h+' ':'')}Debit</th>,<th key={'hc'+i} style={S.thR}>{(h?h+' ':'')}Credit</th>];})}
        {prior&&<><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}</tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={{...S.td,color:T.textBright}}>{r.code}</td><td style={S.td} title={r.name}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td>
        {cols.map((c,i)=>{const x=drcr(r.code,i);const clk=()=>setDrillAcct({...r,from:oneYrBefore(c.to),to:c.to});return[<td key={'d'+i} style={{...S.tdR,cursor:x.dr>0?'pointer':'default',color:x.dr>0?T.accent:undefined}} onClick={()=>x.dr>0&&clk()}>{x.dr>0?fmt(x.dr):''}</td>,<td key={'c'+i} style={{...S.tdR,cursor:x.cr>0?'pointer':'default',color:x.cr>0?T.accent:undefined}} onClick={()=>x.cr>0&&clk()}>{x.cr>0?fmt(x.cr):''}</td>];})}
        {prior&&(()=>{const cN=balAt(r.code,curI),pN=balAt(r.code,priI);return[<td key="dc" style={S.tdR}>{fmt(cN-pN)}</td>,<td key="pc" style={{...S.tdR,color:(cN-pN)>=0?T.green:T.red}}>{pctTxt(rptPct(cN,pN))}</td>];})()}</tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td>{cols.map((c,i)=>[<td key={'d'+i} style={{...S.tdBold,textAlign:'right'}}>${fmt(totDr(i))}</td>,<td key={'c'+i} style={{...S.tdBold,textAlign:'right'}}>${fmt(totCr(i))}</td>])}{prior&&<><td style={S.tdBold}/><td style={S.tdBold}/></>}</tr></tbody></table></div>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(totDr(curI)-totCr(curI))<0.005?T.green:T.red}}>{Math.abs(totDr(curI)-totCr(curI))<0.005?'In balance':'Off by $'+fmt(totDr(curI)-totCr(curI))}</div></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillAcct.from||drillFrom} to={drillAcct.to||asOf} onClose={()=>setDrillAcct(null)} onChanged={()=>setRk(k=>k+1)}/>}
  </div>);
}

// ═══ Account Drill-Down Modal (12-month GL detail from TB) ═══
function AccountDrillDownModal({entityId,entityName,acct,from:fromProp,to:toProp,onClose,onChanged}){
  const[reloadKey,setReloadKey]=useState(0);
  // Committed range that actually drives the query (item 2: only updates when the
  // user clicks Refresh, so the window doesn't reload on every typed digit).
  const[from,setFrom]=useState(fromProp);
  const[to,setTo]=useState(toProp);
  // Draft range bound to the date inputs; committed to from/to on Refresh (or Enter).
  const[fromDraft,setFromDraft]=useState(fromProp);
  const[toDraft,setToDraft]=useState(toProp);
  const applyDates=()=>{
    if(/^\d{4}-\d{2}-\d{2}$/.test(fromDraft))setFrom(fromDraft);
    if(/^\d{4}-\d{2}-\d{2}$/.test(toDraft))setTo(toDraft);
  };
  const datesDirty=fromDraft!==from||toDraft!==to;
  const[lines,setLines]=useState([]);
  const[begBal,setBegBal]=useState(0);
  const[loading,setLoading]=useState(true);
  const[err,setErr]=useState('');
  const[allEntries,setAllEntries]=useState([]);
  const[allAccounts,setAllAccounts]=useState([]);
  const[viewEntry,setViewEntry]=useState(null);
  // Beginning balance is the account balance as of the day before 'from', fetched
  // directly so the window can be any custom range (not just trailing-12mo).
  const prevDay=(d)=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return x.toISOString().slice(0,10);};
  useEffect(()=>{
    if(!/^\d{4}-\d{2}-\d{2}$/.test(from)||!/^\d{4}-\d{2}-\d{2}$/.test(to))return;
    (async()=>{
      setLoading(true);setErr('');
      try{
        const[entries,accts,begBalances]=await Promise.all([
          api.getEntries(entityId,from,to),
          api.getAccounts(entityId),
          api.getBalances(entityId,{as_of:prevDay(from)})
        ]);
        setAllEntries(entries);setAllAccounts(accts);
        const nameMap=Object.fromEntries((accts||[]).map(a=>[a.code,a.name]));
        // Offset / payee account for a line = the other account(s) on the same JE.
        // One distinct other account -> its label; more than one -> "-Split-".
        const offsetOf=(e)=>{const others=[...new Set((e.lines||[]).filter(x=>x.account_code!==acct.code).map(x=>x.account_code))];if(others.length===0)return '';if(others.length===1)return others[0]+' - '+(nameMap[others[0]]||'');return '-Split-';};
        const txns=[];
        entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code)txns.push({date:e.date,entry_num:e.entry_num,jeId:e.id,memo:e.memo,offset:offsetOf(e),vendor:e.vendor||'',debit:l.debit||0,credit:l.credit||0,created_by:e.created_by,created_at:e.created_at,class_name:l.class_name||'',location_name:l.location_name||''});});});
        txns.sort((a,b)=>a.date.localeCompare(b.date)||a.entry_num-b.entry_num);
        const bb=(begBalances||[]).find(x=>x.code===acct.code);
        setBegBal(bb?bb.balance:0);
        setLines(txns);
      }catch(e){setErr(e.message);}finally{setLoading(false);}
    })();
  },[entityId,acct.code,acct.type,from,to,reloadKey]);
  const openJE=jeId=>{const e=allEntries.find(x=>x.id===jeId);if(e)setViewEntry(e);};
  const isDr=acct.type==='Asset'||acct.type==='Expense';
  let running=begBal;
  const totalDr=lines.reduce((s,l)=>s+l.debit,0);
  const totalCr=lines.reduce((s,l)=>s+l.credit,0);
  const doExport=()=>{const acctLabel=acct.code+' - '+acct.name;
    const d=[[entityName||'Account Detail'],[acctLabel],['Period: '+from+' to '+to],[],['Date','JE','Account',classTerm(),'Location','Memo','Offset Account','Vendor/Payee','Debit','Credit','Balance']];
    d.push(['','','','','','Beginning Balance','','','','',begBal]);
    let r=begBal;lines.forEach(l=>{r+=isDr?(l.debit-l.credit):(l.credit-l.debit);d.push([l.date,'JE-'+String(l.entry_num).padStart(4,'0'),acctLabel,l.class_name||'',l.location_name||'',l.memo,l.offset||'',l.vendor||'',l.debit||'',l.credit||'',r]);});
    d.push(['','','','','','Totals','','',totalDr,totalCr,r]);
    exportToExcel(d,'GL_'+acct.code+'_'+to+'.xlsx');};
  return(<div style={S.modal}><div className="cl-modal-box" style={{...S.modalBox,width:'min(1100px,96vw)',maxWidth:'98vw',height:'88vh',maxHeight:'96vh',minWidth:'min(680px,96vw)',minHeight:420,resize:'both',overflow:'hidden',display:'flex',flexDirection:'column'}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14,gap:16}}>
      <div><div style={{fontSize:18,fontWeight:700,color:T.textBright}}>{acct.code} &mdash; {acct.name}</div>
        <div style={{display:'flex',alignItems:'center',gap:8,marginTop:6}}><label style={{fontSize:11,color:T.textMuted}}>From</label><input type="date" value={fromDraft} max={toDraft} onChange={e=>setFromDraft(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyDates();}} style={{...S.inputSm,padding:'4px 8px',fontSize:12}}/><label style={{fontSize:11,color:T.textMuted}}>To</label><input type="date" value={toDraft} min={fromDraft} onChange={e=>setToDraft(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyDates();}} style={{...S.inputSm,padding:'4px 8px',fontSize:12}}/><button onClick={applyDates} disabled={!datesDirty} style={{background:datesDirty?T.accent:T.bgElevated,border:'1px solid '+(datesDirty?T.accent:T.border),borderRadius:6,color:datesDirty?'#fff':T.textMuted,fontSize:11,fontWeight:600,padding:'4px 12px',cursor:datesDirty?'pointer':'default'}}>Refresh</button><button onClick={()=>{setFromDraft(fromProp);setToDraft(toProp);setFrom(fromProp);setTo(toProp);}} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'4px 8px',cursor:'pointer'}}>Reset</button>
        <span style={{width:1,height:16,background:T.border,margin:'0 2px'}}/>{PRESETS.map(([k,lbl])=><button key={k} onClick={()=>{const r=presetRange(k);setFromDraft(r.from);setToDraft(r.to);setFrom(r.from);setTo(r.to);}} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'4px 8px',cursor:'pointer'}}>{lbl}</button>)}</div>
        <div style={{marginTop:6}}><span style={S.tag(acct.type)}>{acct.type}</span></div></div>
      <button style={S.btnExport} onClick={doExport}>Export Excel</button>
    </div>
    {loading?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Loading...</div>:
     err?<div style={S.err}>{err}</div>:
     <div style={{flex:1,overflowY:'auto',border:'1px solid '+T.border,borderRadius:T.radiusSm}}>
       <table className="cl-colresize" style={S.table}>
         <thead style={{position:'sticky',top:0,background:T.bgCard,zIndex:1}}><tr>
           <th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.th}>Offset Account</th><th style={S.th}>Vendor/Payee</th>
           <th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th></tr></thead>
         <tbody>
           <tr style={{background:T.bgElevated}}>
             <td style={{...S.td,color:T.textMuted,fontStyle:'italic'}} colSpan={5}>Beginning balance as of {from}</td>
             <td style={S.tdR}></td><td style={S.tdR}></td>
             <td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(begBal)}</td></tr>
           {lines.length===0?
             <tr><td colSpan={8} style={{...S.td,textAlign:'center',padding:30,color:T.textDim}}>No activity in this period{from>'2015-01-01'&&<div style={{marginTop:10}}><button onClick={()=>{setFromDraft('2015-01-01');setFrom('2015-01-01');}} style={{...S.btnS,fontSize:12,padding:'6px 14px'}}>View all activity</button><div style={{fontSize:11,color:T.textMuted,marginTop:6}}>The default view shows the last 12 months. This account's activity may be older \u2014 e.g. an investment purchase.</div></div>}</td></tr>
             :lines.map((l,i)=>{running+=isDr?(l.debit-l.credit):(l.credit-l.debit);
               const tip=(l.created_by?'Posted by '+l.created_by:'')+(l.created_at?(l.created_by?' on ':'Posted on ')+new Date(l.created_at+(l.created_at.includes('Z')||l.created_at.includes('+')?'':'Z')).toLocaleString('en-US',{timeZone:'America/Los_Angeles',year:'numeric',month:'short',day:'numeric',hour:'numeric',minute:'2-digit',hour12:true,timeZoneName:'short'}):'');
               return<tr key={i}>
                 <td style={{...S.td,color:T.textMuted,whiteSpace:'nowrap'}}>{l.date}</td>
                 <td style={S.td} title={tip}><button style={{background:'none',border:0,padding:0,color:T.accent,fontWeight:600,cursor:'pointer',fontSize:'inherit',fontFamily:'inherit'}} onClick={()=>openJE(l.jeId)}>JE-{String(l.entry_num).padStart(4,'0')}</button></td>
                 <td style={S.td}>{l.memo}</td>
                 <td style={S.td}>{l.offset}</td>
                 <td style={S.td}>{l.vendor}</td>
                 <td style={S.tdR}>{l.debit>0?fmt(l.debit):''}</td>
                 <td style={S.tdR}>{l.credit>0?fmt(l.credit):''}</td>
                 <td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(running)}</td></tr>;})}
           <tr style={S.grandTotalRow}>
             <td style={{...S.tdBold}} colSpan={5}>Period Totals</td>
             <td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalDr)}</td>
             <td style={{...S.tdBold,textAlign:'right'}}>${fmt(totalCr)}</td>
             <td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>${fmt(running)}</td></tr>
         </tbody></table></div>}
    {viewEntry&&<EditJEModal entityId={entityId} entry={viewEntry} accounts={allAccounts} onClose={()=>setViewEntry(null)} onSaved={()=>{setViewEntry(null);setReloadKey(k=>k+1);onChanged&&onChanged();}}/>}
  </div></div>);
}

function BalanceSheet({entityId,entityName,asOf,setAsOf,canEdit=true}){
  const[drillAcct,setDrillAcct]=useState(null);const[rk,setRk]=useState(0);
  const[dateFilter,setDateFilter]=useState('all');const[colMode,setColMode]=useState('total');const[compare,setCompare]=useState(false);
  const anchor=/^\d{4}-\d{2}-\d{2}$/.test(asOf)?asOf:today();
  const periods=useMemo(()=>rptPeriods(dateFilter,colMode,anchor),[dateFilter,colMode,anchor]);
  const prior=(compare&&colMode==='total'&&periods[0]&&periods[0].from)?rptPriorWindow(periods[0]):null;
  const cols=useMemo(()=>prior?[prior,...periods]:periods,[JSON.stringify(prior),JSON.stringify(periods)]);
  const[data,setData]=useState([]);
  useEffect(()=>{let ok=true;Promise.all(cols.map(c=>api.getBalances(entityId,{as_of:c.to,close_pl_before:c.to.slice(0,4)+'-01-01'}).catch(()=>[]))).then(r=>{if(ok)setData(r);});return()=>{ok=false;};},[entityId,JSON.stringify(cols),rk]);
  const meta=useMemo(()=>{const m=new Map();data.forEach(bs=>(bs||[]).forEach(b=>{if(!m.has(b.code))m.set(b.code,{code:b.code,name:b.name,type:b.type});}));return m;},[data]);
  const vmap=useMemo(()=>data.map(bs=>{const mm=new Map();(bs||[]).forEach(b=>mm.set(b.code,b.balance));return mm;}),[data]);
  const val=(code,ci)=>(vmap[ci]&&vmap[ci].get(code))||0;
  const grp=type=>[...meta.values()].filter(a=>a.type===type).filter(a=>vmap.some(mm=>Math.abs((mm&&mm.get(a.code))||0)>0.005));
  const assets=grp('Asset'),liabs=grp('Liability'),eq=grp('Equity');
  const sumC=(items,ci)=>items.reduce((s,a)=>s+val(a.code,ci),0);
  const niCol=ci=>{let r=0,e=0;(data[ci]||[]).forEach(b=>{if(b.type==='Revenue')r+=b.balance;else if(b.type==='Expense')e+=b.balance;});return r-e;};
  const tA=ci=>sumC(assets,ci);const tE=ci=>sumC(eq,ci)+niCol(ci);const tLE=ci=>sumC(liabs,ci)+tE(ci);
  const nCols=cols.length;const curI=nCols-1;const priI=prior?0:-1;
  const pctTxt=p=>p==null?'—':(p>=0?'+':'')+p.toFixed(1)+'%';
  const chgCells=(getter)=>{if(!prior)return null;const c=getter(curI),p=getter(priI);return[<td key="d" style={S.tdR}>{fmt(c-p)}</td>,<td key="p" style={{...S.tdR,color:(c-p)>=0?T.green:T.red}}>{pctTxt(rptPct(c,p))}</td>];};
  const nColSpan=1+nCols+(prior?2:0);
  const oneYrBefore=d=>{const x=new Date(d+'T00:00:00');x.setFullYear(x.getFullYear()-1);x.setDate(x.getDate()+1);return x.toISOString().slice(0,10);};
  const Sec=({title,items,totalGetter,extraNiRow})=>(<><tr><td style={S.sectionHeader} colSpan={nColSpan}>{title}</td></tr>
    {items.map(a=><tr key={a.code}><td style={S.indentTd}>{a.name}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct({code:a.code,name:a.name,from:oneYrBefore(c.to),to:c.to})}>{fmt(val(a.code,i))}</td>)}{chgCells(ci=>val(a.code,ci))}</tr>)}
    {extraNiRow&&<tr><td style={{...S.indentTd,fontStyle:'italic',color:T.textMuted}}>Net Income (current period)</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,fontStyle:'italic'}}>{fmt(niCol(i))}</td>)}{chgCells(niCol)}</tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(totalGetter(i))}</td>)}{chgCells(totalGetter)}</tr></>);
  const colHead=(c,i)=>prior&&i===0?'Prev':(c.label==='Total'?('As of '+anchor):c.label);
  const doExport=()=>{const hdr=['',...cols.map((c,i)=>colHead(c,i)),...(prior?['$ Change','% Change']:[])];
    const d=[[entityName||'Balance Sheet'],['Balance Sheet'],[],hdr];
    const push=(label,getter)=>{const row=[label,...cols.map((c,i)=>getter(i))];if(prior)row.push(getter(curI)-getter(priI),rptPct(getter(curI),getter(priI)));d.push(row);};
    d.push(['Assets']);assets.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));push('Total Assets',tA);
    d.push(['Liabilities']);liabs.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));push('Total Liabilities',ci=>sumC(liabs,ci));
    d.push(['Equity']);eq.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));push('  Net Income (current period)',niCol);push('Total Equity',tE);
    push('Total Liabilities + Equity',tLE);exportToExcel(d,'BS_'+anchor+'.xlsx');};
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end',marginBottom:16,flexWrap:'wrap',gap:10}}>
    <div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
      <ReportControls dateFilter={dateFilter} setDateFilter={setDateFilter} colMode={colMode} setColMode={setColMode} compare={compare} setCompare={setCompare}/>
    </div>
    <div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='bs' currentConfig={{asOf,dateFilter,colMode,compare}} onApply={(c)=>{if(c.asOf)setAsOf(c.asOf);if(c.dateFilter)setDateFilter(c.dateFilter);if(c.colMode)setColMode(c.colMode);if(typeof c.compare==='boolean')setCompare(c.compare);}} canEdit={canEdit}/><button style={S.btnExport} onClick={doExport}>Export Excel</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Balance Sheet</div><div style={{fontSize:13,color:T.textMuted}}>As of {anchor}{colMode!=='total'?(' · '+RPT_COL_MODES.find(m=>m[0]===colMode)[1]):''}</div></div>
    <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:520,margin:'0 auto'}}><thead><tr><th style={S.th}></th>{cols.map((c,i)=><th key={i} style={S.thR}>{colHead(c,i)}</th>)}{prior&&<><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}</tr></thead><tbody>
      <Sec title="Assets" items={assets} totalGetter={tA}/><tr><td colSpan={nColSpan} style={{padding:6}}/></tr>
      <Sec title="Liabilities" items={liabs} totalGetter={ci=>sumC(liabs,ci)}/><tr><td colSpan={nColSpan} style={{padding:3}}/></tr>
      <Sec title="Equity" items={eq} totalGetter={tE} extraNiRow/>
      <tr style={S.grandTotalRow}><td style={S.tdBold}>Total Liabilities + Equity</td>{cols.map((c,i)=><td key={i} style={{...S.tdBold,textAlign:'right',fontSize:15}}>{fmt(tLE(i))}</td>)}{chgCells(tLE)}</tr>
    </tbody></table></div>
    <div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(tA(curI)-tLE(curI))<0.005?T.green:T.red}}>{Math.abs(tA(curI)-tLE(curI))<0.005?'A = L + E':'Off by $'+fmt(tA(curI)-tLE(curI))}</div></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillAcct.from} to={drillAcct.to} onClose={()=>setDrillAcct(null)} onChanged={()=>setRk(k=>k+1)}/>}
    </div>);}

function IncomeStatement({entityId,entityName,from,setFrom,to,setTo,canEdit=true}){
  const[drillAcct,setDrillAcct]=useState(null);const[rk,setRk]=useState(0);
  const[dateFilter,setDateFilter]=useState('all');const[colMode,setColMode]=useState('total');const[compare,setCompare]=useState(false);
  const anchor=/^\d{4}-\d{2}-\d{2}$/.test(to)?to:today();
  const periods=useMemo(()=>rptPeriods(dateFilter,colMode,anchor),[dateFilter,colMode,anchor]);
  const prior=(compare&&colMode==='total'&&periods[0]&&periods[0].from)?rptPriorWindow(periods[0]):null;
  const cols=useMemo(()=>prior?[prior,...periods]:periods,[JSON.stringify(prior),JSON.stringify(periods)]);
  const[data,setData]=useState([]);
  useEffect(()=>{let ok=true;Promise.all(cols.map(c=>api.getBalances(entityId,{from:c.from||undefined,to:c.to}).catch(()=>[]))).then(r=>{if(ok)setData(r);});return()=>{ok=false;};},[entityId,JSON.stringify(cols),rk]);
  const meta=useMemo(()=>{const m=new Map();data.forEach(bs=>(bs||[]).forEach(b=>{if(!m.has(b.code))m.set(b.code,{code:b.code,name:b.name,type:b.type,subtype:b.subtype});}));return m;},[data]);
  const vmap=useMemo(()=>data.map(bs=>{const mm=new Map();(bs||[]).forEach(b=>mm.set(b.code,b.balance));return mm;}),[data]);
  const val=(code,ci)=>(vmap[ci]&&vmap[ci].get(code))||0;
  const grp=pred=>[...meta.values()].filter(pred).filter(a=>vmap.some(mm=>Math.abs((mm&&mm.get(a.code))||0)>0.005));
  const rev=grp(a=>a.type==='Revenue');const cogs=grp(a=>a.type==='Expense'&&a.subtype==='COGS');
  const opex=grp(a=>a.type==='Expense'&&a.subtype==='Operating Expense');const other=grp(a=>a.type==='Expense'&&a.subtype!=='COGS'&&a.subtype!=='Operating Expense');
  const sumC=(items,ci)=>items.reduce((s,a)=>s+val(a.code,ci),0);
  const ni=ci=>sumC(rev,ci)-sumC(cogs,ci)-sumC(opex,ci)-sumC(other,ci);
  const nCols=cols.length;const curI=nCols-1;const priI=prior?0:-1;
  const pctTxt=p=>p==null?'—':(p>=0?'+':'')+p.toFixed(1)+'%';
  const chgCells=(getter)=>{if(!prior)return null;const c=getter(curI),p=getter(priI),pc=rptPct(c,p);return[<td key="d" style={S.tdR}>{fmt(c-p)}</td>,<td key="p" style={{...S.tdR,color:(c-p)>=0?T.green:T.red}}>{pctTxt(pc)}</td>];};
  const nColSpan=1+nCols+(prior?2:0);
  const Sec=({title,items})=>(<><tr><td style={S.sectionHeader} colSpan={nColSpan}>{title}</td></tr>
    {items.map(a=><tr key={a.code}><td style={S.indentTd}>{a.name}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct({code:a.code,name:a.name,from:c.from,to:c.to})}>{fmt(val(a.code,i))}</td>)}{chgCells(ci=>val(a.code,ci))}</tr>)}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(sumC(items,i))}</td>)}{chgCells(ci=>sumC(items,ci))}</tr></>);
  const TotalRow=({label,getter,big})=>(<tr style={big?S.grandTotalRow:{background:T.bgElevated}}><td style={big?{...S.tdBold,fontSize:15}:{...S.td,fontWeight:700,color:T.textBright}}>{label}</td>{cols.map((c,i)=><td key={i} style={big?{...S.tdBold,textAlign:'right',fontSize:16,color:getter(i)>=0?T.green:T.red}:{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(getter(i))}</td>)}{chgCells(getter)}</tr>);
  const colHead=(c,i)=>prior&&i===0?'Prev':(c.label==='Total'?'Amount':c.label);
  const doExport=()=>{const hdr=['Account',...cols.map((c,i)=>colHead(c,i)),...(prior?['$ Change','% Change']:[])];
    const d=[[entityName||'Income Statement'],['Income Statement'],[(dateFilter==='all'?'Through '+anchor:RPT_DATE_FILTERS.find(f=>f[0]===dateFilter)[1]+' ending '+anchor)+(colMode!=='total'?' · '+RPT_COL_MODES.find(m=>m[0]===colMode)[1]:'')],[],hdr];
    const push=(label,getter)=>{const row=[label,...cols.map((c,i)=>getter(i))];if(prior)row.push(getter(curI)-getter(priI),rptPct(getter(curI),getter(priI)));d.push(row);};
    [['Revenue',rev],['Cost of Goods Sold',cogs],['Operating Expenses',opex],['Other Expenses',other]].forEach(([t,items])=>{if(!items.length)return;d.push([t]);items.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));push('Total '+t,ci=>sumC(items,ci));});
    push('Net Income',ci=>ni(ci));exportToExcel(d,'IS_'+anchor+'.xlsx');};
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end',marginBottom:16,flexWrap:'wrap',gap:10}}>
    <div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>Period end</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <ReportControls dateFilter={dateFilter} setDateFilter={setDateFilter} colMode={colMode} setColMode={setColMode} compare={compare} setCompare={setCompare}/>
    </div>
    <div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='is' currentConfig={{to,dateFilter,colMode,compare}} onApply={(c)=>{if(c.to)setTo(c.to);if(c.dateFilter)setDateFilter(c.dateFilter);if(c.colMode)setColMode(c.colMode);if(typeof c.compare==='boolean')setCompare(c.compare);}} canEdit={canEdit}/><button style={S.btnExport} onClick={doExport}>Export Excel</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Income Statement</div><div style={{fontSize:13,color:T.textMuted}}>{dateFilter==='all'?('Through '+anchor):(RPT_DATE_FILTERS.find(f=>f[0]===dateFilter)[1]+' ending '+anchor)}{colMode!=='total'?(' · '+RPT_COL_MODES.find(m=>m[0]===colMode)[1]):''}</div></div>
    <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:520,margin:'0 auto'}}><thead><tr><th style={S.th}>Account</th>{cols.map((c,i)=><th key={i} style={S.thR}>{colHead(c,i)}</th>)}{prior&&<><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}</tr></thead><tbody>
      <Sec title="Revenue" items={rev}/>
      {cogs.length>0&&<><Sec title="Cost of Goods Sold" items={cogs}/><TotalRow label="Gross Profit" getter={ci=>sumC(rev,ci)-sumC(cogs,ci)}/></>}
      <Sec title="Operating Expenses" items={opex}/>
      <TotalRow label="Operating Income" getter={ci=>sumC(rev,ci)-sumC(cogs,ci)-sumC(opex,ci)}/>
      {other.length>0&&<Sec title="Other Expenses" items={other}/>}
      <TotalRow label="Net Income" getter={ci=>ni(ci)} big/>
    </tbody></table></div></div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillAcct.from||from} to={drillAcct.to||to} onClose={()=>setDrillAcct(null)} onChanged={()=>setRk(k=>k+1)}/>}
    </div>);}

// ═══ Custom Detail Report (Q6: multi-account, grouped by class/location, with subtotals) ═══
function CustomDetailReport({entityId,entityName,dimsEnabled,canEdit=true,pendingConfig,clearPending}){
  const[accounts,setAccounts]=useState([]);const[sel,setSel]=useState([]);const[acctSearch,setAcctSearch]=useState('');
  const[from,setFrom]=useState('');const[to,setTo]=useState('');
  const[groupBy,setGroupBy]=useState(dimsEnabled?'class':'none');
  const[colMode,setColMode]=useState('total');const[compare,setCompare]=useState(false);const[priorRows,setPriorRows]=useState([]);
  const[rows,setRows]=useState(null);const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  const[begRows,setBegRows]=useState([]); // beginning balances for selected balance-sheet accounts
  const prevDay=(d)=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return x.toISOString().slice(0,10);};
  const isBS=(t)=>t==='Asset'||t==='Liability'||t==='Equity';
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts).catch(()=>setAccounts([]));},[entityId]);
  useEffect(()=>{if(pendingConfig){if(pendingConfig.sel)setSel(pendingConfig.sel);setFrom(pendingConfig.from||'');setTo(pendingConfig.to||'');if(pendingConfig.groupBy)setGroupBy(pendingConfig.groupBy);clearPending&&clearPending();}},[]);
  useEffect(()=>{if(pendingConfig){if(pendingConfig.sel)setSel(pendingConfig.sel);if(pendingConfig.dim)setDim(pendingConfig.dim);setFrom(pendingConfig.from||'');setTo(pendingConfig.to||'');clearPending&&clearPending();}},[]);
  const toggle=code=>setSel(s=>s.includes(code)?s.filter(c=>c!==code):[...s,code]);
  const filteredAccts=accounts.filter(a=>!acctSearch||acctLabel(a.code,a.name).toLowerCase().includes(acctSearch.toLowerCase()));
  const run=async()=>{
    if(!sel.length){setErr('Select at least one account');return;}
    setLoading(true);setErr('');setRows(null);setBegRows([]);
    try{
      const all=await api.getGLDetail(entityId,{from:from||undefined,to:to||undefined});
      const selSet=new Set(sel);
      setRows((all.lines||all||[]).filter(l=>selSet.has(l.account_code)));
      // Comparative: pull the equal-length prior window's activity for the same accounts.
      if(compare&&from&&to){const pw=priorWindow(from,to);if(pw){const pall=await api.getGLDetail(entityId,{from:pw.from,to:pw.to});setPriorRows((pall.lines||pall||[]).filter(l=>selSet.has(l.account_code)));}else setPriorRows([]);}else setPriorRows([]);
      // Beginning balance only applies to balance-sheet accounts and only when a
      // start date is set (otherwise the period runs from inception → beg = 0).
      const bsSel=accounts.filter(a=>selSet.has(a.code)&&isBS(a.type));
      if(from&&bsSel.length){
        const bals=await api.getBalances(entityId,{as_of:prevDay(from)});
        const byCode=new Map((bals||[]).map(b=>[b.code,b.balance]));
        setBegRows(bsSel.map(a=>({code:a.code,name:a.name,type:a.type,balance:byCode.get(a.code)||0})).filter(r=>r.balance!==0));
      }
    }catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const groupKey=l=>groupBy==='class'?(l.class_name||'(no class)'):groupBy==='location'?(l.location_name||'(no location)'):'All';
  const groups=(()=>{if(!rows)return[];const m=new Map();rows.forEach(l=>{const k=groupKey(l);if(!m.has(k))m.set(k,[]);m.get(k).push(l);});return[...m.entries()].sort((a,b)=>a[0].localeCompare(b[0]));})();
  const amt=l=>(l.debit||0)-(l.credit||0);
  const grand=rows?rows.reduce((s,l)=>s+amt(l),0):0;
  const begTotal=begRows.reduce((s,b)=>s+(b.balance||0),0);
  // Column-display + comparative computeds (Liting #2).
  const cols=buildPeriodCols(from,to,colMode);
  const showTotal=colMode!=='total';
  const colIdxOf=(date)=>{if(colMode==='total')return 0;for(let i=0;i<cols.length;i++){if(date>=cols[i].from&&date<=cols[i].to)return i;}return -1;};
  const sumByCol=(lines)=>{const arr=cols.map(()=>0);let tot=0;(lines||[]).forEach(l=>{const a=amt(l);tot+=a;const ci=colIdxOf(l.date);if(ci>=0)arr[ci]+=a;});return{arr,tot};};
  const priorGroupMap=(()=>{const m=new Map();priorRows.forEach(l=>{const k=groupKey(l);m.set(k,(m.get(k)||0)+amt(l));});return m;})();
  const priorGrand=priorRows.reduce((s,l)=>s+amt(l),0);
  const descCols=4;
  const totalColCount=descCols+cols.length+(showTotal?1:0)+(compare?3:0);
  const pctTxt=p=>p==null?'—':(p>=0?'+':'')+p.toFixed(1)+'%';
  const cmpCells=(cur,pri)=>{const d=cur-pri;const p=pri!==0?(d/Math.abs(pri))*100:null;return[<td key="pp" style={{...S.tdR,fontWeight:700}}>{fmt(pri)}</td>,<td key="dd" style={{...S.tdR,fontWeight:700,color:d>=0?T.green:T.red}}>{fmt(d)}</td>,<td key="pc" style={{...S.tdR,fontWeight:700,color:d>=0?T.green:T.red}}>{pctTxt(p)}</td>];};
  const doExport=()=>{
    const d=[[entityName||'Custom Detail Report'],['Custom Detail Report'],['Period: '+(from||'Begin')+' to '+(to||today())],[]];
    if(begRows.length>0){
      d.push(['Beginning Balances — Balance Sheet accounts as of '+from]);
      d.push(['','Account','','','','','Balance']);
      begRows.forEach(b=>d.push(['',b.code+' '+b.name,'','','','',b.balance]));
      d.push(['','','','','','Total Beginning Balance',begTotal]);d.push([]);
    }
    const amtHdr=cols.map(c=>c.label);const cmpHdr=compare?['Prev Period','$ Change','% Change']:[];
    const pctN=(cur,pri)=>pri!==0?+(((cur-pri)/Math.abs(pri))*100).toFixed(1):'';
    groups.forEach(([g,lines])=>{
      if(groupBy!=='none')d.push([g]);
      d.push(['Account','Date','JE','Description',...amtHdr,...(showTotal?['Total']:[]),...cmpHdr]);
      lines.forEach(l=>{const a=amt(l);const ci=colIdxOf(l.date);const cells=cols.map((c,k)=>(colMode==='total'||k===ci)?a:'');d.push([l.account_code+' '+l.account_name,l.date,'JE-'+String(l.entry_num).padStart(4,'0'),l.description||l.memo||'',...cells,...(showTotal?[a]:[]),...(compare?['','','']:[])]);});
      const {arr,tot}=sumByCol(lines);const pri=priorGroupMap.get(g)||0;
      d.push(['Total'+(groupBy!=='none'?' for '+g:''),'','','',...arr,...(showTotal?[tot]:[]),...(compare?[pri,tot-pri,pctN(tot,pri)]:[])]);d.push([]);
    });
    const gg=sumByCol(rows||[]);
    d.push(['PERIOD ACTIVITY','','','',...gg.arr,...(showTotal?[gg.tot]:[]),...(compare?[priorGrand,gg.tot-priorGrand,pctN(gg.tot,priorGrand)]:[])]);
    if(begRows.length>0)d.push(['ENDING BALANCE (BS accts: beginning + activity)','','','',...(colMode==='total'?[begTotal+grand]:[...cols.map(()=>''),begTotal+grand]),...(compare?['','','']:[])]);
    exportToExcel(d,'Custom_Detail_'+(to||today())+'.xlsx');
  };
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}><div><div style={S.h1}>Custom Detail Report</div><div style={S.sub}>Pick accounts, optionally group by class or location</div></div><div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='customdetail' currentConfig={{sel,from,to,groupBy,colMode,compare}} onApply={(c)=>{setSel(c.sel||[]);setFrom(c.from||'');setTo(c.to||'');if(c.groupBy)setGroupBy(c.groupBy);if(c.colMode)setColMode(c.colMode);if(typeof c.compare==='boolean')setCompare(c.compare);}} canEdit={canEdit}/>{rows&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>
    <div style={S.card}>
      <div style={{display:'flex',gap:24,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 320px',minWidth:280}}>
          <label style={S.label}>Accounts ({sel.length} selected)</label>
          <input style={{...S.inputSm,width:'100%',marginBottom:6}} placeholder="Search accounts..." value={acctSearch} onChange={e=>setAcctSearch(e.target.value)}/>
          <div style={{maxHeight:200,overflowY:'auto',border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:6}}>
            {filteredAccts.map(a=><label key={a.code} style={{display:'flex',alignItems:'center',gap:8,fontSize:12,padding:'3px 4px',cursor:'pointer'}}><input type="checkbox" checked={sel.includes(a.code)} onChange={()=>toggle(a.code)}/>{acctLabel(a.code,a.name)}</label>)}
          </div>
          <div style={{marginTop:4,display:'flex',gap:10}}><button style={{...S.btnGhost,fontSize:11,color:T.accent}} onClick={()=>setSel(filteredAccts.map(a=>a.code))}>Select all shown</button><button style={{...S.btnGhost,fontSize:11,color:T.textMuted}} onClick={()=>setSel([])}>Clear</button></div>
        </div>
        <div style={{flex:'0 0 200px'}}>
          <div style={{marginBottom:10}}><label style={S.label}>From</label><input style={{...S.inputSm,width:'100%'}} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
          <div style={{marginBottom:10}}><label style={S.label}>To</label><input style={{...S.inputSm,width:'100%'}} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
          <div style={{marginBottom:10,display:'flex',gap:6,flexWrap:'wrap'}}>{PRESETS.map(([k,lbl])=><button key={k} onClick={()=>{const r=presetRange(k);setFrom(r.from);setTo(r.to);}} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'5px 9px',cursor:'pointer'}}>{lbl}</button>)}</div>
          <div style={{marginBottom:10}}><label style={S.label}>Columns</label><select style={{...S.inputSm,width:'100%'}} value={colMode} onChange={e=>setColMode(e.target.value)}>{COL_MODES.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>
          <label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,marginBottom:10,cursor:'pointer',color:T.textMuted}}><input type="checkbox" checked={compare} onChange={e=>setCompare(e.target.checked)}/>Compare to prior period</label>
          {dimsEnabled&&<div><label style={S.label}>Group by</label><select style={{...S.inputSm,width:'100%'}} value={groupBy} onChange={e=>setGroupBy(e.target.value)}><option value="none">No grouping</option><option value="class">{classTerm()==='Class'?'Class / Investor':classTerm()}</option><option value="location">Location</option></select></div>}
        </div>
      </div>
      {err&&<div style={S.err}>{err}</div>}
      <button style={{...S.btnP,marginTop:14}} onClick={run} disabled={loading}>{loading?'Running...':'Run Report'}</button>
    </div>
    {rows&&<div style={S.cardFlush}>
      {begRows.length>0&&<table style={{...S.table,marginBottom:0}}><thead><tr><th style={S.th} colSpan={4}>Beginning Balances — Balance Sheet accounts as of {from}</th><th style={S.thR}>Balance</th></tr></thead>
        <tbody>{begRows.map((b,i)=><tr key={'beg'+i}><td style={S.td} colSpan={4}>{b.code} {b.name}</td><td style={{...S.tdR,color:b.balance<0?T.red:T.textBright}}>{fmt(b.balance)}</td></tr>)}
          <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600}} colSpan={4}>Total Beginning Balance</td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(begTotal)}</td></tr></tbody></table>}
      {groups.length===0&&begRows.length===0?<div style={{padding:24,color:T.textDim}}>No activity for the selected accounts/period.</div>:
      <div style={{overflowX:'auto'}}><table className="cl-colresize" style={S.table}><thead><tr>
        <th style={S.th}>Account</th><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Description</th>
        {cols.map((c,i)=><th key={i} style={S.thR}>{c.label}</th>)}
        {showTotal&&<th style={S.thR}>Total</th>}
        {compare&&<><th style={S.thR}>Prev Period</th><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}
      </tr></thead>
      <tbody>{groups.map(([g,lines])=>{const {arr,tot}=sumByCol(lines);const pri=priorGroupMap.get(g)||0;return<Fragment key={g}>
        {groupBy!=='none'&&<tr style={{background:T.bgElevated}}><td style={{...S.tdBold,color:T.textBright}} colSpan={totalColCount}>{g}</td></tr>}
        {lines.map((l,i)=>{const a=amt(l);const ci=colIdxOf(l.date);return<tr key={i}>
          <td style={S.td}>{l.account_code} {l.account_name}</td><td style={{...S.td,whiteSpace:'nowrap'}}>{l.date}</td><td style={S.td}>JE-{String(l.entry_num).padStart(4,'0')}</td><td style={S.td}>{l.description||l.memo||''}</td>
          {cols.map((c,k)=><td key={k} style={{...S.tdR,color:a<0?T.red:T.textBright}}>{(colMode==='total'||k===ci)?fmt(a):''}</td>)}
          {showTotal&&<td style={{...S.tdR,color:a<0?T.red:T.textBright}}>{fmt(a)}</td>}
          {compare&&<><td style={S.tdR}></td><td style={S.tdR}></td><td style={S.tdR}></td></>}
        </tr>;})}
        {groupBy!=='none'&&<tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600}} colSpan={descCols}>Total for {g}</td>{arr.map((v,k)=><td key={k} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(v)}</td>)}{showTotal&&<td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(tot)}</td>}{compare&&cmpCells(tot,pri)}</tr>}
      </Fragment>;})}
        {(()=>{const {arr,tot}=sumByCol(rows||[]);return<tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={descCols}>PERIOD ACTIVITY</td>{arr.map((v,k)=><td key={k} style={{...S.tdBold,textAlign:'right',color:T.textBright}}>{fmt(v)}</td>)}{showTotal&&<td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>{fmt(tot)}</td>}{compare&&cmpCells(tot,priorGrand)}</tr>;})()}
        {begRows.length>0&&<tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={descCols+cols.length-1}>ENDING BALANCE (Balance Sheet accts: beginning + activity)</td><td style={{...S.tdBold,textAlign:'right',color:T.textBright}} colSpan={1+(showTotal?1:0)}>{fmt(begTotal+grand)}</td>{compare&&<><td/><td/><td/></>}</tr>}
      </tbody></table></div>}
    </div>}
  </div>);
}

// ═══ Pivot Summary Report (Q7: class × account matrix, totals by class — for PCAP) ═══
function PivotReport({entityId,entityName,canEdit=true,pendingConfig,clearPending}){
  const[accounts,setAccounts]=useState([]);const[sel,setSel]=useState([]);const[acctSearch,setAcctSearch]=useState('');
  const[dim,setDim]=useState('class');const[from,setFrom]=useState('');const[to,setTo]=useState('');
  const[data,setData]=useState(null);const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts).catch(()=>setAccounts([]));},[entityId]);
  const toggle=code=>setSel(s=>s.includes(code)?s.filter(c=>c!==code):[...s,code]);
  const filteredAccts=accounts.filter(a=>!acctSearch||acctLabel(a.code,a.name).toLowerCase().includes(acctSearch.toLowerCase()));
  const run=async()=>{
    if(!sel.length){setErr('Select at least one account');return;}
    setLoading(true);setErr('');setData(null);
    try{setData(await api.getPivot(entityId,{dim,accounts:sel.join(','),from:from||undefined,to:to||undefined}));}
    catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const doExport=()=>{
    if(!data)return;
    const head=[dim==='class'?(classTerm()==='Class'?'Class / Investor':classTerm()):dim==='location'?'Location':'Project',...data.columns.map(c=>c.code+' '+c.name),'Total'];
    const d=[[entityName||'Pivot Report'],['Pivot Summary by '+(dim==='class'?classTerm():dim)],['Period: '+(from||'Begin')+' to '+(to||today())],[],head];
    data.rows.forEach(r=>d.push([r.name,...data.columns.map(c=>r.cells[c.code]||0),r.total]));
    d.push(['Total',...data.columns.map(c=>data.column_totals[c.code]||0),data.grand_total]);
    exportToExcel(d,'Pivot_'+dim+'_'+(to||today())+'.xlsx');
  };
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}><div><div style={S.h1}>Pivot Summary</div><div style={S.sub}>Totals by class across selected accounts — for PCAP letters</div></div><div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='pivot' currentConfig={{sel,dim,from,to}} onApply={(c)=>{setSel(c.sel||[]);setDim(c.dim||'class');setFrom(c.from||'');setTo(c.to||'');}} canEdit={canEdit}/>{data&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>
    <div style={S.card}>
      <div style={{display:'flex',gap:24,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 320px',minWidth:280}}>
          <label style={S.label}>Accounts ({sel.length} selected)</label>
          <input style={{...S.inputSm,width:'100%',marginBottom:6}} placeholder="Search accounts..." value={acctSearch} onChange={e=>setAcctSearch(e.target.value)}/>
          <div style={{maxHeight:200,overflowY:'auto',border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:6}}>
            {filteredAccts.map(a=><label key={a.code} style={{display:'flex',alignItems:'center',gap:8,fontSize:12,padding:'3px 4px',cursor:'pointer'}}><input type="checkbox" checked={sel.includes(a.code)} onChange={()=>toggle(a.code)}/>{acctLabel(a.code,a.name)}</label>)}
          </div>
          <div style={{marginTop:4,display:'flex',gap:10}}><button style={{...S.btnGhost,fontSize:11,color:T.accent}} onClick={()=>setSel(filteredAccts.map(a=>a.code))}>Select all shown</button><button style={{...S.btnGhost,fontSize:11,color:T.textMuted}} onClick={()=>setSel([])}>Clear</button></div>
        </div>
        <div style={{flex:'0 0 200px'}}>
          <div style={{marginBottom:10}}><label style={S.label}>Pivot by</label><select style={{...S.inputSm,width:'100%'}} value={dim} onChange={e=>setDim(e.target.value)}><option value="class">{classTerm()==='Class'?'Class / Investor':classTerm()}</option><option value="location">Location</option><option value="project">Project</option></select></div>
          <div style={{marginBottom:10}}><label style={S.label}>From</label><input style={{...S.inputSm,width:'100%'}} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
          <div><label style={S.label}>To</label><input style={{...S.inputSm,width:'100%'}} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
        </div>
      </div>
      {err&&<div style={S.err}>{err}</div>}
      <button style={{...S.btnP,marginTop:14}} onClick={run} disabled={loading}>{loading?'Running...':'Run Pivot'}</button>
    </div>
    {data&&<div style={{...S.cardFlush,overflowX:'auto'}}>
      {data.rows.length===0?<div style={{padding:24,color:T.textDim}}>No activity for the selected accounts/period.</div>:
      <table style={S.table}><thead><tr><th style={{...S.th,position:'sticky',left:0,background:T.bgCard}}>{dim==='class'?(classTerm()==='Class'?'Class / Investor':classTerm()):dim==='location'?'Location':'Project'}</th>{data.columns.map(c=><th key={c.code} style={S.thR} title={c.code+' '+c.name}>{c.name||c.code}</th>)}<th style={S.thR}>Total</th></tr></thead>
      <tbody>{data.rows.map(r=><tr key={r.id}><td style={{...S.td,position:'sticky',left:0,background:T.bgCard,fontWeight:500}}>{r.name}</td>{data.columns.map(c=><td key={c.code} style={S.tdR}>{r.cells[c.code]?fmt(r.cells[c.code]):''}</td>)}<td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(r.total)}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={{...S.tdBold,position:'sticky',left:0,background:T.bgCard}}>Total</td>{data.columns.map(c=><td key={c.code} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(data.column_totals[c.code]||0)}</td>)}<td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>{fmt(data.grand_total)}</td></tr>
      </tbody></table>}
    </div>}
  </div>);
}

// ═══ AP Aging Detail (Q5: open bills from Bill.com, bucketed by days past due) ═══
function ApAgingReport({entityId,entityName,canEdit=true,pendingConfig,clearPending}){
  const[asOf,setAsOf]=useState(today());
  useEffect(()=>{if(pendingConfig){if(pendingConfig.asOf)setAsOf(pendingConfig.asOf);clearPending&&clearPending();}},[]);
  const[data,setData]=useState(null);const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  const[viewEntry,setViewEntry]=useState(null);
  const[entryLoading,setEntryLoading]=useState(false);
  // GL rows only carry an entry id; fetch the full entry (with lines) before
  // opening the JE modal, which requires entry.lines to render.
  const openEntry=async(id)=>{if(!id)return;setEntryLoading(true);try{const full=await api.getEntry(entityId,id);setViewEntry(full);}catch(e){alert('Could not open entry: '+e.message);}finally{setEntryLoading(false);}};
  const BK=['current','d1_30','d31_60','d61_90','d91_plus'];
  const run=async()=>{
    setLoading(true);setErr('');setData(null);
    try{setData(await api.getApAging(entityId,asOf||undefined));}
    catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const lbl=d=>data?data.bucket_labels[d]:d;
  const COLS=6; // Date, Type, Num, Vendor, Due Date, Past Due
  const ncols=COLS+BK.length+2; // + GL + Amount
  const doExport=()=>{
    if(!data)return;
    const head=['Date','Type','Num','Vendor','Due Date','Past Due (days)','Current','1-30','31-60','61-90','91+','GL','Amount'];
    const d=[[entityName||'AP Aging Detail'],['A/P Aging Detail — built from GL '+(data.ap_account||'202000')],['As of '+data.as_of],[],head];
    data.vendors.forEach(g=>{
      g.rows.forEach(r=>d.push([r.date,r.type,r.num,r.vendor,r.due_date,r.past_due_days,
        r.bucket==='current'?r.amount:'',r.bucket==='d1_30'?r.amount:'',r.bucket==='d31_60'?r.amount:'',r.bucket==='d61_90'?r.amount:'',r.bucket==='d91_plus'?r.amount:'','',r.amount]));
      d.push(['Total '+g.vendor,'','','','','',g.subtotal.current,g.subtotal.d1_30,g.subtotal.d31_60,g.subtotal.d61_90,g.subtotal.d91_plus,'',g.subtotal.total]);
    });
    if((data.gl_rows||[]).length){
      d.push([]);d.push(['GL ENTRIES (imported / non-Bill.com — not aged)']);
      data.gl_rows.forEach(r=>d.push([r.date,'GL','JE-'+String(r.entry_num).padStart(4,'0'),'',(r.memo||''),'','','','','','',r.amount,r.amount]));
      d.push(['Total GL Entries','','','','','','','','','','',data.gl_total,data.gl_total]);
    }
    const gt=data.grand_total;
    d.push(['TOTAL','','','','','',gt.current,gt.d1_30,gt.d31_60,gt.d61_90,gt.d91_plus,gt.gl,gt.total]);
    d.push(['Reconciliation vs GL '+(data.ap_account||'202000')+' ('+fmt(data.gl_balance)+')','','','','','','','','','','','',data.recon_diff]);
    exportToExcel(d,'AP_Aging_'+data.as_of+'.xlsx');
  };
  const hasAnything=data&&(data.bill_count>0||(data.gl_rows&&data.gl_rows.length>0));
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}><div><div style={S.h1}>A/P Aging Detail</div><div style={S.sub}>Built from GL account {data?data.ap_account:'202000'} &middot; ties to the book{data&&data.billcom_error?' · Bill.com enrich error: '+data.billcom_error:''}</div></div><div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='apaging' currentConfig={{asOf}} onApply={(c)=>{if(c.asOf)setAsOf(c.asOf);}} canEdit={canEdit}/>{hasAnything&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>
    <div style={S.card}>
      <div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div style={{flex:'0 0 180px'}}><label style={S.label}>As of date</label><input style={{...S.inputSm,width:'100%'}} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
        <div style={{display:'flex',alignItems:'flex-end',gap:6,flexWrap:'wrap'}}>{PRESETS.map(([k,lbl])=><button key={k} onClick={()=>setAsOf(k==='all'?today():presetRange(k).to)} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'6px 10px',cursor:'pointer'}}>{lbl}</button>)}</div>
        <button style={{...S.btnP}} onClick={run} disabled={loading}>{loading?'Building from GL…':'Run Aging'}</button>
      </div>
      {err&&<div style={S.err}>{err}</div>}
    </div>
    {data&&<div style={{...S.cardFlush,overflowX:'auto'}}>
      {!hasAnything?<div style={{padding:24,color:T.textDim}}>No open A/P as of {data.as_of}.</div>:
      <table style={S.table}><thead><tr>
        <th style={S.th}>Date</th><th style={S.th}>Type</th><th style={S.th}>Num</th><th style={S.th}>Vendor</th><th style={S.th}>Due Date</th><th style={S.thR}>Past Due</th>
        {BK.map(b=><th key={b} style={S.thR}>{lbl(b)}</th>)}<th style={{...S.thR,color:T.accent}}>GL</th><th style={S.thR}>Amount</th>
      </tr></thead>
      <tbody>{data.vendors.map(g=><Fragment key={g.vendor}>
        <tr><td colSpan={ncols} style={{...S.td,fontWeight:700,color:T.textBright,background:T.bgElevated}}>{g.vendor}</td></tr>
        {g.rows.map((r,i)=><tr key={i}>
          <td style={S.td}>{r.date}</td><td style={S.td}>{r.type}</td><td style={S.td}>{r.num}</td><td style={S.td}>{r.vendor}</td><td style={S.td}>{r.due_date}</td><td style={S.tdR}>{r.past_due_days||''}</td>
          {BK.map(b=><td key={b} style={S.tdR}>{r.bucket===b?fmt(r.amount):''}</td>)}<td style={S.tdR}></td><td style={{...S.tdR,fontWeight:600}}>{fmt(r.amount)}</td>
        </tr>)}
        <tr style={{background:T.bgElevated}}><td colSpan={COLS} style={{...S.td,fontWeight:600,fontStyle:'italic'}}>Total {g.vendor}</td>
          {BK.map(b=><td key={b} style={{...S.tdR,fontWeight:600}}>{g.subtotal[b]?fmt(g.subtotal[b]):''}</td>)}<td style={S.tdR}></td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(g.subtotal.total)}</td></tr>
      </Fragment>)}
        {data.gl_rows&&data.gl_rows.length>0&&<Fragment>
          <tr><td colSpan={ncols} style={{...S.td,fontWeight:700,color:T.accent,background:T.accentDim}}>GL ENTRIES <span style={{fontWeight:400,color:T.textMuted}}>— imported / non-Bill.com &middot; not aged</span></td></tr>
          {data.gl_rows.map((r,i)=><tr key={'gl'+i} onClick={()=>openEntry(r.entry_id)} style={{cursor:r.entry_id?'pointer':'default',opacity:entryLoading?0.6:1}}>
            <td style={S.td}>{r.date}</td><td style={S.td}>GL</td><td style={{...S.td,color:T.accent}}>{r.entry_num!=null?'JE-'+String(r.entry_num).padStart(4,'0'):''}</td><td style={{...S.td,color:T.textMuted}} colSpan={3}>{r.memo}{r.description?' · '+r.description:''}</td>
            {BK.map(b=><td key={b} style={S.tdR}></td>)}<td style={{...S.tdR,fontWeight:600,color:T.accent}}>{fmt(r.amount)}</td><td style={{...S.tdR,fontWeight:600}}>{fmt(r.amount)}</td>
          </tr>)}
          <tr style={{background:T.accentDim}}><td colSpan={COLS} style={{...S.td,fontWeight:600,fontStyle:'italic'}}>Total GL Entries</td>
            {BK.map(b=><td key={b} style={S.tdR}></td>)}<td style={{...S.tdR,fontWeight:700,color:T.accent}}>{fmt(data.gl_total)}</td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(data.gl_total)}</td></tr>
        </Fragment>}
        <tr style={S.grandTotalRow}><td colSpan={COLS} style={S.tdBold}>TOTAL</td>
          {BK.map(b=><td key={b} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(data.grand_total[b]||0)}</td>)}<td style={{...S.tdR,fontWeight:700,color:T.accent}}>{fmt(data.grand_total.gl||0)}</td><td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>{fmt(data.grand_total.total)}</td></tr>
        <tr><td colSpan={ncols} style={{...S.td,textAlign:'right',background:Math.abs(data.recon_diff)<0.005?'#f3faf5':'#fdf2f4',color:Math.abs(data.recon_diff)<0.005?T.green:T.red,fontWeight:600}}>
          Reconciliation vs GL {data.ap_account||'202000'} ({fmt(data.gl_balance)}): {Math.abs(data.recon_diff)<0.005?fmt(0)+' ✓':fmt(data.recon_diff)+' — does not tie'}
        </td></tr>
      </tbody></table>}
    </div>}
    {viewEntry&&<EditJEModal entityId={entityId} entry={viewEntry} accounts={[]} onClose={()=>setViewEntry(null)} onSaved={()=>{setViewEntry(null);run();}}/>}
  </div>);
}

// ═══ Memorized Reports (saved report configurations; shared per entity) ═══
// MemorizeBar renders on each configurable report: a Save button + a dropdown
// of that report's saved configs. onApply restores a saved config's settings.
function MemorizeBar({entityId,reportType,currentConfig,onApply,canEdit=true}){
  const[saved,setSaved]=useState([]);const[open,setOpen]=useState(false);const[saving,setSaving]=useState(false);const[err,setErr]=useState('');
  const load=useCallback(()=>{api.getMemorizedReports(entityId).then(all=>setSaved(all.filter(r=>r.report_type===reportType))).catch(()=>{});},[entityId,reportType]);
  useEffect(()=>{load();},[load]);
  const save=async()=>{const name=prompt('Save this report view as:');if(!name||!name.trim())return;setSaving(true);setErr('');try{await api.createMemorizedReport(entityId,{report_type:reportType,name:name.trim(),config:currentConfig});load();}catch(e){setErr(e.message);alert(e.message);}finally{setSaving(false);}};
  const run=(r)=>{onApply(r.config||{});setOpen(false);};
  const del=async(r,e)=>{e.stopPropagation();if(!confirm('Delete saved report "'+r.name+'"?'))return;try{await api.deleteMemorizedReport(entityId,r.id);load();}catch(ex){alert(ex.message);}};
  return(<div style={{position:'relative',display:'inline-flex',gap:8,alignItems:'center'}}>
    {saved.length>0&&<div style={{position:'relative'}}>
      <button style={{...S.btnGhost,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:'7px 12px',fontSize:12}} onClick={()=>setOpen(!open)}>★ Saved ({saved.length}) ▾</button>
      {open&&<><div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>setOpen(false)}/>
        <div style={{position:'absolute',top:'100%',right:0,marginTop:6,width:280,maxHeight:340,overflowY:'auto',background:'#fff',border:'1px solid '+T.border,borderRadius:T.radius,boxShadow:T.shadowLg,zIndex:100,padding:'6px 0'}}>
          {saved.map(r=><div key={r.id} onClick={()=>run(r)} style={{padding:'8px 14px',cursor:'pointer',fontSize:13,display:'flex',justifyContent:'space-between',alignItems:'center'}} onMouseEnter={e=>e.currentTarget.style.background=T.bgElevated} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
            <span><span style={{fontWeight:600,color:T.textBright}}>{r.name}</span>{r.created_by_name&&<span style={{color:T.textDim,fontSize:11,marginLeft:6}}>· {r.created_by_name}</span>}</span>
            {canEdit&&<button style={{background:'none',border:'none',color:T.red,cursor:'pointer',fontSize:14}} onClick={e=>del(r,e)}>×</button>}</div>)}
        </div></>}
    </div>}
    {canEdit&&<button style={{...S.btnGhost,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:'7px 12px',fontSize:12}} disabled={saving} onClick={save}>{saving?'Saving…':'★ Save view'}</button>}
  </div>);
}

// Dedicated page listing all memorized reports for the entity, grouped by type.
function MemorizedReportsPage({entityId,entityName,canEdit=true,onOpen}){
  const[rows,setRows]=useState(null);const[err,setErr]=useState('');
  const TYPE_LABELS={customdetail:'Custom Detail',pivot:'Pivot Summary',apaging:'AP Aging',drilldown:'Account Drilldown',bs:'Balance Sheet',is:'Income Statement',trial:'Trial Balance'};
  const load=useCallback(()=>{setErr('');api.getMemorizedReports(entityId).then(setRows).catch(e=>setErr(e.message));},[entityId]);
  useEffect(()=>{load();},[load]);
  const del=async(r)=>{if(!confirm('Delete saved report "'+r.name+'"?'))return;try{await api.deleteMemorizedReport(entityId,r.id);load();}catch(ex){alert(ex.message);}};
  const groups={};(rows||[]).forEach(r=>{(groups[r.report_type]=groups[r.report_type]||[]).push(r);});
  return(<div><div style={{marginBottom:8}}><div style={S.h1}>Memorized Reports</div><div style={S.sub}>Saved report views for {entityName} · shared with everyone on this entity</div></div>
    {err&&<div style={S.err}>{err}</div>}
    {rows&&rows.length===0?<div style={{...S.card,textAlign:'center',padding:50,color:T.textDim}}>No saved reports yet. Open any report (Custom Detail, Pivot, AP Aging, etc.), set it up the way you like, and click "★ Save view".</div>:
     !rows?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Loading…</div>:
     <div style={{display:'flex',flexDirection:'column',gap:18}}>{Object.keys(groups).map(tp=><div key={tp}>
       <div style={{fontSize:12,fontWeight:700,letterSpacing:'0.05em',textTransform:'uppercase',color:T.textDim,marginBottom:8}}>{TYPE_LABELS[tp]||tp}</div>
       <div style={{...S.cardFlush}}><table style={S.table}><thead><tr><th style={S.th}>Name</th><th style={S.th}>Saved by</th><th style={S.th}>Saved</th><th style={{...S.th,width:160}}>Actions</th></tr></thead>
         <tbody>{groups[tp].map(r=><tr key={r.id}>
           <td style={{...S.td,fontWeight:600,color:T.textBright}}>{r.name}</td>
           <td style={{...S.td,color:T.textMuted}}>{r.created_by_name||'—'}</td>
           <td style={{...S.td,color:T.textMuted}}>{(r.created_at||'').slice(0,10)}</td>
           <td style={S.td}><div style={{display:'flex',gap:8}}>
             <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>onOpen&&onOpen(r)}>Open</button>
             {canEdit&&<button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={()=>del(r)}>Delete</button>}
           </div></td></tr>)}
         </tbody></table></div>
     </div>)}</div>}
  </div>);
}

// ═══ Workpapers › Management Fee (CLRF) — roll prior quarter forward ═══
function MgmtFeeWorkpaper({entityId,entityName,canEdit=true}){
  const[file,setFile]=useState(null);
  const[analysis,setAnalysis]=useState(null);
  const[rows,setRows]=useState([]); // {name,group,beginning_commitment,change}
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[result,setResult]=useState(null);
  const fmt=n=>n==null?'-':n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
  const onPick=async(f)=>{
    setErr('');setResult(null);setAnalysis(null);setRows([]);setFile(f);
    if(!f)return;
    setBusy(true);
    try{const a=await api.mgmtFeeAnalyze(entityId,f);setAnalysis(a);setRows((a.investors||[]).map(i=>({...i})));}
    catch(e){setErr(e.message);} finally{setBusy(false);}
  };
  const setChange=(i,v)=>setRows(rs=>rs.map((r,idx)=>idx===i?{...r,change:v}:r));
  const totalChange=rows.reduce((s,r)=>s+(Number(String(r.change).replace(/[$,\s]/g,''))||0),0);
  const generate=async()=>{
    setErr('');setBusy(true);setResult(null);
    try{
      const changes=rows.filter(r=>Number(String(r.change).replace(/[$,\s]/g,''))!==0).map(r=>({name:r.name,change:Number(String(r.change).replace(/[$,\s]/g,''))}));
      const out=await api.mgmtFeeGenerate(entityId,file,changes,analysis?.next_quarter?.start);
      if(!out)return;
      const url=URL.createObjectURL(out.blob);const a=document.createElement('a');a.href=url;a.download=out.filename;a.click();URL.revokeObjectURL(url);
      setResult(out.summary||{});
    }catch(e){setErr(e.message);} finally{setBusy(false);}
  };
  return(<div>
    <div style={S.h1}>Management Fee Workpaper</div>
    <div style={{color:T.textMuted,marginBottom:16,fontSize:13,maxWidth:760}}>Upload the prior quarter's management-fee workbook. CloudLedger reads the investor list, group classifications, rate tables and tier splits, rolls the quarter forward (new dates, ending → next beginning), applies any commitment changes you enter, and produces the next quarter's workbook.</div>
    {err&&<div style={S.err}>{err}</div>}

    <div style={{...S.card,marginBottom:16}}>
      <div style={{...S.h2,marginBottom:10}}>1 · Upload prior-quarter workbook</div>
      <input type="file" accept=".xlsx" disabled={busy||!canEdit} onChange={e=>onPick(e.target.files[0])} style={{fontSize:13}}/>
      {file&&<span style={{marginLeft:10,color:T.textMuted,fontSize:12}}>{file.name}</span>}
      {busy&&!analysis&&<div style={{marginTop:8,color:T.textMuted,fontSize:12}}>Reading workbook…</div>}
    </div>

    {analysis&&<>
      <div style={{...S.card,marginBottom:16}}>
        <div style={{...S.h2,marginBottom:10}}>2 · Quarter roll-forward</div>
        <div style={{display:'flex',gap:30,flexWrap:'wrap',fontSize:13}}>
          <div><div style={{color:T.textDim,fontSize:11}}>PRIOR QUARTER</div><div style={{color:T.textBright,fontWeight:600}}>{analysis.prior_quarter||'—'}</div><div style={{color:T.textMuted,fontSize:12}}>starts {analysis.prior_quarter_start}</div></div>
          <div style={{fontSize:20,color:T.textDim,alignSelf:'center'}}>→</div>
          <div><div style={{color:T.textDim,fontSize:11}}>NEW QUARTER</div><div style={{color:T.accent,fontWeight:600}}>{analysis.next_quarter?.label}</div><div style={{color:T.textMuted,fontSize:12}}>{analysis.next_quarter?.start} – {analysis.next_quarter?.end} ({analysis.next_quarter?.days} days)</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>INVESTORS</div><div style={{color:T.textBright,fontWeight:600}}>{analysis.investor_count}</div><div style={{color:T.textMuted,fontSize:12}}>{Object.entries(analysis.groups||{}).map(([g,n])=>g+':'+n).join('  ')}</div></div>
        </div>
      </div>

      <div style={{...S.card,marginBottom:16}}>
        <div style={{...S.h2,marginBottom:6}}>3 · Commitment changes this quarter <span style={{fontWeight:400,color:T.textMuted,fontSize:12}}>(leave 0 if unchanged)</span></div>
        <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>Each investor's prior ending commitment carries to the new beginning. Enter new capital calls, transfers, or redemptions as a positive/negative change.</div>
        <div style={{maxHeight:'46vh',overflowY:'auto'}}>
        <table style={S.table}><thead style={{position:'sticky',top:0,background:T.bgElevated,zIndex:1}}><tr><th style={S.th}>Investor</th><th style={S.th}>Group</th><th style={S.thR}>Beginning</th><th style={S.thR}>Change (+/−)</th><th style={S.thR}>New Ending</th></tr></thead>
        <tbody>{rows.map((r,i)=>{const chg=Number(String(r.change).replace(/[$,\s]/g,''))||0;const end=(r.beginning_commitment||0)+chg;return(
          <tr key={i}><td style={S.td}>{r.name}</td><td style={S.td}><span style={S.tag(r.group)}>{r.group}</span></td>
          <td style={S.tdR}>{fmt(r.beginning_commitment)}</td>
          <td style={{...S.tdR,padding:'2px 8px'}}><input value={r.change} disabled={!canEdit} onChange={e=>setChange(i,e.target.value)} style={{...S.input,width:120,textAlign:'right',padding:'4px 8px',fontSize:12}}/></td>
          <td style={{...S.tdR,color:chg!==0?T.accent:T.text,fontWeight:chg!==0?600:400}}>{fmt(end)}</td></tr>);})}
        </tbody></table>
        </div>
        {totalChange!==0&&<div style={{marginTop:8,fontSize:12,color:T.textMuted}}>Net commitment change: <span style={{color:T.accent,fontWeight:600}}>{fmt(totalChange)}</span></div>}
      </div>

      <div style={{display:'flex',gap:10,alignItems:'center'}}>
        <button style={{...S.btnP,opacity:busy?0.6:1}} disabled={busy||!canEdit} onClick={generate}>{busy?'Generating…':'Generate '+(analysis.next_quarter?.label||'next quarter')+' workbook'}</button>
      </div>

      {result&&<div style={{...S.card,marginTop:16,borderColor:T.green+'55'}}>
        <div style={{...S.h2,marginBottom:8,color:T.green}}>✓ {result.quarter} workbook generated</div>
        <div style={{display:'flex',gap:24,flexWrap:'wrap',fontSize:13}}>
          <div><div style={{color:T.textDim,fontSize:11}}>STANDARD</div><div style={{color:T.textBright}}>{fmt(result.standard)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>BBR</div><div style={{color:T.textBright}}>{fmt(result.bbr)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>GCM</div><div style={{color:T.textBright}}>{fmt(result.gcm)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>USC</div><div style={{color:T.textBright}}>{fmt(result.usc)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>TOTAL QUARTERLY FEE</div><div style={{color:T.accent,fontWeight:700,fontSize:15}}>{fmt(result.total)}</div></div>
        </div>
        <div style={{marginTop:8,fontSize:12,color:T.textMuted}}>The .xlsx has downloaded. Review the calc tab before sending.</div>
      </div>}
    </>}
  </div>);
}

// ═══ Fund Reporting — CLRF-style LP fund statement package (config + generate) ═══
function FundReporting({entityId,entityName}){
  const[asOf,setAsOf]=useState(today());
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[invs,setInvs]=useState(null);
  const[classes,setClasses]=useState(null);
  const[savingId,setSavingId]=useState(null);
  const[classFilter,setClassFilter]=useState('');
  const blank={parent_name:'',name:'',acquisition_date:'',cost:'',fair_value:'',sort_order:''};
  const[draft,setDraft]=useState(blank);

  const load=()=>{
    setErr('');
    api.getFundInvestments(entityId).then(setInvs).catch(e=>setErr(e.message));
    api.getClasses(entityId).then(setClasses).catch(e=>setErr(e.message));
  };
  useEffect(()=>{setInvs(null);setClasses(null);load();},[entityId]);

  const addInv=async()=>{
    if(!draft.name.trim()){setErr('Investment name is required');return;}
    setBusy(true);setErr('');
    try{await api.createFundInvestment(entityId,{...draft,cost:Number(draft.cost)||0,fair_value:Number(draft.fair_value)||0,sort_order:Number(draft.sort_order)||0});setDraft(blank);load();}
    catch(e){setErr(e.message);}finally{setBusy(false);}
  };
  const saveInv=async(row)=>{
    setSavingId(row.id);setErr('');
    try{await api.updateFundInvestment(entityId,row.id,{parent_name:row.parent_name,name:row.name,acquisition_date:row.acquisition_date,cost:Number(row.cost)||0,fair_value:Number(row.fair_value)||0,sort_order:Number(row.sort_order)||0});}
    catch(e){setErr(e.message);}finally{setSavingId(null);}
  };
  const delInv=async(id)=>{setErr('');try{await api.deleteFundInvestment(entityId,id);load();}catch(e){setErr(e.message);}};
  const setInvField=(id,f,v)=>setInvs(list=>list.map(r=>r.id===id?{...r,[f]:v}:r));

  const toggleGP=async(cls)=>{
    const next=(cls.partner_type==='GP')?'LP':'GP';
    setClasses(list=>list.map(c=>c.id===cls.id?{...c,partner_type:next}:c));
    try{await api.setClassPartnerType(entityId,cls.id,next);}catch(e){setErr(e.message);load();}
  };

  const genPdf=async()=>{
    if(!/^\d{4}-\d{2}-\d{2}$/.test(asOf)){setErr('Pick a valid as-of date');return;}
    setBusy(true);setErr('');
    try{const out=await api.getFundStatementsPdf(entityId,asOf);if(!out)return;
      const url=URL.createObjectURL(out.blob);const a=document.createElement('a');a.href=url;a.download=out.filename;a.click();URL.revokeObjectURL(url);}
    catch(e){setErr(e.message);}finally{setBusy(false);}
  };

  const gpList=(classes||[]).filter(c=>c.partner_type==='GP');
  const shownClasses=(classes||[]).filter(c=>!classFilter||c.name.toLowerCase().includes(classFilter.toLowerCase()));
  const th={textAlign:'left',padding:'6px 8px',borderBottom:'2px solid '+T.border,color:T.textDim,fontSize:11,fontWeight:700};
  const td={padding:'4px 8px',borderBottom:'1px solid '+T.border,fontSize:12};
  const cellInput={width:'100%',background:'transparent',border:'1px solid '+T.border,borderRadius:4,padding:'3px 6px',color:T.text,fontSize:12};

  return(<div>
    <div style={{marginBottom:8}}>
      <div style={S.h1}>Fund Reporting</div>
      <div style={S.sub}>Limited-partnership fund statement package (Assets/Liabilities/Partners' Capital, Schedule of Investments, Operations, Changes in Partners' Capital, Cash Flows) · {entityName}</div>
    </div>
    {err&&<div style={{...S.err,marginBottom:12}}>{err}</div>}

    {/* Generate */}
    <div style={{...S.card,marginBottom:16}}>
      <div style={S.h2}>Generate statement package</div>
      <div style={{display:'flex',gap:10,alignItems:'flex-end',marginTop:10,flexWrap:'wrap'}}>
        <div><label style={S.label}>As of date</label><input type='date' value={asOf} onChange={e=>setAsOf(e.target.value)} style={{...S.input,width:170}}/></div>
        <button style={{...S.btnP,opacity:busy?0.6:1}} disabled={busy} onClick={genPdf}>{busy?'Generating…':'Generate PDF'}</button>
      </div>
      <div style={{fontSize:12,color:T.textMuted,marginTop:8}}>Amounts come from the general ledger. The Schedule of Investments and the GP/LP capital split use the settings below.</div>
    </div>

    {/* Schedule of Investments editor */}
    <div style={{...S.card,marginBottom:16}}>
      <div style={S.h2}>Schedule of Investments</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>Per-underlying look-through detail (not in the GL). Group underlyings under a holding company via “Parent”. Percentages of partners' capital are computed at generation.</div>
      {invs===null?<div style={{color:T.textMuted,fontSize:12}}>Loading…</div>:
      <div style={{overflowX:'auto'}}>
        <table style={{borderCollapse:'collapse',width:'100%',minWidth:720}}>
          <thead><tr>
            <th style={th}>Parent (holding co.)</th><th style={th}>Investment</th><th style={th}>Acq. date</th>
            <th style={{...th,textAlign:'right'}}>Cost</th><th style={{...th,textAlign:'right'}}>Fair value</th>
            <th style={{...th,textAlign:'right'}}>Order</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {invs.map(r=>(<tr key={r.id}>
              <td style={td}><input style={cellInput} value={r.parent_name||''} onChange={e=>setInvField(r.id,'parent_name',e.target.value)}/></td>
              <td style={td}><input style={cellInput} value={r.name||''} onChange={e=>setInvField(r.id,'name',e.target.value)}/></td>
              <td style={td}><input style={{...cellInput,width:90}} value={r.acquisition_date||''} placeholder='m/d/yyyy' onChange={e=>setInvField(r.id,'acquisition_date',e.target.value)}/></td>
              <td style={{...td,textAlign:'right'}}><input style={{...cellInput,textAlign:'right'}} value={r.cost} onChange={e=>setInvField(r.id,'cost',e.target.value)}/></td>
              <td style={{...td,textAlign:'right'}}><input style={{...cellInput,textAlign:'right'}} value={r.fair_value} onChange={e=>setInvField(r.id,'fair_value',e.target.value)}/></td>
              <td style={{...td,textAlign:'right',width:60}}><input style={{...cellInput,textAlign:'right'}} value={r.sort_order} onChange={e=>setInvField(r.id,'sort_order',e.target.value)}/></td>
              <td style={{...td,whiteSpace:'nowrap'}}>
                <button style={{...S.btnGhost,color:T.green,fontSize:11,marginRight:6,opacity:savingId===r.id?0.6:1}} disabled={savingId===r.id} onClick={()=>saveInv(r)}>{savingId===r.id?'…':'Save'}</button>
                <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={()=>delInv(r.id)}>Delete</button>
              </td>
            </tr>))}
            {/* add-new row */}
            <tr>
              <td style={td}><input style={cellInput} value={draft.parent_name} placeholder='CLRFI Midco I, LLC' onChange={e=>setDraft({...draft,parent_name:e.target.value})}/></td>
              <td style={td}><input style={cellInput} value={draft.name} placeholder='New investment' onChange={e=>setDraft({...draft,name:e.target.value})}/></td>
              <td style={td}><input style={{...cellInput,width:90}} value={draft.acquisition_date} placeholder='m/d/yyyy' onChange={e=>setDraft({...draft,acquisition_date:e.target.value})}/></td>
              <td style={{...td,textAlign:'right'}}><input style={{...cellInput,textAlign:'right'}} value={draft.cost} placeholder='0' onChange={e=>setDraft({...draft,cost:e.target.value})}/></td>
              <td style={{...td,textAlign:'right'}}><input style={{...cellInput,textAlign:'right'}} value={draft.fair_value} placeholder='0' onChange={e=>setDraft({...draft,fair_value:e.target.value})}/></td>
              <td style={{...td,textAlign:'right',width:60}}><input style={{...cellInput,textAlign:'right'}} value={draft.sort_order} placeholder='0' onChange={e=>setDraft({...draft,sort_order:e.target.value})}/></td>
              <td style={td}><button style={{...S.btnP,padding:'4px 10px',opacity:busy?0.6:1}} disabled={busy} onClick={addInv}>Add</button></td>
            </tr>
          </tbody>
        </table>
      </div>}
    </div>

    {/* GP/LP tagging */}
    <div style={{...S.card}}>
      <div style={S.h2}>General Partner designation</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>Tag the investor classes that are General Partners. Everything else is treated as a Limited Partner for the GP/LP capital split. Currently {gpList.length} class{gpList.length===1?'':'es'} tagged GP.</div>
      <input style={{...S.input,maxWidth:320,marginBottom:10}} placeholder='Filter classes…' value={classFilter} onChange={e=>setClassFilter(e.target.value)}/>
      {classes===null?<div style={{color:T.textMuted,fontSize:12}}>Loading…</div>:
      <div style={{maxHeight:340,overflowY:'auto',border:'1px solid '+T.border,borderRadius:6}}>
        <table style={{borderCollapse:'collapse',width:'100%'}}>
          <thead><tr><th style={th}>Investor class</th><th style={{...th,width:120,textAlign:'center'}}>Type</th></tr></thead>
          <tbody>
            {shownClasses.map(c=>(<tr key={c.id}>
              <td style={td}>{c.name}</td>
              <td style={{...td,textAlign:'center'}}>
                <button onClick={()=>toggleGP(c)} style={{cursor:'pointer',fontSize:11,fontWeight:700,padding:'3px 10px',borderRadius:12,border:'1px solid '+(c.partner_type==='GP'?(T.green||'#2a9d5a'):T.border),background:c.partner_type==='GP'?(T.green||'#2a9d5a')+'22':'transparent',color:c.partner_type==='GP'?(T.green||'#2a9d5a'):T.textDim}}>{c.partner_type==='GP'?'General Partner':'Limited Partner'}</button>
              </td>
            </tr>))}
          </tbody>
        </table>
      </div>}
    </div>
  </div>);
}

// ═══ Trailing 12 Months — P&L with 12 monthly columns + a Total column ═══
function TrailingTwelveMonths({entityId,entityName}){
  const[asOf,setAsOf]=useState(today());
  const[data,setData]=useState(null);
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[analysis,setAnalysis]=useState(null);
  const[analyzing,setAnalyzing]=useState(false);
  const[analysisErr,setAnalysisErr]=useState('');
  const fmt=n=>{const v=Number(n)||0;const t=Math.abs(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});return v<0?'('+t+')':(v===0?'-':t);};
  useEffect(()=>{let cancelled=false;setData(null);setErr('');setAnalysis(null);setAnalysisErr('');
    if(!entityId||!/^\d{4}-\d{2}-\d{2}$/.test(asOf))return;
    setBusy(true);
    api.getTtmPL(entityId,asOf)
      .then(d=>{if(!cancelled)setData(d);})
      .catch(e=>{if(!cancelled)setErr(e.message);})
      .finally(()=>{if(!cancelled)setBusy(false);});
    return()=>{cancelled=true;};
  },[entityId,asOf]);
  const months=data?data.meta.months:[];
  const nCols=months.length; // 12
  // Build the ordered display rows for both the on-screen table and the export.
  const buildRows=()=>{
    if(!data)return[];
    const rows=[];
    const line=(label,vals,total,opt={})=>rows.push({label,vals,total,...opt});
    // Revenue
    line('Revenue',null,null,{header:true});
    data.revenue.forEach(l=>line(l.name,l.vals,l.total,{indent:1}));
    line('Total Revenue',data.totRev.vals,data.totRev.total,{bold:true,rule:true});
    // Cost of Revenue (only if present)
    if(data.hasCogs){
      line('Cost of Revenue',null,null,{header:true});
      data.cogs.forEach(l=>line(l.name,l.vals,l.total,{indent:1}));
      line('Total Cost of Revenue',data.totCogs.vals,data.totCogs.total,{bold:true,rule:true});
      line('Gross Profit',data.grossProfit.vals,data.grossProfit.total,{bold:true,rule:true});
    }
    // Operating Expenses, grouped
    line('Operating Expenses',null,null,{header:true});
    data.opexGroups.forEach(g=>{
      if(data.opexGroups.length>1){
        line(g.title,null,null,{indent:1,sub:true});
        g.lines.forEach(l=>line(l.name,l.vals,l.total,{indent:2}));
        line('Total '+g.title,g.subtotal.vals,g.subtotal.total,{indent:1,rule:true});
      }else{
        g.lines.forEach(l=>line(l.name,l.vals,l.total,{indent:1}));
      }
    });
    line('Total Operating Expenses',data.totOpex.vals,data.totOpex.total,{bold:true,rule:true});
    // Net Income
    line('Net Income (Loss)',data.netIncome.vals,data.netIncome.total,{bold:true,rule:true,dbl:true});
    return rows;
  };
  const runAnalysis=async()=>{
    if(!data)return;setAnalyzing(true);setAnalysisErr('');
    try{const a=await api.analyzeTtmPL(entityId,asOf);if(a)setAnalysis(a);}
    catch(e){setAnalysisErr(e.message);}finally{setAnalyzing(false);}
  };
  const rows=buildRows();
  // Download the styled workbook from the server (ExcelJS): comma-formatted
  // amounts, underlined month header, and underlined subtotal/grand-total rows.
  const doExport=async()=>{
    if(!data)return;
    setErr('');
    try{
      const out=await api.getTtmPLXlsx(entityId,asOf,analysis);
      if(!out)return;
      const url=URL.createObjectURL(out.blob);const a=document.createElement('a');a.href=url;a.download=out.filename;a.click();URL.revokeObjectURL(url);
    }catch(e){setErr(e.message);}
  };
  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
      <div><div style={S.h1}>Trailing 12 Months</div><div style={S.sub}>P&amp;L activity by month for the trailing twelve months &middot; {entityName||'this entity'}</div></div>
      <div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
        <div><label style={S.label}>As of date</label><input type='date' value={asOf} onChange={e=>setAsOf(e.target.value)} style={{...S.input,width:160}}/></div>
        {data&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}
      </div>
    </div>
    {err&&<div style={S.err}>{err}</div>}
    {busy&&!data&&<div style={{color:T.textMuted,fontSize:12,padding:12}}>Computing trailing 12 months…</div>}
    {data&&<div style={{...S.card,padding:0,overflowX:'auto'}}>
      <table style={{borderCollapse:'collapse',width:'100%',fontSize:12,whiteSpace:'nowrap'}}>
        <thead>
          <tr>
            <th style={{position:'sticky',left:0,background:T.cardBg||T.bg,textAlign:'left',padding:'8px 12px',borderBottom:'2px solid '+T.border,color:T.textDim,fontSize:11}}>{data.meta.periodLabel}</th>
            {months.map((m,i)=><th key={i} style={{textAlign:'right',padding:'8px 10px',borderBottom:'2px solid '+T.border,color:T.textDim,fontSize:11}}>{m.label}</th>)}
            <th style={{textAlign:'right',padding:'8px 12px',borderBottom:'2px solid '+T.border,color:T.textBright,fontSize:11,fontWeight:700,borderLeft:'2px solid '+T.border}}>{data.meta.totalLabel||'Total'}</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r,ri)=>(
            <tr key={ri} style={{background:r.header?(T.hover||'transparent'):'transparent'}}>
              <td style={{position:'sticky',left:0,background:T.cardBg||T.bg,padding:'4px 12px',paddingLeft:(12+(r.indent||0)*16)+'px',fontWeight:(r.bold||r.header||r.sub)?600:400,color:r.header?T.textBright:T.text,borderTop:r.rule?'1px solid '+T.border:'none'}}>{r.label}</td>
              {(r.vals||new Array(nCols).fill(null)).map((v,ci)=>(
                <td key={ci} style={{textAlign:'right',padding:'4px 10px',fontWeight:r.bold?600:400,color:T.text,borderTop:r.rule?'1px solid '+T.border:'none',borderBottom:r.dbl?'3px double '+T.border:'none'}}>{r.vals?fmt(v):''}</td>
              ))}
              <td style={{textAlign:'right',padding:'4px 12px',fontWeight:(r.bold||r.header)?700:600,color:T.textBright,borderLeft:'2px solid '+T.border,borderTop:r.rule?'1px solid '+T.border:'none',borderBottom:r.dbl?'3px double '+T.border:'none'}}>{r.total==null?'':fmt(r.total)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>}
    {data&&(()=>{
      return(<div style={{...S.card,marginTop:16}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:12,flexWrap:'wrap'}}>
          <div><div style={S.h2}>Items Needing Attention</div><div style={{fontSize:12,color:T.textMuted,marginTop:2}}>AI review of the trailing-twelve-month trends, in order of importance.</div></div>
          <button style={{...S.btnP,opacity:analyzing?0.6:1}} disabled={analyzing} onClick={runAnalysis}>{analyzing?'Analyzing…':analysis?'Re-analyze with Claude':'Analyze with Claude'}</button>
        </div>
        {analysisErr&&<div style={{...S.err,marginTop:12}}>{analysisErr}</div>}
        {!analysis&&!analyzing&&!analysisErr&&<div style={{fontSize:13,color:T.textMuted,marginTop:12}}>Click “Analyze with Claude” to generate a review of items needing attention. It reads the 12-month P&amp;L and lists the accounts worth a closer look.</div>}
        {analyzing&&<div style={{fontSize:13,color:T.textMuted,marginTop:12}}>Claude is reviewing the trailing twelve months…</div>}
        {analysis&&<div style={{marginTop:12}}>
          {analysis.summary&&<div style={{fontSize:13,color:T.text,marginBottom:analysis.findings.length?14:0,lineHeight:1.5}}>{analysis.summary}</div>}
          {analysis.findings.length>0?analysis.findings.map((it,i)=>(<div key={i} style={{display:'flex',gap:10,alignItems:'flex-start',padding:'9px 0',borderTop:i?'1px solid '+T.border:'none'}}>
            <span style={{flexShrink:0,width:20,textAlign:'right',fontSize:13,fontWeight:700,color:T.textDim}}>{i+1}.</span>
            <div style={{fontSize:13,color:T.text,lineHeight:1.5}}><span style={{fontWeight:600,color:T.textBright}}>{it.account||it.title}</span>{(it.account||it.title)&&(it.reason||it.detail)?' — ':''}{it.reason||it.detail}</div>
          </div>)):<div style={{fontSize:13,color:T.green||'#2a9d5a',marginTop:4}}>Nothing flagged for this period.</div>}
          <div style={{fontSize:11,color:T.textDim,marginTop:12,fontStyle:'italic'}}>Generated by Claude · review before relying on it.</div>
        </div>}
      </div>);
    })()}
  </div>);
}

// ═══ Workpapers › Financial Statements — GL-derived statement package (PDF) ═══
function FinancialStatements({entityId,entityName,canEdit=true}){
  const[asOf,setAsOf]=useState(today());
  const[period,setPeriod]=useState('monthly');
  const[execFile,setExecFile]=useState(null);
  const[reqFile,setReqFile]=useState(null);
  const[preview,setPreview]=useState(null);
  const[busy,setBusy]=useState(false);
  const[gen,setGen]=useState(false);
  const[err,setErr]=useState('');
  const[result,setResult]=useState(null);
  const fmt=n=>n==null?'-':Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
  // Re-run the numeric preview whenever the date or period changes.
  useEffect(()=>{let cancelled=false;setPreview(null);setErr('');setResult(null);
    if(!entityId||!/^\d{4}-\d{2}-\d{2}$/.test(asOf))return;
    setBusy(true);
    api.financialStatementsPreview(entityId,asOf,period)
      .then(p=>{if(!cancelled)setPreview(p);})
      .catch(e=>{if(!cancelled)setErr(e.message);})
      .finally(()=>{if(!cancelled)setBusy(false);});
    return()=>{cancelled=true;};
  },[entityId,asOf,period]);
  const generate=async()=>{
    setErr('');setGen(true);setResult(null);
    try{
      const out=await api.financialStatementsGenerate(entityId,asOf,period,execFile,reqFile);
      if(!out)return;
      const url=URL.createObjectURL(out.blob);const a=document.createElement('a');a.href=url;a.download=out.filename;a.click();URL.revokeObjectURL(url);
      setResult(out.summary||{});
    }catch(e){setErr(e.message);}finally{setGen(false);}
  };
  const periods=[['monthly','Monthly'],['quarterly','Quarterly'],['annually','Annually']];
  const tieOk=preview&&preview.checks&&preview.checks.balanceSheetTies;
  const cfTie=preview&&preview.checks&&preview.checks.cashFlowTies;
  return(<div>
    <div style={S.h1}>Financial Statements</div>
    <div style={{color:T.textMuted,marginBottom:16,fontSize:13,maxWidth:820}}>Generates a GL-derived statement package for {entityName||'this entity'} — Balance Sheet, Statements of Operations, Statement of Cash Flows, and Statement of Changes in Members' Equity — as of a date, then merges it into a single PDF with your uploaded executive summary and requisition report. The requisition report's Current & Prior Invoice Log pages are removed automatically.</div>
    {err&&<div style={S.err}>{err}</div>}

    <div style={{...S.card,marginBottom:16}}>
      <div style={{...S.h2,marginBottom:10}}>1 · Statement date &amp; basis</div>
      <div style={{display:'flex',gap:20,flexWrap:'wrap',alignItems:'flex-end'}}>
        <div><label style={S.label}>As of date</label><input type="date" value={asOf} disabled={!canEdit} onChange={e=>setAsOf(e.target.value)} style={{...S.input,width:170}}/></div>
        <div><label style={S.label}>Period basis</label>
          <div style={{display:'flex',gap:6}}>{periods.map(([v,l])=>(
            <button key={v} onClick={()=>setPeriod(v)} disabled={!canEdit} style={{...(period===v?S.btnP:S.btnS),padding:'7px 14px'}}>{l}</button>
          ))}</div>
        </div>
      </div>
      <div style={{marginTop:10,fontSize:12,color:T.textMuted}}>The operations statement compares the {period==='monthly'?'current month vs. prior month':period==='quarterly'?'current quarter vs. prior quarter':'trailing year vs. prior year'}; the year-to-date column and cash-flow statement are always calendar year-to-date.</div>
    </div>

    <div style={{...S.card,marginBottom:16}}>
      <div style={{...S.h2,marginBottom:10}}>2 · Tie-out preview</div>
      {busy&&!preview&&<div style={{color:T.textMuted,fontSize:12}}>Computing statements…</div>}
      {preview&&<div>
        <div style={{display:'flex',gap:24,flexWrap:'wrap',fontSize:13,marginBottom:10}}>
          <div><div style={{color:T.textDim,fontSize:11}}>TOTAL ASSETS</div><div style={{color:T.textBright,fontWeight:600}}>{fmt(preview.totals.totalAssets.cur)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>TOTAL LIAB + EQUITY</div><div style={{color:T.textBright,fontWeight:600}}>{fmt(preview.totals.totalLiabEquity.cur)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>NET INCOME (YTD)</div><div style={{color:T.textBright,fontWeight:600}}>{fmt(preview.totals.netIncomeYtd)}</div></div>
          <div><div style={{color:T.textDim,fontSize:11}}>CASH, END</div><div style={{color:T.textBright,fontWeight:600}}>{fmt(preview.totals.cashEnd)}</div></div>
        </div>
        <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
          <span style={{fontSize:12,padding:'4px 10px',borderRadius:6,background:(tieOk?T.green:T.red)+'22',color:tieOk?T.green:T.red,fontWeight:600}}>{tieOk?'✓ Balance sheet balances':'✗ Balance sheet out by '+fmt(preview.checks.balanceSheetDiff)}</span>
          <span style={{fontSize:12,padding:'4px 10px',borderRadius:6,background:(cfTie?T.green:T.orange)+'22',color:cfTie?T.green:T.orange,fontWeight:600}}>{cfTie?'✓ Cash flow ties':'⚠ Cash flow off by '+fmt(preview.totals.cashFlowTieOut)}</span>
        </div>
        {!cfTie&&<div style={{marginTop:8,fontSize:12,color:T.textMuted}}>A cash-flow difference is usually a mid-year chart change or an opening-balance gap; the statement still generates, with the residual disclosed in a note.</div>}
      </div>}
    </div>

    <div style={{...S.card,marginBottom:16}}>
      <div style={{...S.h2,marginBottom:10}}>3 · Attach supporting PDFs <span style={{fontWeight:400,color:T.textMuted,fontSize:12}}>(optional)</span></div>
      <div style={{display:'flex',flexDirection:'column',gap:12}}>
        <div>
          <label style={S.label}>Executive summary (merged as-is, after the cover)</label>
          <input type="file" accept=".pdf" disabled={!canEdit} onChange={e=>setExecFile(e.target.files[0]||null)} style={{fontSize:13}}/>
          {execFile&&<span style={{marginLeft:10,color:T.textMuted,fontSize:12}}>{execFile.name}</span>}
        </div>
        <div>
          <label style={S.label}>Requisition report (PDF or Excel &mdash; Invoice Log pages removed automatically)</label>
          <input type="file" accept=".pdf,.xlsx,.xls" disabled={!canEdit} onChange={e=>setReqFile(e.target.files[0]||null)} style={{fontSize:13}}/>
          {reqFile&&<span style={{marginLeft:10,color:T.textMuted,fontSize:12}}>{reqFile.name}</span>}
        </div>
      </div>
    </div>

    <div style={{display:'flex',gap:10,alignItems:'center'}}>
      <button style={{...S.btnP,opacity:(gen||!preview)?0.6:1}} disabled={gen||!preview||!canEdit} onClick={generate}>{gen?'Generating…':'Generate financial statements PDF'}</button>
    </div>

    {result&&<div style={{...S.card,marginTop:16,borderColor:T.green+'55'}}>
      <div style={{...S.h2,marginBottom:8,color:T.green}}>✓ Package generated ({result.pages} pages)</div>
      <div style={{fontSize:13,color:T.text}}>
        {(result.sections||[]).map((s,i)=><span key={i} style={{marginRight:14}}>{s.label}: <b>{s.pages}</b>p</span>)}
      </div>
      {result.reqTotal!=null&&<div style={{marginTop:8,fontSize:12,color:T.textMuted}}>Requisition report{result.reqConvertedFromXlsx?(' (converted from Excel'+(result.reqSheetUsed?', sheet "'+result.reqSheetUsed+'"':'')+')'):''}: kept {result.reqKept} of {result.reqTotal} pages{(result.reqRemoved&&result.reqRemoved.length)?(' (removed '+result.reqRemoved.length+' invoice-log page'+(result.reqRemoved.length>1?'s':'')+')'):''}.</div>}
      {(result.warnings||[]).length>0&&<div style={{marginTop:8,fontSize:12,color:T.orange}}>{result.warnings.map((w,i)=><div key={i}>⚠ {w}</div>)}</div>}
      <div style={{marginTop:8,fontSize:12,color:T.textMuted}}>The PDF has downloaded. Review before distributing.</div>
    </div>}
  </div>);
}

// ═══ Investor Commitments (informational capital register; never posts to GL) ═══
function CommitmentsPage({entityId,entityName,canEdit=true}){
  const[data,setData]=useState(null);const[classes,setClasses]=useState([]);const[err,setErr]=useState('');const[loading,setLoading]=useState(true);
  const[showAdd,setShowAdd]=useState(false);const[form,setForm]=useState({class_id:'',commitment_amount:'',called_amount:'',commit_date:'',notes:''});
  const[editId,setEditId]=useState(null);const[editForm,setEditForm]=useState({});
  const load=useCallback(async()=>{setLoading(true);setErr('');try{const[d,c]=await Promise.all([api.getCommitments(entityId),api.getClasses(entityId)]);setData(d);setClasses(c||[]);}catch(e){setErr(e.message);}finally{setLoading(false);}},[entityId]);
  useEffect(()=>{load();},[load]);
  const pct=v=>(v*100).toFixed(2)+'%';
  const committedClassIds=new Set((data?.investors||[]).map(i=>i.class_id));
  const availClasses=classes.filter(c=>!committedClassIds.has(c.id)||c.id===Number(form.class_id));
  const add=async()=>{if(!form.class_id){setErr('Pick an investor');return;}try{await api.createCommitment(entityId,{class_id:Number(form.class_id),commitment_amount:Number(form.commitment_amount||0),called_amount:Number(form.called_amount||0),commit_date:form.commit_date||null,notes:form.notes||null});setShowAdd(false);setForm({class_id:'',commitment_amount:'',called_amount:'',commit_date:'',notes:''});load();}catch(e){setErr(e.message);}};
  const startEdit=i=>{setEditId(i.id);setEditForm({commitment_amount:i.commitment_amount,called_amount:i.called_amount,commit_date:i.commit_date||'',notes:i.notes||''});};
  const saveEdit=async()=>{try{await api.updateCommitment(entityId,editId,{commitment_amount:Number(editForm.commitment_amount||0),called_amount:Number(editForm.called_amount||0),commit_date:editForm.commit_date||null,notes:editForm.notes||null});setEditId(null);load();}catch(e){setErr(e.message);}};
  const del=async i=>{if(!confirm('Remove commitment for '+i.investor+'?'))return;try{await api.deleteCommitment(entityId,i.id);load();}catch(e){setErr(e.message);}};
  const doExport=()=>{if(!data)return;const d=[[entityName||'Investor Commitments'],['Investor Commitments'],[],['Investor','Code','Commitment','Called to Date','Uncalled','% Called','Ownership %','Commit Date','Notes']];
    data.investors.forEach(i=>d.push([i.investor,i.investor_code||'',i.commitment_amount,i.called_amount,i.uncalled_amount,i.pct_called,i.ownership_pct,i.commit_date||'',i.notes||'']));
    const t=data.totals;d.push(['Total','',t.commitment_amount,t.called_amount,t.uncalled_amount,'','','','']);
    exportToExcel(d,'Investor_Commitments.xlsx');};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
    <div><div style={S.h1}>Investor Commitments</div><div style={S.sub}>Capital commitments by investor &middot; informational only (does not post to the GL)</div></div>
    <div style={{display:'flex',gap:8}}>{data&&data.investors.length>0&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}{canEdit&&<button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setErr('');}}>{showAdd?'Cancel':'+ Add Commitment'}</button>}</div></div>
    {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}>
      <div style={S.row}>
        <div style={{...S.col,flex:2}}><label style={S.label}>Investor (class)</label><select style={S.select} value={form.class_id} onChange={e=>setForm(f=>({...f,class_id:e.target.value}))}><option value=''>Select investor…</option>{availClasses.map(c=><option key={c.id} value={c.id}>{c.code?c.code+' — ':''}{c.name}</option>)}</select></div>
        <div style={S.col}><label style={S.label}>Commitment</label><input style={{...S.input,textAlign:'right'}} value={form.commitment_amount} onChange={e=>setForm(f=>({...f,commitment_amount:e.target.value}))} placeholder='0.00'/></div>
        <div style={S.col}><label style={S.label}>Called to date</label><input style={{...S.input,textAlign:'right'}} value={form.called_amount} onChange={e=>setForm(f=>({...f,called_amount:e.target.value}))} placeholder='0.00'/></div>
        <div style={S.col}><label style={S.label}>Commit date</label><input style={S.input} type='date' value={form.commit_date} onChange={e=>setForm(f=>({...f,commit_date:e.target.value}))}/></div></div>
      <div style={{marginBottom:12}}><label style={S.label}>Notes</label><input style={S.input} value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder='(optional)'/></div>
      {err&&<div style={S.err}>{err}</div>}<button style={S.btnP} onClick={add}>Add Commitment</button></div>}
    {loading?<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Loading…</div>:
     err&&!showAdd?<div style={S.err}>{err}</div>:
     !data||data.investors.length===0?<div style={{...S.card,textAlign:'center',padding:50,color:T.textDim}}>No commitments recorded yet.</div>:
     <div style={{...S.cardFlush,overflowX:'auto'}}><table style={S.table}><thead><tr>
       <th style={S.th}>Investor</th><th style={S.thR}>Commitment</th><th style={S.thR}>Called to Date</th><th style={S.thR}>Uncalled</th><th style={S.thR}>% Called</th><th style={S.thR}>Ownership %</th><th style={S.th}>Commit Date</th>{canEdit&&<th style={{...S.th,width:120}}>Actions</th>}</tr></thead>
       <tbody>{data.investors.map(i=>editId===i.id?(<tr key={i.id} style={{background:T.accentDim}}>
         <td style={S.td}>{i.investor}</td>
         <td style={S.tdR}><input style={{...S.input,textAlign:'right',padding:'4px 8px'}} value={editForm.commitment_amount} onChange={e=>setEditForm(f=>({...f,commitment_amount:e.target.value}))}/></td>
         <td style={S.tdR}><input style={{...S.input,textAlign:'right',padding:'4px 8px'}} value={editForm.called_amount} onChange={e=>setEditForm(f=>({...f,called_amount:e.target.value}))}/></td>
         <td style={{...S.tdR,color:T.textDim}}>{fmt((Number(editForm.commitment_amount)||0)-(Number(editForm.called_amount)||0))}</td>
         <td style={S.tdR}>—</td><td style={S.tdR}>—</td>
         <td style={S.td}><input style={{...S.input,padding:'4px 8px'}} type='date' value={editForm.commit_date} onChange={e=>setEditForm(f=>({...f,commit_date:e.target.value}))}/></td>
         <td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnGhost,color:T.green,fontSize:11}} onClick={saveEdit}>Save</button><button style={{...S.btnGhost,fontSize:11}} onClick={()=>setEditId(null)}>Cancel</button></div></td></tr>):(
       <tr key={i.id}><td style={{...S.td,fontWeight:600,color:T.textBright}}>{i.investor}{i.investor_code&&<span style={{color:T.textDim,fontWeight:400,marginLeft:6,fontSize:11}}>{i.investor_code}</span>}</td>
         <td style={S.tdR}>{fmt(i.commitment_amount)}</td><td style={S.tdR}>{fmt(i.called_amount)}</td>
         <td style={{...S.tdR,fontWeight:600,color:i.uncalled_amount>0.005?T.textBright:T.textDim}}>{fmt(i.uncalled_amount)}</td>
         <td style={S.tdR}>{pct(i.pct_called)}</td><td style={S.tdR}>{pct(i.ownership_pct)}</td><td style={S.td}>{i.commit_date||''}</td>
         {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6}}><button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(i)}>Edit</button><button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={()=>del(i)}>x</button></div></td>}</tr>))}
         <tr style={S.grandTotalRow}><td style={S.tdBold}>Total</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(data.totals.commitment_amount)}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(data.totals.called_amount)}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(data.totals.uncalled_amount)}</td><td style={S.tdR}></td><td style={{...S.tdR,fontWeight:700,color:T.textBright}}>100.00%</td><td style={S.td}></td>{canEdit&&<td style={S.td}></td>}</tr>
       </tbody></table></div>}
  </div>);
}

// ═══ Bank Reconciliation Report (QBO-style summary + detail, printable) ═══
function ReconciliationReportModal({entityId,rec,onClose}){
  const[data,setData]=useState(null);const[err,setErr]=useState('');const[loading,setLoading]=useState(true);
  useEffect(()=>{let alive=true;(async()=>{try{const d=await api.getReconciliationReport(entityId,rec.id);if(alive)setData(d);}catch(e){if(alive)setErr(e.message);}finally{if(alive)setLoading(false);}})();return()=>{alive=false;};},[entityId,rec.id]);
  const money=n=>{const v=Math.abs(Number(n)||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});return (Number(n)||0)<0?'-'+v:v;};
  const print=()=>{
    const el=document.getElementById('cl-recon-report');if(!el)return;
    const w=window.open('','_blank');if(!w)return;
    w.document.write('<html><head><title>Bank Reconciliation Report</title><style>'+
      'body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#111;margin:32px;}'+
      'h1{font-size:16px;margin:0 0 2px;}h2{font-size:13px;margin:20px 0 6px;border-bottom:1px solid #ccc;padding-bottom:3px;}'+
      '.muted{color:#555;font-size:11px;}table{width:100%;border-collapse:collapse;margin-top:4px;}'+
      'th,td{text-align:left;padding:4px 6px;font-size:11px;}th{border-bottom:1px solid #999;}'+
      'td.r,th.r{text-align:right;}tr.sub td{border-top:1px solid #999;font-weight:bold;}'+
      '.sumrow{display:flex;justify-content:space-between;max-width:520px;padding:3px 0;}'+
      '.sumrow.total{border-top:1px solid #999;font-weight:bold;margin-top:4px;padding-top:6px;}'+
      '</style></head><body>'+el.innerHTML+'</body></html>');
    w.document.close();w.focus();setTimeout(()=>{w.print();},250);
  };
  const csv=()=>{
    if(!data)return;
    const rows=[];
    rows.push(['Bank Reconciliation Report']);
    rows.push([data.entity_name]);
    rows.push([data.account_code+' '+data.account_name+', Period Ending '+data.statement_date]);
    rows.push(['Reconciled on',(data.reconciled_on||'').replace('T',' '),'Reconciled by',data.reconciled_by||'']);
    rows.push([]);
    rows.push(['Summary','USD']);
    const s=data.summary;
    rows.push(['Statement beginning balance',s.beginning_balance]);
    rows.push(['Checks and payments cleared ('+s.payments_count+')',s.payments_total]);
    rows.push(['Deposits and other credits cleared ('+s.deposits_count+')',s.deposits_total]);
    rows.push(['Statement ending balance',s.ending_balance]);
    rows.push(['Register balance as of '+data.statement_date,s.register_at_statement_date]);
    rows.push(['Cleared transactions after '+data.statement_date+' ('+s.cleared_after_count+')',s.cleared_after_total]);
    rows.push(['Uncleared transactions after '+data.statement_date+' ('+s.uncleared_after_count+')',s.uncleared_after_total]);
    rows.push(['Register balance as of report date',s.register_as_of_report]);
    const sec=(title,list)=>{rows.push([]);rows.push([title]);rows.push(['DATE','TYPE','REF NO.','PAYEE','AMOUNT (USD)']);list.forEach(l=>rows.push([l.date,l.type,l.ref_no,l.payee,l.amount]));};
    sec('Checks and payments cleared ('+data.payments_cleared.length+')',data.payments_cleared);
    sec('Deposits and other credits cleared ('+data.deposits_cleared.length+')',data.deposits_cleared);
    if(data.uncleared_through.length) sec('Uncleared transactions as of '+data.statement_date+' ('+data.uncleared_through.length+')',data.uncleared_through);
    const out=rows.map(r=>r.map(c=>{const v=c==null?'':String(c);return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;}).join(',')).join('\n');
    const blob=new Blob([out],{type:'text/csv'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);
    a.download=(data.account_code||'account')+' Bank Recon '+data.statement_date+'.csv';document.body.appendChild(a);a.click();a.remove();
  };
  const SumRow=({label,val,total})=> <div className={'sumrow'+(total?' total':'')} style={{display:'flex',justifyContent:'space-between',maxWidth:520,padding:total?'6px 0 3px':'3px 0',borderTop:total?'1px solid #999':'none',fontWeight:total?700:400,marginTop:total?4:0}}><span>{label}</span><span>{money(val)}</span></div>;
  const DetailTable=({list})=>(
    <table style={{width:'100%',borderCollapse:'collapse',marginTop:4}}><thead><tr>
      <th style={{textAlign:'left',padding:'4px 6px',borderBottom:'1px solid #999',fontSize:11}}>DATE</th>
      <th style={{textAlign:'left',padding:'4px 6px',borderBottom:'1px solid #999',fontSize:11}}>TYPE</th>
      <th style={{textAlign:'left',padding:'4px 6px',borderBottom:'1px solid #999',fontSize:11}}>REF NO.</th>
      <th style={{textAlign:'left',padding:'4px 6px',borderBottom:'1px solid #999',fontSize:11}}>PAYEE</th>
      <th style={{textAlign:'right',padding:'4px 6px',borderBottom:'1px solid #999',fontSize:11}}>AMOUNT (USD)</th>
    </tr></thead><tbody>
      {list.map((l,i)=><tr key={i}><td style={{padding:'4px 6px',fontSize:11}}>{l.date}</td><td style={{padding:'4px 6px',fontSize:11}}>{l.type}</td><td style={{padding:'4px 6px',fontSize:11}}>{l.ref_no}</td><td style={{padding:'4px 6px',fontSize:11}}>{l.payee}</td><td style={{padding:'4px 6px',fontSize:11,textAlign:'right'}}>{money(l.amount)}</td></tr>)}
      {list.length===0&&<tr><td colSpan={5} style={{padding:'8px 6px',fontSize:11,color:'#888'}}>None</td></tr>}
    </tbody></table>
  );
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:880,maxHeight:'92vh',display:'flex',flexDirection:'column'}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12,paddingRight:40}}>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright}}>Reconciliation Report</div>
      <div style={{display:'flex',gap:8}}>{data&&<><button style={S.btnS} onClick={csv}>Download CSV</button><button style={S.btnP} onClick={print}>Print / PDF</button></>}</div>
    </div>
    {loading&&<div style={{padding:40,textAlign:'center',color:T.textDim}}>Loading…</div>}
    {err&&<div style={{padding:16,color:T.red,fontSize:13}}>{err}</div>}
    {data&&<div style={{overflowY:'auto'}}><div id="cl-recon-report" style={{background:'#fff',color:'#111',padding:24,borderRadius:8,border:'1px solid '+T.border}}>
      <h1 style={{fontSize:16,margin:'0 0 2px'}}>{data.entity_name}</h1>
      <div style={{fontWeight:600}}>{data.account_code} {data.account_name}, Period Ending {data.statement_date}</div>
      <div style={{fontSize:15,fontWeight:700,margin:'10px 0 2px'}}>RECONCILIATION REPORT</div>
      <div className="muted" style={{color:'#555',fontSize:11}}>Reconciled on: {(data.reconciled_on||'').replace('T',' ').slice(0,19)}</div>
      <div className="muted" style={{color:'#555',fontSize:11}}>Reconciled by: {data.reconciled_by}</div>
      <div className="muted" style={{color:'#555',fontSize:11,marginTop:4}}>Any changes made to transactions after this date aren't included in this report.</div>

      <h2 style={{fontSize:13,margin:'20px 0 6px',borderBottom:'1px solid #ccc',paddingBottom:3}}>Summary <span style={{float:'right'}}>USD</span></h2>
      <SumRow label="Statement beginning balance" val={data.summary.beginning_balance}/>
      <SumRow label={'Checks and payments cleared ('+data.summary.payments_count+')'} val={data.summary.payments_total}/>
      <SumRow label={'Deposits and other credits cleared ('+data.summary.deposits_count+')'} val={data.summary.deposits_total}/>
      <SumRow label="Statement ending balance" val={data.summary.ending_balance} total/>
      <div style={{height:10}}/>
      <SumRow label={'Register balance as of '+data.statement_date} val={data.summary.register_at_statement_date}/>
      <SumRow label={'Cleared transactions after '+data.statement_date+' ('+data.summary.cleared_after_count+')'} val={data.summary.cleared_after_total}/>
      <SumRow label={'Uncleared transactions after '+data.statement_date+' ('+data.summary.uncleared_after_count+')'} val={data.summary.uncleared_after_total}/>
      <SumRow label="Register balance as of report date" val={data.summary.register_as_of_report} total/>

      <h2 style={{fontSize:13,margin:'20px 0 6px',borderBottom:'1px solid #ccc',paddingBottom:3}}>Details</h2>
      <div style={{fontWeight:700,fontSize:12,marginTop:8}}>Checks and payments cleared ({data.payments_cleared.length})</div>
      <DetailTable list={data.payments_cleared}/>
      <div style={{fontWeight:700,fontSize:12,marginTop:16}}>Deposits and other credits cleared ({data.deposits_cleared.length})</div>
      <DetailTable list={data.deposits_cleared}/>
      {data.uncleared_through.length>0&&<><div style={{fontWeight:700,fontSize:12,marginTop:16}}>Uncleared transactions as of {data.statement_date} ({data.uncleared_through.length})</div>
      <DetailTable list={data.uncleared_through}/></>}
    </div></div>}
  </div></div>);
}

// ═══ Bank Reconciliation ═══
function BankReconciliation({entityId,user,canEdit=true}){const[accounts,setAccounts]=useState([]);const[entries,setEntries]=useState([]);const[recs,setRecs]=useState([]);
  const[view,setView]=useState('list');const[selAcct,setSelAcct]=useState('');const[stmtDate,setStmtDate]=useState(today());const[stmtBal,setStmtBal]=useState('');
  const[cleared,setCleared]=useState({});const[checked,setChecked]=useState({});
  const[viewEntry,setViewEntry]=useState(null);
  const[reportRec,setReportRec]=useState(null);
  const load=useCallback(async()=>{const[a,e,r]=await Promise.all([api.getAccounts(entityId),api.getEntries(entityId),api.getReconciliations(entityId)]);setAccounts(a);setEntries(e);setRecs(r);},[entityId]);
  useEffect(()=>{load();},[load]);
  const bankAccts=accounts.filter(a=>a.bank_acct||(['cash','bank','checking','savings'].some(w=>a.name.toLowerCase().includes(w))&&a.type==='Asset'));
  useEffect(()=>{if(selAcct)api.getCleared(entityId,selAcct).then(setCleared);else setCleared({});},[selAcct,entityId]);
  const getTxns=code=>{const txns=[];entries.forEach(e=>{e.lines.forEach((l,li)=>{if(l.account_code===code){const acct=accounts.find(a=>a.code===code);const isDr=acct?.type==='Asset'||acct?.type==='Expense';txns.push({jeId:e.id,jeNum:e.entry_num,lineIdx:li,date:e.date,memo:e.memo,amount:isDr?(l.debit-l.credit):(l.credit-l.debit),debit:l.debit,credit:l.credit,key:e.id+'-'+li});}});});txns.sort((a,b)=>a.date.localeCompare(b.date));return txns;};
  // Reconciliation math is as-of the statement date: only transactions dated on or
  // before the statement date participate (book balance, uncleared list, outstanding
  // items). Later-dated activity belongs to the next reconciliation.
  const txnsAll=selAcct?getTxns(selAcct):[];const txns=stmtDate?txnsAll.filter(t=>t.date<=stmtDate):txnsAll;
  const uncl=txns.filter(t=>!cleared[t.key]);const bookBal=txns.reduce((s,t)=>s+t.amount,0);const stmtNum=parseFloat(stmtBal)||0;
  const outDep=uncl.filter(t=>!checked[t.key]&&t.amount>0).reduce((s,t)=>s+t.amount,0);const outPay=uncl.filter(t=>!checked[t.key]&&t.amount<0).reduce((s,t)=>s+t.amount,0);
  const diff=bookBal-(stmtNum+outDep+outPay);const isRec=Math.abs(diff)<0.005&&stmtNum!==0;
  const finalize=async()=>{if(!isRec)return;const inScope=new Set(uncl.map(t=>t.key));await api.createReconciliation(entityId,{account_code:selAcct,statement_date:stmtDate,statement_balance:stmtNum,book_balance:bookBal,cleared_keys:Object.keys(checked).filter(k=>checked[k]&&inScope.has(k))});setChecked({});setStmtBal('');setView('list');load();};
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
      (()=>{
        // Latest reconciliation per account (by statement_date, then id) — only that
        // one can be undone; later recs depend on the cleared state left by earlier ones.
        const latestByAcct={};recs.forEach(r=>{const cur=latestByAcct[r.account_code];if(!cur||r.statement_date>cur.statement_date||(r.statement_date===cur.statement_date&&r.id>cur.id))latestByAcct[r.account_code]=r;});
        const undoRec=async r=>{if(!confirm('Undo the '+r.statement_date+' reconciliation for account '+r.account_code+'?\n\nIts '+r.cleared_count+' cleared item(s) will return to uncleared and reappear in the next reconciliation. No journal entries are changed.'))return;
          try{await api.deleteReconciliation(entityId,r.id);load();}catch(ex){alert(ex.message);}};
        return(<table style={S.table}><thead><tr><th style={S.th}>Date</th><th style={S.th}>Account</th><th style={S.thR}>Statement</th><th style={S.thR}>Book</th><th style={S.thR}>Cleared</th><th style={S.th}>By</th><th style={S.th}></th></tr></thead>
        <tbody>{recs.map(r=><tr key={r.id}><td style={S.td}>{r.statement_date}</td><td style={S.td}>{(()=>{const a=accounts.find(x=>x.code===r.account_code);return a?acctLabel(a.code,a.name):r.account_code;})()}</td><td style={S.tdR}>${fmt(r.statement_balance)}</td><td style={S.tdR}>${fmt(r.book_balance)}</td><td style={S.tdR}>{r.cleared_count}</td><td style={S.td}>{r.completed_by}</td><td style={S.td}><div style={{display:'flex',gap:6,justifyContent:'flex-end'}}>
          <button style={{...S.btnS,padding:'4px 12px',fontSize:11}} onClick={()=>setReportRec(r)}>Report</button>
          {canEdit&&latestByAcct[r.account_code]?.id===r.id&&<button style={{...S.btnS,padding:'4px 12px',fontSize:11,color:T.red,borderColor:T.red+'40'}} title="Undo this reconciliation — cleared items return to uncleared" onClick={()=>undoRec(r)}>Undo</button>}
        </div></td></tr>)}</tbody></table>);})()}</div>
    {reportRec&&<ReconciliationReportModal entityId={entityId} rec={reportRec} onClose={()=>setReportRec(null)}/>}
    </div>);}

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
      // Detect the Cost Code # and Cost Code Name columns by HEADER text. Layouts
      // differ: SRN/Silsbee put the name in col F (index 5), but HP/Braker put it
      // in col D (index 3) with the Bill # in col F — hard-coding index 5 there
      // reads the bill number as the cost-code name (e.g. 349979). Fall back to the
      // SRN positions (code=2, name=5) only if the headers aren't found.
      let codeIdx=2,nameIdx=5,hdrRow=-1;
      for(let i=0;i<Math.min(rows.length,8);i++){
        const cells=(rows[i]||[]).map(c=>String(c==null?'':c).toLowerCase().replace(/\s+/g,' ').trim());
        const ci=cells.findIndex(t=>/cost code\s*#|cost code\s*(number|no)\b|^cost code$/.test(t));
        const ni=cells.findIndex(t=>/cost code name/.test(t));
        if(ci>=0&&ni>=0){codeIdx=ci;nameIdx=ni;hdrRow=i;break;}
      }
      const m={};
      for(let i=(hdrRow>=0?hdrRow+1:0);i<rows.length;i++){
        const row=rows[i];
        const code=row&&row[codeIdx]!=null?String(row[codeIdx]).trim():'';
        const name=row&&row[nameIdx]!=null?String(row[nameIdx]).trim():'';
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
          // Prefer the name from THIS workbook's cost-code catalog (authoritative
          // for the requisition) over the server prediction, whose learned history
          // may carry a mis-columned name for templates like HP/Braker.
          cost_code_name:(r.cost_code&&wbCoaMap[String(r.cost_code).trim()])||r.cost_code_name||'',
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
      const {blob,filename,summary,failedChecks,workpaperFolder,workpaperSaved,packetFileId,packetFileName,forced,devFee}=await api.rollForwardRequisition(entityId,rfFile,newCurrent,{reqNumber:rfReqNum,asOfDate:rfAsOf,invoices,force});
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
      setReqState(cur=>({...cur,cards:[],file:null,reqNum:'',detail:null,result:{filename,summary,failedChecks,count:newCurrent.length,workpaperFolder,workpaperSaved,forced,devFee}}));
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
        {rfResult.devFee&&(rfResult.devFee.needs_review?
          <div style={{marginTop:12,padding:'10px 12px',borderRadius:8,background:T.orangeDim,border:'1px solid '+T.orange+'40'}}>
            <div style={{fontSize:11,fontWeight:700,color:T.orange,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:4}}>Development fee — manual entry needed</div>
            <div style={{fontSize:12,color:T.text}}>{rfResult.devFee.note||'CloudLedger could not confirm this project\u2019s dev-fee method from the prior report. Enter the development fee for this period by hand.'}{rfResult.devFee.prior&&rfResult.devFee.prior.fee!=null&&<span style={{color:T.textMuted}}> Prior period: {Number(rfResult.devFee.prior.fee).toLocaleString(undefined,{style:'currency',currency:'USD'})} on {Number(rfResult.devFee.prior.base).toLocaleString(undefined,{style:'currency',currency:'USD'})} of costs.</span>}</div>
          </div>
        :rfResult.devFee.amount!=null?
          <div style={{marginTop:12,padding:'10px 12px',borderRadius:8,background:T.bgElevated,border:'1px solid '+T.greenBorder}}>
            <div style={{fontSize:11,fontWeight:700,color:T.green,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:4}}>Development fee added</div>
            <div style={{fontSize:13,color:T.textBright,fontWeight:600}}>{Number(rfResult.devFee.amount).toLocaleString(undefined,{style:'currency',currency:'USD'})}{rfResult.devFee.rate_text?<span style={{fontWeight:400,color:T.text}}> &mdash; {rfResult.devFee.rate_text}</span>:null}</div>
            <div style={{fontSize:11,color:T.textMuted,marginTop:3}}>
              {rfResult.devFee.base!=null&&<>Base: {Number(rfResult.devFee.base).toLocaleString(undefined,{style:'currency',currency:'USD'})} of new costs. </>}
              Method {rfResult.devFee.source==='claude'?'inferred by Claude':'read from the prior report\u2019s formulas'}{rfResult.devFee.validated?', matched the prior period':''}.
            </div>
          </div>
        :null)}
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
  const[name,setName]=useState('');const[newType,setNewType]=useState('accounting');const[newDisplayId,setNewDisplayId]=useState('');const[bulkText,setBulkText]=useState('');const[bulkType,setBulkType]=useState('accounting');const[bulkBusy,setBulkBusy]=useState(false);const[err,setErr]=useState('');
  const[typeBusy,setTypeBusy]=useState(null);// entity id whose type is being toggled
  const[openType,setOpenType]=useState({accounting:false,development:false,shell:false});
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
      const s=r.suggested||{};setGlMap({account_number:s.account_number||'',account_name:s.account_name||'',transaction_date:s.transaction_date||'',description:s.description||'',memo:s.memo||'',debit:s.debit||'',credit:s.credit||'',reference:s.reference||'',running_balance:s.running_balance||'',project:s.project||'',class:s.class||'',location:s.location||''});
      setGlFused(!!s.fused);setGlFusedDelim('auto');setGlStep('map');}
    catch(ex){setGlErr(ex.message);}finally{setGlBusy(false);}};
  const runGlImport=async()=>{if(!glFile||!glEntity)return;setGlBusy(true);setGlErr('');setGlUnbalanced(null);
    const mapping={...glMap,fused:glFused,fused_column:glFused?(glMap.account_number||glPreview?.suggested?.fused_column):null,fused_delimiter:glFusedDelim==='auto'?null:glFusedDelim};
    try{const r=await api.importGL(glEntity,glFile,mapping);setGlResult(r);setGlStep('done');}
    catch(ex){setGlErr(ex.message);if(ex.detail&&ex.detail.unbalanced_groups)setGlUnbalanced(ex.detail);}finally{setGlBusy(false);}};
  const GL_FIELDS=[{k:'account_number',label:'Account Number',req:true},{k:'account_name',label:'Account Name',req:true},{k:'transaction_date',label:'Transaction Date',req:true},{k:'description',label:'Description',req:false},{k:'memo',label:'Memo',req:false},{k:'debit',label:'Debit',req:true},{k:'credit',label:'Credit',req:true},{k:'reference',label:'Reference / Doc # (groups lines into JEs)',req:false},{k:'running_balance',label:'Running Balance (verification only, not stored)',req:false},{k:'project',label:'Project (Intacct project / QBO class — tracked per line)',req:false},{k:'class',label:'Class (e.g. investor — tracked per line)',req:false},{k:'location',label:'Location (e.g. deal / asset — tracked per line)',req:false}];
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
    {bulk&&(()=>{const TYPE_ALIASES={accounting:'accounting',acct:'accounting',acc:'accounting',development:'development','development project':'development',dev:'development',devproject:'development',shell:'shell'};
      const parseBulk=()=>bulkText.split('\n').map(l=>l.trim()).filter(Boolean).map(line=>{
        let name=line,type=null;
        const parts=line.includes('\t')?line.split('\t'):line.split(',');
        if(parts.length>1){const last=parts[parts.length-1].trim().toLowerCase();
          if(TYPE_ALIASES[last]){type=TYPE_ALIASES[last];name=parts.slice(0,-1).join(',').trim();}}
        return{name,type:type||bulkType};});
      const rows=parseBulk();
      return(<div style={{...S.card,borderColor:T.accent+'40'}}><div style={{...S.h2,marginBottom:8}}>Bulk Import Entities</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>One entity per line. Optionally add a type column after a comma or tab &mdash; <span style={{fontFamily:'monospace'}}>accounting</span>, <span style={{fontFamily:'monospace'}}>development</span>, or <span style={{fontFamily:'monospace'}}>shell</span>. Lines without a type use the default below.<br/><span style={{fontFamily:'monospace',fontSize:11}}>e.g.&nbsp; CLR Fund II LP, accounting&nbsp;&nbsp;|&nbsp;&nbsp;Sabine Yard Expansion, development</span></div>
      <textarea style={{...S.input,height:160,fontFamily:'monospace',fontSize:12,resize:'vertical'}} value={bulkText} onChange={e=>setBulkText(e.target.value)}/>
      <div style={{display:'flex',gap:12,alignItems:'center',marginTop:10,flexWrap:'wrap'}}>
        <label style={{...S.label,marginBottom:0}}>Default type</label>
        <select style={{...S.inputSm,width:'auto'}} value={bulkType} onChange={e=>setBulkType(e.target.value)}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option></select>
        {rows.length>0&&<span style={{fontSize:11,color:T.textMuted}}>{rows.length} entit{rows.length===1?'y':'ies'}: {['accounting','development','shell'].map(t=>({t,n:rows.filter(r=>r.type===t).length})).filter(x=>x.n>0).map(x=>x.n+' '+x.t).join(', ')}</span>}
      </div>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:10}} disabled={bulkBusy} onClick={async()=>{if(!rows.length){setErr('None');return;}setBulkBusy(true);setErr('');try{for(const r of rows)await api.createEntity(r.name,r.type);setBulkText('');setBulk(false);refresh();}catch(e){setErr(e.message);}finally{setBulkBusy(false);}}}>{bulkBusy?'Importing...':'Import'}</button></div>);})()}
    <div style={{...S.cardFlush,overflowX:'auto'}}><table style={{...S.table,minWidth:1180}}><thead><tr><th style={{...S.th,minWidth:240}}>Entity</th><th style={{...S.th,width:760,minWidth:760}}>Actions</th></tr></thead>
      <tbody>{ENTITY_TYPES.map(t=>{const grp=entities.filter(e=>entTypeOf(e)===t.key).sort((a,b)=>a.name.localeCompare(b.name));const isOpen=openType[t.key];return(<Fragment key={t.key}>
        <tr style={{cursor:'pointer',background:T.bgElevated,borderTop:'2px solid '+T.border}} onClick={()=>setOpenType(o=>({...o,[t.key]:!o[t.key]}))}>
          <td colSpan={2} style={{...S.td,fontWeight:700,color:T.textBright}}><span style={{marginRight:6,fontSize:12,color:T.textMuted}}>{isOpen?'▾':'▸'}</span><span style={{marginRight:8}}>{t.icon}</span>{t.label}<span style={{marginLeft:8,fontSize:11,fontWeight:600,color:T.textMuted}}>({grp.length})</span></td>
        </tr>
        {isOpen&&grp.length===0&&<tr><td colSpan={2} style={{...S.td,color:T.textMuted,padding:'10px 20px 10px 44px',fontSize:12}}>No {t.label.toLowerCase()} entities.</td></tr>}
        {isOpen&&grp.map(e=><tr key={e.id} style={e.id===activeEntity?{background:T.accentDim}:{}}>
        <td style={{...S.td,fontWeight:600,color:T.textBright,paddingLeft:32}}>{e.display_id&&<span style={{marginRight:8,fontSize:11,fontWeight:700,color:T.textMuted,fontFamily:'monospace'}}>{e.display_id}</span>}{e.name}{e.entity_type==='development'&&<span style={{marginLeft:8,fontSize:9,fontWeight:700,color:T.green,background:T.greenDim,border:'1px solid '+T.greenBorder,borderRadius:4,padding:'2px 6px',textTransform:'uppercase',letterSpacing:'0.05em',verticalAlign:'middle'}}>Dev Project</span>}{e.entity_type==='shell'&&<span style={{marginLeft:8,fontSize:9,fontWeight:700,color:T.teal,background:T.tealDim,border:'1px solid '+T.teal+'40',borderRadius:4,padding:'2px 6px',textTransform:'uppercase',letterSpacing:'0.05em',verticalAlign:'middle'}}>Shell</span>}</td>
        <td style={S.td}><div style={{display:'flex',gap:8,flexWrap:'nowrap',whiteSpace:'nowrap'}}>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0}} onClick={()=>setActiveEntity(e.id)}>Select</button>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0,color:T.accent,borderColor:T.accent+'40'}} onClick={()=>{setImporting(e.id);setImportMsg('');setImportErr('');}}>Import Trial Balance</button>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0,color:T.accent,borderColor:T.accent+'40'}} onClick={()=>{resetGl();setGlEntity(e.id);}}>Import General Ledger Detail</button>
          <select style={{...S.inputSm,padding:'5px 8px',fontSize:11,flexShrink:0,width:'auto'}} disabled={typeBusy===e.id} title="Entity type" value={e.entity_type||'accounting'} onChange={async(ev)=>{const next=ev.target.value;if(next===e.entity_type)return;if(!confirm('Set "'+e.name+'" to '+({accounting:'Accounting',development:'Development Project',shell:'Shell'}[next]||next)+'?'))return;setTypeBusy(e.id);try{await api.updateEntity(e.id,{entity_type:next});await refresh();}catch(ex){alert(ex.message);}finally{setTypeBusy(null);}}}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option></select>
          <button style={{...S.btnS,padding:'5px 12px',fontSize:11,flexShrink:0}} title="Set the short Entity ID used as the invoice-packet filename prefix" onClick={async()=>{const cur=e.display_id||'';const v=prompt('Entity ID for "'+e.name+'"\n(used as the invoice-packet filename prefix; leave blank to use the entity name):',cur);if(v===null)return;try{await api.updateEntity(e.id,{display_id:v.trim()});await refresh();}catch(ex){alert(ex.message);}}}>Edit ID</button>
          <button style={{...S.btnD,padding:'5px 12px',fontSize:11,flexShrink:0}} onClick={async()=>{if(!confirm('Delete entity '+e.name+' and all its data?'))return;await api.deleteEntity(e.id);const r=await refresh();if(activeEntity===e.id)setActiveEntity(r[0]?.id||null);}}>Delete</button>
        </div></td></tr>)}
      </Fragment>);})}</tbody></table></div>
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
        <div style={{overflow:'auto',maxHeight:320,marginBottom:14,border:'1px solid '+T.borderLight,borderRadius:6}}>
          <table style={{...S.table,fontSize:10,minWidth:1100,width:'max-content'}}><thead><tr>{glPreview.columns.map(c=><th key={c} style={{...S.th,fontSize:10,whiteSpace:'nowrap',padding:'5px 8px',position:'sticky',top:0,background:T.bgElevated,zIndex:1}}>{c}</th>)}</tr></thead>
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
  const[accessGroups,setAccessGroups]=useState([]); // groups the user belongs to (with entity_ids)
  const[accessEffective,setAccessEffective]=useState(null); // null=all, else union of individual+group
  const openAccess=async(u)=>{
    setAccessUser(u);setAccessErr('');setAccessSaving(false);setAccessGroups([]);setAccessEffective(null);
    try{
      const[ents,acc]=await Promise.all([api.getEntities(),api.getUserEntityAccess(u.id)]);
      setAccessAllEntities(ents);
      setAccessEntities(acc.entity_ids||[]);
      setAccessGroups(acc.groups||[]);
      setAccessEffective(acc.effective===undefined?null:acc.effective);
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
  // ── User groups (e.g. CLA): bundle members + grant entity access at once ──
  const[groups,setGroups]=useState([]);
  const[groupModal,setGroupModal]=useState(null);
  const[gMembers,setGMembers]=useState([]);const[gEntities,setGEntities]=useState([]);const[gAllEntities,setGAllEntities]=useState([]);
  const[gSaving,setGSaving]=useState(false);const[gErr,setGErr]=useState('');
  const loadGroups=useCallback(()=>{api.getGroups().then(setGroups).catch(()=>{});},[]);
  useEffect(()=>{loadGroups();},[loadGroups]);
  const openGroup=async(g)=>{setGroupModal(g);setGErr('');setGSaving(false);try{const[detail,ents]=await Promise.all([api.getGroup(g.id),api.getEntities()]);setGMembers(detail.member_ids||[]);setGEntities(detail.entity_ids||[]);setGAllEntities(ents);}catch(e){setGErr(e.message);}};
  const toggleGMember=(id)=>setGMembers(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const toggleGEntity=(id)=>setGEntities(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const saveGroup=async()=>{setGSaving(true);setGErr('');try{await api.setGroupMembers(groupModal.id,gMembers);await api.setGroupEntities(groupModal.id,gEntities);setGroupModal(null);loadGroups();}catch(e){setGErr(e.message);}setGSaving(false);};
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
    {groups.length>0&&<div style={{...S.card,marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:4}}>User Groups</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:12}}>Grant entity access to a whole team at once. Everyone in a group can access all entities assigned to that group (in addition to any individual access). Adding someone to a group limits them to the group's entities.</div>
      <table style={S.table}><thead><tr><th style={S.th}>Group</th><th style={S.th}>Members</th><th style={S.th}>Entities</th><th style={{...S.th,width:120}}></th></tr></thead>
      <tbody>{groups.map(g=><tr key={g.id}>
        <td style={{...S.td,fontWeight:600,color:T.textBright}}>{g.name}</td>
        <td style={S.td}>{g.member_count}</td>
        <td style={S.td}>{g.entity_count}</td>
        <td style={S.td}><button style={{...S.btnS,padding:'5px 12px',fontSize:11}} onClick={()=>openGroup(g)}>Manage</button></td>
      </tr>)}</tbody></table>
    </div>}
    <div style={S.cardFlush}>
      <table style={S.table}><thead><tr>
        <th style={S.th}>Name</th>
        <th style={S.th}>Login Email</th>
        <th style={S.th}>Role</th>
        <th style={{...S.th,width:240}}>Actions</th></tr></thead>
      <tbody>{users.length===0&&!loadErr?<tr><td colSpan={4} style={{...S.td,textAlign:'center',padding:40,color:T.textDim}}>No users found</td></tr>:
        users.map(u=><tr key={u.id}>
          <td style={{...S.td,fontWeight:600,color:T.textBright}}>{titleName(u.name)}{u.id===currentUser.id?<span style={{color:T.accent,fontSize:10,marginLeft:8,fontWeight:500}}>(you)</span>:''}</td>
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
      <div style={{fontSize:13,color:T.textMuted,marginBottom:14}}>User: <strong style={{color:T.textBright}}>{titleName(accessUser.name)}</strong> ({accessUser.role})</div>
      {accessGroups.length>0&&<div style={{fontSize:12,color:T.textBright,marginBottom:8,padding:'8px 12px',background:T.accent+'12',border:'1px solid '+T.accent+'33',borderRadius:6}}>
        Member of {accessGroups.map(g=>g.name).join(', ')} — also has access to {new Set(accessGroups.flatMap(g=>g.entity_ids)).size} entit{new Set(accessGroups.flatMap(g=>g.entity_ids)).size===1?'y':'ies'} via group{accessGroups.length>1?'s':''} (marked "via …" below). Manage group entities in User Groups.
      </div>}
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10,padding:'8px 12px',background:T.bgInset,borderRadius:6}}>
        {accessEffective===null
          ? 'This user can currently access ALL entities (no individual or group restrictions). Check entities below to restrict to only those.'
          : 'Effective access: '+accessEffective.length+' entit'+(accessEffective.length===1?'y':'ies')+' (individual grants + groups). The checkboxes below set this user’s INDIVIDUAL access, added on top of any group access.'}
      </div>
      <div style={{maxHeight:320,overflowY:'auto',border:'1px solid '+T.border,borderRadius:6,marginBottom:12}}>
        {accessAllEntities.map(e=>{
          const viaGroups=accessGroups.filter(g=>g.entity_ids.includes(e.id)).map(g=>g.name);
          const inGroup=viaGroups.length>0;
          const checked=inGroup||accessEntities.includes(e.id);
          return (
          <label key={e.id} title={inGroup?'Granted via '+viaGroups.join(', ')+' — manage in User Groups':''} style={{display:'flex',alignItems:'center',padding:'8px 12px',borderBottom:'1px solid '+T.border,cursor:inGroup?'default':'pointer',gap:10}}>
            <input type="checkbox" checked={checked} disabled={inGroup} onChange={()=>{if(!inGroup)toggleAccessEntity(e.id);}}/>
            <span style={{color:T.textBright,fontSize:13}}>{e.name}</span>
            {e.code&&<span style={{color:T.textMuted,fontSize:11,fontFamily:'monospace'}}>{e.code}</span>}
            {inGroup&&<span style={{marginLeft:'auto',fontSize:10,color:T.accent,background:T.accent+'18',padding:'2px 6px',borderRadius:4,whiteSpace:'nowrap'}}>via {viaGroups.join(', ')}</span>}
          </label>
          );
        })}
        {accessAllEntities.length===0&&<div style={{padding:16,color:T.textMuted,textAlign:'center'}}>No entities</div>}
      </div>
      {accessErr&&<div style={S.err}>{accessErr}</div>}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
        <button style={S.btnGhost} onClick={()=>setAccessEntities([])} disabled={accessSaving}>{accessGroups.length>0?'Clear individual grants':'Clear (= all access)'}</button>
        <div style={{display:'flex',gap:8}}>
          <button style={S.btnGhost} onClick={()=>setAccessUser(null)} disabled={accessSaving}>Cancel</button>
          <button style={S.btnP} onClick={saveAccess} disabled={accessSaving}>{accessSaving?'Saving...':'Save'}</button>
        </div>
      </div>
    </div></div>}
    {groupModal&&<div style={S.modal} onClick={()=>setGroupModal(null)}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:640}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setGroupModal(null)}>&times;</button>
      <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:6}}>{groupModal.name} Group</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:14}}>Everyone checked under Members gets access to every entity checked under Entities (plus any individual access they already have).</div>
      <div style={{display:'flex',gap:16}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:12,fontWeight:600,color:T.textBright,marginBottom:6}}>Members ({gMembers.length})</div>
          <div style={{maxHeight:300,overflowY:'auto',border:'1px solid '+T.border,borderRadius:6}}>
            {users.filter(u=>u.role!=='Admin').map(u=>(
              <label key={u.id} style={{display:'flex',alignItems:'center',padding:'7px 10px',borderBottom:'1px solid '+T.border,cursor:'pointer',gap:8}}>
                <input type="checkbox" checked={gMembers.includes(u.id)} onChange={()=>toggleGMember(u.id)}/>
                <span style={{color:T.textBright,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{titleName(u.name)}</span>
              </label>
            ))}
            {users.filter(u=>u.role!=='Admin').length===0&&<div style={{padding:12,color:T.textMuted,fontSize:12,textAlign:'center'}}>No non-admin users</div>}
          </div>
          <div style={{fontSize:10,color:T.textMuted,marginTop:4}}>Admins already have all-entity access, so they aren't listed.</div>
        </div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:12,fontWeight:600,color:T.textBright,marginBottom:6}}>Entities ({gEntities.length})</div>
          <div style={{maxHeight:300,overflowY:'auto',border:'1px solid '+T.border,borderRadius:6}}>
            {gAllEntities.map(e=>(
              <label key={e.id} style={{display:'flex',alignItems:'center',padding:'7px 10px',borderBottom:'1px solid '+T.border,cursor:'pointer',gap:8}}>
                <input type="checkbox" checked={gEntities.includes(e.id)} onChange={()=>toggleGEntity(e.id)}/>
                <span style={{color:T.textBright,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{e.name}</span>
                {e.code&&<span style={{color:T.textMuted,fontSize:10,fontFamily:'monospace'}}>{e.code}</span>}
              </label>
            ))}
            {gAllEntities.length===0&&<div style={{padding:12,color:T.textMuted,fontSize:12,textAlign:'center'}}>No entities</div>}
          </div>
        </div>
      </div>
      {gErr&&<div style={{...S.err,marginTop:10}}>{gErr}</div>}
      <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:14}}>
        <button style={S.btnGhost} onClick={()=>setGroupModal(null)} disabled={gSaving}>Cancel</button>
        <button style={S.btnP} onClick={saveGroup} disabled={gSaving}>{gSaving?'Saving...':'Save'}</button>
      </div>
    </div></div>}
        {resetId&&<div style={S.modal} onClick={()=>setResetId(null)}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:400,textAlign:'center'}} onClick={e=>e.stopPropagation()}>
      <button style={S.modalClose} onClick={()=>setResetId(null)}>&times;</button><div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:20}}>Reset Password</div>
      <div style={{fontSize:13,color:T.textMuted,marginBottom:6}}>User: <strong style={{color:T.textBright}}>{titleName(users.find(u=>u.id===resetId)?.name)}</strong></div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:16,fontFamily:'monospace'}}>{users.find(u=>u.id===resetId)?.email}</div>
      <input style={S.input} type="password" placeholder="New password" value={resetPw} onChange={e=>{setResetPw(e.target.value);setResetMsg('');}}/>
      {resetMsg&&<div style={{fontSize:12,marginTop:8,color:resetMsg.includes('!')?T.green:T.red}}>{resetMsg}</div>}
      <button style={{...S.btnP,width:'100%',padding:11,marginTop:12}} onClick={async()=>{if(resetPw.length<3){setResetMsg('Min 3 chars');return;}try{await api.adminResetPassword(resetId,resetPw);setResetMsg('Password reset!');setTimeout(()=>setResetId(null),1500);}catch(e){setResetMsg(e.message);}}}>Reset Password</button>
    </div></div>}</div>);}
