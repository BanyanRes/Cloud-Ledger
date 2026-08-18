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
// Local calendar date as YYYY-MM-DD. Uses local date parts (NOT toISOString(),
// which converts to UTC and can roll a user in a far-from-UTC timezone — e.g.
// UTC+8 Philippines — onto the previous/next calendar day, causing reports run
// with the same "as of" default to differ from a US user's).
const today = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
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
  { key:'operating',   label:'Operating',   icon:'🏢' },
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
// Active entity's display id (its "entity id"), or its name when there's no
// display id — prepended to exported Excel filenames.
let _activeEntityFileTag = '';
const CLASS_DIM_LABELS = { TURNKEYR: 'Pay Application' };
const classTerm = () => CLASS_DIM_LABELS[_activeEntityCode] || 'Class';
function exportToExcel(data, fn, opts) { opts = opts || {}; const moneyFmt = opts.numFmt || '#,##0.00;(#,##0.00)'; const plain = new Set(opts.plainCols || []); const ws = XLSX.utils.aoa_to_sheet(data); const range = XLSX.utils.decode_range(ws['!ref']); for (let R = range.s.r; R <= range.e.r; R++) { for (let C = range.s.c; C <= range.e.c; C++) { if (plain.has(C)) continue; const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })]; if (cell && cell.t === 'n') cell.z = moneyFmt; } } const fmtLen = v => (typeof v === 'number' ? v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).length : String(v == null ? '' : v).length); const nCols = data.reduce((m, r) => Math.max(m, (r ? r.length : 0)), 0); const colW = []; for (let c = 0; c < nCols; c++) { let w = 8; for (const r of data) { if (r && r[c] != null && r[c] !== '') w = Math.max(w, fmtLen(r[c])); } colW.push({ wch: Math.min(w + 2, 60) }); } ws['!cols'] = colW; if (opts.formulas) { for (const g of (opts.formulas || [])) { if (!g || !g.f) continue; const addr = XLSX.utils.encode_cell({ r: g.r, c: g.c }); const prev = ws[addr]; const cell = { t: 'n', f: g.f, z: moneyFmt }; if (prev && prev.v != null && !isNaN(prev.v)) cell.v = prev.v; ws[addr] = cell; } } const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); const _pfx = String(_activeEntityFileTag || '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, ''); const _out = (_pfx && fn.indexOf(_pfx + '_') !== 0) ? (_pfx + '_' + fn) : fn; XLSX.writeFile(wb, _out); }
// ── Excel formula helpers (used by every report's Export Excel) ──
// 0-based column index → A1 letter (0→A, 1→B, … 26→AA).
const XLC = n => { let s = ''; n = n + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
// Push SUM() formulas onto total row `tr` (0-based) for each 0-based column in
// `dcols`, summing detail rows `first..last` (0-based, inclusive). No-op if empty.
const sumCols = (F, tr, dcols, first, last) => { if (last < first || tr == null) return; for (const c of dcols) F.push({ r: tr, c, f: 'SUM(' + XLC(c) + (first + 1) + ':' + XLC(c) + (last + 1) + ')' }); };
// Push SUM() formulas that add a specific, explicit set of rows (e.g. a grand total
// that sums subtotal rows rather than a contiguous block). `rows` = 0-based row idxs.
const sumRows = (F, tr, dcols, rows) => { if (!rows || !rows.length || tr == null) return; for (const c of dcols) { const L = XLC(c); F.push({ r: tr, c, f: rows.map(r => L + (r + 1)).join('+') }); } };
const BLANK_JE = () => ({date:today(),memo:'',lines:[{account_code:'',debit:'',credit:'',description:''},{account_code:'',debit:'',credit:'',description:''}]});
const SIDEBAR_KEY = 'cl_sidebar';

// ─── Cloud Ledger Logo SVG ───
function Logo({size=32}) {
  return (<svg width={size} height={size} viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
    {/* Cloud outline on stacked ledger layers */}
    <path d="M7 30 L20 36 L33 30" fill="none" stroke="#0f172a" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
    <path d="M7 24.5 L20 30.5 L33 24.5" fill="none" stroke="#0f172a" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
    <path d="M7 19 L20 25 L33 19" fill="none" stroke="#2563eb" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
    <path d="M11.5 18.8 C7.8 18.8 7.3 13.6 11 13 C10.8 8 17.4 6.4 20 10.3 C22.4 7.6 27.3 9.1 26.9 13 C30.6 13.2 30.4 18.8 26.6 18.8" fill="none" stroke="#2563eb" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round"/>
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

// Wide tables used to put their horizontal scrollbar at the BOTTOM OF THE TABLE,
// so reaching it meant scrolling the whole page down past every row. Capping the
// container's height keeps both scrollbars on screen at all times, and the CSS
// below pins the header row so the columns stay labelled while scrolling.
const scrollBox = extra => ({
  background: T.bgCard, border: '1px solid ' + T.border, borderRadius: T.radius,
  marginBottom: 20, boxShadow: T.shadow,
  overflow: 'auto', maxHeight: 'max(340px, calc(100vh - 300px))',
  ...(extra || {}),
});
const SCROLL_CSS = `
.cl-scroll thead th { position: sticky; top: 0; z-index: 2; }
.cl-scroll { scrollbar-color: #cbd5e1 #f1f5f9; }
.cl-scroll::-webkit-scrollbar { height: 12px; width: 12px; }
.cl-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 6px; }
.cl-scroll::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
.cl-scroll::-webkit-scrollbar-track { background: #f1f5f9; }
`;

const S = {
  app: { fontFamily: "'Inter',-apple-system,sans-serif", background: T.bg, color: T.text, minHeight: '100vh', display: 'flex', flexDirection: 'column', fontSize: 13, lineHeight: 1.5 },
  topBar: { background: T.bgCard, borderBottom: '1px solid '+T.border, padding: '0 24px', height: 56, display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0, zIndex: 10, boxShadow: T.shadow },
  body: { display: 'flex', flex: 1, overflow: 'hidden' },
  sidebar: col => ({ width: col ? 56 : 256, background: T.sidebarBg, padding: col ? '12px 4px' : '16px 0', flexShrink: 0, overflowY: 'auto', overflowX: 'hidden', transition: 'width 0.2s ease', display: 'flex', flexDirection: 'column' }),
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
          <div style={{display:'flex',justifyContent:'flex-end',marginTop:18}}><button style={S.link} onClick={()=>setMode('signup')}>Create account</button></div>
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
function EntityPicker({entities,activeId,onSelect,onManage,defaultId,onSetDefault}){const[open,setOpen]=useState(false);const[search,setSearch]=useState('');const active=entities.find(e=>e.id===activeId);
  const activeRef=useRef(null);
  // When the list opens, scroll the currently-selected entity into view so it's visible.
  useEffect(()=>{if(open&&!search&&activeRef.current){try{activeRef.current.scrollIntoView({block:'center'});}catch(_){}}},[open]);
  const filtered=entities.filter(e=>e.name.toLowerCase().includes(search.toLowerCase()));
  return(<div style={{position:'relative'}}><div style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',padding:'6px 14px',borderRadius:T.radiusSm,background:T.bgElevated,border:'1px solid '+T.border}} onClick={()=>setOpen(!open)}>
    <span style={{fontWeight:600,color:T.textBright,fontSize:13,maxWidth:240,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{active?.name||'Select entity'}</span>
    <span style={{color:T.textDim,fontSize:9}}>{'\u25bc'}</span></div>
    {open&&<><div style={{position:'fixed',inset:0,zIndex:50}} onClick={()=>{setOpen(false);setSearch('');}}/>
      <div style={{position:'absolute',top:'100%',left:0,background:'#fff',border:'1px solid '+T.border,borderRadius:T.radius,maxHeight:380,overflowY:'auto',zIndex:100,boxShadow:T.shadowLg,width:340,marginTop:6}}>
        <div style={{position:'sticky',top:0,padding:12,background:'#fff',borderBottom:'1px solid '+T.border}}><input style={S.input} placeholder={'Search '+entities.length+' entities...'} value={search} onChange={e=>setSearch(e.target.value)} autoFocus/></div>
        {filtered.map(e=><div key={e.id} ref={e.id===activeId?activeRef:null} style={{display:'flex',alignItems:'center',gap:8,padding:'10px 16px',cursor:'pointer',background:e.id===activeId?T.accentDim:'transparent',borderLeft:e.id===activeId?'3px solid '+T.accent:'3px solid transparent'}} onClick={()=>{onSelect(e.id);setOpen(false);setSearch('');}}>
          <span style={{flex:1,fontWeight:600,color:T.textBright,fontSize:13}}>{e.name}</span>
          {e.id===activeId&&<span style={{fontSize:10,fontWeight:700,color:T.accent,background:T.accent+'18',padding:'2px 6px',borderRadius:4,whiteSpace:'nowrap'}}>Current</span>}
          {onSetDefault&&<button title={e.id===defaultId?'Default entity — loads on refresh (click to unset)':'Set as default entity (loads on refresh)'} onClick={ev=>{ev.stopPropagation();onSetDefault(e.id===defaultId?null:e.id);}} style={{background:'none',border:'none',cursor:'pointer',fontSize:16,lineHeight:1,color:e.id===defaultId?T.accent:T.textDim,padding:'0 2px'}}>{e.id===defaultId?'★':'☆'}</button>}</div>)}
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

// ═══ Sidebar category flyout ═══
// Clicking a sidebar category opens this panel to the right of the rail. Items can
// be dragged into a different order; the new order is handed back to the caller,
// which persists it per-user on the server. Positioned `fixed` rather than absolute
// so the sidebar's own overflow can't clip it.
function NavFlyout({catKey,catLabel,items,pos,page,canReorder,onPick,onClose,onReorder}){
  const ids=items.map(i=>i.id);
  const[list,setList]=useState(ids);
  const[drag,setDrag]=useState(null);
  const key=ids.join(',');
  useEffect(()=>{setList(ids);setDrag(null);},[catKey,key]);
  useEffect(()=>{const h=e=>{if(e.key==='Escape')onClose();};window.addEventListener('keydown',h);return()=>window.removeEventListener('keydown',h);},[onClose]);
  const byId={};items.forEach(i=>{byId[i.id]=i;});
  const move=(from,to)=>setList(l=>{const n=l.slice();const[x]=n.splice(from,1);n.splice(to,0,x);return n;});
  const commit=()=>{setDrag(null);if(list.join(',')!==key)onReorder(list);};
  const est=list.length*33+(canReorder?70:34);
  const top=Math.max(12,Math.min(pos.top,Math.max(12,window.innerHeight-est-12)));
  return(<>
    <div style={{position:'fixed',inset:0,zIndex:70}} onClick={onClose}/>
    <div style={{position:'fixed',top,left:pos.left,zIndex:71,width:256,background:T.bgCard,borderRadius:T.radiusSm,border:'1px solid '+T.border,boxShadow:T.shadowLg,padding:'8px 0',maxHeight:'calc(100vh - 24px)',overflowY:'auto'}}>
      <div style={{padding:'2px 14px 8px',fontSize:9,fontWeight:700,color:T.textDim,textTransform:'uppercase',letterSpacing:'0.12em'}}>{catLabel}</div>
      {list.map((id,i)=>{
        const n=byId[id];if(!n)return null;
        const active=page===id;
        return(<div key={id} draggable={canReorder}
          onDragStart={()=>setDrag(i)}
          onDragEnter={()=>{if(drag!=null&&drag!==i){move(drag,i);setDrag(i);}}}
          onDragOver={e=>e.preventDefault()}
          onDragEnd={commit}
          onDrop={e=>{e.preventDefault();commit();}}
          onClick={()=>onPick(id)}
          style={{display:'flex',alignItems:'center',gap:8,padding:'7px 14px',cursor:'pointer',fontSize:13,
            fontWeight:active?600:400,color:active?T.accent:T.text,
            background:drag===i?T.bgHover:(active?T.accentDim:'transparent'),
            borderLeft:'3px solid '+(active?T.accent:'transparent')}}>
          {canReorder&&<span onClick={e=>e.stopPropagation()} title="Drag to reorder" style={{color:T.textDim,fontSize:12,cursor:'grab',lineHeight:1}}>⠿</span>}
          <span style={{width:20,textAlign:'center'}}>{n.icon}</span>
          <span style={{flex:1,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}}>{n.label}</span>
        </div>);
      })}
      {canReorder&&list.length>1&&<div style={{borderTop:'1px solid '+T.borderLight,marginTop:6,paddingTop:7,display:'flex',justifyContent:'space-between',alignItems:'center',padding:'7px 14px 0'}}>
        <span style={{fontSize:10,color:T.textDim}}>Drag to reorder</span>
        <button style={{background:'none',border:'none',color:T.textMuted,fontSize:10,cursor:'pointer',padding:0}} onClick={e=>{e.stopPropagation();onReorder(null);}}>Reset</button></div>}
    </div></>);
}

export default function App(){
  const[user,setUser]=useState(null);const[entities,setEntities]=useState([]);const[activeEntity,setActiveEntity]=useState(null);
  const[defaultEntityId,setDefaultEntityId]=useState(null); // per-user preferred entity loaded on refresh
  // Workpapers modal openable from the header (any page), for the active entity.
  const[wpEntity,setWpEntity]=useState(null);
  const[page,setPage]=useState('dashboard');const[loading,setLoading]=useState(true);
  // Sidebar: which category flyout is open ({key,top,left}), and the per-user
  // item order inside each category ({CATEGORY:[pageId,...]}), loaded from and
  // saved to the server so it follows the user across browsers and machines.
  const[openCat,setOpenCat]=useState(null);
  const[navOrder,setNavOrder]=useState({});
  useEffect(()=>{
    if(!user)return;let live=true;
    (async()=>{try{const p=await api.getMyPrefs();
      if(live&&p&&p.navOrder&&typeof p.navOrder==='object'&&!Array.isArray(p.navOrder))setNavOrder(p.navOrder);
    }catch(e){console.error('[nav] could not load prefs:',e.message);}})();
    return()=>{live=false;};
  },[user]);
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
  useEffect(()=>{if(user)Promise.all([api.getEntities(),api.getMyPrefs().catch(()=>({}))]).then(([e,p])=>{setEntities(e);if(p&&p.defaultEntityId!=null)setDefaultEntityId(p.defaultEntityId);if(e.length>0&&!activeEntity){const def=p&&p.defaultEntityId;setActiveEntity((def!=null&&e.find(x=>x.id===def))?def:e[0].id);}});},[user]);
  const setDefaultEntity=(id)=>{setDefaultEntityId(id);api.saveMyPrefs({defaultEntityId:id}).catch(err=>console.error('[prefs] save default entity failed:',err.message));};
  const refreshEntities=useCallback(async()=>{const e=await api.getEntities();setEntities(e);return e;},[]);
  const canAccess=s=>{if(!user)return false;if(user.role==='Admin')return true;return({Accountant:['entries','reports','coa','bankrec','billcom','workpapers','intercompany'],Viewer:['entries','reports','coa','bankrec','workpapers']}[user.role]||[]).includes(s);};
  // Read-only users (Viewer) SEE the same sections as an Accountant but cannot edit.
  // canEdit gates every write control; it must never be derived from mere visibility.
  const canEdit = !!user && (user.role==='Admin' || (()=>{ const ae=activeEntity?entities.find(e=>e.id===activeEntity):null; return ae&&ae.access_level ? ae.access_level==='full' : user.role==='Accountant'; })());
  if(loading)return<div style={{...S.app,display:'flex',alignItems:'center',justifyContent:'center'}}><div style={{color:T.textMuted}}>Loading...</div></div>;
  const _resetToken=(()=>{try{return new URLSearchParams(window.location.search).get('reset_token');}catch{return null;}})();
  if(_resetToken)return<ResetPasswordScreen token={_resetToken}/>;
  if(!user)return<AuthScreen onLogin={setUser}/>;
  const jeHasContent=jeForm.memo||jeForm.lines.some(l=>l.account_code||l.debit||l.credit)||jePendingFiles.length>0;

  const _activeEnt = entities.find(e=>e.id===activeEntity);
  _activeEntityCode = _activeEnt ? _activeEnt.code : null; // drives per-entity Class relabel (classTerm)
  _activeEntityFileTag = _activeEnt ? (_activeEnt.display_id || _activeEnt.name || '') : '';
  const isTurnkeyEntity = !!(_activeEnt && (_activeEnt.code==='TURNKEYR' || /turnkey\s*rail/i.test(_activeEnt.name||'')));
  const isDevEntity = !!(_activeEnt && _activeEnt.entity_type==='development');
  // County Line Rail Fund — the only entity with the Management Fee workpaper for now.
  const isCLRF = !!(_activeEnt && (_activeEnt.code==='CLRF' || /county\s*line\s*rail\s*fund/i.test(_activeEnt.name||'')));
  // Banyan Residential — pays the health-insurance premium and prepares the
  // monthly Insurance Allocation workpaper. Not a development entity, so it has
  // no Requisitions; the allocation is its only workpaper.
  const isBanyanRes = !!(_activeEnt && (_activeEnt.code==='BANYANRE1' || /^banyan\s*residential$/i.test((_activeEnt.name||'').trim())));
  const isShellEntity = !!(_activeEnt && _activeEnt.entity_type==='shell');
  const dimsEnabled = !!_activeEnt && !isShellEntity && !_activeEnt.hide_dims;// dimensions on every entity EXCEPT shell, unless the entity is flagged hide_dims (e.g. SRN, CLR Silsbee)
  // AR module scope is per-entity by code. Full AR (customers/invoices/recurring
  // + aging) for these six; A/R Aging only for Turnkey (invoices sync in from
  // Turnkey Rail, not created here); no AR anywhere else.
  const AR_FULL_CODES=['BANYANRE1','CLIPPROP','CLRSILSB2','SABINERI','CLRBUNAP','COUNTYLI3'];
  const arFull = !!_activeEnt && AR_FULL_CODES.includes(_activeEnt.code);
  const arAgingEnabled = arFull || !!(_activeEnt && _activeEnt.code==='TURNKEYR');
  // Six top-level categories: Accounting, Banking, A/R, Reports, Workpapers,
  // Administration. The rail shows only the categories; clicking one opens a
  // flyout to its right listing that category's pages. Entity-conditional pages
  // live inside a permanent category so switching entities never reshuffles the
  // rail, and a category with nothing visible to this user is dropped entirely.
  const hasWorkpapers = isDevEntity || isCLRF || isBanyanRes;
  const navTree=[
    {key:'ACCOUNTING',label:'Accounting',icon:'📘',items:[
      {id:'journal',label:'Journal Entries',icon:NI.journal,section:'entries'},
      {id:'coa',label:'Chart of Accounts',icon:NI.coa,section:'coa'},
      ...(dimsEnabled?[{id:'dimensions',label:'Dimensions',icon:'🏷️',section:'coa'}]:[]),
    ]},
    {key:'BANKING',label:'Banking',icon:'🏛️',items:[
      {id:'banktxn',label:'Bank Transactions',icon:NI.banktxn,section:'bankrec'},
      {id:'bankrec',label:'Bank Reconciliation',icon:NI.bankrec,section:'bankrec'},
    ]},
    ...(arAgingEnabled?[{key:'AR',label:'Accounts Receivable',icon:'🧾',items:[
      ...(arFull?[
        {id:'ar_customers',label:'Customers',icon:'👥',section:'coa'},
        {id:'ar_invoices',label:'Invoices',icon:'🧾',section:'coa'},
        {id:'ar_recurring',label:'Recurring',icon:'🔁',section:'coa'},
      ]:[]),
      {id:'ar_aging',label:'A/R Aging',icon:'⏱️',section:'reports'},
    ]}]:[]),
    {key:'AP',label:'Accounts Payable',icon:'📥',items:[
      {id:'apaging',label:'A/P Aging',icon:'⏳',section:'reports'},
      {id:'billcom_sync',label:'Bill.com Sync',icon:'🔄',section:'billcom'},
    ]},
    {key:'INTERCOMPANY',label:'Intercompany',icon:'🔗',items:[
      {id:'ic_recon',label:'IC Reconciliation',icon:'⚖️',section:'intercompany'},
      {id:'ic_mapping',label:'IC Mapping',icon:'🗺️',section:'intercompany'},
      {id:'external_tb',label:'External Entities TB',icon:'📥',section:'intercompany'},
      {id:'org_structure',label:'Org Structure',icon:'🏛️',section:'intercompany'},
    ]},
    {key:'REPORTS',label:'Reports',icon:'📊',items:[
      {id:'wp_finstmts',label:'Financial Statements',icon:'📑',section:'reports'},
      {id:'bs',label:'Balance Sheet',icon:NI.bs,section:'reports'},
      {id:'is',label:'Income Statement',icon:NI.is,section:'reports'},
      {id:'trial',label:'Trial Balance',icon:NI.trial,section:'reports'},
      {id:'ledger',label:'General Ledger',icon:NI.ledger,section:'reports'},
      {id:'ttm',label:'Trailing 12 Months',icon:'📈',section:'reports'},
      {id:'fundrep',label:'Fund Reporting',icon:'🏦',section:'reports'},
      {id:'customdetail',label:'Custom Detail',icon:'📋',section:'reports'},
      ...(dimsEnabled?[{id:'pivot',label:'Pivot Summary',icon:'📊',section:'reports'}]:[]),
      {id:'commitments',label:'Commitments',icon:'🤝',section:'reports'},
      ...(isTurnkeyEntity?[{id:'wip',label:'WIP Schedule',icon:NI.wip,section:'reports'}]:[]),
      {id:'memorized',label:'Memorized Reports',icon:'★',section:'reports'},
    ]},
    ...(hasWorkpapers?[{key:'WORKPAPERS',label:'Workpapers',icon:'📄',items:[
      ...(isDevEntity?[{id:'requisitions',label:'Requisitions',icon:'🏗️',section:'reports'}]:[]),
      ...(isCLRF?[{id:'wp_mgmtfee',label:'Management Fee',icon:'📄',section:'workpapers'}]:[]),
      ...(isCLRF?[{id:'wp_gpfees',label:'GP Fees & Expenses',icon:'🤝',section:'workpapers'}]:[]),
      ...(isCLRF?[{id:'wp_valuation',label:'Investment & Valuation',icon:'🏦',section:'workpapers'}]:[]),
      ...(isBanyanRes?[{id:'wp_insalloc',label:'Insurance Allocation',icon:'🩺',section:'workpapers'}]:[]),
    ]}]:[]),
    {key:'ADMINISTRATION',label:'Administration',icon:'⚙️',items:[
      {id:'entities',label:'Entities ('+entities.length+')',icon:NI.entities,section:'all'},
      {id:'users',label:'Users',icon:NI.users,section:'all'},
      {id:'billcom',label:'Bill.com Setup',icon:'💳',section:'billcom'},
    ]},
  ];
  // Access filter, then the user's saved order. Items with no saved position sort
  // to the end in their declared order (Array.sort is stable), so a newly shipped
  // page appears at the bottom of its category rather than disappearing.
  const visibleItems=cat=>cat.items.filter(n=>n.section==='all'?user.role==='Admin':canAccess(n.section));
  const orderedItems=cat=>{
    const ord=navOrder[cat.key]||[];
    const rank=id=>{const i=ord.indexOf(id);return i<0?1e6:i;};
    return visibleItems(cat).slice().sort((a,b)=>rank(a.id)-rank(b.id));
  };
  const navCats=navTree.filter(c=>visibleItems(c).length>0);
  const activeCat=navCats.find(c=>c.items.some(n=>n.id===page));
  const openCatDef=openCat?navCats.find(c=>c.key===openCat.key):null;
  const saveNavOrder=async(catKey,ids)=>{
    const next={...navOrder};
    if(ids)next[catKey]=ids;else delete next[catKey];
    setNavOrder(next);
    try{await api.saveMyPrefs({navOrder:next});}catch(e){console.error('[nav] could not save order:',e.message);}
  };

  return(<div style={S.app}>
    <style>{SCROLL_CSS}</style>
    <div style={S.topBar}><div style={{display:'flex',alignItems:'center',gap:16}}>
      <button style={{...S.btnGhost,fontSize:18,padding:'4px 6px',color:T.textMuted}} onClick={()=>setSidebarCol(c=>!c)}>{sidebarCol?'\u2630':'\u2190'}</button>
      <div style={{display:'flex',alignItems:'center',gap:10}}><Logo size={32}/>{!sidebarCol&&<div style={{fontSize:17,fontWeight:800,color:T.textBright}}>CloudLedger</div>}</div>
      <div style={{width:1,height:28,background:T.border}}/>{_activeEnt&&<button style={{...S.btnGhost,fontSize:18,padding:'4px 8px',lineHeight:1}} title={'Open '+_activeEnt.name+' Workpapers'} onClick={()=>setWpEntity(_activeEnt)}>📁</button>}<EntityPicker entities={entities} activeId={activeEntity} onSelect={setActiveEntity} onManage={()=>setPage('entities')} defaultId={defaultEntityId} onSetDefault={setDefaultEntity}/>{entities.length>0&&<GlobalSearch entities={entities} activeEntity={activeEntity} onSelectEntity={setActiveEntity} onGo={setPage} onPickJE={(id)=>{setPendingJEId(id);setPage('journal');}} onPickAccount={(code)=>{setPendingAcctCode(code);setPage('coa');}}/>}</div>
      <div style={{display:'flex',alignItems:'center',gap:10}}>
        {canEdit&&activeEntity&&<button style={{...S.btnP,position:'relative'}} onClick={()=>setShowJE(true)}>+ Journal Entry{jeHasContent&&<span style={{position:'absolute',top:-3,right:-3,width:8,height:8,borderRadius:4,background:T.orange,border:'2px solid #fff'}}/>}</button>}
        <span style={{fontSize:13,fontWeight:500}}>{titleName(user.name)}</span><span style={S.badge}>{user.role}</span>
        <button style={S.btnS} onClick={()=>setShowChangePw(true)}>Settings</button>
        <button style={S.btnS} onClick={()=>{api.clearToken();setUser(null);}}>Sign Out</button></div></div>
    <div style={S.body}><div style={S.sidebar(sidebarCol)}>
      <div style={S.navItem(page==='dashboard',sidebarCol)} onClick={()=>{setOpenCat(null);setPage('dashboard');}} title="Dashboard">
        {sidebarCol?<span style={{display:'inline-block',width:18,textAlign:'center',fontSize:15}}>{NI.dashboard}</span>
          :<span><span style={{display:'inline-block',width:22,textAlign:'center',marginRight:8}}>{NI.dashboard}</span>Dashboard</span>}</div>
      <div style={{height:10}}/>
      {navCats.map(c=>{
        const isOpen=!!openCat&&openCat.key===c.key;
        return(<div key={c.key} title={c.label}
          style={{...S.navItem(isOpen||(activeCat&&activeCat.key===c.key),sidebarCol),display:'flex',alignItems:'center',justifyContent:sidebarCol?'center':'space-between'}}
          onClick={e=>{const r=e.currentTarget.getBoundingClientRect();setOpenCat(isOpen?null:{key:c.key,top:r.top,left:r.right+6});}}>
          {sidebarCol?<span style={{fontSize:15}}>{c.icon}</span>:<>
            <span style={{flex:1,minWidth:0,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',paddingRight:25}}><span style={{display:'inline-block',width:22,textAlign:'center',marginRight:8}}>{c.icon}</span>{c.label}</span>
            <span style={{fontSize:9,marginRight:8,opacity:0.75,flexShrink:0,display:'inline-block',transform:isOpen?'rotate(90deg)':'none',transition:'transform 0.12s'}}>{'\u25B6'}</span></>}
        </div>);
      })}</div>
      {openCat&&openCatDef&&<NavFlyout catKey={openCatDef.key} catLabel={openCatDef.label} items={orderedItems(openCatDef)}
        pos={openCat} page={page} canReorder={true}
        onPick={id=>{setPage(id);setOpenCat(null);}} onClose={()=>setOpenCat(null)}
        onReorder={ids=>saveNavOrder(openCatDef.key,ids)}/>}
      <div style={S.main}>{(()=>{const en=entities.find(e=>e.id===activeEntity);const entityName=en?en.name:'';return<>
        {page==='dashboard'&&<Dashboard entityId={activeEntity} setActiveEntity={setActiveEntity} setPage={setPage} user={user} key={rk}/>}
        {page==='journal'&&activeEntity&&<JournalList entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} canEdit={canEdit} key={activeEntity+'-'+rk} onNewEntry={()=>setShowJE(true)} openJEId={pendingJEId} clearOpenJE={()=>setPendingJEId(null)}/>}
        {page==='coa'&&activeEntity&&<ChartOfAccounts entityId={activeEntity} entityName={entityName} canEdit={canEdit}/>}
        {page==='dimensions'&&activeEntity&&dimsEnabled&&<DimensionsManager entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ar_customers'&&activeEntity&&arFull&&<CustomersManager entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='ar_invoices'&&activeEntity&&arFull&&<ArInvoices entityId={activeEntity} entityName={entityName} canEdit={canEdit} dimsEnabled={dimsEnabled} isBanyanRes={isBanyanRes} key={activeEntity+'-'+rk}/>}
        {page==='ar_recurring'&&activeEntity&&arFull&&<ArRecurring entityId={activeEntity} entityName={entityName} canEdit={canEdit} dimsEnabled={dimsEnabled} key={activeEntity+'-'+rk}/>}
        {page==='ar_aging'&&activeEntity&&arAgingEnabled&&<ArAgingReport entityId={activeEntity} entityName={entityName} key={activeEntity+'-'+rk}/>}
        {page==='ledger'&&activeEntity&&<GeneralLedger entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} key={activeEntity+'-'+rk} from={glFrom} setFrom={setGlFrom} to={glTo} setTo={setGlTo} filter={glFilter} setFilter={setGlFilter}/>}
        {page==='banktxn'&&activeEntity&&<BankTransactions entityId={activeEntity} canEdit={canEdit} dimsEnabled={dimsEnabled} bankSelAcct={bankSelAcct} setBankSelAcct={setBankSelAcct} bankTxns={bankTxns} setBankTxns={setBankTxns} bankUploading={bankUploading} setBankUploading={setBankUploading} bankStatusFilter={bankStatusFilter} setBankStatusFilter={setBankStatusFilter}/>}
        {page==='bankrec'&&activeEntity&&<BankReconciliation entityId={activeEntity} user={user} canEdit={canEdit}/>}
        {page==='trial'&&activeEntity&&<TrialBalance entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} isClrf={_activeEnt?.code==='COUNTYLI1'} key={activeEntity+'-'+rk} asOf={tbAsOf} setAsOf={setTbAsOf} canEdit={canEdit}/>}
        {page==='bs'&&activeEntity&&<BalanceSheet entityId={activeEntity} entityName={entityName} asOf={bsAsOf} setAsOf={setBsAsOf} canEdit={canEdit}/>}
        {page==='is'&&activeEntity&&<IncomeStatement entityId={activeEntity} entityName={entityName} from={isFrom} setFrom={setIsFrom} to={isTo} setTo={setIsTo} canEdit={canEdit}/>}
        {page==='customdetail'&&activeEntity&&<CustomDetailReport entityId={activeEntity} entityName={entityName} dimsEnabled={dimsEnabled} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='customdetail'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='pivot'&&activeEntity&&dimsEnabled&&<PivotReport entityId={activeEntity} entityName={entityName} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='pivot'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='ic_recon'&&canAccess('intercompany')&&<IntercompanyReconciliation entities={entities} activeEntity={activeEntity} setPage={setPage} key={'icr-'+rk}/>}
        {page==='external_tb'&&canAccess('intercompany')&&<ExternalTbPage canEdit={canEdit} key={'etb-'+rk}/>}
        {page==='org_structure'&&canAccess('intercompany')&&<OrgStructurePage entities={entities} canEdit={canEdit} key={'org-'+rk}/>}
        {page==='ic_mapping'&&canAccess('intercompany')&&<IntercompanyMapping entities={entities} activeEntity={activeEntity} canEdit={canEdit} key={'icm-'+rk}/>}
        {page==='apaging'&&activeEntity&&<ApAgingReport entityId={activeEntity} entityName={entityName} canEdit={canEdit} pendingConfig={pendingReportConfig&&pendingReportConfig.type==='apaging'?pendingReportConfig.config:null} clearPending={()=>setPendingReportConfig(null)} key={activeEntity+'-'+rk}/>}
        {page==='commitments'&&activeEntity&&<CommitmentsPage entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='memorized'&&activeEntity&&<MemorizedReportsPage entityId={activeEntity} entityName={entityName} canEdit={canEdit} onOpen={(r)=>{const c=r.config||{};if(r.report_type==='trial'&&c.asOf)setTbAsOf(c.asOf);else if(r.report_type==='bs'&&c.asOf)setBsAsOf(c.asOf);else if(r.report_type==='is'){if(c.from)setIsFrom(c.from);if(c.to)setIsTo(c.to);}else setPendingReportConfig({type:r.report_type,config:c});setPage(r.report_type==='drilldown'?'coa':r.report_type);}} key={activeEntity+'-'+rk}/>}
        {page==='wip'&&activeEntity&&<WipSchedule entityName={entityName} asOf={wipAsOf} setAsOf={setWipAsOf}/>}
        {page==='entities'&&<EntityManagement refresh={refreshEntities} entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='users'&&<UserManagement currentUser={user}/>}
        {page==='billcom'&&<BillcomSetup entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity}/>}
        {page==='billcom_sync'&&<BillcomSetup entities={entities} activeEntity={activeEntity} setActiveEntity={setActiveEntity} initialTab='sync'/>}
        {page==='requisitions'&&activeEntity&&isDevEntity&&<Requisitions entityId={activeEntity} entityName={entityName} canEdit={canEdit} reqState={reqState} setReqState={setReqState}/>}
        {page==='wp_mgmtfee'&&activeEntity&&isCLRF&&<MgmtFeeWorkpaper entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='wp_gpfees'&&activeEntity&&isCLRF&&<GpFeesWorkpaper entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='wp_valuation'&&activeEntity&&isCLRF&&<ValuationWorkpaper entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='wp_insalloc'&&activeEntity&&isBanyanRes&&<InsuranceAllocationWorkpaper entityId={activeEntity} entityName={entityName} canEdit={canEdit} key={activeEntity+'-'+rk}/>}
        {page==='wp_finstmts'&&activeEntity&&<FinancialStatements entityId={activeEntity} entityName={entityName} canEdit={canEdit} isDevEntity={isDevEntity} key={activeEntity+'-'+rk}/>}
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


function BillcomSetup({entities,activeEntity,setActiveEntity,initialTab}) {
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
  const[syncCutoffDate,setSyncCutoffDate]=useState(''); // only sync invoices dated on/after this
  const[unsyncing,setUnsyncing]=useState(false);
  const[cfgIds,setCfgIds]=useState(null); // entity_ids that have Bill.com configured
  const[showAllEnts,setShowAllEnts]=useState(false);
  useEffect(()=>{api.getBillcomEntities().then(r=>setCfgIds(new Set(r.entity_ids||[]))).catch(()=>setCfgIds(new Set()));},[]);

  // Phase 2: account mapping state
  const[tab,setTab]=useState(initialTab||'config'); // 'config' | 'mapping' | 'sync'
  const syncOnly=initialTab==='sync'; // A/P "Bill.com Sync" view: only the Sync panel, no setup chrome
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

  // When opened directly on the Sync view (the Bill.com Sync nav item), load the
  // sync log once config is known so it shows without a manual Refresh Log click.
  useEffect(()=>{ if((syncOnly||tab==='sync')&&cfg&&cfg.configured)loadSyncLogs(); },[cfg,tab,syncOnly,loadSyncLogs]);

  const runSync=async()=>{
    if(!selectedEntity)return;
    setSyncing(true);setSyncMsg('');setSyncErr('');setSyncResult(null);
    // The server processes a bounded batch per call (~25 bills) and sets
    // bills.budget_reached=true when more remain. Auto-continue: keep calling
    // until it's caught up, so one click drains the whole backlog. Hard cap on
    // iterations guarantees this can never loop forever.
    const MAX_ROUNDS=40;
    let round=0, last=null;
    // synced + errors are cumulative (real new work each batch); skipped is NOT
    // summed because each batch re-scans and re-skips the same pre-cutoff /
    // already-synced bills, so we report the final batch's skipped count only.
    const tot={bs:0,be:0,ps:0,pe:0};
    const agingOv=new Map();
    try{
      while(round<MAX_ROUNDS){
        round++;
        setSyncMsg('Syncing batch '+round+'\u2026'+(round>1?' ('+tot.bs+' bills, '+tot.ps+' payments posted so far)':''));
        const r=await api.syncBillcom(selectedEntity);
        last=r;
        const b=r.bills||{},py=r.payments||{};
        tot.bs+=(b.synced||0);tot.be+=(b.errors||0);
        tot.ps+=(py.synced||0);tot.pe+=(py.errors||0);
        (b.aging_overlaps||[]).forEach(o=>agingOv.set(o.id,o));
        if(!(b.budget_reached)) break; // caught up
      }
      if(last&&last.bills)last.bills.aging_overlaps=[...agingOv.values()];
      setSyncResult(last);
      const lb=(last&&last.bills)||{},lp=(last&&last.payments)||{};
      const capNote=(round>=MAX_ROUNDS)?' (stopped at batch limit \u2014 click Sync Now again to continue)':'';
      setSyncMsg('Done in '+round+' batch'+(round===1?'':'es')+'. Bills: '+tot.bs+' posted, '+tot.be+' errors, '+(lb.skipped||0)+' skipped. Payments: '+tot.ps+' posted, '+tot.pe+' errors, '+(lp.skipped||0)+' skipped.'+(agingOv.size?(' '+agingOv.size+' already in A/P aging.'):'')+capNote);
      loadSyncLogs();
    }catch(e){setSyncErr('Sync failed after '+round+' batch'+(round===1?'':'es')+': '+e.message+(tot.bs?(' ('+tot.bs+' bills already posted before the error)'):''));}
    setSyncing(false);
  };

  const runUnsync=async()=>{
    if(!selectedEntity)return;
    let count=0;
    try{ const dry=await api.unsyncBillcom(selectedEntity,true); count=dry.would_delete_entries||0; }
    catch(e){ setSyncErr('Un-sync check failed: '+e.message); return; }
    if(count===0){
      // No Bill.com-created journal entries to delete, but there may still be
      // sync-log rows (e.g. skipped-bill records) that must be cleared so a
      // re-sync re-evaluates every bill under the current rules. Clear the log.
      setUnsyncing(true);setSyncMsg('');setSyncErr('');
      try{ const r=await api.unsyncBillcom(selectedEntity,false); setSyncMsg('No Bill.com journal entries to delete; cleared the sync log'+(r.sync_log_cleared?'':' (nothing to clear)')+' so a re-sync will re-evaluate every bill from scratch.'); loadSyncLogs(); }
      catch(e){ setSyncErr('Un-sync (clear log) failed: '+e.message); }
      setUnsyncing(false); return;
    }
    if(!window.confirm('Un-sync will DELETE '+count+' journal entr'+(count===1?'y':'ies')+' that Bill.com syncs created on this entity, and clear the sync log so a corrected sync can re-pull. It does NOT touch imported-GL or manual entries. Continue?')) return;
    setUnsyncing(true);setSyncMsg('');setSyncErr('');setSyncResult(null);
    try{
      const r=await api.unsyncBillcom(selectedEntity,false);
      setSyncMsg('Un-synced. Removed '+(r.deleted_entries||0)+' Bill.com journal entr'+((r.deleted_entries===1)?'y':'ies')+' and cleared the sync log. Set a cutoff date if needed, then Sync Now to re-pull correctly.');
      loadSyncLogs();
    }catch(e){setSyncErr('Un-sync failed: '+e.message);}
    setUnsyncing(false);
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
        setSyncCutoffDate(r.sync_cutoff_date||'');
        setPassword('');setDevKey('');
      }else{
        setEnv('sandbox');setUsername('');setOrgId('');setDefaultApAcct('');setDefaultCashAcct('');setDefaultClearingAcct('');setSyncCutoffDate('');
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

  // Follow the top-level entity picker: when the active entity changes up top,
  // this in-page selector switches to it too (no stale sub-selection).
  useEffect(()=>{ if(activeEntity&&activeEntity!==selectedEntity)setSelectedEntity(activeEntity); },[activeEntity]);

  const save=async()=>{
    if(!selectedEntity){setErr('Select an entity first');return;}
    setSaving(true);setMsg('');setErr('');
    try{
      const body={environment:env,username,org_id:orgId,default_ap_account:defaultApAcct||null,default_cash_account:defaultCashAcct||null,default_clearing_account:defaultClearingAcct||null,sync_cutoff_date:syncCutoffDate||null};
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
        <div style={{fontSize:22,fontWeight:700,color:T.textBright}}>{syncOnly?'Bill.com Sync':'Bill.com Setup'}</div>
        <div style={{fontSize:13,color:T.textMuted,marginTop:4}}>{syncOnly?('Pull approved bills and payments from Bill.com into '+(en?en.name:'this entity')+' as journal entries.'):'Connect a CloudLedger entity to a Bill.com Organization for AP integration.'}</div>
      </div>
    </div>

    {!syncOnly&&<div style={{...S.card,padding:20,marginBottom:20}}>
      <div style={{fontSize:12,fontWeight:600,color:T.textMuted,marginBottom:6}}>ENTITY</div>
      {(()=>{const listed=(showAllEnts||!cfgIds)?entities:entities.filter(e=>cfgIds.has(e.id)||e.id===selectedEntity);return(
      <select value={selectedEntity||''} onChange={e=>setSelectedEntity(parseInt(e.target.value)||null)} style={{...S.input,maxWidth:400}}>
        <option value="">-- Select entity --</option>
        {listed.map(e=><option key={e.id} value={e.id}>{e.name}</option>)}
      </select>);})()}
      <label style={{display:'inline-flex',alignItems:'center',gap:6,marginLeft:12,fontSize:12,color:T.textMuted,cursor:'pointer'}}><input type="checkbox" checked={showAllEnts} onChange={e=>setShowAllEnts(e.target.checked)}/>Show all entities</label>
    </div>}

    {selectedEntity&&(loading?<div style={{color:T.textMuted}}>Loading...</div>:<>
      {!syncOnly&&cfg&&cfg.configured&&<div style={{...S.card,padding:16,marginBottom:16,background:'#f0fdf4',border:'1px solid #86efac'}}>
        <div style={{fontSize:13,fontWeight:600,color:'#15803d'}}>Configured</div>
        <div style={{fontSize:12,color:T.textMuted,marginTop:6}}>
          Last tested: {cfg.last_tested_at?new Date(cfg.last_tested_at).toLocaleString():'never'}
          {cfg.last_test_status&&<> · Status: <b style={{color:cfg.last_test_status==='success'?'#15803d':T.red}}>{cfg.last_test_status}</b></>}
        </div>
        {cfg.last_test_message&&<div style={{fontSize:11,color:T.textMuted,marginTop:4,fontFamily:'monospace'}}>{cfg.last_test_message}</div>}
      </div>}

      {/* Phase 2: Tab bar */}
      {!syncOnly&&<div style={{display:'flex',gap:0,borderBottom:'1px solid '+T.border,marginBottom:16}}>
        <button onClick={()=>setTab('config')} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='config'?'2px solid '+T.accent:'2px solid transparent',color:tab==='config'?T.textBright:T.textMuted,cursor:'pointer'}}>Config</button>
        <button onClick={()=>{setTab('mapping');if(cfg&&cfg.configured)loadMapping();}} disabled={!cfg||!cfg.configured} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='mapping'?'2px solid '+T.accent:'2px solid transparent',color:tab==='mapping'?T.textBright:T.textMuted,cursor:cfg&&cfg.configured?'pointer':'not-allowed',opacity:cfg&&cfg.configured?1:0.5}}>Account Mapping</button>
        <button onClick={()=>{setTab('sync');if(cfg&&cfg.configured)loadSyncLogs();}} disabled={!cfg||!cfg.configured} style={{padding:'10px 18px',fontSize:13,fontWeight:600,background:'transparent',border:'none',borderBottom:tab==='sync'?'2px solid '+T.accent:'2px solid transparent',color:tab==='sync'?T.textBright:T.textMuted,cursor:cfg&&cfg.configured?'pointer':'not-allowed',opacity:cfg&&cfg.configured?1:0.5}}>Sync</button>
      </div>}

      {(!syncOnly&&tab==='config')?<>
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
          <div>
            <label style={S.label}>Sync Cutoff Date</label>
            <input type="date" value={syncCutoffDate} onChange={e=>setSyncCutoffDate(e.target.value)} style={S.input}/>
            <div style={{fontSize:11,color:T.textMuted,marginTop:4}}>Only bills <b>approved on or after</b> this date are synced (AP is booked on approval, so a late invoice with an older invoice date still syncs once it's approved). Set this to your CloudLedger go-live date for this entity so bills already booked in the prior system aren't duplicated. Leave blank to sync everything.</div>
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
      </>:(!syncOnly&&tab==='mapping')?<>
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
      {syncOnly&&(!cfg||!cfg.configured)?<div style={{...S.card,padding:20,background:'#fffbeb',border:'1px solid #fcd34d'}}>
        <div style={{fontSize:14,fontWeight:600,color:'#92400e',marginBottom:4}}>Bill.com isn't configured for {en?en.name:'this entity'}</div>
        <div style={{fontSize:12,color:'#92400e'}}>An admin can set it up under Administration &rarr; Bill.com Setup. Once it's configured, come back here to sync.</div>
      </div>:<>
      <div style={{...S.card,padding:20}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:14}}>
          <div>
            <div style={{fontSize:14,fontWeight:600,color:T.textBright}}>Sync</div>
            <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>Pull approved bills and payments from Bill.com and create journal entries. Already-synced items are skipped.</div>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button style={S.btnS} onClick={loadSyncLogs} disabled={syncLogsLoading||syncing}>{syncLogsLoading?'Loading...':'Refresh Log'}</button>
            <button style={{...S.btnS,color:'#b91c1c',borderColor:'#fca5a5'}} onClick={runUnsync} disabled={syncing||unsyncing}>{unsyncing?'Un-syncing...':'Un-sync'}</button>
            <button style={S.btnP} onClick={runSync} disabled={syncing||unsyncing}>{syncing?'Syncing...':'Sync Now'}</button>
          </div>
        </div>

        {syncErr&&<div style={{padding:10,marginBottom:12,background:'#fef2f2',border:'1px solid #fecaca',borderRadius:T.radiusSm,color:T.red,fontSize:12}}>{syncErr}</div>}
        {syncMsg&&<div style={{padding:10,marginBottom:12,background:'#f0fdf4',border:'1px solid #86efac',borderRadius:T.radiusSm,color:'#15803d',fontSize:12}}>{syncMsg}</div>}

        {syncResult&&syncResult.bills&&(syncResult.bills.aging_overlaps||[]).length>0&&(()=>{ const ov=syncResult.bills.aging_overlaps; return (<div style={{padding:12,marginBottom:12,background:'#fffbeb',border:'1px solid #fcd34d',borderRadius:T.radiusSm,fontSize:12}}>
          <div style={{fontWeight:600,color:'#92400e',marginBottom:4}}>{ov.length+' bill'+(ov.length===1?'':'s')+' skipped — already in the A/P aging (already booked in the GL)'}</div>
          <div style={{overflowX:'auto'}}><table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}><thead><tr>
            <th style={{textAlign:'left',padding:'4px 6px',color:T.textMuted}}>Invoice #</th><th style={{textAlign:'left',padding:'4px 6px',color:T.textMuted}}>Vendor</th><th style={{textAlign:'left',padding:'4px 6px',color:T.textMuted}}>Date</th><th style={{textAlign:'right',padding:'4px 6px',color:T.textMuted}}>Amount</th><th style={{textAlign:'left',padding:'4px 6px',color:T.textMuted}}>Matched on</th>
          </tr></thead><tbody>{ov.map((o,i)=><tr key={i} style={{borderTop:'1px solid '+T.border}}>
            <td style={{padding:'4px 6px'}}>{o.invoice_number}</td><td style={{padding:'4px 6px'}}>{o.vendor||'—'}</td><td style={{padding:'4px 6px'}}>{o.date||'—'}</td><td style={{padding:'4px 6px',textAlign:'right'}}>{o.amount!=null?fmt(o.amount):'—'}</td><td style={{padding:'4px 6px'}}>{o.matched_on}</td>
          </tr>)}</tbody></table></div>
          <button style={{...S.btnS,marginTop:8}} onClick={()=>{const rows=[['Invoice #','Vendor','Date','Amount','Matched on'],...ov.map(o=>[o.invoice_number,o.vendor||'',o.date||'',o.amount!=null?o.amount:'',o.matched_on])];exportToExcel(rows,'AP_aging_skipped.xlsx');}}>Download skipped report</button>
        </div>); })()}

        {syncResult&&syncResult.missing_mappings&&syncResult.missing_mappings.length>0&&(()=>{
          const mm=syncResult.missing_mappings;
          const suggested=mm.filter(m=>m.suggested_gl);
          const ambiguous=mm.filter(m=>m.ambiguous);
          const nomatch=mm.filter(m=>!m.suggested_gl&&!m.ambiguous);
          return <div style={{padding:12,marginBottom:14,background:'#fffbeb',border:'1px solid #fcd34d',borderRadius:T.radiusSm}}>
          <div style={{fontSize:13,fontWeight:600,color:'#92400e',marginBottom:6}}>Missing GL Mappings ({mm.length})</div>
          <div style={{fontSize:12,color:'#92400e',marginBottom:10}}>These Bill.com accounts appeared in bills but aren't mapped to a CloudLedger GL account yet. Here's exactly what each one needs:</div>
          {suggested.length>0&&<div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:600,color:'#15803d',marginBottom:4}}>✓ Ready to map — the Bill.com name matches a GL account exactly:</div>
            <ul style={{margin:0,paddingLeft:20,fontSize:12,color:'#166534'}}>
              {suggested.map(m=><li key={m.billcom_account_id} style={{marginBottom:2}}><b>{m.billcom_account_name||m.name}</b> → <b>{m.suggested_gl.cl_account_code} {m.suggested_gl.cl_account_name}</b> <span style={{color:'#4d7c0f'}}>({m.affected_bills} bill{m.affected_bills===1?'':'s'})</span></li>)}
            </ul>
            <div style={{fontSize:11,color:'#166534',marginTop:4,fontStyle:'italic'}}>These will be mapped and posted automatically on the next sync — just click Sync Now again.</div>
          </div>}
          {ambiguous.length>0&&<div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:600,color:'#b45309',marginBottom:4}}>⚠ Needs your decision — the name matches more than one GL account:</div>
            <ul style={{margin:0,paddingLeft:20,fontSize:12,color:'#78350f'}}>
              {ambiguous.map(m=><li key={m.billcom_account_id} style={{marginBottom:2}}><b>{m.billcom_account_name||m.name}</b> <span style={{color:'#a16207'}}>({m.affected_bills} bill{m.affected_bills===1?'':'s'})</span> — pick the correct GL account in the Account Mapping tab.</li>)}
            </ul>
          </div>}
          {nomatch.length>0&&<div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:600,color:'#b91c1c',marginBottom:4}}>✕ No matching GL account — these need a mapping chosen manually:</div>
            <ul style={{margin:0,paddingLeft:20,fontSize:12,color:'#78350f'}}>
              {nomatch.map(m=><li key={m.billcom_account_id} style={{marginBottom:2}}><b>{m.billcom_account_name||m.name}</b> <span style={{color:'#a16207'}}>({m.affected_bills} bill{m.affected_bills===1?'':'s'})</span> — no GL account has this name; map it in the Account Mapping tab.</li>)}
            </ul>
          </div>}
          <button style={{...S.btnS,marginTop:4}} onClick={()=>{setTab('mapping');if(cfg&&cfg.configured)loadMapping();}}>Open Account Mapping tab</button>
        </div>;})()}

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
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Invoice #</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Status</th>
                <th style={{textAlign:'left',padding:'8px 10px',fontWeight:600,color:T.textMuted,fontSize:11,textTransform:'uppercase'}}>Message</th>
              </tr>
            </thead>
            <tbody>
              {syncLogs.map(l=>(
                <tr key={l.id} style={{borderTop:'1px solid '+T.border}}>
                  <td style={{padding:'8px 10px',color:T.textMuted,whiteSpace:'nowrap'}}>{l.created_at?new Date(l.created_at).toLocaleString():''}</td>
                  <td style={{padding:'8px 10px'}}>{l.sync_type}</td>
                  <td style={{padding:'8px 10px',fontFamily:'monospace',fontSize:11}} title={l.billcom_id?('Bill.com ID: '+l.billcom_id):''}>{l.invoice_number||l.billcom_id}</td>
                  <td style={{padding:'8px 10px'}}><span style={{color:l.status==='success'?'#15803d':T.red,fontWeight:600}}>{l.status}</span></td>
                  <td style={{padding:'8px 10px',color:T.textMuted}}>{l.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
         </div>}
      </div>
      </>}
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
function EditJEModal({entityId,dimsEnabled=true,isTurnkeyEntity=false,entry,accounts:initAccounts,onClose,onSaved}){
  // dimsEnabled defaults to true: report drilldowns (AccountDrillDownModal, ApAgingReport,
  // FundReporting) open this modal without the prop and only ever do so for real
  // accounting/fund entities, so the full Location/Class/Project dimension list must show —
  // matching the New JE modal. Callers that manage shell entities (JournalList, GeneralLedger)
  // still pass the explicit value, so shell entities correctly get dimsEnabled=false there.
  const[accounts,setAccounts]=useState(initAccounts||[]);const[showAddAcct,setShowAddAcct]=useState(false);const[err,setErr]=useState('');const[saving,setSaving]=useState(false);
  const[projects,setProjects]=useState([]);const[dimProjects,setDimProjects]=useState([]);
  useEffect(()=>{api.getTurnkeyProjects().then(setProjects).catch(()=>setProjects([]));api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>setDimProjects([]));},[entityId]);
  // Whether THIS entity is Turnkey Rail is decided by the entity (passed from the
  // parent when known), NOT by whether any Turnkey projects exist globally.
  // getTurnkeyProjects() returns a GLOBAL list, so keying off its length made every
  // entity look like a Turnkey entity whenever a single Turnkey project existed —
  // which suppressed the project-dimension list on ordinary accounting entities
  // (e.g. Banyan Residential, whose Dimension dropdown then showed only the lone
  // global Turnkey project). Fallback for callers that don't pass the flag: an entity
  // that has its OWN dimension projects uses them; only an entity with none falls back
  // to the global Turnkey project list. Turnkey Rail has no dimension projects, so it
  // still gets the Turnkey list; Banyan (278 dimension projects) gets its projects.
  const isTurnkeyHere=isTurnkeyEntity||(projects.length>0&&dimProjects.length===0);
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
  const startColDrag=(key,ev)=>{ev.preventDefault();ev.stopPropagation();const sx=ev.clientX;const sw=colW[key];const min=key==='acct'?60:key==='desc'?50:44;document.body.style.userSelect='none';const mv=e=>setColW(p=>({...p,[key]:Math.max(min,sw+(e.clientX-sx))}));const up=()=>{document.body.style.userSelect='';window.removeEventListener('mousemove',mv);window.removeEventListener('mouseup',up);};window.addEventListener('mousemove',mv);window.addEventListener('mouseup',up);};
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
          <tbody>{e.lines.map((l,i)=><tr key={i}><td style={{...S.td,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={acctLabel(l.account_code,acctName(l.account_code))}>{acctLabel(l.account_code,acctName(l.account_code))}</td>
            <td style={{...S.td,color:T.textMuted,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={[l.description||'',l.project_name?('Project: '+(l.project_code&&l.project_code!==l.project_name?l.project_code:l.project_name)):'',l.location_name?('Location: '+l.location_name):'',l.class_name?('Class: '+l.class_name):''].filter(Boolean).join('  ·  ')}>{l.description||''}{(l.project_name||l.location_name||l.class_name)&&<span style={{marginLeft:l.description?8:0,fontSize:10,color:T.accent}}>{[l.project_name?('▦ '+(l.project_code&&l.project_code!==l.project_name?l.project_code:l.project_name)):'',l.location_name,l.class_name].filter(Boolean).join(' · ')}</span>}</td>
            <td style={{...S.tdR,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.debit>0?fmt(l.debit):''}</td><td style={{...S.tdR,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{l.credit>0?fmt(l.credit):''}</td></tr>)}</tbody></table></div>
        {e.attachments?.length>0&&<div style={{marginTop:10,display:'flex',flexWrap:'wrap',gap:4}}>{e.attachments.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>)}</div>}
      </div>)}</div>}
    {editEntry&&<EditJEModal entityId={entityId} dimsEnabled={dimsEnabled} isTurnkeyEntity={/turnkey\s*rail/i.test(entityName||'')} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={load}/>}
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
    <div style={{marginTop:14,fontSize:12,color:T.textMuted}}>Customers defined here feed Invoices and Recurring under RECEIVABLES. The email on a customer is where its invoices get sent.</div>
  </div>);
}

// ═══ Accounts Receivable: settings, invoices, recurring templates, aging ═══
// Invoices post an accrual JE on creation (Dr A/R, Cr Revenue). Sending emails
// the PDF and files a copy under Workpapers > Invoices/<year>. Receipts post
// Dr Bank / Cr A/R, so the aging always ties back to the GL A/R account.

const AR_ST={draft:'#94a3b8',sent:'#2563eb',paid:'#16a34a',void:'#ef4444'};
function ArBadge({inv}){
  const st=inv.status||'draft';
  const late=st!=='paid'&&st!=='void'&&inv.due_date&&inv.due_date<today()&&(inv.open_amount==null||inv.open_amount>0.005);
  const isCm=inv.doc_type==='credit_memo';
  const label=st==='void'?'Void':isCm?(st==='paid'?'Credit · applied':'Credit memo'):st==='paid'?'Paid':late?'Overdue':st==='sent'?'Sent':'Draft';
  const color=isCm?T.orange:late?T.orange:(AR_ST[st]||T.textMuted);
  return <span style={{fontSize:11,fontWeight:600,color,border:'1px solid '+color+'55',borderRadius:10,padding:'2px 8px',whiteSpace:'nowrap'}}>{label}</span>;
}

// Accounts offered for an A/R invoice/credit-memo line. Revenue first (the usual
// case), then Expense, then balance-sheet accounts (Asset/Liability/Equity), so
// any GL account is selectable while the common choice stays at the top. Within
// each group, sort by code for a predictable order.
function arLineAccounts(accounts){
  const order={Revenue:0,Income:0,Expense:1,Asset:2,Liability:3,Equity:4};
  const rank=a=>(order[a.type]!==undefined?order[a.type]:5);
  return (accounts||[]).slice().sort((a,b)=>{
    const r=rank(a)-rank(b); if(r!==0) return r;
    return String(a.code).localeCompare(String(b.code),undefined,{numeric:true});
  });
}

// Shared line-item editor for both one-off invoices and recurring templates.
function ArLines({lines,setLines,revAccts,classes,locations,projects,dimsEnabled,isBanyanRes}){
  const upd=(i,k,v)=>setLines(ls=>ls.map((l,ix)=>ix===i?{...l,[k]:v}:l));
  const add=()=>setLines(ls=>[...ls,{description:'',qty:1,rate:'',revenue_account_code:ls.length?ls[ls.length-1].revenue_account_code:'',class_id:'',location_id:'',project_id:''}]);
  const rm=i=>setLines(ls=>ls.length<=1?ls:ls.filter((_,ix)=>ix!==i));
  const total=lines.reduce((s,l)=>s+(Number(l.qty)||0)*(Number(l.rate)||0),0);
  // Banyan Residential codes every transaction to a project (its only dimension),
  // so instead of the empty Location/Class dropdowns it gets a single "Dimension"
  // dropdown listing all project codes. The chosen project flows through to the
  // posted journal entry's revenue line.
  const projDim=!!(isBanyanRes&&projects&&projects.length);
  const projLabel=p=>p.code&&p.code!==p.name?p.code+' — '+p.name:(p.code||p.name);
  return(<div>
    <div style={{overflowX:'auto'}}><table style={S.table}><thead><tr>
      <th style={S.th}>Description</th><th style={{...S.th,width:190}}>Revenue account</th>
      {projDim&&<th style={{...S.th,width:220}}>Dimension</th>}
      {!projDim&&dimsEnabled&&<th style={{...S.th,width:130}}>Location</th>}
      {!projDim&&dimsEnabled&&<th style={{...S.th,width:130}}>{classTerm()}</th>}
      <th style={{...S.thR,width:70}}>Qty</th><th style={{...S.thR,width:110}}>Rate</th><th style={{...S.thR,width:110}}>Amount</th><th style={{...S.th,width:28}}></th></tr></thead>
      <tbody>{lines.map((l,i)=><tr key={i}>
        <td style={{padding:'4px 6px'}}><input style={{...S.inputSm,width:'100%'}} value={l.description} onChange={e=>upd(i,'description',e.target.value)} placeholder="e.g. Land lease — May 2026"/></td>
        <td style={{padding:'4px 6px'}}><select style={{...S.select,padding:'6px 8px',fontSize:12}} value={l.revenue_account_code} onChange={e=>upd(i,'revenue_account_code',e.target.value)}>
          <option value="">— select —</option>{revAccts.map(a=><option key={a.code} value={a.code}>{a.code} {a.name}</option>)}</select></td>
        {projDim&&<td style={{padding:'4px 6px'}}><select style={{...S.select,padding:'6px 8px',fontSize:12}} value={l.project_id||''} onChange={e=>upd(i,'project_id',e.target.value)}>
          <option value="">— select project —</option>{projects.map(p=><option key={p.id} value={p.id}>{projLabel(p)}</option>)}</select></td>}
        {!projDim&&dimsEnabled&&<td style={{padding:'4px 6px'}}><select style={{...S.select,padding:'6px 8px',fontSize:12}} value={l.location_id||''} onChange={e=>upd(i,'location_id',e.target.value)}>
          <option value="">—</option>{locations.map(o=><option key={o.id} value={o.id}>{o.name}</option>)}</select></td>}
        {!projDim&&dimsEnabled&&<td style={{padding:'4px 6px'}}><select style={{...S.select,padding:'6px 8px',fontSize:12}} value={l.class_id||''} onChange={e=>upd(i,'class_id',e.target.value)}>
          <option value="">—</option>{classes.map(o=><option key={o.id} value={o.id}>{o.name}</option>)}</select></td>}
        <td style={{padding:'4px 6px'}}><input style={{...S.inputSm,width:'100%',textAlign:'right'}} value={l.qty} onChange={e=>upd(i,'qty',e.target.value)}/></td>
        <td style={{padding:'4px 6px'}}><input style={{...S.inputSm,width:'100%',textAlign:'right'}} value={l.rate} onChange={e=>upd(i,'rate',e.target.value)} placeholder="0.00"/></td>
        <td style={{...S.td,textAlign:'right',color:T.textBright}}>{fmt((Number(l.qty)||0)*(Number(l.rate)||0))}</td>
        <td style={S.td}>{lines.length>1&&<button style={{...S.btnGhost,color:T.red,fontSize:11,padding:'2px 4px'}} onClick={()=>rm(i)}>x</button>}</td></tr>)}
      </tbody></table></div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginTop:8}}>
      <button style={{...S.btnGhost,color:T.accent,fontSize:12}} onClick={add}>+ Add line</button>
      <div style={{fontSize:14,fontWeight:700,color:T.textBright}}>Total {fmt(total)}</div></div>
  </div>);
}

// Per-entity invoice presentation + numbering settings.
function ArSettingsPanel({entityId,accounts,onClose,onSaved}){
  const[f,setF]=useState(null);const[err,setErr]=useState('');const[saving,setSaving]=useState(false);
  useEffect(()=>{(async()=>{try{const s=await api.getArSettings(entityId);setF({bill_from:s.bill_from||'',remit_to:s.remit_to||'',footer_note:s.footer_note||'',reply_to:s.reply_to||'',invoice_prefix:s.invoice_prefix||'',default_ar_account:s.default_ar_account||'',resolved:s.resolved_ar_account,email_configured:s.email_configured});}catch(e){setErr(e.message);}})();},[entityId]);
  const save=async()=>{setSaving(true);setErr('');try{await api.saveArSettings(entityId,{bill_from:f.bill_from,remit_to:f.remit_to,footer_note:f.footer_note,reply_to:f.reply_to,invoice_prefix:f.invoice_prefix,default_ar_account:f.default_ar_account});onSaved&&onSaved();onClose();}catch(e){setErr(e.message);}finally{setSaving(false);}};
  if(!f)return <div style={{...S.card,color:T.textMuted}}>Loading settings…</div>;
  const arAccts=accounts.filter(a=>a.type==='Asset');
  return(<div style={{...S.card,borderColor:T.accent+'40',marginBottom:16}}>
    <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:12}}>Invoice Settings</div>
    <div style={{display:'flex',gap:14,flexWrap:'wrap'}}>
      <div style={{flex:'1 1 260px'}}><label style={S.label}>Bill-from block (top-left of the invoice)</label>
        <textarea style={{...S.input,minHeight:70,resize:'vertical'}} value={f.bill_from} onChange={e=>setF(v=>({...v,bill_from:e.target.value}))} placeholder={'Legal entity name, LLC\n123 Main St\nCity, TX 77000'}/></div>
      <div style={{flex:'1 1 260px'}}><label style={S.label}>Remit-to block (invoice footer)</label>
        <textarea style={{...S.input,minHeight:70,resize:'vertical'}} value={f.remit_to} onChange={e=>setF(v=>({...v,remit_to:e.target.value}))} placeholder={'Wire: Bank name\nRouting / Account'}/></div>
    </div>
    <div style={{display:'flex',gap:14,flexWrap:'wrap',marginTop:10}}>
      <div style={{flex:'0 0 150px'}}><label style={S.label}>Invoice prefix</label><input style={S.input} value={f.invoice_prefix} onChange={e=>setF(v=>({...v,invoice_prefix:e.target.value}))} placeholder="INV"/></div>
      <div style={{flex:'1 1 220px'}}><label style={S.label}>A/R account</label>
        <select style={S.select} value={f.default_ar_account} onChange={e=>setF(v=>({...v,default_ar_account:e.target.value}))}>
          <option value="">Auto-detect ({f.resolved||'none found'})</option>{arAccts.map(a=><option key={a.code} value={a.code}>{a.code} {a.name}</option>)}</select></div>
      <div style={{flex:'1 1 220px'}}><label style={S.label}>Reply-to email</label><input style={S.input} type="email" value={f.reply_to} onChange={e=>setF(v=>({...v,reply_to:e.target.value}))} placeholder="ar@banyanres.com"/></div>
    </div>
    <div style={{marginTop:10}}><label style={S.label}>Footer note (payment terms, late fees)</label>
      <input style={S.input} value={f.footer_note} onChange={e=>setF(v=>({...v,footer_note:e.target.value}))} placeholder="Payment due within 30 days of invoice date."/></div>
    {!f.email_configured&&<div style={{marginTop:10,fontSize:12,color:T.orange}}>Email sending is not configured on the server (RESEND_API_KEY). You can still generate and download invoice PDFs; the Send button will report the error until the key is set.</div>}
    {err&&<div style={{...S.err,marginTop:8}}>{err}</div>}
    <div style={{display:'flex',gap:10,marginTop:14}}><button style={S.btnP} onClick={save} disabled={saving}>{saving?'Saving…':'Save Settings'}</button><button style={S.btnS} onClick={onClose}>Cancel</button></div>
  </div>);
}

function ArInvoices({entityId,entityName,canEdit,dimsEnabled,isBanyanRes}){
  const[customers,setCustomers]=useState([]);const[accounts,setAccounts]=useState([]);
  const[classes,setClasses]=useState([]);const[locations,setLocations]=useState([]);const[projects,setProjects]=useState([]);
  const[invoices,setInvoices]=useState([]);const[loading,setLoading]=useState(true);
  const[statusF,setStatusF]=useState('');const[err,setErr]=useState('');
  const[showNew,setShowNew]=useState(false);const[showSettings,setShowSettings]=useState(false);
  const[showCM,setShowCM]=useState(false);
  const[detail,setDetail]=useState(null);const[busy,setBusy]=useState('');
  const blankLine={description:'',qty:1,rate:'',revenue_account_code:'',class_id:'',location_id:'',project_id:''};
  const[form,setForm]=useState({customer_id:'',invoice_num:'',invoice_date:today(),due_date:'',memo:''});
  const[lines,setLines]=useState([{...blankLine}]);
  const[cmForm,setCmForm]=useState({customer_id:'',invoice_date:today(),memo:''});
  const[cmLines,setCmLines]=useState([{...blankLine}]);
  // Revenue first (the usual choice for an invoice line), then Expense, then
  // balance-sheet accounts, so any GL account can be chosen (e.g. a contra-
  // revenue or a balance-sheet clearing account) while the common case stays on top.
  const revAccts=arLineAccounts(accounts);
  const bankAccts=accounts.filter(a=>a.bank_acct);
  const loadRefs=useCallback(async()=>{
    const[c,a,cl,lo,pr]=await Promise.all([api.getArCustomers(entityId),api.getAccounts(entityId),
      api.getClasses(entityId).catch(()=>[]),api.getLocations(entityId).catch(()=>[]),api.getProjects(entityId).catch(()=>[])]);
    setCustomers((c||[]).filter(x=>x.active));setAccounts(a||[]);setClasses(cl||[]);setLocations(lo||[]);setProjects(pr||[]);
  },[entityId]);
  const load=useCallback(async()=>{setLoading(true);setErr('');
    try{setInvoices(await api.getArInvoices(entityId,statusF?{status:statusF}:{})||[]);}
    catch(e){setErr(e.message);}finally{setLoading(false);}},[entityId,statusF]);
  useEffect(()=>{loadRefs();},[loadRefs]);
  useEffect(()=>{load();},[load]);
  const resetForm=()=>{setForm({customer_id:'',invoice_num:'',invoice_date:today(),due_date:'',memo:''});setLines([{...blankLine}]);};
  const resetCm=()=>{setCmForm({customer_id:'',invoice_date:today(),memo:''});setCmLines([{...blankLine}]);};
  const createCm=async()=>{
    setErr('');
    if(!cmForm.customer_id){setErr('Pick a customer for the credit memo');return;}
    setBusy('createcm');
    try{
      const cm=await api.createCreditMemo(entityId,{customer_id:+cmForm.customer_id,invoice_date:cmForm.invoice_date,memo:cmForm.memo,
        lines:cmLines.map(l=>({description:l.description,qty:Number(l.qty),rate:Number(l.rate),
          revenue_account_code:l.revenue_account_code,class_id:l.class_id?+l.class_id:null,location_id:l.location_id?+l.location_id:null}))});
      resetCm();setShowCM(false);await load();setDetail(cm);
    }catch(e){setErr(e.message);}finally{setBusy('');}
  };
  const create=async()=>{
    setErr('');
    if(!form.customer_id){setErr('Pick a customer');return;}
    setBusy('create');
    try{
      const inv=await api.createArInvoice(entityId,{customer_id:+form.customer_id,invoice_num:form.invoice_num.trim()||undefined,invoice_date:form.invoice_date,
        due_date:form.due_date||undefined,memo:form.memo,
        lines:lines.map(l=>({description:l.description,qty:Number(l.qty),rate:Number(l.rate),
          revenue_account_code:l.revenue_account_code,class_id:l.class_id?+l.class_id:null,location_id:l.location_id?+l.location_id:null,project_id:l.project_id?+l.project_id:null}))});
      resetForm();setShowNew(false);await load();setDetail(inv);
    }catch(e){setErr(e.message);}finally{setBusy('');}
  };
  const open=async(inv)=>{try{setDetail(await api.getArInvoice(entityId,inv.id));}catch(e){alert(e.message);}};
  const act=async(fn,label)=>{setBusy(label);try{const r=await fn();await load();if(detail)setDetail(await api.getArInvoice(entityId,detail.id));return r;}catch(e){alert(e.message);}finally{setBusy('');}};
  const totals=invoices.reduce((s,i)=>{if(i.status!=='void'){s.billed+=Number(i.total)||0;s.open+=Number(i.open_amount)||0;}return s;},{billed:0,open:0});
  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
      <div><div style={S.h1}>Invoices</div><div style={S.sub}>{entityName} — each invoice posts Dr A/R / Cr Revenue when created, then you review and send.</div></div>
      <div style={{display:'flex',gap:8}}>
        {canEdit&&<button style={S.btnS} onClick={()=>setShowSettings(s=>!s)}>Settings</button>}
        {canEdit&&<button style={S.btnP} onClick={()=>{setShowNew(n=>!n);setShowCM(false);setErr('');}}>{showNew?'Cancel':'+ New Invoice'}</button>}
        {canEdit&&<button style={S.btnS} onClick={()=>{setShowCM(n=>!n);setShowNew(false);setErr('');}}>{showCM?'Cancel':'+ New Credit Memo'}</button>}</div></div>
    {showSettings&&<ArSettingsPanel entityId={entityId} accounts={accounts} onClose={()=>setShowSettings(false)} onSaved={load}/>}
    {showNew&&<div style={{...S.card,borderColor:T.green+'40',marginBottom:16}}>
      <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 240px'}}><label style={S.label}>Customer</label>
          <select style={S.select} value={form.customer_id} onChange={e=>setForm(f=>({...f,customer_id:e.target.value}))}>
            <option value="">— select —</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
        <div style={{flex:'0 0 180px'}}><label style={S.label}>Invoice # (blank = auto)</label><input style={S.input} value={form.invoice_num} onChange={e=>setForm(f=>({...f,invoice_num:e.target.value}))} placeholder="Auto-generated"/></div>
        <div style={{flex:'0 0 160px'}}><label style={S.label}>Invoice date</label><input style={S.input} type="date" value={form.invoice_date} onChange={e=>setForm(f=>({...f,invoice_date:e.target.value}))}/></div>
        <div style={{flex:'0 0 160px'}}><label style={S.label}>Due date (blank = terms)</label><input style={S.input} type="date" value={form.due_date} onChange={e=>setForm(f=>({...f,due_date:e.target.value}))}/></div>
        <div style={{flex:'1 1 240px'}}><label style={S.label}>Memo / "Re:" line</label><input style={S.input} value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))} placeholder="April 2026 services"/></div>
      </div>
      <div style={{marginTop:14}}><ArLines lines={lines} setLines={setLines} revAccts={revAccts} classes={classes} locations={locations} projects={projects} isBanyanRes={isBanyanRes} dimsEnabled={dimsEnabled}/></div>
      {err&&<div style={{...S.err,marginTop:10}}>{err}</div>}
      {revAccts.length===0&&<div style={{marginTop:10,fontSize:12,color:T.orange}}>This entity has no Revenue accounts yet — add one in the Chart of Accounts first.</div>}
      <div style={{display:'flex',justifyContent:'flex-end',marginTop:12}}><button style={S.btnP} onClick={create} disabled={busy==='create'}>{busy==='create'?'Posting…':'Create Invoice (posts JE)'}</button></div>
    </div>}
    {showCM&&<div style={{...S.card,borderColor:T.orange+'55',marginBottom:16}}>
      <div style={{fontSize:13,fontWeight:700,color:T.textBright,marginBottom:4}}>New Credit Memo</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:12}}>A credit memo posts the reverse of an invoice (Dr Revenue / Cr A/R). Enter amounts as positive numbers; it is recorded as a credit. Apply it to an open invoice from that invoice's detail view.</div>
      <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 240px'}}><label style={S.label}>Customer</label>
          <select style={S.select} value={cmForm.customer_id} onChange={e=>setCmForm(f=>({...f,customer_id:e.target.value}))}>
            <option value="">— select —</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
        <div style={{flex:'0 0 160px'}}><label style={S.label}>Credit memo date</label><input style={S.input} type="date" value={cmForm.invoice_date} onChange={e=>setCmForm(f=>({...f,invoice_date:e.target.value}))}/></div>
        <div style={{flex:'1 1 240px'}}><label style={S.label}>Memo / reason</label><input style={S.input} value={cmForm.memo} onChange={e=>setCmForm(f=>({...f,memo:e.target.value}))} placeholder="Credit for overbilling on INV-2026-0001"/></div>
      </div>
      <div style={{marginTop:14}}><ArLines lines={cmLines} setLines={setCmLines} revAccts={revAccts} classes={classes} locations={locations} dimsEnabled={dimsEnabled}/></div>
      {err&&<div style={{...S.err,marginTop:10}}>{err}</div>}
      <div style={{display:'flex',justifyContent:'flex-end',marginTop:12}}><button style={S.btnP} onClick={createCm} disabled={busy==='createcm'}>{busy==='createcm'?'Posting…':'Create Credit Memo (posts JE)'}</button></div>
    </div>}
    <div style={{...S.card,display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap',marginBottom:0}}>
      <div style={{flex:'0 0 170px'}}><label style={S.label}>Status</label>
        <select style={S.select} value={statusF} onChange={e=>setStatusF(e.target.value)}>
          <option value="">All</option><option value="draft">Draft</option><option value="sent">Sent</option><option value="paid">Paid</option><option value="void">Void</option></select></div>
      <div style={{fontSize:12,color:T.textMuted}}>Billed (ex-void) <span style={{color:T.textBright,fontWeight:600}}>{fmt(totals.billed)}</span>
        <span style={{margin:'0 10px'}}>·</span>Open <span style={{color:T.textBright,fontWeight:600}}>{fmt(totals.open)}</span></div>
    </div>
    {err&&!showNew&&<div style={S.err}>{err}</div>}
    <div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
      <th style={S.th}>Invoice #</th><th style={S.th}>Date</th><th style={S.th}>Customer</th><th style={S.th}>Memo</th>
      <th style={S.th}>Due</th><th style={{...S.th,width:100}}>Status</th><th style={S.thR}>Total</th><th style={S.thR}>Open</th></tr></thead>
      <tbody>
        {loading&&<tr><td colSpan={8} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>Loading…</td></tr>}
        {!loading&&invoices.length===0&&<tr><td colSpan={8} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>No invoices yet.</td></tr>}
        {!loading&&invoices.map(i=><tr key={i.id} style={{cursor:'pointer',...(i.status==='void'?{opacity:0.55}:{})}} onClick={()=>open(i)}>
          <td style={{...S.td,color:T.textBright,fontWeight:600}}>{i.invoice_num}</td>
          <td style={S.td}>{i.invoice_date}</td><td style={S.td}>{i.customer_name}</td>
          <td style={{...S.td,color:T.textMuted}}>{i.memo||''}</td><td style={S.td}>{i.due_date||''}</td>
          <td style={S.td}><ArBadge inv={i}/></td>
          <td style={{...S.td,textAlign:'right'}}>{fmt(i.total)}</td>
          <td style={{...S.td,textAlign:'right',color:i.open_amount>0.005?T.textBright:T.textMuted}}>{fmt(i.open_amount)}</td></tr>)}
      </tbody></table></div>
    {detail&&<ArInvoiceDetail entityId={entityId} invoice={detail} bankAccts={bankAccts} canEdit={canEdit} busy={busy} act={act}
      onClose={()=>setDetail(null)} onChanged={load}/>}
  </div>);
}

function ArInvoiceDetail({entityId,invoice,bankAccts,canEdit,busy,act,onClose}){
  const inv=invoice;
  const[sendTo,setSendTo]=useState(inv.customer_email||'');
  const[pay,setPay]=useState({date:today(),amount:'',bank_account_code:bankAccts[0]?bankAccts[0].code:'',memo:''});
  const isCm=inv.doc_type==='credit_memo';
  // Open credit memos for THIS invoice's customer, offered for application
  // when the invoice still has an open balance. Not loaded for a credit memo.
  const[credits,setCredits]=useState([]);
  const[applyCm,setApplyCm]=useState({credit_memo_id:'',amount:'',date:today()});
  const loadCredits=useCallback(async()=>{
    if(isCm||inv.open_amount<=0.005){setCredits([]);return;}
    try{const all=await api.getCreditMemos(entityId,true);
      setCredits((all||[]).filter(c=>c.customer_id===inv.customer_id&&c.remaining>0.005));}catch(_){setCredits([]);}
  },[entityId,inv.id,inv.customer_id,inv.open_amount,isCm]);
  useEffect(()=>{setSendTo(inv.customer_email||'');setPay(p=>({...p,amount:'',date:today()}));setApplyCm({credit_memo_id:'',amount:'',date:today()});loadCredits();},[inv.id,loadCredits]);
  const isDraft=inv.status==='draft';const isVoid=inv.status==='void';
  return(<div style={{position:'fixed',inset:0,background:'rgba(15,23,42,0.55)',display:'flex',justifyContent:'center',alignItems:'flex-start',zIndex:60,overflowY:'auto',padding:'40px 16px'}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
    <div style={{background:'#fff',borderRadius:12,maxWidth:820,width:'100%',padding:22,boxShadow:'0 20px 60px rgba(0,0,0,0.3)'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:6}}>
        <div><div style={{fontSize:18,fontWeight:700,color:T.textBright}}>{inv.invoice_num} <ArBadge inv={inv}/></div>
          <div style={{fontSize:13,color:T.textMuted,marginTop:4}}>{inv.customer_name}{inv.customer_email?' · '+inv.customer_email:''}</div></div>
        <button style={S.btnGhost} onClick={onClose}>Close</button></div>
      <div style={{display:'flex',gap:24,flexWrap:'wrap',fontSize:12,color:T.textMuted,margin:'12px 0 16px'}}>
        <div>Invoice date<div style={{color:T.textBright,fontSize:13}}>{inv.invoice_date}</div></div>
        <div>Due date<div style={{color:T.textBright,fontSize:13}}>{inv.due_date||'—'}</div></div>
        <div>A/R account<div style={{color:T.textBright,fontSize:13}}>{inv.ar_account_code}</div></div>
        <div>Total<div style={{color:T.textBright,fontSize:13}}>{fmt(inv.total)}</div></div>
        <div>Paid<div style={{color:T.textBright,fontSize:13}}>{fmt(inv.paid_amount)}</div></div>
        <div>Open<div style={{color:T.textBright,fontSize:13,fontWeight:700}}>{fmt(inv.open_amount)}</div></div>
        {inv.sent_at&&<div>Sent<div style={{color:T.textBright,fontSize:13}}>{String(inv.sent_at).slice(0,10)}</div></div>}
      </div>
      {inv.memo&&<div style={{fontSize:13,marginBottom:12}}>Re: {inv.memo}</div>}
      <table style={{...S.table,marginBottom:16}}><thead><tr>
        <th style={S.th}>Description</th><th style={S.th}>Account</th><th style={S.thR}>Qty</th><th style={S.thR}>Rate</th><th style={S.thR}>Amount</th></tr></thead>
        <tbody>{(inv.lines||[]).map(l=><tr key={l.id}>
          <td style={S.td}>{l.description}</td><td style={S.td}>{l.revenue_account_code}</td>
          <td style={{...S.td,textAlign:'right'}}>{l.qty}</td><td style={{...S.td,textAlign:'right'}}>{fmt(l.rate)}</td>
          <td style={{...S.td,textAlign:'right'}}>{fmt(l.amount)}</td></tr>)}</tbody></table>
      {(inv.receipts||[]).length>0&&<div style={{marginBottom:16}}>
        <div style={{fontSize:12,fontWeight:600,color:T.textMuted,marginBottom:6}}>{isCm?'APPLIED TO INVOICES':'PAYMENTS & CREDITS APPLIED'}</div>
        <table style={S.table}><tbody>{inv.receipts.map(r=>{const ca=r.kind==='credit_application';return <tr key={r.id}>
          <td style={S.td}>{r.date}</td><td style={S.td}>{ca?<span style={{color:T.orange}}>Credit applied</span>:r.bank_account_code}</td><td style={{...S.td,color:T.textMuted}}>{r.memo||''}</td>
          <td style={{...S.td,textAlign:'right'}}>{fmt(r.amount)}</td>
          {canEdit&&<td style={{...S.td,width:60}}><button style={{...S.btnGhost,color:T.red,fontSize:11}} disabled={!!busy}
            onClick={()=>{if(ca){if(confirm('Remove this credit application? The credit returns to the memo and the invoice reopens.'))act(()=>api.deleteCreditApplication(entityId,r.id),'rmca');}else{if(confirm('Delete this payment and its journal entry?'))act(()=>api.deleteArReceipt(entityId,inv.id,r.id),'rmrec');}}}>x</button></td>}</tr>;})}</tbody></table></div>}
      <div style={{display:'flex',gap:8,flexWrap:'wrap',marginBottom:16}}>
        <button style={S.btnS} onClick={()=>window.open(api.arInvoicePdfUrl(entityId,inv.id),'_blank')}>View PDF</button>
        {canEdit&&!isVoid&&<button style={S.btnS} disabled={!!busy} onClick={()=>act(()=>api.saveArInvoicePdf(entityId,inv.id),'savepdf')}>{busy==='savepdf'?'Filing…':'File to Workpapers'}</button>}
        {canEdit&&isDraft&&<button style={S.btnS} disabled={!!busy} onClick={()=>act(()=>api.markArInvoiceSent(entityId,inv.id),'marksent')}>Mark Sent (no email)</button>}
        {canEdit&&isDraft&&<button style={{...S.btnGhost,color:T.red}} disabled={!!busy}
          onClick={()=>{if(confirm('Delete draft '+inv.invoice_num+' and its journal entry?'))act(()=>api.deleteArInvoice(entityId,inv.id),'del').then(()=>onClose());}}>Delete Draft</button>}
        {canEdit&&!isVoid&&<button style={{...S.btnGhost,color:T.orange}} disabled={!!busy}
          onClick={()=>{if(confirm(isDraft?'Void this draft? Its journal entry will be removed.':'Void '+inv.invoice_num+'? A reversing journal entry will be posted.'))act(()=>api.voidArInvoice(entityId,inv.id),'void');}}>Void</button>}
      </div>
      {canEdit&&!isVoid&&<div style={{...S.card,marginBottom:12}}>
        <div style={{fontSize:13,fontWeight:600,color:T.textBright,marginBottom:8}}>Send to customer</div>
        <div style={{display:'flex',gap:10,alignItems:'flex-end',flexWrap:'wrap'}}>
          <div style={{flex:'1 1 260px'}}><label style={S.label}>Recipient email</label><input style={S.input} type="email" value={sendTo} onChange={e=>setSendTo(e.target.value)} placeholder="ar@customer.com"/></div>
          <button style={S.btnP} disabled={!!busy||!sendTo.trim()} onClick={()=>act(()=>api.sendArInvoice(entityId,inv.id,{to:sendTo.trim()}),'send')}>{busy==='send'?'Sending…':'Email Invoice'}</button></div>
        <div style={{fontSize:11,color:T.textMuted,marginTop:8}}>Attaches the PDF, files a copy under Workpapers &gt; Invoices/{String(inv.invoice_date).slice(0,4)}, and marks the invoice sent.</div>
      </div>}
      {canEdit&&!isVoid&&!isCm&&inv.open_amount>0.005&&credits.length>0&&<div style={{...S.card,marginBottom:12,borderColor:T.orange+'55'}}>
        <div style={{fontSize:13,fontWeight:600,color:T.textBright,marginBottom:8}}>Apply a credit memo</div>
        <div style={{display:'flex',gap:10,alignItems:'flex-end',flexWrap:'wrap'}}>
          <div style={{flex:'1 1 260px'}}><label style={S.label}>Credit memo</label>
            <select style={S.select} value={applyCm.credit_memo_id} onChange={e=>{const id=e.target.value;const c=credits.find(x=>String(x.id)===String(id));setApplyCm(a=>({...a,credit_memo_id:id,amount:c?String(Math.min(c.remaining,inv.open_amount)):''}));}}>
              <option value="">— select —</option>{credits.map(c=><option key={c.id} value={c.id}>{c.invoice_num} · {fmt(c.remaining)} available</option>)}</select></div>
          <div style={{flex:'0 0 140px'}}><label style={S.label}>Amount</label><input style={{...S.input,textAlign:'right'}} value={applyCm.amount} onChange={e=>setApplyCm(a=>({...a,amount:e.target.value}))} placeholder={String(inv.open_amount)}/></div>
          <div style={{flex:'0 0 150px'}}><label style={S.label}>Date</label><input style={S.input} type="date" value={applyCm.date} onChange={e=>setApplyCm(a=>({...a,date:e.target.value}))}/></div>
          <button style={S.btnP} disabled={!!busy||!applyCm.credit_memo_id} onClick={()=>act(()=>api.applyCreditMemo(entityId,applyCm.credit_memo_id,{invoice_id:inv.id,amount:applyCm.amount===''?undefined:Number(applyCm.amount),date:applyCm.date}),'applycm').then(loadCredits)}>{busy==='applycm'?'Applying…':'Apply Credit'}</button></div>
        <div style={{fontSize:11,color:T.textMuted,marginTop:8}}>Applies the credit to this invoice's open balance. No journal entry — the credit memo already posted Dr Revenue / Cr A/R when it was created.</div>
      </div>}
      {canEdit&&!isVoid&&inv.open_amount>0.005&&<div style={S.card}>
        <div style={{fontSize:13,fontWeight:600,color:T.textBright,marginBottom:8}}>Record payment</div>
        <div style={{display:'flex',gap:10,alignItems:'flex-end',flexWrap:'wrap'}}>
          <div style={{flex:'0 0 150px'}}><label style={S.label}>Date</label><input style={S.input} type="date" value={pay.date} onChange={e=>setPay(p=>({...p,date:e.target.value}))}/></div>
          <div style={{flex:'0 0 140px'}}><label style={S.label}>Amount</label><input style={{...S.input,textAlign:'right'}} value={pay.amount} onChange={e=>setPay(p=>({...p,amount:e.target.value}))} placeholder={String(inv.open_amount)}/></div>
          <div style={{flex:'1 1 200px'}}><label style={S.label}>Bank account</label>
            <select style={S.select} value={pay.bank_account_code} onChange={e=>setPay(p=>({...p,bank_account_code:e.target.value}))}>
              <option value="">— select —</option>{bankAccts.map(a=><option key={a.code} value={a.code}>{a.code} {a.name}</option>)}</select></div>
          <div style={{flex:'1 1 180px'}}><label style={S.label}>Memo</label><input style={S.input} value={pay.memo} onChange={e=>setPay(p=>({...p,memo:e.target.value}))} placeholder="Check 10482"/></div>
          <button style={S.btnP} disabled={!!busy||!pay.bank_account_code} onClick={()=>act(()=>api.addArReceipt(entityId,inv.id,{date:pay.date,amount:pay.amount===''?undefined:Number(pay.amount),bank_account_code:pay.bank_account_code,memo:pay.memo}),'pay')}>{busy==='pay'?'Posting…':'Post Receipt'}</button></div>
        <div style={{fontSize:11,color:T.textMuted,marginTop:8}}>Blank amount pays the full open balance. Posts Dr {pay.bank_account_code||'bank'} / Cr {inv.ar_account_code}.</div>
        {bankAccts.length===0&&<div style={{fontSize:12,color:T.orange,marginTop:8}}>No account is flagged as a bank/cash account in this entity's Chart of Accounts.</div>}
      </div>}
    </div></div>);
}

function ArRecurring({entityId,entityName,canEdit,dimsEnabled}){
  const[templates,setTemplates]=useState([]);const[customers,setCustomers]=useState([]);
  const[accounts,setAccounts]=useState([]);const[classes,setClasses]=useState([]);const[locations,setLocations]=useState([]);
  const[loading,setLoading]=useState(true);const[err,setErr]=useState('');const[busy,setBusy]=useState('');
  const[showNew,setShowNew]=useState(false);const[editId,setEditId]=useState(null);
  const blankLine={description:'',qty:1,rate:'',revenue_account_code:'',class_id:'',location_id:''};
  const[form,setForm]=useState({customer_id:'',memo:'',frequency:'monthly',day_of_month:1,next_run:''});
  const[lines,setLines]=useState([{...blankLine}]);
  const revAccts=arLineAccounts(accounts);
  const load=useCallback(async()=>{setLoading(true);setErr('');
    try{const[t,c,a,cl,lo]=await Promise.all([api.getArTemplates(entityId),api.getArCustomers(entityId),api.getAccounts(entityId),
      api.getClasses(entityId).catch(()=>[]),api.getLocations(entityId).catch(()=>[])]);
      setTemplates(t||[]);setCustomers((c||[]).filter(x=>x.active));setAccounts(a||[]);setClasses(cl||[]);setLocations(lo||[]);
    }catch(e){setErr(e.message);}finally{setLoading(false);}},[entityId]);
  useEffect(()=>{load();},[load]);
  const reset=()=>{setForm({customer_id:'',memo:'',frequency:'monthly',day_of_month:1,next_run:''});setLines([{...blankLine}]);setEditId(null);setShowNew(false);};
  const payload=()=>({customer_id:+form.customer_id,memo:form.memo,frequency:form.frequency,day_of_month:+form.day_of_month||1,
    next_run:form.next_run||undefined,
    lines:lines.map(l=>({description:l.description,qty:Number(l.qty),rate:Number(l.rate),revenue_account_code:l.revenue_account_code,
      class_id:l.class_id?+l.class_id:null,location_id:l.location_id?+l.location_id:null}))});
  const save=async()=>{
    setErr('');if(!form.customer_id){setErr('Pick a customer');return;}
    setBusy('save');
    try{if(editId)await api.updateArTemplate(entityId,editId,payload());else await api.createArTemplate(entityId,payload());reset();await load();}
    catch(e){setErr(e.message);}finally{setBusy('');}
  };
  const startEdit=t=>{setEditId(t.id);setShowNew(true);setErr('');
    setForm({customer_id:String(t.customer_id),memo:t.memo||'',frequency:t.frequency||'monthly',day_of_month:t.day_of_month||1,next_run:t.next_run||''});
    setLines((t.lines||[]).map(l=>({description:l.description,qty:l.qty,rate:l.rate,revenue_account_code:l.revenue_account_code,class_id:l.class_id||'',location_id:l.location_id||''})));};
  const run=async(t)=>{const d=prompt('Invoice date for this run:',t.next_run||today());if(!d)return;
    setBusy('gen'+t.id);
    try{const inv=await api.generateArInvoice(entityId,t.id,d);await load();alert('Created '+inv.invoice_num+' for '+fmt(inv.total)+' (draft). Open Invoices to review and send.');}
    catch(e){alert(e.message);}finally{setBusy('');}};
  const dueNow=t=>t.active&&t.next_run&&t.next_run<=today();
  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:16}}>
      <div><div style={S.h1}>Recurring Invoices</div><div style={S.sub}>{entityName} — templates you run each period. Generating creates a draft invoice and posts its accrual entry.</div></div>
      {canEdit&&<button style={S.btnP} onClick={()=>{if(showNew)reset();else{setShowNew(true);setErr('');}}}>{showNew?'Cancel':'+ New Template'}</button>}</div>
    {showNew&&<div style={{...S.card,borderColor:T.green+'40',marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:600,color:T.textBright,marginBottom:12}}>{editId?'Edit template':'New recurring template'}</div>
      <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
        <div style={{flex:'1 1 220px'}}><label style={S.label}>Customer</label>
          <select style={S.select} value={form.customer_id} onChange={e=>setForm(f=>({...f,customer_id:e.target.value}))}>
            <option value="">— select —</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
        <div style={{flex:'0 0 150px'}}><label style={S.label}>Frequency</label>
          <select style={S.select} value={form.frequency} onChange={e=>setForm(f=>({...f,frequency:e.target.value}))}>
            <option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="annual">Annual</option></select></div>
        <div style={{flex:'0 0 120px'}}><label style={S.label}>Day of month</label><input style={S.input} type="number" min="1" max="28" value={form.day_of_month} onChange={e=>setForm(f=>({...f,day_of_month:e.target.value}))}/></div>
        <div style={{flex:'0 0 160px'}}><label style={S.label}>Next run</label><input style={S.input} type="date" value={form.next_run} onChange={e=>setForm(f=>({...f,next_run:e.target.value}))}/></div>
        <div style={{flex:'1 1 220px'}}><label style={S.label}>Memo</label><input style={S.input} value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))} placeholder="Monthly land lease"/></div>
      </div>
      <div style={{marginTop:14}}><ArLines lines={lines} setLines={setLines} revAccts={revAccts} classes={classes} locations={locations} dimsEnabled={dimsEnabled}/></div>
      {err&&<div style={{...S.err,marginTop:10}}>{err}</div>}
      <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:12}}>
        <button style={S.btnS} onClick={reset}>Cancel</button>
        <button style={S.btnP} onClick={save} disabled={busy==='save'}>{busy==='save'?'Saving…':(editId?'Save Template':'Create Template')}</button></div>
    </div>}
    {err&&!showNew&&<div style={S.err}>{err}</div>}
    <div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
      <th style={S.th}>Customer</th><th style={S.th}>Memo</th><th style={S.th}>Frequency</th><th style={S.th}>Next run</th>
      <th style={S.thR}>Amount</th><th style={{...S.th,width:90}}>Status</th>{canEdit&&<th style={{...S.th,width:210}}>Actions</th>}</tr></thead>
      <tbody>
        {loading&&<tr><td colSpan={canEdit?7:6} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>Loading…</td></tr>}
        {!loading&&templates.length===0&&<tr><td colSpan={canEdit?7:6} style={{...S.td,textAlign:'center',color:T.textMuted,padding:18}}>No recurring templates yet.</td></tr>}
        {!loading&&templates.map(t=><tr key={t.id} style={t.active?undefined:{opacity:0.55}}>
          <td style={{...S.td,color:T.textBright}}>{t.customer_name}</td>
          <td style={{...S.td,color:T.textMuted}}>{t.memo||''}</td>
          <td style={S.td}>{t.frequency}{t.frequency!=='annual'?' (day '+t.day_of_month+')':''}</td>
          <td style={S.td}>{t.next_run||'—'}{dueNow(t)&&<span style={{marginLeft:6,fontSize:11,color:T.orange,fontWeight:600}}>due</span>}</td>
          <td style={{...S.td,textAlign:'right'}}>{fmt(t.amount)}</td>
          <td style={S.td}><span style={{fontSize:12,color:t.active?T.green:T.textMuted}}>{t.active?'Active':'Paused'}</span></td>
          {canEdit&&<td style={S.td}><div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
            <button style={{...S.btnGhost,color:T.green,fontSize:11}} disabled={!!busy} onClick={()=>run(t)}>{busy==='gen'+t.id?'…':'Generate'}</button>
            <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>startEdit(t)}>Edit</button>
            <button style={{...S.btnGhost,fontSize:11}} onClick={async()=>{try{await api.updateArTemplate(entityId,t.id,{active:t.active?0:1});load();}catch(e){alert(e.message);}}}>{t.active?'Pause':'Resume'}</button>
            <button style={{...S.btnGhost,color:T.red,fontSize:11}} onClick={async()=>{if(!confirm('Delete this template? Invoices already generated are not affected.'))return;try{await api.deleteArTemplate(entityId,t.id);load();}catch(e){alert(e.message);}}}>x</button>
          </div></td>}</tr>)}
      </tbody></table></div>
  </div>);
}

function ArOpeningUploadModal({entityId,asOf,onClose,onImported}){
  const [parsing,setParsing]=useState(false);
  const [preview,setPreview]=useState(null);
  const [err,setErr]=useState('');
  const [force,setForce]=useState(false);
  const [allowOver,setAllowOver]=useState(false);
  const [saving,setSaving]=useState(false);
  const [fileName,setFileName]=useState('');
  const fmtDate=(v)=>{ if(v===null||v===undefined||v==='')return null; if(v instanceof Date&&!isNaN(v)){return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');} const s=String(v).trim(); if(!s)return null; if(s.indexOf('/')>=0){const p=s.split(' ')[0].split('/'); if(p.length===3){let y=p[2]; if(y.length===2)y='20'+y; return y+'-'+p[0].padStart(2,'0')+'-'+p[1].padStart(2,'0');}} if(s.length>=10&&s.charAt(4)==='-')return s.slice(0,10); return null; };
  const num=(v)=>{ if(v===null||v===undefined||v==='')return null; if(typeof v==='number')return v; const n=parseFloat(String(v).split(',').join('').split('$').join('')); return isNaN(n)?null:n; };
  const onFile=async(file)=>{
    setErr('');setPreview(null);setParsing(true);setAllowOver(false);setFileName(file.name);
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
    try{ const res=await api.arOpeningImport(entityId,preview.items,force,{as_of:preview.asOf,allow_over_gl:allowOver}); onImported(res); }
    catch(e){ setErr(e.message||String(e)); } finally{ setSaving(false); }
  };
  const tie=preview&&preview.recon!==null?Math.abs(preview.recon)<0.005:null;
  // recon = GL control balance - file total. Negative means the detail claims MORE
  // open A/R than the control account holds, which is never a valid subledger.
  const overGl=!!(preview&&preview.recon!==null&&preview.recon<-0.005);
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
      {preview.recon!==null&&!tie&&!overGl&&<div style={{fontSize:12,color:T.orange,marginTop:6}}>This detail does not tie to the GL control balance as of {preview.asOf}. You can still import; the aging will carry the difference as an un-itemized residual.</div>}
      {overGl&&<div style={{fontSize:12,color:T.red,marginTop:6,fontWeight:600}}>This detail is ${fmt(Math.abs(preview.recon))} MORE than the GL control balance as of {preview.asOf}. A subledger cannot hold more open A/R than its control account. This usually means the wrong entity is selected, or this file belongs to a different entity. Check the entity selector before importing.</div>}
    </div>}
    {preview&&<label style={{display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.textMuted,marginBottom:12}}><input type="checkbox" style={S.checkbox} checked={force} onChange={e=>setForce(e.target.checked)}/>Replace even if cash receipts were already applied to opening items</label>}
    {overGl&&<label style={{display:'flex',alignItems:'center',gap:8,fontSize:12,color:T.red,marginBottom:12}}><input type="checkbox" style={S.checkbox} checked={allowOver} onChange={e=>setAllowOver(e.target.checked)}/>I checked the entity and the file &mdash; import anyway</label>}
    <div style={{display:'flex',gap:10,justifyContent:'flex-end'}}>
      <button style={S.btnS} onClick={onClose} disabled={saving}>Cancel</button>
      <button style={{...S.btnP,opacity:(!preview||saving||(overGl&&!allowOver))?0.5:1}} onClick={doImport} disabled={!preview||saving||(overGl&&!allowOver)}>{saving?'Importing…':(preview?'Import '+preview.count+' items':'Import')}</button>
    </div>
  </div></div>);
}

function ArAgingReport({entityId,entityName}){
  const[asOf,setAsOf]=useState(today());
  const[data,setData]=useState(null);const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  const[showDetail,setShowDetail]=useState(true);
  const[viewEntry,setViewEntry]=useState(null);const[entryLoading,setEntryLoading]=useState(false);const[showUpload,setShowUpload]=useState(false);
  // GL rows carry only an entry id; fetch the full entry (with lines) before
  // opening the JE modal so the user can see/clear the legacy balance.
  const openEntry=async(id)=>{if(!id)return;setEntryLoading(true);try{setViewEntry(await api.getEntry(entityId,id));}catch(e){alert('Could not open entry: '+e.message);}finally{setEntryLoading(false);}};
  const BK=[['current','Current'],['d1_30','1-30'],['d31_60','31-60'],['d61_90','61-90'],['d90_plus','90+']];
  const arLabel=(data&&(data.ar_account||(data.ar_accounts||[]).join(', ')))||'12000';
  const run=async()=>{setLoading(true);setErr('');setData(null);
    try{setData(await api.getArAging(entityId,asOf));}catch(e){setErr(e.message);}finally{setLoading(false);}};
  useEffect(()=>{run();},[entityId]);
  const jeNum=n=>n!=null?'JE-'+String(n).padStart(4,'0'):'';
  const doExport=()=>{
    if(!data)return;
    const head=['Customer','Invoice #','Invoice date','Due date','Days past due','Current','1-30','31-60','61-90','90+','GL','Amount'];
    const d=[[entityName||'A/R Aging Detail'],['A/R Aging Detail — built from GL '+arLabel],['As of '+data.as_of],[],head];
    const F=[];const custTotRows=[];
    data.rows.forEach(r=>{
      const cFst=d.length;
      data.detail.filter(x=>x.customer===r.customer).forEach(x=>d.push([r.customer,x.invoice_num,x.invoice_date,x.due_date,x.days_past_due,
        x.bucket==='current'?x.open:'',x.bucket==='d1_30'?x.open:'',x.bucket==='d31_60'?x.open:'',x.bucket==='d61_90'?x.open:'',x.bucket==='d90_plus'?x.open:'','',x.open]));
      const cLst=d.length-1;const cT=d.length;d.push(['Total '+r.customer,'','','','',r.current,r.d1_30,r.d31_60,r.d61_90,r.d90_plus,'',r.total]);
      sumCols(F,cT,[5,6,7,8,9,11],cFst,cLst);custTotRows.push(cT);
    });
    let glTotRow=null;
    if((data.gl_rows||[]).length){
      d.push([]);d.push(['GL ENTRIES (imported / manual — not aged)']);
      const gFst=d.length;data.gl_rows.forEach(r=>d.push([r.memo||'GL detail import',jeNum(r.entry_num),r.date,'','','','','','','',r.amount,r.amount]));const gLst=d.length-1;
      glTotRow=d.length;d.push(['Total GL Entries','','','','','','','','','',data.gl_total,data.gl_total]);
      sumCols(F,glTotRow,[10,11],gFst,gLst);
    }
    const t=data.totals;
    const totRow=d.length;d.push(['TOTAL','','','','',t.current,t.d1_30,t.d31_60,t.d61_90,t.d90_plus,t.gl,t.total]);
    sumRows(F,totRow,[5,6,7,8,9,10,11],[...custTotRows,...(glTotRow!=null?[glTotRow]:[])]);
    d.push(['Reconciliation vs GL '+arLabel+' ('+fmt(data.gl_ar_balance)+')','','','','','','','','','','',data.recon_diff]);
    if(Number(data.opening_residual||0)>=0.005)d.push(['Of which not itemized on the subledger','','','','','','','','','','',data.opening_residual]);
    exportToExcel(d,'AR_Aging_'+data.as_of+'.xlsx',{plainCols:[4],formulas:F});
  };
  const hasGL=data&&data.gl_rows&&data.gl_rows.length>0;
  const hasAnything=data&&(data.rows.length>0||hasGL);
  const ncols=1+BK.length+2;
  return(<div>
    {showUpload&&<ArOpeningUploadModal entityId={entityId} asOf={asOf} onClose={()=>setShowUpload(false)} onImported={()=>{setShowUpload(false);run();}}/>}
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
      <div><div style={S.h1}>A/R Aging</div><div style={S.sub}>{entityName} — invoices and imported opening items are aged; any remaining un-itemized A/R on {arLabel} shows in the GL column.</div></div>
      <div style={{display:'flex',gap:8,alignItems:'center'}}><button style={S.btnS} onClick={()=>setShowUpload(true)}>Upload aging detail</button>{hasAnything&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>
    <div style={S.card}><div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div style={{flex:'0 0 180px'}}><label style={S.label}>As of date</label><input style={{...S.inputSm,width:'100%'}} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
      <button style={S.btnP} onClick={run} disabled={loading}>{loading?'Building…':'Run Aging'}</button>
      <label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}><input type="checkbox" style={S.checkbox} checked={showDetail} onChange={e=>setShowDetail(e.target.checked)}/>Show invoice detail</label>
    </div>{err&&<div style={S.err}>{err}</div>}</div>
    {data&&<>
      <div className="cl-scroll" style={scrollBox()}>
        {!hasAnything?<div style={{padding:24,color:T.textMuted}}>No open A/R as of {data.as_of}.</div>:
        <table style={S.table}><thead><tr><th style={S.th}>Customer</th>{BK.map(b=><th key={b[0]} style={S.thR}>{b[1]}</th>)}<th style={{...S.thR,color:T.accent}}>GL</th><th style={S.thR}>Total</th></tr></thead>
          <tbody>
            {data.rows.map(r=><Fragment key={r.customer}>
              <tr><td style={{...S.td,color:T.textBright,fontWeight:600}}>{r.customer}</td>
                {BK.map(b=><td key={b[0]} style={{...S.td,textAlign:'right'}}>{r[b[0]]?fmt(r[b[0]]):''}</td>)}
                <td style={{...S.td,textAlign:'right'}}></td>
                <td style={{...S.td,textAlign:'right',fontWeight:600}}>{fmt(r.total)}</td></tr>
              {showDetail&&data.detail.filter(d=>d.customer===r.customer).map(d=><tr key={d.invoice_id}>
                <td style={{...S.td,paddingLeft:26,color:T.textMuted,fontSize:12}}>{d.invoice_num} · {d.invoice_date} · due {d.due_date||'—'}{d.days_past_due>0?' · '+d.days_past_due+'d late':''}</td>
                {BK.map(b=><td key={b[0]} style={{...S.td,textAlign:'right',fontSize:12,color:T.textMuted}}>{d.bucket===b[0]?fmt(d.open):''}</td>)}
                <td style={{...S.td,textAlign:'right'}}></td>
                <td style={{...S.td,textAlign:'right',fontSize:12,color:T.textMuted}}>{fmt(d.open)}</td></tr>)}
            </Fragment>)}
            {hasGL&&<Fragment>
              <tr><td colSpan={ncols} style={{...S.td,fontWeight:700,color:T.accent,background:T.accentDim}}>GL ENTRIES <span style={{fontWeight:400,color:T.textMuted}}>— imported / manual &middot; not aged &middot; clear by journal entry</span></td></tr>
              {data.gl_rows.map((r,i)=><tr key={'gl'+i} onClick={()=>openEntry(r.entry_id)} style={{cursor:r.entry_id?'pointer':'default',opacity:entryLoading?0.6:1}}>
                <td style={{...S.td,fontSize:12}}><span style={{color:T.accent}}>{jeNum(r.entry_num)}</span> <span style={{color:T.textMuted}}>· {r.date} · {r.memo}</span></td>
                {BK.map(b=><td key={b[0]} style={{...S.td,textAlign:'right'}}></td>)}
                <td style={{...S.td,textAlign:'right',fontWeight:600,color:T.accent}}>{fmt(r.amount)}</td>
                <td style={{...S.td,textAlign:'right',fontWeight:600}}>{fmt(r.amount)}</td></tr>)}
              <tr style={{background:T.accentDim}}><td style={{...S.td,fontWeight:600,fontStyle:'italic'}}>Total GL Entries</td>
                {BK.map(b=><td key={b[0]} style={{...S.td,textAlign:'right'}}></td>)}
                <td style={{...S.td,textAlign:'right',fontWeight:700,color:T.accent}}>{fmt(data.gl_total)}</td>
                <td style={{...S.td,textAlign:'right',fontWeight:700,color:T.textBright}}>{fmt(data.gl_total)}</td></tr>
            </Fragment>}
            <tr style={{borderTop:'2px solid '+T.border}}><td style={{...S.td,fontWeight:700,color:T.textBright}}>Total</td>
              {BK.map(b=><td key={b[0]} style={{...S.td,textAlign:'right',fontWeight:700}}>{fmt(data.totals[b[0]])}</td>)}
              <td style={{...S.td,textAlign:'right',fontWeight:700,color:T.accent}}>{fmt(data.totals.gl||0)}</td>
              <td style={{...S.td,textAlign:'right',fontWeight:700,color:T.textBright}}>{fmt(data.totals.total)}</td></tr>
            {(()=>{const ties=Math.abs(data.recon_diff)<0.005;const resid=Number(data.opening_residual||0);const partial=ties&&resid>=0.005;
              const bg=!ties?'#fdf2f4':(partial?'#fff8ed':'#f3faf5');const fg=!ties?T.red:(partial?T.orange:T.green);
              const txt=!ties?fmt(data.recon_diff)+' — does not tie'
                :(partial?fmt(0)+' ✓ — but '+fmt(resid)+' of the control balance is not itemized on the subledger'
                  :fmt(0)+' ✓');
              return <tr><td colSpan={ncols} style={{...S.td,textAlign:'right',background:bg,color:fg,fontWeight:600}}>
                Reconciliation vs GL {arLabel} ({fmt(data.gl_ar_balance)}): {txt}</td></tr>;})()}
          </tbody></table>}
      </div>
    </>}
    {viewEntry&&<EditJEModal entityId={entityId} entry={viewEntry} accounts={[]} onClose={()=>setViewEntry(null)} onSaved={()=>{setViewEntry(null);run();}}/>}
  </div>);
}

// ═══ Chart of Accounts ═══
function ChartOfAccounts({entityId,entityName,canEdit}){const[accounts,setAccounts]=useState([]);const[showAdd,setShowAdd]=useState(false);const[q,setQ]=useState('');
  const[form,setForm]=useState({code:'',name:'',type:'Asset',subtype:'',bank_acct:false});const[err,setErr]=useState('');
  const[editing,setEditing]=useState(null);const[editForm,setEditForm]=useState({});const[editErr,setEditErr]=useState('');
  const[balByCode,setBalByCode]=useState({});const[drillAcct,setDrillAcct]=useState(null);
  const asOf=today();
  const yearAgo=(()=>{const d=new Date();d.setFullYear(d.getFullYear()-1);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');})();
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
  const[projects,setProjects]=useState([]);const[projFilter,setProjFilter]=useState(''); // dim_projects list + selected project id
  const[begBals,setBegBals]=useState({}); // account_code -> opening balance as of day before `from`
  // Local-parts date (no UTC shift): the day before `from`, for the opening balance.
  const _prevDay=(d)=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0')+'-'+String(x.getDate()).padStart(2,'0');};
  const reload=useCallback(()=>{
    Promise.all([api.getEntries(entityId,from||undefined,to||undefined),api.getAccounts(entityId)]).then(([e,a])=>{setEntries(e);setAccounts(a);});
    // Opening balances only matter when a start date is set; without `from` the
    // ledger runs from inception so every account correctly opens at 0.
    if(/^\d{4}-\d{2}-\d{2}$/.test(from)){
      api.getBalances(entityId,{as_of:_prevDay(from),...(projFilter?{project_id:projFilter}:{})}).then(bals=>{setBegBals(Object.fromEntries((bals||[]).map(b=>[b.code,b.balance||0])));}).catch(()=>setBegBals({}));
    } else setBegBals({});
  },[entityId,from,to,projFilter]);
  useEffect(()=>{reload();},[reload]);
  useEffect(()=>{api.getProjects(entityId).then(setProjects).catch(()=>setProjects([]));},[entityId]);
  const filtered=accounts.filter(a=>!filter||a.code===filter).sort((a,b)=>a.code.localeCompare(b.code));
  const entryAtts={};entries.forEach(e=>{if(e.attachments?.length>0)entryAtts[e.id]=e.attachments;});
  const doExport=()=>{const rows=[[entityName||'General Ledger'],['General Ledger'],['Period: '+(from||'Begin')+' to '+(to||today())],[]];const F=[];filtered.forEach(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code&&(!projFilter||Number(l.project_id)===Number(projFilter)))txns.push({date:e.date,je:'JE-'+String(e.entry_num).padStart(4,'0'),memo:e.memo,debit:l.debit,credit:l.credit});});});if(txns.length===0&&!filter)return;rows.push([acctLabel(acct.code,acct.name)]);rows.push(['Date','JE','Memo','Debit','Credit','Balance']);const isDr=acct.type==='Asset'||acct.type==='Expense';let run=from?(begBals[acct.code]||0):0;let prevRow=null;if(from){rows.push(['','','Beginning Balance','','',run]);prevRow=rows.length-1;}txns.sort((a,b)=>a.date.localeCompare(b.date)).forEach(t=>{run+=isDr?(t.debit-t.credit):(t.credit-t.debit);rows.push([t.date,t.je,t.memo,t.debit||'',t.credit||'',run]);const r=rows.length-1,rr=r+1;const delta=isDr?('D'+rr+'-E'+rr):('E'+rr+'-D'+rr);F.push({r,c:5,f:prevRow!=null?('F'+(prevRow+1)+'+'+delta):delta});prevRow=r;});rows.push([]);});exportToExcel(rows,'GL.xlsx',{formulas:F});};
  return(<div><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}><div><div style={S.h1}>General Ledger</div>{entityName&&<div style={{fontSize:13,color:T.textMuted}}>{entityName}</div>}</div><button style={S.btnExport} onClick={doExport}>Export Excel</button></div><div style={S.sub}/>
    <div style={S.filterBar}><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={from} onChange={e=>setFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <div style={{maxWidth:280}}><label style={S.label}>Account</label><select style={{...S.inputSm,width:'100%'}} value={filter} onChange={e=>setFilter(e.target.value)}><option value="">All accounts</option>{accounts.sort((a,b)=>a.code.localeCompare(b.code)).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>{projects.length>0&&<div style={{maxWidth:280}}><label style={S.label}>Project</label><select style={{...S.inputSm,width:'100%'}} value={projFilter} onChange={e=>setProjFilter(e.target.value)}><option value="">All projects</option>{projects.map(p=><option key={p.id} value={p.id}>{p.code?p.code+' — '+p.name:p.name}</option>)}</select></div>}</div>
    {filtered.map(acct=>{const txns=[];entries.forEach(e=>{e.lines.forEach(l=>{if(l.account_code===acct.code&&(!projFilter||Number(l.project_id)===Number(projFilter)))txns.push({...l,date:e.date,memo:e.memo,jeNum:e.entry_num,jeId:e.id});});});
      if(txns.length===0&&!filter)return null;txns.sort((a,b)=>a.date.localeCompare(b.date));const dr=acct.type==='Asset'||acct.type==='Expense';const beg=from?(begBals[acct.code]||0):0;let run=beg;
      return(<div key={acct.code} style={S.cardFlush}><div style={{padding:'14px 20px',display:'flex',alignItems:'center',gap:10,borderBottom:'1px solid '+T.border}}>
        <span style={{fontWeight:700,color:T.textBright,fontSize:14}}>{acct.code}</span><span>{acct.name}</span><span style={S.tag(acct.type)}>{acct.type}</span></div>
        {txns.length===0?<div style={{padding:20,color:T.textDim}}>No transactions</div>:
        <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:900}}><thead><tr><th style={S.th}>Date</th><th style={S.th}>JE</th><th style={S.th}>Memo</th><th style={S.thR}>Debit</th><th style={S.thR}>Credit</th><th style={S.thR}>Balance</th><th style={{...S.th,width:100}}>Docs</th></tr></thead>
          <tbody>{from&&<tr><td style={{...S.td,color:T.textMuted}}>{from}</td><td style={S.td}></td><td style={{...S.td,fontStyle:'italic',color:T.textMuted}}>Beginning Balance</td><td style={S.tdR}></td><td style={S.tdR}></td><td style={{...S.tdR,fontWeight:600,color:T.textMuted}}>{fmt(beg)}</td><td style={S.td}></td></tr>}{txns.map((t,i)=>{run+=dr?(t.debit-t.credit):(t.credit-t.debit);const atts=entryAtts[t.jeId];return<tr key={i}><td style={{...S.td,color:T.textMuted}}>{t.date}</td><td style={S.td}><button style={{background:'none',border:0,padding:0,color:T.accent,fontWeight:600,cursor:'pointer',fontSize:'inherit',fontFamily:'inherit'}} onClick={()=>{const e=entries.find(x=>x.id===t.jeId);if(e)setEditEntry(e);}}>JE-{String(t.jeNum).padStart(4,'0')}</button></td><td style={S.td} title={t.description||t.memo}>{t.description||t.memo}</td><td style={S.tdR}>{t.debit>0?fmt(t.debit):''}</td><td style={S.tdR}>{t.credit>0?fmt(t.credit):''}</td><td style={{...S.tdR,fontWeight:600,color:T.textBright}}>{fmt(run)}</td>
            <td style={S.td}>{atts?atts.map(a=><a key={a.id} href={api.downloadAttachment(a.id)} target="_blank" rel="noreferrer" style={S.attachLink}>{a.original_name}</a>):''}</td></tr>;})}</tbody></table></div>}</div>);})}
    {editEntry&&<EditJEModal entityId={entityId} dimsEnabled={dimsEnabled} isTurnkeyEntity={/turnkey\s*rail/i.test(entityName||'')} entry={editEntry} accounts={accounts} onClose={()=>setEditEntry(null)} onSaved={()=>{setEditEntry(null);reload();}}/>}
    </div>);}

// ═══ Wire Coding Notes Modal ═══
// Leave a note during the month describing how a wire should be coded. On the
// next statement upload, a row whose amount (within tolerance) and exact date
// match the note is auto-populated with the note's GL coding + dimension and
// arrives 'coded' for review. The note is kept for reference after it fires.
function WireNotesModal({entityId,selAcct,bankAccts,accounts,setAccounts,setBankAccts,locations=[],classes=[],dimProjects=[],dimsEnabled=true,canEdit=true,onClose}){
  const[notes,setNotes]=useState([]);const[err,setErr]=useState('');const[msg,setMsg]=useState('');const[showAddAcct,setShowAddAcct]=useState(false);
  const blank={bank_account_code:selAcct||'',note:'',match_amount:'',amount_tolerance:'0',match_date:'',desc_keyword:'',account_code:'',memo:'',dim:'',one_shot:true};
  const[form,setForm]=useState(blank);const[editId,setEditId]=useState(null);
  // Refs to the native date/amount inputs. Some browsers (and autofill/date
  // pickers) can set the field's displayed value without firing React's
  // onChange, leaving form.match_date empty at submit. On save we read the live
  // input value as the source of truth so a visible date is never rejected.
  const dateRef=useRef(null);const amountRef=useRef(null);
  // Files staged in the form. For a new note they're uploaded right after the
  // note is created; when editing an existing note they upload immediately.
  const[stagedFiles,setStagedFiles]=useState([]);const[uploadingFiles,setUploadingFiles]=useState(false);
  const load=useCallback(()=>{api.getBankCodingNotes(entityId).then(setNotes).catch(e=>setErr(e.message));},[entityId]);
  useEffect(()=>{load();},[load]);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const reset=()=>{setForm(blank);setEditId(null);setStagedFiles([]);};
  // One tagged dimension per note (Project / Location / Class), mirroring the coding grid.
  const projOpts=dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)}));
  const locOpts=locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name}));
  const clsOpts=classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name}));
  const dimOpts=[...projOpts,...locOpts,...clsOpts];const showDims=dimsEnabled&&dimOpts.length>0;
  const dimFromNote=n=>n.project_id?'project:'+n.project_id:n.location_id?'location:'+n.location_id:n.class_id?'class:'+n.class_id:'';
  const dimLabel=n=>{const v=dimFromNote(n);const o=dimOpts.find(x=>x.v===v);return o?o.label:'';};
  const startEdit=n=>{setEditId(n.id);setStagedFiles([]);setForm({bank_account_code:n.bank_account_code||'',note:n.note||'',match_amount:String(n.match_amount),amount_tolerance:String(n.amount_tolerance||0),match_date:n.date_from||n.date_to||'',desc_keyword:n.desc_keyword||'',account_code:n.account_code||'',memo:n.memo||'',dim:dimFromNote(n),one_shot:!!n.one_shot});};
  // Attach picked files. When the note is already saved (editId), upload now;
  // otherwise stage them to upload right after the note is created on save.
  const onPickFiles=async e=>{const files=Array.from(e.target.files||[]);e.target.value='';if(!files.length)return;
    if(editId){setUploadingFiles(true);setErr('');try{await api.uploadBankCodingNoteFiles(entityId,editId,files);setMsg('Support added');load();}catch(ex){setErr(ex.message);}finally{setUploadingFiles(false);}}
    else{setStagedFiles(prev=>[...prev,...files]);}};
  const removeStaged=i=>setStagedFiles(prev=>prev.filter((_,ix)=>ix!==i));
  const delAttachment=async(aid)=>{if(!confirm('Remove this supporting document?'))return;try{await api.deleteBankCodingNoteFile(aid);load();}catch(e){setErr(e.message);}};
  const save=async()=>{setErr('');setMsg('');
    const[dk,di]=form.dim?form.dim.split(':'):['',''];
    // Read live input values as source of truth (guards against a displayed
    // date/amount that never fired onChange, e.g. via autofill or the picker).
    const matchDate=(dateRef.current&&dateRef.current.value)||form.match_date||'';
    const matchAmountRaw=(amountRef.current&&amountRef.current.value)||form.match_amount||'';
    // Precise date is stored as a single-day window (date_from == date_to).
    const body={bank_account_code:form.bank_account_code||null,note:form.note||null,match_amount:Number(matchAmountRaw),amount_tolerance:Number(form.amount_tolerance)||0,date_from:matchDate||null,date_to:matchDate||null,desc_keyword:form.desc_keyword||null,account_code:form.account_code||null,memo:form.memo||null,project_id:dk==='project'?di:null,class_id:dk==='class'?Number(di):null,location_id:dk==='location'?Number(di):null,one_shot:!!form.one_shot};
    if(!isFinite(body.match_amount)||body.match_amount===0){setErr('Enter a signed wire amount (negative for money out).');return;}
    if(!body.date_from){setErr('Enter the transaction date.');return;}
    if(!body.account_code){setErr('Choose the GL account to code the wire to.');return;}
    try{
      if(editId){await api.updateBankCodingNote(entityId,editId,body);}
      else{const r=await api.createBankCodingNote(entityId,body);if(stagedFiles.length&&r&&r.id){setUploadingFiles(true);try{await api.uploadBankCodingNoteFiles(entityId,r.id,stagedFiles);}finally{setUploadingFiles(false);}}}
      setMsg(editId?'Note updated':'Note saved');reset();load();
    }catch(e){setErr(e.message);}};
  const del=async id=>{if(!confirm('Delete this wire note?'))return;try{await api.deleteBankCodingNote(entityId,id);load();}catch(e){setErr(e.message);}};
  const toggleActive=async n=>{try{await api.updateBankCodingNote(entityId,n.id,{active:!n.active});load();}catch(e){setErr(e.message);}};
  const acctName=code=>{const a=accounts.find(x=>x.code===code);return a?acctLabel(a.code,a.name):code;};
  const bankName=code=>{if(!code)return'Any account';const a=(bankAccts||[]).find(x=>x.code===code);return a?acctLabel(a.code,a.name):code;};

  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:920}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Wire Coding Notes</div>
    <div style={{fontSize:12,color:T.textMuted,marginBottom:18}}>Leave a note for a wire processed this month. When you upload the bank statement, the row matching the amount and date is auto-coded (status Coded) for your review before posting.</div>
    {err&&<div style={S.err}>{err}</div>}{msg&&<div style={S.success}>{msg}</div>}
    {canEdit&&<div style={{border:'1px solid '+T.border,borderRadius:T.radiusXs,padding:16,marginBottom:20,background:T.bgSoft||'#fafafa'}}>
      <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:12}}>{editId?'Edit note':'New note'}</div>
      <div style={S.row}>
        <div style={S.col}><label style={S.label}>Bank Account</label><select style={S.select} value={form.bank_account_code} onChange={e=>set('bank_account_code',e.target.value)}><option value="">Any account</option>{(bankAccts||[]).map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select></div>
        <div style={S.col}><label style={S.label}>Transaction Date</label><input ref={dateRef} style={S.input} type="date" value={form.match_date} onChange={e=>set('match_date',e.target.value)}/></div>
      </div>
      <div style={S.row}>
        <div style={S.col}><label style={S.label}>Wire Amount (signed)</label><input ref={amountRef} style={S.input} type="number" step="0.01" placeholder="-250000.00 (out) / 250000.00 (in)" value={form.match_amount} onChange={e=>set('match_amount',e.target.value)}/></div>
        <div style={S.col}><label style={S.label}>Amount Tolerance (±$)</label><input style={S.input} type="number" step="0.01" min="0" value={form.amount_tolerance} onChange={e=>set('amount_tolerance',e.target.value)}/></div>
        <div style={S.col}><label style={S.label}>Description contains (optional)</label><input style={S.input} placeholder="e.g. WIRE, FEDWIRE" value={form.desc_keyword} onChange={e=>set('desc_keyword',e.target.value)}/></div>
      </div>
      <div style={S.row}>
        <div style={{...S.col,flex:2}}><label style={S.label}>Code to GL Account</label>
          <div style={{display:'flex',gap:6}}>
            <select style={{...S.select,flex:1}} value={form.account_code} onChange={e=>set('account_code',e.target.value)}><option value="">Select account...</option>{accounts.map(a=><option key={a.code} value={a.code}>{acctLabel(a.code,a.name)}</option>)}</select>
            <button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40',whiteSpace:'nowrap'}} onClick={()=>setShowAddAcct(true)} title="Create a new GL account">+ New</button>
          </div>
        </div>
        {showDims&&<div style={{...S.col,flex:2}}><label style={S.label}>Dimension (optional)</label><select style={S.select} value={form.dim} onChange={e=>set('dim',e.target.value)}><option value="">None</option>{dimOpts.map(o=><option key={o.v} value={o.v}>{o.label}</option>)}</select></div>}
      </div>
      <div style={S.row}>
        <div style={{...S.col,flex:2}}><label style={S.label}>Memo (optional)</label><input style={S.input} value={form.memo} onChange={e=>set('memo',e.target.value)}/></div>
      </div>
      <div style={{marginTop:4}}>
        <label style={S.label}>Support (email copy, PDF, Excel — optional)</label>
        <div style={{display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
          <div style={{position:'relative',display:'inline-block',overflow:'hidden'}}>
            <button style={{...S.btnS,pointerEvents:'none'}} disabled={uploadingFiles}>{uploadingFiles?'Uploading...':'+ Attach files'}</button>
            <input type="file" multiple accept=".pdf,.xlsx,.xls,.csv,.eml,.msg,.png,.jpg,.jpeg,.doc,.docx,.txt" style={{position:'absolute',top:0,left:0,width:'100%',height:'100%',opacity:0,cursor:'pointer'}} onChange={onPickFiles}/>
          </div>
          {/* Already-saved attachments (edit mode) */}
          {editId&&(notes.find(n=>n.id===editId)?.attachments||[]).map(a=><span key={a.id} style={{display:'inline-flex',alignItems:'center',gap:6,fontSize:11,background:T.bgSoft||'#f1f1f1',border:'1px solid '+T.border,borderRadius:12,padding:'3px 9px'}}>
            <a href={api.bankCodingNoteFileUrl(a.id)} target="_blank" rel="noreferrer" style={{color:T.accent,textDecoration:'none'}} title={a.original_name}>{a.original_name.length>24?a.original_name.slice(0,24)+'…':a.original_name}</a>
            <span style={{cursor:'pointer',color:T.red,fontWeight:700}} onClick={()=>delAttachment(a.id)}>×</span>
          </span>)}
          {/* Staged files not yet uploaded (new note) */}
          {stagedFiles.map((f,i)=><span key={i} style={{display:'inline-flex',alignItems:'center',gap:6,fontSize:11,background:T.orange+'18',border:'1px solid '+T.orange+'55',borderRadius:12,padding:'3px 9px'}} title="Will upload when you save">
            {f.name.length>24?f.name.slice(0,24)+'…':f.name}
            <span style={{cursor:'pointer',color:T.red,fontWeight:700}} onClick={()=>removeStaged(i)}>×</span>
          </span>)}
        </div>
      </div>
      <div style={{display:'flex',alignItems:'center',gap:16,marginTop:10}}>
        <label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,color:T.text,cursor:'pointer'}}><input type="checkbox" checked={form.one_shot} onChange={e=>set('one_shot',e.target.checked)}/>Match only one wire (recommended)</label>
        <div style={{flex:1}}/>
        {editId&&<button style={S.btnS} onClick={reset}>Cancel</button>}
        <button style={S.btnP} onClick={save}>{editId?'Update note':'Save note'}</button>
      </div>
      <div style={{fontSize:11,color:T.textMuted,marginTop:10}}>The row must match the transaction date exactly. Widen the amount tolerance if the statement figure may differ slightly (e.g. wire fees). Add a description keyword only if the amount alone might collide with another transaction.</div>
    </div>}
    <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:8}}>Saved notes ({notes.length})</div>
    {notes.length===0?<div style={{fontSize:12,color:T.textMuted,padding:'12px 0'}}>No notes yet.</div>:
    <table style={S.table}><thead><tr>
      <th style={S.th}>Bank Acct</th><th style={S.th}>Date</th><th style={S.thR}>Amount (±tol)</th><th style={S.th}>Keyword</th><th style={S.th}>Codes To</th><th style={S.th}>Dimension</th><th style={S.th}>Note</th><th style={S.th}>Support</th><th style={S.th}>Status</th><th style={S.th}></th>
    </tr></thead><tbody>{notes.map(n=><tr key={n.id} style={n.active?{}:{opacity:0.5}}>
      <td style={{...S.td,fontSize:11}}>{bankName(n.bank_account_code)}</td>
      <td style={{...S.td,fontSize:11,color:T.textMuted}}>{n.date_from||n.date_to||'—'}</td>
      <td style={{...S.tdR,fontWeight:700,color:n.match_amount>=0?T.green:T.red}}>{n.match_amount>=0?'+':''}{fmt(n.match_amount)}{Number(n.amount_tolerance)>0?<span style={{color:T.textMuted,fontWeight:400,fontSize:11}}> ±{fmt(n.amount_tolerance)}</span>:null}</td>
      <td style={{...S.td,fontSize:11}}>{n.desc_keyword||'—'}</td>
      <td style={{...S.td,fontSize:11}} title={n.account_code?acctName(n.account_code):''}>{n.account_code?acctName(n.account_code):'—'}</td>
      <td style={{...S.td,fontSize:11,color:T.textMuted}} title={dimLabel(n)}>{dimLabel(n)?(dimLabel(n).length>22?dimLabel(n).slice(0,22)+'…':dimLabel(n)):'—'}</td>
      <td style={{...S.td,fontSize:11,color:T.textMuted}} title={n.note||''}>{n.note?(n.note.length>24?n.note.slice(0,24)+'…':n.note):'—'}</td>
      <td style={{...S.td,fontSize:11}}>{(n.attachments&&n.attachments.length)?<span style={{display:'inline-flex',gap:4,flexWrap:'wrap'}}>{n.attachments.map(a=><a key={a.id} href={api.bankCodingNoteFileUrl(a.id)} target="_blank" rel="noreferrer" style={{color:T.accent,textDecoration:'none'}} title={a.original_name}>📎</a>)}<span style={{color:T.textMuted}}>{n.attachments.length}</span></span>:'—'}</td>
      <td style={{...S.td,fontSize:11}}>{n.matched_count>0?<span style={{color:T.teal,fontWeight:600}} title={'Matched '+n.matched_count+'× · last '+(n.last_matched_at||'')}>Matched {n.matched_count}×</span>:(n.active?<span style={{color:T.orange}}>Waiting</span>:<span style={{color:T.textMuted}}>Inactive</span>)}{n.one_shot?'':<span style={{color:T.textMuted,fontSize:10}} title="Recurring"> ↻</span>}</td>
      <td style={{...S.td}}>{canEdit&&<div style={{display:'flex',gap:4}}>
        <button style={{...S.btnGhost,fontSize:11,padding:'4px 6px'}} onClick={()=>startEdit(n)}>Edit</button>
        <button style={{...S.btnGhost,fontSize:11,padding:'4px 6px',color:T.textMuted}} onClick={()=>toggleActive(n)}>{n.active?'Disable':'Enable'}</button>
        <button style={{...S.btnGhost,fontSize:11,padding:'4px 6px',color:T.red}} onClick={()=>del(n.id)}>Delete</button>
      </div>}</td>
    </tr>)}</tbody></table>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>{if(setAccounts)setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));if(a.bank_acct&&setBankAccts)setBankAccts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));set('account_code',a.code);}}/>}
  </div></div>);
}

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

function SplitBankTransactionModal({txn, accounts, excludeCode, entityId, dimsEnabled=true, onClose, onSaved}){
  const target = Math.abs(txn.amount);
  const initialLines = (txn.splits && txn.splits.length > 0)
    ? txn.splits.map(s => ({ account_code: s.account_code, amount: String(s.amount), memo: s.memo || '', project_id: s.project_id||null, class_id: s.class_id||null, location_id: s.location_id||null, invoice_id: s.invoice_id||null }))
    : (txn.account_code
        ? [{ account_code: txn.account_code, amount: target.toFixed(2), memo: txn.memo || '', project_id: txn.project_id||null, class_id: txn.class_id||null, location_id: txn.location_id||null }, { account_code: '', amount: '', memo: '' }]
        : [{ account_code: '', amount: '', memo: '' }, { account_code: '', amount: '', memo: '' }]);
  const [lines, setLines] = useState(initialLines);
  const [err, setErr] = useState('');
  const [saving, setSaving] = useState(false);
  // A/R cash application: a deposit line coded to the A/R control account can be
  // applied to a specific open invoice, tagging the split with invoice_id so the
  // aging report clears that document on post (no extra JE).
  const arCodes = (accounts||[]).filter(a => /accounts?\s*receivable/i.test(a.name||'') && !/other|note|interest/i.test(a.name||'')).map(a => String(a.code));
  const isDeposit = txn.amount > 0;
  const [openInvoices, setOpenInvoices] = useState([]);
  // Exclude THIS transaction's own splits from the pending-allocation math, so
  // re-opening the modal on a deposit that already applied to an invoice doesn't
  // count that application against itself (which would show 0 remaining on its
  // own invoice). effective_open then = invoice open − OTHER unposted deposits.
  useEffect(() => { if (isDeposit && arCodes.length) api.getArOpenInvoices(entityId, null, txn.id).then(d => setOpenInvoices((d && d.invoices) || [])).catch(() => {}); }, [entityId]);
  const isArLine = l => arCodes.includes(String(l.account_code));
  // Effective remaining: prefer the server's effective_open (invoice open minus
  // amounts already spoken for by other coded-but-unposted deposits); fall back
  // to plain open for older server responses.
  const effOpen = o => (o.effective_open != null ? o.effective_open : o.open);
  const pickInvoice = (i, invId) => setLines(prev => prev.map((l, idx) => {
    if (idx !== i) return l;
    if (!invId) return { ...l, invoice_id: null };
    const inv = openInvoices.find(o => String(o.id) === String(invId));
    return inv ? { ...l, invoice_id: inv.id, amount: String(effOpen(inv) > 0.005 ? effOpen(inv) : inv.open), memo: l.memo || inv.invoice_num } : l;
  }));
  const [locations, setLocations] = useState([]); const [classes, setClasses] = useState([]); const [dimProjects, setDimProjects] = useState([]);
  useEffect(() => { api.getLocations(entityId).then(d=>setLocations(d||[])).catch(()=>{}); api.getClasses(entityId).then(d=>setClasses(d||[])).catch(()=>{}); api.getProjects(entityId).then(d=>setDimProjects(d||[])).catch(()=>{}); }, [entityId]);
  const dimOpts = [
    ...dimProjects.map(pr=>({v:'project:'+pr.id,label:'Project — '+(pr.code&&pr.code!==pr.name?pr.code+' — '+pr.name:pr.name)})),
    ...locations.map(loc=>({v:'location:'+loc.id,label:'Location — '+(loc.code?loc.code+' — ':'')+loc.name})),
    ...classes.map(c=>({v:'class:'+c.id,label:classTerm()+' — '+(c.code?c.code+' — ':'')+c.name})),
  ];
  const showDims = dimsEnabled && dimOpts.length > 0;
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
    // Keep any line with an account and a non-zero amount. A negative amount is a
    // credit-memo / offset line (e.g. a receipt applied net of a credit memo); the
    // signed total still has to net to the transaction amount.
    const valid = lines.filter(l => l.account_code && parseAmt(l.amount) !== 0);
    if (valid.length === 0) { setErr('Add at least one account with an amount'); return; }
    if (valid.reduce((s, l) => s + parseAmt(l.amount), 0) <= 0) { setErr('Splits must net to a positive amount matching the transaction'); return; }
    if (!balanced) { setErr('Splits must total ' + fmt(target) + ' (currently off by ' + fmt(remaining) + ')'); return; }
    setSaving(true);
    try {
      await api.splitBankTransaction(entityId, txn.id, valid.map(l => ({ account_code: l.account_code, amount: parseAmt(l.amount), memo: l.memo || null, project_id: l.project_id||null, class_id: l.class_id||null, location_id: l.location_id||null, invoice_id: l.invoice_id || null })));
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
        <td style={{...S.td,padding:'4px 6px'}}>{isDeposit && isArLine(l)
          ? (()=>{ const selInv=openInvoices.find(o=>String(o.id)===String(l.invoice_id));
              const spokenFor=selInv&&selInv.pending_applied>0.005;
              return <div>
                <select style={{...S.inputSm,width:'100%'}} value={l.invoice_id||''} onChange={e => pickInvoice(i, e.target.value)}>
                  <option value="">Apply to invoice&hellip; (on account)</option>
                  {openInvoices.slice()
                    .sort((a,b)=>(Math.abs(effOpen(b)-target)<0.005)-(Math.abs(effOpen(a)-target)<0.005))
                    .map(o => { const rem=effOpen(o); const full=rem<0.005;
                      // Fully spoken-for invoices are disabled — unless it's the one already
                      // selected on THIS line (so an existing choice always stays visible/selectable).
                      const isSelected=String(o.id)===String(l.invoice_id);
                      return <option key={o.id} value={o.id} disabled={full && !isSelected}>
                        {Math.abs(rem-target)<0.005?'\u2713 ':''}{o.invoice_num} {'\u2014'} {o.customer} {'\u2014'} ${fmt(rem)}{o.pending_applied>0.005?' left (of $'+fmt(o.open)+')':''} ({o.bucket==='current'?'current':o.days_past_due+'d'}){full?' \u2014 fully applied':''}
                      </option>; })}
                </select>
                {spokenFor && <div style={{fontSize:10,color:T.orange,marginTop:2,lineHeight:1.3}}>
                  ${fmt(selInv.pending_applied)} already applied by unposted {selInv.pending_txns&&selInv.pending_txns.length?('txn '+selInv.pending_txns.map(t=>'#'+t.txn_id).join(', ')):'deposit(s)'} &mdash; ${fmt(effOpen(selInv))} remaining
                </div>}
              </div>; })()
          : <input style={S.inputSm} value={l.memo} onChange={e => updateLine(i, 'memo', e.target.value)} placeholder="Memo"/>}</td>
        <td style={{...S.td,padding:'4px 6px',textAlign:'center'}}>{lines.length > 1 && <button style={S.btnGhost} onClick={() => removeLine(i)}>x</button>}</td>
      </tr>)}</tbody>
    </table>
    <div style={{fontSize:11,color:T.textMuted,marginBottom:8}}>Enter a negative amount for a credit memo or offset line; splits net to the transaction amount.</div>
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

function BankTransactions({entityId,canEdit=true,dimsEnabled=true,bankSelAcct:selAcct,setBankSelAcct:setSelAcct,bankTxns:txns,setBankTxns:setTxns,bankUploading:uploading,setBankUploading:setUploading,bankStatusFilter:statusFilter,setBankStatusFilter:setStatusFilter}){
  const[accounts,setAccounts]=useState([]);const[bankAccts,setBankAccts]=useState([]);
  const[err,setErr]=useState('');const[msg,setMsg]=useState('');const[showAddAcct,setShowAddAcct]=useState(false);
  const[uploadProgress,setUploadProgress]=useState('');const[discarding,setDiscarding]=useState(false);
  const[splitTxn,setSplitTxn]=useState(null);
  const[matchTxn,setMatchTxn]=useState(null);
  const[showNotes,setShowNotes]=useState(false);
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
      setMsg(r.count+' imported'+(r.auto_coded>0?', '+r.auto_coded+' auto-coded from wire notes':'')+(auto>0?', '+auto+' auto-categorized':''));loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}finally{setUploading(false);setUploadProgress('');}};
  const cancelUpload=()=>{setUploading(false);setUploadProgress('');setMsg('Upload cancelled');};
  // Only genuinely un-worked rows are discardable. "Matched" (linked to an existing
  // JE) and "posted" (has its own JE) are finished states — discarding a matched row
  // would break its reconciliation link to an already-booked entry, and a posted row
  // has a live GL entry. "Coded" rows have a chosen account but aren't booked yet, so
  // they're still safe to discard. Discardable = pending OR coded, never matched/posted.
  const discardable=t=>t.status==='pending'||t.status==='coded';
  const discardAllUnposted=async()=>{const drop=txns.filter(discardable);if(!drop.length){setErr('Nothing to discard');return;}
    const batchIds=[...new Set(drop.map(t=>t.batch_id).filter(Boolean))];
    if(!confirm('Discard all '+drop.length+' un-posted transaction(s) from this account? This cannot be undone. Matched and posted transactions will be kept.'))return;
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
  const dimOpts=[...projOpts,...locOpts,...clsOpts];const showDims=dimsEnabled&&dimOpts.length>0;
  const txnDimValue=t=>t.project_id?'project:'+t.project_id:t.location_id?'location:'+t.location_id:t.class_id?'class:'+t.class_id:'';
  const setTxnDim=(t,val)=>{const[kind,id]=val?val.split(':'):['',''];codeTransaction(t.id,t.account_code||null,t.memo,{project_id:kind==='project'?id:null,class_id:kind==='class'?id:null,location_id:kind==='location'?id:null});};
  const postCoded=async()=>{const ids=txns.filter(t=>t.status==='coded').map(t=>t.id);if(!ids.length){setErr('Nothing coded');return;}try{const r=await api.postBankTransactions(entityId,ids);setMsg(r.posted+' JEs created');loadTxns(selAcct,statusFilter);}catch(ex){setErr(ex.message);}};
  const changeAcct=v=>{setSelAcct(v);setTxns([]);if(v)loadTxns(v,statusFilter);};
  const changeStatus=v=>{setStatusFilter(v);if(selAcct)loadTxns(selAcct,v);};

  const filteredTxns=txns;
  const totalIn=filteredTxns.filter(t=>t.amount>0).reduce((s,t)=>s+t.amount,0);const totalOut=filteredTxns.filter(t=>t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const uncat=filteredTxns.filter(t=>t.status==='pending').length;
  const nPosted=filteredTxns.filter(t=>t.status==='posted').length;const nMatched=filteredTxns.filter(t=>t.status==='matched').length;const nCoded=filteredTxns.filter(t=>t.status==='coded').length;const nPending=filteredTxns.filter(t=>t.status==='pending').length;
  const statusBits=[nPosted&&nPosted+' posted',nMatched&&nMatched+' matched',nCoded&&nCoded+' coded',nPending&&nPending+' pending'].filter(Boolean);

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
        <button style={{...S.btnS,color:T.orange,borderColor:T.orange+'40'}} onClick={()=>setShowNotes(true)} title="Leave a note during the month so a wire is auto-coded when the statement is uploaded">Wire Notes</button>
      </div>}
    {showNotes&&<WireNotesModal entityId={entityId} selAcct={selAcct} bankAccts={bankAccts} accounts={accounts} setAccounts={setAccounts} setBankAccts={setBankAccts} locations={locations} classes={classes} dimProjects={dimProjects} dimsEnabled={dimsEnabled} canEdit={canEdit} onClose={()=>setShowNotes(false)}/>}
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
        <div><div style={S.h2}>{filteredTxns.length} Transactions</div>{statusBits.length>0&&<div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{statusBits.join(' \u00b7 ')}</div>}</div>
        <div style={{display:'flex',gap:10}}>
          {canEdit&&<button style={{...S.btnS,color:T.teal,borderColor:T.teal+'40'}} onClick={()=>setShowAddAcct(true)}>+ New Account</button>}
          {canEdit&&filteredTxns.some(discardable)&&<button style={{...S.btnD,padding:'8px 14px',fontSize:12}} disabled={discarding} onClick={discardAllUnposted}>{discarding?'Discarding...':'Discard '+filteredTxns.filter(discardable).length+' un-posted'}</button>}
          {canEdit&&filteredTxns.some(t=>t.status==='coded')&&<button style={S.btnP} onClick={postCoded}>Post {filteredTxns.filter(t=>t.status==='coded').length} to GL</button>}</div></div>
      <div style={{overflowX:'auto',width:'100%'}}>
      <table style={{...S.table,tableLayout:'fixed',width:colW.date+colW.desc+colW.amount+colW.gl+colW.memo+colW.status+36,minWidth:'100%'}}>
        <colgroup><col style={{width:colW.date}}/><col style={{width:colW.desc}}/><col style={{width:colW.amount}}/><col style={{width:colW.gl}}/><col style={{width:colW.memo}}/><col style={{width:colW.status}}/><col style={{width:36}}/></colgroup>
        <thead><tr>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Date{resizeHandle('date')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Description{resizeHandle('desc')}</th>
          <th style={{...S.thR,position:'relative',borderRight:'1px solid '+T.borderLight}}>Amount{resizeHandle('amount')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>GL Account{resizeHandle('gl')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Memo{resizeHandle('memo')}</th>
          <th style={{...S.th,position:'relative',borderRight:'1px solid '+T.borderLight}}>Status{resizeHandle('status')}</th>
          <th style={{...S.th,width:36}}></th></tr></thead>
        <tbody>{filteredTxns.map(t=><tr key={t.id} style={(t.status==='posted'||t.status==='matched')?{opacity:0.5}:{}}>
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
        </tr>)}</tbody></table></div></div>}
    {selAcct&&filteredTxns.length===0&&!uploading&&<div style={{...S.card,textAlign:'center',padding:60,color:T.textDim}}>No transactions yet. Upload a bank statement above.</div>}
    {showAddAcct&&<QuickAddAccountModal entityId={entityId} onClose={()=>setShowAddAcct(false)} onCreated={a=>{setAccounts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));if(a.bank_acct)setBankAccts(p=>[...p,a].sort((x,y)=>x.code.localeCompare(y.code)));}}/>}
    {splitTxn&&<SplitBankTransactionModal txn={splitTxn} accounts={accounts} excludeCode={selAcct} entityId={entityId} dimsEnabled={dimsEnabled} onClose={()=>setSplitTxn(null)} onSaved={()=>{setSplitTxn(null);loadTxns(selAcct,statusFilter);}}/>}
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
    const first=d.length;
    rows.forEach(r=>d.push([r.project_code||r.turnkey_project_id,r.project_name||'',r.contract_amount,r.revised_contract,r.costs_to_date,r.estimated_cost_to_complete,r.estimated_total_cost,r.estimated_gross_profit,(r.percent_complete||0)/100,r.earned_revenue,r.billed_to_date,r.over_under_billing]));
    const last=d.length-1;const F=[];
    if(tot){d.push([]);const tr=d.length;d.push(['','Total',tot.contract_amount,tot.revised_contract,tot.costs_to_date,tot.estimated_cost_to_complete,tot.estimated_total_cost,tot.estimated_gross_profit,'',tot.earned_revenue,tot.billed_to_date,tot.over_under_billing]);sumCols(F,tr,[2,3,4,5,6,7,9,10,11],first,last);}
    exportToExcel(d,'WIP_'+validAsOf+'.xlsx',{formulas:F});
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
// Format a Date as YYYY-MM-DD from LOCAL parts. Must NOT use toISOString(),
// which converts to UTC: _mkDate parses at local midnight, so a user in a
// positive-UTC-offset zone (e.g. UTC+8 Philippines) would have local midnight
// fall on the PREVIOUS calendar day in UTC, shifting the report's as-of/period
// boundaries back a day and producing a Trial Balance that drops that day's
// activity vs a US user. Local getters round-trip _mkDate correctly.
const _ymd=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
const _mkDate=s=>new Date((/^\d{4}-\d{2}-\d{2}$/.test(s)?s:today())+'T00:00:00');
// Overall [from,to] window for a date preset, anchored at `anchor` (YYYY-MM-DD).
// A preset returns the most recent COMPLETE calendar period whose end falls ON OR
// BEFORE the period-end (whole-period, calendar-aligned). For a 06/30/2026 period
// end (which is itself a month/quarter end): Last Month = June 2026, Last Quarter =
// Q2 2026 (Apr–Jun), Last Year = 2025. A mid-period period-end steps back to the
// last completed period (e.g. 05/15 → Last Quarter = Q1, Last Month = April).
//   NOTE (bug fix, Liting #1): the old version returned a trailing window
//   (anchor − N months + 1 day), e.g. Last Quarter → 2026-03-31…06-30. Split into
//   quarterly columns that produced a bogus 1-day "Q1 '26" sliver column (only 3/31
//   activity) plus a "Q2 '26" column, so the displayed quarter never matched. Whole-
//   period alignment makes each column a full calendar period. With Compare (prior
//   period) on, Last Quarter now shows Q2 vs its prior quarter Q1.
// JS Date normalizes out-of-range month indexes (month -1 → prior Dec), so periods
// that cross a year boundary are handled for free.
function rptWindow(filter,anchor){
  const a=_mkDate(anchor);const to=_ymd(a);const ay=_ymd(a);
  const Y=a.getFullYear(),M=a.getMonth();
  if(filter==='month'){const b=(ay>=_ymd(new Date(Y,M+1,0)))?M:M-1;return{from:_ymd(new Date(Y,b,1)),to:_ymd(new Date(Y,b+1,0))};}
  if(filter==='quarter'){const q=Math.floor(M/3);const b=(ay>=_ymd(new Date(Y,q*3+3,0)))?q:q-1;return{from:_ymd(new Date(Y,b*3,1)),to:_ymd(new Date(Y,b*3+3,0))};}
  if(filter==='year'){const b=(ay>=_ymd(new Date(Y,11,31)))?Y:Y-1;return{from:_ymd(new Date(b,0,1)),to:_ymd(new Date(b,11,31))};}
  return{from:null,to};// 'all' = inception → anchor
}
// Human label for a preset's actual window, e.g. "Last Quarter · 2026-01-01 → 2026-03-31".
// Presets now resolve to a prior whole period whose end is NOT the anchor, so the header
// must show the real range rather than the old (misleading) "… ending <anchor>".
function rptRangeLabel(filter,anchor){
  if(filter==='all')return 'Through '+anchor;
  const w=rptWindow(filter,anchor);const nm=RPT_DATE_FILTERS.find(f=>f[0]===filter);
  return (nm?nm[1]:filter)+' · '+w.from+' → '+w.to;
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
  // Whole-month windows (month/quarter/year) → the previous equivalent CALENDAR
  // period, so a Q2 (Apr–Jun) window compares against Q1 (Jan–Mar), not a ragged
  // day-count window. Otherwise fall back to the equal-length window right before.
  const monthStart=from.getDate()===1;
  const monthEnd=(()=>{const n=new Date(to);n.setDate(n.getDate()+1);return n.getDate()===1;})();
  if(monthStart&&monthEnd){
    const months=(to.getFullYear()-from.getFullYear())*12+(to.getMonth()-from.getMonth())+1;
    return{label:'Prev',from:_ymd(new Date(from.getFullYear(),from.getMonth()-months,1)),to:_ymd(new Date(from.getFullYear(),from.getMonth(),0))};
  }
  const days=Math.round((to-from)/86400000)+1;
  const pTo=new Date(from);pTo.setDate(pTo.getDate()-1);
  const pFrom=new Date(pTo);pFrom.setDate(pFrom.getDate()-days+1);
  return{label:'Prev',from:_ymd(pFrom),to:_ymd(pTo)};
}
// Concise label for a whole-period window: "Q2 2026", "Jun 2026", or "2026".
// Returns '' when the window is not exactly one calendar month/quarter/year, so
// callers can fall back to a generic heading.
function periodLabel(from,to){
  if(!from||!to)return '';
  const s=_mkDate(from),e=_mkDate(to);
  const monthStart=s.getDate()===1;
  const monthEnd=(()=>{const n=new Date(e);n.setDate(n.getDate()+1);return n.getDate()===1;})();
  if(monthStart&&monthEnd){
    const months=(e.getFullYear()-s.getFullYear())*12+(e.getMonth()-s.getMonth())+1;
    if(months===1)return s.toLocaleString('en-US',{month:'short'})+' '+s.getFullYear();
    if(months===3&&s.getMonth()%3===0)return 'Q'+(Math.floor(s.getMonth()/3)+1)+' '+s.getFullYear();
    if(months===12&&s.getMonth()===0)return String(s.getFullYear());
  }
  return '';
}
// "6/30/26" style date (M/D/YY, no leading zeros) for report headings.
const fmtMDY=iso=>{if(!iso)return '';const d=_mkDate(iso);return (d.getMonth()+1)+'/'+d.getDate()+'/'+String(d.getFullYear()).slice(2);};
// Period noun for a window: Quarter / Month / Year / Period (generic).
const periodNoun=(from,to)=>{const pl=periodLabel(from,to);if(/^Q/.test(pl))return 'Quarter';if(/^\d{4}$/.test(pl))return 'Year';if(pl)return 'Month';return 'Period';};
// Total-Only report subtitle naming each displayed column by its period-end date,
// e.g. "For the Quarters Ended 6/30/26 and 3/31/26" (most recent first). This avoids
// mislabeling — it does not call any single column "Last Quarter"; it just states the
// quarter-end dates the columns cover. Inception-to-date ("All") stays "Through <date>".
function endedHeading(dispCols,dateFilter,anchor){
  if(dateFilter==='all'||dispCols.some(c=>!c.from))return 'Through '+anchor;
  const sorted=[...dispCols].sort((a,b)=>(a.to<b.to?1:a.to>b.to?-1:0));
  const nouns=new Set(dispCols.map(c=>periodNoun(c.from,c.to)));const noun=nouns.size===1?[...nouns][0]:'Period';
  const dates=sorted.map(c=>fmtMDY(c.to));
  const list=dates.length<=1?(dates[0]||''):dates.slice(0,-1).join(', ')+' and '+dates[dates.length-1];
  return 'For the '+(dates.length>1?noun+'s':noun)+' Ended '+list;
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
  // Report format: 'balances' = classic single-date Debit/Credit snapshot (unchanged default);
  // 'activity' = Sage-style Beginning Balance / Period Debits / Period Credits / Ending Balance
  // over a From→To window. Retained Earnings closes annually (P&L rolls at Jan 1, not per period).
  const[format,setFormat]=useState('balances');
  const[fromDate,setFromDate]=useState(()=>{const y=(/^\d{4}/.test(asOf||'')?asOf:today()).slice(0,4);return y+'-01-01';});
  const[actData,setActData]=useState(null);
  // Filter rule: only County Line Rail Fund uses Location/Investor dimensions.
  // Every other entity filters by Project instead.
  const showLocInv=dimsEnabled&&isClrf;
  const showProj=dimsEnabled&&!isClrf;
  // Guard: while the user is editing the date input, asOf can briefly be '' or a partial string like '2026-'.
  // Avoid crashing the page on Invalid Date — fall back to today() until a complete YYYY-MM-DD is entered.
  const validAsOf=/^\d{4}-\d{2}-\d{2}$/.test(asOf)&&!isNaN(new Date(asOf+'T00:00:00').getTime())?asOf:today();
  const fyS=validAsOf.slice(0,4)+'-01-01';
  const validFrom=/^\d{4}-\d{2}-\d{2}$/.test(fromDate)&&!isNaN(new Date(fromDate+'T00:00:00').getTime())?fromDate:(validAsOf.slice(0,4)+'-01-01');
  const isActivity=format==='activity';
  const yStart=d=>d.slice(0,4)+'-01-01';
  const prevDayTB=d=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return _ymd(x);};
  const locName=locId?(locations.find(l=>String(l.id)===String(locId))?.name||''):'';
  const className=classId?(classes.find(c=>String(c.id)===String(classId))?.name||''):'';
  const projName=projId?(()=>{const p=projects.find(p=>String(p.id)===String(projId));return p?(p.code&&p.code!==p.name?p.code:p.name):'';})():'';
  const dimmed=!!(locId||classId||projId); // any dimension selected → activity view
  const scopeLabel=[locName,className,projName].filter(Boolean).join(' · ');
  // 12-month window ending at asOf. If asOf is 2026-04-10, window is 2025-04-11 → 2026-04-10.
  const drillFrom=useMemo(()=>{const d=new Date(validAsOf+'T00:00:00');d.setFullYear(d.getFullYear()-1);d.setDate(d.getDate()+1);return _ymd(d);},[validAsOf]);
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
  const oneYrBefore=d=>{const x=new Date(d+'T00:00:00');x.setFullYear(x.getFullYear()-1);x.setDate(x.getDate()+1);return _ymd(x);};
  const colHead=(c,i)=>prior&&i===0?'Prev':(c.label==='Total'?'':c.label);
  const fnameTag=[locName,className,projName].filter(Boolean).map(s=>s.replace(/[^A-Za-z0-9]+/g,'_')).join('_');
  // ── Activity (Sage-style) data ──
  // Beginning balances = balances as of the day before the From date, with P&L closed at that
  // year's Jan 1 (annual close). Ending balances = balances as of the To date, P&L closed at Jan 1.
  // Period debits/credits = gross debit/credit activity strictly between From and To.
  // Balances are shown signed debit-positive (liabilities/equity/revenue negative), matching Sage.
  useEffect(()=>{let ok=true;
    if(!isActivity){return;}
    const from=validFrom,to=validAsOf,pd=prevDayTB(validFrom);
    const begArgs=dimmed?{as_of:pd,...dimArgs}:{as_of:pd,close_pl_before:yStart(from)};
    const endArgs=dimmed?{as_of:to,...dimArgs}:{as_of:to,close_pl_before:yStart(to)};
    const actArgs={from,to,...(dimmed?dimArgs:{})};
    Promise.all([api.getBalances(entityId,begArgs).catch(()=>[]),api.getBalances(entityId,actArgs).catch(()=>[]),api.getBalances(entityId,endArgs).catch(()=>[])])
      .then(([beg,act,end])=>{if(ok)setActData({beg:beg||[],act:act||[],end:end||[]});});
    return()=>{ok=false;};
  },[entityId,isActivity,validFrom,validAsOf,JSON.stringify(dimArgs),dimmed,rk]);
  const signOf=(type,bal)=>(type==='Asset'||type==='Expense')?bal:-bal; // debit-positive signed balance
  const actRows=useMemo(()=>{
    if(!actData)return[];
    const m=new Map();
    const put=arr=>(arr||[]).forEach(b=>{if(!m.has(b.code))m.set(b.code,{code:b.code,name:b.name,type:b.type,beg:0,dr:0,cr:0,end:0});});
    put(actData.beg);put(actData.act);put(actData.end);
    (actData.beg||[]).forEach(b=>{const r=m.get(b.code);if(r)r.beg=signOf(b.type,b.balance||0);});
    (actData.end||[]).forEach(b=>{const r=m.get(b.code);if(r)r.end=signOf(b.type,b.balance||0);});
    (actData.act||[]).forEach(b=>{const r=m.get(b.code);if(r){r.dr=b.total_debit||0;r.cr=b.total_credit||0;}});
    return[...m.values()].filter(r=>Math.abs(r.beg)>0.005||Math.abs(r.end)>0.005||Math.abs(r.dr)>0.005||Math.abs(r.cr)>0.005).sort((a,b)=>String(a.code).localeCompare(String(b.code)));
  },[actData]);
  const actTot=useMemo(()=>actRows.reduce((t,r)=>({beg:t.beg+r.beg,dr:t.dr+r.dr,cr:t.cr+r.cr,end:t.end+r.end}),{beg:0,dr:0,cr:0,end:0}),[actRows]);
  const rnd=n=>{const v=Math.round((n||0)*100)/100;return Math.abs(v)<0.005?0:v;};
  const sfmt=n=>Math.abs(n||0)<0.005?'':fmt(n); // signed money, blank at zero, negatives in ()
  const doExportActivity=()=>{const lbl=scopeLabel?(' — '+scopeLabel):'';
    const d=[[entityName||'Trial Balance'],['Trial Balance'+lbl],['Beginning '+validFrom+' through Ending '+validAsOf],[],
      ['Code','Account','Type','Beginning Balance (on '+validFrom+')','Period Debits','Period Credits','Ending Balance (on '+validAsOf+')']];
    const first=d.length;
    actRows.forEach(r=>d.push([r.code,r.name,r.type,rnd(r.beg)||'',r.dr||'',r.cr||'',rnd(r.end)||'']));
    const last=d.length-1;
    d.push([]);const tr=d.length;d.push(['','','Totals',rnd(actTot.beg),actTot.dr,actTot.cr,rnd(actTot.end)]);
    const F=[];sumCols(F,tr,[3,4,5,6],first,last);
    exportToExcel(d,'TB_Activity'+(fnameTag?'_'+fnameTag:'')+'_'+validFrom+'_'+validAsOf+'.xlsx',{formulas:F});};
  const doExport=()=>{const lbl=scopeLabel?(' — '+scopeLabel):'';const hdr=['Code','Account','Type'];cols.forEach((c,i)=>{const h=colHead(c,i);hdr.push((h?h+' ':'')+'Debit',(h?h+' ':'')+'Credit');});const d=[[entityName||'Trial Balance'],['Trial Balance'+lbl],['As of '+anchor],[],hdr];
    const first=d.length;
    rows.forEach(r=>{const row=[r.code,r.name,r.type];cols.forEach((c,i)=>{const x=drcr(r.code,i);row.push(x.dr||'',x.cr||'');});d.push(row);});
    const last=d.length-1;
    const tot=['','','Total'];cols.forEach((c,i)=>{tot.push(totDr(i),totCr(i));});d.push([]);const tr=d.length;d.push(tot);
    const F=[],dcols=[];cols.forEach((c,i)=>{dcols.push(3+2*i,4+2*i);});sumCols(F,tr,dcols,first,last);
    exportToExcel(d,'TB'+(fnameTag?'_'+fnameTag:'')+'_'+anchor+'.xlsx',{formulas:F});};
  const amtStyle={...S.tdR,cursor:'pointer'};
  // GL detail export (optionally scoped to the selected location and/or investor).
  // Pulls flat lines with running balance from /gl-detail through the as-of date;
  // dimension-tagged lines only when a dimension is selected.
  const doExportGL=async()=>{
    try{
      const r=await api.getGLDetail(entityId,{to:validAsOf,...(locId?{location_id:locId}:{}),...(classId?{class_id:classId}:{}),...(projId?{project_id:projId}:{})});
      const lbl=scopeLabel?(' — '+scopeLabel):'';
      const d=[[entityName||'General Ledger'],['GL Detail'+lbl],['Through '+asOf],[],['Date','Entry #','Account','Account Name','Memo / Description','Project','Location',classTerm(),'Debit','Credit','Running Bal']];
      const first=d.length;
      (r.lines||[]).forEach(l=>d.push([l.date,l.entry_num,l.account_code,l.account_name,l.description||l.memo||'',l.project_code&&l.project_code!==l.project_name?l.project_code:(l.project_name||''),l.location_name,l.class_name,l.debit||'',l.credit||'',l.running_balance]));
      const last=d.length-1;
      d.push([]);d.push(['','','','','','','','','Total Dr','Total Cr','']);
      const tr=d.length;d.push(['','','','','','','','',r.total_debit,r.total_credit,'']);
      const F=[];sumCols(F,tr,[8,9],first,last);
      exportToExcel(d,'GL'+(fnameTag?'_'+fnameTag:'')+'_'+asOf+'.xlsx',{formulas:F});
    }catch(e){alert('GL export failed: '+e.message);}
  };
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
    <div style={S.filterBar}><div><label style={S.label}>Format</label><select style={S.inputSm} value={format} onChange={e=>setFormat(e.target.value)}><option value="balances">Balances (Debit / Credit)</option><option value="activity">Activity (Beginning → Ending)</option></select></div>
      {isActivity?<><div><label style={S.label}>From date</label><input style={S.inputSm} type="date" value={fromDate} onChange={e=>setFromDate(e.target.value)}/></div><div><label style={S.label}>To date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div></>:<><div><label style={S.label}>As of Date</label><input style={S.inputSm} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div><ReportControls dateFilter={dateFilter} setDateFilter={setDateFilter} colMode={colMode} setColMode={setColMode} compare={compare} setCompare={setCompare}/></>}
      {showProj&&<div><label style={S.label}>Project</label><select style={S.inputSm} value={projId} onChange={e=>setProjId(e.target.value)}><option value="">All (whole entity)</option>{projects.map(p=><option key={p.id} value={p.id}>{p.code&&p.code!==p.name?p.code+' — '+p.name:p.name}{p.line_count!=null?(' ('+p.line_count+')'):''}</option>)}</select></div>}
      {showLocInv&&<div><label style={S.label}>Location</label><select style={S.inputSm} value={locId} onChange={e=>setLocId(e.target.value)}><option value="">All (whole entity)</option>{locations.map(l=><option key={l.id} value={l.id}>{l.name}{l.line_count!=null?(' ('+l.line_count+')'):''}</option>)}</select></div>}
      {showLocInv&&<div><label style={S.label}>Investor (Class)</label><select style={S.inputSm} value={classId} onChange={e=>setClassId(e.target.value)}><option value="">All investors</option>{classes.map(c=><option key={c.id} value={c.id}>{c.name}{c.line_count!=null?(' ('+c.line_count+')'):''}</option>)}</select></div>}</div>
    <div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='trial' currentConfig={{asOf,dateFilter,colMode,compare,format,fromDate}} onApply={(c)=>{if(c.asOf)setAsOf(c.asOf);if(c.dateFilter)setDateFilter(c.dateFilter);if(c.colMode)setColMode(c.colMode);if(typeof c.compare==='boolean')setCompare(c.compare);if(c.format)setFormat(c.format);if(c.fromDate)setFromDate(c.fromDate);}} canEdit={canEdit}/><button style={S.btnExport} onClick={doExportGL} title="Export flat GL detail (dimension-tagged only when a location/investor is selected)">Export GL Detail</button><button style={S.btnExport} onClick={isActivity?doExportActivity:doExport}>Export TB</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Trial Balance{scopeLabel?(' — '+scopeLabel):''}</div><div style={{fontSize:13,color:T.textMuted}}>{isActivity?('Beginning '+validFrom+' → Ending '+asOf):('As of '+asOf)}{dimmed?' · dimension-tagged activity only':''}</div></div>
    {isActivity&&<div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:640}}>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th>
        <th style={S.thR}>Beginning Balance<div style={{fontSize:9,fontWeight:400,textTransform:'none',letterSpacing:0}}>on {validFrom}</div></th>
        <th style={S.thR}>Period Debits</th><th style={S.thR}>Period Credits</th>
        <th style={S.thR}>Ending Balance<div style={{fontSize:9,fontWeight:400,textTransform:'none',letterSpacing:0}}>on {validAsOf}</div></th></tr></thead>
      <tbody>{actRows.map(r=><tr key={r.code}><td style={{...S.td,color:T.textBright}}>{r.code}</td><td style={S.td} title={r.name}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td>
        <td style={{...S.tdR,color:r.beg<0?T.red:undefined}}>{sfmt(r.beg)}</td><td style={S.tdR}>{sfmt(r.dr)}</td><td style={S.tdR}>{sfmt(r.cr)}</td><td style={{...S.tdR,color:r.end<0?T.red:undefined}}>{sfmt(r.end)}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Totals</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(rnd(actTot.beg))}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(actTot.dr)}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(actTot.cr)}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(rnd(actTot.end))}</td></tr></tbody></table></div>}
    {!isActivity&&<div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:520}}>
      <thead><tr><th style={S.th}>Code</th><th style={S.th}>Account</th><th style={S.th}>Type</th>
        {cols.map((c,i)=>{const h=colHead(c,i);return[<th key={'hd'+i} style={S.thR}>{(h?h+' ':'')}Debit</th>,<th key={'hc'+i} style={S.thR}>{(h?h+' ':'')}Credit</th>];})}
        {prior&&<><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}</tr></thead>
      <tbody>{rows.map(r=><tr key={r.code}><td style={{...S.td,color:T.textBright}}>{r.code}</td><td style={S.td} title={r.name}>{r.name}</td><td style={S.td}><span style={S.tag(r.type)}>{r.type}</span></td>
        {cols.map((c,i)=>{const x=drcr(r.code,i);const clk=()=>setDrillAcct({...r,from:oneYrBefore(c.to),to:c.to});return[<td key={'d'+i} style={{...S.tdR,cursor:x.dr>0?'pointer':'default',color:x.dr>0?T.accent:undefined}} onClick={()=>x.dr>0&&clk()}>{x.dr>0?fmt(x.dr):''}</td>,<td key={'c'+i} style={{...S.tdR,cursor:x.cr>0?'pointer':'default',color:x.cr>0?T.accent:undefined}} onClick={()=>x.cr>0&&clk()}>{x.cr>0?fmt(x.cr):''}</td>];})}
        {prior&&(()=>{const cN=balAt(r.code,curI),pN=balAt(r.code,priI);return[<td key="dc" style={S.tdR}>{fmt(cN-pN)}</td>,<td key="pc" style={{...S.tdR,color:(cN-pN)>=0?T.green:T.red}}>{pctTxt(rptPct(cN,pN))}</td>];})()}</tr>)}
        <tr style={S.grandTotalRow}><td style={S.tdBold} colSpan={3}>Total</td>{cols.map((c,i)=>[<td key={'d'+i} style={{...S.tdBold,textAlign:'right'}}>${fmt(totDr(i))}</td>,<td key={'c'+i} style={{...S.tdBold,textAlign:'right'}}>${fmt(totCr(i))}</td>])}{prior&&<><td style={S.tdBold}/><td style={S.tdBold}/></>}</tr></tbody></table></div>}
    {!isActivity&&<div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(totDr(curI)-totCr(curI))<0.005?T.green:T.red}}>{Math.abs(totDr(curI)-totCr(curI))<0.005?'In balance':'Off by $'+fmt(totDr(curI)-totCr(curI))}</div>}
    {isActivity&&<div style={{textAlign:'center',marginTop:14,fontSize:13,fontWeight:600,color:Math.abs(actTot.dr-actTot.cr)<0.005?T.green:T.red}}>{Math.abs(actTot.dr-actTot.cr)<0.005?'In balance':'Off by $'+fmt(actTot.dr-actTot.cr)}</div>}</div>
    {drillAcct&&<AccountDrillDownModal entityId={entityId} entityName={entityName} acct={drillAcct} from={drillAcct.from||drillFrom} to={drillAcct.to||asOf} onClose={()=>setDrillAcct(null)} onChanged={()=>setRk(k=>k+1)}/>}
  </div>);
}

// ═══ Account Drill-Down Modal (12-month GL detail from TB) ═══
function AccountDrillDownModal({entityId,entityName,acct,from:fromProp,to:toProp,onClose,onChanged}){
  const[reloadKey,setReloadKey]=useState(0);
  const[q,setQ]=useState(''); // free-text filter over the visible transaction rows
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
  const prevDay=(d)=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return _ymd(x);};
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
  const matchQ=(l)=>{if(!q)return true;const s=q.toLowerCase();return (l.memo||'').toLowerCase().includes(s)||(l.offset||'').toLowerCase().includes(s)||(l.vendor||'').toLowerCase().includes(s)||('je-'+String(l.entry_num).padStart(4,'0')).includes(s)||(l.date||'').includes(s)||String(l.debit).includes(s)||String(l.credit).includes(s);};
  const doExport=()=>{const acctLabel=acct.code+' - '+acct.name;
    const d=[[entityName||'Account Detail'],[acctLabel],['Period: '+from+' to '+to],[],['Date','JE','Account',classTerm(),'Location','Memo','Offset Account','Vendor/Payee','Debit','Credit','Balance']];
    const F=[];
    d.push(['','','','','','Beginning Balance','','','','',begBal]);let prevRow=d.length-1;const begRow=prevRow;
    let r=begBal;const first=d.length;lines.forEach(l=>{r+=isDr?(l.debit-l.credit):(l.credit-l.debit);d.push([l.date,'JE-'+String(l.entry_num).padStart(4,'0'),acctLabel,l.class_name||'',l.location_name||'',l.memo,l.offset||'',l.vendor||'',l.debit||'',l.credit||'',r]);const cur=d.length-1,rr=cur+1;const delta=isDr?('I'+rr+'-J'+rr):('J'+rr+'-I'+rr);F.push({r:cur,c:10,f:'K'+(prevRow+1)+'+'+delta});prevRow=cur;});const last=d.length-1;
    const tr=d.length;d.push(['','','','','','Totals','','',totalDr,totalCr,r]);
    sumCols(F,tr,[8,9],first,last);F.push({r:tr,c:10,f:'K'+(begRow+1)+'+'+(isDr?('I'+(tr+1)+'-J'+(tr+1)):('J'+(tr+1)+'-I'+(tr+1)))});
    exportToExcel(d,'GL_'+acct.code+'_'+to+'.xlsx',{formulas:F});};
  return(<div style={S.modal}><div className="cl-modal-box" style={{...S.modalBox,width:'min(1100px,96vw)',maxWidth:'98vw',height:'88vh',maxHeight:'96vh',minWidth:'min(680px,96vw)',minHeight:420,resize:'both',overflow:'hidden',display:'flex',flexDirection:'column'}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14,gap:16}}>
      <div><div style={{fontSize:18,fontWeight:700,color:T.textBright}}>{acct.code} &mdash; {acct.name}</div>
        <div style={{display:'flex',alignItems:'center',gap:8,marginTop:6}}><label style={{fontSize:11,color:T.textMuted}}>From</label><input type="date" value={fromDraft} max={toDraft} onChange={e=>setFromDraft(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyDates();}} style={{...S.inputSm,padding:'4px 8px',fontSize:12}}/><label style={{fontSize:11,color:T.textMuted}}>To</label><input type="date" value={toDraft} min={fromDraft} onChange={e=>setToDraft(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')applyDates();}} style={{...S.inputSm,padding:'4px 8px',fontSize:12}}/><button onClick={applyDates} disabled={!datesDirty} style={{background:datesDirty?T.accent:T.bgElevated,border:'1px solid '+(datesDirty?T.accent:T.border),borderRadius:6,color:datesDirty?'#fff':T.textMuted,fontSize:11,fontWeight:600,padding:'4px 12px',cursor:datesDirty?'pointer':'default'}}>Refresh</button><button onClick={()=>{setFromDraft(fromProp);setToDraft(toProp);setFrom(fromProp);setTo(toProp);}} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'4px 8px',cursor:'pointer'}}>Reset</button>
        <span style={{width:1,height:16,background:T.border,margin:'0 2px'}}/>{PRESETS.map(([k,lbl])=><button key={k} onClick={()=>{const r=presetRange(k);setFromDraft(r.from);setToDraft(r.to);setFrom(r.from);setTo(r.to);}} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'4px 8px',cursor:'pointer'}}>{lbl}</button>)}<span style={{width:1,height:16,background:T.border,margin:'0 2px'}}/><input value={q} onChange={e=>setQ(e.target.value)} placeholder={'\\uD83D\\uDD0D Search memo / offset / vendor / JE\\u2026'} style={{...S.inputSm,padding:'4px 10px',fontSize:12,width:250}}/>{q&&<button onClick={()=>setQ('')} title="Clear search" style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'4px 8px',cursor:'pointer'}}>Clear</button>}</div>
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
               if(!matchQ(l))return null;
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
  const oneYrBefore=d=>{const x=new Date(d+'T00:00:00');x.setFullYear(x.getFullYear()-1);x.setDate(x.getDate()+1);return _ymd(x);};
  const Sec=({title,items,totalGetter,extraNiRow})=>(<><tr><td style={S.sectionHeader} colSpan={nColSpan}>{title}</td></tr>
    {items.map(a=><tr key={a.code}><td style={S.indentTd}>{a.name}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct({code:a.code,name:a.name,type:a.type,from:oneYrBefore(c.to),to:c.to})}>{fmt(val(a.code,i))}</td>)}{chgCells(ci=>val(a.code,ci))}</tr>)}
    {extraNiRow&&<tr><td style={{...S.indentTd,fontStyle:'italic',color:T.textMuted}}>Net Income (current period)</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,fontStyle:'italic'}}>{fmt(niCol(i))}</td>)}{chgCells(niCol)}</tr>}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td>{cols.map((c,i)=><td key={i} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(totalGetter(i))}</td>)}{chgCells(totalGetter)}</tr></>);
  const colHead=(c,i)=>prior&&i===0?'Prev':(c.label==='Total'?('As of '+anchor):c.label);
  const doExport=()=>{const hdr=['',...cols.map((c,i)=>colHead(c,i)),...(prior?['$ Change','% Change']:[])];
    const d=[[entityName||'Balance Sheet'],['Balance Sheet'],[],hdr];
    const F=[];const nP=cols.length;const chgC=nP+1,pctC=nP+2;const pcols=cols.map((c,i)=>1+i);
    // Every total is a live SUM(); Total L+E sums the two subtotal rows; $/%
    // change columns are formulas on every row.
    const push=(label,getter)=>{const row=[label,...cols.map((c,i)=>getter(i))];if(prior)row.push(getter(curI)-getter(priI),rptPct(getter(curI),getter(priI)));d.push(row);const r=d.length-1;
      if(prior){const rr=r+1,cc=XLC(1+curI),pc=XLC(1+priI);F.push({r,c:chgC,f:cc+rr+'-'+pc+rr});F.push({r,c:pctC,f:'IF(ABS('+pc+rr+')<0.005,"",('+cc+rr+'-'+pc+rr+')/ABS('+pc+rr+')*100)'});}
      return r;};
    d.push(['Assets']);const aF=d.length;assets.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));const aL=d.length-1;const aT=push('Total Assets',tA);sumCols(F,aT,pcols,aF,aL);
    d.push(['Liabilities']);const lF=d.length;liabs.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));const lL=d.length-1;const lT=push('Total Liabilities',ci=>sumC(liabs,ci));sumCols(F,lT,pcols,lF,lL);
    d.push(['Equity']);const eF=d.length;eq.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));push('  Net Income (current period)',niCol);const eL=d.length-1;const eT=push('Total Equity',tE);sumCols(F,eT,pcols,eF,eL);
    const leT=push('Total Liabilities + Equity',tLE);sumRows(F,leT,pcols,[lT,eT]);
    exportToExcel(d,'BS_'+anchor+'.xlsx',{formulas:F});};
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

// ── Income Statement date-range windows (period P&L), all anchored to the period-end ──
//   ytd     = Jan 1 of the period-end's year → period end (default).
//   month   = the full calendar month BEFORE the period-end's month (6/30 → May 2026).
//   quarter = the full calendar quarter BEFORE the period-end's quarter (6/30 → Q1 2026).
//   year    = the prior full calendar year (2026 → 2025).
//   custom  = the user-entered [from,to] (to defaults to the period end).
function isPnlWindow(filter,anchor,cf,ct){
  const a=_mkDate(anchor),Y=a.getFullYear(),M=a.getMonth();
  if(filter==='ytd')return{from:Y+'-01-01',to:anchor};
  if(filter==='month')return{from:_ymd(new Date(Y,M-1,1)),to:_ymd(new Date(Y,M,0))};
  if(filter==='quarter'){const ps=(Math.floor(M/3)-1)*3;return{from:_ymd(new Date(Y,ps,1)),to:_ymd(new Date(Y,ps+3,0))};}
  if(filter==='year')return{from:(Y-1)+'-01-01',to:(Y-1)+'-12-31'};
  if(filter==='custom')return{from:(/^\d{4}-\d{2}-\d{2}$/.test(cf)?cf:null),to:(/^\d{4}-\d{2}-\d{2}$/.test(ct)?ct:anchor)};
  return{from:null,to:anchor};
}
// The period-end's OWN period through the period end (the "current" period for
// Compare): month → 1st of that month, quarter → 1st of that quarter, year → Jan 1;
// each ending at the period end. For 6/30/26: month=Jun 1–30, quarter=Apr 1–Jun 30
// (Q2), year=Jan 1–Jun 30.
function isCurrentPeriodWindow(filter,anchor){
  const a=_mkDate(anchor),Y=a.getFullYear(),M=a.getMonth();
  if(filter==='month')return{from:_ymd(new Date(Y,M,1)),to:anchor};
  if(filter==='quarter')return{from:_ymd(new Date(Y,Math.floor(M/3)*3,1)),to:anchor};
  if(filter==='year')return{from:Y+'-01-01',to:anchor};
  return{from:null,to:anchor};
}
// Prior comparative window: YTD compares to the SAME dates a year earlier (prior-year
// YTD); everything else uses the previous equivalent calendar period.
function isPnlPrior(filter,win,anchor){
  if(!win||!win.from)return null;
  if(filter==='ytd'){const a=_mkDate(anchor);return{label:'Prev',from:(a.getFullYear()-1)+'-01-01',to:_ymd(new Date(a.getFullYear()-1,a.getMonth(),a.getDate()))};}
  // Custom → the SAME calendar dates one year earlier (prior-year same period), e.g.
  // 1/1/26–2/28/26 → 1/1/25–2/28/25.
  if(filter==='custom'){const f=_mkDate(win.from),t=_mkDate(win.to);return{label:'Prev',from:_ymd(new Date(f.getFullYear()-1,f.getMonth(),f.getDate())),to:_ymd(new Date(t.getFullYear()-1,t.getMonth(),t.getDate()))};}
  return rptPriorWindow(win);
}
const IS_DATE_FILTERS=[['ytd','Year to Date'],['month','Last Month'],['quarter','Last Quarter'],['year','Last Year'],['custom','Custom']];
function IncomeStatement({entityId,entityName,from,setFrom,to,setTo,canEdit=true}){
  const[drillAcct,setDrillAcct]=useState(null);const[rk,setRk]=useState(0);
  const[dateFilter,setDateFilter]=useState('ytd');const[compare,setCompare]=useState(false);
  const[customFrom,setCustomFrom]=useState('');const[customTo,setCustomTo]=useState('');
  const colMode='total';// Income Statement is always a single period column (Columns option removed).
  const anchor=/^\d{4}-\d{2}-\d{2}$/.test(to)?to:today();
  const win=useMemo(()=>isPnlWindow(dateFilter,anchor,customFrom,customTo),[dateFilter,anchor,customFrom,customTo]);
  // Columns. Compare off → the single selected window. Compare on → the CURRENT
  // (most recent) period FIRST, then the comparison period; $ / % change is
  // measured current − comparison:
  //   • Last Month/Quarter/Year → CURRENT period (the period-end's own month/quarter/
  //     year through the period end) then the selected "last" period. So Last Quarter
  //     @ 6/30 → Q2 (current), Q1 (last); Last Month → June, May.
  //   • Year to Date → current YTD, then prior-year YTD.
  //   • Custom → the custom window, then the immediately preceding equal-length window.
  const cols=useMemo(()=>{
    const sel={label:'Total',from:win.from,to:win.to};
    if(!compare)return[sel];
    if(dateFilter==='ytd'||dateFilter==='custom'){const p=isPnlPrior(dateFilter,win,anchor);return p?[sel,p]:[sel];}
    const cur=isCurrentPeriodWindow(dateFilter,anchor);
    return[{label:'Total',from:cur.from,to:cur.to},sel];
  },[compare,dateFilter,win.from,win.to,anchor]);
  const prior=(compare&&cols.length>1)?cols[cols.length-1]:null;
  // Column reorder (Liting #2): `colOrder` holds display order as original-column
  // indices into `cols`. Drag a period-column header to rearrange. Resets whenever
  // the underlying set of columns changes (different filter/column mode).
  const[colOrder,setColOrder]=useState(null);const dragFrom=useRef(null);
  const colSig=cols.map(c=>c.label+'|'+(c.from||'')+'|'+c.to).join(',');
  useEffect(()=>{setColOrder(null);},[colSig]);
  const ord=(colOrder&&colOrder.length===cols.length)?colOrder:cols.map((_,i)=>i);
  const moveCol=(fromPos,toPos)=>{setColOrder(prev=>{const base=(prev&&prev.length===cols.length)?[...prev]:cols.map((_,i)=>i);const[x]=base.splice(fromPos,1);base.splice(toPos,0,x);return base;});};
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
  const nCols=cols.length;const curI=0;const priI=(compare&&nCols>1)?nCols-1:-1;// current column is first
  const pctTxt=p=>p==null?'—':(p>=0?'+':'')+p.toFixed(1)+'%';
  const chgCells=(getter)=>{if(!prior)return null;const c=getter(curI),p=getter(priI),pc=rptPct(c,p);return[<td key="d" style={S.tdR}>{fmt(c-p)}</td>,<td key="p" style={{...S.tdR,color:(c-p)>=0?T.green:T.red}}>{pctTxt(pc)}</td>];};
  const nColSpan=1+nCols+(prior?2:0);
  const Sec=({title,items})=>(<><tr><td style={S.sectionHeader} colSpan={nColSpan}>{title}</td></tr>
    {items.map(a=><tr key={a.code}><td style={S.indentTd}>{a.name}</td>{ord.map(oi=><td key={oi} style={{...S.tdR,borderBottom:'1px solid '+T.borderLight,cursor:'pointer'}} onClick={()=>setDrillAcct({code:a.code,name:a.name,type:a.type,from:cols[oi].from,to:cols[oi].to})}>{fmt(val(a.code,oi))}</td>)}{chgCells(ci=>val(a.code,ci))}</tr>)}
    <tr style={S.subtotalRow}><td style={{...S.td,fontWeight:600,paddingLeft:14}}>Total {title}</td>{ord.map(oi=><td key={oi} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(sumC(items,oi))}</td>)}{chgCells(ci=>sumC(items,ci))}</tr></>);
  const TotalRow=({label,getter,big})=>(<tr style={big?S.grandTotalRow:{background:T.bgElevated}}><td style={big?{...S.tdBold,fontSize:15}:{...S.td,fontWeight:700,color:T.textBright}}>{label}</td>{ord.map(oi=><td key={oi} style={big?{...S.tdBold,textAlign:'right',fontSize:16,color:getter(oi)>=0?T.green:T.red}:{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(getter(oi))}</td>)}{chgCells(getter)}</tr>);
  // Column heading names the calendar period the column's data covers — a clean
  // "Q2 2026" / "May 2026" / "2025" when it's a whole month/quarter/year, otherwise
  // the period-end date (e.g. "6/30/26" for a year-to-date or partial window). Never
  // shows a bare "Amount" — every column states its period. Updates with the selection.
  const colHead=(c)=>periodLabel(c.from,c.to)||(c.from?fmtMDY(c.to):'Amount');
  // Report subtitle. Total-Only names the displayed columns by their period-end
  // dates ("For the Quarters Ended 6/30/26 and 3/31/26"); other column modes keep
  // the range + column-mode label.
  const dispCols=ord.map(oi=>cols[oi]);
  const subtitle=(()=>{
    if(dateFilter==='ytd')return 'Year to Date Ended '+fmtMDY(win.to)+(prior?' vs '+fmtMDY(prior.to):'');
    if(dateFilter==='custom')return win.from?('For the Period '+fmtMDY(win.from)+' – '+fmtMDY(win.to)+(prior?' vs '+fmtMDY(prior.from)+' – '+fmtMDY(prior.to):'')):('Through '+fmtMDY(win.to));
    return endedHeading(dispCols,dateFilter,anchor);
  })();
  const doExport=()=>{
    // 0-based column index -> A1 letter (B, C, … AA, …).
    const colL=n=>{let s='';n=n+1;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
    const hdr=['Account',...ord.map(oi=>colHead(cols[oi],oi)),...(prior?['$ Change','% Change']:[])];
    const d=[[entityName||'Income Statement'],['Income Statement'],[subtitle],[],hdr];
    // Formula cells (Liting #3): EVERY calculated cell is a live formula, not just
    // subtotals — subtotal rows use SUM(); Net Income references the section totals;
    // and the $ Change / % Change columns are formulas on every row (detail, subtotal,
    // Net Income). Only the individual account amounts stay as values, since they are
    // the source figures with nothing to compute from. Values are also written as
    // cached fallbacks by exportToExcel so the sheet reads correctly before recalc.
    const formulas=[];
    const push=(label,getter)=>{
      const row=[label,...ord.map(oi=>getter(oi))];
      if(prior)row.push(getter(curI)-getter(priI),rptPct(getter(curI),getter(priI)));
      d.push(row);const r=d.length-1;
      if(prior){const rr=r+1,curCol=colL(1+ord.indexOf(curI)),priCol=colL(1+ord.indexOf(priI)),chgC=1+ord.length,pctC=2+ord.length;
        formulas.push({r,c:chgC,f:curCol+rr+'-'+priCol+rr});
        formulas.push({r,c:pctC,f:'IF(ABS('+priCol+rr+')<0.005,"",('+curCol+rr+'-'+priCol+rr+')/ABS('+priCol+rr+')*100)'});}
      return r;};
    const secRow={};
    const sections=[['Revenue',rev,1],['Cost of Goods Sold',cogs,-1],['Operating Expenses',opex,-1],['Other Expenses',other,-1]];
    sections.forEach(([t,items])=>{if(!items.length)return;
      d.push([t]);
      const firstIdx=d.length; // 0-based index of the first item row (next push)
      items.forEach(a=>push('  '+a.name,ci=>val(a.code,ci)));
      const lastIdx=d.length-1;
      const totIdx=push('Total '+t,ci=>sumC(items,ci));secRow[t]=totIdx;
      ord.forEach((oi,pos)=>{const c=1+pos;formulas.push({r:totIdx,c,f:'SUM('+colL(c)+(firstIdx+1)+':'+colL(c)+(lastIdx+1)+')'});});
    });
    const niIdx=push('Net Income',ci=>ni(ci));
    const present=sections.filter(([t,items])=>items.length);
    ord.forEach((oi,pos)=>{const c=1+pos;let f=present.map(([t,items,sign])=>(sign>0?'+':'-')+colL(c)+(secRow[t]+1)).join('');if(f[0]==='+')f=f.slice(1);if(f)formulas.push({r:niIdx,c,f});});
    exportToExcel(d,'IS_'+anchor+'.xlsx',{formulas});};
  return(<div><div style={S.card}><div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-end',marginBottom:16,flexWrap:'wrap',gap:10}}>
    <div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>Period end</label><input style={S.inputSm} type="date" value={to} onChange={e=>setTo(e.target.value)}/></div>
      <div><label style={S.label}>Date range</label><select style={S.inputSm} value={dateFilter} onChange={e=>setDateFilter(e.target.value)}>{IS_DATE_FILTERS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>
      {dateFilter==='custom'&&<><div><label style={S.label}>From</label><input style={S.inputSm} type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)}/></div>
      <div><label style={S.label}>To</label><input style={S.inputSm} type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)}/></div></>}
      <div><label style={S.label}>Compare</label><label style={{display:'flex',alignItems:'center',gap:6,fontSize:12,color:T.textMuted,height:34}} title="Adds a prior-period column with $ and % change"><input type="checkbox" checked={compare} onChange={e=>setCompare(e.target.checked)}/> Prev period ($ / %)</label></div>
    </div>
    <div style={{display:'flex',gap:8,alignItems:'center'}}><MemorizeBar entityId={entityId} reportType='is' currentConfig={{to,dateFilter,compare,customFrom,customTo}} onApply={(c)=>{if(c.to)setTo(c.to);if(c.dateFilter)setDateFilter(c.dateFilter);if(typeof c.compare==='boolean')setCompare(c.compare);setCustomFrom(c.customFrom||'');setCustomTo(c.customTo||'');}} canEdit={canEdit}/><button style={S.btnExport} onClick={doExport}>Export Excel</button></div></div>
    <div style={S.reportHeader}>{entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}<div style={{fontSize:20,fontWeight:700,color:T.textBright}}>Income Statement</div><div style={{fontSize:13,color:T.textMuted}}>{subtitle}</div></div>
    <div style={{overflowX:'auto'}}><table style={{...S.table,minWidth:520,margin:'0 auto'}}><thead><tr><th style={S.th}>Account</th>{ord.map((oi,pos)=><th key={oi} draggable onDragStart={()=>{dragFrom.current=pos;}} onDragOver={e=>e.preventDefault()} onDrop={()=>{if(dragFrom.current!=null&&dragFrom.current!==pos)moveCol(dragFrom.current,pos);dragFrom.current=null;}} style={{...S.thR,cursor:cols.length>1?'grab':'default',userSelect:'none'}} title={cols.length>1?'Drag to reorder columns':undefined}>{colHead(cols[oi],oi)}</th>)}{prior&&<><th style={S.thR}>$ Change</th><th style={S.thR}>% Change</th></>}</tr></thead><tbody>
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
  const[projects,setProjects]=useState([]);const[projFilter,setProjFilter]=useState(''); // optional single-project filter
  const[runToken,setRunToken]=useState(0); // bumped after each successful run to auto-export
  const prevDay=(d)=>{const x=new Date(d+'T00:00:00');x.setDate(x.getDate()-1);return _ymd(x);};
  const isBS=(t)=>t==='Asset'||t==='Liability'||t==='Equity';
  useEffect(()=>{api.getAccounts(entityId).then(setAccounts).catch(()=>setAccounts([]));},[entityId]);
  useEffect(()=>{api.getProjects(entityId).then(setProjects).catch(()=>setProjects([]));},[entityId]);
  useEffect(()=>{if(pendingConfig){if(pendingConfig.sel)setSel(pendingConfig.sel);setFrom(pendingConfig.from||'');setTo(pendingConfig.to||'');if(pendingConfig.groupBy)setGroupBy(pendingConfig.groupBy);clearPending&&clearPending();}},[]);
  useEffect(()=>{if(pendingConfig){if(pendingConfig.sel)setSel(pendingConfig.sel);if(pendingConfig.dim)setDim(pendingConfig.dim);setFrom(pendingConfig.from||'');setTo(pendingConfig.to||'');clearPending&&clearPending();}},[]);
  const toggle=code=>setSel(s=>s.includes(code)?s.filter(c=>c!==code):[...s,code]);
  const filteredAccts=accounts.filter(a=>!acctSearch||acctLabel(a.code,a.name).toLowerCase().includes(acctSearch.toLowerCase()));
  const run=async()=>{
    if(!sel.length){setErr('Select at least one account');return;}
    setLoading(true);setErr('');setRows(null);setBegRows([]);
    try{
      const all=await api.getGLDetail(entityId,{from:from||undefined,to:to||undefined,...(projFilter?{project_id:projFilter}:{})});
      const selSet=new Set(sel);
      setRows((all.lines||all||[]).filter(l=>selSet.has(l.account_code)));
      // Comparative: pull the equal-length prior window's activity for the same accounts.
      if(compare&&from&&to){const pw=priorWindow(from,to);if(pw){const pall=await api.getGLDetail(entityId,{from:pw.from,to:pw.to,...(projFilter?{project_id:projFilter}:{})});setPriorRows((pall.lines||pall||[]).filter(l=>selSet.has(l.account_code)));}else setPriorRows([]);}else setPriorRows([]);
      // Beginning balance only applies to balance-sheet accounts and only when a
      // start date is set (otherwise the period runs from inception → beg = 0).
      const bsSel=accounts.filter(a=>selSet.has(a.code)&&isBS(a.type));
      if(from&&bsSel.length){
        const bals=await api.getBalances(entityId,{as_of:prevDay(from),...(projFilter?{project_id:projFilter}:{})});
        const byCode=new Map((bals||[]).map(b=>[b.code,b.balance]));
        setBegRows(bsSel.map(a=>({code:a.code,name:a.name,type:a.type,balance:byCode.get(a.code)||0})).filter(r=>r.balance!==0));
      }
      setRunToken(x=>x+1);
    }catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const groupKey=l=>groupBy==='class'?(l.class_name||'(no class)'):groupBy==='location'?(l.location_name||'(no location)'):groupBy==='project'?(l.project_name||'(no project)'):'All';
  const groups=(()=>{if(!rows)return[];const m=new Map();rows.forEach(l=>{const k=groupKey(l);if(!m.has(k))m.set(k,[]);m.get(k).push(l);});return[...m.entries()].sort((a,b)=>a[0].localeCompare(b[0]));})();
  const amt=l=>{const isDr=l.account_type==='Asset'||l.account_type==='Expense';return isDr?((l.debit||0)-(l.credit||0)):((l.credit||0)-(l.debit||0));};
  const bsCodesSet=new Set(accounts.filter(a=>sel.includes(a.code)&&isBS(a.type)).map(a=>a.code));
  const grand=rows?rows.filter(l=>bsCodesSet.has(l.account_code)).reduce((s,l)=>s+amt(l),0):0;
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
    const _projObj=projFilter?projects.find(p=>String(p.id)===String(projFilter)):null;
    const _projLabel=_projObj?(_projObj.code?_projObj.code+' — '+_projObj.name:_projObj.name):'';
    const d=[[entityName||'Custom Detail Report'],['Custom Detail Report'],...(_projLabel?[['Project: '+_projLabel]]:[]),['Period: '+(from||'Begin')+' to '+(to||today())],[]];
    const F=[];// account subtotals SUM detail rows; group totals SUM account subtotals; PERIOD ACTIVITY SUMs group totals.
    if(begRows.length>0){
      d.push(['Beginning Balances — Balance Sheet accounts as of '+from]);
      d.push(['','Account','','','','','Balance']);
      const bF=d.length;begRows.forEach(b=>d.push(['',b.code+' '+b.name,'','','','',b.balance]));const bL=d.length-1;
      const bT=d.length;d.push(['','','','','','Total Beginning Balance',begTotal]);d.push([]);
      sumCols(F,bT,[6],bF,bL);
    }
    const amtHdr=cols.map(c=>c.label);const cmpHdr=compare?['Prev Period','$ Change','% Change']:[];
    const pctN=(cur,pri)=>pri!==0?+(((cur-pri)/Math.abs(pri))*100).toFixed(1):'';
    const pcols=[];for(let k=0;k<cols.length;k++)pcols.push(4+k);if(showTotal)pcols.push(4+cols.length);
    const groupTotRows=[];
    groups.forEach(([g,lines])=>{
      if(groupBy!=='none')d.push([g]);
      d.push(['Account','Date','JE','Description',...amtHdr,...(showTotal?['Total']:[]),...cmpHdr]);
      const _byA=[];const _im=new Map();lines.forEach(l=>{const k=l.account_code;if(!_im.has(k)){_im.set(k,_byA.length);_byA.push([k,l.account_name,[]]);}_byA[_im.get(k)][2].push(l);});
      const acctTotRows=[];
      _byA.forEach(([acode,aname,alines])=>{
        const aF=d.length;
        alines.forEach(l=>{const a=amt(l);const ci=colIdxOf(l.date);const cells=cols.map((c,k)=>(colMode==='total'||k===ci)?a:'');d.push([l.account_code+' '+l.account_name,l.date,'JE-'+String(l.entry_num).padStart(4,'0'),l.description||l.memo||'',...cells,...(showTotal?[a]:[]),...(compare?['','','']:[])]);});
        const aL=d.length-1;const as=sumByCol(alines);const aT=d.length;d.push(['Total '+acode+' '+aname,'','','',...as.arr,...(showTotal?[as.tot]:[]),...(compare?['','','']:[])]);
        sumCols(F,aT,pcols,aF,aL);acctTotRows.push(aT);
      });
      const {arr,tot}=sumByCol(lines);const pri=priorGroupMap.get(g)||0;
      const gT=d.length;d.push(['Total'+(groupBy!=='none'?' for '+g:''),'','','',...arr,...(showTotal?[tot]:[]),...(compare?[pri,tot-pri,pctN(tot,pri)]:[])]);d.push([]);
      sumRows(F,gT,pcols,acctTotRows);groupTotRows.push(gT);
    });
    const gg=sumByCol(rows||[]);
    const pa=d.length;d.push(['PERIOD ACTIVITY','','','',...gg.arr,...(showTotal?[gg.tot]:[]),...(compare?[priorGrand,gg.tot-priorGrand,pctN(gg.tot,priorGrand)]:[])]);
    sumRows(F,pa,pcols,groupTotRows);
    if(begRows.length>0)d.push(['ENDING BALANCE (BS accts: beginning + activity)','','','',...(colMode==='total'?[begTotal+grand]:[...cols.map(()=>''),begTotal+grand]),...(compare?['','','']:[])]);
    exportToExcel(d,'Custom_Detail_'+(to||today())+'.xlsx',{formulas:F});
  };
  useEffect(()=>{if(runToken>0)doExport();},[runToken]); // auto-export after each Run Report
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
          {dimsEnabled&&<div><label style={S.label}>Group by</label><select style={{...S.inputSm,width:'100%'}} value={groupBy} onChange={e=>setGroupBy(e.target.value)}><option value="none">No grouping</option><option value="class">{classTerm()==='Class'?'Class / Investor':classTerm()}</option><option value="location">Location</option><option value="project">Project</option></select></div>}
          {projects.length>0&&<div style={{marginTop:10}}><label style={S.label}>Project</label><select style={{...S.inputSm,width:'100%'}} value={projFilter} onChange={e=>setProjFilter(e.target.value)}><option value="">All projects</option>{projects.map(p=><option key={p.id} value={p.id}>{p.code?p.code+' — '+p.name:p.name}</option>)}</select></div>}
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
        {(()=>{const byA=[];const im=new Map();lines.forEach(l=>{const k=l.account_code;if(!im.has(k)){im.set(k,byA.length);byA.push([k,l.account_name,[]]);}byA[im.get(k)][2].push(l);});return byA.map(([acode,aname,alines])=>{const as=sumByCol(alines);return<Fragment key={'acct'+acode}>
          {alines.map((l,i)=>{const a=amt(l);const ci=colIdxOf(l.date);return<tr key={acode+'-'+i}>
            <td style={S.td}>{l.account_code} {l.account_name}</td><td style={{...S.td,whiteSpace:'nowrap'}}>{l.date}</td><td style={S.td}>JE-{String(l.entry_num).padStart(4,'0')}</td><td style={S.td}>{l.description||l.memo||''}</td>
            {cols.map((c,k)=><td key={k} style={{...S.tdR,color:a<0?T.red:T.textBright}}>{(colMode==='total'||k===ci)?fmt(a):''}</td>)}
            {showTotal&&<td style={{...S.tdR,color:a<0?T.red:T.textBright}}>{fmt(a)}</td>}
            {compare&&<><td style={S.tdR}></td><td style={S.tdR}></td><td style={S.tdR}></td></>}
          </tr>;})}
          <tr style={{borderTop:'1px solid '+T.border}}><td style={{...S.td,fontWeight:600,fontStyle:'italic',color:T.textMuted}} colSpan={descCols}>Total {acode} {aname}</td>{as.arr.map((v,k)=><td key={k} style={{...S.tdR,fontWeight:600}}>{fmt(v)}</td>)}{showTotal&&<td style={{...S.tdR,fontWeight:600}}>{fmt(as.tot)}</td>}{compare&&<><td style={S.tdR}></td><td style={S.tdR}></td><td style={S.tdR}></td></>}</tr>
        </Fragment>;});})()}
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
    const F=[];const nC=data.columns.length;const totC=1+nC;const dataCols=[];for(let k=1;k<=nC;k++)dataCols.push(k);
    const first=d.length;
    data.rows.forEach(r=>d.push([r.name,...data.columns.map(c=>r.cells[c.code]||0),r.total]));
    const last=d.length-1;
    for(let r=first;r<=last;r++)F.push({r,c:totC,f:'SUM(B'+(r+1)+':'+XLC(nC)+(r+1)+')'});// row total = across its columns
    const tr=d.length;d.push(['Total',...data.columns.map(c=>data.column_totals[c.code]||0),data.grand_total]);
    sumCols(F,tr,dataCols,first,last);F.push({r:tr,c:totC,f:'SUM(B'+(tr+1)+':'+XLC(nC)+(tr+1)+')'});
    exportToExcel(d,'Pivot_'+dim+'_'+(to||today())+'.xlsx',{formulas:F});
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
    {data&&<div className="cl-scroll" style={scrollBox()}>
      {data.rows.length===0?<div style={{padding:24,color:T.textDim}}>No activity for the selected accounts/period.</div>:
      <table style={S.table}><thead><tr><th style={{...S.th,position:'sticky',left:0,background:T.bgCard}}>{dim==='class'?(classTerm()==='Class'?'Class / Investor':classTerm()):dim==='location'?'Location':'Project'}</th>{data.columns.map(c=><th key={c.code} style={S.thR} title={c.code+' '+c.name}>{c.name||c.code}</th>)}<th style={S.thR}>Total</th></tr></thead>
      <tbody>{data.rows.map(r=><tr key={r.id}><td style={{...S.td,position:'sticky',left:0,background:T.bgCard,fontWeight:500}}>{r.name}</td>{data.columns.map(c=><td key={c.code} style={S.tdR}>{r.cells[c.code]?fmt(r.cells[c.code]):''}</td>)}<td style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(r.total)}</td></tr>)}
        <tr style={S.grandTotalRow}><td style={{...S.tdBold,position:'sticky',left:0,background:T.bgCard}}>Total</td>{data.columns.map(c=><td key={c.code} style={{...S.tdR,fontWeight:700,color:T.textBright}}>{fmt(data.column_totals[c.code]||0)}</td>)}<td style={{...S.tdBold,textAlign:'right',color:T.textBright}}>{fmt(data.grand_total)}</td></tr>
      </tbody></table>}
    </div>}
  </div>);
}

// ═══ AP Aging Detail (Q5: open bills from Bill.com, bucketed by days past due) ═══
// Upload an A/P aging detail (the same GL/prior-system report the GL import came
// from) to set the Bill.com sync cutoff. The latest bill date on the report is
// the last invoice already booked in the GL, so we store cutoff = latest + 1 day
// (the sync engine's cutoff is exclusive: a bill syncs when invoiceDate >= cutoff),
// making Bill.com skip everything already in the GL and pull only newer bills.
function ApAgingCutoffModal({entityId,entityName,onClose,onDone}){
  const [parsing,setParsing]=useState(false);
  const [preview,setPreview]=useState(null);
  const [err,setErr]=useState('');
  const [saving,setSaving]=useState(false);
  const [fileName,setFileName]=useState('');
  const [curCutoff,setCurCutoff]=useState(null);
  useEffect(()=>{ let ok=true; (async()=>{ try{ const c=await api.getBillcomConfig(entityId); if(ok)setCurCutoff((c&&c.sync_cutoff_date)||''); }catch(e){ if(ok)setCurCutoff(''); } })(); return()=>{ok=false;}; },[]);
  const fmtDate=(v)=>{ if(v===null||v===undefined||v==='')return null; if(v instanceof Date&&!isNaN(v)){return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');} const s=String(v).trim(); if(!s)return null; if(s.indexOf('/')>=0){const p=s.split(' ')[0].split('/'); if(p.length===3){let y=p[2]; if(y.length===2)y='20'+y; return y+'-'+p[0].padStart(2,'0')+'-'+p[1].padStart(2,'0');}} if(s.length>=10&&s.charAt(4)==='-')return s.slice(0,10); return null; };
  const num=(v)=>{ if(v===null||v===undefined||v==='')return null; if(typeof v==='number')return v; const n=parseFloat(String(v).split(',').join('').split('$').join('').split('(').join('-').split(')').join('')); return isNaN(n)?null:n; };
  const addDay=(ymd)=>{ const p=ymd.split('-').map(Number); const dt=new Date(Date.UTC(p[0],p[1]-1,p[2])); dt.setUTCDate(dt.getUTCDate()+1); return dt.toISOString().slice(0,10); };
  const findIn=(H,names)=>{ for(const nm of names){ for(let i=0;i<H.length;i++){ if(H[i].indexOf(nm)>=0)return i; } } return -1; };
  const onFile=async(file)=>{
    setErr('');setPreview(null);setParsing(true);setFileName(file.name);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array',cellDates:true});
      let rows=null,headerIdx=-1,ci=null;
      for(const name of wb.SheetNames){
        const rws=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''});
        for(let i=0;i<Math.min(rws.length,25);i++){
          const H=rws[i].map(c=>String(c).toLowerCase().trim());
          const bd=findIn(H,['bill date','invoice date','bill dt','document date']);
          const amt=findIn(H,['open balance','balance','amount','total','current']);
          const dcol=bd>=0?bd:findIn(H,['date']);
          if(dcol>=0&&amt>=0){ rows=rws;headerIdx=i; ci={ date:dcol, amt, vendor:findIn(H,['vendor name','vendor','payee','name']), doc:findIn(H,['bill no','bill number','document no','invoice no','ref no','ref','num']) }; break; }
        }
        if(rows)break;
      }
      if(!rows)throw new Error('Could not find an A/P aging sheet — need a header row with a Bill date column and an Amount/Balance column.');
      let fileAsOf=null; for(let i=0;i<headerIdx;i++){ const rr=rows[i]||[]; for(let c=0;c<rr.length;c++){ if(String(rr[c]||'').toLowerCase().indexOf('as of')>=0){ const mm=String(rr[c]).match(/(\d{4}-\d{2}-\d{2})/); if(mm)fileAsOf=mm[1]; else if(c+1<rr.length){ const ff=fmtDate(rr[c+1]); if(ff)fileAsOf=ff; } } } }
      const items=[]; let latest=null; let vend=null;
      for(let i=headerIdx+1;i<rows.length;i++){
        const r=rows[i];
        const vcell=String(ci.vendor>=0?r[ci.vendor]:'').trim();
        const low=vcell=>vcell.toLowerCase();
        if(low(vcell).indexOf('grand total')>=0)break;
        if(low(vcell).indexOf('total for')>=0||low(vcell).indexOf('total ')===0)continue;
        const dateCell=fmtDate(r[ci.date]);
        const amt=num(r[ci.amt]);
        if(ci.vendor>=0&&vcell&&(dateCell===null||amt===null)){ vend=vcell; continue; }
        if(dateCell===null||amt===null)continue;
        if(ci.vendor>=0&&vcell)vend=vcell;
        const docCell=ci.doc>=0?String(r[ci.doc]||'').trim():'';
        items.push({vendor:vend||'(no vendor)',invoice_number:docCell,bill_date:dateCell,amount:Math.round(amt*100)/100});
        if(!latest||dateCell>latest)latest=dateCell;
      }
      if(!items.length)throw new Error('No bill rows with a date and amount were found below the header row.');
      if(!latest)throw new Error('Could not read a bill date from any row.');
      const total=Math.round(items.reduce((s,x)=>s+x.amount,0)*100)/100;
      const vendCount=new Set(items.map(x=>x.vendor)).size;
      const cutoff=addDay(latest);
      let glBalance=null,recon=null;
      try{ const ag=await api.getApAging(entityId,latest); glBalance=ag.gl_balance; recon=Math.round((glBalance-total)*100)/100; }catch(e){}
      setPreview({count:items.length,total,vendCount,latest,cutoff,glBalance,recon,items,asOf:fileAsOf});
    }catch(e){ setErr(e.message||String(e)); } finally{ setParsing(false); }
  };
  const doSave=async()=>{ if(!preview)return; setSaving(true);setErr('');
    try{ await api.setBillcomCutoff(entityId,preview.cutoff,preview.items,preview.asOf); onDone&&onDone(preview.cutoff); }
    catch(e){ setErr(e.message||String(e)); } finally{ setSaving(false); }
  };
  const tie=preview&&preview.recon!==null?Math.abs(preview.recon)<0.005:null;
  return(<div style={S.modal} onClick={onClose}><div className="cl-modal-box" style={{...S.modalBox,maxWidth:640}} onClick={e=>e.stopPropagation()}>
    <button style={S.modalClose} onClick={onClose}>&times;</button>
    <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginBottom:4}}>Set Bill.com sync cutoff from A/P aging</div>
    <div style={{fontSize:12,color:T.textMuted,marginBottom:16}}>Upload the A/P aging detail (.xlsx) for the same period as your GL import. The latest bill date on the report is the last bill already booked in the GL, so Bill.com will sync only bills dated after it — no double-counting. This sets the cutoff on {entityName}'s Bill.com config; it does not import bills.</div>
    <div style={{...S.card,background:T.bgElevated,padding:16,marginBottom:14,textAlign:'center'}}>
      <input id="ap-cutoff-file" type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>{const f=e.target.files[0];if(f)onFile(f); e.target.value='';}}/>
      <label htmlFor="ap-cutoff-file" style={{...S.btnP,display:'inline-block',cursor:'pointer'}}>{parsing?'Reading file…':'Choose .xlsx file'}</label>
      {fileName&&<div style={{fontSize:12,color:T.textMuted,marginTop:8}}>{fileName}</div>}
    </div>
    {err&&<div style={{...S.err,marginBottom:12}}>{err}</div>}
    {preview&&<div style={{...S.card,padding:14,marginBottom:14}}>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Bills read</span><span style={{fontWeight:600}}>{preview.count} across {preview.vendCount} vendors</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Grand total</span><span style={{fontWeight:600}}>{'$'+fmt(preview.total)}</span></div>
      {preview.glBalance!==null&&<div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>GL control balance (as of {preview.latest})</span><span>{'$'+fmt(preview.glBalance)}</span></div>}
      {preview.recon!==null&&<div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Difference</span><span style={{fontWeight:700,color:tie?T.green:T.orange}}>{tie?'ties out':'$'+fmt(preview.recon)+' off'}</span></div>}
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'6px 0 3px',marginTop:6,borderTop:'1px solid '+T.border}}><span style={{color:T.textMuted}}>Latest bill date in report</span><span style={{fontWeight:700}}>{preview.latest}</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>Current cutoff</span><span>{curCutoff?curCutoff:'(none set)'}</span></div>
      <div style={{display:'flex',justifyContent:'space-between',fontSize:13,padding:'3px 0'}}><span style={{color:T.textMuted}}>New cutoff</span><span style={{fontWeight:700,color:T.accent}}>{preview.cutoff}</span></div>
      <div style={{fontSize:12,color:T.textMuted,marginTop:8}}>Bill.com will sync bills dated {preview.cutoff} and later. Everything through {preview.latest} is treated as already in the GL.</div>
    </div>}
    <div style={{display:'flex',gap:10,justifyContent:'flex-end'}}>
      <button style={S.btnS} onClick={onClose} disabled={saving}>Cancel</button>
      <button style={{...S.btnP,opacity:(!preview||saving)?0.5:1}} onClick={doSave} disabled={!preview||saving}>{saving?'Saving…':(preview?'Set cutoff to '+preview.cutoff:'Set cutoff')}</button>
    </div>
  </div></div>);
}

function ApAgingReport({entityId,entityName,canEdit=true,pendingConfig,clearPending}){
  const[asOf,setAsOf]=useState(today());
  useEffect(()=>{if(pendingConfig){if(pendingConfig.asOf)setAsOf(pendingConfig.asOf);clearPending&&clearPending();}},[]);
  const[data,setData]=useState(null);const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  const[viewEntry,setViewEntry]=useState(null);
  const[entryLoading,setEntryLoading]=useState(false);
  const[showUpload,setShowUpload]=useState(false);
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
    const F=[];const venTotRows=[];
    data.vendors.forEach(g=>{
      const vF=d.length;
      g.rows.forEach(r=>d.push([r.date,r.type,r.num,r.vendor,r.due_date,r.past_due_days,
        r.bucket==='current'?r.amount:'',r.bucket==='d1_30'?r.amount:'',r.bucket==='d31_60'?r.amount:'',r.bucket==='d61_90'?r.amount:'',r.bucket==='d91_plus'?r.amount:'','',r.amount]));
      const vL=d.length-1;const vT=d.length;d.push(['Total '+g.vendor,'','','','','',g.subtotal.current,g.subtotal.d1_30,g.subtotal.d31_60,g.subtotal.d61_90,g.subtotal.d91_plus,'',g.subtotal.total]);
      sumCols(F,vT,[6,7,8,9,10,12],vF,vL);venTotRows.push(vT);
    });
    let glTotRow=null;
    if((data.gl_rows||[]).length){
      d.push([]);d.push(['GL ENTRIES (imported / non-Bill.com — not aged)']);
      const gF=d.length;data.gl_rows.forEach(r=>d.push([r.date,'GL','JE-'+String(r.entry_num).padStart(4,'0'),'',(r.memo||''),'','','','','','',r.amount,r.amount]));const gL=d.length-1;
      glTotRow=d.length;d.push(['Total GL Entries','','','','','','','','','','',data.gl_total,data.gl_total]);
      sumCols(F,glTotRow,[11,12],gF,gL);
    }
    const gt=data.grand_total;
    const totRow=d.length;d.push(['TOTAL','','','','','',gt.current,gt.d1_30,gt.d31_60,gt.d61_90,gt.d91_plus,gt.gl,gt.total]);
    sumRows(F,totRow,[6,7,8,9,10,11,12],[...venTotRows,...(glTotRow!=null?[glTotRow]:[])]);
    d.push(['Reconciliation vs GL '+(data.ap_account||'202000')+' ('+fmt(data.gl_balance)+')','','','','','','','','','','','',data.recon_diff]);
    exportToExcel(d,'AP_Aging_'+data.as_of+'.xlsx',{plainCols:[5],formulas:F});
  };
  const hasAnything=data&&(data.bill_count>0||(data.gl_rows&&data.gl_rows.length>0));
  return(<div>{showUpload&&<ApAgingCutoffModal entityId={entityId} entityName={entityName} onClose={()=>setShowUpload(false)} onDone={()=>{setShowUpload(false);}}/>}<div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}><div><div style={S.h1}>A/P Aging Detail</div><div style={S.sub}>Built from GL account {data?data.ap_account:'202000'} &middot; ties to the book{data&&data.billcom_error?' · Bill.com enrich error: '+data.billcom_error:''}</div></div><div style={{display:'flex',gap:8,alignItems:'center'}}><button style={S.btnS} onClick={()=>setShowUpload(true)}>Upload aging detail</button><MemorizeBar entityId={entityId} reportType='apaging' currentConfig={{asOf}} onApply={(c)=>{if(c.asOf)setAsOf(c.asOf);}} canEdit={canEdit}/>{hasAnything&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>
    <div style={S.card}>
      <div style={{display:'flex',gap:16,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div style={{flex:'0 0 180px'}}><label style={S.label}>As of date</label><input style={{...S.inputSm,width:'100%'}} type="date" value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
        <div style={{display:'flex',alignItems:'flex-end',gap:6,flexWrap:'wrap'}}>{PRESETS.map(([k,lbl])=><button key={k} onClick={()=>setAsOf(k==='all'?today():presetRange(k).to)} style={{background:'none',border:'1px solid '+T.border,borderRadius:6,color:T.textMuted,fontSize:11,padding:'6px 10px',cursor:'pointer'}}>{lbl}</button>)}</div>
        <button style={{...S.btnP}} onClick={run} disabled={loading}>{loading?'Building from GL…':'Run Aging'}</button>
      </div>
      {err&&<div style={S.err}>{err}</div>}
    </div>
    {data&&<div className="cl-scroll" style={scrollBox()}>
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
// ─── Workpapers › GP Fees & Expenses (CLRF, quarterly) ──────────────────────
// Enter a quarter end, run the report. The server reads the four portfolio-company
// ledgers, files one copy per period under the entity's workpaper folder, and
// returns the .xlsx, which is downloaded here.
function GpFeesWorkpaper({entityId,entityName,canEdit=true}){
  const QUARTER_ENDS=['03-31','06-30','09-30','12-31'];
  const isQuarterEnd=(d)=>/^\d{4}-\d{2}-\d{2}$/.test(d)&&QUARTER_ENDS.includes(d.slice(5));
  // Default to the most recent quarter end that has already passed.
  const defaultQE=()=>{const t=today();const y=Number(t.slice(0,4));const cands=[];
    for(const yy of [y,y-1])for(const mm of QUARTER_ENDS)cands.push(yy+'-'+mm);
    const past=cands.filter(d=>d<=t).sort();return past.length?past[past.length-1]:(y-1)+'-12-31';};
  const[qe,setQe]=useState(defaultQE());
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[result,setResult]=useState(null);
  const valid=isQuarterEnd(qe);
  const run=async()=>{
    if(!valid)return;
    setBusy(true);setErr('');setResult(null);
    try{
      const r=await api.gpFeesGenerate(entityId,qe);
      if(!r)return;
      const url=URL.createObjectURL(r.blob);
      const a=document.createElement('a');a.href=url;a.download=r.filename;
      document.body.appendChild(a);a.click();document.body.removeChild(a);
      setTimeout(()=>URL.revokeObjectURL(url),4000);
      setResult(r.summary||{});
    }catch(e){ setErr(e.message||String(e)); }
    finally{ setBusy(false); }
  };
  const t=result&&result.totals?result.totals:null;
  return(<div><div style={S.card}>
    {entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}
    <div style={{fontSize:20,fontWeight:700,color:T.textBright,marginBottom:4}}>GP Fees &amp; Expenses</div>
    <div style={{fontSize:13,color:T.textMuted,marginBottom:18,maxWidth:760,lineHeight:1.5}}>
      Quarterly schedule of fees paid to the General Partner and affiliates - management fee, employee
      compensation reimbursement and development fee - built from the portfolio-company trial balances.
      A copy is filed under Workpapers &rsaquo; GP Fees &amp; Expenses by year and quarter; re-running a
      quarter replaces that quarter&rsquo;s file.
    </div>
    <div style={{display:'flex',gap:14,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>Quarter End Date</label>
        <input style={S.inputSm} type="date" value={qe} onChange={e=>{setQe(e.target.value);setErr('');setResult(null);}}/></div>
      <button style={{...S.btnP,opacity:(!valid||busy||!canEdit)?0.5:1}} disabled={!valid||busy||!canEdit} onClick={run}>
        {busy?'Running…':'Run Report'}</button>
    </div>
    {!valid&&qe&&<div style={{fontSize:12,color:T.orange,marginTop:10}}>
      Enter a quarter end date: March 31, June 30, September 30 or December 31.</div>}
    {err&&<div style={{fontSize:12,color:T.red,marginTop:12,fontWeight:600}}>{err}</div>}
    {result&&<div style={{...S.card,marginTop:18,padding:14,background:'#f3faf5'}}>
      <div style={{fontWeight:700,color:T.green,marginBottom:8}}>
        {result.quarter} report downloaded{result.replaced>0?' · replaced the previous copy':''}</div>
      {t&&<table style={{...S.table,minWidth:360,marginBottom:10}}><tbody>
        <tr><td style={S.td}>Management fee</td><td style={S.tdR}>{fmt(t.management_fee)}</td></tr>
        <tr><td style={S.td}>Employee compensation reimbursement</td><td style={S.tdR}>{fmt(t.comp_reimbursement)}</td></tr>
        <tr><td style={S.td}>Development fee</td><td style={S.tdR}>{fmt(t.development_fee)}</td></tr>
        <tr style={S.grandTotalRow}><td style={S.tdBold}>Total fees to GP/Affiliates</td>
          <td style={{...S.tdBold,textAlign:'right'}}>{fmt(t.total)}</td></tr>
      </tbody></table>}
      {result.saved_to&&<div style={{fontSize:12,color:T.textMuted}}>Filed at <strong>{result.saved_to}</strong></div>}
      {result.entities&&<div style={{fontSize:12,color:T.textMuted,marginTop:4}}>
        Portfolio companies included: {result.entities.join(' · ')}</div>}
    </div>}
  </div></div>);
}

// ─── Workpapers › Investment & Valuation (CLRF, quarterly) ──────────────────
// One run generates TWO workbooks under Workpapers › Investment & Valuation:
// the Investment workpaper (portfolio TBs, NWC/loans, waterfall, solved
// valuations under the frozen-unrealized-gain convention) and the Valuation
// workbook produced against those solved amounts so its Summary matches the
// investment workpaper's Valuations tab exactly.
function ValuationWorkpaper({entityId,entityName,canEdit=true}){
  const QUARTER_ENDS=['03-31','06-30','09-30','12-31'];
  const isQuarterEnd=(d)=>/^\d{4}-\d{2}-\d{2}$/.test(d)&&QUARTER_ENDS.includes(d.slice(5));
  const defaultQE=()=>{const t=today();const y=Number(t.slice(0,4));const cands=[];
    for(const yy of [y,y-1])for(const mm of QUARTER_ENDS)cands.push(yy+'-'+mm);
    const past=cands.filter(d=>d<=t).sort();return past.length?past[past.length-1]:(y-1)+'-12-31';};
  const[qe,setQe]=useState(defaultQE());
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[result,setResult]=useState(null);
  const valid=isQuarterEnd(qe);
  const run=async()=>{
    if(!valid)return;
    setBusy(true);setErr('');setResult(null);
    try{
      const r=await api.investmentValuationGenerate(entityId,qe);
      if(!r)return;
      setResult(r);
      const dl=(id)=>{if(!id)return;const a=document.createElement('a');
        a.href=api.downloadEntityFile(id);document.body.appendChild(a);a.click();document.body.removeChild(a);};
      if(r.investment)dl(r.investment.file_id);
      if(r.valuation)setTimeout(()=>dl(r.valuation.file_id),1200);
    }catch(e){ setErr(e.message||String(e)); }
    finally{ setBusy(false); }
  };
  const solve=result&&result.solve?result.solve:null;
  const val=result&&result.valuation?result.valuation:null;
  const inv=result&&result.investment?result.investment:null;
  return(<div><div style={S.card}>
    {entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}
    <div style={{fontSize:20,fontWeight:700,color:T.textBright,marginBottom:4}}>Investment &amp; Valuation</div>
    <div style={{fontSize:13,color:T.textMuted,marginBottom:18,maxWidth:760,lineHeight:1.5}}>
      Run Report will generate two separate workpapers &mdash; Investment and Valuation.
    </div>
    <div style={{display:'flex',gap:14,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>Quarter End Date</label>
        <input style={S.inputSm} type="date" value={qe} onChange={e=>{setQe(e.target.value);setErr('');setResult(null);}}/></div>
      <button style={{...S.btnP,opacity:(!valid||busy||!canEdit)?0.5:1}} disabled={!valid||busy||!canEdit} onClick={run}>
        {busy?'Running…':'Run Report'}</button>
    </div>
    {!valid&&qe&&<div style={{fontSize:12,color:T.orange,marginTop:10}}>
      Enter a quarter end date: March 31, June 30, September 30 or December 31.</div>}
    {err&&<div style={{fontSize:12,color:T.red,marginTop:12,fontWeight:600}}>{err}</div>}
    {result&&<div style={{...S.card,marginTop:18,padding:14,background:'#f3faf5'}}>
      <div style={{fontWeight:700,color:T.green,marginBottom:8}}>
        {result.quarter} workpapers generated &middot; filed under {result.folder}</div>
      {solve&&<table style={{...S.table,minWidth:520,marginBottom:10}}><tbody>
        <tr><td style={S.tdBold}>Investment</td><td style={{...S.tdBold,textAlign:'right'}}>Valuation</td>
          <td style={{...S.tdBold,textAlign:'right'}}>Est. Proceeds</td><td style={{...S.tdBold,textAlign:'right'}}>Book Carrying Value</td><td style={{...S.tdBold,textAlign:'right'}}>Unrealized G/(L)</td></tr>
        {[['CLIP','clip'],['Silsbee','silsbee'],['Buna','buna'],['SRN','srn']].map(([lbl,k])=>(
          <tr key={k}><td style={S.td}>{lbl}</td>
            <td style={S.tdR}>{fmt(solve[k].valuation)}</td>
            <td style={S.tdR}>{fmt(solve[k].proceeds)}</td>
            <td style={S.tdR}>{inv&&inv.J?fmt(inv.J[k]):'—'}</td>
            <td style={S.tdR}>{result.unrealized?fmt(result.unrealized[k]):'—'}</td></tr>))}
      </tbody></table>}
      {inv&&<div style={{fontSize:12,color:T.textMuted}}>Investment workpaper: <strong>{inv.folder_path}/{inv.original_name}</strong>{inv.replaced>0?' (replaced prior copy)':''}</div>}
      {val&&<div style={{fontSize:12,color:T.textMuted,marginTop:2}}>Valuation workbook: <strong>{val.folder_path}/{val.original_name}</strong>{val.replaced>0?' (replaced prior copy)':''}</div>}
    </div>}
  </div></div>);
}

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
  const[odysseyAmt,setOdysseyAmt]=useState('');
  const[savingOdyssey,setSavingOdyssey]=useState(false);
  const[alloc,setAlloc]=useState(null);

  const load=()=>{
    setErr('');
    api.getFundInvestments(entityId).then(setInvs).catch(e=>setErr(e.message));
    api.getClasses(entityId).then(setClasses).catch(e=>setErr(e.message));
  };
  useEffect(()=>{setInvs(null);setClasses(null);load();},[entityId]);

  const odysseyClass=(classes||[]).find(c=>/odyssey/i.test(c.name||''));
  const loadAlloc=()=>{api.getFundAllocation(entityId,asOf).then(a=>{setAlloc(a);}).catch(e=>setErr(e.message));};
  useEffect(()=>{if(classes!==null)loadAlloc();},[classes,asOf,entityId]);
  useEffect(()=>{
    if(odysseyClass&&alloc&&alloc.gpDetail){const row=alloc.gpDetail.find(g=>g.class_id===odysseyClass.id);if(row&&odysseyAmt==='')setOdysseyAmt(String(row.commitment_amount||''));}
  },[alloc,odysseyClass]);
  const saveOdyssey=async()=>{
    if(!odysseyClass){setErr('No investor class named "Odyssey" is tagged. Tag it as GP below first.');return;}
    setSavingOdyssey(true);setErr('');
    try{await api.setClassCommitment(entityId,odysseyClass.id,Number(odysseyAmt)||0);loadAlloc();}
    catch(e){setErr(e.message);}finally{setSavingOdyssey(false);}
  };

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

    {/* Odyssey commitment + GP/LP allocation */}
    <div style={{...S.card,marginBottom:16}}>
      <div style={S.h2}>Odyssey commitment &amp; net-loss allocation</div>
      <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>Odyssey's commitment changes monthly. Enter the current amount; the GP ownership % and the GP/LP net-loss split update from it (GP net loss = fund net loss × GP commitment ÷ total commitment).</div>
      <div style={{display:'flex',gap:10,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div>
          <label style={S.label}>Odyssey Holdings commitment ($)</label>
          <input type='number' value={odysseyAmt} onChange={e=>setOdysseyAmt(e.target.value)} placeholder='e.g. 369577' style={{...S.input,width:200}}/>
        </div>
        <button style={{...S.btnP,opacity:savingOdyssey?0.6:1}} disabled={savingOdyssey} onClick={saveOdyssey}>{savingOdyssey?'Saving…':'Save Odyssey commitment'}</button>
        {!odysseyClass&&<div style={{fontSize:12,color:T.orange||'#d08a2a',alignSelf:'center'}}>Tag an “Odyssey” class as GP below to enable this.</div>}
      </div>
      {alloc&&<div style={{marginTop:14,display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:12}}>
        {(()=>{const money=n=>n==null?'—':(n<0?'('+Math.abs(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})+')':n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}));
          const stat=(label,val,sub)=>(<div style={{border:'1px solid '+T.border,borderRadius:8,padding:'10px 12px'}}>
            <div style={{fontSize:11,color:T.textDim,textTransform:'uppercase',letterSpacing:'.04em'}}>{label}</div>
            <div style={{fontSize:18,fontWeight:700,color:T.textBright,marginTop:3}}>{val}</div>
            {sub&&<div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{sub}</div>}
          </div>);
          return(<>
            {stat('GP ownership %', alloc.hasCommitments?alloc.gpSharePct.toFixed(4)+'%':'—', 'GP '+money(alloc.gpCommitment)+' / total '+money(alloc.totalCommitment))}
            {stat('Fund net loss (YTD)', money(alloc.netLossYtd), alloc.asOf?('as of '+alloc.asOf):'set as-of date above')}
            {stat('GP net loss', money(alloc.gpNetLoss), 'allocated by ownership %')}
            {stat('LP net loss', money(alloc.lpNetLoss), 'remainder')}
          </>);})()}
      </div>}
      {alloc&&alloc.gpDetail&&alloc.gpDetail.length>0&&<div style={{marginTop:12,fontSize:12,color:T.textMuted}}>
        GP commitments: {alloc.gpDetail.map(g=>g.name+' '+(g.commitment_amount||0).toLocaleString('en-US')).join(' · ')}
      </div>}
      {alloc&&!alloc.hasCommitments&&<div style={{marginTop:12,fontSize:12,color:T.orange||'#d08a2a'}}>No commitments loaded yet, so net loss falls entirely to LP. Once commitments are loaded, the split activates automatically.</div>}
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
function FinancialStatements({entityId,entityName,canEdit=true,isDevEntity=false}){
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
      const out=await api.financialStatementsGenerate(entityId,asOf,period,execFile,isDevEntity?reqFile:null);
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
    <div style={{color:T.textMuted,marginBottom:16,fontSize:13,maxWidth:820}}>Generates a GL-derived statement package for {entityName||'this entity'} — Balance Sheet, Statements of Operations, Statement of Cash Flows, and Statement of Changes in Members' Equity — as of a date, then merges it into a single PDF with your uploaded executive summary{isDevEntity?' and requisition report. The requisition report\u2019s Current & Prior Invoice Log pages are removed automatically.':'.'}</div>
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
        {isDevEntity&&<div>
          <label style={S.label}>Requisition report (PDF or Excel &mdash; Invoice Log pages removed automatically)</label>
          <input type="file" accept=".pdf,.xlsx,.xls" disabled={!canEdit} onChange={e=>setReqFile(e.target.files[0]||null)} style={{fontSize:13}}/>
          {reqFile&&<span style={{marginLeft:10,color:T.textMuted,fontSize:12}}>{reqFile.name}</span>}
        </div>}
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
    const F=[];const first=d.length;
    data.investors.forEach(i=>d.push([i.investor,i.investor_code||'',i.commitment_amount,i.called_amount,i.uncalled_amount,i.pct_called,i.ownership_pct,i.commit_date||'',i.notes||'']));
    const last=d.length-1;
    for(let r=first;r<=last;r++){const rr=r+1;F.push({r,c:4,f:'C'+rr+'-D'+rr});}// Uncalled = Commitment − Called
    const t=data.totals;const tr=d.length;d.push(['Total','',t.commitment_amount,t.called_amount,t.uncalled_amount,'','','','']);
    sumCols(F,tr,[2,3,4],first,last);
    exportToExcel(d,'Investor_Commitments.xlsx',{formulas:F});};
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
     <div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
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
// A statement ending balance that was TYPED as 0.00 is a real balance, not a
// missing one: an account that sat at zero all month still has to be reconciled
// and signed off. So the ending-balance box is parsed to null when it is blank
// (or unparseable) and to a number otherwise — blank is what blocks finalizing,
// never the value zero. Tolerant of "$", thousands commas and (parentheses) for
// negatives, which is how a balance arrives when pasted off a statement.
function parseStmtBalance(s){
  const t=String(s??'').trim(); if(!t) return null;
  const neg=/^\(.*\)$/.test(t);
  const n=parseFloat(t.replace(/[()$,\s]/g,''));
  return Number.isFinite(n)?(neg?-Math.abs(n):n):null;
}
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
  const uncl=txns.filter(t=>!cleared[t.key]);const bookBal=txns.reduce((s,t)=>s+t.amount,0);const stmtVal=parseStmtBalance(stmtBal);const hasStmt=stmtVal!==null;const stmtNum=hasStmt?stmtVal:0;
  const outDep=uncl.filter(t=>!checked[t.key]&&t.amount>0).reduce((s,t)=>s+t.amount,0);const outPay=uncl.filter(t=>!checked[t.key]&&t.amount<0).reduce((s,t)=>s+t.amount,0);
  const diff=bookBal-(stmtNum+outDep+outPay);const isRec=hasStmt&&Math.abs(diff)<0.005;
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
      <div style={{display:'flex',justifyContent:'flex-end',marginTop:16}}><button style={{...S.btnP,padding:'10px 28px',fontSize:14,opacity:isRec?1:.5,cursor:isRec?'pointer':'not-allowed'}} onClick={finalize}>{isRec?'Finalize Reconciliation':(hasStmt?'Difference must be $0.00':'Enter the statement ending balance')}</button></div></>}
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
// ═══ Workpapers › Insurance Allocation (Banyan Residential) ═══
// Upload the carrier billing invoice + the consolidated billing report; the
// server computes each entity's employer/employee split, files the workpaper
// under Workpapers, and returns the .xlsx (auto-downloaded) plus a summary.
function InsuranceAllocationWorkpaper({entityId,entityName,canEdit=true}){
  const[invoice,setInvoice]=useState(null);
  const[consolidated,setConsolidated]=useState(null);
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');
  const[result,setResult]=useState(null);
  const run=async()=>{
    if(!invoice||!consolidated)return;
    setBusy(true);setErr('');setResult(null);
    try{
      const r=await api.insuranceAllocationGenerate(entityId,invoice,consolidated);
      if(!r)return;
      const url=URL.createObjectURL(r.blob);
      const a=document.createElement('a');a.href=url;a.download=r.filename;
      document.body.appendChild(a);a.click();document.body.removeChild(a);
      setTimeout(()=>URL.revokeObjectURL(url),4000);
      setResult(r.summary||{});
    }catch(e){ setErr(e.message||String(e)); }
    finally{ setBusy(false); }
  };
  const s=result;
  return(<div><div style={S.card}>
    {entityName&&<div style={{fontSize:14,fontWeight:600,color:T.textMuted,marginBottom:4}}>{entityName}</div>}
    <div style={{fontSize:20,fontWeight:700,color:T.textBright,marginBottom:4}}>Insurance Allocation</div>
    <div style={{fontSize:13,color:T.textMuted,marginBottom:18,maxWidth:780,lineHeight:1.5}}>
      Allocates the monthly health-insurance premium across the four commonly-owned entities
      (County Line Rail Operations, Banyan Residential, Sabine River &amp; Northern, TurnKey Rail).
      Upload the carrier billing invoice and the consolidated billing report; a copy is filed under
      Workpapers &rsaquo; Insurance Allocation by year, and the workbook downloads to your computer.
    </div>
    <div style={{display:'flex',gap:18,alignItems:'flex-end',flexWrap:'wrap'}}>
      <div><label style={S.label}>Health insurance billing invoice (.xlsx)</label>
        <input style={S.inputSm} type="file" accept=".xlsx,.xls" onChange={e=>{setInvoice(e.target.files[0]||null);setErr('');setResult(null);}}/></div>
      <div><label style={S.label}>Consolidated billing report (.xlsx)</label>
        <input style={S.inputSm} type="file" accept=".xlsx,.xls" onChange={e=>{setConsolidated(e.target.files[0]||null);setErr('');setResult(null);}}/></div>
      <button style={{...S.btnP,opacity:(!invoice||!consolidated||busy||!canEdit)?0.5:1}} disabled={!invoice||!consolidated||busy||!canEdit} onClick={run}>
        {busy?'Running…':'Generate Allocation'}</button>
    </div>
    {!canEdit&&<div style={{fontSize:12,color:T.textMuted,marginTop:10}}>This workpaper is read-only for your account.</div>}
    {err&&<div style={{fontSize:12,color:T.red,marginTop:12,fontWeight:600}}>{err}</div>}
    {s&&<div style={{...S.card,marginTop:18,padding:16,background:'#f3faf5'}}>
      <div style={{fontWeight:700,color:s.reconciled?T.green:T.orange,marginBottom:10}}>
        Allocation generated{s.period?(' · '+s.period):''}{s.reconciled?' · balanced':' · review unmatched subscribers'}</div>
      <table style={{...S.table,minWidth:580,marginBottom:8}}><tbody>
        <tr><td style={S.tdBold}>Entity</td><td style={{...S.tdBold,textAlign:'right'}}>Premium</td>
          <td style={{...S.tdBold,textAlign:'right'}}>Employer</td><td style={{...S.tdBold,textAlign:'right'}}>Employee</td>
          <td style={{...S.tdBold,textAlign:'right'}}>Eligibility</td></tr>
        {(s.entities||[]).map(e=>(<tr key={e.code}><td style={S.td}>{e.name} &middot; {e.code}</td>
          <td style={S.tdR}>{fmt((e.premium||0)+(e.eligibility||0))}</td>
          <td style={S.tdR}>{fmt((e.employer||0)+(e.eligibility||0))}</td>
          <td style={S.tdR}>{fmt(e.employee)}</td>
          <td style={S.tdR}>{e.eligibility?fmt(e.eligibility):'—'}</td></tr>))}
        <tr style={S.grandTotalRow}><td style={S.tdBold}>Total billed</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(s.totalBilled)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>{fmt(s.employerTotal)}</td><td style={{...S.tdBold,textAlign:'right'}}>{fmt(s.employeeTotal)}</td>
          <td style={{...S.tdBold,textAlign:'right'}}>{s.eligibilityTotal?fmt(s.eligibilityTotal):'—'}</td></tr>
      </tbody></table>
      {s.eligibilityTotal?<div style={{fontSize:11,color:T.textMuted,marginBottom:10}}>Eligibility charges are billed 100% to each entity&rsquo;s employer and are included in the Premium and Employer columns above.</div>:null}
      {(s.flags||[]).length>0&&<div style={{marginBottom:8}}>
        <div style={{fontSize:12,fontWeight:700,color:T.textMuted,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:4}}>Review flags</div>
        {(s.flags||[]).map((f,i)=>(<div key={i} style={{fontSize:12,color:T.textMuted,marginBottom:2}}>
          {f.type==='reclass'?('Entity reclassified — '+f.name+': billed '+f.from+', allocated to '+f.to+'.'):
           f.type==='employerPaid'?('Employer-paid in full — '+f.name+': employee share $0.00 ('+f.entity+').'):
           f.type==='eligibility'?('Eligibility change — '+f.name+': '+fmt(f.amount)+' booked 100% employer → '+(f.entity||'unmatched')+'.'):''}
        </div>))}
      </div>}
      {(s.unmatched||[]).length>0&&<div style={{fontSize:12,color:T.orange,marginBottom:6}}>
        Unmatched (no entity — review): {(s.unmatched||[]).join(', ')}</div>}
      {s.savedFolder&&<div style={{fontSize:12,color:T.textMuted}}>Filed at <strong>{s.savedFolder}/{s.savedName}</strong> &middot; also downloaded to your computer.</div>}
    </div>}
  </div></div>);
}

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
  // Vendor(normalized) -> {code,name,budget} and code -> {name,budget}, parsed
  // from the uploaded workbook's invoice logs. Used to auto-code invoices for
  // entities with no CloudLedger coding history (e.g. Braker, whose reports were
  // built in Excel), and to populate the Budget Code column (Braker-only).
  const[wbVendorMap,setWbVendorMap]=useState({});
  const[wbBudgetMap,setWbBudgetMap]=useState({});
  // Normalize a vendor string for matching: lowercase, strip punctuation and
  // common company suffixes, collapse whitespace. Mirrors the server's normVendor
  // closely enough for the workbook-seeded fallback (exact/looser matches only).
  const normVend=(s)=>String(s==null?'':s).toLowerCase()
    .replace(/\(reclass\)/g,' ').replace(/[.,]/g,' ')
    .replace(/\b(llc|inc|lp|llp|pllc|pc|ltd|co|corp|company|incorporated)\b/g,' ')
    .replace(/\s+/g,' ').trim();
  const parseWorkbookCoaMap=async(file)=>{
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      // Scan BOTH invoice logs. The Prior log is the richest source (it carries
      // every historical vendor->code->name->budget pairing); the Current log adds
      // this-period vendors. Later sheets don't overwrite an existing mapping, so
      // the Prior log (parsed first) wins on conflicts.
      const sheetNames=['Prior Invoice Log','Current Invoice Log'];
      const m={};                 // code -> cost_code_name (existing behavior)
      // Frequency tallies. A vendor or a cost code can carry more than one
      // coding across history; we pick the MOST COMMON pairing (the mode) rather
      // than the top-most row, so e.g. Strategic HFC's code 11634 resolves to its
      // 9-of-10 'HFC Construction Monitoring Fee' budget, not the single stray
      // 'Sales Tax Exemption' row. Ties fall back to the first seen.
      const vTally={};            // vendorKey -> Map('code||name||budget' -> {count,code,name,budget,order})
      const bTally={};            // code -> Map(budget -> {count,order}); also firstName per code
      const bName={};             // code -> first non-empty cost_code_name
      let ord=0,sawAny=false;
      const bumpV=(k,code,name,budget)=>{const key=code+'||'+name+'||'+budget;let t=vTally[k];if(!t){t=new Map();vTally[k]=t;}let e=t.get(key);if(!e){e={count:0,code,name,budget,order:ord++};t.set(key,e);}e.count++;};
      const bumpB=(code,budget)=>{let t=bTally[code];if(!t){t=new Map();bTally[code]=t;}let e=t.get(budget);if(!e){e={count:0,order:ord++};t.set(budget,e);}e.count++;};
      for(const sn of sheetNames){
        const ws=wb.Sheets[sn];
        if(!ws)continue;
        sawAny=true;
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,blankrows:false});
        // Detect columns by HEADER text. Layouts differ (SRN name in col F,
        // HP/Braker name in col D with Bill # in col F), and Braker carries an
        // extra "Budget Code" column. Fall back to SRN positions if not found.
        let codeIdx=2,nameIdx=5,budgetIdx=-1,vendorIdx=-1,hdrRow=-1;
        for(let i=0;i<Math.min(rows.length,8);i++){
          const cells=(rows[i]||[]).map(c=>String(c==null?'':c).toLowerCase().replace(/\s+/g,' ').trim());
          const ci=cells.findIndex(t=>/cost code\s*#|cost code\s*(number|no)\b|^cost code$/.test(t));
          const ni=cells.findIndex(t=>/cost code name/.test(t));
          if(ci>=0&&ni>=0){
            codeIdx=ci;nameIdx=ni;hdrRow=i;
            budgetIdx=cells.findIndex(t=>/budget\s*code/.test(t));  // Braker-only; -1 elsewhere
            vendorIdx=cells.findIndex(t=>/vendor|payee/.test(t));
            break;
          }
        }
        for(let i=(hdrRow>=0?hdrRow+1:0);i<rows.length;i++){
          const row=rows[i];if(!row)continue;
          const code=row[codeIdx]!=null?String(row[codeIdx]).trim():'';
          const name=row[nameIdx]!=null?String(row[nameIdx]).trim():'';
          if(!code||!/\d/.test(code))continue;          // skip headers / subtotal rows (no numeric code)
          if(/total/i.test(name))continue;               // skip "X Total" subtotal label rows
          const budget=budgetIdx>=0&&row[budgetIdx]!=null?String(row[budgetIdx]).trim():'';
          const vendor=vendorIdx>=0&&row[vendorIdx]!=null?String(row[vendorIdx]).trim():'';
          if(name&&!m[code])m[code]=name;                // first (top-most) name for a code wins
          if(name&&!bName[code])bName[code]=name;
          if(budget)bumpB(code,budget);                  // tally budgets per code
          if(vendor){const vk=normVend(vendor);if(vk)bumpV(vk,code,name,budget);}  // tally codings per vendor
        }
      }
      // Also fold in the full Chart of Accounts (the "COA - Intacct" tab):
      // Account no. -> Title. Scanned AFTER the invoice logs so a code already
      // used in the logs keeps its report label; this only FILLS IN codes not yet
      // invoiced. That's what lets you CORRECT a coding to a brand-new (but valid)
      // cost code and still have the Cost Code Name auto-fill.
      for(const coaSheet of Object.keys(wb.Sheets)){
        if(!/coa|chart of accounts/i.test(coaSheet))continue;
        const ws=wb.Sheets[coaSheet];if(!ws)continue;
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,blankrows:false});
        let acctIdx=0,titleIdx=1,hdr=-1;
        for(let i=0;i<Math.min(rows.length,12);i++){
          const cells=(rows[i]||[]).map(c=>String(c==null?'':c).toLowerCase().replace(/\s+/g,' ').trim());
          const ai=cells.findIndex(t=>/account\s*(no|number|#)|^acct\b/.test(t));
          const ti=cells.findIndex(t=>/^title$|account\s*name|description/.test(t));
          if(ai>=0&&ti>=0){acctIdx=ai;titleIdx=ti;hdr=i;break;}
        }
        for(let i=(hdr>=0?hdr+1:0);i<rows.length;i++){
          const row=rows[i];if(!row)continue;
          const code=row[acctIdx]!=null?String(row[acctIdx]).trim():'';
          const name=row[titleIdx]!=null?String(row[titleIdx]).trim():'';
          if(!code||!/^\d+$/.test(code)||!name)continue;   // account numbers only
          sawAny=true;
          if(!m[code])m[code]=name;                        // logs win; COA fills the gaps
          if(!bName[code])bName[code]=name;
        }
      }
      if(!sawAny){setWbCoaMap({});setWbVendorMap({});setWbBudgetMap({});return;}
      // Finalize: pick the modal budget per code, and the modal coding per vendor.
      const pickMode=(t)=>{let best=null;for(const[k,e]of t){if(!best||e.count>best.count||(e.count===best.count&&e.order<best.order))best={key:k,...e};}return best;};
      const bmap={};              // code -> {name,budget}
      for(const code of Object.keys(bTally)){const best=pickMode(bTally[code]);bmap[code]={name:bName[code]||'',budget:best?best.key:''};}
      const vmap={};              // normalized vendor -> {code,name,budget}
      for(const vk of Object.keys(vTally)){const best=pickMode(vTally[vk]);if(best)vmap[vk]={code:best.code||'',name:best.name||'',budget:best.budget||''};}
      setWbCoaMap(m);setWbVendorMap(vmap);setWbBudgetMap(bmap);
    }catch{setWbCoaMap({});setWbVendorMap({});setWbBudgetMap({});}
  };

  // Read each uploaded invoice with Claude and append an editable card.
  const onRfInvoices=async(e)=>{const files=[...e.target.files];e.target.value='';if(!files.length)return;
    setRfReadErr('');setRfReading(n=>n+files.length);
    for(const f of files){
      try{const r=await api.readRequisitionInvoice(entityId,f);
        // Server prediction relies on CloudLedger coding history. Entities whose
        // reports were built in Excel (e.g. Braker) have no history, so cost_code
        // comes back blank. Fall back to the uploaded workbook: match the invoice
        // vendor against the workbook's vendor->coding map to seed cost code, name,
        // and Budget Code. Every field remains editable in the card.
        let cc=r.cost_code||'';
        let ccName=(cc&&wbCoaMap[String(cc).trim()])||r.cost_code_name||'';
        let budget='';
        if(!cc&&r.vendor){const seed=wbVendorMap[normVend(r.vendor)];if(seed){cc=seed.code||'';if(!ccName)ccName=seed.name||'';budget=seed.budget||'';}}
        // Budget Code follows the cost code when the workbook has a Budget Code
        // column (Braker-only). The vendor seed's budget is vendor-specific and
        // wins; only fall back to the per-code modal budget when the vendor seed
        // gave none (e.g. the code was typed by hand). Empty without the column.
        if(!budget&&cc){const b=wbBudgetMap[String(cc).trim()];if(b&&b.budget)budget=b.budget;}
        setRfCards(cards=>[...cards,{
          _id:Date.now()+'-'+Math.random().toString(36).slice(2,7),
          filename:r.filename||f.name,
          cost_code:cc,
          // Prefer the name from THIS workbook's cost-code catalog (authoritative
          // for the requisition) over the server prediction, whose learned history
          // may carry a mis-columned name for templates like HP/Braker.
          cost_code_name:ccName,
          budget_code:budget,
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
    // Track when the user hand-edits the Name or Budget Code, so a later Cost
    // Code change won't clobber a value they typed themselves.
    if(field==='cost_code_name')next._nameEdited=true;
    if(field==='budget_code')next._budgetEdited=true;
    // When the Cost Code changes, follow it: fill the Cost Code Name and Budget
    // Code from the catalog (uploaded report's map wins; server catalog is the
    // fallback) WHENEVER the code resolves to a known value — unless the user
    // manually edited that field. An empty lookup (e.g. mid-typing an
    // incomplete code) never clears an existing value, so partial keystrokes
    // don't wipe the field and don't block the final code from filling it in.
    if(field==='cost_code'){
      const k=String(val).trim();
      const newName=wbCoaMap[k]||(coaMap[k]&&coaMap[k].cost_code_name)||'';
      if(newName&&!next._nameEdited)next.cost_code_name=newName;
      const b=wbBudgetMap[k];
      const newBudget=b&&b.budget?b.budget:'';
      if(newBudget&&!next._budgetEdited)next.budget_code=newBudget;
    }
    return next;
  }));
  // Back-fill coding on existing cards whenever the code catalog becomes
  // available or changes — the uploaded report's map (wbCoaMap/wbBudgetMap) or
  // the server catalog (coaMap). This covers coding an invoice BEFORE the prior
  // report was loaded: as soon as the catalog arrives, any card with a cost code
  // but a blank Cost Code Name (or blank Budget Code) is filled in. Only blanks
  // are touched, so manual entries and existing codings are never overwritten.
  // Universal — runs for every entity in the requisition module.
  useEffect(()=>{
    setRfCards(cards=>{
      let changed=false;
      const next=cards.map(c=>{
        if(!c.cost_code)return c;
        const k=String(c.cost_code).trim();
        const patch={};
        if(!c.cost_code_name){
          const nm=wbCoaMap[k]||(coaMap[k]&&coaMap[k].cost_code_name)||'';
          if(nm)patch.cost_code_name=nm;
        }
        if(!c.budget_code){
          const b=wbBudgetMap[k];
          if(b&&b.budget)patch.budget_code=b.budget;
        }
        if(Object.keys(patch).length){changed=true;return{...c,...patch};}
        return c;
      });
      return changed?next:cards;
    });
  },[wbCoaMap,wbBudgetMap,coaMap]);
  const removeCard=id=>setRfCards(cards=>cards.filter(c=>c._id!==id));

  const runRollForward=async(force=false)=>{
    if(!rfFile){setRfErr('Upload the prior requisition workbook (.xlsx) first.');return;}
    const newCurrent=rfCards.map(c=>{
      const amount=c.amount!==''?parseFloat(String(c.amount).replace(/[$,]/g,'')):NaN;
      return {code:c.cost_code||undefined,name:c.cost_code_name||undefined,budgetcode:c.budget_code||undefined,vendor:c.vendor||undefined,bill:c.bill||undefined,date:c.date||undefined,...(Number.isFinite(amount)?{amount}:{})};
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
              {Object.values(wbBudgetMap).some(b=>b&&b.budget)&&(
              <div style={{flex:'1 1 130px'}}><label style={S.label}>Budget Code</label><input style={S.input} value={c.budget_code||''} onChange={e=>updateCard(c._id,'budget_code',e.target.value)}/></div>
              )}
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
        <div style={{...S.col,flex:2}}><label style={S.label}>Entity Type</label><select style={S.input} value={newType} onChange={e=>setNewType(e.target.value)}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option><option value="operating">Operating</option></select></div></div>
      {err&&<div style={S.err}>{err}</div>}
      <div style={{fontSize:11,color:T.textMuted,marginBottom:10}}>A default chart of accounts will be created. You can replace it by importing a trial balance from the entity row. Development-project entities unlock the Requisitions coding tools.</div>
      <button style={S.btnP} onClick={async()=>{if(!name.trim()){setErr('Name required');return;}try{await api.createEntity(name.trim(),newType,newDisplayId.trim());setName('');setNewType('accounting');setNewDisplayId('');setShowAdd(false);setErr('');refresh();}catch(e){setErr(e.message);}}}>Create Entity</button></div>}
    {bulk&&(()=>{const TYPE_ALIASES={accounting:'accounting',acct:'accounting',acc:'accounting',development:'development','development project':'development',dev:'development',devproject:'development',shell:'shell',operating:'operating',op:'operating'};
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
        <select style={{...S.inputSm,width:'auto'}} value={bulkType} onChange={e=>setBulkType(e.target.value)}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option><option value="operating">Operating</option></select>
        {rows.length>0&&<span style={{fontSize:11,color:T.textMuted}}>{rows.length} entit{rows.length===1?'y':'ies'}: {['accounting','development','shell','operating'].map(t=>({t,n:rows.filter(r=>r.type===t).length})).filter(x=>x.n>0).map(x=>x.n+' '+x.t).join(', ')}</span>}
      </div>
      {err&&<div style={S.err}>{err}</div>}<button style={{...S.btnP,marginTop:10}} disabled={bulkBusy} onClick={async()=>{if(!rows.length){setErr('None');return;}setBulkBusy(true);setErr('');try{for(const r of rows)await api.createEntity(r.name,r.type);setBulkText('');setBulk(false);refresh();}catch(e){setErr(e.message);}finally{setBulkBusy(false);}}}>{bulkBusy?'Importing...':'Import'}</button></div>);})()}
    <div className="cl-scroll" style={scrollBox()}><table style={{...S.table,minWidth:1180}}><thead><tr><th style={{...S.th,minWidth:240}}>Entity</th><th style={{...S.th,width:760,minWidth:760}}>Actions</th></tr></thead>
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
          <select style={{...S.inputSm,padding:'5px 8px',fontSize:11,flexShrink:0,width:'auto'}} disabled={typeBusy===e.id} title="Entity type" value={e.entity_type||'accounting'} onChange={async(ev)=>{const next=ev.target.value;if(next===e.entity_type)return;if(!confirm('Set "'+e.name+'" to '+({accounting:'Accounting',development:'Development Project',shell:'Shell'}[next]||next)+'?'))return;setTypeBusy(e.id);try{await api.updateEntity(e.id,{entity_type:next});await refresh();}catch(ex){alert(ex.message);}finally{setTypeBusy(null);}}}><option value="accounting">Accounting</option><option value="development">Development Project</option><option value="shell">Shell</option><option value="operating">Operating</option></select>
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
  const[accessEffective,setAccessEffective]=useState(null); // null=all, else union of individual+group minus exclusions
  const[accessExclusions,setAccessExclusions]=useState([]); // per-user entity ids to subtract (negative overrides)
  const[accessLevels,setAccessLevels]=useState({}); // entity_id -> 'full' | 'view' for individual grants
  const openAccess=async(u)=>{
    setAccessUser(u);setAccessErr('');setAccessSaving(false);setAccessGroups([]);setAccessEffective(null);setAccessExclusions([]);setAccessLevels({});
    try{
      const[ents,acc]=await Promise.all([api.getEntities(),api.getUserEntityAccess(u.id)]);
      setAccessAllEntities(ents);
      setAccessEntities(acc.entity_ids||[]);
      setAccessGroups(acc.groups||[]);
      setAccessExclusions(acc.exclusions||[]);
      setAccessLevels(acc.levels||{});
      setAccessEffective(acc.effective===undefined?null:acc.effective);
    }catch(e){setAccessErr(e.message);}
  };
  const toggleAccessEntity=(id)=>setAccessEntities(prev=>prev.includes(id)?prev.filter(x=>x!==id):[...prev,id]);
  // Toggle an entity granted via group membership: unchecking adds it to the
  // per-user exclusion list (subtracted from effective access); rechecking removes it.
  const toggleAccessExclusion=(id)=>setAccessExclusions(prev=>prev.includes(id)?prev.filter(x=>x!==id):[...prev,id]);
  const setAccessLevel=(id,lvl)=>setAccessLevels(prev=>({...prev,[id]:lvl}));
  const saveAccess=async()=>{
    setAccessSaving(true);setAccessErr('');
    try{await api.setUserEntityAccess(accessUser.id,accessEntities,accessExclusions,accessLevels);setAccessUser(null);}
    catch(e){setAccessErr(e.message);}
    setAccessSaving(false);
  };
  useEffect(()=>{loadUsers();},[loadUsers]);
  // ── User groups (e.g. CLA): bundle members + grant entity access at once ──
  const[groups,setGroups]=useState([]);
  const[groupModal,setGroupModal]=useState(null);
  const[gMembers,setGMembers]=useState([]);const[gEntities,setGEntities]=useState([]);const[gAllEntities,setGAllEntities]=useState([]);const[gLevels,setGLevels]=useState({});
  const[gSaving,setGSaving]=useState(false);const[gErr,setGErr]=useState('');
  const loadGroups=useCallback(()=>{api.getGroups().then(setGroups).catch(()=>{});},[]);
  useEffect(()=>{loadGroups();},[loadGroups]);
  const openGroup=async(g)=>{setGroupModal(g);setGErr('');setGSaving(false);try{const[detail,ents]=await Promise.all([api.getGroup(g.id),api.getEntities()]);setGMembers(detail.member_ids||[]);setGEntities(detail.entity_ids||[]);setGLevels(detail.levels||{});setGAllEntities(ents);}catch(e){setGErr(e.message);}};
  const toggleGMember=(id)=>setGMembers(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const toggleGEntity=(id)=>setGEntities(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const setGLevel=(id,lvl)=>setGLevels(p=>({...p,[id]:lvl}));
  const saveGroup=async()=>{setGSaving(true);setGErr('');try{await api.setGroupMembers(groupModal.id,gMembers);await api.setGroupEntities(groupModal.id,gEntities,gLevels);setGroupModal(null);loadGroups();}catch(e){setGErr(e.message);}setGSaving(false);};
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
          const excluded=accessExclusions.includes(e.id);
          // A group-granted entity is checked unless it's been explicitly excluded.
          // Unchecking a group entity toggles the exclusion; an individual grant
          // toggles the normal grant. Excluding only applies to group entities.
          const checked=!excluded&&(inGroup||accessEntities.includes(e.id));
          const onToggle=()=>{ if(inGroup) toggleAccessExclusion(e.id); else toggleAccessEntity(e.id); };
          return (
          <label key={e.id} title={inGroup?(excluded?'Excluded for this user (overrides '+viaGroups.join(', ')+')':'Granted via '+viaGroups.join(', ')+' — uncheck to exclude for this user only'):''} style={{display:'flex',alignItems:'center',padding:'8px 12px',borderBottom:'1px solid '+T.border,cursor:'pointer',gap:10}}>
            <input type="checkbox" checked={checked} onChange={onToggle}/>
            <span style={{color:excluded?T.textMuted:T.textBright,fontSize:13,textDecoration:excluded?'line-through':'none'}}>{e.name}</span>
            {e.code&&<span style={{color:T.textMuted,fontSize:11,fontFamily:'monospace'}}>{e.code}</span>}
            {inGroup&&!excluded&&<span style={{marginLeft:'auto',fontSize:10,color:T.accent,background:T.accent+'18',padding:'2px 6px',borderRadius:4,whiteSpace:'nowrap'}}>via {viaGroups.join(', ')}</span>}
            {inGroup&&excluded&&<span style={{marginLeft:'auto',fontSize:10,color:T.danger||'#c0392b',background:(T.danger||'#c0392b')+'18',padding:'2px 6px',borderRadius:4,whiteSpace:'nowrap'}}>excluded</span>}
            {accessEntities.includes(e.id)&&!inGroup&&<span style={{marginLeft:'auto',display:'flex',gap:4}} onClick={ev=>ev.preventDefault()}>
              {[['full','Full'],['view','View only']].map(([lv,lbl])=><button key={lv} type="button" onClick={ev=>{ev.preventDefault();ev.stopPropagation();setAccessLevel(e.id,lv);}} style={{fontSize:10,padding:'2px 8px',borderRadius:4,border:'1px solid '+T.border,cursor:'pointer',background:((accessLevels[e.id]||'full')===lv)?T.accent:'transparent',color:((accessLevels[e.id]||'full')===lv)?'#fff':T.textMuted}}>{lbl}</button>)}
            </span>}
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
                {gEntities.includes(e.id)&&<span style={{marginLeft:'auto',display:'flex',gap:4}} onClick={ev=>ev.preventDefault()}>
                  {[['full','Full'],['view','View']].map(([lv,lbl])=><button key={lv} type="button" onClick={ev=>{ev.preventDefault();ev.stopPropagation();setGLevel(e.id,lv);}} style={{fontSize:9,padding:'1px 6px',borderRadius:4,border:'1px solid '+T.border,cursor:'pointer',background:((gLevels[e.id]||'full')===lv)?T.accent:'transparent',color:((gLevels[e.id]||'full')===lv)?'#fff':T.textMuted}}>{lbl}</button>)}
                </span>}
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

// ═══════════════════════════ Intercompany ═══════════════════════════
// Two pages under the Intercompany nav section.
//
//   IntercompanyReconciliation — Due from / Due to (transactional) and
//     Investment / Contributed capital (structural), run against a saved group
//     as of a date.
//   IntercompanyMapping — the setup page. Tells the reconciliation which GL
//     account faces which entity, and defines the groups.
//
// The rule the whole section exists to protect: a counterparty OUTSIDE the
// selected group is tagged "no elim" and never eliminated. It is shown in its
// own panel so the amount is visible rather than quietly netted away.

const IC_TYPE_LABEL={due_from:'Due from',due_to:'Due to',investment:'Investment',contributed_capital:'Contributed capital'};

function IcBadge({kind,children}){
  const C={ok:[T.green,T.greenDim],warn:[T.orange,T.orangeDim],bad:[T.red,T.redDim],info:[T.accent,T.accentDim],mute:[T.textMuted,T.bgElevated]}[kind]||[T.textMuted,T.bgElevated];
  return<span style={{display:'inline-block',padding:'2px 9px',borderRadius:20,fontSize:10,fontWeight:700,color:C[0],background:C[1],whiteSpace:'nowrap'}}>{children}</span>;
}
function IcTile({label,value,kind}){
  const c={ok:T.green,warn:T.orange,bad:T.red,info:T.accent}[kind]||T.textBright;
  return<div style={{flex:1,minWidth:150,background:T.bgCard,border:'1px solid '+T.border,borderRadius:T.radiusSm,padding:'12px 16px'}}>
    <div style={{fontSize:10,color:T.textMuted,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.06em',marginBottom:4}}>{label}</div>
    <div style={{fontSize:19,fontWeight:700,color:c,fontVariantNumeric:'tabular-nums'}}>{value}</div></div>;
}
function IcEmpty({children}){return<div style={{...S.card,textAlign:'center',padding:40,color:T.textDim}}>{children}</div>;}


// ═══════════════════ IC Reconciliation — one entity ═══════════════════
// Pick an entity; every counterparty relationship it has is listed with the
// account CODE and NAME on BOTH sides. A netted difference tells you the two
// ledgers disagree — the account numbers tell you which two entries to open,
// which is the whole point of running this.

const IC_STATUS={
  matched:['ok','Matched'],
  mismatch:['bad','Mismatch'],
  one_sided:['warn','Counterparty not mapped'],
  no_ledger:['mute','No counterparty ledger'],
  settled:['mute','Settled'],
  self:['bad','Points at itself'],
};
function IcStatus({status}){const s=IC_STATUS[status]||['mute',status];return<IcBadge kind={s[0]}>{s[1]}</IcBadge>;}

// Two ledgers rarely use the same number of accounts for one relationship, so
// the sides are laid out row by row and the shorter one is padded. Keeping them
// on the same row makes the comparison readable without inventing a pairing
// that does not exist.
function zipLegs(a,b){
  const theirs=b||[],ours=a||[];
  const used=new Set(),pairs=[];
  // A mapping that names the counterparty account puts the two entries on the
  // SAME line. Without a pin the sides are still zipped by order and padded —
  // that stays deliberate, because inventing a pairing is worse than admitting
  // there isn't one.
  for(const l of ours){
    let mate=null;
    if(l&&l.counterparty_account_code){
      const j=theirs.findIndex((t,i)=>!used.has(i)&&t&&String(t.account_code)===String(l.counterparty_account_code));
      if(j>=0){used.add(j);mate=theirs[j];}
    }
    pairs.push([l||null,mate]);
  }
  // Whatever they hold that no pin claimed drops into the free slots, in order.
  const rest=theirs.filter((_,i)=>!used.has(i));
  let k=0;
  for(const p of pairs){if(!p[1]&&k<rest.length)p[1]=rest[k++];}
  while(k<rest.length)pairs.push([null,rest[k++]]);
  return pairs.length?pairs:[[null,null]];
}
const legCells=(l,align)=>l
  ?<><td style={{...S.td,color:T.textBright,fontWeight:600,width:88}}>{l.account_code}</td>
     <td style={{...S.td,whiteSpace:'normal',color:T.textMuted}}>{l.account_name}
       {l.ic_type&&<span style={{marginLeft:6,fontSize:10,color:T.textDim}}>{IC_TYPE_LABEL[l.ic_type]}</span>}</td>
     <td style={{...S.tdR,width:130}}>{fmt(l.amount)}</td></>
  :<><td style={S.td}></td><td style={S.td}></td><td style={S.tdR}></td></>;

// ── Shared intercompany pieces ──
// The counterparty picker and the one-row "map this account" form are used by
// BOTH pages. An account that needs mapping is almost always noticed on the
// reconciliation — that is the page that shows you something is missing — so
// the fix has to be available there, not only on the setup page.

// A counterparty is a CloudLedger entity, a registered company with no ledger
// here, or an outside party. Encoded 'e<id>' / 'n<id>' / '__ext' so a single
// <select> carries all three, plus '__new' to register a company inline.
const icCpValue=(entId,nodeId,isExt)=>isExt?'__ext':(nodeId?'n'+nodeId:(entId?'e'+entId:''));

function IcCpSelect({entityOpts,companies,entId,nodeId,isExt,onPick,seedName,onCompanyAdded,onMsg,onErr,style}){
  const offLedger=(companies||[]).filter(c=>c.entity_id==null);
  return<select style={{...S.selectSm,minWidth:200,...(style||{})}} value={icCpValue(entId,nodeId,isExt)} onChange={async e=>{
    const v=e.target.value;
    if(v==='__ext'){onPick({entity_id:null,node_id:null,is_external:true});return;}
    if(v==='__new'){
      const name=prompt('Register a company that has no ledger in CloudLedger:',seedName||'');
      if(!name||!name.trim())return;
      try{const r=await api.createIcCompany({name:name.trim()});
        if(onCompanyAdded)await onCompanyAdded();
        onPick({entity_id:null,node_id:r.id,is_external:false});
        if(onMsg)onMsg(r.existing?('"'+r.name+'" was already registered — selected it.'):('Registered "'+r.name+'".'));
      }catch(ex){if(onErr)onErr(ex.message);}
      return;}
    if(v.startsWith('n'))onPick({entity_id:null,node_id:Number(v.slice(1)),is_external:false});
    else if(v.startsWith('e'))onPick({entity_id:Number(v.slice(1)),node_id:null,is_external:false});
    else onPick({entity_id:null,node_id:null,is_external:false});
  }}>
    <option value=''>{'\u2014 pick counterparty \u2014'}</option>
    <optgroup label='CloudLedger entities'>
      {(entityOpts||[]).map(x=><option key={x.id} value={'e'+x.id}>{x.name}</option>)}</optgroup>
    {offLedger.length>0&&<optgroup label='Registered companies (no ledger)'>
      {offLedger.map(c=><option key={c.id} value={'n'+c.id}>{c.name}</option>)}</optgroup>}
    <optgroup label='Other'>
      <option value='__ext'>External (a vendor or third party)</option>
      <option value='__new'>+ Register a new company…</option></optgroup>
  </select>;
}

function IcSearch({value,onChange,placeholder,width}){
  return<span style={{position:'relative',display:'inline-block'}}>
    <input style={{...S.inputSm,width:width||220,paddingRight:value?24:10}} value={value}
      onChange={e=>onChange(e.target.value)} placeholder={placeholder||'Search…'}/>
    {value&&<button title='Clear' onClick={()=>onChange('')}
      style={{position:'absolute',right:2,top:'50%',transform:'translateY(-50%)',background:'none',border:'none',
        color:T.textMuted,cursor:'pointer',fontSize:15,lineHeight:1,padding:'0 5px'}}>×</button>}
  </span>;
}

// Search over a reconciliation row: the counterparty's name, and every account
// code and name on either side. Someone hunting for 18307 should find it
// without having to know which counterparty it sits under.
function icRowMatches(r,n){
  if(!n)return true;
  if(String(r.counterparty_name||'').toLowerCase().includes(n))return true;
  return [...(r.our_legs||[]),...(r.their_legs||[])].some(l=>l&&(
    String(l.account_code||'').toLowerCase().includes(n)||
    String(l.account_name||'').toLowerCase().includes(n)));
}

// Which account on the counterparty's ledger answers ours.
//
// The list is ranked by the server: does their account name resolve back to us
// first, how closely the balances offset second. That order matters — only 2 of
// 28 relationships in this portfolio tie to the penny, so an account that names
// us while disagreeing on the number is still the right pick, and the
// disagreement is the finding rather than a reason to reject it.
function IcCpAccountSelect({entityId,accountCode,counterpartyEntityId,value,onChange,onLoaded,width,showGap=true}){
  const[cand,setCand]=useState(null);
  const[busy,setBusy]=useState(false);
  useEffect(()=>{
    let dead=false;
    if(!entityId||!accountCode||!counterpartyEntityId){setCand(null);return;}
    setBusy(true);
    api.getIcCounterpartyAccounts({entity_id:Number(entityId),
      counterparty_entity_id:Number(counterpartyEntityId),account_code:accountCode})
      .then(r=>{if(dead)return;setCand(r);if(onLoaded)onLoaded(r);})
      .catch(e=>{if(!dead){setCand(null);console.error('[ic] counterparty accounts:',e.message);}})
      .finally(()=>{if(!dead)setBusy(false);});
    return()=>{dead=true;};
  },[entityId,accountCode,counterpartyEntityId]);// eslint-disable-line react-hooks/exhaustive-deps

  if(!counterpartyEntityId)return null;
  const picked=cand&&value?cand.candidates.find(x=>String(x.account_code)===String(value)):null;
  return<>
    <select style={{...S.selectSm,minWidth:width||260}} value={value||''} onChange={e=>onChange(e.target.value)}
      title='Optional. The reconciliation nets by counterparty either way — naming the account lines the two entries up and brings their balance in even if they have not mapped their side.'>
      <option value=''>{busy?'Reading their ledger…':'— their account (optional) —'}</option>
      {cand&&cand.candidates.map(x=><option key={x.account_code} value={x.account_code}>
        {x.account_code+' '+x.account_name+'  '+fmt(x.balance)}</option>)}</select>
    {/* No gap amounts here. Whether the two sides agree is IC Reconciliation's
        job; this page only says WHICH account answers which. */}
    {showGap&&cand&&<div style={{flexBasis:'100%',fontSize:11.5,color:T.textDim,marginTop:2}}>
      {picked
        ?(picked.already_mapped_to&&!picked.already_mapped_to_us
          ?<span style={{color:T.red}}>that account is already mapped to {picked.already_mapped_to}</span>
          :null)
        :busy?'Looking through '+cand.counterparty.name+'\u2019s ledger…'
        :cand.candidates.length===0
          ?<span>Nothing on {cand.counterparty.name}{'\u2019'}s ledger names {cand.entity.name}. Leave it blank, or create the account.</span>
          :<span>{cand.candidates.length} candidate{cand.candidates.length===1?'':'s'} on {cand.counterparty.name}{'\u2019'}s ledger.</span>}
    </div>}
  </>;
}

// One account's mapping, saved from wherever the gap was noticed. `acct` only
// has to carry account_code and account_name; anything else it knows — the kind
// the name reads as, the counterparty the matcher resolved — seeds the form.
function IcMapForm({entityId,acct,entityOpts,companies,onCompanyAdded,onSaved,onCancel,onMsg,onErr}){
  const[f,setF]=useState({
    ic_type:acct.ic_type||'due_from',
    counterparty_entity_id:acct.counterparty_entity_id||'',
    counterparty_node_id:acct.counterparty_node_id||'',
    counterparty_account_code:acct.counterparty_account_code||'',
    is_external:!!acct.is_external,
    notes:''});
  const[busy,setBusy]=useState(false);
  // Only an entity can have a counterparty account — an off-ledger company has
  // no GL here to point at, and an external party has none by definition.
  const cpEid=f.is_external?null:(f.counterparty_entity_id?Number(f.counterparty_entity_id):null);

  const save=async()=>{
    if(!f.is_external&&!f.counterparty_entity_id&&!f.counterparty_node_id){
      if(onErr)onErr('Pick a counterparty for '+acct.account_code+', or mark it external.');return;}
    setBusy(true);
    try{
      await api.createIcMapping({entity_id:Number(entityId),
        account_code:acct.account_code,account_name:acct.account_name,ic_type:f.ic_type,
        counterparty_entity_id:f.is_external?null:(f.counterparty_entity_id?Number(f.counterparty_entity_id):null),
        counterparty_node_id:f.is_external?null:(f.counterparty_node_id?Number(f.counterparty_node_id):null),
        counterparty_account_code:f.is_external?null:(f.counterparty_account_code||null),
        is_external:f.is_external?1:0,notes:f.notes||null});
      if(onErr)onErr('');
      if(onMsg)onMsg('Mapped '+acct.account_code+' '+(acct.account_name||'')+'.');
      if(onSaved)await onSaved();
    }catch(e){if(onErr)onErr(e.message);}
    finally{setBusy(false);}
  };
  return<div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
    <select style={{...S.selectSm,minWidth:150}} value={f.ic_type} onChange={e=>setF(x=>({...x,ic_type:e.target.value}))}>
      {Object.keys(IC_TYPE_LABEL).map(k=><option key={k} value={k}>{IC_TYPE_LABEL[k]}</option>)}</select>
    <IcCpSelect entityOpts={entityOpts} companies={companies}
      entId={f.counterparty_entity_id} nodeId={f.counterparty_node_id} isExt={f.is_external}
      onPick={pk=>setF(x=>({...x,counterparty_entity_id:pk.entity_id||'',counterparty_node_id:pk.node_id||'',
        is_external:!!pk.is_external,counterparty_account_code:''}))}
      seedName={acct.counterparty_label||acct.account_name}
      onCompanyAdded={onCompanyAdded} onMsg={onMsg} onErr={onErr}/>
    <IcCpAccountSelect entityId={entityId} accountCode={acct.account_code} counterpartyEntityId={cpEid}
      value={f.counterparty_account_code}
      onChange={v=>setF(x=>({...x,counterparty_account_code:v}))}
      onLoaded={r=>setF(x=>({...x,counterparty_account_code:r.best||''}))} showGap={false}/>
    <input style={{...S.inputSm,width:170}} value={f.notes} placeholder='Notes (optional)'
      onChange={e=>setF(x=>({...x,notes:e.target.value}))}/>
    <button style={{...S.btnP,padding:'6px 12px',fontSize:12}} onClick={save} disabled={busy}>{busy?'Saving…':'Save mapping'}</button>
    <button style={{...S.btnGhost,fontSize:12}} onClick={onCancel}>Cancel</button>

    {/* The gap is stated in the row below rather than left to be worked out from
        two numbers in a dropdown. Only 2 of 28 relationships in this portfolio
        tie to the penny — a counterparty account that disagrees is still the
        right account, and the disagreement is the finding. */}
    <div style={{flexBasis:'100%'}}>
      <IcCpAccountSelect entityId={entityId} accountCode={acct.account_code} counterpartyEntityId={cpEid}
        value={f.counterparty_account_code} onChange={v=>setF(x=>({...x,counterparty_account_code:v}))}
        width={0} showGap={true}/></div>
  </div>;
}

// The mapped accounts: counterparty entity AND its GL account, both named.
// That pair is what makes a mapping checkable — two codes, two balances, and
// either they agree or they do not. A row missing either half is not here; it
// is work, and it shows up in the unmapped panel instead.
function IcMappedAccounts({entityId,entityName,canEdit,reloadKey,onMapped,onMsg,onErr}){
  const[data,setData]=useState(null);
  const[loading,setLoading]=useState(false);
  const[q,setQ]=useState('');
  const[pinning,setPinning]=useState(null);

  const load=useCallback(async()=>{
    if(!entityId){setData(null);return;}
    setLoading(true);
    try{setData(await api.getIcMappedAccounts(entityId));if(onErr)onErr('');}
    catch(e){if(onErr)onErr(e.message);setData(null);}
    finally{setLoading(false);}
  },[entityId,reloadKey]);// eslint-disable-line react-hooks/exhaustive-deps
  useEffect(()=>{load();},[load]);

  const n=q.trim().toLowerCase();
  const shown=(data?data.rows:[]).filter(r=>{
    if(!n)return true;
    return [r.account_code,r.account_name,r.counterparty_name,r.their_account_code,r.their_account_name]
      .some(v=>String(v||'').toLowerCase().includes(n));
  });

  return<div style={{...S.cardFlush,border:'1px solid '+T.border,marginBottom:16,overflow:'hidden'}}>
    <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border,
      display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
      <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>
        Mapped accounts{data?' ('+shown.length+')':''}</span>
      {data&&data.broken>0&&<span style={{color:T.red,fontSize:11.5}}>
        · {data.broken} name an account that is gone</span>}
      <IcSearch value={q} onChange={setQ} placeholder="Either side's code or name…" width={220}/>
      <button style={{...S.btnGhost,fontSize:12,marginLeft:'auto'}} onClick={load} disabled={loading}>
        {loading?'Loading…':'Refresh'}</button></div>

    {loading&&!data?<div style={{textAlign:'center',padding:30,color:T.textMuted}}>Reading both ledgers…</div>
    :!data||data.count===0?<div style={{padding:'22px 16px',textAlign:'center',color:T.textDim}}>
      <b style={{color:T.textBright}}>Nothing on {entityName||'this entity'} is fully mapped yet.</b><br/>
      <span style={{fontSize:12}}>A mapping counts as mapped once it names the counterparty entity
      {' '}<b>and</b> that entity{'\u2019'}s GL account.</span>
      {data&&data.incomplete_count>0&&<div style={{fontSize:11.5,marginTop:8}}>
        {data.incomplete_count} name{data.incomplete_count===1?'s':''} the entity but not the account
        {' '}— they are in <b>Show unmapped accounts</b>, under {'\u201c'}Mapped, but no counterparty account{'\u201d'}.</div>}</div>
    :shown.length===0?<div style={{padding:'22px 16px',textAlign:'center',color:T.textDim}}>
      Nothing matches. {data.count} mapped in total.</div>
    :<div className="cl-scroll" style={{overflow:'auto',maxHeight:'max(320px, calc(100vh - 460px))'}}>
      <table style={S.table}><thead><tr>
        <th colSpan={2} style={{...S.th,borderRight:'2px solid '+T.border}}>{entityName||'This entity'}</th>
        <th colSpan={2} style={S.th}>Counterparty</th>
        {canEdit&&<th style={{...S.th,width:110}}></th>}</tr></thead>
      <tbody>{shown.map(r=><Fragment key={r.id}>
        <tr style={pinning===r.id?{background:T.accentDim}:(r.self?{background:T.redDim}:null)}>
          <td style={{...S.td,whiteSpace:'normal',minWidth:200}}>
            <span style={{fontWeight:600,color:T.textBright}}>{r.account_code}</span>{' '}
            <span style={{color:T.textMuted}}>{r.account_name}</span>
            <span style={{marginLeft:6,fontSize:10,color:T.textDim}}>{IC_TYPE_LABEL[r.ic_type]||r.ic_type}</span></td>
          <td style={{...S.tdR,width:130,borderRight:'2px solid '+T.border}}>{fmt(r.balance)}</td>
          <td style={{...S.td,whiteSpace:'normal',minWidth:200}}>
            <div style={{color:T.textBright,fontWeight:600}}>{r.counterparty_name}
              {r.self&&<span style={{marginLeft:6}}><IcBadge kind="bad">points at itself</IcBadge></span>}</div>
            {r.investor_capital
              ?<div style={{marginTop:3}}><IcBadge kind="ok">investor capital — account not required</IcBadge></div>
              :<div style={{fontSize:12,color:T.textMuted,marginTop:2}}>
                <b style={{color:T.textBright}}>{r.their_account_code}</b> {r.their_account_name||''}
                {r.from_tb&&r.tb_as_of&&<span style={{marginLeft:6}}><IcBadge kind="info">their TB · {r.tb_as_of}</IcBadge></span>}
                {r.from_tb&&r.tb_missing&&<span style={{marginLeft:6}}><IcBadge kind="bad">TB deleted — re-upload it</IcBadge></span>}
                {r.their_account_missing&&<span style={{marginLeft:6}}><IcBadge kind="bad">{r.from_tb?'not in the uploaded TB':'not in their ledger'}</IcBadge></span>}</div>}</td>
          <td style={{...S.tdR,width:130}}>{r.investor_capital?<span style={{color:T.textDim}}>—</span>:(r.their_balance==null?'':fmt(r.their_balance))}</td>
          {canEdit&&<td style={{...S.td,textAlign:'right'}}>
            {!r.self&&!r.investor_capital&&!r.from_tb&&<button style={{...S.btnGhost,color:T.accent,fontSize:11}}
              onClick={()=>setPinning(pinning===r.id?null:r.id)}>
              {pinning===r.id?'Close':'Change'}</button>}
            <button style={{...S.btnGhost,color:T.red,fontSize:11}} title='Remove the mapping. The account returns to the unmapped worklist.'
              onClick={async()=>{if(!window.confirm('Remove the mapping for '+r.account_code+' '+(r.account_name||'')+'?'))return;
                try{await api.deleteIcMapping(r.id);if(onMsg)onMsg('Removed the mapping for '+r.account_code+'.');
                  await load();if(onMapped)await onMapped();}
                catch(e){if(onErr)onErr(e.message);}}}>x</button></td>}</tr>
        {canEdit&&pinning===r.id&&<tr style={{background:T.accentDim}}>
          <td colSpan={5} style={{...S.td,whiteSpace:'normal'}}>
            <IcPinForm mapping={{...r,counterparty_account_code:r.their_account_code}} entityId={entityId}
              onCancel={()=>setPinning(null)} onMsg={onMsg} onErr={onErr}
              onSaved={async()=>{setPinning(null);await load();if(onMapped)await onMapped();}}/></td></tr>}
      </Fragment>)}</tbody></table></div>}

    {/* What is not on this list, and why. Nothing is dropped without saying so. */}
    {data&&(data.incomplete_count>0||data.external_count>0||data.off_ledger_count>0)&&
      <div style={{padding:'10px 16px',borderTop:'1px solid '+T.border,color:T.textDim,fontSize:11.5}}>
        {data.incomplete_count>0&&<div>
          <b style={{color:T.textMuted}}>{data.incomplete_count}</b> name the counterparty entity but not its
          {' '}GL account — not mapped yet, and listed as work under <b>Show unmapped accounts</b>.</div>}
        {(data.external_count>0||data.off_ledger_count>0)&&<div style={{marginTop:4}}>
          <b style={{color:T.textMuted}}>{data.external_count+data.off_ledger_count}</b> face a counterparty with no
          {' '}ledger in CloudLedger{data.external_count>0&&data.off_ledger_count>0
            ?' ('+data.external_count+' external, '+data.off_ledger_count+' off-ledger companies)'
            :(data.external_count>0?' (external parties)':' (off-ledger companies)')}
          {' '}— there is no GL account to name, so they can never be mapped.</div>}</div>}
  </div>;
}

// Naming the counterparty account on a mapping that already names the entity.
// Same PUT the edit row uses — the whole mapping goes back, because the server
// validates it as a whole.
function IcPinForm({mapping,entityId,onSaved,onCancel,onMsg,onErr}){
  const[v,setV]=useState(mapping.counterparty_account_code||'');
  const[busy,setBusy]=useState(false);
  const save=async()=>{
    setBusy(true);
    try{
      await api.updateIcMapping(mapping.id,{
        account_code:mapping.account_code,account_name:mapping.account_name,
        ic_type:mapping.ic_type,
        counterparty_entity_id:mapping.counterparty_entity_id,
        counterparty_node_id:mapping.counterparty_node_id||null,
        counterparty_account_code:v||null,
        is_external:0,notes:mapping.notes||null});
      if(onErr)onErr('');
      if(onMsg)onMsg(v?(mapping.account_code+' now answers their '+v+'.')
        :('Cleared the counterparty account on '+mapping.account_code+'.'));
      if(onSaved)await onSaved();
    }catch(e){if(onErr)onErr(e.message);}
    finally{setBusy(false);}
  };
  return<div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
    <div style={{flexBasis:'100%',display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
      <IcCpAccountSelect entityId={entityId} accountCode={mapping.account_code}
        counterpartyEntityId={mapping.counterparty_entity_id} value={v} onChange={setV}
        onLoaded={r=>setV(x=>x||r.best||'')}/>
      <button style={{...S.btnP,padding:'6px 12px',fontSize:12}} onClick={save} disabled={busy}>
        {busy?'Saving…':'Save'}</button>
      <button style={{...S.btnGhost,fontSize:12}} onClick={onCancel}>Cancel</button></div>
  </div>;
}

// ── The unmapped worklist ──
// One list of everything standing between this entity and a fully mapped
// intercompany position: accounts with no mapping at all, and mappings that
// name the counterparty entity but not its GL account. Each row arrives with
// the counterparty AND its account already filled in — an existing account
// when the counterparty's ledger holds one, otherwise a draft of the account
// to create, named by mirroring ours. Confirm is the only click a correct row
// needs; nothing is saved or created before it.
function IcUnmappedAccounts({entityId,entityName,entityOpts,companies,canEdit,reloadKey,onCompanyAdded,onMapped,onMsg,onErr}){
  const[data,setData]=useState(null);
  const[loading,setLoading]=useState(false);
  const[q,setQ]=useState('');
  const[mapping,setMapping]=useState(null);
  // Per-row overrides, keyed by row: which answer mode is active and what the
  // person typed. Untouched rows confirm the server's suggestion as-is.
  const[ov,setOv]=useState({});
  const[busyKey,setBusyKey]=useState(null);

  const load=useCallback(async()=>{
    if(!entityId){setData(null);return;}
    setLoading(true);
    try{setData(await api.getIcUnmappedAccounts(entityId));setOv({});if(onErr)onErr('');}
    catch(e){if(onErr)onErr(e.message);setData(null);}
    finally{setLoading(false);}
  },[entityId,reloadKey]);// eslint-disable-line react-hooks/exhaustive-deps
  useEffect(()=>{load();},[load]);

  const rowKey=r=>r.mapping_id?('m'+r.mapping_id):('a'+r.account_code);
  const setSt=(r,patch)=>setOv(o=>({...o,[rowKey(r)]:{...(o[rowKey(r)]||{}),...patch}}));
  const modeOf=r=>{const st=ov[rowKey(r)]||{};
    return st.mode||(r.suggested_existing?'existing':(r.suggested_new?'new':null));};

  const n=q.trim().toLowerCase();
  const shown=(data?data.accounts:[]).filter(a=>{
    if(!n)return true;
    return [a.account_code,a.account_name,a.counterparty_label,a.counterparty_name,
      a.suggested_existing&&a.suggested_existing.account_code,
      a.suggested_existing&&a.suggested_existing.account_name,
      a.suggested_new&&a.suggested_new.account_name]
      .some(v=>String(v||'').toLowerCase().includes(n));
  });

  // Ready = the answer is already on screen: a pre-filled account, an editable
  // draft, external, or investor capital. A row still waiting on a human
  // choice is not ready, and Confirm all leaves it untouched.
  const isReady=r=>{
    if(r.individual)return false;
    if(r.is_external||r.investor_capital)return true;
    const st=ov[rowKey(r)]||{},mode=modeOf(r);
    // Answered from an uploaded TB: ready once a line is chosen (or pre-filled).
    if(r.tb_as_of)return mode==='existing'||!!st.pickCode;
    const cpEff=st.cpOverride||r.counterparty_entity_id;
    if(!cpEff||!mode)return false;
    if(mode==='pick'&&!st.pickCode)return false;
    return true;
  };
  const readyCount=shown.filter(isReady).length;

  // The row-level save, shared by Confirm and Confirm all. Throws on failure
  // so the caller decides how to report it.
  const saveRow=async r=>{
    const k=rowKey(r),st=ov[k]||{},mode=modeOf(r);
    const cpEff=st.cpOverride||r.counterparty_entity_id;
    let cpCode=null;
    if(!r.is_external&&!r.investor_capital){
      if(r.tb_as_of){
        // The counterparty's ledger is an uploaded TB — a line is picked,
        // never drafted: a TB is a statement someone issued, not a chart.
        cpCode=(mode==='existing'&&r.suggested_existing)?r.suggested_existing.account_code:(st.pickCode||null);
        if(!cpCode)throw new Error('Pick the counterparty account from the uploaded TB for '+r.account_code+' first.');
      }else if(mode==='new'){
        // ONE box: leading token is the code, the rest is the name.
        const raw=String(st.acct!=null?st.acct:(r.suggested_new.account_code+' '+r.suggested_new.account_name)).trim();
        const m2=raw.match(/^(\S+)\s+(.+)$/);
        if(!m2)throw new Error('Enter the new account for '+r.account_code+' as “code name” — e.g. “23000 Due to Banyan Residential”.');
        // Created on the COUNTERPARTY's ledger, at zero. The reconciliation
        // will show our balance against their zero until they book the entry
        // — that gap is the truth, not a defect.
        await api.createAccount(cpEff,{code:m2[1],name:m2[2],type:r.suggested_new.type,subtype:'',bank_acct:false});
        cpCode=m2[1];
      }else if(mode==='pick')cpCode=st.pickCode||null;
      else if(mode==='existing')cpCode=r.suggested_existing.account_code;
      if(!cpCode)throw new Error('Pick or create the counterparty account for '+r.account_code+' first.');
    }
    // Investor capital completes WITHOUT an account: the contributor keeps no
    // ledger in CloudLedger. Saved to the registered company if one matched,
    // external otherwise, with the investor's name kept in the notes.
    const payload=r.investor_capital
      ?{account_code:r.account_code,account_name:r.account_name,ic_type:r.ic_type,
        counterparty_entity_id:null,
        counterparty_node_id:r.counterparty_node_id||null,
        counterparty_account_code:null,
        is_external:r.counterparty_node_id?0:1,
        notes:r.notes||r.counterparty_label||null}
      :{account_code:r.account_code,account_name:r.account_name,ic_type:r.ic_type,
        counterparty_entity_id:r.is_external?null:(r.tb_as_of?null:cpEff),
        counterparty_node_id:r.tb_as_of?(r.counterparty_node_id||null):null,
        counterparty_account_code:r.is_external?null:cpCode,
        is_external:r.is_external?1:0,notes:r.notes||(r.is_external?(r.counterparty_label||null):null)};
    if(r.mapping_id)await api.updateIcMapping(r.mapping_id,payload);
    else await api.createIcMapping({...payload,entity_id:Number(entityId)});
    return{mode,cpCode};
  };

  const confirmRow=async r=>{
    setBusyKey(rowKey(r));
    try{
      const res=await saveRow(r);
      if(onErr)onErr('');
      if(onMsg)onMsg(r.investor_capital?('Mapped '+r.account_code+' as investor capital — no counterparty account required.')
        :r.is_external?('Mapped '+r.account_code+' as external.')
        :(r.account_code+' now answers '+res.cpCode+(res.mode==='new'?' (created)':'')+'.'));
      await load();if(onMapped)await onMapped();
    }catch(e){if(onErr)onErr(e.message);}
    finally{setBusyKey(null);}
  };

  const confirmAll=async()=>{
    const ready=shown.filter(isReady);
    if(!ready.length)return;
    const creates=ready.filter(r=>!r.is_external&&!r.investor_capital&&modeOf(r)==='new').length;
    // One dialog, stating the consequential part out loud: how many accounts
    // get CREATED on counterparties' charts of accounts.
    if(!window.confirm('Confirm '+ready.length+' row'+(ready.length===1?'':'s')+' as shown?'
      +(creates?('\n\n'+creates+' new account'+(creates===1?'':'s')+' will be created on the counterparties’ charts of accounts.'):'')))return;
    setBusyKey('__all');
    const failed=[];
    for(const r of ready){
      try{await saveRow(r);}
      catch(e){failed.push(r.account_code+': '+e.message);}
    }
    // Failures are named, not counted — a row that silently stayed unmapped
    // would look identical to one that was confirmed.
    if(onErr)onErr(failed.length?('Could not confirm '+failed.length+' — '+failed.join(' · ')):'');
    if(onMsg)onMsg('Confirmed '+(ready.length-failed.length)+' of '+ready.length+' row'+(ready.length===1?'':'s')+'.');
    await load();if(onMapped)await onMapped();
    setBusyKey(null);
  };

  return<div style={{...S.cardFlush,border:'1px solid '+T.border,marginBottom:16,overflow:'hidden'}}>
    <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border,
      display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
      <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>
        Unmapped accounts{data?' ('+shown.length+')':''}</span>
      <span style={{color:T.textDim,fontSize:11.5}} title='Due from / Due to and Investment / Contributed capital, with a balance. A mapping without the counterparty GL account counts as unmapped.'>
        both sides pre-filled — Confirm saves the row</span>
      {data&&data.shell_entity&&<IcBadge kind="info">shell entity — contributed capital left out</IcBadge>}
      {data&&data.skipped_zero_balance>0&&<span style={{color:T.textDim,fontSize:11.5}}
        title='A zero balance has nothing for a counterparty ledger to disagree with, so mapping it would change no reconciliation.'>
        · {data.skipped_zero_balance} at zero not listed</span>}
      <IcSearch value={q} onChange={setQ} placeholder='Account code or name…' width={200}/>
      <span style={{marginLeft:'auto',display:'flex',gap:8,alignItems:'center'}}>
        {canEdit&&readyCount>0&&<button style={{...S.btnP,padding:'7px 14px',fontSize:12.5}}
          onClick={confirmAll} disabled={busyKey!=null||loading}>
          {busyKey==='__all'?'Confirming…':'Confirm all ('+readyCount+')'}</button>}
        <button style={{...S.btnGhost,fontSize:12}} onClick={load} disabled={loading}>
          {loading?'Loading…':'Refresh'}</button></span></div>

    {loading&&!data?<div style={{textAlign:'center',padding:30,color:T.textMuted}}>Reading both sides…</div>
    :!data||data.count===0?<div style={{padding:'22px 16px',textAlign:'center',color:T.textDim}}>
      <b style={{color:T.textBright}}>Nothing here needs mapping.</b><br/>
      <span style={{fontSize:12}}>Every due from / due to, investment and contributed-capital account
      on {entityName||'this entity'} that carries a balance is mapped, counterparty account included.</span>
      {data&&data.skipped_zero_balance>0&&<div style={{fontSize:11.5,marginTop:8}}>
        {data.skipped_zero_balance} at a zero balance — nothing to reconcile, so nothing to map.</div>}</div>
    :shown.length===0?<div style={{padding:'22px 16px',textAlign:'center',color:T.textDim}}>
      No unmapped account matches “{q}”. {data.count} {data.count===1?'is':'are'} unmapped in total.</div>
    :<div className="cl-scroll" style={{overflow:'auto',maxHeight:'max(340px, calc(100vh - 420px))'}}>
      <table style={S.table}><thead><tr>
        <th style={S.th}>Account</th><th style={S.thR}>Balance</th>
        <th style={S.th}>Counterparty</th><th style={S.th}>Counterparty account</th>
        {canEdit&&<th style={{...S.th,width:130}}></th>}</tr></thead>
      <tbody>{shown.map(r=>{
        const k=rowKey(r),st=ov[k]||{},mode=modeOf(r),busy=busyKey===k||busyKey==='__all';
        // The accountant may re-point the counterparty; the account lookup on
        // the right then runs against the newly chosen entity.
        const cpEff=st.cpOverride||r.counterparty_entity_id;
        const confirmable=!r.individual&&(r.is_external||r.investor_capital
          ||(r.tb_as_of?(mode==='existing'||!!st.pickCode):(cpEff&&mode)));
        return<Fragment key={k}>
        <tr style={mapping===k?{background:T.accentDim}:null}>
          <td style={{...S.td,whiteSpace:'normal',minWidth:220}}>
            <span style={{fontWeight:600,color:T.textBright}}>{r.account_code}</span>{' '}
            <span style={{color:T.textMuted}}>{r.account_name}</span>
            {r.individual&&<span style={{marginLeft:6}}><IcBadge kind="mute">individual</IcBadge></span>}</td>
          <td style={{...S.tdR,width:120,fontWeight:600,color:T.textBright}}>{fmt(r.balance)}</td>
          <td style={{...S.td,whiteSpace:'normal'}}>
            {r.investor_capital?<><IcBadge kind="warn">investor — not in CL</IcBadge>
              {r.counterparty_label&&<div style={{color:T.textMuted,fontSize:12,marginTop:2}}>{r.counterparty_label}</div>}</>
            :r.is_external?<IcBadge kind="warn">external</IcBadge>
            :r.counterparty_entity_id?<select style={{...S.selectSm,minWidth:210}} value={String(cpEff)}
                onChange={e=>setSt(r,{cpOverride:Number(e.target.value),mode:'pick',pickCode:''})}>
                {entityOpts.map(x=><option key={x.id} value={String(x.id)}>{x.name}</option>)}</select>
            :r.tb_as_of?<>{r.counterparty_label||r.counterparty_name} <IcBadge kind="info">their TB · {r.tb_as_of}</IcBadge></>
            :r.counterparty_node_id?<>{r.counterparty_label} <IcBadge kind="info">no ledger</IcBadge></>
            :<span style={{color:T.textDim,fontStyle:'italic'}}>{r.counterparty_label||'not recognised'}</span>}</td>
          <td style={{...S.td,whiteSpace:'normal',minWidth:280}}>
            {r.investor_capital?<span style={{color:T.green,fontSize:12.5,fontWeight:600}}>No counterparty account needed
                <span style={{display:'block',fontWeight:400,color:T.textMuted,fontSize:11.5}}>This investor keeps no ledger in CloudLedger, so there is no investment account to match.</span></span>
            :r.is_external?<span style={{color:T.textDim}}>outside the group — nothing to map to</span>
            :r.tb_as_of?(mode==='existing'&&r.suggested_existing
              ?<div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                <span style={{display:'inline-flex',gap:6,alignItems:'center',border:'1px solid '+T.border,borderRadius:7,padding:'6px 10px'}}>
                  <span style={{fontWeight:600,color:T.textBright}}>{r.suggested_existing.account_code}</span>
                  <span style={{color:T.textMuted}}>{r.suggested_existing.account_name}</span></span>
                <IcBadge kind="info">from uploaded TB</IcBadge>
                <button style={S.link} onClick={()=>setSt(r,{mode:'pick',pickCode:''})}>change…</button></div>
              :<div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                <IcTbAccountSelect nodeId={r.counterparty_node_id} asOf={r.tb_as_of}
                  value={st.pickCode||''} onChange={v=>setSt(r,{mode:'pick',pickCode:v})}/>
                {r.suggested_existing&&<button style={S.link} onClick={()=>setSt(r,{mode:'existing'})}>back to suggestion</button>}</div>)
            :!cpEff?<span style={{color:T.textDim}}>—</span>
            :mode==='pick'?<div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                <IcCpAccountSelect entityId={entityId} accountCode={r.account_code}
                  counterpartyEntityId={cpEff} value={st.pickCode||''}
                  onChange={v=>setSt(r,{mode:'pick',pickCode:v})}
                  onLoaded={res=>{if(!st.pickCode&&res.best)setSt(r,{mode:'pick',pickCode:res.best});}}
                  width={280} showGap={false}/>
                {r.suggested_new&&!st.cpOverride&&<button style={S.link} onClick={()=>setSt(r,{mode:'new'})}>new account…</button>}</div>
            :mode==='new'?<div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                {/* ONE box: the leading token is the code, the rest is the name. */}
                <input style={{...S.inputSm,width:270,borderStyle:'dashed',borderColor:T.accent}}
                  value={st.acct!=null?st.acct:(r.suggested_new.account_code+' '+r.suggested_new.account_name)}
                  onChange={e=>setSt(r,{mode:'new',acct:e.target.value})}/>
                <IcBadge kind="info">will be created when confirmed</IcBadge>
                <button style={S.link} onClick={()=>setSt(r,{mode:'pick'})}>pick existing…</button>
                <div style={{flexBasis:'100%',fontSize:11,color:T.textDim}}>
                  This account does not exist yet. Confirm adds it to the counterparty{'’'}s chart of accounts
                  (at a zero balance) and saves the mapping in one step.</div></div>
            :mode==='existing'?<div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
                <span style={{display:'inline-flex',gap:6,alignItems:'center',border:'1px solid '+T.border,borderRadius:7,padding:'6px 10px'}}>
                  <span style={{fontWeight:600,color:T.textBright}}>{r.suggested_existing.account_code}</span>
                  <span style={{color:T.textMuted}}>{r.suggested_existing.account_name}</span></span>
                <button style={S.link} onClick={()=>setSt(r,{mode:'pick'})}>change…</button></div>
            :<span style={{color:T.textDim}}>—</span>}</td>
          {canEdit&&<td style={{...S.td,textAlign:'right'}}>
            {confirmable
              ?<><button style={{...S.btnP,padding:'5px 12px',fontSize:12}} disabled={busy}
                  onClick={()=>confirmRow(r)}>{busy?'Saving…':'Confirm'}</button>
                {!r.mapping_id&&<button style={{...S.btnGhost,fontSize:10.5}}
                  onClick={()=>setMapping(mapping===k?null:k)}>{mapping===k?'close':'other…'}</button>}</>
              :!r.mapping_id&&<button style={{...S.btnGhost,color:T.accent,fontSize:11}}
                  onClick={()=>setMapping(mapping===k?null:k)}>{mapping===k?'Close':'Map…'}</button>}</td>}</tr>
        {canEdit&&mapping===k&&<tr style={{background:T.accentDim}}>
          <td colSpan={5} style={{...S.td,whiteSpace:'normal'}}>
            <IcMapForm entityId={entityId} acct={r} entityOpts={entityOpts} companies={companies}
              onCompanyAdded={onCompanyAdded} onCancel={()=>setMapping(null)} onMsg={onMsg} onErr={onErr}
              onSaved={async()=>{setMapping(null);await load();if(onMapped)await onMapped();}}/></td></tr>}
      </Fragment>;})}</tbody></table></div>}
  </div>;
}

function IntercompanyReconciliation({entities,activeEntity,setPage,canEdit=true}){
  const[tab,setTab]=useState('due');
  const[eid,setEid]=useState(activeEntity||'');
  const[asOf,setAsOf]=useState(today());
  const[data,setData]=useState(null);
  const[loading,setLoading]=useState(false);const[err,setErr]=useState('');
  const[msg,setMsg]=useState('');
  const[q,setQ]=useState('');
  // An unmapped account is mapped here, on the page that found it.
  const[companies,setCompanies]=useState([]);
  const[mapping,setMapping]=useState(null);
  const entityOpts=[...(entities||[])].sort((a,b)=>String(a.name).localeCompare(String(b.name),undefined,{sensitivity:'base'}));
  const loadCompanies=useCallback(async()=>{
    try{setCompanies(await api.getIcCompanies());}catch(e){console.error('[ic] companies:',e.message);}},[]);
  useEffect(()=>{loadCompanies();},[loadCompanies]);

  const run=useCallback(async()=>{
    if(!eid){setData(null);return;}
    setLoading(true);setErr('');
    try{setData(await api.reconcileIcEntity(eid,asOf));}
    catch(e){setErr(e.message);setData(null);}
    finally{setLoading(false);}
  },[eid,asOf]);
  useEffect(()=>{run();},[eid]);// eslint-disable-line react-hooks/exhaustive-deps

  const side=data?(tab==='due'?data.due:data.investment):null;
  const us=data?data.entity.name:'Selected entity';

  // Search narrows what is SHOWN, never what was computed. The totals line
  // keeps reporting the whole entity, so filtering the table can never make an
  // unexplained difference look smaller than it is.
  const needle=q.trim().toLowerCase();
  const acctMatches=a=>!needle
    ||String(a.account_code||'').toLowerCase().includes(needle)
    ||String(a.account_name||'').toLowerCase().includes(needle);
  const shownRows=side?side.rows.filter(r=>icRowMatches(r,needle)):[];
  const shownUnmapped=side?(side.unmapped||[]).filter(acctMatches):[];
  const shownExternal=side?(side.external||[]).filter(acctMatches):[];

  const doExport=()=>{
    if(!data)return;
    const d=[['Intercompany Reconciliation'],[data.entity.name+' — as of '+(data.as_of||asOf)],[]];
    const dump=(title,rows,cols)=>{
      d.push([title]);
      d.push(['Counterparty','Status','Difference',data.entity.name+' acct','Account name','Amount','Counterparty acct','Account name','Amount']);
      rows.forEach(r=>{
        zipLegs(r.our_legs,r.their_legs).forEach(([a,b],i)=>d.push([
          i===0?r.counterparty_name:'',i===0?(IC_STATUS[r.status]||['','' ])[1]:'',i===0?r.difference:'',
          a?a.account_code:'',a?a.account_name:'',a?a.amount:'',
          b?b.account_code:'',b?b.account_name:'',b?b.amount:'']));
        if(cols)d.push(['','','','Net',...['',r.our_net,'','',r.their_net]]);
      });
      d.push([]);
    };
    dump('DUE FROM / DUE TO',data.due.rows,true);
    if(data.investment.findings.length){
      d.push(['INVESTMENT IN ITSELF']);
      d.push(['Account','Account name','Amount']);
      data.investment.findings.forEach(f=>d.push([f.account_code,f.account_name,f.amount]));
      d.push([]);
    }
    dump('INVESTMENT / CONTRIBUTED CAPITAL',data.investment.rows,false);
    const ext=[...data.due.external,...data.investment.external];
    if(ext.length){
      d.push(['OUTSIDE THE GROUP']);
      d.push(['Account','Account name','Type','Our balance','Their value (entered)','Difference']);
      ext.forEach(x=>d.push([x.account_code,x.account_name,IC_TYPE_LABEL[x.ic_type]||x.ic_type,x.amount,
        x.manual_balance==null?'':x.manual_balance,x.difference==null?'':x.difference]));
      d.push([]);
    }
    exportToExcel(d,'IC_Recon_'+String(data.entity.code||data.entity.id)+'_'+String(data.as_of||asOf).replace(/-/g,'')+'.xlsx');
  };

  const StatusLine=({t})=>{
    if(!t)return null;
    const bits=[['matched',t.matched],['mismatch',t.mismatched],['one_sided',t.one_sided],['no_ledger',t.no_ledger]]
      .filter(([,n])=>n>0);
    return<div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap',marginBottom:14}}>
      {bits.length?bits.map(([k,n])=><span key={k}>{n} <IcStatus status={k}/></span>)
        :<span style={{color:T.textDim,fontSize:12}}>nothing with a balance</span>}
      {t.settled>0&&<span style={{color:T.textDim,fontSize:12}}>· {t.settled} settled</span>}
      {t.abs_difference>0.005&&<span style={{marginLeft:'auto',fontWeight:700,color:T.red,fontVariantNumeric:'tabular-nums'}}>
        {fmt(t.abs_difference)} unexplained</span>}
    </div>;
  };

  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8,flexWrap:'wrap',gap:12}}>
      <div><div style={S.h1}>IC Reconciliation</div>
        <div style={S.sub}>Every counterparty of the selected entity, with the account number and name on both sides.</div></div>
      <div style={{display:'flex',gap:8,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div><label style={S.label}>Entity</label>
          <select style={{...S.selectSm,minWidth:240}} value={eid} onChange={e=>setEid(e.target.value)}>
            <option value=''>Select an entity…</option>
            {entityOpts.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select></div>
        <div><label style={S.label}>As of</label>
          <input style={S.inputSm} type='date' value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
        <div><label style={S.label}>Search</label>
          <IcSearch value={q} onChange={setQ} placeholder='Counterparty, account code or name…' width={250}/></div>
        <button style={S.btnP} onClick={run} disabled={!eid||loading}>{loading?'Running…':'Run'}</button>
        {data&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</div></div>

    {err&&<div style={{...S.card,borderColor:T.red+'40'}}><div style={S.err}>{err}</div></div>}
    {msg&&<div style={{...S.card,borderColor:T.green+'40',padding:14}}><div style={S.success}>{msg}</div></div>}
    {!eid&&!err&&<IcEmpty>Pick an entity to reconcile.</IcEmpty>}

    {eid&&<div style={{display:'flex',gap:4,marginBottom:16,borderBottom:'1px solid '+T.border}}>
      {[['due','Due from / Due to'],['investment','Investment / Contributed capital']].map(([k,l])=>
        <button key={k} onClick={()=>setTab(k)} style={{background:'none',border:'none',borderBottom:'2px solid '+(tab===k?T.accent:'transparent'),color:tab===k?T.accent:T.textMuted,fontWeight:tab===k?700:500,fontSize:13,padding:'9px 16px',cursor:'pointer',marginBottom:-1}}>{l}
          {k==='due'&&data&&data.due.totals.mismatched>0&&<span style={{marginLeft:7}}><IcBadge kind="bad">{data.due.totals.mismatched}</IcBadge></span>}
          {k==='investment'&&data&&(data.investment.findings.length+data.investment.totals.mismatched)>0&&
            <span style={{marginLeft:7}}><IcBadge kind="bad">{data.investment.findings.length+data.investment.totals.mismatched}</IcBadge></span>}</button>)}</div>}

    {loading&&<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Reading balances…</div>}

    {!loading&&data&&<>
      <StatusLine t={side.totals}/>

      {/* The investment tab says so plainly when the entity simply has none,
          rather than showing an empty table that reads like a failure. */}
      {tab==='investment'&&!data.investment.has_any
        ?<IcEmpty><b>{data.entity.name}</b> has no investment or contributed-capital accounts.<br/>
          <span style={{fontSize:12}}>Nothing to reconcile on this tab for this entity.</span></IcEmpty>
      :side.rows.length===0&&!(side.external||[]).length&&(tab==='due'||!data.investment.findings.length)
        ?<IcEmpty>No mapped {tab==='due'?'due from / due to':'investment'} account on <b>{data.entity.name}</b> faces another group entity at {data.as_of||asOf}.
          {' '}Map its accounts on the <button style={S.link} onClick={()=>setPage&&setPage('ic_mapping')}>IC Mapping</button> page.</IcEmpty>
      :side.rows.length===0&&(side.external||[]).length>0
        ?<IcEmpty>Every mapped {tab==='due'?'due from / due to':'investment'} account on <b>{data.entity.name}</b> faces a party outside the group,
          {' '}so there is nothing to reconcile. They are listed below.</IcEmpty>
      :null}

      {tab==='investment'&&data.investment.findings.length>0&&
        <div style={{...S.cardFlush,border:'1px solid '+T.red+'40',overflow:'hidden'}}>
          <div style={{padding:'12px 16px',background:T.redDim,borderBottom:'1px solid '+T.border}}>
            <span style={{fontWeight:700,color:T.red,fontSize:13}}>Investment in itself ({data.investment.findings.length})</span>
            <span style={{color:T.textMuted,fontSize:12,marginLeft:8}}>Inflates this entity's own assets and equity by the same amount.</span></div>
          <table style={S.table}><tbody>{data.investment.findings.map(f=><tr key={f.account_code}>
            <td style={{...S.td,fontWeight:600,color:T.textBright,width:88}}>{f.account_code}</td>
            <td style={{...S.td,whiteSpace:'normal'}}>{f.account_name}</td>
            <td style={{...S.tdR,fontWeight:700,color:T.red,width:150}}>{fmt(f.amount)}</td></tr>)}</tbody></table></div>}

      {shownRows.length>0&&<div className="cl-scroll" style={scrollBox()}>
        <table style={S.table}>
          <thead><tr>
            <th colSpan={3} style={{...S.th,borderRight:'2px solid '+T.border}}>{us}</th>
            <th colSpan={3} style={S.th}>Counterparty</th>
            <th style={S.thR}>Difference</th></tr></thead>
          <tbody>{shownRows.map(r=>{
            const pairs=zipLegs(r.our_legs,r.their_legs);
            const bad=r.status==='mismatch'||r.status==='self';
            const warn=r.status==='one_sided';
            return<Fragment key={(r.counterparty_entity_id||r.counterparty_node_id||r.counterparty_name)+'-'+r.status}>
              <tr style={{background:bad?T.redDim:(warn?T.orangeDim:T.bgElevated)}}>
                <td colSpan={6} style={{...S.td,fontWeight:700,color:T.textBright}}>
                  {r.counterparty_name}
                  {r.tb_as_of?<span style={{marginLeft:8}}><IcBadge kind="info">their TB · {r.tb_as_of}</IcBadge></span>
                    :r.off_ledger&&<span style={{marginLeft:8}}><IcBadge kind="info">no ledger</IcBadge></span>}
                  <span style={{marginLeft:8}}><IcStatus status={r.status}/></span></td>
                <td style={{...S.tdR,fontWeight:700,color:bad?T.red:(warn?T.orange:T.green)}}>{fmt(r.difference)}</td></tr>
              {pairs.map(([a,b],i)=><tr key={i}>
                {legCells(a)}
                <td style={{...S.td,borderLeft:'2px solid '+T.border,width:88,color:T.textBright,fontWeight:600}}>{b?b.account_code:''}</td>
                <td style={{...S.td,whiteSpace:'normal',color:T.textMuted}}>{b?b.account_name:
                  (i===0&&!r.their_legs.length?<span style={{color:T.textDim,fontStyle:'italic'}}>
                    {r.off_ledger?'no ledger in CloudLedger':'this entity has not mapped its side'}</span>:'')}
                  {b&&b.ic_type&&<span style={{marginLeft:6,fontSize:10,color:T.textDim}}>{IC_TYPE_LABEL[b.ic_type]}</span>}
                  {b&&b.from_pin&&<span style={{marginLeft:6}} title={'Read straight from '+r.counterparty_name+'\u2019s ledger because the mapping on this side names this account. '+r.counterparty_name+' has not mapped it.'}>
                    <IcBadge kind="info">their GL</IcBadge></span>}
                  {b&&b.from_tb&&<span style={{marginLeft:6}} title={'Read from the uploaded trial balance as of '+(r.tb_as_of||'')+'.'}>
                    <IcBadge kind="info">their TB</IcBadge></span>}</td>
                <td style={{...S.tdR,width:130}}>{b?fmt(b.amount):''}</td>
                <td style={S.tdR}></td></tr>)}
              {tab==='due'&&<tr style={{background:T.bgElevated}}>
                <td style={S.td}></td>
                <td style={{...S.td,fontWeight:600,color:T.textMuted}}>Net receivable</td>
                <td style={{...S.tdR,fontWeight:700}}>{fmt(r.our_net)}</td>
                <td style={{...S.td,borderLeft:'2px solid '+T.border}}></td>
                <td style={{...S.td,fontWeight:600,color:T.textMuted}}>Net receivable</td>
                <td style={{...S.tdR,fontWeight:700}}>{fmt(r.their_net)}</td>
                <td style={S.tdR}></td></tr>}
              {tab==='investment'&&<tr style={{background:T.bgElevated}}>
                <td style={S.td}></td>
                <td style={{...S.td,fontWeight:600,color:T.textMuted}}>Investment / Capital</td>
                <td style={{...S.tdR,fontWeight:700}}>{fmt(r.our_investment)} / {fmt(r.our_capital)}</td>
                <td style={{...S.td,borderLeft:'2px solid '+T.border}}></td>
                <td style={{...S.td,fontWeight:600,color:T.textMuted}}>Investment / Capital</td>
                <td style={{...S.tdR,fontWeight:700}}>{fmt(r.their_investment)} / {fmt(r.their_capital)}</td>
                <td style={S.tdR}></td></tr>}
            </Fragment>;})}</tbody></table></div>}

      {needle&&side.rows.length>0&&shownRows.length===0&&
        <IcEmpty>No counterparty or account on this tab matches “{q}”.</IcEmpty>}

      {/* Mapped right here. Walking to the setup page to retype an account code
          that is already on the screen is the reason these sat unmapped. */}
      {shownUnmapped.length>0&&<div className="cl-scroll" style={scrollBox()}>
        <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border}}>
          <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>Not mapped yet ({shownUnmapped.length}{shownUnmapped.length!==side.unmapped.length?' of '+side.unmapped.length:''})</span>
          <span style={{color:T.textMuted,fontSize:12,marginLeft:8}}>These accounts on {data.entity.name} carry a balance and look intercompany, but no mapping says who they face.</span>
          <button style={{...S.link,marginLeft:10}} onClick={()=>setPage&&setPage('ic_mapping')}>Open IC Mapping</button></div>
        <table style={S.table}><tbody>{shownUnmapped.map(u=><Fragment key={u.account_code}>
          <tr style={mapping===u.account_code?{background:T.accentDim}:null}>
            <td style={{...S.td,fontWeight:600,color:T.textBright,width:88}}>{u.account_code}</td>
            <td style={{...S.td,whiteSpace:'normal'}}>{u.account_name}
              {u.reason&&<span style={{color:T.textDim,fontSize:11,marginLeft:8}}>{u.reason}</span>}</td>
            <td style={{...S.tdR,width:150}}>{fmt(u.balance)}</td>
            {canEdit&&<td style={{...S.td,width:80,textAlign:'right'}}>
              <button style={{...S.btnGhost,color:T.accent,fontSize:11}}
                onClick={()=>{setMapping(mapping===u.account_code?null:u.account_code);setMsg('');}}>
                {mapping===u.account_code?'Close':'Map'}</button></td>}</tr>
          {canEdit&&mapping===u.account_code&&<tr style={{background:T.accentDim}}>
            <td colSpan={4} style={{...S.td,whiteSpace:'normal'}}>
              <IcMapForm entityId={data.entity.id} acct={u} entityOpts={entityOpts} companies={companies}
                onCompanyAdded={loadCompanies} onCancel={()=>setMapping(null)} onMsg={setMsg} onErr={setErr}
                onSaved={async()=>{setMapping(null);await run();}}/></td></tr>}
        </Fragment>)}</tbody></table></div>}

      {/* Not reconciled — there is no counterparty ledger to agree with, so a
          difference would be meaningless. Listed anyway so the balance is not
          silently missing from the page. */}
      {shownExternal.length>0&&<details open={shownExternal.some(x=>x.manual_balance!=null)} style={{...S.cardFlush,marginTop:14,padding:'10px 16px'}}>
        <summary style={{cursor:'pointer',color:T.textMuted,fontSize:12.5}}>
          <b style={{color:T.textBright}}>{shownExternal.length}</b> account{shownExternal.length>1?'s':''} face
          {shownExternal.length>1?'':'s'} a party outside the group
          <span style={{marginLeft:8,fontVariantNumeric:'tabular-nums'}}>({fmt(shownExternal.reduce((s,x)=>s+Math.abs(x.amount),0))})</span>
          <span style={{marginLeft:8,color:T.textDim}}>— no ledger here; enter their value to compare</span>
        </summary>
        <table style={{...S.table,marginTop:8}}>
          <thead><tr><th style={S.th}>Account</th><th style={S.th}></th>
            <th style={S.thR}>Our balance</th><th style={S.thR}>Their value (entered)</th><th style={S.thR}>Difference</th></tr></thead>
          <tbody>{shownExternal.map(x=><tr key={x.account_code}>
          <td style={{...S.td,fontWeight:600,color:T.textBright,width:88}}>{x.account_code}</td>
          <td style={{...S.td,whiteSpace:'normal',color:T.textMuted}}>{x.account_name}
            <span style={{marginLeft:6,fontSize:10,color:T.textDim}}>{IC_TYPE_LABEL[x.ic_type]}</span></td>
          <td style={{...S.tdR,width:130}}>{fmt(x.amount)}</td>
          <td style={{...S.tdR,width:170}}>
            {canEdit?<input key={x.mapping_id+':'+(x.manual_balance==null?'':x.manual_balance)}
              style={{...S.inputSm,width:140,textAlign:'right'}} placeholder='from their books…'
              defaultValue={x.manual_balance==null?'':x.manual_balance}
              onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}}
              onBlur={async e=>{
                const v=e.target.value.trim().replace(/[,$\s]/g,'');
                const cur=x.manual_balance==null?'':String(x.manual_balance);
                if(v===cur)return;
                if(v!==''&&!isFinite(Number(v))){setErr('Enter a number for '+x.account_code+'.');return;}
                try{await api.setIcManualBalance(x.mapping_id,(data.as_of||asOf)||null,v===''?null:Number(v));setErr('');await run();}
                catch(e2){setErr(e2.message);}}}/>
              :(x.manual_balance==null?<span style={{color:T.textDim}}>—</span>:fmt(x.manual_balance))}</td>
          <td style={{...S.tdR,width:120,fontWeight:700,color:x.difference==null?T.textDim:(Math.abs(x.difference)<0.01?T.green:T.red)}}>
            {x.difference==null?'':fmt(x.difference)}</td></tr>)}</tbody></table></details>}
    </>}
  </div>);
}

// ── External Entities TB ──
// Counterparties with no ledger in CloudLedger (JVs, outside ventures) still
// issue trial balances. Uploaded here per company + month, a TB stands in for
// the missing ledger: IC Mapping answers from its lines and the
// reconciliation compares against it like an in-ledger pair.
function ExternalTbPage({canEdit=true}){
  const[companies,setCompanies]=useState([]);
  const[tbs,setTbs]=useState([]);
  const[nodeId,setNodeId]=useState('');
  const[asOf,setAsOf]=useState(today());
  const[file,setFile]=useState(null);
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState('');const[msg,setMsg]=useState('');
  const[newCo,setNewCo]=useState('');const[showNew,setShowNew]=useState(false);
  const[view,setView]=useState(null);
  const[lines,setLines]=useState([]);
  const fileRef=useRef(null);
  const load=useCallback(async()=>{
    try{const[c,t]=await Promise.all([api.getIcCompanies(),api.getExternalTbs()]);
      setCompanies(c);setTbs(t);}catch(e){setErr(e.message);}},[]);
  useEffect(()=>{load();},[load]);
  const coName=id=>{const c=companies.find(x=>String(x.id)===String(id));return c?c.name:('Company '+id);};
  const upload=async()=>{
    if(!nodeId){setErr('Pick the external company the TB belongs to.');return;}
    if(!asOf){setErr('Set the as-of date (month end).');return;}
    if(!file){setErr('Choose the TB file (.xlsx or .csv).');return;}
    setBusy(true);setErr('');setMsg('');
    try{const r=await api.uploadExternalTb(Number(nodeId),asOf,file);
      setMsg(r.count+' lines saved for '+r.node_name+' as of '+r.as_of+'. Its intercompany accounts can now be mapped and reconciled.');
      setFile(null);if(fileRef.current)fileRef.current.value='';
      await load();}
    catch(e){setErr(e.message);}finally{setBusy(false);}};
  const registerCo=async()=>{
    const name=newCo.trim();if(!name)return;
    try{const r=await api.createIcCompany({name});setNewCo('');setShowNew(false);
      await load();setNodeId(String(r.id));
      setMsg(r.existing?(r.name+' was already registered — selected.'):('Registered '+name+'.'));}
    catch(e){setErr(e.message);}};
  const openLines=async t=>{
    if(view&&view.node_id===t.node_id&&view.as_of===t.as_of){setView(null);return;}
    try{setView(t);setLines(await api.getExternalTbLines(t.node_id,t.as_of));}
    catch(e){setErr(e.message);setView(null);}};
  // Companies that are actual CL entities keep their ledger here — no upload.
  const uploadable=companies.filter(c=>!c.entity_id);
  return(<div>
    <div style={S.h1}>External Entities TB</div>
    <div style={S.sub}>Monthly trial balances for counterparties that keep no ledger in CloudLedger.
      {' '}Once a TB is uploaded, IC Mapping pre-fills the counterparty account from its lines and the
      {' '}reconciliation compares against it — the recon reads the latest TB on or before its as-of date.</div>
    {err&&<div style={{...S.card,borderColor:T.red+'40'}}><div style={S.err}>{err}</div></div>}
    {msg&&<div style={{...S.card,borderColor:T.green+'40',padding:14}}><div style={S.success}>{msg}</div></div>}

    {canEdit&&<div style={{...S.card,padding:16,marginBottom:18}}>
      <div style={{display:'flex',gap:10,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div><label style={S.label}>External company</label>
          <select style={{...S.selectSm,minWidth:260}} value={nodeId} onChange={e=>setNodeId(e.target.value)}>
            <option value=''>Select a company…</option>
            {uploadable.map(c=><option key={c.id} value={String(c.id)}>{c.name}</option>)}</select></div>
        <button style={S.btnS} onClick={()=>setShowNew(!showNew)}>{showNew?'cancel':'+ Register new'}</button>
        <div><label style={S.label}>As of (month end)</label>
          <input style={S.inputSm} type='date' value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
        <div><label style={S.label}>Trial balance file</label>
          <input ref={fileRef} type='file' accept='.xlsx,.xls,.csv' style={{fontSize:12.5}}
            onChange={e=>setFile(e.target.files&&e.target.files[0]?e.target.files[0]:null)}/></div>
        <button style={S.btnP} onClick={upload} disabled={busy}>{busy?'Uploading…':'Upload TB'}</button></div>
      {showNew&&<div style={{display:'flex',gap:8,alignItems:'center',marginTop:10}}>
        <input style={{...S.inputSm,width:300}} placeholder='Company name — e.g. CPI/BYN South Mountain SFR Venture'
          value={newCo} onChange={e=>setNewCo(e.target.value)}
          onKeyDown={e=>{if(e.key==='Enter')registerCo();}}/>
        <button style={S.btnS} onClick={registerCo}>Register</button></div>}
      <div style={{fontSize:11.5,color:T.textDim,marginTop:10}}>
        Columns detected automatically: account number + name, then either Debit/Credit columns or one
        {' '}Balance column. Re-uploading the same company + date replaces that TB in full.</div></div>}

    <div style={S.cardFlush}>
      <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border}}>
        <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>Uploaded TBs ({tbs.length})</span></div>
      {tbs.length===0?<div style={{padding:24,color:T.textDim,fontSize:12.5,textAlign:'center'}}>
        Nothing uploaded yet. Upload a counterparty's monthly TB above to start reconciling against it.</div>
      :<table style={S.table}>
        <thead><tr><th style={S.th}>Company</th><th style={S.th}>As of</th>
          <th style={S.thR}>Lines</th><th style={S.th}>File</th><th style={S.th}>Uploaded</th>
          <th style={{...S.th,width:130}}></th></tr></thead>
        <tbody>{tbs.map(t=><Fragment key={t.node_id+'-'+t.as_of}>
          <tr>
            <td style={{...S.td,fontWeight:600,color:T.textBright,whiteSpace:'normal'}}>{t.node_name}</td>
            <td style={S.td}>{t.as_of}</td>
            <td style={S.tdR}>{t.line_count}</td>
            <td style={{...S.td,color:T.textMuted,fontSize:11.5,whiteSpace:'normal'}}>{t.filename||''}</td>
            <td style={{...S.td,color:T.textMuted,fontSize:11.5}}>{String(t.uploaded_at||'').slice(0,10)}{t.uploaded_by?' · '+t.uploaded_by:''}</td>
            <td style={{...S.td,textAlign:'right'}}>
              <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>openLines(t)}>
                {view&&view.node_id===t.node_id&&view.as_of===t.as_of?'Close':'View'}</button>
              {canEdit&&<button style={{...S.btnGhost,color:T.red,fontSize:11}}
                onClick={async()=>{if(!window.confirm('Delete the '+t.as_of+' TB for '+t.node_name+'? Mappings that point at it will show as missing until a TB is re-uploaded.'))return;
                  try{await api.deleteExternalTb(t.node_id,t.as_of);setMsg('Deleted.');await load();if(view&&view.node_id===t.node_id&&view.as_of===t.as_of)setView(null);}
                  catch(e){setErr(e.message);}}}>x</button>}</td></tr>
          {view&&view.node_id===t.node_id&&view.as_of===t.as_of&&<tr style={{background:T.accentDim}}>
            <td colSpan={6} style={{...S.td,whiteSpace:'normal',padding:0}}>
              <table style={S.table}><tbody>{lines.map(l=><tr key={l.account_code}>
                <td style={{...S.td,fontWeight:600,color:T.textBright,width:100}}>{l.account_code}</td>
                <td style={{...S.td,whiteSpace:'normal',color:T.textMuted}}>{l.account_name}
                  <span style={{marginLeft:6,fontSize:10,color:T.textDim}}>{l.type}</span></td>
                <td style={{...S.tdR,width:150}}>{fmt(l.balance)}</td></tr>)}</tbody></table></td></tr>}
        </Fragment>)}</tbody></table>}</div>
  </div>);
}

// Pick a counterparty account from an UPLOADED trial balance — the TB
// equivalent of the ledger-backed account picker.
function IcTbAccountSelect({nodeId,asOf,value,onChange,width=300}){
  const[lines,setLines]=useState(null);
  useEffect(()=>{let ok=true;
    api.getExternalTbLines(nodeId,asOf).then(l=>{if(ok)setLines(l);}).catch(()=>{if(ok)setLines([]);});
    return()=>{ok=false;};},[nodeId,asOf]);
  if(lines===null)return<span style={{color:T.textDim,fontSize:12}}>reading the uploaded TB…</span>;
  if(!lines.length)return<span style={{color:T.textDim,fontSize:12}}>the uploaded TB has no lines</span>;
  return<select style={{...S.selectSm,maxWidth:width}} value={value} onChange={e=>onChange(e.target.value)}>
    <option value=''>— pick from the uploaded TB —</option>
    {lines.map(l=><option key={l.account_code} value={l.account_code}>
      {l.account_code+' '+(l.account_name||'')+' ('+fmt(l.balance)+')'}</option>)}</select>;
}

// ── IC Mapping: the setup page a person owns ──
// Account names alone can't be trusted (CloudLedger copies the same chart of
// accounts across entities, so an entity can own an account named "Due from
// <itself>"). Suggestions read the names and propose; a person confirms.
function IntercompanyMapping({entities,activeEntity,canEdit=true}){
  const[tab,setTab]=useState('mappings');
  const[eid,setEid]=useState(activeEntity||'');
  const[rows,setRows]=useState([]);
  const[loading,setLoading]=useState(false);const[err,setErr]=useState('');const[msg,setMsg]=useState('');
  const[showAdd,setShowAdd]=useState(false);
  const[showUnmapped,setShowUnmapped]=useState(false);
  const[showMapped,setShowMapped]=useState(false);
  const[showExternal,setShowExternal]=useState(false);
  // Bumped by every change to the mappings, so the unmapped panel reloads too.
  const[mapVersion,setMapVersion]=useState(0);
  const[form,setForm]=useState({account_code:'',account_name:'',ic_type:'due_from',counterparty_entity_id:'',counterparty_node_id:'',counterparty_account_code:'',is_external:false,notes:''});
  const[companies,setCompanies]=useState([]);
  const[people,setPeople]=useState([]);
  // Sorted by the name shown, not by the entity code the API sorts on — see the
  // BROZ FUND I LLC case (code "1000", so it sorted above every other entity).
  const entityOpts=[...(entities||[])].sort((a,b)=>String(a.name).localeCompare(String(b.name),undefined,{sensitivity:'base'}));
  const offLedgerCompanies=companies.filter(c=>c.entity_id==null);
  // An external mapping is a decision already made: the counterparty is outside
  // the group, CloudLedger holds no ledger for it, and the reconciliation
  // excludes it by design. Leaving them in the working list buries the rows that
  // still need attention — on Banyan Residential they are 66 of 89. They are
  // hidden by default, never deleted, and one click away.
  // External CAPITAL is investor money — complete by rule, shown under Show
  // mapped with the investor-capital tag — so it stays out of this list.
  const externalCount=rows.filter(r=>r.is_external&&r.ic_type!=='contributed_capital').length;
  const externalRows=rows.filter(r=>r.is_external&&r.ic_type!=='contributed_capital');

  const entName=id=>{const e=(entities||[]).find(x=>x.id===Number(id));return e?e.name:'';};
  const loadCompanies=useCallback(async()=>{try{setCompanies(await api.getIcCompanies());}catch(e){console.error('[ic] companies:',e.message);}},[]);
  const loadPeople=useCallback(async()=>{try{setPeople(await api.getIcPeople());}catch(e){console.error('[ic] people:',e.message);}},[]);

  const loadRows=useCallback(async()=>{
    if(!eid){setRows([]);return;}
    setLoading(true);setErr('');
    try{setRows(await api.getIcMappings({entity_id:eid}));}catch(e){setErr(e.message);}finally{setLoading(false);}
  },[eid]);
  useEffect(()=>{loadRows();},[loadRows]);
  useEffect(()=>{loadCompanies();},[loadCompanies]);
  // Every mutation goes through this, not through loadRows directly: anything
  // that changes the mappings also invalidates the unmapped panel.
  const refreshMappings=useCallback(()=>{loadRows();setMapVersion(v=>v+1);},[loadRows]);
  useEffect(()=>{loadPeople();},[loadPeople]);

  const add=async()=>{
    if(!eid){setErr('Pick an entity first.');return;}
    try{await api.createIcMapping({...form,entity_id:Number(eid),
      counterparty_entity_id:form.is_external?null:(form.counterparty_entity_id?Number(form.counterparty_entity_id):null),
      counterparty_node_id:form.is_external?null:(form.counterparty_node_id?Number(form.counterparty_node_id):null),
      counterparty_account_code:form.is_external?null:(form.counterparty_account_code||null),
      is_external:form.is_external?1:0});
      setShowAdd(false);setForm({account_code:'',account_name:'',ic_type:'due_from',counterparty_entity_id:'',counterparty_node_id:'',counterparty_account_code:'',is_external:false,notes:''});
      setErr('');refreshMappings();}catch(e){setErr(e.message);}
  };
  const del=async r=>{if(!confirm('Remove the mapping for '+r.account_code+' '+(r.account_name||'')+'?'))return;
    try{await api.deleteIcMapping(r.id);refreshMappings();}catch(e){setErr(e.message);}};


  // Counterparty = a CloudLedger entity, a registered company (a name on the org
  // charts with no ledger here), or "external". Encoded as 'e<id>' / 'n<id>' /
  // '__ext' so one <select> can carry all three without a second control.
  // '__new' registers a company inline: leaving the page to create a QOZB before
  // a mapping can be finished is the wrong shape for this job.
  // ONE callback carrying the whole choice. The three fields are mutually
  // exclusive, so they must be written together — setting them through separate
  // setters meant the second call undid the first. The control itself is
  // IcCpSelect, shared with the reconciliation page's inline mapping form.
  const cpCell=(entId,nodeId,isExt,onPick,seedName)=>(
    <IcCpSelect entityOpts={entityOpts} companies={companies}
      entId={entId} nodeId={nodeId} isExt={isExt} onPick={onPick} seedName={seedName}
      onCompanyAdded={loadCompanies} onMsg={setMsg} onErr={setErr}/>);

  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8,flexWrap:'wrap',gap:12}}>
      <div><div style={S.h1}>IC Mapping</div>
        <div style={S.sub}>Tells the reconciliation which account faces which entity. Account names are only a suggestion — a person confirms every row.</div></div></div>

    <div style={{display:'flex',gap:4,marginBottom:16,borderBottom:'1px solid '+T.border}}>
      {[['mappings','Account mappings'],['companies','Companies ('+offLedgerCompanies.length+')']].map(([k,l])=>
        <button key={k} onClick={()=>{setTab(k);setErr('');setMsg('');}} style={{background:'none',border:'none',borderBottom:'2px solid '+(tab===k?T.accent:'transparent'),color:tab===k?T.accent:T.textMuted,fontWeight:tab===k?700:500,fontSize:13,padding:'9px 16px',cursor:'pointer',marginBottom:-1}}>{l}</button>)}</div>

    {err&&<div style={{...S.card,borderColor:T.red+'40',padding:14}}><div style={S.err}>{err}</div></div>}
    {msg&&<div style={{...S.card,borderColor:T.green+'40',padding:14}}><div style={S.success}>{msg}</div></div>}

    {tab==='mappings'&&<>
      <div style={{...S.row,alignItems:'flex-end'}}>
        <div style={{minWidth:280}}><label style={S.label}>Entity</label>
          <select style={S.select} value={eid} onChange={e=>{setEid(e.target.value);setMsg('');}}>
            <option value=''>Select an entity…</option>
            {entityOpts.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select></div>
        {eid&&externalCount>0&&<button style={S.btnS} title='Accounts facing a party outside the group. There is no ledger on the other side, so there is nothing to map to.'
          onClick={()=>setShowExternal(v=>!v)}>
          {showExternal?'Hide external ('+externalCount+')':'Show external ('+externalCount+')'}</button>}
        {canEdit&&eid&&<><button style={showUnmapped?{...S.btnS,borderColor:T.accent,color:T.accent}:S.btnS}
            onClick={()=>{setShowUnmapped(v=>!v);setErr('');setMsg('');}}>
            Show unmapped accounts</button>
          <button style={showMapped?{...S.btnS,borderColor:T.accent,color:T.accent}:S.btnS}
            onClick={()=>{setShowMapped(v=>!v);setErr('');setMsg('');}}>
            Show mapped</button>
          <button style={S.btnP} onClick={()=>{setShowAdd(!showAdd);setErr('');}}>{showAdd?'Cancel':'+ Add mapping'}</button></>}</div>

      {eid&&showMapped&&<IcMappedAccounts entityId={eid} entityName={entName(eid)} canEdit={canEdit}
        reloadKey={mapVersion} onMapped={loadRows} onMsg={setMsg} onErr={setErr}/>}

      {canEdit&&eid&&showUnmapped&&<IcUnmappedAccounts entityId={eid} entityName={entName(eid)}
        entityOpts={entityOpts} companies={companies} canEdit={canEdit} reloadKey={mapVersion}
        onCompanyAdded={loadCompanies} onMapped={loadRows} onMsg={setMsg} onErr={setErr}/>}

      {showAdd&&<div style={{...S.card,borderColor:T.green+'40'}}>
        <div style={S.row}>
          <div style={S.col}><label style={S.label}>Account code</label><input style={S.input} value={form.account_code} onChange={e=>setForm(f=>({...f,account_code:e.target.value}))} placeholder='18307'/></div>
          <div style={{...S.col,flex:2}}><label style={S.label}>Account name</label><input style={S.input} value={form.account_name} onChange={e=>setForm(f=>({...f,account_name:e.target.value}))} placeholder='Due from …'/></div>
          <div style={S.col}><label style={S.label}>Kind</label>
            <select style={S.select} value={form.ic_type} onChange={e=>setForm(f=>({...f,ic_type:e.target.value}))}>
              {Object.keys(IC_TYPE_LABEL).map(k=><option key={k} value={k}>{IC_TYPE_LABEL[k]}</option>)}</select></div>
          <div style={{...S.col,flex:2}}><label style={S.label}>Counterparty</label>
            {cpCell(form.counterparty_entity_id,form.counterparty_node_id,form.is_external,
              p=>setForm(f=>({...f,counterparty_entity_id:p.entity_id||'',counterparty_node_id:p.node_id||'',
                is_external:!!p.is_external,counterparty_account_code:''})),
              form.account_name)}</div></div>
        {!form.is_external&&form.counterparty_entity_id&&form.account_code&&
          <div style={{marginBottom:12}}><label style={S.label}>Their account</label>
            <div style={{display:'flex',flexWrap:'wrap',alignItems:'center',gap:8}}>
              <IcCpAccountSelect entityId={eid} accountCode={form.account_code}
                counterpartyEntityId={form.counterparty_entity_id} value={form.counterparty_account_code}
                onChange={v=>setForm(f=>({...f,counterparty_account_code:v}))}
                onLoaded={r=>setForm(f=>f.counterparty_account_code?f:({...f,counterparty_account_code:r.best||''}))}/>
            </div></div>}
        <div style={{marginBottom:12}}><label style={S.label}>Notes</label><input style={S.input} value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder='(optional — e.g. the outside party name)'/></div>
        <button style={S.btnP} onClick={add}>Add mapping</button></div>}


      {!eid&&<IcEmpty>Pick an entity to see and edit its intercompany account mappings.</IcEmpty>}

      {/* The page keeps no standing table of its own. The unmapped worklist and
          Show mapped are the same accounts split by the only question that
          matters — is the pair complete — so a third copy of the rows under
          them was noise. Externals live behind their own toggle: no ledger on
          the other side, nothing to map to, never part of either list. */}
      {eid&&!showUnmapped&&!showMapped&&!showExternal&&!showAdd&&
        <IcEmpty><b>{entName(eid)}</b>: {rows.length-externalCount} mapping{(rows.length-externalCount)===1?'':'s'} facing the group{externalCount>0?', '+externalCount+' external':''}.<br/>
          <span style={{fontSize:12}}>Open <b>Show unmapped accounts</b> to work through what still needs mapping,
          {' '}or <b>Show mapped</b> to review the finished pairs.</span></IcEmpty>}

      {eid&&showExternal&&<div style={{...S.cardFlush,border:'1px solid '+T.border,overflow:'hidden'}}>
        <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border}}>
          <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>External mappings ({externalRows.length})</span>
          <span style={{color:T.textMuted,fontSize:12,marginLeft:8}}>These face a party outside the group. The reconciliation reports them separately and never nets them.</span></div>
        {externalRows.length===0?<div style={{padding:'18px 16px',textAlign:'center',color:T.textDim}}>None.</div>
        :<div className="cl-scroll" style={{overflow:'auto',maxHeight:400}}><table style={S.table}><tbody>
          {externalRows.map(r=><tr key={r.id}>
            <td style={{...S.td,whiteSpace:'normal'}}><span style={{fontWeight:600,color:T.textBright}}>{r.account_code}</span> <span style={{color:T.textMuted}}>{r.account_name}</span></td>
            <td style={{...S.td,color:T.textMuted,whiteSpace:'normal'}}>{r.notes||''}</td>
            {canEdit&&<td style={{...S.td,textAlign:'right',width:60}}>
              <button style={{...S.btnGhost,color:T.red,fontSize:11}} title='Remove the mapping. The account returns to the unmapped worklist.'
                onClick={()=>del(r)}>x</button></td>}</tr>)}
        </tbody></table></div>}</div>}
    </>}

    {tab==='companies'&&<>
      <div style={{marginBottom:16,color:T.textMuted,fontSize:13}}>
        Companies that appear on the org charts as counterparties but keep no ledger in CloudLedger — QOZBs, joint ventures, sponsor holding companies. Register one here and it becomes selectable as a counterparty everywhere, and account names that mention it start matching to it automatically.
        {canEdit&&<button style={{...S.btnP,marginLeft:12}} onClick={async()=>{
          const name=prompt('Company name:');if(!name||!name.trim())return;
          try{const r=await api.createIcCompany({name:name.trim()});await loadCompanies();
            setErr('');setMsg(r.existing?('"'+r.name+'" was already registered.'):('Registered "'+r.name+'".'));
          }catch(e){setErr(e.message);}}}>+ Register company</button>}</div>

      {offLedgerCompanies.length===0?<IcEmpty>No off-ledger companies registered yet. When a mapping names a company CloudLedger doesn’t have, register it from the counterparty picker (“+ Register a new company…”).</IcEmpty>
      :<div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
        <th style={S.th}>Company</th><th style={S.th}>Kind</th><th style={S.th}>Notes</th></tr></thead>
      <tbody>{offLedgerCompanies.map(c=><tr key={c.id}>
        <td style={{...S.td,fontWeight:600,color:T.textBright}}>{c.name} <span style={{marginLeft:6}}><IcBadge kind="info">no ledger</IcBadge></span></td>
        <td style={S.td}>{c.node_type||'shell'}</td>
        <td style={{...S.td,color:T.textMuted,whiteSpace:'normal'}}>{c.notes||''}</td></tr>)}</tbody></table></div>}

      {people.length>0&&<div style={{marginTop:24}}>
        <div style={{fontWeight:700,color:T.textBright,fontSize:13,marginBottom:6}}>Marked as people ({people.length})</div>
        <div style={{color:T.textMuted,fontSize:12,marginBottom:10}}>These names are never proposed as counterparties. Individuals on capital accounts are detected automatically; these were marked by hand, which is how a person on a due-from or due-to account is handled.</div>
        <div style={{...S.cardFlush,overflow:'hidden'}}><table style={S.table}><tbody>
          {people.map(p=><tr key={p.id}>
            <td style={{...S.td,fontWeight:600,color:T.textBright}}>{p.name}</td>
            {canEdit&&<td style={{...S.td,textAlign:'right',width:110}}>
              <button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={async()=>{
                try{await api.unmarkIcPerson(p.id);await loadPeople();setMsg('"'+p.name+'" is no longer marked a person.');}
                catch(e){setErr(e.message);}}}>Undo</button></td>}</tr>)}
        </tbody></table></div></div>}

      {companies.some(c=>c.entity_id!=null)&&<div style={{color:T.textMuted,fontSize:12,marginTop:12}}>
        {companies.filter(c=>c.entity_id!=null).length} more compan{companies.filter(c=>c.entity_id!=null).length===1?'y is':'ies are'} recorded on the Org Structure page and backed by a CloudLedger entity, so they are offered under “CloudLedger entities” instead.</div>}
    </>}

  </div>);
}

// ═════════════════════════ Org Structure ═════════════════════════
// The ownership tree from the legal org charts, and the tie-out it makes
// possible: a fund's investment account against the contributed capital of the
// company it funded, looked through the holding companies in between.
//
// Nodes come in two kinds. A node with an entity is a CloudLedger ledger; a
// node without one is a shell that holds an investment balance but keeps no
// ledger here (CLRFI CLIP Sponsor, County Line SRN, CLRFI Midco II). The
// shells are why a plain parent column on the entities table wouldn't do: drop
// them and the 76.51% / 54.53% steps that create the minority interest vanish.

const ORG_NODE_TYPE_LABEL={fund:'Fund',holdco:'Holding co.',company:'Company',property_owner:'Property owner',shell:'Shell (no ledger)'};

function OrgPctBadge({pct}){
  if(pct==null)return null;
  const full=Math.abs(pct-100)<0.00005;
  return<span style={{display:'inline-block',padding:'1px 7px',borderRadius:20,fontSize:10,fontWeight:700,
    color:full?T.textMuted:T.orange,background:full?T.bgElevated:T.orangeDim,whiteSpace:'nowrap'}}>{Number(pct).toFixed(2)}%</span>;
}

function OrgTreeRows({node,onEdit,canEdit,depth=0}){
  if(!node)return null;
  const shell=node.entity_id==null;
  return(<>
    <tr style={shell?{background:T.bgElevated}:null}>
      <td style={{...S.td,whiteSpace:'nowrap'}}>
        <span style={{display:'inline-block',width:depth*22}}/>
        {depth>0&&<span style={{color:T.textDim,marginRight:6}}>{'└'}</span>}
        <span style={{fontWeight:shell?500:600,color:shell?T.textMuted:T.textBright,fontStyle:shell?'italic':'normal'}}>{node.name}</span>
        {shell&&<span style={{marginLeft:8}}><IcBadge kind="mute">no ledger</IcBadge></span>}
      </td>
      <td style={S.td}>{ORG_NODE_TYPE_LABEL[node.node_type]||node.node_type}</td>
      <td style={S.tdR}>{depth===0?<span style={{color:T.textDim}}>{'—'}</span>:<OrgPctBadge pct={node.ownership_pct}/>}</td>
      <td style={{...S.tdR,fontWeight:600}}>{Number(node.effective_pct).toFixed(2)}%</td>
      <td style={{...S.tdR,color:node.nci_pct>0.00005?T.orange:T.textDim,fontWeight:node.nci_pct>0.00005?600:400}}>
        {node.nci_pct>0.00005?Number(node.nci_pct).toFixed(2)+'%':'—'}</td>
      <td style={{...S.td,color:T.textMuted,whiteSpace:'normal',fontSize:12}}>{node.notes||''}</td>
      {canEdit&&<td style={S.td}><button style={{...S.btnGhost,color:T.accent,fontSize:11}} onClick={()=>onEdit(node)}>Edit</button></td>}
    </tr>
    {(node.children||[]).map(c=><OrgTreeRows key={c.id+'-'+c.edge_id} node={c} onEdit={onEdit} canEdit={canEdit} depth={depth+1}/>)}
  </>);
}

function OrgStructurePage({entities,canEdit=true}){
  const[tab,setTab]=useState('structure');
  const[roots,setRoots]=useState([]);const[rootId,setRootId]=useState('');
  const[nodes,setNodes]=useState([]);const[tree,setTree]=useState(null);const[cycles,setCycles]=useState([]);
  const[asOf,setAsOf]=useState('2026-06-30');
  const[recon,setRecon]=useState(null);
  const[loading,setLoading]=useState(false);const[err,setErr]=useState('');const[msg,setMsg]=useState('');
  const[edit,setEdit]=useState(null);// node being edited
  const[addUnder,setAddUnder]=useState(null);
  const[form,setForm]=useState({name:'',entity_id:'',node_type:'company',notes:'',ownership_pct:'100'});
  const[open,setOpen]=useState({});
  // By displayed name, not by entity code — see the note in IntercompanyMapping.
  const entityOpts=[...(entities||[])].sort((a,b)=>String(a.name).localeCompare(String(b.name),undefined,{sensitivity:'base'}));

  const loadRoots=useCallback(async()=>{
    try{const d=await api.getOrgStructure();setRoots(d.roots||[]);setNodes(d.nodes||[]);
      if(!rootId&&d.roots&&d.roots.length)setRootId(String(d.roots[0].id));
    }catch(e){setErr(e.message);}
  },[rootId]);
  useEffect(()=>{loadRoots();},[]);// eslint-disable-line react-hooks/exhaustive-deps

  const loadTree=useCallback(async()=>{
    if(!rootId)return;
    setLoading(true);setErr('');
    try{
      const t=await api.getOrgTree(rootId);
      setTree(t.tree);setCycles(t.cycles||[]);
    }catch(e){setErr(e.message);setTree(null);}
    finally{setLoading(false);}
  },[rootId]);
  useEffect(()=>{loadTree();},[loadTree]);

  const runRecon=useCallback(async()=>{
    if(!rootId)return;
    setLoading(true);setErr('');
    try{setRecon(await api.reconcileOrgInvestments(rootId,asOf));}
    catch(e){setErr(e.message);setRecon(null);}
    finally{setLoading(false);}
  },[rootId,asOf]);
  useEffect(()=>{if(tab==='tieout'&&rootId&&!recon)runRecon();},[tab,rootId]);// eslint-disable-line react-hooks/exhaustive-deps

  const seed=async()=>{
    setErr('');setMsg('');
    try{const r=await api.seedOrgClrf();
      setMsg(r.created?('Loaded the County Line Rail Fund I chart — '+r.nodes+' nodes.'):'That chart is already loaded.');
      const d=await api.getOrgStructure();setRoots(d.roots||[]);setNodes(d.nodes||[]);
      setRootId(String(r.root_node_id));
    }catch(e){setErr(e.message);}
  };

  const saveNode=async()=>{
    setErr('');
    try{
      if(edit){
        await api.saveOrgNode(edit.id,{name:form.name,entity_id:form.entity_id?Number(form.entity_id):null,
          node_type:form.node_type,notes:form.notes||null,sort_order:edit.sort_order||0});
        if(edit.edge_id!=null)await api.saveOrgEdge(edit.edge_id,{parent_node_id:edit.parent_node_id_for_edit,
          child_node_id:edit.id,ownership_pct:Number(form.ownership_pct)});
      }else if(addUnder){
        const r=await api.createOrgNode({name:form.name,entity_id:form.entity_id?Number(form.entity_id):null,
          node_type:form.node_type,notes:form.notes||null});
        await api.createOrgEdge({parent_node_id:addUnder.id,child_node_id:r.id,ownership_pct:Number(form.ownership_pct)});
      }
      setEdit(null);setAddUnder(null);setRecon(null);
      await loadRoots();await loadTree();
    }catch(e){setErr(e.message);}
  };
  const removeNode=async()=>{
    if(!edit)return;
    if(!confirm('Remove "'+edit.name+'" from the structure? Everything under it is detached too. The CloudLedger entity itself is not touched.'))return;
    try{await api.deleteOrgNode(edit.id);setEdit(null);setRecon(null);await loadRoots();await loadTree();}
    catch(e){setErr(e.message);}
  };

  const startEdit=n=>{
    // Find this node's parent so the ownership % on its incoming edge is editable.
    let parentId=null;
    const walk=(p)=>{for(const c of (p.children||[])){if(c.id===n.id&&c.edge_id===n.edge_id)parentId=p.id;walk(c);}};
    if(tree)walk(tree);
    setAddUnder(null);
    setEdit({...n,parent_node_id_for_edit:parentId});
    setForm({name:n.name,entity_id:n.entity_id||'',node_type:n.node_type,notes:n.notes||'',
      ownership_pct:String(n.ownership_pct==null?100:n.ownership_pct)});
  };
  const startAdd=()=>{
    const flat=[];const walk=n=>{flat.push(n);(n.children||[]).forEach(walk);};if(tree)walk(tree);
    setEdit(null);setAddUnder(tree);
    setForm({name:'',entity_id:'',node_type:'company',notes:'',ownership_pct:'100'});
  };

  const flat=(()=>{const out=[];const walk=n=>{out.push(n);(n.children||[]).forEach(walk);};if(tree)walk(tree);return out;})();
  const statusBadge=s=>s==='ties_on_chain'||s==='ties_on_total'?<IcBadge kind="ok">ties</IcBadge>
    :s==='ties_rounding'?<IcBadge kind="ok">ties (rounding)</IcBadge>
    :s==='not_in_tree'?<IcBadge kind="mute">not in structure</IcBadge>
    :<IcBadge kind="bad">difference</IcBadge>;

  const doExport=()=>{
    if(!recon)return;
    const d=[['Org structure — fund investment vs. subsidiary capital'],
      [(recon.root&&recon.root.name)||'','as of '+(recon.as_of||asOf)],[],
      ['Investor','Account','Account name','Names','Compared against','Chain','Investment','Capital on this chain','Total capital','Difference (chain)','Difference (total)','Status']];
    recon.rows.forEach(r=>d.push([r.investor_name,r.account_code,r.account_name,r.named_name||'',r.subsidiary_name||'',
      (r.chain||[]).map(c=>c.name+(c.is_shell?' [shell]':'')).join(' > '),
      r.investment,r.chain_capital,r.total_capital,r.chain_difference,r.total_difference,r.status]));
    d.push([]);d.push(['Non-controlling interest (indicative, from GL equity)']);
    d.push(['Entity','Ownership %','Effective %','NCI %','Entity equity','NCI equity']);
    recon.nci.forEach(n=>d.push([n.name,n.ownership_pct,n.effective_pct,n.nci_pct,n.entity_equity,n.nci_equity]));
    d.push(['Total','','','','',recon.totals.nci_equity_total]);
    exportToExcel(d,'Org_Structure_Tieout_'+String(recon.as_of||asOf).replace(/-/g,'')+'.xlsx');
  };

  return(<div>
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8,flexWrap:'wrap',gap:12}}>
      <div><div style={S.h1}>Org Structure</div>
        <div style={S.sub}>Who owns whom, from the legal org charts — including holding companies that hold investment balances but keep no ledger here.</div></div>
      <div style={{display:'flex',gap:8,alignItems:'flex-end',flexWrap:'wrap'}}>
        <div><label style={S.label}>Structure</label>
          <select style={{...S.selectSm,minWidth:220}} value={rootId} onChange={e=>{setRootId(e.target.value);setRecon(null);}}>
            <option value=''>Select…</option>
            {roots.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}</select></div>
        {tab==='tieout'&&<>
          <div><label style={S.label}>As of</label>
            <input style={S.inputSm} type='date' value={asOf} onChange={e=>setAsOf(e.target.value)}/></div>
          <button style={S.btnP} onClick={runRecon} disabled={!rootId||loading}>{loading?'Running…':'Run'}</button>
          {recon&&<button style={S.btnExport} onClick={doExport}>Export Excel</button>}</>}
      </div></div>

    <div style={{display:'flex',gap:4,marginBottom:16,borderBottom:'1px solid '+T.border}}>
      {[['structure','Ownership tree'],['tieout','Investment tie-out']].map(([k,l])=>
        <button key={k} onClick={()=>{setTab(k);setErr('');setMsg('');}} style={{background:'none',border:'none',borderBottom:'2px solid '+(tab===k?T.accent:'transparent'),color:tab===k?T.accent:T.textMuted,fontWeight:tab===k?700:500,fontSize:13,padding:'9px 16px',cursor:'pointer',marginBottom:-1}}>{l}</button>)}</div>

    {err&&<div style={{...S.card,borderColor:T.red+'40',padding:14}}><div style={S.err}>{err}</div></div>}
    {msg&&<div style={{...S.card,borderColor:T.green+'40',padding:14}}><div style={S.success}>{msg}</div></div>}
    {cycles&&cycles.length>0&&<div style={{...S.card,borderColor:T.orange+'40',padding:14}}>
      <div style={{color:T.orange,fontSize:12,fontWeight:600}}>{cycles.length} circular ownership link{cycles.length===1?'':'s'} were skipped while drawing this tree.</div></div>}

    {!roots.length&&!err&&<div style={{...S.card,textAlign:'center',padding:40}}>
      <div style={{color:T.textDim,marginBottom:14}}>No ownership structure recorded yet.</div>
      {canEdit&&<button style={S.btnP} onClick={seed}>Load the County Line Rail Fund I chart</button>}
      <div style={{color:T.textMuted,fontSize:12,marginTop:10}}>Creates the CLRF I → Midco I chain from the 4/13/2026 org chart, including the four holding companies that have no ledger in CloudLedger.</div></div>}

    {tab==='structure'&&tree&&<>
      {canEdit&&<div style={{marginBottom:12,display:'flex',gap:8}}>
        <button style={S.btnS} onClick={startAdd}>+ Add company</button>
        {!roots.some(r=>r.name==='County Line Rail Fund I, LP')&&<button style={S.btnS} onClick={seed}>Load CLRF I chart</button>}</div>}

      {(edit||addUnder)&&<div style={{...S.card,borderColor:T.green+'40'}}>
        <div style={{fontWeight:700,color:T.textBright,marginBottom:12}}>{edit?'Edit '+edit.name:'Add a company'}</div>
        <div style={S.row}>
          <div style={{...S.col,flex:2}}><label style={S.label}>Name</label>
            <input style={S.input} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder='CLRFI CLIP Sponsor, LLC'/></div>
          <div style={{...S.col,flex:2}}><label style={S.label}>CloudLedger entity</label>
            <select style={S.select} value={form.entity_id} onChange={e=>setForm(f=>({...f,entity_id:e.target.value,node_type:e.target.value?(f.node_type==='shell'?'company':f.node_type):'shell'}))}>
              <option value=''>None — shell with no ledger</option>
              {entityOpts.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}</select></div>
          <div style={S.col}><label style={S.label}>Kind</label>
            <select style={S.select} value={form.node_type} onChange={e=>setForm(f=>({...f,node_type:e.target.value}))}>
              {Object.keys(ORG_NODE_TYPE_LABEL).map(k=><option key={k} value={k}>{ORG_NODE_TYPE_LABEL[k]}</option>)}</select></div>
          <div style={S.col}><label style={S.label}>Owned %</label>
            <input style={{...S.input,textAlign:'right'}} value={form.ownership_pct} onChange={e=>setForm(f=>({...f,ownership_pct:e.target.value}))} placeholder='100'/></div></div>
        {addUnder&&<div style={{marginBottom:12}}><label style={S.label}>Owned by</label>
          <select style={S.select} value={addUnder.id} onChange={e=>{const id=Number(e.target.value);setAddUnder(flat.find(n=>n.id===id)||tree);}}>
            {flat.map(n=><option key={n.id+'-'+n.edge_id} value={n.id}>{' '.repeat(n.depth*3)}{n.name}</option>)}</select></div>}
        <div style={{marginBottom:12}}><label style={S.label}>Notes</label>
          <input style={S.input} value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} placeholder='(optional)'/></div>
        <div style={{display:'flex',gap:8}}>
          <button style={S.btnP} onClick={saveNode}>{edit?'Save':'Add'}</button>
          <button style={S.btnS} onClick={()=>{setEdit(null);setAddUnder(null);setErr('');}}>Cancel</button>
          {edit&&<button style={{...S.btnD,marginLeft:'auto'}} onClick={removeNode}>Remove from structure</button>}</div></div>}

      <div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
        <th style={S.th}>Company</th><th style={S.th}>Kind</th><th style={S.thR}>Owned %</th>
        <th style={S.thR}>Effective %</th><th style={S.thR}>Minority %</th><th style={S.th}>Notes</th>
        {canEdit&&<th style={{...S.th,width:70}}></th>}</tr></thead>
      <tbody><OrgTreeRows node={tree} onEdit={startEdit} canEdit={canEdit}/></tbody></table></div>

      <div style={{color:T.textMuted,fontSize:12,marginTop:-8}}>
        <b>Effective %</b> multiplies the ownership steps from the top of the structure down, so a 76.51% step reduces everything beneath it. <b>Minority %</b> is the rest — the share owned by someone outside this structure.</div>
    </>}

    {tab==='tieout'&&<>
      {loading&&<div style={{textAlign:'center',padding:40,color:T.textMuted}}>Reading balances…</div>}
      {!loading&&recon&&<>
        <div style={{...S.row,marginBottom:20}}>
          <IcTile label='Investments tied' value={recon.totals.ties+' of '+recon.rows.length} kind={recon.totals.differences?'warn':'ok'}/>
          <IcTile label='Investment total' value={fmt(recon.totals.investment_total)} kind='info'/>
          <IcTile label='Unexplained difference' value={fmt(recon.totals.abs_difference)} kind={recon.totals.abs_difference>1?'bad':'ok'}/>
          <IcTile label='Minority interest' value={fmt(recon.totals.nci_equity_total)} kind='warn'/>
          <IcTile label='Shells in the chain' value={String(recon.totals.shell_count)} kind='info'/></div>

        {recon.rows.length===0?<IcEmpty>No investment accounts are mapped on the entities in this structure. Map them on the <b>IC Mapping</b> page first.</IcEmpty>
        :<div className="cl-scroll" style={scrollBox()}><table style={S.table}><thead><tr>
          <th style={{...S.th,width:28}}></th><th style={S.th}>Investor</th><th style={S.th}>Account</th>
          <th style={S.th}>Compared against</th><th style={S.thR}>Investment</th>
          <th style={S.thR}>Capital on this chain</th><th style={S.thR}>Total capital</th>
          <th style={S.thR}>Difference</th><th style={S.th}>Status</th></tr></thead>
        <tbody>{recon.rows.map(r=>{const k=r.investor_entity_id+'-'+r.account_code;const isOpen=!!open[k];
          const bad=r.status==='difference'||r.status==='not_in_tree';
          return<Fragment key={k}>
          <tr style={{cursor:'pointer',background:bad?T.redDim:'transparent'}} onClick={()=>setOpen(o=>({...o,[k]:!o[k]}))}>
            <td style={{...S.td,color:T.textDim,textAlign:'center'}}>{isOpen?'▼':'▶'}</td>
            <td style={{...S.td,fontWeight:600,color:T.textBright}}>{r.investor_name}</td>
            <td style={S.td}><span style={{fontWeight:600,color:T.textBright}}>{r.account_code}</span> <span style={{color:T.textMuted}}>{r.account_name}</span></td>
            <td style={S.td}>{r.subsidiary_name||<span style={{color:T.textDim}}>{'—'}</span>}
              {r.retargeted&&<span style={{marginLeft:6}}><IcBadge kind="info">re-aimed</IcBadge></span>}</td>
            <td style={{...S.tdR,fontWeight:600}}>{fmt(r.investment)}</td>
            <td style={S.tdR}>{fmt(r.chain_capital)}</td>
            <td style={S.tdR}>{fmt(r.total_capital)}</td>
            <td style={{...S.tdR,fontWeight:700,color:bad?T.red:T.green}}>{fmt(r.best_difference==null?r.total_difference:r.best_difference)}</td>
            <td style={S.td}>{statusBadge(r.status)}</td></tr>
          {isOpen&&<tr><td colSpan={9} style={{padding:'12px 16px 14px 40px',background:T.bgElevated,borderBottom:'1px solid '+T.border,whiteSpace:'normal'}}>
            {r.message&&<div style={{color:T.textMuted,fontSize:12,marginBottom:10}}>{r.message}</div>}
            {r.retarget_note&&<div style={{color:T.textMuted,fontSize:12,marginBottom:10}}>{r.retarget_note}</div>}
            {r.chain&&r.chain.length>0&&<div style={{marginBottom:10,fontSize:12}}>
              <span style={{color:T.textMuted,fontWeight:600,marginRight:6}}>Chain:</span>
              {r.chain.map((c,i)=><span key={c.id}>{i>0&&<span style={{color:T.textDim,margin:'0 6px'}}>{'›'}</span>}
                <span style={{color:c.is_shell?T.textMuted:T.textBright,fontStyle:c.is_shell?'italic':'normal'}}>{c.name}</span>
                {c.is_shell&&<span style={{color:T.textDim,fontSize:11}}> (no ledger)</span>}
                {Math.abs(c.ownership_pct-100)>0.00005&&<span style={{marginLeft:5}}><OrgPctBadge pct={c.ownership_pct}/></span>}</span>)}</div>}
            {r.chain_legs&&r.chain_legs.length>0&&<div style={{marginBottom:8}}>
              <div style={{fontSize:11,fontWeight:700,color:T.textMuted,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:4}}>Capital from this chain</div>
              {r.chain_legs.map(c=><div key={c.account_code} style={{fontSize:12,padding:'2px 0'}}>
                <span style={{fontWeight:600,color:T.textBright}}>{c.account_code}</span> <span style={{color:T.textMuted}}>{c.account_name}</span>
                <span style={{float:'right',fontVariantNumeric:'tabular-nums'}}>{fmt(c.amount)}</span></div>)}</div>}
            {r.other_legs&&r.other_legs.length>0&&<div>
              <div style={{fontSize:11,fontWeight:700,color:T.textMuted,textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:4}}>Other capital on this company</div>
              {r.other_legs.map(c=><div key={c.account_code} style={{fontSize:12,padding:'2px 0',color:T.textMuted}}>
                <span style={{fontWeight:600}}>{c.account_code}</span> {c.account_name}
                {!c.mapped&&<span style={{marginLeft:6}}><IcBadge kind="mute">not mapped</IcBadge></span>}
                <span style={{float:'right',fontVariantNumeric:'tabular-nums'}}>{fmt(c.amount)}</span></div>)}</div>}
          </td></tr>}
          </Fragment>;})}
        </tbody></table></div>}

        {recon.nci&&recon.nci.length>0&&<div className="cl-scroll" style={scrollBox()}>
          <div style={{padding:'12px 16px',background:T.bgElevated,borderBottom:'1px solid '+T.border}}>
            <span style={{fontWeight:700,color:T.textBright,fontSize:13}}>Minority interest — indicative</span>
            <span style={{color:T.textMuted,fontSize:12,marginLeft:8}}>Each company's total equity at this date multiplied by the share owned outside the structure. Only the companies where the dilution happens are counted, so the same minority isn't charged twice down the chain. This is a GL-derived estimate, not a computed NCI roll-forward — it includes any self-referential gross-up and excludes results not yet closed to equity.</span></div>
          <table style={S.table}><thead><tr>
            <th style={S.th}>Company</th><th style={S.thR}>Owned %</th><th style={S.thR}>Effective %</th>
            <th style={S.thR}>Minority %</th><th style={S.thR}>Total equity</th><th style={S.thR}>Minority share</th></tr></thead>
          <tbody>{recon.nci.map(n=><tr key={n.node_id}>
            <td style={{...S.td,fontWeight:600,color:T.textBright}}>{n.name}</td>
            <td style={S.tdR}>{Number(n.ownership_pct).toFixed(2)}%</td>
            <td style={S.tdR}>{Number(n.effective_pct).toFixed(2)}%</td>
            <td style={{...S.tdR,color:T.orange,fontWeight:600}}>{Number(n.nci_pct).toFixed(2)}%</td>
            <td style={S.tdR}>{fmt(n.entity_equity)}</td>
            <td style={{...S.tdR,fontWeight:600}}>{fmt(n.nci_equity)}</td></tr>)}
            <tr style={S.grandTotalRow}><td style={S.tdBold}>Total</td><td style={S.tdR}></td><td style={S.tdR}></td><td style={S.tdR}></td><td style={S.tdR}></td>
              <td style={{...S.tdBold,textAlign:'right'}}>{fmt(recon.totals.nci_equity_total)}</td></tr>
          </tbody></table></div>}
      </>}
    </>}
  </div>);
}
