// ═══ Administration: Compliance Certificates (quarterly lender reporting packages) ═══
//
// Given a quarter-end date, rebuilds the quarterly lender compliance packages
// the same way they were prepared by hand for Q3 2026:
//
//   Skye Canyon (Trimont)       - Yardi 3-month statement (from Cushman's month-end
//                                 "MergedExcel" package), the month-end financial
//                                 reports PDF, and an unsigned certified rent roll.
//   South Mountain / Orchards   - Mark-Taylor month-end financial PDFs (SM + Preserve),
//   (Trimont)                     a T3 (last 3 months) built from each 12-month P&L,
//                                 and two unsigned certified rent rolls.
//   Silo Ridge - JPM            - DSCR calc: quarter revenue / opex / interest from the
//                                 three monthly Greystar variance reports (jpm_is tree).
//   Silo Ridge - Wells Fargo    - balance sheet / income statement / 12-month pages
//                                 pulled from Greystar's quarter-end operating report,
//                                 plus an unsigned guarantor compliance certificate.
//   Midco I (Bank of Texas)     - entity quarterly financial statements, the CLRFI
//                                 compliance certificate (quarter-end date stamped),
//                                 and the financial covenant workbook (DSCR on a
//                                 trailing-N-month annualized basis per Credit Agreement
//                                 §10.1, LTV vs appraisals, Debt Yield).
//
// Everything is read from and written to SharePoint via Microsoft Graph (app-only),
// into  Accounting > General/3. Projects/Compliance Certificates/{YYYY}/Q{n}/...
// For local runs the same code can read the OneDrive-synced folders instead
// (CC_LOCAL_ROOT), which is how it was verified against the hand-built Q3 2026 set.
//
// Nothing is emailed or signed - certificates are produced unsigned with the
// signature/delivery dates left blank.

const path = require('path');
const fs = require('fs');

const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = MON.map((m) => m.slice(0, 3));
const pad2 = (n) => String(n).padStart(2, '0');
const ord = (d) => d + (d % 10 === 1 && d !== 11 ? 'st' : d % 10 === 2 && d !== 12 ? 'nd' : d % 10 === 3 && d !== 13 ? 'rd' : 'th');

const DEST_BASE = 'General/3. Projects/Compliance Certificates';
const DIRS = {
  midco: 'Midco I',
  jpm: 'Silo Ridge/JPM',
  wf: 'Silo Ridge/Wells Fargo',
  skye: 'Skye Canyon',
  smo: 'South Mountain and Orchards',
};

// Editable configuration (stored in compliance_cert_config). Paths are relative
// to each SharePoint library root: accounting = "Accounting" site Documents,
// cla = "CLA" site Documents, shared = "Banyan Res Shared Drive" Documents.
const DEFAULTS = {
  sp_drives: { accounting: '', cla: '', shared: '' },   // blank = env or auto-discover
  sp_sites: { accounting: 'Accounting', cla: 'CLA', shared: 'Banyan Res Shared Drive' },
  sfr_base: 'Projects/Asset Management/SFR/Financials',
  // draft email sent with the Silo Ridge DSCR calc (saved as an unsent .eml next to it)
  silo_jpm_email: {
    to: ['REA Finance Operations CASE <rea.finance.case@jpmorgan.com>', 'Kale Easterling <kale.easterling@greystar.com>',
      'Michael Bie <michael.bie@greystar.com>', 'Banyan Silo Ridge Manager <siloridgemgr@greystar.com>',
      'Shea Johnson <shea.johnson@greystar.com>', 'Rachel Raiford <rraiford@greystar.com>', 'Chandan Singh <chandan.singh@greystar.com>'],
    cc: ['REA Finance - CASE <REA_Finance_-_CASE@jpmorgan.com>', 'Irvin Bermudez <IBermudez@banyanres.com>'],
    from: 'Jimmy Yun <JYun@banyanres.com>',
  },
  midco: {
    dscr_months: 6,            // §10.1: trailing 6 months annualized until the Reserve Account Release Date
    dscr_min_through_2026: 1.10,
    dscr_min_after_2026: 1.20,
    ltv_max: 0.60,
    debt_yield_min: 0.085,
    appraisals: { CLIP: 121800000, Silsbee: 31000000, Buna: 20700000, SRN: 108100000 },
    appraisal_label: 'Q1 2026 Appraisals',
    nonrecurring_addback: 0,
    entities: [
      { key: 'CLIP', fin: '02 CLIP/00 Financials', req: '02 CLIP/01 Requisition Report', fs: /quarterly consolidated financial statements/i },
      { key: 'Silsbee', fin: '01 Silsbee/01 Financials', fs: /quarterly (financial statements|fs)\b/i },
      { key: 'SRN', fin: '04 SRN/01 Financials', fs: /quarterly (financial statements|fs)\b/i },
      { key: 'Buna', fin: '05 Buna/00 Financials', fs: /quarter ended/i },
    ],
  },
};

// ── quarter helpers ─────────────────────────────────────────────────────────
function parseQuarterEnd(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Quarter-end date must be YYYY-MM-DD');
  const y = +m[1], mo = +m[2], d = +m[3];
  if (![3, 6, 9, 12].includes(mo)) throw new Error('Not a quarter-end month (use Mar/Jun/Sep/Dec)');
  const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  if (d !== last) throw new Error('Quarter-end must be the last day of the month (' + y + '-' + pad2(mo) + '-' + last + ')');
  const q = mo / 3;
  const back = (k) => { const dt = new Date(Date.UTC(y, mo - 1 - k, 1)); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1 }; };
  return {
    iso: s, y, m: mo, d, q,
    label: 'Q' + q + ' ' + y, tag: q + 'Q' + y,          // "Q3 2026", "3Q2026"
    months: [back(2), back(1), back(0)],                    // quarter months, oldest first
    back,
  };
}
const mLabel = (p) => MON3[p.m - 1] + ' ' + p.y;           // "Sep 2026"

// ── storage: SharePoint (Graph) or local OneDrive-synced folders ─────────────
let tokenCache = null;
async function graphToken() {
  const t = process.env.MS_TENANT_ID, c = process.env.MS_CLIENT_ID, s = process.env.MS_CLIENT_SECRET;
  if (!t || !c || !s) return null;
  if (tokenCache && tokenCache.exp > Date.now() + 60000) return tokenCache.tok;
  const r = await fetch('https://login.microsoftonline.com/' + t + '/oauth2/v2.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: c, client_secret: s, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!r.ok) throw new Error('Graph token ' + r.status + ': ' + (await r.text()).slice(0, 200));
  const j = await r.json(); tokenCache = { tok: j.access_token, exp: Date.now() + (j.expires_in || 3000) * 1000 };
  return tokenCache.tok;
}
const graphConfigured = () => !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
const spEnc = (p) => String(p).split('/').filter(Boolean).map(encodeURIComponent).join('/');

function makeGraphStore(cfg) {
  const driveCache = {};
  const ENV = { accounting: 'CC_SP_DRIVE_ACCOUNTING', cla: 'CC_SP_DRIVE_CLA', shared: 'CC_SP_DRIVE_SHARED' };
  async function gfetch(url, opts = {}) {
    const token = await graphToken(); if (!token) throw new Error('Microsoft Graph is not configured (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET)');
    const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url,
      { method: opts.method || 'GET', headers: Object.assign({ Authorization: 'Bearer ' + token }, opts.headers || {}), body: opts.body });
    if (!r.ok && !(opts.ok || []).includes(r.status)) throw new Error('Graph ' + (opts.method || 'GET') + ' ' + r.status + ': ' + (await r.text()).slice(0, 250));
    return r;
  }
  async function drive(lib) {
    if (driveCache[lib]) return driveCache[lib];
    let id = (cfg.sp_drives && cfg.sp_drives[lib]) || process.env[ENV[lib]] || (lib === 'accounting' ? (process.env.AMFEE_SP_DRIVE_ID || process.env.USC_SP_DRIVE_ID) : '');
    if (!id) {   // auto-discover the site's default document library
      const name = (cfg.sp_sites || {})[lib];
      const sr = await (await gfetch('/sites?search=' + encodeURIComponent(name))).json();
      const site = (sr.value || []).find((s) => (s.displayName || '').toLowerCase() === name.toLowerCase()) || (sr.value || [])[0];
      if (!site) throw new Error('SharePoint site "' + name + '" not found - set its drive id in Compliance Certificates config');
      id = (await (await gfetch('/sites/' + site.id + '/drive')).json()).id;
    }
    driveCache[lib] = id; return id;
  }
  return {
    kind: 'sharepoint',
    async list(lib, p) {
      const d = await drive(lib);
      const r = await gfetch('/drives/' + d + '/root:/' + spEnc(p) + ':/children?$top=999', { ok: [404] });
      if (r.status === 404) return null;
      return ((await r.json()).value || []).map((x) => ({ name: x.name, folder: !!x.folder, size: x.size }));
    },
    async read(lib, p) {
      const d = await drive(lib);
      return Buffer.from(await (await gfetch('/drives/' + d + '/root:/' + spEnc(p) + ':/content')).arrayBuffer());
    },
    async write(lib, p, buf) {
      const d = await drive(lib);
      const parts = p.split('/').filter(Boolean); const fname = parts.pop();
      let cur = '';
      for (const part of parts) {
        const parent = cur; cur = cur ? cur + '/' + part : part;
        await gfetch('/drives/' + d + (parent ? '/root:/' + spEnc(parent) + ':/children' : '/root/children'), {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, ok: [409],
          body: JSON.stringify({ name: part, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
        });
      }
      const full = parts.concat([fname]).join('/');
      let j;
      if (buf.length <= 4 * 1024 * 1024) {
        j = await (await gfetch('/drives/' + d + '/root:/' + spEnc(full) + ':/content', { method: 'PUT', headers: { 'Content-Type': 'application/octet-stream' }, body: buf })).json();
      } else {   // upload session for large files
        const sess = await (await gfetch('/drives/' + d + '/root:/' + spEnc(full) + ':/createUploadSession', {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) })).json();
        const CH = 320 * 1024 * 30;   // ~9.4 MiB, a multiple of 320 KiB as Graph requires
        for (let off = 0; off < buf.length; off += CH) {
          const chunk = buf.subarray(off, Math.min(off + CH, buf.length));
          const r = await fetch(sess.uploadUrl, { method: 'PUT', headers: { 'Content-Length': String(chunk.length), 'Content-Range': 'bytes ' + off + '-' + (off + chunk.length - 1) + '/' + buf.length }, body: chunk });
          if (!r.ok && r.status !== 202) throw new Error('Upload session ' + r.status + ': ' + (await r.text()).slice(0, 200));
          if (r.status === 200 || r.status === 201) j = await r.json();
        }
      }
      return (j && j.webUrl) || null;
    },
  };
}

function makeLocalStore() {
  const root = process.env.CC_LOCAL_ROOT;   // e.g. C:\Users\JimmyYun\banyanres.com
  const writeRoot = process.env.CC_LOCAL_WRITE_ROOT || root;
  const LIB = { accounting: 'Accounting - Documents', cla: 'CLA - Documents', shared: 'Banyan Res Shared Drive - Documents' };
  const abs = (r, lib, p) => path.join(r, LIB[lib], ...String(p).split('/').filter(Boolean));
  return {
    kind: 'local',
    async list(lib, p) {
      const d = abs(root, lib, p); if (!fs.existsSync(d)) return null;
      return fs.readdirSync(d, { withFileTypes: true }).map((e) => ({ name: e.name, folder: e.isDirectory() }));
    },
    async read(lib, p) { return fs.readFileSync(abs(root, lib, p)); },
    async write(lib, p, buf) { const f = abs(writeRoot, lib, p); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, buf); return f; },
  };
}
function makeStore(cfg) {
  if (process.env.CC_STORAGE === 'local' || (!graphConfigured() && process.env.CC_LOCAL_ROOT)) return makeLocalStore();
  return makeGraphStore(cfg);
}

// month folder finder: "09.2026", "09 Sep 2026", "09 September 2026", "09 Sept 2026"
function monthFolder(names, y, m) {
  const MM = pad2(m), yr = String(y);
  const c = names.filter((n) => n.includes(yr) && new RegExp('^' + MM + '(\\D|$)').test(n.trim()));
  return c.sort((a, b) => a.length - b.length)[0] || null;
}
async function findMonthDir(store, lib, base, y, m) {
  for (const yb of [base + '/' + y, base]) {
    const kids = await store.list(lib, yb); if (!kids) continue;
    const mf = monthFolder(kids.filter((k) => k.folder).map((k) => k.name), y, m);
    if (mf) return yb + '/' + mf;
  }
  return null;
}
async function findFile(store, lib, dir, re, ext) {
  if (!dir) return null;
  const kids = await store.list(lib, dir); if (!kids) return null;
  const hits = kids.filter((k) => !k.folder && re.test(k.name) && (!ext || ext.test(k.name)) && !/^~\$/.test(k.name));
  hits.sort((a, b) => a.name.length - b.name.length);
  return hits.length ? dir + '/' + hits[0].name : null;
}

// ── parsing helpers ─────────────────────────────────────────────────────────
function sheetRows(buf, sheetName) {
  const XLSX = require('xlsx');
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const ws = sheetName ? wb.Sheets[sheetName] : wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('Sheet "' + sheetName + '" not found');
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}
const num = (v) => (typeof v === 'number' ? v : (v == null || v === '' ? null : (isFinite(+String(v).replace(/[,$\s]/g, '')) ? +String(v).replace(/[,$\s]/g, '') : null)));
const lab = (v) => (v == null ? '' : String(v).trim());
// Month header cells come as "Sep 2026", "September 2026", "09/30/2026" or real dates
function parseMonthCell(v) {
  if (v instanceof Date) return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1 };
  const s = lab(v); let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return { y: +m[3], m: +m[1] };
  m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{4})$/);
  if (m) { const i = MON3.findIndex((x) => x.toLowerCase() === m[1].slice(0, 3).toLowerCase()); if (i >= 0) return { y: +m[2], m: i + 1 }; }
  return null;
}
// find the header row with the most month columns; map "YYYY-M" -> column index (first occurrence)
function findMonthHeader(rows) {
  let best = { row: -1, map: {}, n: 0 };
  for (let r = 0; r < Math.min(rows.length, 12); r++) {
    const map = {}; let n = 0;
    (rows[r] || []).forEach((v, c) => { const p = parseMonthCell(v); if (p) { const k = p.y + '-' + p.m; if (!(k in map)) { map[k] = c; n++; } } });
    if (n > best.n) best = { row: r, map, n };
  }
  return best;
}

async function pdfPagesText(buf) {
  const pdf = require('pdf-parse/lib/pdf-parse.js');
  const pages = [];
  await pdf(buf, {
    pagerender: async (pd) => {
      const tc = await pd.getTextContent();
      const t = tc.items.map((i) => i.str).join(' ').replace(/\s+/g, ' ');
      pages[pd.pageIndex] = t; return t;
    },
    max: 0,
  });
  return pages;
}

// ── workbook helpers (ExcelJS) ──────────────────────────────────────────────
const NF = '#,##0.00_);(#,##0.00)';
function newWb() { const ExcelJS = require('exceljs'); const wb = new ExcelJS.Workbook(); wb.calcProperties.fullCalcOnLoad = true; return wb; }
const YEL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
// Widen every column to fit its contents (numbers as displayed, labels, formula
// results). Never narrows a column that was given a width on purpose. Long title
// lines (> 60 chars) are ignored so they keep spilling across, as in Excel.
function autofitWorkbook(wb) {
  for (const ws of wb.worksheets) {
    if (ws.__noAutofit) continue;   // copied from a template: keep its exact widths
    const need = {};
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (c) => {
      let v = c.value;
      if (v && typeof v === 'object' && !(v instanceof Date)) {
        if ('result' in v) v = v.result; else if (v.richText) v = v.richText.map((t) => t.text).join(''); else if ('text' in v) v = v.text; else return;
      }
      if (v == null || v === '') return;
      let len;
      if (v instanceof Date) len = 10;
      else if (typeof v === 'number') {
        const fmt = String(c.numFmt || '');
        const dec = /0\.0000/.test(fmt) ? 4 : /0\.00/.test(fmt) ? 2 : /%/.test(fmt) ? 2 : (Number.isInteger(v) || /#,##0[^.]|#,##0$/.test(fmt) ? 0 : 2);
        let t = Math.abs(fmt.includes('%') ? v * 100 : v).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
        if (fmt.includes('%')) t += '%';
        if (/"x"/.test(fmt)) t += 'x';
        len = t.length + 2;   // room for parentheses / currency padding
      } else {
        len = String(v).split(/\r?\n/).reduce((m, l) => Math.max(m, l.length), 0);
        if (len > 60) return;
      }
      if (c.font && c.font.bold) len *= 1.1;
      need[c.col] = Math.max(need[c.col] || 0, len);
    }));
    for (const [col, len] of Object.entries(need)) {
      const column = ws.getColumn(+col); const want = Math.min(65, Math.ceil(len + 2));
      if (!column.width || column.width < want) column.width = want;
    }
  }
}
async function wbBuffer(wb) { if (!wb.__noAutofit) autofitWorkbook(wb); return Buffer.from(await wb.xlsx.writeBuffer()); }
function setF(ws, addr, formula, result, fmt) { const c = ws.getCell(addr); c.value = { formula: formula.replace(/^=/, ''), result }; if (fmt) c.numFmt = fmt; return c; }
function colL(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
const qs = (name) => "'" + String(name).replace(/'/g, "''") + "'";

// ── PDF certificate stamping ────────────────────────────────────────────────
const TPL_DIR = path.join(__dirname, 'templates', 'compliance');
async function stampCertificate(key, values) {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
  const meta = JSON.parse(fs.readFileSync(path.join(TPL_DIR, 'fields.json'), 'utf8'))[key];
  const doc = await PDFDocument.load(fs.readFileSync(path.join(TPL_DIR, meta.file)));
  const font = await doc.embedFont(key === 'bot_clrfi' ? StandardFonts.Helvetica : StandardFonts.TimesRoman);
  const page = doc.getPages()[meta.page || 0];
  for (const f of meta.fields) {
    if (!(f.name in values)) continue;
    const ov = (values.__override || {})[f.name] || {};
    const x = ov.x != null ? ov.x : f.x, w = ov.w != null ? ov.w : f.w;
    page.drawRectangle({ x: x - 0.5, y: f.y, width: w + 1, height: f.h, color: rgb(1, 1, 1) });
    if (values[f.name]) page.drawText(values[f.name], { x, y: f.baseline, size: f.size, font, color: rgb(0, 0, 0) });
    if (f.underline) page.drawRectangle({ x: f.underline.x, y: f.underline.y, width: f.underline.w, height: f.underline.h, color: rgb(0, 0, 0) });   // restore the signature-date line
  }
  return Buffer.from(await doc.save());
}


// ── CloudLedger as a source ────────────────────────────────────────────────
// When a statement isn't in the SharePoint folders yet, generate it straight from
// CloudLedger with the same generators used by hand (Workpapers > Financial
// Statements, the CLRF fund statements, the Trailing 12 Months P&L).
// opts.cl = { url, token, entities:[{id,code,name,entity_type}] } is supplied by
// the runner (token = short-lived login of the person who clicked Report).
// CloudLedger entities. 'fs' = the entity whose quarterly statements go in the
// lender package (as in the Q2 2026 package); 'ops' = the entity that carries the
// operating numbers used in the covenant calc (T12, loan, reserve, interest).
// Matched by exact name first, then by entity code.
const CL_MAP = {
  CLIP: { fs: ['County Line Industrial Park', 'COUNTYLI2'], ops: ['CLIP Property Owner', 'CLIPPROP'] },
  Silsbee: { fs: ['County Line Rail Silsbee', 'COUNTYLI5'], ops: ['CLR Silsbee Property Owner', 'CLRSILSB2'] },
  SRN: { fs: ['Sabine River & Northern Railroad', 'SABINERI'], ops: ['Sabine River & Northern Railroad', 'SABINERI'] },
  Buna: { fs: ['CLR Buna Property Owner', 'CLRBUNAP'], ops: ['CLR Buna Property Owner', 'CLRBUNAP'] },
  Midco: { fs: ['CLRFI Midco I', 'CLRFIMID'] },
  CLRF: { fs: ['County Line Rail Fund', 'COUNTYLI1'] },
};
function clEntity(cl, key, role = 'ops') {
  if (!cl || !cl.entities || !CL_MAP[key]) return null;
  const want = CL_MAP[key][role] || CL_MAP[key].fs; if (!want) return null;
  const norm = (x) => String(x || '').trim().toLowerCase().replace(/,?\s*(llc|lp|l\.p\.)$/, '');
  return cl.entities.find((e) => norm(e.name) === norm(want[0])) || cl.entities.find((e) => String(e.code) === want[1]) || null;
}
async function clFetch(cl, p, opts = {}) {
  const r = await fetch(cl.url.replace(/\/$/, '') + p, { method: opts.method || 'GET', headers: { Authorization: 'Bearer ' + cl.token }, body: opts.body });
  if (!r.ok) throw new Error('CloudLedger ' + p.split('?')[0] + ' -> HTTP ' + r.status + ' ' + (await r.text()).slice(0, 200));
  return r;
}
// Account balances from the CloudLedger GL: {as_of} = balance at a date,
// {from,to} = activity in a period. balance is debit-positive.
async function clBalances(cl, ent, q) {
  const qs2 = Object.entries(q).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
  return (await clFetch(cl, '/api/entities/' + ent.id + '/balances?' + qs2)).json();
}
// cash interest on the BOT loan = interest expense + interest capitalized to the project
const isInterestExp = (r) => r.type === 'Expense' && /interest/i.test(r.name || '') && !/income/i.test(r.name || '');
const isCapInterest = (r) => r.type === 'Asset' && (/^(12321|12325)\b/.test(String(r.code)) || /construction period interest|capitalized interest/i.test(r.name || ''));
const isBotLoan = (r) => /^25063\b/.test(String(r.code)) || (r.type === 'Liability' && /loan payable/i.test(r.name || '') && /(bot|bank of texas|midco)/i.test(r.name || ''));

// Quarterly financial statements PDF (Workpapers > Financial Statements generator)
async function clQuarterlyFs(cl, ent, asOf) {
  const fd = new FormData(); fd.append('as_of', asOf); fd.append('period', 'quarterly');
  const r = await clFetch(cl, '/api/workpapers/financial-statements/' + ent.id + '/generate', { method: 'POST', body: fd });
  return Buffer.from(await r.arrayBuffer());
}
async function clFundFs(cl, ent, asOf) {
  const r = await clFetch(cl, '/api/entities/' + ent.id + '/fund-statements.pdf?as_of=' + asOf);
  return Buffer.from(await r.arrayBuffer());
}
// Trailing 12 Months P&L -> rows in the same shape as the folder T12 exports
// (row 5 = month headers; column A labels incl. Total Revenue / Gross Profit /
// Total Operating Expenses / Depreciation Expense / Amortization Expense).
async function clT12Rows(cl, ent, asOf) {
  const t = await (await clFetch(cl, '/api/entities/' + ent.id + '/ttm-pl?as_of=' + asOf)).json();
  const months = t.meta.months.map((m) => m.label);
  const row = (label, o) => [label].concat(o ? o.vals : months.map(() => null)).concat([o ? o.total : null]);
  const sum = (lines) => { const v = months.map(() => 0); let tot = 0; lines.forEach((l) => { l.vals.forEach((x, i) => { v[i] += x || 0; }); tot += l.total || 0; }); return { vals: v, total: tot }; };
  const opex = t.opex || [];
  const rows = [[ent.name], ['Trailing 12 Months'], [t.meta.periodLabel + ' (from CloudLedger)'], [], [''].concat(months).concat(['Total'])];
  rows.push(['Revenue']); (t.revenue || []).forEach((l) => rows.push(row(l.name, l))); rows.push(row('Total Revenue', t.totRev));
  if (t.hasCogs) { (t.cogs || []).forEach((l) => rows.push(row(l.name, l))); rows.push(row('Gross Profit', t.grossProfit)); }
  rows.push(['Operating Expenses']);
  (t.opexGroups || []).forEach((g) => { rows.push([g.title]); g.lines.forEach((l) => rows.push(row(l.name, l))); });
  rows.push(row('Depreciation Expense', sum(opex.filter((l) => /depreciation/i.test(l.name)))));
  rows.push(row('Amortization Expense', sum(opex.filter((l) => /amortization/i.test(l.name)))));
  rows.push(row('Total Operating Expenses', t.totOpex));
  rows.push(row('Net Income (Loss)', t.netIncome));
  return rows;
}
// ── unsent email draft (.eml) ──────────────────────────────────────────────
// X-Unsent: 1 makes Outlook open the file as a new editable message (with the
// attachment) instead of a received one. Nothing is sent.
function buildDraftEml({ from, to, cc, subject, text, html, attachments }) {
  const enc = (v) => (/^[\x20-\x7e]*$/.test(v) ? v : '=?UTF-8?B?' + Buffer.from(v, 'utf8').toString('base64') + '?=');
  const b64 = (buf) => buf.toString('base64').replace(/.{76}/g, '$&\r\n');
  const mix = 'mix_' + Date.now().toString(36), alt = 'alt_' + Date.now().toString(36);
  const L = [];
  L.push('X-Unsent: 1', 'MIME-Version: 1.0', 'From: ' + from, 'To: ' + to.join(', '));
  if (cc && cc.length) L.push('Cc: ' + cc.join(', '));
  L.push('Subject: ' + enc(subject), 'Content-Type: multipart/mixed; boundary="' + mix + '"', '');
  L.push('--' + mix, 'Content-Type: multipart/alternative; boundary="' + alt + '"', '');
  L.push('--' + alt, 'Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: base64', '', b64(Buffer.from(text, 'utf8')));
  L.push('--' + alt, 'Content-Type: text/html; charset=utf-8', 'Content-Transfer-Encoding: base64', '', b64(Buffer.from(html, 'utf8')));
  L.push('--' + alt + '--', '');
  for (const a of attachments || []) {
    L.push('--' + mix, 'Content-Type: ' + a.type + '; name="' + a.name + '"', 'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="' + a.name + '"', '', b64(a.data));
  }
  L.push('--' + mix + '--', '');
  return Buffer.from(L.join('\r\n'), 'utf8');
}
const SIGNATURE_TEXT = ['Jimmy Yun', 'Chief Accounting Officer', 'M: 714.321.6021', '2381 Rosecrans Ave Suite 425, El Segundo, CA 90245', 'banyanres.com'];


// ── Mark-Taylor / ResMan "Trailing Profit And Loss Detail" (T3) ────────────
// Rebuilt in exactly the layout of the prior quarter's T3 (the ResMan export):
// outline INCOME > RENTAL REVENUE > GROSS POT RENT > accounts, group totals,
// TOTAL INCOME / TOTAL EXPENSE / NET OPERATING INCOME / NET INCOME, and a
// Variance column vs budget (income: actual - budget; expense: budget - actual).
const acctNo = (v) => { const m = String(v == null ? '' : v).trim().match(/^(\d{4}(?:\.\d+)?)(?=\s|$)/); return m ? m[1] : null; };   // keeps sub-accounts like 5002.01 distinct from 5002
function t3Outline(rows) {   // parse the prior quarter's T3 into sections > categories > (subgroups | accounts)
  const sections = []; let sec = null, cat = null, sub = null;
  for (let r = 6; r < rows.length; r++) {
    const R = rows[r] || []; let col = -1;
    for (let c = 0; c < 5; c++) if (R[c] != null && String(R[c]).trim() !== '') { col = c; break; }
    if (col < 0) continue;
    const lab = String(R[col]).trim(); const hasNum = [5, 6, 7, 8, 9].some((c) => typeof R[c] === 'number');
    const no = acctNo(lab);
    if (no) { const a = { no, label: lab }; if (col === 4 && sub) sub.accts.push(a); else if (cat) { sub = null; cat.items.push({ acct: a }); } continue; }
    if (hasNum || /^(total|net)\b/i.test(lab) || /^©/.test(lab)) { if (/^total\s/i.test(lab) && sub && lab.slice(6).trim().toUpperCase() === sub.name.toUpperCase()) sub = null; continue; }
    if (col === 1) { sec = { name: lab, cats: [] }; sections.push(sec); cat = null; sub = null; }
    else if (col === 2 && sec) { cat = { name: lab, items: [] }; sec.cats.push(cat); sub = null; }
    else if (col === 3 && cat) { sub = { name: lab, accts: [] }; cat.items.push({ sub }); }
  }
  return sections;
}
function mtrAccounts(rows, cols) {   // account -> {label, vals[]} from a Mark-Taylor MTR report (P&L only: stops at NET INCOME, so capital items below it are left out)
  const out = {};
  for (const R of rows) {
    if (R && R.slice(0, 5).some((x) => x != null && /^NET INCOME$/i.test(String(x).trim()))) break;
    if (!R) continue; let lab = null;
    for (let c = 0; c < 5; c++) if (acctNo(R[c])) { lab = String(R[c]).trim(); break; }
    if (!lab) continue; const no = acctNo(lab);
    if (/^[12]/.test(no)) continue;   // balance-sheet / capital items (e.g. 14xx capex) are not on the ResMan T3
    out[no] = { label: lab, vals: cols.map((c) => (typeof R[c] === 'number' ? R[c] : 0)) };
  }
  return out;
}
function mtrRowByLabel(rows, re, cols) {
  for (const R of rows) { if (!R) continue; for (let c = 0; c < 5; c++) if (R[c] != null && re.test(String(R[c]).trim())) return cols.map((k) => (typeof R[k] === 'number' ? R[k] : 0)); }
  return null;
}

// Formatting profiles taken from the prior quarter's T3 (ResMan export): for each kind
// of row, the exact cell styles, row height and merged spans of its first occurrence.
async function xlsxCellFontSizes(buf) {   // 'E10' -> 7.5, read from styles.xml (ExcelJS parses font sizes with parseInt)
  const JSZip = require('jszip'); const zip = await JSZip.loadAsync(buf);
  const styles = await zip.file('xl/styles.xml').async('string');
  const sheetName = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];
  const sheet = await zip.file(sheetName).async('string');
  const fontsXml = (styles.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/) || [])[1] || '';
  const sizes = (fontsXml.match(/<font\b[\s\S]*?<\/font>|<font\b[^>]*\/>/g) || []).map((f) => { const m = f.match(/<sz val="([\d.]+)"/); return m ? parseFloat(m[1]) : null; });
  const xfsXml = (styles.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
  const xfFont = (xfsXml.match(/<xf\b[^>]*>/g) || []).map((x) => { const m = x.match(/fontId="(\d+)"/); return m ? +m[1] : 0; });
  const out = {}; const re = /<c\b([^>]*)>/g; let m;   // attribute order varies (t="s" r=".." s="..")
  while ((m = re.exec(sheet))) { const rm = m[1].match(/\br="([A-Z]+\d+)"/), sm = m[1].match(/\bs="(\d+)"/); if (rm && sm) out[rm[1]] = sizes[xfFont[+sm[1]]]; }
  return out;
}
async function t3StyleProfiles(tplBuf) {
  const ExcelJS = require('exceljs'); const twb = new ExcelJS.Workbook(); await twb.xlsx.load(tplBuf);
  const tws = twb.worksheets[0];
  let fsz = {}; try { fsz = await xlsxCellFontSizes(tplBuf); } catch (e) { /* keep ExcelJS sizes */ }
  const colL2 = (n) => { let t = ''; while (n > 0) { const q = (n - 1) % 26; t = String.fromCharCode(65 + q) + t; n = Math.floor((n - 1) / 26); } return t; };
  const merges = {}; (tws.model.merges || []).forEach((m) => {
    const mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if (!mm || mm[2] !== mm[4]) return;
    const col = (L) => L.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    (merges[+mm[2]] = merges[+mm[2]] || []).push([col(mm[1]), col(mm[3])]);
  });
  const snap = (r) => ({ height: tws.getRow(r).height, merges: merges[r] || [],
    cells: Array.from({ length: 11 }, (_, i) => {
      const st = JSON.parse(JSON.stringify(tws.getCell(r, i + 1).style || {}));
      const exact = fsz[colL2(i + 1) + r]; if (exact && st.font) st.font.size = exact;   // keep 7.5pt etc.
      return st;
    }) });
  const prof = {}; const put = (k, r) => { if (!prof[k]) prof[k] = snap(r); };
  for (let r = 1; r <= 6; r++) prof['head' + r] = Object.assign(snap(r), { values: Array.from({ length: 11 }, (_, i) => tws.getCell(r, i + 1).value) });
  const open = {}; let afterNoi = 0;
  for (let r = 7; r <= tws.rowCount; r++) {
    let col = 0, lab = null;
    for (let c = 1; c <= 5; c++) { const v = tws.getCell(r, c).value; if (v != null && String(v).trim() !== '') { col = c; lab = String(v).trim(); break; } }
    if (!lab) { if (prof.noi && !prof.gapAfterNoi) { afterNoi++; prof.gapHeights = (prof.gapHeights || []).concat([tws.getRow(r).height]); } continue; }
    if (prof.noi && afterNoi) prof.gapAfterNoi = afterNoi;
    const isNum = typeof tws.getCell(r, 6).value === 'number';
    if (/^\u00a9/.test(lab)) { put('footer', r); continue; }
    if (acctNo(lab)) { const shaded = !!(tws.getCell(r, col).fill && tws.getCell(r, col).fill.fgColor && tws.getCell(r, col).fill.fgColor.argb); put((col === 5 ? 'detE' : 'detD') + (shaded ? 'S' : ''), r); continue; }
    if (/^NET OPERATING INCOME/i.test(lab)) { put('noi', r); continue; }
    if (/^NET INCOME/i.test(lab)) { put('ni', r); continue; }
    if (/^TOTAL\s/.test(lab) && lab === lab.toUpperCase()) { put('totSec', r); continue; }
    if (/^Total\s/i.test(lab)) { const n = lab.slice(6).trim().toUpperCase(); put(open[4] === n ? 'totSub' : 'totCat', r); continue; }
    if (!isNum) { open[col] = lab.toUpperCase(); put(col === 2 ? 'sec' : col === 3 ? 'cat' : 'sub', r); }
  }
  return { prof, widths: Array.from({ length: 11 }, (_, i) => tws.getColumn(i + 1).width), views: tws.views };
}

function buildResmanT3({ prop, Q, outline, actual, budgetQ, months, styles }) {
  // place accounts that have quarter activity/budget but are not in the outline, next to their nearest lower account
  const all = []; outline.forEach((s) => s.cats.forEach((c) => c.items.forEach((it) => { if (it.acct) all.push({ a: it.acct, list: c.items, isItem: true }); else it.sub.accts.forEach((a) => all.push({ a, list: it.sub.accts, isItem: false })); })));
  const known = new Set(all.map((x) => x.a.no));
  const extra = Object.keys(actual).concat(Object.keys(budgetQ)).filter((no, i, arr) => arr.indexOf(no) === i && !known.has(no))
    .filter((no) => (actual[no] && actual[no].vals.some((v) => Math.abs(v) > 0.004)) || Math.abs(budgetQ[no] || 0) > 0.004).sort();
  const added = [];
  for (const no of extra) {
    const lower = all.filter((x) => x.a.no < no).sort((x, y) => (x.a.no < y.a.no ? 1 : -1))[0] || all[0];
    if (!lower) continue;
    const a = { no, label: (actual[no] && actual[no].label) || no };
    const idx = lower.list.findIndex((e) => (lower.isItem ? e.acct === lower.a : e === lower.a));
    lower.list.splice(idx + 1, 0, lower.isItem ? { acct: a } : a);
    all.push({ a, list: lower.list, isItem: lower.isItem }); added.push(a.label);
  }
  const ExcelJS = require('exceljs'); const wb = new ExcelJS.Workbook(); wb.__noAutofit = true;
  const ws = wb.addWorksheet('Sheet', { views: (styles && styles.views) || [{ showGridLines: false }] });
  const P = (styles && styles.prof) || {};
  // stamp a template row's exact formatting (styles, height, merges) onto output row rr
  const applyProf = (rr, k) => {
    const pr = P[k]; if (!pr) return false;
    pr.cells.forEach((st, i) => { if (st && Object.keys(st).length) ws.getCell(rr, i + 1).style = JSON.parse(JSON.stringify(st)); });
    if (pr.height) ws.getRow(rr).height = pr.height;
    pr.merges.forEach(([c1, c2]) => { try { ws.mergeCells(rr, c1, rr, c2); } catch (e) { /* overlapping */ } });
    return true;
  };
  const NUMF = '[$-1010409]#,##0.00;(#,##0.00)', TF = 'Trebuchet MS', thin = { style: 'thin' }, SHADE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
  ((styles && styles.widths) || [1.71, 1.57, 1.71, 1.57, 31.29, 10, 10, 10, 10, 10, 1.43]).forEach((w, i) => { if (w) ws.getColumn(i + 1).width = w; });
  const title = (r, text, font, h) => { ws.mergeCells('A' + r + ':J' + r); const c = ws.getCell('A' + r); c.value = text; c.font = Object.assign({ name: TF }, font); c.alignment = { horizontal: 'center' }; ws.getRow(r).height = h; };
  title(1, prop, { size: 12, bold: true }, 15.75);
  title(2, 'Mark-Taylor Residential, Inc.', { size: 10 }, 12.75);
  title(3, 'Trailing Profit And Loss Detail', { size: 10 }, 13.5);
  title(4, MON[Q.m - 1] + ' ' + Q.y + ' - Accrual - Accounting Book: Default', { size: 10, italic: true }, 12.75);
  const printed = new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles', month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit', second: '2-digit' }).replace(',', '');
  ws.getCell('A5').value = 'Printed ' + printed; ws.getCell('A5').font = { name: TF, size: 8 }; ws.getRow(5).height = 11.25;
  for (let c = 1; c <= 11; c++) { ws.getCell(5, c).border = { bottom: thin }; ws.getCell(6, c).border = { bottom: thin }; }
  ws.getCell('A6').value = 'Account'; ws.getCell('A6').font = { name: TF, size: 8, bold: true };
  months.map((p) => mLabel(p) + ' Actual').concat(['Adjusted Total', 'Variance']).forEach((h, i) => { const c = ws.getCell(6, 6 + i); c.value = h; c.font = { name: TF, size: 8, bold: true }; c.alignment = { horizontal: 'right', vertical: 'bottom', wrapText: true }; });
  ws.getRow(6).height = 24;
  if (P.head1) {   // exact title block from the template
    ws.unMergeCells('A1:J1'); ws.unMergeCells('A2:J2'); ws.unMergeCells('A3:J3'); ws.unMergeCells('A4:J4');
    for (let rr = 1; rr <= 6; rr++) {
      const pr = P['head' + rr]; if (!pr) continue;
      for (let c = 1; c <= 11; c++) { const cell = ws.getCell(rr, c); const keep = cell.value; cell.style = JSON.parse(JSON.stringify(pr.cells[c - 1] || {})); cell.value = keep != null ? keep : pr.values[c - 1]; }
      if (pr.height) ws.getRow(rr).height = pr.height;
      pr.merges.forEach(([c1, c2]) => { try { ws.mergeCells(rr, c1, rr, c2); } catch (e) { /* */ } });
    }
  }
  let r = 7, shadeN = 0;
  const hdr = (col, text, h = 12) => {
    const k = col === 2 ? 'sec' : col === 3 ? 'cat' : 'sub';
    if (!applyProf(r, k)) { ws.getCell(r, col).numFmt = '@'; ws.getCell(r, col).font = { name: TF, size: 8, bold: true }; ws.getRow(r).height = h; }
    ws.getCell(r, col).value = text; r++;
  };
  const nums = (vals, opts = {}) => vals.forEach((v, i) => { const c = ws.getCell(r, 6 + i); c.value = Math.round(v * 100) / 100; c.numFmt = NUMF; c.font = { name: TF, size: 7 }; c.alignment = { horizontal: 'right' }; if (opts.top) c.border = { top: thin }; if (opts.fill) c.fill = SHADE; });
  const line = (a, col, income) => {   // account line: 3 months, total, variance
    const v = (actual[a.no] && actual[a.no].vals) || [0, 0, 0]; const tot = v[0] + v[1] + v[2]; const bud = budgetQ[a.no] || 0;
    const variance = income ? tot - bud : bud - tot;
    const shaded = (shadeN++ % 2 === 1);
    const k = (col === 5 ? 'detE' : 'detD') + (shaded ? 'S' : '');
    const styled = applyProf(r, k) || applyProf(r, col === 5 ? 'detE' : 'detD');
    const c = ws.getCell(r, col); c.value = (actual[a.no] && actual[a.no].label) || a.label;
    if (!styled) { c.numFmt = '@'; c.font = { name: TF, size: 7.5 }; if (shaded) c.fill = SHADE; if (col === 4) ws.mergeCells(r, 4, r, 5); ws.getRow(r).height = 12; }
    else if (shaded && !P[k]) { for (let cc = col; cc <= 10; cc++) ws.getCell(r, cc).fill = SHADE; }
    if (styled) [v[0], v[1], v[2], tot, variance].forEach((x, i) => { ws.getCell(r, 6 + i).value = Math.round(x * 100) / 100; });
    else nums([v[0], v[1], v[2], tot, variance], { fill: shaded });
    r++;
    return { v: [v[0], v[1], v[2], tot], variance };
  };
  const addT = (acc, x) => { for (let i = 0; i < 4; i++) acc.v[i] += x.v[i]; acc.variance += x.variance; };
  const total = (label, acc, indent, h = 16.5, kind) => {
    const vals = acc.v.concat([acc.variance]);
    if (kind && applyProf(r, kind)) {
      ws.getCell(r, 1).value = label; vals.forEach((x, i) => { ws.getCell(r, 6 + i).value = Math.round(x * 100) / 100; }); r++; return;
    }
    ws.mergeCells(r, 1, r, 5); const c = ws.getCell(r, 1); c.value = label; c.numFmt = '@'; c.font = { name: TF, size: 8, bold: true }; c.alignment = { horizontal: 'left', indent };
    nums(vals, { top: true }); ws.getRow(r).height = h; r++;
  };
  const zero = () => ({ v: [0, 0, 0, 0], variance: 0 });
  const secTot = {}; let noiRow = null;
  for (const sec of outline) {
    const income = /^INCOME$/i.test(sec.name.trim());
    if (/^NON/i.test(sec.name) && noiRow) { const gaps = P.gapHeights || [null, null]; gaps.forEach((h) => { if (h) ws.getRow(r).height = h; r++; }); }
    hdr(2, sec.name); const sAcc = zero();
    for (const cat of sec.cats) {
      hdr(3, cat.name, 11.25); const cAcc = zero();
      for (const it of cat.items) {
        if (it.acct) { addT(cAcc, line(it.acct, 4, income)); continue; }
        hdr(4, it.sub.name, 12); const gAcc = zero();
        for (const a of it.sub.accts) addT(gAcc, line(a, 5, income));
        total('Total ' + it.sub.name, gAcc, 3, 16.5, 'totSub'); addT(cAcc, gAcc);
      }
      total('Total ' + cat.name, cAcc, 2, 16.5, 'totCat'); addT(sAcc, cAcc);
    }
    total('TOTAL ' + sec.name.toUpperCase(), sAcc, 1, 19.5, 'totSec'); secTot[sec.name.toUpperCase()] = sAcc;
    if (/^EXPENSE$/i.test(sec.name.trim()) && secTot.INCOME) {
      const ti = secTot.INCOME, noi = { v: ti.v.map((x, i) => x - sAcc.v[i]), variance: ti.variance + sAcc.variance };
      total('NET OPERATING INCOME', noi, 0, 19.5, 'noi'); noiRow = noi;
    }
  }
  const nonop = Object.keys(secTot).find((k) => /^NON/.test(k));
  const base = noiRow || zero();
  const ni = nonop ? { v: base.v.map((x, i) => x - secTot[nonop].v[i]), variance: base.variance + secTot[nonop].variance } : base;
  total('NET INCOME', ni, 0, 11.25, 'ni');
  r++;
  if (applyProf(r, 'footer')) ws.getCell(r, 1).value = '\u00a9 ResMan, LLC';
  else { ws.mergeCells(r, 1, r, 10); const f = ws.getCell(r, 1); f.value = '\u00a9 ResMan, LLC'; f.font = { name: 'Arial', size: 9, bold: true }; f.alignment = { horizontal: 'right' }; }
  return { wb, totals: { income: secTot.INCOME ? secTot.INCOME.v : null, net: ni.v }, added };
}


// ── workbook templating helpers (Midco covenant workbook = last quarter's, refreshed) ──
function copySheet(srcWb, src, dstWb, name, opt = {}) {
  const ws = dstWb.addWorksheet(name, { views: src.views, properties: Object.assign({}, src.properties), pageSetup: Object.assign({}, src.pageSetup) });
  ws.__noAutofit = true;
  (src.model.merges || []).forEach((m) => { try { ws.mergeCells(m); } catch (e) { /* overlap */ } });   // merge first, then styles, so borders survive
  for (let c = 1; c <= Math.max(src.columnCount || 0, 30); c++) { const col = src.getColumn(c); if (col.width) ws.getColumn(c).width = col.width; if (col.hidden) ws.getColumn(c).hidden = true; }
  src.eachRow({ includeEmpty: true }, (row, r) => {
    const out = ws.getRow(r); if (row.height) out.height = row.height; if (row.hidden) out.hidden = true;
    row.eachCell({ includeEmpty: true }, (cell, c) => {
      const d = ws.getCell(r, c); let v = cell.value;
      if (v && typeof v === 'object' && (v.sharedFormula || v.formula)) { const f = cell.formula; v = f ? { formula: f, result: cell.result } : cell.result; }
      d.value = v; d.style = JSON.parse(JSON.stringify(cell.style || {}));
    });
  });
  if (opt.images !== false) src.getImages().forEach((im) => {
    const m = srcWb.getImage(im.imageId); if (!m || !m.buffer) return;
    const id = dstWb.addImage({ buffer: m.buffer, extension: m.extension || 'png' });
    ws.addImage(id, { tl: im.range.tl, br: im.range.br, editAs: im.range.editAs || 'oneCell' });
  });
  return ws;
}
function copyRowStyle(ws, fromRow, toRow, cols = 12) {
  for (let c = 1; c <= cols; c++) ws.getCell(toRow, c).style = JSON.parse(JSON.stringify(ws.getCell(fromRow, c).style || {}));
  if (ws.getRow(fromRow).height) ws.getRow(toRow).height = ws.getRow(fromRow).height;
}
function imageSize(buf) {   // {w,h} in pixels for PNG / JPEG
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf[0] === 0xff && buf[1] === 0xd8) {
    let i = 2;
    while (i < buf.length) { if (buf[i] !== 0xff) { i++; continue; } const mk = buf[i + 1]; const len = buf.readUInt16BE(i + 2);
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + len; }
  }
  return { w: 900, h: 600 };
}
async function loadPrevCovenant(store, Q) {   // last quarter's "Midco Financial Covenant_*.xlsx" (highest _vN)
  const pq = Q.back(3); const dir = DEST_BASE + '/' + pq.y + '/Q' + Math.ceil(pq.m / 3) + '/Midco I';
  const kids = (await store.list('accounting', dir)) || [];
  const vnum = (n) => { const m = n.match(/_v(\d+)\.xlsx$/i); return m ? +m[1] : 0; };
  const c = kids.filter((k) => !k.folder && /financial covenant/i.test(k.name) && /\.xlsx$/i.test(k.name) && !/^~\$/.test(k.name))
    .sort((a, b) => (vnum(b.name) - vnum(a.name)) || b.name.localeCompare(a.name));
  if (!c.length) return null;
  const ExcelJS = require('exceljs'); const w = new ExcelJS.Workbook(); await w.xlsx.load(await store.read('accounting', dir + '/' + c[0].name));
  return w.getWorksheet('Calc') ? { wb: w, name: c[0].name } : null;
}
// lender-package file names, as in the Q2 2026 package
const FS_NAME = {
  SRN: (Q) => '00 A1 County Line SRN LLC - Quarterly Financial Statements ' + pad2(Q.m) + '.' + pad2(Q.d) + '.' + Q.y + '.pdf',
  CLIP: (Q) => '00 A1a County Line Industrial Park LLC - Quarterly Consolidated Financial Statements ' + pad2(Q.m) + '.' + Q.y + '.pdf',
  Buna: (Q) => '00 A2 CLR Buna Property Owner LLC Financial Statements Quarter Ended ' + MON[Q.m - 1] + ' ' + Q.d + ', ' + Q.y + '.pdf',
  Silsbee: (Q) => '00 A2 County Line Rail Silsbee LLC - Quarterly FS ' + pad2(Q.m) + '.' + Q.y + '.pdf',
};

// ════════════════════════════════════════════════════════════════════════════
// The report job
// ════════════════════════════════════════════════════════════════════════════
async function runReport(quarterEnd, cfg, log, opts = {}) {
  const cl = opts.cl && opts.cl.token ? opts.cl : null;
  const Q = parseQuarterEnd(quarterEnd);
  const store = makeStore(cfg);
  const out = { quarter_end: Q.iso, quarter: Q.label, storage: store.kind, dest: DEST_BASE + '/' + Q.y + '/Q' + Q.q, items: [], metrics: {}, warnings: [] };
  const say = (s) => { if (log) log(s); };
  const item = (pkg, file, status, note, url) => { out.items.push({ pkg, file, status, note: note || '', url: url || null }); say(pkg + ': ' + status + ' - ' + file + (note ? ' (' + note + ')' : '')); };
  const destPath = (dir, f) => out.dest + '/' + DIRS[dir] + '/' + f;
  async function put(pkg, dir, fname, buf, note) {
    let url;
    try { url = await store.write('accounting', destPath(dir, fname), buf); }
    catch (e) { if (/EBUSY|EPERM|locked/i.test(e.message)) throw new Error('the file is open (probably in Excel) - close it and click Report again'); throw e; }
    item(pkg, fname, 'created', note, url);
  }
  async function copy(pkg, dir, lib, src, fname, note) {
    if (!src) { item(pkg, fname, 'missing', note || 'source file not found'); return false; }
    const buf = await store.read(lib, src);
    const url = await store.write('accounting', destPath(dir, fname || path.posix.basename(src)), buf);
    item(pkg, fname || path.posix.basename(src), 'copied', 'from ' + lib + ':/' + src, url); return true;
  }
  async function safe(pkg, label, fn) { try { await fn(); } catch (e) { item(pkg, label, 'error', e.message); } }
  const QE = Q.months[2], sfr = cfg.sfr_base || DEFAULTS.sfr_base;
  const sfrMonth = (prop, p) => sfr + '/' + prop + '/' + pad2(p.m) + '.' + p.y;
  const asOf = pad2(Q.m) + '/' + pad2(Q.d) + '/' + Q.y;
  // date the certificates are prepared (Pacific time), shown on the Date: line
  const prepared = new Date().toLocaleDateString('en-US', { timeZone: 'America/Los_Angeles', month: '2-digit', day: '2-digit', year: 'numeric' });

  // ── 1. Skye Canyon ────────────────────────────────────────────────────────
  await safe('Skye Canyon', 'package', async () => {
    const dir = sfrMonth('Banyan Brighton', QE);
    const xl = await findFile(store, 'shared', dir, /^mergedexcel_/i, /\.xlsx$/i);
    const pdfF = await findFile(store, 'shared', dir, /^mergedexcel_/i, /\.pdf$/i);
    if (xl) {
      const rows = sheetRows(await store.read('shared', xl), '12 Month Income Statement');
      const hdr = rows[4] || []; const cols = Q.months.map((p) => hdr.findIndex((h) => lab(h) === mLabel(p)));
      if (cols.some((c) => c < 0)) throw new Error('quarter months not all in Cushman 12-month statement');
      const wb = newWb(); const ws = wb.addWorksheet('Report1');
      ws.getCell('A1').value = 'Banyan Brighton (res68831)'; ws.getCell('A2').value = 'Statement (12 months)';
      ws.getCell('A3').value = 'Period = ' + mLabel(Q.months[0]) + '-' + mLabel(QE); ws.getCell('A4').value = 'Book = Accrual ; Tree = ysi_is';
      Q.months.forEach((p, i) => { ws.getCell(5, 3 + i).value = mLabel(p); }); ws.getCell(5, 6).value = 'Total';
      for (let r = 5; r < rows.length; r++) {
        const R = rows[r] || []; const xr = r + 1;
        if (R[0] != null) ws.getCell(xr, 1).value = R[0]; if (R[1] != null) ws.getCell(xr, 2).value = R[1];
        const v = cols.map((c) => num(R[c]));
        if (v.some((x) => x != null)) {
          v.forEach((x, i) => { const c = ws.getCell(xr, 3 + i); c.value = x || 0; c.numFmt = NF; });
          setF(ws, 'F' + xr, `SUM(C${xr}:E${xr})`, v.reduce((a, b) => a + (b || 0), 0), NF);
        }
      }
      ws.getColumn(1).width = 11; ws.getColumn(2).width = 42;
      await put('Skye Canyon', 'skye', '12_Month_Statement_res68831_Accrual.xlsx', await wbBuffer(wb), 'from ' + path.posix.basename(xl));
    } else item('Skye Canyon', '12_Month_Statement_res68831_Accrual.xlsx', 'missing', 'no MergedExcel_*.xlsx in ' + dir);
    await copy('Skye Canyon', 'skye', 'shared', pdfF, 'Skye Canyon ' + MON[QE.m - 1] + ' ' + QE.y + ' Financial Reports.pdf');
    await put('Skye Canyon', 'skye', 'Certified Rent Roll_Skye Canyon_' + Q.tag + '.pdf', await stampCertificate('rentroll_skye', { asof: asOf, signdate: prepared }), 'unsigned');
  });

  // ── 2. South Mountain / Orchards (Preserve) ───────────────────────────────
  for (const [prop, short] of [['Banyan South Mountain', 'SM'], ['Banyan Preserve', 'Preserve']]) {
    await safe('South Mountain / Orchards', prop, async () => {
      const dir = sfrMonth(prop, QE);
      const pl = await findFile(store, 'shared', dir, new RegExp(prop.replace(/ /g, '\\s+') + '\\s+12 Month Profit and Loss', 'i'), /\.xlsx$/i);
      await copy('South Mountain / Orchards', 'smo', 'shared', await findFile(store, 'shared', dir, new RegExp(prop.replace(/ /g, '\\s+') + '\\s+Financial', 'i'), /\.pdf$/i));
      const fname = 'Banyan ' + short + ' - Q' + Q.q + ' T3.xlsx';
      if (!pl) { item('South Mountain / Orchards', fname, 'missing', 'no 12 Month Profit and Loss in ' + dir); return; }
      const rows = sheetRows(await store.read('shared', pl));
      const hdrRow = rows[4] || [];
      const cols = Q.months.map((p) => hdrRow.findIndex((h) => h instanceof Date ? (h.getUTCFullYear() === p.y && h.getUTCMonth() + 1 === p.m) : false));
      if (cols.some((c) => c < 0)) throw new Error('quarter months not all in Mark-Taylor 12-month P&L');
      // prior quarter's T3 = the layout template (ResMan export)
      const pq = Q.back(3), pQ = { y: pq.y, q: Math.ceil(pq.m / 3) };
      let tpl = null, tplBuf = null;
      for (const folder of ['South Mountain and Orchards', 'South Mountian and Orchards']) {
        const f2 = DEST_BASE + '/' + pQ.y + '/Q' + pQ.q + '/' + folder + '/Banyan ' + short + ' - Q' + pQ.q + ' T3.xlsx';
        try { tplBuf = await store.read('accounting', f2); tpl = sheetRows(tplBuf); break; } catch (e) { /* try next */ }
      }
      // budget for the quarter = YTD budget at quarter end - YTD budget at prior quarter end (0 at year start)
      const bvYtd = async (p) => {
        const d2 = sfrMonth(prop, p); const bf = await findFile(store, 'shared', d2, new RegExp(prop.replace(/ /g, '\\s+') + '\\s+Budget Variance', 'i'), /\.xlsx$/i);
        if (!bf) return null; const br = sheetRows(await store.read('shared', bf));
        const out2 = {}; Object.entries(mtrAccounts(br, [12])).forEach(([k, v]) => { out2[k] = v.vals[0]; }); return out2;
      };
      const bNow = await bvYtd(QE); const bPrev = pq.y === QE.y ? await bvYtd(pq) : {};
      const outline = tpl ? t3Outline(tpl) : null;
      if (outline && outline.length && bNow && bPrev) {
        const budgetQ = {}; new Set(Object.keys(bNow).concat(Object.keys(bPrev))).forEach((k) => { budgetQ[k] = (bNow[k] || 0) - (bPrev[k] || 0); });
        const actual = mtrAccounts(rows, cols);
        let styles = null; try { styles = await t3StyleProfiles(tplBuf); } catch (e) { out.warnings.push('South Mountain / Orchards: could not read formatting from last quarter\u2019s T3 (' + e.message + ')'); }
        const built = buildResmanT3({ prop, Q, outline, actual, budgetQ, months: Q.months, styles });
        // tie-out to Mark-Taylor's own totals
        const mtrInc = mtrRowByLabel(rows, /^TOTAL INCOME$/i, cols), mtrNet0 = mtrRowByLabel(rows, /^NET INCOME$/i, cols), capex = mtrRowByLabel(rows, /^Total Capital Expenditure$/i, cols) || [0, 0, 0];
        const mtrNet = mtrNet0 && mtrNet0.map((v, i) => v + capex[i]);   // Mark-Taylor deducts capex before NET INCOME; the ResMan T3 does not
        const off = (a, b2) => a && b2 && [0, 1, 2].some((i) => Math.abs(a[i] - b2[i]) > 0.05);
        let note = 'ResMan T3 layout from Q' + pQ.q + ' ' + pQ.y + '; ties to Mark-Taylor TOTAL INCOME and NET INCOME';
        if (off(built.totals.income, mtrInc) || off(built.totals.net, mtrNet)) {
          note = 'ResMan T3 layout - WARNING: totals do not tie to Mark-Taylor\u2019s 12-month P&L';
          out.warnings.push('South Mountain / Orchards: ' + fname + ' totals do not tie to the Mark-Taylor 12-month P&L - check the account mapping');
        }
        if (built.added.length) out.warnings.push('South Mountain / Orchards: ' + fname + ' - new account(s) not in last quarter\u2019s T3, placed next to their neighbors: ' + built.added.join(', '));
        await put('South Mountain / Orchards', 'smo', fname, await wbBuffer(built.wb), note);
        return;
      }
      out.warnings.push('South Mountain / Orchards: ' + fname + ' - prior quarter T3 or budget report not found, used the simple layout');
      const hdr = rows[4] || [];
      const cols2 = Q.months.map((p) => hdr.findIndex((h) => h instanceof Date ? (h.getUTCFullYear() === p.y && h.getUTCMonth() + 1 === p.m) : false));
      const wb = newWb(); const ws = wb.addWorksheet('Sheet');
      ws.getCell('A1').value = prop; ws.getCell('A2').value = 'Mark-Taylor Residential, Inc.'; ws.getCell('A3').value = 'Trailing Profit And Loss Detail (T3)';
      ws.getCell('A4').value = MON[QE.m - 1] + ' ' + QE.y + ' - Accrual  |  ' + mLabel(Q.months[0]) + '-' + mLabel(QE) + ' columns from Mark-Taylor Twelve Month Profit and Loss (MTR)';
      ws.getCell('A6').value = 'Account';
      Q.months.forEach((p, i) => { ws.getCell(6, 6 + i).value = mLabel(p) + ' Actual'; }); ws.getCell(6, 9).value = 'Adjusted Total';
      let xr = 7;
      for (let r = 5; r < rows.length; r++) {
        const R = rows[r] || []; const labels = R.slice(0, 5); const v = cols.map((c) => num(R[c]));
        if (labels.every((x) => x == null) && v.every((x) => x == null)) continue;
        labels.forEach((x, i) => { if (x != null) ws.getCell(xr, 1 + i).value = x; });
        if (v.some((x) => x != null)) {
          v.forEach((x, i) => { const c = ws.getCell(xr, 6 + i); c.value = x || 0; c.numFmt = NF; });
          setF(ws, 'I' + xr, `SUM(F${xr}:H${xr})`, v.reduce((a, b) => a + (b || 0), 0), NF);
        }
        xr++;
      }
      ws.getColumn(4).width = 36;
      await put('South Mountain / Orchards', 'smo', fname, await wbBuffer(wb), 'from ' + path.posix.basename(pl));
    });
  }
  await safe('South Mountain / Orchards', 'rent rolls', async () => {
    await put('South Mountain / Orchards', 'smo', 'Certified Rent Roll_Orchards ' + Q.tag + '.pdf', await stampCertificate('rentroll_orchards', { asof: asOf, signdate: prepared }), 'unsigned');
    await put('South Mountain / Orchards', 'smo', 'Certified Rent Roll_South Mountain ' + Q.tag + '.pdf', await stampCertificate('rentroll_sm', { asof: asOf, signdate: prepared }), 'unsigned');
  });

  // ── 3. Silo Ridge - JPM DSCR ──────────────────────────────────────────────
  await safe('Silo Ridge - JPM', 'DSCR calc', async () => {
    const data = new Map(); const orders = [];
    for (let i = 0; i < 3; i++) {
      const p = Q.months[i]; const dir = sfrMonth('Banyan Silo Ridge', p);
      const f = await findFile(store, 'shared', dir, /variance report/i, /\.xls[xm]$/i);
      if (!f) throw new Error('no Variance Report for ' + mLabel(p) + ' in ' + dir);
      const rows = sheetRows(await store.read('shared', f), 'IS Scoping');
      const seen = {}; const od = [];
      for (let r = 26; r < Math.min(rows.length, 260); r++) {
        const R = rows[r] || []; const desc = lab(R[3]); if (!desc) continue;
        const base = lab(R[2]) + '|' + desc; const n = seen[base] || 0; seen[base] = n + 1; const k = base + '|' + n;
        if (!data.has(k)) data.set(k, { code: lab(R[2]), desc, v: [null, null, null] });
        data.get(k).v[i] = num(R[4]); od.push(k);
      }
      orders.push(od);
    }
    const final = orders[2].slice();
    for (const od of orders.slice(0, 2)) od.forEach((k, idx) => {
      if (final.includes(k)) return; let pos = final.length;
      for (let j = idx - 1; j >= 0; j--) if (final.includes(od[j])) { pos = final.indexOf(od[j]) + 1; break; }
      final.splice(pos, 0, k);
    });
    const wb = newWb(); const calc = wb.addWorksheet('Calc'); const tabName = Q.label + ' (jpm_is)'; const ws = wb.addWorksheet(tabName);
    ws.getCell('A1').value = 'Silo Ridge (p0005057)'; ws.getCell('A2').value = 'Statement - ' + Q.label + ' (monthly PTD actuals)';
    ws.getCell('A3').value = 'Period = ' + mLabel(Q.months[0]) + '-' + mLabel(QE);
    ws.getCell('A4').value = 'Book = Accrual ; Tree = jpm_is  |  Source: Greystar monthly Variance Reports (IS Scoping = Yardi Budget Comparison, PTD Actual)';
    Q.months.forEach((p, i) => { ws.getCell(5, 3 + i).value = mLabel(p); }); ws.getCell(5, 6).value = 'Total';
    const rowOf = {}; const tot = {};
    final.forEach((k, n) => {
      const r = 6 + n, d = data.get(k);
      if (d.code) ws.getCell(r, 1).value = d.code; ws.getCell(r, 2).value = d.desc;
      if (d.v.some((x) => x != null)) {
        d.v.forEach((x, i) => { const c = ws.getCell(r, 3 + i); c.value = x || 0; c.numFmt = NF; });
        const s = d.v.reduce((a, b) => a + (b || 0), 0); setF(ws, 'F' + r, `SUM(C${r}:E${r})`, s, NF);
        if (['Total Revenue', 'Total Operating Expenses', 'Total Interest Expense'].includes(d.desc) && !(d.desc in rowOf)) { rowOf[d.desc] = r; tot[d.desc] = s; ws.getRow(r).font = { bold: true }; }
      }
    });
    for (const k of ['Total Revenue', 'Total Operating Expenses', 'Total Interest Expense']) if (!(k in rowOf)) throw new Error('"' + k + '" row not found in variance reports');
    ws.getColumn(1).width = 13; ws.getColumn(2).width = 44; [3, 4, 5, 6].forEach((c) => { ws.getColumn(c).width = 15; });
    const rev = tot['Total Revenue'], opex = tot['Total Operating Expenses'], int = tot['Total Interest Expense'];
    const noi = rev - opex, dscr = int ? noi / int : null;
    // Calc tab - same layout/format as the hand-built JPM DSCR workbook
    const T = qs(tabName);
    const NUM0 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
    const BLUE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    const GRAY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
    const thin = { style: 'thin' };
    const put2 = (addr, v, o = {}) => {
      const c = calc.getCell(addr);
      c.value = v && v.f ? { formula: v.f, result: v.r } : v;
      c.font = { name: 'Calibri', size: 11, bold: !!o.b, italic: !!o.i, color: o.white ? { argb: 'FFFFFFFF' } : undefined };
      if (o.nf) c.numFmt = o.nf; if (o.fill) c.fill = o.fill; if (o.center) c.alignment = { horizontal: 'center' };
      return c;
    };
    put2('B2', 'Silo Ridge', { b: true, white: true, fill: BLUE });
    put2('C2', Q.label, { b: true, white: true, fill: BLUE, center: true });
    put2('B3', 'Total Revenue', { b: true });
    put2('C3', { f: `SUM(${T}!C${rowOf['Total Revenue']}:E${rowOf['Total Revenue']})`, r: rev }, { b: true, nf: NUM0 });
    put2('B4', 'Less: Operating Expenses', { i: true });
    put2('C4', { f: `-SUM(${T}!C${rowOf['Total Operating Expenses']}:E${rowOf['Total Operating Expenses']})`, r: -opex }, { nf: NUM0 });
    put2('B5', 'Add back: Reserve Account Balance', { i: true }); put2('C5', 0, { nf: NUM0 });
    put2('B6', 'Total Net Operating Income', { b: true });
    put2('C6', { f: 'SUM(C3:C5)', r: noi }, { b: true, nf: NUM0 });
    put2('B8', 'Cash Interest Expense');
    put2('C8', { f: `SUM(${T}!C${rowOf['Total Interest Expense']}:E${rowOf['Total Interest Expense']})`, r: int }, { nf: NUM0 });
    put2('B9', 'Principal Payments Paid'); put2('C9', 0, { nf: NUM0 });
    put2('B10', 'Total Debt Service', { b: true });
    put2('C10', { f: '+C8+C9', r: int }, { nf: NUM0 });
    put2('B12', 'DSCR');
    put2('C12', { f: 'C6/C10', r: dscr }, { nf: '0.00', fill: GRAY });
    put2('C13', null, { fill: GRAY });
    calc.getRow(5).hidden = true; calc.getRow(9).hidden = true;   // kept for the formulas, hidden as in the original
    // borders: outer box B2:C13, rule above the totals and above DSCR
    for (let r = 2; r <= 13; r++) {
      const bc = calc.getCell('B' + r), cc = calc.getCell('C' + r);
      const top = (r === 2 || r === 6 || r === 10 || r === 12) ? thin : undefined;
      const bottom = r === 13 ? thin : undefined;
      bc.border = { left: thin, top, bottom };
      cc.border = { right: thin, top, bottom: bottom || ((r === 4 || r === 5 || r === 8 || r === 9) ? thin : undefined) };
    }
    calc.getColumn(1).width = 13; calc.getColumn(2).width = 40.18; calc.getColumn(3).width = 16.18;
    calc.views = [{ zoomScale: 120 }];
    out.metrics.silo_jpm = { revenue: rev, opex, noi, interest: int, dscr };
    const xlsxName = 'Silo Ridge DSCR Calc_' + Q.tag + '.xlsx';
    const xlsxBuf = await wbBuffer(wb);
    await put('Silo Ridge - JPM', 'jpm', xlsxName, xlsxBuf, 'DSCR ' + (dscr ? dscr.toFixed(2) : 'n/a') + 'x');
    // draft email to JPM / Greystar with the calc attached (same recipients as the prior quarter's reply)
    const em = Object.assign({}, DEFAULTS.silo_jpm_email, cfg.silo_jpm_email || {});
    const subject = "Silo Ridge - Q" + Q.q + "'" + Q.y + ' Debt Covenant';
    const msgLine = 'Please see the attached ' + Q.label + ' debt covenant (DSCR) calculation for Silo Ridge.';
    const text = 'Hi Team,\r\n\r\n' + msgLine + '\r\n\r\nThanks.\r\n\r\n' + SIGNATURE_TEXT.join('\r\n') + '\r\n';
    const html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">Hi Team,<br><br>' + msgLine + '<br><br>Thanks.<br><br>' +
      '<b>' + SIGNATURE_TEXT[0] + '</b><br>' + SIGNATURE_TEXT.slice(1).join('<br>') + '</div>';
    await put('Silo Ridge - JPM', 'jpm', subject + '.eml', buildDraftEml({ from: em.from, to: em.to, cc: em.cc, subject, text, html,
      attachments: [{ name: xlsxName, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data: xlsxBuf }] }),
      'draft email - open in Outlook, review and send');
  });

  // ── 4. Silo Ridge - Wells Fargo ───────────────────────────────────────────
  await safe('Silo Ridge - Wells Fargo', 'financial statements', async () => {
    const dir = sfrMonth('Banyan Silo Ridge', QE);
    const f = await findFile(store, 'shared', dir, /operating report/i, /\.pdf$/i);
    const fname = 'Silo Ridge ' + MON[QE.m - 1] + ' ' + QE.y + ' Financial Statements (Greystar).pdf';
    if (!f) { item('Silo Ridge - Wells Fargo', fname, 'missing', 'no Operating Report PDF in ' + dir); return; }
    const buf = await store.read('shared', f); const texts = await pdfPagesText(buf);
    const per = 'Period = ' + mLabel(QE);
    const keep = [];
    texts.forEach((t, i) => {
      // Yardi report titles are drawn last, so they appear at the END of the page text
      if (!t || /Table of Contents/i.test(t)) return; const tail = t.slice(-500);
      if ((/Balance Sheet \(With Period Change\)/.test(tail) && tail.includes(per)) ||
          (/\bIncome Statement\b/.test(tail) && !/Scoping|Budget Comparison/.test(tail) && tail.includes(per) && /Tree = ysi_is/.test(tail)) ||
          (/12 Month Actual to Budget/.test(tail))) keep.push(i);
    });
    if (!keep.length) { item('Silo Ridge - Wells Fargo', fname, 'missing', 'statement pages not found in ' + path.posix.basename(f)); return; }
    const { PDFDocument } = require('pdf-lib');
    const src = await PDFDocument.load(buf, { ignoreEncryption: true }); const dst = await PDFDocument.create();
    (await dst.copyPages(src, keep)).forEach((pg) => dst.addPage(pg));
    await put('Silo Ridge - Wells Fargo', 'wf', fname, Buffer.from(await dst.save()), keep.length + ' pages from ' + path.posix.basename(f));
    out.warnings.push('Wells Fargo: Greystar sends PDFs only - if WF wants the Yardi Excel exports (12-month, BS period change, IS for the quarter), request them from Greystar.');
  });
  await safe('Silo Ridge - Wells Fargo', 'certificate', async () => {
    await put('Silo Ridge - Wells Fargo', 'wf', 'Silo Ridge Compliance Certificate ' + Q.label + '.pdf', await stampCertificate('wf_silo', {
      testdate: ord(Q.d) + ' day of ' + MON[Q.m - 1] + ' ' + Q.y + '.',
      delivered: '____ day of ________________, ' + Q.y + '.',
      __override: { testdate: { x: 162.7, w: 397 } },
    }), 'unsigned');
  });

  // ── 5. Midco I ────────────────────────────────────────────────────────────
  const mc = Object.assign({}, DEFAULTS.midco, cfg.midco || {});
  const ents = DEFAULTS.midco.entities;   // regexes aren't JSON-serializable; keep code-defined
  // 5a. entity financial statements
  for (const e of ents) {
    await safe('Midco I', e.key + ' FS', async () => {
      const ent = clEntity(cl, e.key, 'fs');
      if (!ent) {   // CloudLedger not reachable from this run - fall back to the SharePoint folders
        const dir = await findMonthDir(store, 'cla', e.fin, QE.y, QE.m);
        const f = await findFile(store, 'cla', dir, e.fs, /\.pdf$/i);
        if (f) await copy('Midco I', 'midco', 'cla', f, FS_NAME[e.key](Q));
        else item('Midco I', e.key + ' quarterly financial statements', 'missing', cl ? 'no matching CloudLedger entity' : 'CloudLedger not available and not found in ' + (dir || e.fin));
        return;
      }
      await put('Midco I', 'midco', FS_NAME[e.key](Q),
        await clQuarterlyFs(cl, ent, Q.iso), 'generated from CloudLedger (' + ent.name + ')');
    });
  }
  await safe('Midco I', 'CLRFI Midco I consolidated FS', async () => {
    const ent = clEntity(cl, 'Midco', 'fs');
    if (!ent) {
      const f = await findFile(store, 'cla', '61 Midco I/00 Financials/' + QE.y, new RegExp('consolidated financial statements ' + pad2(QE.m) + '\\.' + QE.y, 'i'), /\.pdf$/i);
      if (f) await copy('Midco I', 'midco', 'cla', f); else item('Midco I', 'CLRFI Midco I consolidated FS', 'missing', cl ? 'no matching CloudLedger entity' : 'CloudLedger not available');
      return;
    }
    await put('Midco I', 'midco', 'CLRFI Midco I LLC - Consolidated Financial Statements ' + pad2(QE.m) + '.' + QE.y + '.pdf',
      await clQuarterlyFs(cl, ent, Q.iso), 'generated from CloudLedger (' + ent.name + ', lender package)');
  });
  await safe('Midco I', 'CLRF quarterly FS', async () => {
    const ent = clEntity(cl, 'CLRF', 'fs');
    if (!ent) {
      const f = await findFile(store, 'cla', '30 County Line Rail Fund/Financial Statements/' + QE.y + ' Q' + Q.q + ' Financials', /FS SCF Wksht_Final/i, /\.pdf$/i);
      if (f) await copy('Midco I', 'midco', 'cla', f); else item('Midco I', 'CLRF quarterly FS', 'missing', cl ? 'no matching CloudLedger entity' : 'CloudLedger not available');
      return;
    }
    await put('Midco I', 'midco', 'CLRF - Q' + Q.q + ' ' + QE.y + ' FS SCF Wksht_Final.pdf',
      await clFundFs(cl, ent, Q.iso), 'generated from CloudLedger (' + ent.name + ' fund statements)');
  });
  await safe('Midco I', 'certificate', async () => {
    await put('Midco I', 'midco', 'CLRFI Blank Compliance Certificate_' + Q.label.replace(/^Q(\d) /, 'Q$1 ') + '.pdf',
      await stampCertificate('bot_clrfi', { qend: MON[Q.m - 1] + ' ' + ord(Q.d), yy: String(Q.y).slice(2) }), 'quarter-end stamped, unsigned');
  });

  // 5b. covenant workbook
  await safe('Midco I', 'covenant workbook', async () => {
    const N = Math.max(1, +mc.dscr_months || 6), ann = 12 / N;
    const win = []; for (let k = N - 1; k >= 0; k--) win.push(Q.back(k));
    const wb = newWb();
    // last quarter's covenant workbook is the template: Calc / Appraisals / NonRecurring Expenses are copied
    // exactly (formatting, pictures, history) and refreshed; the Interest Reserve tab gets this quarter's statement.
    let tplCov = null; try { tplCov = await loadPrevCovenant(store, Q); } catch (e) { out.warnings.push('Midco: could not read last quarter\u2019s covenant workbook (' + e.message + ')'); }
    const scratchWb = newWb(); const tplSheets = {}; let tplCalc = null;
    if (tplCov) {
      if (tplCov.wb._themes) wb._themes = tplCov.wb._themes;   // same theme colors as the template
      tplCalc = copySheet(tplCov.wb, tplCov.wb.getWorksheet('Calc'), wb, 'Calc');
      for (const n of ['Appraisals', 'Interest Reserve', 'NonRecurring Expenses']) {
        const sh = tplCov.wb.getWorksheet(n); if (sh) tplSheets[n] = copySheet(tplCov.wb, sh, wb, n, { images: n !== 'Interest Reserve' });
      }
    }
    // without a template the Calc tab is built from scratch; with one, the formulas are computed on a scratch sheet first
    const calc = tplCov ? scratchWb.addWorksheet('Calc') : wb.addWorksheet('Calc');
    const pending = [];
    const res = {};
    // T12 per entity
    for (const e of ents) {
      let f = null, srcMonth = null, rows = null, tab = null;
      const clEnt = clEntity(cl, e.key);
      if (clEnt) {   // primary source: the CloudLedger GL
        try { rows = await clT12Rows(cl, clEnt, Q.iso); tab = (e.key + ' T12 ' + mLabel(QE)).slice(0, 31); f = 'CloudLedger: ' + clEnt.name; }
        catch (err) { out.warnings.push('Midco: ' + e.key + ' T12 from CloudLedger failed (' + err.message + ') - fell back to the SharePoint T12'); rows = null; }
      }
      if (!rows) {   // fallback: latest T12 export in the SharePoint folders
        for (let k = 0; k < 3 && !f; k++) {
          const p = Q.back(k); const dir = await findMonthDir(store, 'cla', e.fin, p.y, p.m);
          f = await findFile(store, 'cla', dir, /trailing.?12|\bT12\b/i, /\.xlsx$/i); if (f) srcMonth = p;
        }
      }
      if (!rows && !f) { pending.push(e.key + ': no T12 found'); res[e.key] = null; continue; }
      if (!rows) { rows = sheetRows(await store.read('cla', f)); tab = (e.key + ' T12 ' + mLabel(srcMonth)).slice(0, 31); }
      const ws = wb.addWorksheet(tab);
      rows.forEach((R, r) => (R || []).forEach((v, c) => { if (v != null) { const cell = ws.getCell(r + 1, c + 1); cell.value = v instanceof Date ? v : v; if (typeof v === 'number') cell.numFmt = NF; } }));
      ws.getColumn(1).width = 40;
      const H = findMonthHeader(rows); const hRow = H.row >= 0 ? H.row : 4; let extra = Math.max(0, ...rows.map((R) => (R || []).length));
      const find = (re) => rows.findIndex((R) => R && re.test(lab(R[0])));
      const rRev = find(/^gross profit$/i) >= 0 ? find(/^gross profit$/i) : find(/^total revenue$/i);
      const rOpx = find(/^total operating expenses$/i), rDep = find(/^depreciation expense$/i), rAmo = find(/^amortization expense$/i);
      if (rRev < 0 || rOpx < 0) throw new Error(e.key + ' T12: revenue/opex rows not found');
      const cells = []; let rev = 0, opx = 0;
      for (const p of win) {
        let ci = H.map[p.y + '-' + p.m];
        if (ci == null) {   // month not closed yet - add a highlighted placeholder column
          ci = extra++; ws.getCell(hRow + 1, ci + 1).value = mLabel(p) + ' (PENDING)'; ws.getCell(hRow + 1, ci + 1).fill = YEL;
          [rRev, rOpx, rDep, rAmo].filter((x) => x >= 0).forEach((x) => { const c = ws.getCell(x + 1, ci + 1); c.value = 0; c.fill = YEL; });
          pending.push(e.key + ' ' + mLabel(p) + ' (not closed - enter on tab "' + tab + '")');
        }
        const g = (r) => (r >= 0 ? num((rows[r] || [])[ci]) || 0 : 0);
        rev += g(rRev); opx += g(rOpx) - g(rDep) - g(rAmo);
        cells.push(colL(ci + 1));
      }
      const ref = (r) => cells.map((c) => qs(tab) + '!' + c + (r + 1)).join(',');
      res[e.key] = { tab, rev, opx, fRev: `SUM(${ref(rRev)})`, fOpx: `-(SUM(${ref(rOpx)})` + (rDep >= 0 ? `-SUM(${ref(rDep)})` : '') + (rAmo >= 0 ? `-SUM(${ref(rAmo)})` : '') + ')', src: f };
    }
    // interest allocations
    const ia = wb.addWorksheet('Interest Allocation');
    ia.getCell('A1').value = 'CLRFI Midco I - BOT loan interest allocation by month (cash interest)'; ia.getRow(3).values = ['Month', 'CLIP', 'Silsbee', 'SRN', 'Buna', 'Total', 'Source'];
    ia.getRow(3).font = { bold: true };
    const intTot = { CLIP: 0, Silsbee: 0, SRN: 0, Buna: 0 };
    win.forEach((p, i) => {
      const r = 4 + i; ia.getCell(r, 1).value = mLabel(p);
    });
    const clIaEnts = cl ? ['CLIP', 'Silsbee', 'SRN', 'Buna'].map((k) => [k, clEntity(cl, k)]) : [];
    if (cl && clIaEnts.every(([, e2]) => e2)) {
      // primary: interest per entity per month straight from the CloudLedger GL
      ia.getCell('A1').value = 'CLRFI Midco I - BOT loan cash interest by month, from the CloudLedger GL (interest expense + capitalized construction-period interest)';
      const accts = {};
      for (let i = 0; i < win.length; i++) {
        const p = win[i], r = 4 + i;
        const from = p.y + '-' + pad2(p.m) + '-01', to = new Date(Date.UTC(p.y, p.m, 0)).toISOString().slice(0, 10);
        let rowTot = 0;
        for (let j = 0; j < clIaEnts.length; j++) {
          const [k, e2] = clIaEnts[j];
          const rows = await clBalances(cl, e2, { from, to });
          const hits = (rows || []).filter((x) => isInterestExp(x) || isCapInterest(x));
          // expense accounts: net debits; capitalized interest: additions only (credits are reclasses to fixed assets)
          const v = hits.reduce((a, x) => a + (isCapInterest(x) ? (+x.total_debit || 0) : ((+x.total_debit || 0) - (+x.total_credit || 0))), 0);
          hits.forEach((x) => { (accts[k] = accts[k] || new Set()).add(x.code + ' ' + x.name); });
          const c = ia.getCell(r, 2 + j); c.value = v; c.numFmt = NF; intTot[k] += v; rowTot += v;
        }
        setF(ia, 'F' + r, `SUM(B${r}:E${r})`, rowTot, NF); ia.getCell(r, 7).value = 'CloudLedger GL ' + from + ' to ' + to;
      }
      let ar = 5 + win.length; ia.getCell(ar, 1).value = 'Accounts used'; ia.getCell(ar, 1).font = { bold: true };
      Object.entries(accts).forEach(([k, set]) => { ar++; ia.getCell(ar, 1).value = k; ia.getCell(ar, 2).value = [...set].join('; '); });
    } else {
      for (let i = 0; i < win.length; i++) {
        const p = win[i], r = 4 + i;
        const dir = await findMonthDir(store, 'cla', '61 Midco I/Loan Bank Statements', p.y, p.m);
        const f = await findFile(store, 'cla', dir, /interest allocation/i, /\.xlsx$/i);
        if (!f) { pending.push('Interest allocation ' + mLabel(p) + ' not found'); ia.getCell(r, 7).value = 'MISSING'; ia.getRow(r).fill = YEL; continue; }
        const rows = sheetRows(await store.read('cla', f)); const got = {};
        rows.slice(0, 40).forEach((R) => { const k = lab((R || [])[1]); if (k in intTot && !(k in got)) got[k] = num(R[3]) || 0; });
        ['CLIP', 'Silsbee', 'SRN', 'Buna'].forEach((k, j) => { const c = ia.getCell(r, 2 + j); c.value = got[k] || 0; c.numFmt = NF; intTot[k] += got[k] || 0; });
        setF(ia, 'F' + r, `SUM(B${r}:E${r})`, Object.values(got).reduce((a, b) => a + b, 0), NF); ia.getCell(r, 7).value = path.posix.basename(f);
      }
    }
    const iaLast = 3 + win.length; ia.getColumn(1).width = 12; ia.getColumn(7).width = 50; [2, 3, 4, 5, 6].forEach((c) => { ia.getColumn(c).width = 15; });
    // loan balance at quarter end (BOT statements bracketing the quarter end)
    let loan = null, loanSrc = '';
    const clLnEnts = cl ? ['CLIP', 'Silsbee', 'SRN', 'Buna'].map((k) => [k, clEntity(cl, k)]) : [];
    if (cl && clLnEnts.every(([, e2]) => e2)) {
      // primary: BOT loan payable (25063) on each borrower's balance sheet in CloudLedger
      const parts = [];
      for (const [k, e2] of clLnEnts) {
        const rows = await clBalances(cl, e2, { as_of: Q.iso });
        const v = (rows || []).filter(isBotLoan).reduce((a, x) => a + ((+x.total_credit || 0) - (+x.total_debit || 0)), 0);
        if (v) parts.push(k + ' ' + v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
        loan = (loan || 0) + v;
      }
      loanSrc = 'CloudLedger GL - Loan Payable BOT (25063) at ' + Q.iso + ': ' + parts.join(', ');
      if (!loan) { loan = null; out.warnings.push('Midco: no BOT loan balance (25063) found in CloudLedger at ' + Q.iso + ' - checked the BOT loan statements instead'); }
    }
    if (loan == null) {   // fallback: BOT loan statements in SharePoint
      for (const p of [Q.back(-1), QE]) {
        const dir = await findMonthDir(store, 'cla', '61 Midco I/Loan Bank Statements', p.y, p.m); if (!dir || loan) continue;
        const kids = (await store.list('cla', dir)) || [];
        for (const k of kids.filter((x) => !x.folder && /loan/i.test(x.name) && /\.pdf$/i.test(x.name))) {
          const t = (await pdfPagesText(await store.read('cla', dir + '/' + k.name))).join(' ');
          const re = /(\d\d)-(\d\d)-(\d\d)\s*(\d\d)-(\d\d)-(\d\d)\s*(\d+)\s*\$([\d,]+\.\d\d)/g; let m;
          while ((m = re.exec(t))) {
            const from = '20' + m[3] + '-' + m[1] + '-' + m[2], to = '20' + m[6] + '-' + m[4] + '-' + m[5];
            if (from <= Q.iso && Q.iso < to) { loan = +m[8].replace(/,/g, ''); loanSrc = k.name + ' (accrual ' + from + ' to ' + to + ')'; break; }
          }
          if (loan) break;
        }
      }
    }
    if (loan == null) pending.push('Loan balance at ' + Q.iso + ' not found in BOT loan statements - enter on "Loan Balance" tab');
    const lb = wb.addWorksheet('Loan Balance'); lb.getCell('A1').value = 'BOT loan principal balance at ' + Q.iso; const lbc = lb.getCell('B3'); lbc.value = loan || 0; lbc.numFmt = NF; if (loan == null) lbc.fill = YEL;
    lb.getCell('A3').value = 'Principal balance'; lb.getCell('A5').value = 'Source: ' + (loanSrc || 'NOT FOUND'); lb.getColumn(1).width = 22; lb.getColumn(2).width = 18;
    // interest reserve (acct 13100) per entity TB at quarter end
    const ir = (tplCov ? scratchWb : wb).addWorksheet('Interest Reserve'); ir.getCell('A1').value = 'Interest Reserve (acct 13100) at ' + Q.iso; ir.getRow(3).values = ['Entity', 'Balance', 'Source']; ir.getRow(3).font = { bold: true };
    let irTot = 0, irRow = 4;
    for (const e of ents) {
      const clEnt = clEntity(cl, e.key);
      if (clEnt) {
        let bal = null, ok = false;
        try {
          const rows = await (await clFetch(cl, '/api/entities/' + clEnt.id + '/balances?as_of=' + Q.iso)).json(); ok = true;
          const hit = (rows || []).filter((r) => /^13100\b/.test(String(r.code)));
          if (hit.length) bal = hit.reduce((a, r) => a + ((+r.total_debit || 0) - (+r.total_credit || 0)), 0);
        } catch (err) { out.warnings.push('Midco: ' + e.key + ' interest reserve from CloudLedger failed (' + err.message + ')'); }
        if (ok) {   // CloudLedger answered: no 13100 account means no reserve
          if (bal == null) bal = 0;
          ir.getCell(irRow, 1).value = e.key; const c = ir.getCell(irRow, 2); c.value = bal; c.numFmt = NF;
          ir.getCell(irRow, 3).value = 'CloudLedger GL - ' + clEnt.name + ', acct 13100 at ' + Q.iso;
          irTot += bal; irRow++; continue;
        }
      }
      const dir = await findMonthDir(store, 'cla', e.fin, QE.y, QE.m);
      let f = await findFile(store, 'cla', dir, /(\bTB\b|trial balance)/i, /\.xlsx$/i), draft = false;
      if (!f && e.req) { const d2 = await findMonthDir(store, 'cla', e.req, QE.y, QE.m); f = await findFile(store, 'cla', d2, /TB_Activity/i, /\.xlsx$/i); draft = !!f; }
      let bal = null;
      if (f) {
        const rows = sheetRows(await store.read('cla', f));
        const R = rows.find((x) => x && x.some((v) => /^13100\b/.test(lab(v))));
        if (R) { const nums = R.map(num).filter((x) => x != null && Math.abs(x) >= 0.01 && !/^13100$/.test(String(x))); bal = nums.length ? nums[nums.length - 1] : null; }
      }
      ir.getCell(irRow, 1).value = e.key; const c = ir.getCell(irRow, 2); c.value = bal || 0; c.numFmt = NF;
      ir.getCell(irRow, 3).value = f ? path.posix.basename(f) + (draft ? ' (DRAFT requisition TB - month not closed)' : '') : 'NOT FOUND';
      if (bal == null || draft) { c.fill = YEL; pending.push('Interest reserve ' + e.key + (bal == null ? ' not found' : ' from draft TB')); }
      irTot += bal || 0; irRow++;
    }
    setF(ir, 'B' + irRow, `SUM(B4:B${irRow - 1})`, irTot, NF); ir.getCell(irRow, 1).value = 'Total'; ir.getRow(irRow).font = { bold: true };
    ir.getColumn(1).width = 12; ir.getColumn(2).width = 18; ir.getColumn(3).width = 60;
    const irRef = qs('Interest Reserve') + '!B' + irRow;
    // appraisals
    const ap = (tplCov ? scratchWb : wb).addWorksheet('Appraisals'); ap.getCell('A1').value = mc.appraisal_label || 'Appraisals';
    const apKeys = ['CLIP', 'Silsbee', 'Buna', 'SRN']; apKeys.forEach((k, i) => { ap.getCell(3 + i, 1).value = k; const c = ap.getCell(3 + i, 2); c.value = (mc.appraisals || {})[k] || 0; c.numFmt = '#,##0'; });
    let apTot = apKeys.reduce((a, k) => a + ((mc.appraisals || {})[k] || 0), 0);
    ap.getCell(7, 1).value = 'Total'; setF(ap, 'B7', 'SUM(B3:B6)', apTot, '#,##0'); ap.getColumn(1).width = 14; ap.getColumn(2).width = 16;
    if (tplSheets.Appraisals) { const j = tplSheets.Appraisals.getCell('J16').value; const jv = j && typeof j === 'object' ? j.result : j; if (typeof jv === 'number') apTot = jv; }
    // interest reserve = ending balance of the bank's reserve statement uploaded on the Compliance Certificates page
    const irUp = (opts.uploads || []).find((u) => u.kind === 'interest_reserve');
    if (irUp && irUp.balance != null && isFinite(+irUp.balance)) irTot = +irUp.balance;
    else pending.push('Interest reserve statement not uploaded for ' + Q.iso + ' - used CloudLedger 13100 balances');
    let irBalRow = 33;
    {
      const irw = tplSheets['Interest Reserve'] || null;
      if (irw) {
        let top = 2, bottom = 28;
        if (irUp && irUp.image) {
          const sz = imageSize(irUp.image); const W = 941; const w = Math.min(W, sz.w); const h = Math.round(sz.h * (w / sz.w));
          const id = wb.addImage({ buffer: irUp.image, extension: irUp.ext || 'png' });
          irw.addImage(id, { tl: { col: 1, row: 1 }, ext: { width: w, height: h }, editAs: 'oneCell' });
          bottom = 1 + Math.ceil(h / 20);
        } else {
          irw.getCell('B2').value = 'Upload the BOT interest reserve statement for ' + Q.label + ' on the Compliance Certificates page, then run Report again.';
          irw.getCell('B2').font = { bold: true, color: { argb: 'FFC00000' } };
        }
        irBalRow = Math.max(33, bottom + 4);
        if (irBalRow !== 33) {   // keep the balance underneath the picture
          copyRowStyle(irw, 33, irBalRow); irw.getCell(irBalRow, 2).value = irw.getCell('B33').value; irw.getCell('B33').value = null; irw.getCell('D33').value = null;
        }
        const d = irw.getCell(irBalRow, 4); d.value = irTot; if (!d.numFmt || d.numFmt === 'General') d.numFmt = NF;
        if (!irUp) d.fill = YEL;
        if (irUp) { irw.getCell(irBalRow + 1, 2).value = 'Statement: ' + (irUp.filename || '') + (irUp.mime && /pdf/i.test(irUp.mime) ? ' (first page)' : ''); irw.getCell(irBalRow + 1, 2).font = { italic: true, size: 9, color: { argb: 'FF666666' } }; }
      }
    }
    // NonRecurring Expenses: add this quarter's block (same look as the earlier blocks); Calc picks up its Total
    let nrRef = null;
    {
      const nw = tplSheets['NonRecurring Expenses'];
      if (nw) {
        let last = 0, lab = null, hdr = null, tot = null;
        nw.eachRow({ includeEmpty: false }, (row, r) => {
          last = r; const bv = row.getCell(2).value;
          if (typeof bv === 'string' && /^\d+Q\d{4}$/.test(bv.trim())) { lab = r; hdr = r + 1; tot = null; }
          if (lab && !tot && String(row.getCell(4).value || '').trim() === 'Total') tot = r;
        });
        const b0 = last + 3; const tag = Q.q + 'Q' + Q.y;
        if (lab) copyRowStyle(nw, lab, b0, 8); nw.getCell(b0, 2).value = tag;
        if (hdr) { copyRowStyle(nw, hdr, b0 + 1, 8); for (let c = 2; c <= 7; c++) nw.getCell(b0 + 1, c).value = nw.getCell(hdr, c).value; }
        for (let k = 2; k <= 5; k++) if (hdr) copyRowStyle(nw, hdr + 1, b0 + k, 8);
        if (tot) copyRowStyle(nw, tot, b0 + 6, 8);
        nw.getCell(b0 + 6, 4).value = 'Total'; nw.getCell(b0 + 6, 5).value = { formula: 'SUM(E' + (b0 + 2) + ':E' + (b0 + 5) + ')', result: 0 };
        nrRef = "'NonRecurring Expenses'!E" + (b0 + 6);
      }
    }

    // Calc sheet (same layout as the hand-built workbook)
    const cols = { CLIP: 'G', Silsbee: 'H', SRN: 'I', Buna: 'J' }; const iaCol = { CLIP: 'B', Silsbee: 'C', SRN: 'D', Buna: 'E' };
    calc.getCell('B2').value = Q.label; calc.getCell('F2').value = 'Interest Reserve';
    Object.entries(cols).forEach(([k, c]) => { calc.getCell(c + '2').value = k; });
    calc.getCell('L2').value = 'Annualized T' + N;
    ['Total Revenue', 'Less: Operating Expenses', 'Less: Cash Distributions', 'Add back: Onetime-NonRecurring Expenses', 'Add back: Reserve Account Balance', 'Total Net Operating Income'].forEach((t, i) => { calc.getCell('B' + (3 + i)).value = t; });
    let sRev = 0, sOpx = 0, sInt = 0;
    for (const [k, c] of Object.entries(cols)) {
      const r = res[k];
      if (r) { setF(calc, c + '3', r.fRev, r.rev, NF); setF(calc, c + '4', r.fOpx, -r.opx, NF); sRev += r.rev; sOpx += r.opx; }
      setF(calc, c + '8', `SUM(${c}3:${c}7)`, r ? r.rev - r.opx : 0, NF);
      setF(calc, c + '10', `SUM(${qs('Interest Allocation')}!${iaCol[k]}4:${iaCol[k]}${iaLast})`, intTot[k], NF); sInt += intTot[k];
      setF(calc, c + '12', `+${c}10+${c}11`, intTot[k], NF);
    }
    const nr = +mc.nonrecurring_addback || 0;
    calc.getCell('I6').value = nr; calc.getCell('I6').numFmt = NF;
    setF(calc, 'F7', '=' + irRef, irTot, NF); setF(calc, 'F8', 'SUM(F3:F7)', irTot, NF);
    const L3 = sRev * ann, L4 = -sOpx * ann, L8 = L3 + L4 + nr + irTot, L10 = sInt * ann, dscr = L10 ? L8 / L10 : null;
    setF(calc, 'L3', `SUM(F3:J3)*${ann}`, L3, NF); setF(calc, 'L4', `SUM(F4:J4)*${ann}`, L4, NF);
    setF(calc, 'L5', 'SUM(F5:J5)', 0, NF); setF(calc, 'L6', 'SUM(F6:J6)', nr, NF); setF(calc, 'L7', 'SUM(F7:J7)', irTot, NF);
    setF(calc, 'L8', 'L3+L4+L6+L7', L8, NF);
    calc.getCell('B10').value = 'Cash Interest Expense'; calc.getCell('B11').value = 'Principal Payments Paid'; calc.getCell('B12').value = 'Total Debt Service';
    calc.getCell('B14').value = 'DSCR'; calc.getCell('B15').value = 'Covenant';
    setF(calc, 'L10', `SUM(F10:J10)*${ann}`, L10, NF); setF(calc, 'L11', 'SUM(F11:J11)', 0, NF); setF(calc, 'L12', 'L10+L11', L10, NF);
    setF(calc, 'L14', 'L8/L12', dscr, '0.00"x"');
    const dscrMin = Q.iso <= '2026-12-31' ? mc.dscr_min_through_2026 : mc.dscr_min_after_2026; calc.getCell('L15').value = dscrMin;
    // LTV & Debt yield
    calc.getCell('N2').value = 'Maximum Loan-to-Value Ratio'; calc.getCell('N3').value = Q.label + ' Loan Balance'; setF(calc, 'O3', qs('Loan Balance') + '!B3', loan || 0, NF);
    calc.getCell('N4').value = 'Total Rail Asset Appraisal'; setF(calc, 'O4', qs('Appraisals') + '!B7', apTot, '#,##0');
    const ltv = apTot ? (loan || 0) / apTot : null; setF(calc, 'O5', 'O3/O4', ltv, '0.00%'); calc.getCell('N7').value = 'Covenant'; calc.getCell('O7').value = mc.ltv_max;
    calc.getCell('Q2').value = 'Minimum Debt Yield'; calc.getCell('Q3').value = 'Annualized T' + N + ' NOI'; setF(calc, 'R3', 'L8', L8, NF);
    calc.getCell('Q4').value = Q.label + ' Loan Balance'; setF(calc, 'R4', 'O3', loan || 0, NF);
    const dy = loan ? L8 / loan : null; setF(calc, 'R5', 'R3/R4', dy, '0.00%'); calc.getCell('Q7').value = 'Covenant'; calc.getCell('R7').value = mc.debt_yield_min;
    calc.getCell('B17').value = 'Basis: Credit Agreement §10.1 - trailing ' + N + '-month period, annualized (x' + (Math.round(ann * 1000) / 1000) + '), window ' + mLabel(win[0]) + ' - ' + mLabel(QE) + '.';
    calc.getCell('B18').value = 'NOI excludes depreciation and amortization. Yellow cells = pending / not found - fill in and the calc updates.';
    calc.getColumn(2).width = 40; ['F', 'G', 'H', 'I', 'J', 'L', 'O', 'R'].forEach((c) => { calc.getColumn(c).width = 16; }); calc.getColumn('N').width = 28; calc.getColumn('Q').width = 24;
    if (tplCalc) {
      const T = tplCalc;
      const own = new Set(['G3', 'H3', 'I3', 'J3', 'G4', 'H4', 'I4', 'J4', 'G10', 'H10', 'I10', 'J10', 'L3', 'L4', 'L10', 'F7', 'I6', 'O3']);
      // drop cached results of the template's own formulas so Excel recalculates everything on open
      T.eachRow((row) => row.eachCell((c) => { const v = c.value; if (v && typeof v === 'object' && v.formula && !own.has(c.address)) c.value = { formula: v.formula }; }));
      ['G3', 'H3', 'I3', 'J3', 'G4', 'H4', 'I4', 'J4', 'G10', 'H10', 'I10', 'J10', 'L3', 'L4', 'L10'].forEach((a) => { T.getCell(a).value = calc.getCell(a).value; });
      T.getCell('B2').value = Q.label; T.getCell('L2').value = 'Annualized T' + N;
      T.getCell('F7').value = { formula: "'Interest Reserve'!D" + irBalRow, result: irTot };
      T.getCell('I6').value = nrRef ? { formula: nrRef, result: 0 } : nr;
      T.getCell('O3').value = { formula: "'Loan Balance'!B3", result: loan || 0 };
      T.getCell('N3').value = Q.label + ' Loan Balance'; T.getCell('Q4').value = Q.label + ' Loan Balance'; T.getCell('Q3').value = 'Annualized T' + N + ' NOI';
      T.getCell('L15').value = dscrMin;
    }
    out.metrics.midco = {
      basis: 'T' + N + ' annualized', window: mLabel(win[0]) + ' - ' + mLabel(QE),
      noi_annualized: L8, interest_annualized: L10, dscr, dscr_min: dscrMin, dscr_pass: dscr != null ? dscr >= dscrMin : null,
      loan_balance: loan, appraisal_total: apTot, ltv, ltv_max: mc.ltv_max, ltv_pass: ltv != null ? ltv <= mc.ltv_max : null,
      debt_yield: dy, debt_yield_min: mc.debt_yield_min, dy_pass: dy != null ? dy >= mc.debt_yield_min : null,
      interest_reserve: irTot, pending,
    };
    pending.forEach((p) => out.warnings.push('Midco: ' + p));
    await put('Midco I', 'midco', 'Midco Financial Covenant_' + Q.tag + '.xlsx', await wbBuffer(wb),
      'DSCR ' + (dscr ? dscr.toFixed(2) : 'n/a') + 'x' + (pending.length ? ' - DRAFT, ' + pending.length + ' pending input(s)' : ''));
  });

  // ── 6. summary workbook ───────────────────────────────────────────────────
  await safe('Summary', 'report', async () => {
    const wb = newWb(); const ws = wb.addWorksheet('Report');
    ws.getCell('A1').value = 'Compliance Certificates - ' + Q.label + ' (quarter end ' + Q.iso + ')'; ws.getCell('A1').font = { bold: true, size: 13 };
    ws.getCell('A2').value = 'Generated ' + new Date().toISOString().replace('T', ' ').slice(0, 16) + ' UTC by CloudLedger';
    ws.getRow(4).values = ['Package', 'File', 'Status', 'Note']; ws.getRow(4).font = { bold: true };
    out.items.forEach((it, i) => { ws.getRow(5 + i).values = [it.pkg, it.file, it.status, it.note]; if (it.status !== 'created' && it.status !== 'copied') ws.getRow(5 + i).fill = YEL; });
    let r = 6 + out.items.length;
    if (out.warnings.length) { ws.getCell('A' + r).value = 'Warnings'; ws.getCell('A' + r).font = { bold: true }; out.warnings.forEach((w) => { r++; ws.getCell('A' + r).value = w; }); }
    [26, 70, 10, 80].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    const url = await store.write('accounting', out.dest + '/Compliance Report_' + Q.tag + '.xlsx', await wbBuffer(wb));
    out.report_url = url;
  });
  out.counts = out.items.reduce((a, it) => { a[it.status] = (a[it.status] || 0) + 1; return a; }, {});
  return out;
}

// ════════════════════════════════════════════════════════════════════════════
// Routes
// ════════════════════════════════════════════════════════════════════════════
function ensureSchema(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS compliance_cert_config (id INTEGER PRIMARY KEY CHECK (id = 1), data TEXT NOT NULL, updated_by TEXT, updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE IF NOT EXISTS compliance_cert_runs (id INTEGER PRIMARY KEY AUTOINCREMENT, quarter_end TEXT NOT NULL, status TEXT NOT NULL, log TEXT, result TEXT, error TEXT,
      by_who TEXT, started_at TEXT DEFAULT (datetime('now')), finished_at TEXT);`);
  try { db.exec('ALTER TABLE compliance_cert_runs ADD COLUMN claimed_at TEXT'); } catch (e) { /* already there */ }
  db.exec(`CREATE TABLE IF NOT EXISTS compliance_cert_uploads (id INTEGER PRIMARY KEY AUTOINCREMENT, quarter_end TEXT NOT NULL, kind TEXT NOT NULL,
    filename TEXT, mime TEXT, data BLOB, balance REAL, by_who TEXT, created_at TEXT DEFAULT (datetime('now')), UNIQUE (quarter_end, kind))`);
  if (!db.prepare('SELECT id FROM compliance_cert_config WHERE id = 1').get()) {
    const d = JSON.parse(JSON.stringify(Object.assign({}, DEFAULTS, { midco: Object.assign({}, DEFAULTS.midco, { entities: undefined }) })));
    db.prepare('INSERT INTO compliance_cert_config (id, data, updated_by) VALUES (1, ?, ?)').run(JSON.stringify(d), 'seed');
  }
}
function getConfig(db) {
  const row = db.prepare('SELECT data FROM compliance_cert_config WHERE id = 1').get();
  const saved = row ? JSON.parse(row.data) : {};
  return Object.assign({}, DEFAULTS, saved, { midco: Object.assign({}, DEFAULTS.midco, saved.midco || {}, { entities: DEFAULTS.midco.entities }) });
}
function lastQuarterEnd(today = new Date()) {
  const y = today.getUTCFullYear(), m = today.getUTCMonth() + 1; const qm = Math.floor((m - 1) / 3) * 3;   // last completed quarter's end month
  const dt = new Date(Date.UTC(y, qm, 0)); return dt.toISOString().slice(0, 10);
}

function registerComplianceCertRoutes(app, ctx) {
  const { db, auth, requireRole } = ctx;
  ensureSchema(db);
  const running = new Map();
  const who = (req) => (req.user && (req.user.email || req.user.username || req.user.name)) || 'unknown';
  // How reports execute:
  //   sharepoint - in this server, via Microsoft Graph (MS_TENANT_ID/CLIENT_ID/CLIENT_SECRET set)
  //   local      - in this server, against OneDrive-synced folders (CC_LOCAL_ROOT; dev machine)
  //   runner     - queued for the runner program on Jimmy's PC (CC_RUNNER_KEY set), which
  //                builds the package from the PC's OneDrive folders (scripts/compliance-runner.js)
  // accept the variable name in any capitalization (e.g. CC_Runner_key) and ignore stray whitespace
  const runnerKey = () => { const n = Object.keys(process.env).find((k) => k.toUpperCase() === 'CC_RUNNER_KEY'); return n ? String(process.env[n]).trim() : ''; };
  const mode = () => (graphConfigured() ? 'sharepoint' : (process.env.CC_STORAGE === 'local' || process.env.CC_LOCAL_ROOT) ? 'local' : runnerKey() ? 'runner' : 'none');
  let runnerSeen = 0, runnerHost = '';
  const crypto = require('crypto');
  const runnerAuth = (req, res, next) => {
    const k = runnerKey(), got = String(req.get('x-runner-key') || '');
    if (!k || got.length !== k.length || !crypto.timingSafeEqual(Buffer.from(got), Buffer.from(k))) return res.status(401).json({ error: 'invalid runner key' });
    runnerSeen = Date.now(); runnerHost = String(req.get('x-runner-host') || '').slice(0, 80); next();
  };
  const publicConfig = () => { const c = JSON.parse(JSON.stringify(getConfig(db))); if (c.midco) delete c.midco.entities; return c; };

  app.get('/api/admin/compliance-certs', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const cfg = getConfig(db); const c = JSON.parse(JSON.stringify(cfg)); if (c.midco) delete c.midco.entities;
    const runs = db.prepare('SELECT id, quarter_end, status, by_who, started_at, finished_at, error FROM compliance_cert_runs ORDER BY id DESC LIMIT 15').all();
    res.json({ config: c, graph_configured: graphConfigured(), storage: mode(),
      runner_configured: !!runnerKey(), runner_online: Date.now() - runnerSeen < 45000, runner_host: runnerHost,
      runner_last_seen: runnerSeen ? new Date(runnerSeen).toISOString() : null,
      default_quarter_end: lastQuarterEnd(), dest_base: DEST_BASE, runs });
  });
  app.put('/api/admin/compliance-certs/config', auth, requireRole('Admin'), (req, res) => {
    const cur = getConfig(db); const b = req.body || {};
    const next = Object.assign({}, cur, b, { midco: Object.assign({}, cur.midco, b.midco || {}) }); delete next.midco.entities;
    db.prepare("UPDATE compliance_cert_config SET data = ?, updated_by = ?, updated_at = datetime('now') WHERE id = 1").run(JSON.stringify(next), who(req));
    res.json({ ok: true });
  });
  app.post('/api/admin/compliance-certs/report', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    let Q; try { Q = parseQuarterEnd((req.body || {}).quarter_end); } catch (e) { return res.status(400).json({ error: e.message }); }
    if ([...running.values()].some((r) => r.quarter_end === Q.iso) ||
        db.prepare("SELECT id FROM compliance_cert_runs WHERE quarter_end = ? AND status IN ('queued','running')").get(Q.iso))
      return res.status(409).json({ error: 'A report for ' + Q.label + ' is already queued or running' });
    if (mode() === 'none') return res.status(400).json({ error: 'Report runner is not configured (set CC_RUNNER_KEY, or Microsoft Graph credentials)' });
    if (mode() === 'runner') {
      const qid = db.prepare("INSERT INTO compliance_cert_runs (quarter_end, status, log, by_who) VALUES (?, 'queued', '', ?)").run(Q.iso, who(req)).lastInsertRowid;
      return res.json({ id: qid, quarter: Q.label, queued: true });
    }
    const id = db.prepare("INSERT INTO compliance_cert_runs (quarter_end, status, log, by_who) VALUES (?, 'running', '', ?)").run(Q.iso, who(req)).lastInsertRowid;
    const state = { quarter_end: Q.iso, log: [] }; running.set(id, state);
    const upd = db.prepare('UPDATE compliance_cert_runs SET log = ? WHERE id = ?');
    runReport(Q.iso, getConfig(db), (line) => { state.log.push(line); upd.run(state.log.join('\n'), id); })
      .then((result) => db.prepare("UPDATE compliance_cert_runs SET status = 'done', result = ?, finished_at = datetime('now') WHERE id = ?").run(JSON.stringify(result), id))
      .catch((e) => db.prepare("UPDATE compliance_cert_runs SET status = 'failed', error = ?, finished_at = datetime('now') WHERE id = ?").run(String(e && e.message || e), id))
      .finally(() => running.delete(id));
    res.json({ id, quarter: Q.label });
  });
  // ── endpoints used by the runner on Jimmy's PC (authenticated by CC_RUNNER_KEY) ──
  app.post('/api/compliance-runner/heartbeat', runnerAuth, (req, res) => res.json({ ok: true }));
  // entity list (names/types only) so entity matching can be checked from the PC
  app.post('/api/compliance-runner/entities', runnerAuth, (req, res) => { const cols = db.prepare('PRAGMA table_info(entities)').all().map((c) => c.name); const pick = ['id', 'code', 'name', 'entity_type', 'parent_id', 'consolidation_group_id', 'status', 'is_active'].filter((c) => cols.includes(c)); res.json({ columns: cols, entities: db.prepare('SELECT ' + pick.join(', ') + ' FROM entities ORDER BY id').all() }); });
  app.post('/api/compliance-runner/claim', runnerAuth, (req, res) => {
    // a job claimed more than 30 minutes ago without finishing is treated as abandoned
    db.prepare("UPDATE compliance_cert_runs SET status = 'failed', error = 'Runner stopped before finishing (timed out)', finished_at = datetime('now') WHERE status = 'running' AND claimed_at IS NOT NULL AND claimed_at < datetime('now', '-30 minutes')").run();
    const job = db.transaction(() => {
      const r = db.prepare("SELECT id, quarter_end FROM compliance_cert_runs WHERE status = 'queued' ORDER BY id LIMIT 1").get();
      if (!r) return null;
      db.prepare("UPDATE compliance_cert_runs SET status = 'running', claimed_at = datetime('now'), log = ? WHERE id = ?").run('Picked up by ' + (runnerHost || 'runner'), r.id);
      return r;
    })();
    let cl = null;
    if (job && ctx.signToken) {
      const by = db.prepare('SELECT by_who FROM compliance_cert_runs WHERE id = ?').get(job.id);
      const u = by && db.prepare('SELECT id, name, email, role FROM users WHERE email = ?').get(by.by_who);
      if (u && (u.role === 'Admin' || u.role === 'Accountant')) {
        cl = { token: ctx.signToken(u), entities: db.prepare('SELECT id, code, name, entity_type FROM entities').all() };
      }
    }
    const uploads = job ? db.prepare('SELECT kind, filename, mime, balance, data FROM compliance_cert_uploads WHERE quarter_end = ?').all(job.quarter_end)
      .map((u) => ({ kind: u.kind, filename: u.filename, mime: u.mime, balance: u.balance, data: u.data ? Buffer.from(u.data).toString('base64') : null })) : [];
    res.json({ job: job ? { id: job.id, quarter_end: job.quarter_end, config: publicConfig(), cl, uploads } : null });
  });
  app.post('/api/compliance-runner/runs/:id/log', runnerAuth, (req, res) => {
    const r = db.prepare('SELECT log FROM compliance_cert_runs WHERE id = ?').get(+req.params.id); if (!r) return res.status(404).json({ error: 'not found' });
    const lines = ((req.body || {}).lines || []).map(String).slice(0, 200);
    db.prepare('UPDATE compliance_cert_runs SET log = ? WHERE id = ?').run([r.log || ''].concat(lines).filter(Boolean).join('\n'), +req.params.id);
    res.json({ ok: true });
  });
  app.post('/api/compliance-runner/runs/:id/finish', runnerAuth, (req, res) => {
    const b = req.body || {};
    if (b.ok) db.prepare("UPDATE compliance_cert_runs SET status = 'done', result = ?, finished_at = datetime('now') WHERE id = ?").run(JSON.stringify(b.result || {}), +req.params.id);
    else db.prepare("UPDATE compliance_cert_runs SET status = 'failed', error = ?, finished_at = datetime('now') WHERE id = ?").run(String(b.error || 'unknown error').slice(0, 2000), +req.params.id);
    res.json({ ok: true });
  });

  // statements uploaded on the page (e.g. the BOT interest reserve statement for the quarter)
  const multer = require('multer');
  const upMem = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
  app.get('/api/admin/compliance-certs/uploads', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const q = String(req.query.quarter_end || '');
    res.json(db.prepare('SELECT id, quarter_end, kind, filename, mime, balance, by_who, created_at, length(data) AS size FROM compliance_cert_uploads WHERE quarter_end = ? ORDER BY kind').all(q));
  });
  app.post('/api/admin/compliance-certs/uploads', auth, requireRole('Admin', 'Accountant'), upMem.single('file'), (req, res) => {
    let Q; try { Q = parseQuarterEnd((req.body || {}).quarter_end); } catch (e) { return res.status(400).json({ error: e.message }); }
    const kind = String((req.body || {}).kind || 'interest_reserve'); if (kind !== 'interest_reserve') return res.status(400).json({ error: 'unknown upload kind' });
    const bal = +String((req.body || {}).balance || '').replace(/[$,\s]/g, '');
    if (!req.file) return res.status(400).json({ error: 'Choose the statement file (PDF, PNG or JPG)' });
    if (!/pdf|png|jpe?g/i.test(req.file.mimetype + ' ' + req.file.originalname)) return res.status(400).json({ error: 'The statement must be a PDF, PNG or JPG' });
    if (!isFinite(bal) || !String((req.body || {}).balance || '').trim()) return res.status(400).json({ error: 'Enter the ending balance shown on the statement' });
    db.prepare(`INSERT INTO compliance_cert_uploads (quarter_end, kind, filename, mime, data, balance, by_who) VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT (quarter_end, kind) DO UPDATE SET filename = excluded.filename, mime = excluded.mime, data = excluded.data, balance = excluded.balance, by_who = excluded.by_who, created_at = datetime('now')`)
      .run(Q.iso, kind, req.file.originalname, req.file.mimetype, req.file.buffer, bal, who(req));
    res.json({ ok: true });
  });
  app.delete('/api/admin/compliance-certs/uploads/:id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    db.prepare('DELETE FROM compliance_cert_uploads WHERE id = ?').run(+req.params.id); res.json({ ok: true });
  });

  app.get('/api/admin/compliance-certs/runs/:id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const r = db.prepare('SELECT * FROM compliance_cert_runs WHERE id = ?').get(+req.params.id);
    if (!r) return res.status(404).json({ error: 'not found' });
    res.json(Object.assign({}, r, { result: r.result ? JSON.parse(r.result) : null, log: r.log ? r.log.split('\n') : [] }));
  });
}

module.exports = { registerComplianceCertRoutes, runReport, parseQuarterEnd, stampCertificate, DEFAULTS };
