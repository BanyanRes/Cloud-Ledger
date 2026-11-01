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
  }
  return Buffer.from(await doc.save());
}

// ════════════════════════════════════════════════════════════════════════════
// The report job
// ════════════════════════════════════════════════════════════════════════════
async function runReport(quarterEnd, cfg, log) {
  const Q = parseQuarterEnd(quarterEnd);
  const store = makeStore(cfg);
  const out = { quarter_end: Q.iso, quarter: Q.label, storage: store.kind, dest: DEST_BASE + '/' + Q.y + '/Q' + Q.q, items: [], metrics: {}, warnings: [] };
  const say = (s) => { if (log) log(s); };
  const item = (pkg, file, status, note, url) => { out.items.push({ pkg, file, status, note: note || '', url: url || null }); say(pkg + ': ' + status + ' - ' + file + (note ? ' (' + note + ')' : '')); };
  const destPath = (dir, f) => out.dest + '/' + DIRS[dir] + '/' + f;
  async function put(pkg, dir, fname, buf, note) { const url = await store.write('accounting', destPath(dir, fname), buf); item(pkg, fname, 'created', note, url); }
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
      await put('Skye Canyon', 'skye', '12_Month_Statement_res68831_Accrual.xlsx', await wb.xlsx.writeBuffer(), 'from ' + path.posix.basename(xl));
    } else item('Skye Canyon', '12_Month_Statement_res68831_Accrual.xlsx', 'missing', 'no MergedExcel_*.xlsx in ' + dir);
    await copy('Skye Canyon', 'skye', 'shared', pdfF, 'Skye Canyon ' + MON[QE.m - 1] + ' ' + QE.y + ' Financial Reports.pdf');
    await put('Skye Canyon', 'skye', 'Certified Rent Roll_Skye Canyon_' + Q.tag + '.pdf', await stampCertificate('rentroll_skye', { asof: asOf, signdate: '' }), 'unsigned');
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
      const hdr = rows[4] || [];
      const cols = Q.months.map((p) => hdr.findIndex((h) => h instanceof Date ? (h.getUTCFullYear() === p.y && h.getUTCMonth() + 1 === p.m) : false));
      if (cols.some((c) => c < 0)) throw new Error('quarter months not all in Mark-Taylor 12-month P&L');
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
      await put('South Mountain / Orchards', 'smo', fname, await wb.xlsx.writeBuffer(), 'from ' + path.posix.basename(pl));
    });
  }
  await safe('South Mountain / Orchards', 'rent rolls', async () => {
    await put('South Mountain / Orchards', 'smo', 'Certified Rent Roll_Orchards ' + Q.tag + '.pdf', await stampCertificate('rentroll_orchards', { asof: asOf, signdate: '' }), 'unsigned');
    await put('South Mountain / Orchards', 'smo', 'Certified Rent Roll_South Mountain ' + Q.tag + '.pdf', await stampCertificate('rentroll_sm', { asof: asOf, signdate: '' }), 'unsigned');
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
    calc.getCell('B2').value = 'Silo Ridge'; calc.getCell('C2').value = Q.label;
    const T = qs(tabName);
    calc.getCell('B3').value = 'Total Revenue'; setF(calc, 'C3', `SUM(${T}!C${rowOf['Total Revenue']}:E${rowOf['Total Revenue']})`, rev, NF);
    calc.getCell('B4').value = 'Less: Operating Expenses'; setF(calc, 'C4', `-SUM(${T}!C${rowOf['Total Operating Expenses']}:E${rowOf['Total Operating Expenses']})`, -opex, NF);
    calc.getCell('B5').value = 'Add back: Reserve Account Balance'; calc.getCell('C5').value = 0;
    calc.getCell('B6').value = 'Total Net Operating Income'; setF(calc, 'C6', 'SUM(C3:C5)', noi, NF);
    calc.getCell('B8').value = 'Cash Interest Expense'; setF(calc, 'C8', `SUM(${T}!C${rowOf['Total Interest Expense']}:E${rowOf['Total Interest Expense']})`, int, NF);
    calc.getCell('B9').value = 'Principal Payments Paid'; calc.getCell('C9').value = 0;
    calc.getCell('B10').value = 'Total Debt Service'; setF(calc, 'C10', '+C8+C9', int, NF);
    calc.getCell('B12').value = 'DSCR'; setF(calc, 'C12', 'C6/C10', dscr, '0.00"x"');
    calc.getColumn(2).width = 34; calc.getColumn(3).width = 16;
    out.metrics.silo_jpm = { revenue: rev, opex, noi, interest: int, dscr };
    await put('Silo Ridge - JPM', 'jpm', 'Silo Ridge DSCR Calc_' + Q.tag + '.xlsx', await wb.xlsx.writeBuffer(), 'DSCR ' + (dscr ? dscr.toFixed(2) : 'n/a') + 'x');
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
      const dir = await findMonthDir(store, 'cla', e.fin, QE.y, QE.m);
      const f = await findFile(store, 'cla', dir, e.fs, /\.pdf$/i);
      if (!f) item('Midco I', e.key + ' quarterly financial statements', 'missing', 'not found in ' + (dir || e.fin + '/' + QE.y + '/' + pad2(QE.m) + '*'));
      else await copy('Midco I', 'midco', 'cla', f);
    });
  }
  await safe('Midco I', 'CLRFI Midco I consolidated FS', async () => {
    const f = await findFile(store, 'cla', '61 Midco I/00 Financials/' + QE.y, new RegExp('consolidated financial statements ' + pad2(QE.m) + '\\.' + QE.y, 'i'), /\.pdf$/i);
    if (!f) item('Midco I', 'CLRFI Midco I consolidated FS', 'missing', 'not found in 61 Midco I/00 Financials/' + QE.y); else await copy('Midco I', 'midco', 'cla', f);
  });
  await safe('Midco I', 'CLRF quarterly FS', async () => {
    const f = await findFile(store, 'cla', '30 County Line Rail Fund/Financial Statements/' + QE.y + ' Q' + Q.q + ' Financials', /FS SCF Wksht_Final/i, /\.pdf$/i);
    if (!f) item('Midco I', 'CLRF Q' + Q.q + ' FS SCF Wksht', 'missing', 'not found in 30 County Line Rail Fund/Financial Statements/' + QE.y + ' Q' + Q.q + ' Financials'); else await copy('Midco I', 'midco', 'cla', f);
  });
  await safe('Midco I', 'certificate', async () => {
    await put('Midco I', 'midco', 'CLRFI Blank Compliance Certificate_' + Q.label.replace(/^Q(\d) /, 'Q$1 ') + '.pdf',
      await stampCertificate('bot_clrfi', { qend: MON[Q.m - 1] + ' ' + ord(Q.d), yy: String(Q.y).slice(2) }), 'quarter-end stamped, unsigned');
  });

  // 5b. covenant workbook
  await safe('Midco I', 'covenant workbook', async () => {
    const N = Math.max(1, +mc.dscr_months || 6), ann = 12 / N;
    const win = []; for (let k = N - 1; k >= 0; k--) win.push(Q.back(k));
    const wb = newWb(); const calc = wb.addWorksheet('Calc');
    const pending = [];
    const res = {};
    // T12 per entity
    for (const e of ents) {
      let f = null, srcMonth = null;
      for (let k = 0; k < 3 && !f; k++) {   // QE month, else the latest month in the quarter (flag pending)
        const p = Q.back(k); const dir = await findMonthDir(store, 'cla', e.fin, p.y, p.m);
        f = await findFile(store, 'cla', dir, /trailing.?12|\bT12\b/i, /\.xlsx$/i); if (f) srcMonth = p;
      }
      if (!f) { pending.push(e.key + ': no T12 found'); res[e.key] = null; continue; }
      const rows = sheetRows(await store.read('cla', f));
      const tab = (e.key + ' T12 ' + mLabel(srcMonth)).slice(0, 31);
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
    const iaLast = 3 + win.length; ia.getColumn(1).width = 12; ia.getColumn(7).width = 50; [2, 3, 4, 5, 6].forEach((c) => { ia.getColumn(c).width = 15; });
    // loan balance at quarter end (BOT statements bracketing the quarter end)
    let loan = null, loanSrc = '';
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
    if (loan == null) pending.push('Loan balance at ' + Q.iso + ' not found in BOT loan statements - enter on "Loan Balance" tab');
    const lb = wb.addWorksheet('Loan Balance'); lb.getCell('A1').value = 'BOT loan principal balance at ' + Q.iso; const lbc = lb.getCell('B3'); lbc.value = loan || 0; lbc.numFmt = NF; if (loan == null) lbc.fill = YEL;
    lb.getCell('A3').value = 'Principal balance'; lb.getCell('A5').value = 'Source: ' + (loanSrc || 'NOT FOUND'); lb.getColumn(1).width = 22; lb.getColumn(2).width = 18;
    // interest reserve (acct 13100) per entity TB at quarter end
    const ir = wb.addWorksheet('Interest Reserve'); ir.getCell('A1').value = 'Interest Reserve (acct 13100) at ' + Q.iso; ir.getRow(3).values = ['Entity', 'Balance', 'Source']; ir.getRow(3).font = { bold: true };
    let irTot = 0, irRow = 4;
    for (const e of ents) {
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
    const ap = wb.addWorksheet('Appraisals'); ap.getCell('A1').value = mc.appraisal_label || 'Appraisals';
    const apKeys = ['CLIP', 'Silsbee', 'Buna', 'SRN']; apKeys.forEach((k, i) => { ap.getCell(3 + i, 1).value = k; const c = ap.getCell(3 + i, 2); c.value = (mc.appraisals || {})[k] || 0; c.numFmt = '#,##0'; });
    const apTot = apKeys.reduce((a, k) => a + ((mc.appraisals || {})[k] || 0), 0);
    ap.getCell(7, 1).value = 'Total'; setF(ap, 'B7', 'SUM(B3:B6)', apTot, '#,##0'); ap.getColumn(1).width = 14; ap.getColumn(2).width = 16;

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
    out.metrics.midco = {
      basis: 'T' + N + ' annualized', window: mLabel(win[0]) + ' - ' + mLabel(QE),
      noi_annualized: L8, interest_annualized: L10, dscr, dscr_min: dscrMin, dscr_pass: dscr != null ? dscr >= dscrMin : null,
      loan_balance: loan, appraisal_total: apTot, ltv, ltv_max: mc.ltv_max, ltv_pass: ltv != null ? ltv <= mc.ltv_max : null,
      debt_yield: dy, debt_yield_min: mc.debt_yield_min, dy_pass: dy != null ? dy >= mc.debt_yield_min : null,
      interest_reserve: irTot, pending,
    };
    pending.forEach((p) => out.warnings.push('Midco: ' + p));
    await put('Midco I', 'midco', 'Midco Financial Covenant_' + Q.tag + '.xlsx', await wb.xlsx.writeBuffer(),
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
    const url = await store.write('accounting', out.dest + '/Compliance Report_' + Q.tag + '.xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
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
  app.post('/api/compliance-runner/claim', runnerAuth, (req, res) => {
    // a job claimed more than 30 minutes ago without finishing is treated as abandoned
    db.prepare("UPDATE compliance_cert_runs SET status = 'failed', error = 'Runner stopped before finishing (timed out)', finished_at = datetime('now') WHERE status = 'running' AND claimed_at IS NOT NULL AND claimed_at < datetime('now', '-30 minutes')").run();
    const job = db.transaction(() => {
      const r = db.prepare("SELECT id, quarter_end FROM compliance_cert_runs WHERE status = 'queued' ORDER BY id LIMIT 1").get();
      if (!r) return null;
      db.prepare("UPDATE compliance_cert_runs SET status = 'running', claimed_at = datetime('now'), log = ? WHERE id = ?").run('Picked up by ' + (runnerHost || 'runner'), r.id);
      return r;
    })();
    res.json({ job: job ? { id: job.id, quarter_end: job.quarter_end, config: publicConfig() } : null });
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

  app.get('/api/admin/compliance-certs/runs/:id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const r = db.prepare('SELECT * FROM compliance_cert_runs WHERE id = ?').get(+req.params.id);
    if (!r) return res.status(404).json({ error: 'not found' });
    res.json(Object.assign({}, r, { result: r.result ? JSON.parse(r.result) : null, log: r.log ? r.log.split('\n') : [] }));
  });
}

module.exports = { registerComplianceCertRoutes, runReport, parseQuarterEnd, stampCertificate, DEFAULTS };
