// ═══════════════════════════════════════════════════════════════════════════
// Dissolution — Delaware Certificate of Cancellation generator (Administration)
//
// Given an entity (LLC) name, produces a filled Delaware Certificate of
// Cancellation (DE085) ready for execution:
//   • the LLC name on line 1,
//   • the Certificate of Formation filing date on line 2,
//   • optional execution date (day / month / year) and signer name.
//
// The formation date is resolved by findFormationDate():
//   1. CloudLedger's own entity_formation store (a date confirmed/entered once,
//      reused thereafter), then
//   2. a live SharePoint search (Microsoft Graph, app-only) for the entity's
//      Certificate of Formation document — the filing date is read out of the
//      filed certificate's cover ("FILED IN THIS OFFICE ON THE <ORDINAL> DAY OF
//      <MONTH>, A.D. <YEAR>") or, for image-only scans, from the filename.
//
// The live SharePoint search only runs when Microsoft Graph is configured on
// the server (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET). When it is not,
// the store + manual entry keep the feature fully usable, and any date the
// search finds is written back to the store so it is instant next time.
//
// The blank template is NOT shipped in the repo. It lives on the persistent
// data volume under <dataDir>/templates/de_certificate_of_cancellation.pdf and
// is uploaded once through the Dissolution page (Admin only), so the form can
// be revised without a code deploy. The template is a Wolters Kluwer fillable
// AcroForm; we fill its named text fields with pdf-lib.
//
//   Field map (verified against DE085, US-Letter 612×792):
//     Text9 → LLC name (line 1)        Text4 → "filed on" date (line 2)
//     Text5 → execution day            Text3 → execution month
//     Text6 → execution year (A.D.)    Text8 → signer name ("Print or Type")
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

const TEMPLATE_FILE = 'de_certificate_of_cancellation.pdf';
const SIGNATURE_FILE = 'dissolution_signature.png'; // authorized-person signature, stamped on the By line

// ── SharePoint drives to search for formation documents ──────────────────────
// Default to the two libraries where Banyan's filed org docs actually live
// (BanyanResSharedDrive and the CLA site). Override with DISSOLUTION_SP_DRIVES
// (comma-separated Graph driveIds) or DISSOLUTION_SP_SITES (comma-separated site
// names, resolved to their default document library at runtime).
const DEFAULT_DRIVES = [
  'b!cbaGfpSSX0evsGx2H5r0DLLjp034YxNLqDnReYAjik854JV4zX_5QJceAZw9RQva', // BanyanResSharedDrive / Shared Documents
  'b!iI-WnoafAkuB4pUHxideeuZiSs2QftJDtLQF82oB1LI0jFi_EQNxQK7vnO48B5XV', // CLA / Shared Documents
];
const DEFAULT_SITES = ['BanyanResSharedDrive', 'CLA'];

// ── Microsoft Graph (app-only) ───────────────────────────────────────────────
let tokenCache = null;
const graphConfigured = () => !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
async function graphToken() {
  const t = process.env.MS_TENANT_ID, c = process.env.MS_CLIENT_ID, s = process.env.MS_CLIENT_SECRET;
  if (!t || !c || !s) return null;
  if (tokenCache && tokenCache.exp > Date.now() + 60000) return tokenCache.tok;
  const r = await fetch('https://login.microsoftonline.com/' + t + '/oauth2/v2.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: c, client_secret: s, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!r.ok) throw new Error('Graph token ' + r.status + ': ' + (await r.text()).slice(0, 200));
  const j = await r.json();
  tokenCache = { tok: j.access_token, exp: Date.now() + (j.expires_in || 3000) * 1000 };
  return tokenCache.tok;
}
async function gfetch(url, opts = {}) {
  const token = await graphToken();
  if (!token) throw new Error('Microsoft Graph is not configured (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET)');
  const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url,
    { method: opts.method || 'GET', headers: Object.assign({ Authorization: 'Bearer ' + token }, opts.headers || {}), body: opts.body });
  if (!r.ok && !(opts.ok || []).includes(r.status)) throw new Error('Graph ' + (opts.method || 'GET') + ' ' + r.status + ': ' + (await r.text()).slice(0, 250));
  return r;
}
// Resolve the configured drive ids (explicit ids first, then site-name lookups).
let driveIdsCache = null;
async function resolveDriveIds() {
  if (driveIdsCache) return driveIdsCache;
  const ids = [];
  const explicit = (process.env.DISSOLUTION_SP_DRIVES || '').split(',').map((x) => x.trim()).filter(Boolean);
  for (const id of explicit) if (!ids.includes(id)) ids.push(id);
  const sites = (process.env.DISSOLUTION_SP_SITES || '').split(',').map((x) => x.trim()).filter(Boolean);
  for (const name of sites) {
    try {
      const sr = await (await gfetch('/sites?search=' + encodeURIComponent(name))).json();
      const site = (sr.value || []).find((s) => (s.displayName || '').toLowerCase() === name.toLowerCase()) || (sr.value || [])[0];
      if (site) {
        const d = (await (await gfetch('/sites/' + site.id + '/drive')).json()).id;
        if (d && !ids.includes(d)) ids.push(d);
      }
    } catch { /* non-fatal; fall through to defaults */ }
  }
  if (!ids.length) {
    for (const id of DEFAULT_DRIVES) ids.push(id);
    // also attempt default sites if the hardcoded drive ids ever change
    for (const name of DEFAULT_SITES) {
      try {
        const sr = await (await gfetch('/sites?search=' + encodeURIComponent(name))).json();
        const site = (sr.value || []).find((s) => (s.displayName || '').toLowerCase() === name.toLowerCase());
        if (site) { const d = (await (await gfetch('/sites/' + site.id + '/drive')).json()).id; if (d && !ids.includes(d)) ids.push(d); }
      } catch { /* ignore */ }
    }
  }
  driveIdsCache = ids;
  return ids;
}

// ── date helpers ─────────────────────────────────────────────────────────────
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const ORD = {
  first: 1, second: 2, third: 3, fourth: 4, fifth: 5, sixth: 6, seventh: 7, eighth: 8, ninth: 9, tenth: 10,
  eleventh: 11, twelfth: 12, thirteenth: 13, fourteenth: 14, fifteenth: 15, sixteenth: 16, seventeenth: 17,
  eighteenth: 18, nineteenth: 19, twentieth: 20, thirtieth: 30,
};
const TENS = { twenty: 20, thirty: 30 };
function ordinalWordToNum(w) {
  if (!w) return null;
  const s = w.toLowerCase().trim().replace(/\s+/g, '-');
  if (ORD[s] != null) return ORD[s];
  if (s.includes('-')) {
    const [a, b] = s.split('-');
    if (TENS[a] != null && ORD[b] != null) return TENS[a] + ORD[b];
  }
  return null;
}
function fmtDate(y, m, d) {
  // m is 1-based. Returns { iso, display } or null.
  if (!y || !m || !d || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const yy = y < 100 ? 2000 + y : y;
  const iso = yy + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  const display = MONTHS[m - 1].replace(/^./, (c) => c.toUpperCase()) + ' ' + d + ', ' + yy;
  return { iso, display, y: yy, m, d };
}
// Pull the filing date out of a filed Certificate of Formation's body text.
function extractFormationDateFromText(text) {
  if (!text) return null;
  const T = text.replace(/\s+/g, ' ');
  // Delaware filed-certificate cover: "...FILED IN THIS OFFICE ON THE
  // TWENTY-SIXTH DAY OF MAY, A.D. 2026, AT 4:41 O`CLOCK P.M." (the leading
  // "Date: 05-27-26" on the same page is the certification date, NOT this.)
  let m = T.match(/FILED\s+IN\s+THIS\s+OFFICE\s+ON\s+THE\s+([A-Za-z][A-Za-z-]+(?:\s+[A-Za-z]+)?)\s+DAY\s+OF\s+([A-Za-z]+)\s*,?\s*A\.?\s*D\.?\s*,?\s*(\d{4})/i);
  if (m) {
    const day = ordinalWordToNum(m[1]);
    const mon = MONTHS.indexOf(m[2].toLowerCase()) + 1;
    const r = fmtDate(parseInt(m[3], 10), mon, day);
    if (r) { r.phrase = m[0].trim(); return r; }
  }
  // Generic "filed on 05/26/2026" / "effective ... 5/26/26"
  m = T.match(/\bfiled\b[^.]{0,40}?\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i);
  if (m) { const r = fmtDate(parseInt(m[3], 10), parseInt(m[1], 10), parseInt(m[2], 10)); if (r) { r.phrase = m[0].trim(); return r; } }
  // Generic "filed ... May 26, 2026"
  m = T.match(/\bfiled\b[^.]{0,40}?\b([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/i);
  if (m) { const mon = MONTHS.indexOf(m[1].toLowerCase()) + 1; const r = fmtDate(parseInt(m[3], 10), mon, parseInt(m[2], 10)); if (r) { r.phrase = m[0].trim(); return r; } }
  return null;
}
// Fallback: a date embedded in the filename (e.g. "...Formation 11.11.2025.pdf").
function extractDateFromFilename(name) {
  if (!name) return null;
  const base = name.replace(/\.[a-z0-9]+$/i, '');
  let m = base.match(/\b(20\d{2})[.\-_](\d{1,2})[.\-_](\d{1,2})\b/); // YYYY-MM-DD
  if (m) { const r = fmtDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10)); if (r) return r; }
  m = base.match(/\b(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2,4})\b/); // MM.DD.YYYY (US)
  if (m) { const r = fmtDate(parseInt(m[3], 10), parseInt(m[1], 10), parseInt(m[2], 10)); if (r) return r; }
  m = base.match(/\b([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})\b/); // Month D, YYYY
  if (m) { const mon = MONTHS.indexOf(m[1].toLowerCase()) + 1; const r = fmtDate(parseInt(m[3], 10), mon, parseInt(m[2], 10)); if (r) return r; }
  return null;
}

// ── entity-name matching / candidate ranking ────────────────────────────────
const LEGAL_SUFFIX = /[,\s]+(?:llc|l\.l\.c\.|lp|l\.p\.|ltd|inc|incorporated|corp|corporation|company|co|gp|l\.?p\.?)\.?$/i;
function coreName(entity) {
  let s = String(entity || '').trim();
  let prev;
  do { prev = s; s = s.replace(LEGAL_SUFFIX, '').trim(); } while (s !== prev);
  return s;
}
function tokens(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter((t) => t && t.length >= 2 && !['the', 'llc', 'lp', 'inc'].includes(t));
}
const FORMATION_HINT = /certificate\s+of\s+formation|cert.*formation|articles\s+of\s+organization|\bformation\b/i;
const NEGATIVE_HINT = /cancellation|dissolution|dissolv|amend|restat|operating\s+agreement|llc\s+agreement|resolution|consent|good\s*standing|draft|template/i;
function scoreCandidate(entity, name) {
  const n = String(name || '').toLowerCase();
  const toks = tokens(entity);
  if (!toks.length) return -999;
  const hits = toks.filter((t) => n.includes(t)).length;
  const coverage = hits / toks.length;
  if (coverage < 0.6) return -999; // not this entity
  let sc = hits * 3 + coverage * 2;
  if (/certificate\s+of\s+formation/.test(n)) sc += 5;
  else if (FORMATION_HINT.test(n)) sc += 2;
  if (/\bfiled\b/.test(n)) sc += 2;
  if (NEGATIVE_HINT.test(n)) sc -= 6;
  if (/\.pdf$/.test(n)) sc += 1;
  return sc;
}

// Live SharePoint search for an entity's filed Certificate of Formation.
async function searchFormationFromSharePoint(entity) {
  const drives = await resolveDriveIds();
  const core = coreName(entity);
  const queries = [core + ' certificate of formation', core];
  const seen = new Set();
  let candidates = [];
  for (const drive of drives) {
    for (const q of queries) {
      let items;
      try {
        const r = await gfetch('/drives/' + drive + "/root/search(q='" + encodeURIComponent(q).replace(/'/g, "''") + "')?$top=25&$select=id,name,webUrl,file", { ok: [404] });
        if (r.status === 404) continue;
        items = ((await r.json()).value || []);
      } catch { continue; }
      for (const it of items) {
        if (!it || it.folder) continue;
        const key = drive + '|' + it.id;
        if (seen.has(key)) continue;
        seen.add(key);
        const sc = scoreCandidate(entity, it.name);
        if (sc <= 0) continue;
        candidates.push({ driveId: drive, itemId: it.id, name: it.name, webUrl: it.webUrl, score: sc });
      }
      if (candidates.length) break; // first query that matched this drive is enough
    }
  }
  candidates.sort((a, b) => b.score - a.score);
  const out = { candidates: candidates.slice(0, 8) };
  // Try to read a date from the best candidates in order.
  let pdfParse;
  try { pdfParse = require('pdf-parse'); } catch { pdfParse = null; }
  for (const c of candidates.slice(0, 4)) {
    let dated = null;
    if (/\.pdf$/i.test(c.name) && pdfParse) {
      try {
        const buf = Buffer.from(await (await gfetch('/drives/' + c.driveId + '/items/' + c.itemId + '/content')).arrayBuffer());
        const parsed = await pdfParse(buf).catch(() => null);
        if (parsed && parsed.text) dated = extractFormationDateFromText(parsed.text);
      } catch { /* ignore, try filename */ }
    }
    if (!dated) dated = extractDateFromFilename(c.name);
    if (dated) {
      out.formation_date = dated.display;
      out.formation_date_iso = dated.iso;
      out.source_file = c.name;
      out.source_url = c.webUrl;
      out.source = 'sharepoint';
      if (dated.phrase) out.phrase = dated.phrase;
      break;
    }
  }
  return out;
}

// ── store (entity_formation) ─────────────────────────────────────────────────
function ensureTable(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS entity_formation (
    entity_key   TEXT PRIMARY KEY,
    entity_name  TEXT NOT NULL,
    formation_date TEXT,            -- display, e.g. "May 26, 2026"
    formation_date_iso TEXT,        -- "2026-05-26"
    state        TEXT,
    source_file  TEXT,
    source_url   TEXT,
    source       TEXT,              -- 'sharepoint' | 'manual'
    updated_at   TEXT DEFAULT (datetime('now'))
  )`);
}
const keyOf = (name) => String(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
function getStored(db, name) {
  return db.prepare('SELECT * FROM entity_formation WHERE entity_key = ?').get(keyOf(name)) || null;
}
function upsertStored(db, rec) {
  db.prepare(`INSERT INTO entity_formation (entity_key, entity_name, formation_date, formation_date_iso, state, source_file, source_url, source, updated_at)
    VALUES (@entity_key, @entity_name, @formation_date, @formation_date_iso, @state, @source_file, @source_url, @source, datetime('now'))
    ON CONFLICT(entity_key) DO UPDATE SET
      entity_name=excluded.entity_name, formation_date=excluded.formation_date, formation_date_iso=excluded.formation_date_iso,
      state=excluded.state, source_file=excluded.source_file, source_url=excluded.source_url, source=excluded.source,
      updated_at=datetime('now')`).run({
    entity_key: keyOf(rec.entity_name), entity_name: rec.entity_name,
    formation_date: rec.formation_date || null, formation_date_iso: rec.formation_date_iso || null,
    state: rec.state || null, source_file: rec.source_file || null, source_url: rec.source_url || null,
    source: rec.source || 'manual',
  });
}

// Resolve a formation date for an entity: store first, then live SharePoint.
async function findFormationDate(db, entity, { refresh = false } = {}) {
  const res = { entity, graph_configured: graphConfigured() };
  if (!refresh) {
    const s = getStored(db, entity);
    if (s && s.formation_date) {
      return Object.assign(res, {
        found: true, source: s.source || 'store', formation_date: s.formation_date,
        formation_date_iso: s.formation_date_iso, state: s.state, source_file: s.source_file, source_url: s.source_url,
      });
    }
  }
  if (graphConfigured()) {
    try {
      const sp = await searchFormationFromSharePoint(entity);
      if (sp.formation_date) {
        upsertStored(db, { entity_name: entity, formation_date: sp.formation_date, formation_date_iso: sp.formation_date_iso, source_file: sp.source_file, source_url: sp.source_url, source: 'sharepoint' });
        return Object.assign(res, { found: true, source: 'sharepoint', candidates: sp.candidates, phrase: sp.phrase,
          formation_date: sp.formation_date, formation_date_iso: sp.formation_date_iso, source_file: sp.source_file, source_url: sp.source_url });
      }
      return Object.assign(res, { found: false, candidates: sp.candidates || [] });
    } catch (e) {
      return Object.assign(res, { found: false, error: e.message, candidates: [] });
    }
  }
  return Object.assign(res, { found: false, candidates: [] });
}

// ── built-in blank form ──────────────────────────────────────────────────────
// Recreates the Delaware Certificate of Cancellation (DE085) as a fillable
// AcroForm with the same field names fillCancellation() expects, so generation
// works with zero setup. If an Admin uploads the official Wolters Kluwer PDF it
// is used instead (e.g. to carry a pre-drawn signature). US-Letter, 612×792.
async function buildBlankCancellation() {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([612, 792]);
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const form = pdf.getForm();
  const black = rgb(0, 0, 0);
  const draw = (t, x, y, size, f) => page.drawText(t, { x, y, size, font: f || reg, color: black });
  const center = (t, y, size, f) => { const ff = f || reg; draw(t, (612 - ff.widthOfTextAtSize(t, size)) / 2, y, size, ff); };
  const line = (x1, y, x2) => page.drawLine({ start: { x: x1, y }, end: { x: x2, y }, thickness: 0.8, color: black });
  const field = (name, x, y, w, h, size, def) => {
    const fld = form.createTextField(name);
    fld.addToPage(page, { x, y, width: w, height: h, borderWidth: 0 });
    if (size) try { fld.setFontSize(size); } catch { /* ignore */ }
    if (def) fld.setText(def);
    return fld;
  };
  const L = 72, R = 540;
  center('STATE OF DELAWARE', 712, 15, bold);
  center('CERTIFICATE OF CANCELLATION', 692, 15, bold);
  // 1. name of the LLC
  draw('1.', L, 646, 11, bold);
  draw('The name of the limited liability company is', L + 22, 646, 11);
  line(L, 622, R); field('Text9', L + 2, 625, R - (L + 2), 16, 11);
  line(L, 600, R); field('Text2', L + 2, 603, R - (L + 2), 16, 11);
  // 2. formation filing date
  draw('2.', L, 560, 11, bold);
  draw('The Certificate of Formation of the limited liability company was filed on', L + 22, 560, 11);
  line(L, 536, R); field('Text4', L + 2, 539, R - (L + 2), 16, 11);
  // execution clause
  draw('IN WITNESS WHEREOF, the undersigned has executed this Certificate of', L, 486, 11, bold);
  let x = L; const y = 466;
  const seg = (t, f) => { draw(t, x, y, 11, f); x += (f || reg).widthOfTextAtSize(t, 11); };
  seg('Cancellation this '); line(x, y - 2, x + 42); field('Text5', x, y - 1, 42, 14, 11); x += 46;
  seg(' day of '); line(x, y - 2, x + 92); field('Text3', x, y - 1, 92, 14, 11); x += 96;
  seg(', A.D. '); line(x, y - 2, x + 46); field('Text6', x, y - 1, 46, 14, 11); x += 48;
  seg('.');
  // signature block
  draw('By:', 322, 420, 11);
  line(346, 418, 560); field('Text7', 348, 421, 212, 16, 11);
  draw('Authorized Person(s)', 388, 405, 9);
  draw('Name:', 322, 378, 11);
  line(360, 376, 560); field('Text8', 362, 379, 198, 16, 11, 'Jimmy Yun');
  draw('Print or Type', 396, 363, 9);
  draw('DE085 — Delaware Certificate of Cancellation', L, 40, 7);
  return Buffer.from(await pdf.save());
}

// ── template fill ────────────────────────────────────────────────────────────
async function fillCancellation(templateBuf, f) {
  const pdf = await PDFDocument.load(templateBuf);
  const form = pdf.getForm();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const setField = (name, val, size) => {
    if (val == null || val === '') return;
    try {
      const fld = form.getTextField(name);
      if (size) { try { fld.setFontSize(size); } catch { /* ignore */ } }
      fld.setText(String(val));
    } catch { /* field absent — ignore */ }
  };
  // LLC name: shrink to fit the ~184pt line (Text9) for long names.
  const nm = String(f.entity || '');
  const nmSize = Math.max(7, Math.min(11, Math.floor(360 / Math.max(nm.length, 1))));
  setField('Text9', nm, nmSize);
  setField('Text4', f.formationDate, 11);  // "filed on" date
  setField('Text5', f.execDay);            // this __ day
  setField('Text3', f.execMonth);          // of ____
  setField('Text6', f.execYear);           // A.D. ____
  if (f.signerName) setField('Text8', f.signerName);
  // Stamp the authorized-person signature image on the By line, if provided.
  if (f.signatureBytes && f.signatureBytes.length) {
    try {
      const b = f.signatureBytes;
      const img = (b[0] === 0x89 && b[1] === 0x50) ? await pdf.embedPng(b)
        : (b[0] === 0xFF && b[1] === 0xD8) ? await pdf.embedJpg(b) : null;
      if (img) {
        const page = pdf.getPages()[0];
        const w = 105, h = w * (img.height / img.width);
        page.drawImage(img, { x: 356, y: 418 - h * 0.32, width: w, height: h });
      }
    } catch { /* ignore signature errors */ }
  }
  try { form.updateFieldAppearances(font); } catch { /* ignore */ }
  if (f.flatten !== false) { try { form.flatten(); } catch { /* ignore */ } }
  return Buffer.from(await pdf.save());
}

// ── ordinal day for execution date (11 -> "11th") ────────────────────────────
function ordinalDay(d) {
  d = parseInt(d, 10); if (!d) return '';
  const s = ['th', 'st', 'nd', 'rd'], v = d % 100;
  return d + (s[(v - 20) % 10] || s[v] || s[0]);
}

// ── routes ───────────────────────────────────────────────────────────────────
function registerDissolutionRoutes(app, ctx) {
  const { db, auth, requireRole, memUpload, templatesDir } = ctx;
  ensureTable(db);
  try { fs.mkdirSync(templatesDir, { recursive: true }); } catch { /* ignore */ }
  const tplPath = () => path.join(templatesDir, TEMPLATE_FILE);
  const sigPath = () => path.join(templatesDir, SIGNATURE_FILE);
  const bundledSigPath = path.join(__dirname, 'assets', SIGNATURE_FILE);
  // Prefer an uploaded override on the data volume; else the repo-bundled default.
  const activeSigPath = () => (fs.existsSync(sigPath()) ? sigPath() : (fs.existsSync(bundledSigPath) ? bundledSigPath : null));
  const BOTH = ['Admin', 'Accountant'];

  // Feature status — drives the page. Generation always works via the built-in
  // form; template_installed indicates an uploaded override is present.
  app.get('/api/admin/dissolution/status', auth, requireRole(...BOTH), (req, res) => {
    res.json({ template_installed: fs.existsSync(tplPath()), signature_installed: !!activeSigPath(), built_in: true, graph_configured: graphConfigured() });
  });

  // Install / replace the authorized-person signature image (Admin). PNG/JPG;
  // a transparent PNG stamps cleanly on the By line.
  app.post('/api/admin/dissolution/signature', auth, requireRole('Admin'), memUpload.single('file'), (req, res) => {
    try {
      const b = req.file && req.file.buffer;
      if (!b || !b.length) return res.status(400).json({ error: 'No file uploaded.' });
      const isPng = b[0] === 0x89 && b[1] === 0x50, isJpg = b[0] === 0xFF && b[1] === 0xD8;
      if (!isPng && !isJpg) return res.status(400).json({ error: 'Signature must be a PNG or JPG image.' });
      fs.writeFileSync(sigPath(), b);
      res.json({ note: 'Signature installed.', signature_installed: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
  app.delete('/api/admin/dissolution/signature', auth, requireRole('Admin'), (req, res) => {
    try { if (fs.existsSync(sigPath())) fs.unlinkSync(sigPath()); res.json({ signature_installed: false }); }
    catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Install / replace the blank Certificate of Cancellation template (Admin).
  app.post('/api/admin/dissolution/template', auth, requireRole('Admin'), memUpload.single('file'), async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) return res.status(400).json({ error: 'No file uploaded.' });
      // Validate it's a fillable PDF with the expected name/date fields.
      let ok = false;
      try { const d = await PDFDocument.load(req.file.buffer); const names = d.getForm().getFields().map((x) => x.getName()); ok = names.includes('Text9') && names.includes('Text4'); } catch { ok = false; }
      if (!ok) return res.status(400).json({ error: 'That does not look like the DE Certificate of Cancellation fillable form (expected fields Text9 / Text4).' });
      fs.writeFileSync(tplPath(), req.file.buffer);
      res.json({ note: 'Template installed.', template_installed: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Look up an entity's formation date (store, then live SharePoint if configured).
  app.get('/api/admin/dissolution/lookup', auth, requireRole(...BOTH), async (req, res) => {
    try {
      const entity = (req.query.entity || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      const refresh = req.query.refresh === '1' || req.query.refresh === 'true';
      res.json(await findFormationDate(db, entity, { refresh }));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Confirm / override a formation date in the store (manual entry).
  app.post('/api/admin/dissolution/formation', auth, requireRole(...BOTH), (req, res) => {
    try {
      const b = req.body || {};
      const entity = (b.entity || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      let iso = (b.formation_date_iso || '').toString().trim() || null;
      let display = (b.formation_date || '').toString().trim();
      // If an ISO date was given but no display, derive a clean display; vice-versa leave as typed.
      if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso) && !display) {
        const [y, m, d] = iso.split('-').map((x) => parseInt(x, 10));
        const r = fmtDate(y, m, d); if (r) display = r.display;
      }
      upsertStored(db, { entity_name: entity, formation_date: display || null, formation_date_iso: iso, state: b.state || null, source_file: b.source_file || null, source_url: b.source_url || null, source: 'manual' });
      res.json(Object.assign({ saved: true }, getStored(db, entity)));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // List everything in the store (small management table).
  app.get('/api/admin/dissolution/formations', auth, requireRole(...BOTH), (req, res) => {
    try { res.json({ rows: db.prepare('SELECT * FROM entity_formation ORDER BY entity_name').all() }); }
    catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Generate the filled Certificate of Cancellation PDF (uploaded template if
  // present, otherwise the built-in DE085 form).
  app.post('/api/admin/dissolution/generate', auth, requireRole(...BOTH), async (req, res) => {
    try {
      const b = req.body || {};
      const entity = (b.entity || '').toString().trim();
      const formationDate = (b.formation_date || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      // formation_date is optional — if omitted, line 2 is left blank to fill by hand.
      // Optional execution date (YYYY-MM-DD) → day / month / year blanks.
      let execDay = '', execMonth = '', execYear = '';
      const ed = (b.execution_date || '').toString().trim();
      if (ed && /^\d{4}-\d{2}-\d{2}$/.test(ed)) {
        const [y, m, d] = ed.split('-').map((x) => parseInt(x, 10));
        execDay = ordinalDay(d); execMonth = (MONTHS[m - 1] || '').replace(/^./, (c) => c.toUpperCase()); execYear = String(y);
      }
      const signerName = (b.signer_name || '').toString().trim(); // blank → keep template's existing Name
      const template = fs.existsSync(tplPath()) ? fs.readFileSync(tplPath()) : await buildBlankCancellation();
      const sp = activeSigPath();
      const signatureBytes = (b.include_signature !== false && sp) ? fs.readFileSync(sp) : null;
      const buf = await fillCancellation(template, { entity, formationDate, execDay, execMonth, execYear, signerName, signatureBytes, flatten: b.flatten !== false });
      const safe = entity.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'Entity';
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="Certificate_of_Cancellation_' + safe + '.pdf"');
      res.send(buf);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
}

module.exports = {
  registerDissolutionRoutes,
  // exported for local testing:
  buildBlankCancellation, fillCancellation, findFormationDate, searchFormationFromSharePoint,
  extractFormationDateFromText, extractDateFromFilename, ordinalWordToNum, coreName, scoreCandidate,
};
