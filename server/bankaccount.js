// ═══════════════════════════════════════════════════════════════════════════
// Bank Account Application — MapleMark New Account Application filler (Admin)
//
// Given an entity name, produces a partially-filled MapleMark Bank "New Account
// Application (Business & Trust)" packet — the 4-page account application plus
// the 2-page Beneficial Ownership Certification — ready for manual completion
// and signature.
//
// Only ENTITY-IDENTITY fields are auto-filled:
//   • entity legal name (Entity Name — appears on p.1 and on the Beneficial
//     Ownership page as "name of legal entity"),
//   • entity type (derived from the legal-name suffix),
//   • tax identification number / EIN,
//   • physical + mailing address (street / city / state / zip),
//   • entity email + phone,
//   • nature of business / industry,
//   • the "NONE OF THE FOLLOWING" entity-attributes box (a real-estate fund
//     entity is none of the regulated categories).
//
// Everything that is PII or an account-specific choice is deliberately left
// BLANK for a human to complete: all signer rows (names, SSN/ITIN, phone,
// email, residency, ID, authorities), the beneficial-owner rows (names, DOB,
// SSN, % owned), the managing-individual row, every account-profile frequency/
// amount checkbox, and the account title / purpose / type / opening-deposit.
//
// The identity is resolved by resolveProfile():
//   1. CloudLedger's own entity_bank_profile store (confirmed/entered once,
//      reused thereafter), then
//   2. a live SharePoint search (Microsoft Graph, app-only) for the entity's
//      W-9, EIN confirmation letter (CP-575 / SS-4), or Certificate of
//      Formation — the EIN and address are read out of whichever is found.
//
// The live SharePoint search only runs when Microsoft Graph is configured
// (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET). When it is not, the store
// + manual entry keep the feature fully usable, and anything the search finds
// is written back to the store so it is instant next time.
//
// The blank packet is NOT shipped in the repo. It lives on the persistent data
// volume under <dataDir>/templates/maplemark_new_account_application.pdf and is
// uploaded once through the Bank Account Application page (Admin), so the form
// can be revised without a code deploy. An optional repo-bundled copy at
// server/assets/maplemark_new_account_application.pdf is used as a fallback if
// present. We fill the packet's named AcroForm text fields with pdf-lib.
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const { PDFDocument, StandardFonts } = require('pdf-lib');

const TEMPLATE_FILE = 'maplemark_new_account_application.pdf';

// Fields that must exist in a valid MapleMark packet (sanity-check on upload).
const REQUIRED_FIELDS = ['Entity Name', 'TIN', 'Entity Type'];

// ── SharePoint drives to search for org documents ────────────────────────────
// Default to the two libraries where Banyan's filed org docs live (same as the
// Dissolution feature). Override with BANKAPP_SP_DRIVES (comma-separated Graph
// driveIds) or BANKAPP_SP_SITES (comma-separated site names).
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
let driveIdsCache = null;
async function resolveDriveIds() {
  if (driveIdsCache) return driveIdsCache;
  const ids = [];
  const explicit = (process.env.BANKAPP_SP_DRIVES || '').split(',').map((x) => x.trim()).filter(Boolean);
  for (const id of explicit) if (!ids.includes(id)) ids.push(id);
  const sites = (process.env.BANKAPP_SP_SITES || '').split(',').map((x) => x.trim()).filter(Boolean);
  for (const name of sites) {
    try {
      const sr = await (await gfetch('/sites?search=' + encodeURIComponent(name))).json();
      const site = (sr.value || []).find((s) => (s.displayName || '').toLowerCase() === name.toLowerCase()) || (sr.value || [])[0];
      if (site) { const d = (await (await gfetch('/sites/' + site.id + '/drive')).json()).id; if (d && !ids.includes(d)) ids.push(d); }
    } catch { /* non-fatal */ }
  }
  if (!ids.length) {
    for (const id of DEFAULT_DRIVES) ids.push(id);
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

// ── entity-name helpers / entity type ────────────────────────────────────────
const LEGAL_SUFFIX = /[,\s]+(?:llc|l\.l\.c\.|lp|l\.p\.|llp|ltd|inc|incorporated|corp|corporation|company|co|gp|l\.?p\.?)\.?$/i;
function coreName(entity) {
  let s = String(entity || '').trim(); let prev;
  do { prev = s; s = s.replace(LEGAL_SUFFIX, '').trim(); } while (s !== prev);
  return s;
}
function tokens(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter((t) => t && t.length >= 2 && !['the', 'llc', 'lp', 'inc'].includes(t));
}
// Derive the entity type for the "ENTITY TYPE" line from the legal-name suffix.
function entityTypeFromName(name) {
  const n = ' ' + String(name || '').toLowerCase().replace(/[.,]/g, ' ').replace(/\s+/g, ' ') + ' ';
  if (/\bl l c\b|\bllc\b/.test(n)) return 'Limited Liability Company';
  if (/\bl l p\b|\bllp\b/.test(n)) return 'Limited Liability Partnership';
  if (/\bl p\b|\blp\b|\blimited partnership\b/.test(n)) return 'Limited Partnership';
  if (/\bgp\b|\bgeneral partnership\b/.test(n)) return 'General Partnership';
  if (/\binc\b|\bincorporated\b|\bcorp\b|\bcorporation\b/.test(n)) return 'Corporation';
  if (/\btrust\b/.test(n)) return 'Trust';
  if (/\bltd\b|\blimited\b/.test(n)) return 'Limited';
  return '';
}

// ── text extraction (EIN + address) ──────────────────────────────────────────
function extractEin(text) {
  if (!text) return null;
  const T = text.replace(/\s+/g, ' ');
  // Prefer an EIN that sits next to an EIN label.
  let m = T.match(/(?:EIN|Employer\s+Identification(?:\s+Number)?|Federal\s+(?:Tax\s+)?ID(?:entification)?(?:\s+Number)?|Taxpayer\s+Identification(?:\s+Number)?)\b[^0-9]{0,24}(\d{2})-?(\d{7})\b/i);
  if (m) return m[1] + '-' + m[2];
  // Otherwise the first well-formed EIN (NN-NNNNNNN).
  m = T.match(/\b(\d{2})-(\d{7})\b/);
  if (m) return m[1] + '-' + m[2];
  return null;
}
const US_STATES = 'AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC';
// Best-effort address block: a "City, ST ZIP" line, with the preceding line as street.
function extractAddress(text) {
  if (!text) return null;
  const lines = String(text).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const cityRe = new RegExp('^(.+?),\\s*(' + US_STATES + ')\\.?\\s+(\\d{5}(?:-\\d{4})?)$', 'i');
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(cityRe);
    if (!m) continue;
    const city = m[1].trim(), state = m[2].toUpperCase(), zip = m[3];
    // Street = the nearest preceding line that looks like a street address.
    let street = '';
    for (let j = i - 1; j >= 0 && j >= i - 3; j--) {
      const p = lines[j];
      if (/\d/.test(p) && /[A-Za-z]/.test(p) && !/^\d{2}-?\d{7}$/.test(p) && p.length <= 80 &&
          !/employer|identification|internal revenue|department|form\s|request for|taxpayer/i.test(p)) { street = p; break; }
    }
    return { street, city, state, zip };
  }
  // Fallback: inline "City, ST ZIP" anywhere.
  const T = text.replace(/\s+/g, ' ');
  const m2 = T.match(new RegExp('([A-Za-z .\'\\-]{2,40}),\\s*(' + US_STATES + ')\\.?\\s+(\\d{5}(?:-\\d{4})?)', 'i'));
  if (m2) return { street: '', city: m2[1].trim(), state: m2[2].toUpperCase(), zip: m2[3] };
  return null;
}

// ── candidate scoring ────────────────────────────────────────────────────────
const DOC_HINT = /w-?9|cp.?575|ss.?4|employer\s+identification|ein|certificate\s+of\s+formation|articles\s+of\s+organization/i;
const NEGATIVE_HINT = /cancellation|dissolution|dissolv|amend|restat|operating\s+agreement|llc\s+agreement|resolution|consent|good\s*standing|draft|template/i;
function scoreCandidate(entity, name) {
  const n = String(name || '').toLowerCase();
  const toks = tokens(entity);
  if (!toks.length) return -999;
  const hits = toks.filter((t) => n.includes(t)).length;
  const coverage = hits / toks.length;
  if (coverage < 0.6) return -999;
  let sc = hits * 3 + coverage * 2;
  if (/\bw-?9\b/.test(n)) sc += 6;              // W-9: name + address + EIN in one form
  else if (/cp.?575|ss.?4|employer\s+identification|\bein\b/.test(n)) sc += 5; // EIN letter
  else if (/certificate\s+of\s+formation|articles\s+of\s+organization/.test(n)) sc += 3;
  else if (DOC_HINT.test(n)) sc += 1;
  if (/\.pdf$/.test(n)) sc += 1;
  if (NEGATIVE_HINT.test(n)) sc -= 6;
  return sc;
}

// Live SharePoint search for an entity's identity (EIN + address).
async function searchProfileFromSharePoint(entity) {
  const drives = await resolveDriveIds();
  const core = coreName(entity);
  const queries = [core + ' W-9', core + ' EIN', core + ' certificate of formation', core];
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
    }
  }
  candidates.sort((a, b) => b.score - a.score);
  const out = { candidates: candidates.slice(0, 8) };
  let pdfParse; try { pdfParse = require('pdf-parse'); } catch { pdfParse = null; }
  for (const c of candidates.slice(0, 4)) {
    if (!/\.pdf$/i.test(c.name) || !pdfParse) continue;
    let text = null;
    try {
      const buf = Buffer.from(await (await gfetch('/drives/' + c.driveId + '/items/' + c.itemId + '/content')).arrayBuffer());
      const parsed = await pdfParse(buf).catch(() => null);
      if (parsed && parsed.text) text = parsed.text;
    } catch { /* try next candidate */ }
    if (!text) continue;
    if (!out.ein) { const e = extractEin(text); if (e) { out.ein = e; out.ein_source = c.name; out.source_url = c.webUrl; } }
    if (!out.address) { const a = extractAddress(text); if (a) { out.address = a; out.address_source = c.name; if (!out.source_url) out.source_url = c.webUrl; } }
    if (out.ein && out.address) break;
  }
  if (out.ein || out.address) { out.source = 'sharepoint'; out.source_file = out.ein_source || out.address_source; }
  return out;
}

// ── store (entity_bank_profile) ──────────────────────────────────────────────
function ensureTable(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS entity_bank_profile (
    entity_key   TEXT PRIMARY KEY,
    entity_name  TEXT NOT NULL,
    entity_type  TEXT,
    ein          TEXT,
    dba          TEXT,
    phys_address TEXT, phys_city TEXT, phys_state TEXT, phys_zip TEXT,
    mail_address TEXT, mail_city TEXT, mail_state TEXT, mail_zip TEXT,
    email        TEXT,
    phone        TEXT,
    nature       TEXT,
    source       TEXT,          -- 'sharepoint' | 'manual' | 'derived'
    source_file  TEXT,
    source_url   TEXT,
    updated_at   TEXT DEFAULT (datetime('now'))
  )`);
}
const COLS = ['entity_name', 'entity_type', 'ein', 'dba', 'phys_address', 'phys_city', 'phys_state', 'phys_zip',
  'mail_address', 'mail_city', 'mail_state', 'mail_zip', 'email', 'phone', 'nature', 'source', 'source_file', 'source_url'];
const keyOf = (name) => String(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
function getStored(db, name) {
  return db.prepare('SELECT * FROM entity_bank_profile WHERE entity_key = ?').get(keyOf(name)) || null;
}
function upsertStored(db, rec) {
  const row = {}; for (const c of COLS) row[c] = rec[c] != null ? rec[c] : null;
  row.entity_key = keyOf(rec.entity_name);
  db.prepare(`INSERT INTO entity_bank_profile (entity_key, ${COLS.join(', ')}, updated_at)
    VALUES (@entity_key, ${COLS.map((c) => '@' + c).join(', ')}, datetime('now'))
    ON CONFLICT(entity_key) DO UPDATE SET
      ${COLS.map((c) => c + '=excluded.' + c).join(', ')}, updated_at=datetime('now')`).run(row);
}

// Resolve an entity's bank profile: store first, then live SharePoint.
async function resolveProfile(db, entity, { refresh = false } = {}) {
  const res = { entity, graph_configured: graphConfigured(), entity_type: entityTypeFromName(entity) };
  if (!refresh) {
    const s = getStored(db, entity);
    if (s && (s.ein || s.phys_address || s.source === 'manual')) {
      return Object.assign(res, { found: true, source: s.source || 'store', stored: s,
        entity_type: s.entity_type || res.entity_type, ein: s.ein,
        phys: { address: s.phys_address, city: s.phys_city, state: s.phys_state, zip: s.phys_zip },
        mail: { address: s.mail_address, city: s.mail_city, state: s.mail_state, zip: s.mail_zip },
        dba: s.dba, email: s.email, phone: s.phone, nature: s.nature,
        source_file: s.source_file, source_url: s.source_url });
    }
  }
  if (graphConfigured()) {
    try {
      const sp = await searchProfileFromSharePoint(entity);
      const phys = sp.address ? { address: sp.address.street, city: sp.address.city, state: sp.address.state, zip: sp.address.zip } : {};
      if (sp.ein || sp.address) {
        upsertStored(db, { entity_name: entity, entity_type: res.entity_type, ein: sp.ein || null,
          phys_address: phys.address || null, phys_city: phys.city || null, phys_state: phys.state || null, phys_zip: phys.zip || null,
          source: 'sharepoint', source_file: sp.source_file || null, source_url: sp.source_url || null });
        return Object.assign(res, { found: true, source: 'sharepoint', candidates: sp.candidates,
          ein: sp.ein, phys, mail: {}, source_file: sp.source_file, source_url: sp.source_url });
      }
      return Object.assign(res, { found: false, candidates: sp.candidates || [] });
    } catch (e) {
      return Object.assign(res, { found: false, error: e.message, candidates: [] });
    }
  }
  return Object.assign(res, { found: false, candidates: [] });
}

// ── template fill ────────────────────────────────────────────────────────────
async function fillBankApplication(templateBuf, f) {
  const pdf = await PDFDocument.load(templateBuf);
  const form = pdf.getForm();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const setT = (name, val) => {
    if (val == null || val === '') return;
    try { form.getTextField(name).setText(String(val)); } catch { /* field absent — ignore */ }
  };
  const check = (name) => { try { form.getCheckBox(name).check(); } catch { /* ignore */ } };

  // Entity identity (page 1; Entity Name also drives the Beneficial Ownership page).
  setT('Entity Name', f.entity);
  setT('Entity Type', f.entityType);
  setT('TIN', f.ein);
  setT('DBA', f.dba);
  setT('Nature of Business/Industry', f.nature);
  setT('Entity Email Address', f.email);
  setT('Entity Phone Number', f.phone);
  // Physical address
  setT('Address.Physical', f.physAddress);
  setT('City.Phyiscal', f.physCity);   // field name misspelled in the form — keep verbatim
  setT('State.Physical', f.physState);
  setT('ZIP.Physical', f.physZip);
  // Mailing address (mirror physical when no distinct mailing given)
  setT('Address.Mailing', f.mailAddress || f.physAddress);
  setT('City.Mailing', f.mailCity || f.physCity);
  setT('State.Mailing', f.mailState || f.physState);
  setT('ZIP.Mailing', f.mailZip || f.physZip);
  // Entity attributes: a real-estate fund entity is none of the regulated categories.
  if (f.attributesNone !== false) check('NONE OF THE FOLLOWING');
  // Optional: person opening the account (Beneficial Ownership p.2 top line). Left
  // blank unless explicitly provided — it is a human, not entity-identity data.
  setT('Individual Name 1', f.personOpening);

  try { form.updateFieldAppearances(font); } catch { /* ignore */ }
  if (f.flatten === true) { try { form.flatten(); } catch { /* ignore */ } }
  return Buffer.from(await pdf.save());
}

// ── routes ───────────────────────────────────────────────────────────────────
function registerBankAccountRoutes(app, ctx) {
  const { db, auth, requireRole, memUpload, templatesDir } = ctx;
  ensureTable(db);
  try { fs.mkdirSync(templatesDir, { recursive: true }); } catch { /* ignore */ }
  const tplPath = () => path.join(templatesDir, TEMPLATE_FILE);
  const bundledTplPath = path.join(__dirname, 'assets', TEMPLATE_FILE);
  // Prefer an uploaded override on the data volume; else the repo-bundled blank.
  const activeTplPath = () => (fs.existsSync(tplPath()) ? tplPath() : (fs.existsSync(bundledTplPath) ? bundledTplPath : null));
  const BOTH = ['Admin', 'Accountant'];

  // Feature status — drives the page.
  app.get('/api/admin/bankapp/status', auth, requireRole(...BOTH), (req, res) => {
    res.json({
      template_installed: fs.existsSync(tplPath()),
      bundled_template: fs.existsSync(bundledTplPath),
      template_available: !!activeTplPath(),
      graph_configured: graphConfigured(),
    });
  });

  // Install / replace the blank MapleMark packet (Admin).
  app.post('/api/admin/bankapp/template', auth, requireRole('Admin'), memUpload.single('file'), async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) return res.status(400).json({ error: 'No file uploaded.' });
      let ok = false;
      try {
        const d = await PDFDocument.load(req.file.buffer);
        const names = d.getForm().getFields().map((x) => x.getName());
        ok = REQUIRED_FIELDS.every((n) => names.includes(n));
      } catch { ok = false; }
      if (!ok) return res.status(400).json({ error: 'That does not look like the MapleMark New Account Application packet (expected fillable fields Entity Name / TIN / Entity Type).' });
      fs.writeFileSync(tplPath(), req.file.buffer);
      res.json({ note: 'Template installed.', template_installed: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Look up an entity's bank profile (store, then live SharePoint if configured).
  app.get('/api/admin/bankapp/lookup', auth, requireRole(...BOTH), async (req, res) => {
    try {
      const entity = (req.query.entity || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      const refresh = req.query.refresh === '1' || req.query.refresh === 'true';
      res.json(await resolveProfile(db, entity, { refresh }));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Confirm / override a bank profile in the store (manual entry).
  app.post('/api/admin/bankapp/profile', auth, requireRole(...BOTH), (req, res) => {
    try {
      const b = req.body || {};
      const entity = (b.entity || b.entity_name || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      upsertStored(db, {
        entity_name: entity,
        entity_type: b.entity_type != null ? b.entity_type : entityTypeFromName(entity),
        ein: b.ein, dba: b.dba,
        phys_address: b.phys_address, phys_city: b.phys_city, phys_state: b.phys_state, phys_zip: b.phys_zip,
        mail_address: b.mail_address, mail_city: b.mail_city, mail_state: b.mail_state, mail_zip: b.mail_zip,
        email: b.email, phone: b.phone, nature: b.nature,
        source: 'manual', source_file: b.source_file || null, source_url: b.source_url || null,
      });
      res.json(Object.assign({ saved: true }, getStored(db, entity)));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // List everything in the store (small management table).
  app.get('/api/admin/bankapp/profiles', auth, requireRole(...BOTH), (req, res) => {
    try { res.json({ rows: db.prepare('SELECT * FROM entity_bank_profile ORDER BY entity_name').all() }); }
    catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Generate the filled MapleMark packet.
  app.post('/api/admin/bankapp/generate', auth, requireRole(...BOTH), async (req, res) => {
    try {
      const b = req.body || {};
      const entity = (b.entity || '').toString().trim();
      if (!entity) return res.status(400).json({ error: 'entity is required' });
      const active = activeTplPath();
      if (!active) return res.status(400).json({ error: 'No MapleMark packet is installed. Upload the blank New Account Application packet once (Admin) to enable generation.' });
      const template = fs.readFileSync(active);
      const buf = await fillBankApplication(template, {
        entity,
        entityType: (b.entity_type != null ? b.entity_type : entityTypeFromName(entity)) || '',
        ein: b.ein || '', dba: b.dba || '', nature: b.nature || '',
        email: b.email || '', phone: b.phone || '',
        physAddress: b.phys_address || '', physCity: b.phys_city || '', physState: b.phys_state || '', physZip: b.phys_zip || '',
        mailAddress: b.mail_address || '', mailCity: b.mail_city || '', mailState: b.mail_state || '', mailZip: b.mail_zip || '',
        attributesNone: b.attributes_none !== false,
        personOpening: b.person_opening || '',
        flatten: b.flatten === true,
      });
      const safe = entity.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'Entity';
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="MapleMark_New_Account_Application_' + safe + '.pdf"');
      res.send(buf);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
}

module.exports = {
  registerBankAccountRoutes,
  // exported for local testing:
  fillBankApplication, resolveProfile, searchProfileFromSharePoint,
  extractEin, extractAddress, entityTypeFromName, coreName, scoreCandidate,
};
