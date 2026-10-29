// ─── Administration: USC Partner Capital Account Statements (CLIP + Silsbee) ──
//
// Quarterly "Partner Capital Quarterly Account Statement" for the University of
// Southern California's membership interest in County Line Industrial Park LLC
// (CLIP Property Owner, entity 54) and County Line Rail Silsbee LLC (CLR Silsbee
// Property Owner, entity 39), reproducing the fund administrator's one-page
// statement and the cover email to BNY Mellon et al.
//
// ── Data model ───────────────────────────────────────────────────────────────
// CloudLedger's ledgers for these entities track USC's CONTRIBUTED capital as a
// single named account (34261 "Contributed Capital - USC") — $8,000,000 (CLIP)
// and $2,000,000 (Silsbee). They do NOT carry USC's full partner capital account
// (beginning balance incl. cumulative allocated earnings / unrealized, the
// ownership %, or the year-end net-loss / unrealized allocations): those live on
// the administrator's maintained capital-account schedule. So this generator is
// a carry-forward model:
//   • Fixed annual inputs per (entity, fiscal year) — beginning 1/1 balance,
//     committed capital, ownership %, and any year-end allocation amounts —
//     are stored in usc_capital_config (seeded 2026 from the Q2-2026 statement).
//   • CONTRIBUTIONS are pulled LIVE from the GL (year-to-date movement in account
//     34261), which also validates committed capital against the live balance.
//   • Change in unrealized appreciation and net loss are held at 0 intra-year
//     (the administrator freezes them until the annual valuation / allocation),
//     and can be entered on the config row for a year-end (12/31) statement.
//   • Ending = beginning + contributions + distributions + unrealized + net loss.
// This reproduces the Q2-2026 statement to the dollar and rolls the date forward.
//
// Prepared like the CLRF workpapers (pcap.js): buildData -> render PDF ->
// saveToWorkpapers, plus a "Download All" zip and an optional filing of both
// statements into SharePoint (Accounting > General > Investors > Capital
// Statements > USC > <quarter>) when Microsoft Graph is configured.

const path = require('path');
const fs = require('fs');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

// SharePoint target (the investor Capital Statements library). driveId and base
// folder are overridable by env; a dated subfolder (e.g. 9-30-26) is created per
// quarter to match the existing filing convention.
const SP_DRIVE_ID = process.env.USC_SP_DRIVE_ID
  || 'b!xwItU4t_F0ytk_y7mfLzczKCgKnLYEZEn631Bad3dowjH1fJ6JpNQp29UcjVRg0R';
const SP_BASE_PATH = process.env.USC_SP_BASE_PATH || 'General/Investors/Capital Statements/USC';

function resolveQuarter(asOf) {
  if (!isDate(asOf)) throw new Error('as_of must be a date in YYYY-MM-DD form');
  const [y, m, d] = asOf.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) {
    throw new Error('as_of must be a quarter end date: 03-31, 06-30, 09-30 or 12-31. Received ' + asOf);
  }
  const n = m / 3; // 1..4
  const ORD = { 1: '1st', 2: '2nd', 3: '3rd', 4: '4th' };
  return {
    year: y, quarter: n, label: n + 'Q' + y, ordinal: ORD[n],
    end: asOf, year_start: y + '-01-01', prior_year_end: (y - 1) + '-12-31',
    spelled: MONTHS[m - 1] + ' ' + d + ', ' + y,     // "September 30, 2026"
    short: m + '/' + d + '/' + String(y).slice(2),   // "9/30/26"
    fileDate: m + '-' + d + '-' + String(y).slice(2), // "9-30-26"
  };
}

// ── Config (carry-forward annual inputs), seeded for 2026 ─────────────────────
function ensureConfig(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS usc_capital_config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id INTEGER NOT NULL,
    fiscal_year INTEGER NOT NULL,
    short_name TEXT NOT NULL,
    legal_name TEXT NOT NULL,
    usc_account_code TEXT NOT NULL,
    committed_capital REAL NOT NULL,
    ownership_pct REAL NOT NULL,
    pct_decimals INTEGER NOT NULL DEFAULT 2,
    beginning_balance REAL NOT NULL,
    ytd_distribution REAL NOT NULL DEFAULT 0,
    ytd_unrealized REAL NOT NULL DEFAULT 0,
    ytd_net_loss REAL NOT NULL DEFAULT 0,
    sort_order INTEGER NOT NULL DEFAULT 0,
    UNIQUE(entity_id, fiscal_year)
  )`);
  // Seed 2026 from the fund administrator's 2Q2026 statement. INSERT OR IGNORE
  // preserves any later edits Jimmy makes (year rollover updates beginning_balance).
  const ins = db.prepare(`INSERT OR IGNORE INTO usc_capital_config
    (entity_id, fiscal_year, short_name, legal_name, usc_account_code,
     committed_capital, ownership_pct, pct_decimals, beginning_balance, sort_order)
    VALUES (?,?,?,?,?,?,?,?,?,?)`);
  ins.run(54, 2026, 'CLIP', 'County Line Industrial Park LLC', '34261', 8000000, 0.11327, 3, 9444247, 1);
  ins.run(39, 2026, 'Silsbee', 'County Line Rail Silsbee LLC', '34261', 2000000, 0.1707, 2, 2635906, 2);
}

function getConfig(db, year) {
  return db.prepare('SELECT * FROM usc_capital_config WHERE fiscal_year = ? ORDER BY sort_order, id').all(year);
}

// Point-in-time balance of one account (equity: balance = credit - debit, a
// contribution reads positive), from the live GL via computeBalances.
function acctBalance(ctx, eid, code, asOf) {
  const rows = ctx.computeBalances(eid, { as_of: asOf }) || [];
  const row = rows.find((r) => String(r.code) === String(code));
  return row ? r2(Number(row.balance) || 0) : 0;
}

function buildEntity(ctx, cfg, q) {
  const contributedNow = acctBalance(ctx, cfg.entity_id, cfg.usc_account_code, q.end);
  const contributedPriorYE = acctBalance(ctx, cfg.entity_id, cfg.usc_account_code, q.prior_year_end);
  const ytdContribution = r2(contributedNow - contributedPriorYE); // net calls less returns, this FY
  const distribution = r2(cfg.ytd_distribution || 0);
  const unrealized = r2(cfg.ytd_unrealized || 0);
  const netLoss = r2(cfg.ytd_net_loss || 0);
  const beginning = r2(cfg.beginning_balance);
  const ending = r2(beginning + ytdContribution + distribution + unrealized + netLoss);
  const committed = r2(cfg.committed_capital);
  const unfunded = r2(committed - contributedNow);
  const warnings = [];
  if (Math.abs(contributedNow - committed) > 0.5) {
    warnings.push(cfg.short_name + ': live contributed capital (' + contributedNow.toLocaleString('en-US')
      + ') does not equal committed capital on file (' + committed.toLocaleString('en-US') + ').');
  }
  return {
    entity_id: cfg.entity_id, short_name: cfg.short_name, legal_name: cfg.legal_name,
    committed, ownership_pct: cfg.ownership_pct, pct_decimals: cfg.pct_decimals,
    contributed: contributedNow, unfunded,
    beginning, contribution: ytdContribution, distribution, unrealized, netLoss, ending,
    warnings,
  };
}

function buildData(ctx, q) {
  ensureConfig(ctx.db);
  const cfgs = getConfig(ctx.db, q.year);
  if (!cfgs.length) throw new Error('No USC capital config for fiscal year ' + q.year + '. Add one on the config.');
  const entities = cfgs.map((c) => buildEntity(ctx, c, q));
  const warnings = entities.flatMap((e) => e.warnings);
  return { quarter: q, investor: 'University of Southern California', entities, warnings };
}

// ── PDF (one page per entity, matching the administrator's statement) ─────────
async function renderStatementPdf(ent, q) {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.TimesRoman);
  const bold = await pdf.embedFont(StandardFonts.TimesRomanBold);
  const W = 612, PH = 792, mL = 72, RIGHT = 540;
  const page = pdf.addPage([W, PH]);
  const txt = (s, x, y, f, size) => page.drawText(String(s), { x, y, size, font: f });
  const rt = (s, xR, y, f, size) => { const w = f.widthOfTextAtSize(String(s), size); page.drawText(String(s), { x: xR - w, y, size, font: f }); };
  const ctr = (s, y, f, size) => { const w = f.widthOfTextAtSize(String(s), size); page.drawText(String(s), { x: (W - w) / 2, y, size, font: f }); };
  const rule = (x0, x1, y, th) => page.drawLine({ start: { x: x0, y }, end: { x: x1, y }, thickness: th || 0.75, color: rgb(0, 0, 0) });
  const money = (v) => { const n = Math.round(Number(v) || 0); if (n === 0) return '-'; const s = Math.abs(n).toLocaleString('en-US'); return n < 0 ? '(' + s + ')' : s; };
  const dollar = (v) => { const n = Math.round(Number(v) || 0); if (n === 0) return '$  -'; const s = Math.abs(n).toLocaleString('en-US'); return n < 0 ? '$ (' + s + ')' : '$ ' + s; };
  const pct = (f, dec) => ((Number(f) || 0) * 100).toFixed(dec) + '%';

  let y = PH - 72;
  txt(ent.legal_name, mL, y, bold, 12); y -= 18;
  txt('Partner Capital Quarterly Account Statement', mL, y, reg, 10); y -= 14;
  txt(q.spelled, mL, y, reg, 10); y -= 42;
  txt('University of Southern California', mL, y, bold, 18); y -= 40;

  // Section 1 — Partnership Interest (commitment)
  ctr('Partnership Interest', y, reg, 11); rule(mL, RIGHT, y - 4, 1.1); y -= 24;
  txt('Committed Capital', mL, y, reg, 10); rt(dollar(ent.committed), RIGHT, y, reg, 10); y -= 18;
  txt('Partnership Ownership %', mL, y, reg, 10); rt(pct(ent.ownership_pct, ent.pct_decimals), RIGHT, y, reg, 10); y -= 18;
  txt('Unfunded Commitment', mL, y, reg, 10); rt(money(ent.unfunded), RIGHT, y, reg, 10); y -= 36;

  // Section 2 — Partnership Interest (capital roll-forward)
  ctr('Partnership Interest', y, reg, 11); rule(mL, RIGHT, y - 4, 1.1); y -= 24;
  txt('Beginning Capital Account Balance 1/1/' + q.year, mL, y, reg, 10); rt(money(ent.beginning), RIGHT, y, reg, 10); y -= 22;
  txt('Contribution', mL + 18, y, reg, 10); rt(money(ent.contribution), RIGHT, y, reg, 10); y -= 16;
  txt('Distribution', mL + 18, y, reg, 10); rt(money(ent.distribution), RIGHT, y, reg, 10); y -= 16;
  txt('Change in unrealized appreciation on investments', mL + 18, y, reg, 10); rt(money(ent.unrealized), RIGHT, y, reg, 10); y -= 16;
  txt('Net loss', mL + 18, y, reg, 10); rt(money(ent.netLoss), RIGHT, y, reg, 10); y -= 20;
  rule(RIGHT - 110, RIGHT, y + 12, 0.6); // thin rule above the total
  txt('Ending Capital Account Balance ' + q.short, mL, y, reg, 10); rt(money(ent.ending), RIGHT, y, reg, 10);
  rule(RIGHT - 110, RIGHT, y - 3, 0.6); rule(RIGHT - 110, RIGHT, y - 5, 0.6); // double rule below

  return Buffer.from(await pdf.save());
}

function fileNameFor(ent, q) { return 'USC ' + ent.short_name + ' Capital Statements ' + q.fileDate + '.pdf'; }

// ── "Download All" zip (both statements) ─────────────────────────────────────
async function buildZip(data) {
  const JSZip = require('jszip');
  const zip = new JSZip();
  for (const ent of data.entities) {
    const buf = await renderStatementPdf(ent, data.quarter);
    zip.file(fileNameFor(ent, data.quarter), buf);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

// ── Optional: file both statements into SharePoint via MS Graph ──────────────
// App-only (client credentials). Requires MS_TENANT_ID / MS_CLIENT_ID /
// MS_CLIENT_SECRET and the Graph application permission Sites.ReadWrite.All
// (admin consent). Targets the investor library at SP_BASE_PATH, creating a
// dated subfolder per quarter. Degrades gracefully when unconfigured: the
// statements are still generated, saved to Workpapers and returned.
async function graphToken() {
  const t = process.env.MS_TENANT_ID, c = process.env.MS_CLIENT_ID, s = process.env.MS_CLIENT_SECRET;
  if (!t || !c || !s) return null;
  const r = await fetch('https://login.microsoftonline.com/' + t + '/oauth2/v2.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: c, client_secret: s, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!r.ok) throw new Error('Graph token ' + r.status + ': ' + (await r.text()).slice(0, 200));
  return (await r.json()).access_token;
}

const spEncode = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

async function uploadToSharePoint(pdfs, q) {
  const driveId = SP_DRIVE_ID;
  if (!driveId) return { uploaded: false, reason: 'USC_SP_DRIVE_ID not set.' };
  let token;
  try { token = await graphToken(); } catch (e) { return { uploaded: false, reason: e.message }; }
  if (!token) return { uploaded: false, reason: 'Microsoft Graph not configured (set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET).' };
  const auth = { Authorization: 'Bearer ' + token };
  const driveRoot = 'https://graph.microsoft.com/v1.0/drives/' + driveId;
  // Ensure the dated subfolder exists (idempotent: a 409 means it is already there).
  try {
    const mk = await fetch(driveRoot + '/root:/' + spEncode(SP_BASE_PATH) + ':/children', {
      method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, auth),
      body: JSON.stringify({ name: q.fileDate, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    if (!mk.ok && mk.status !== 409) return { uploaded: false, reason: 'create folder ' + mk.status + ': ' + (await mk.text()).slice(0, 200) };
  } catch (e) { return { uploaded: false, reason: e.message }; }
  const folderPath = SP_BASE_PATH + '/' + q.fileDate;
  const files = [];
  for (const p of pdfs) {
    const up = await fetch(driveRoot + '/root:/' + spEncode(folderPath + '/' + p.filename) + ':/content', {
      method: 'PUT', headers: Object.assign({ 'Content-Type': 'application/pdf' }, auth), body: p.buf,
    });
    if (!up.ok) return { uploaded: false, partial: files.length > 0, folder: folderPath, files, reason: 'upload "' + p.filename + '" ' + up.status + ': ' + (await up.text()).slice(0, 200) };
    const j = await up.json().catch(() => ({}));
    files.push({ filename: p.filename, webUrl: j.webUrl });
  }
  return { uploaded: true, folder: folderPath, files };
}

// ── Persist each statement into the entity's Workpapers ───────────────────────
function saveToWorkpapers(ctx, eid, q, filename, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = 'Workpapers/USC Capital Statements/' + q.year + '/Q' + q.quarter;
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, filename);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + filename.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, filename, buf.length, 'application/pdf', who);
  return folder + '/' + filename;
}

function registerUscCapitalRoutes(app, ctx) {
  const { db, auth, requireRole } = ctx;
  ensureConfig(db);

  // Config for a fiscal year (+ live contributed capital for validation display).
  app.get('/api/admin/usc-capital/config', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    try {
      const year = Number((req.query && req.query.year) || new Date().getFullYear());
      ensureConfig(db);
      const rows = getConfig(db, year).map((c) => Object.assign({}, c, {
        live_contributed: (() => { try { return acctBalance(ctx, c.entity_id, c.usc_account_code, year + '-12-31'); } catch (e) { return null; } })(),
      }));
      res.json({ fiscal_year: year, rows });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Single statement PDF (per entity) — used by the page's download buttons.
  app.get('/api/admin/usc-capital/statement.pdf', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const q = resolveQuarter((req.query && req.query.as_of) || '');
      const eid = Number(req.query && req.query.entity_id);
      const data = buildData(ctx, q);
      const ent = data.entities.find((e) => e.entity_id === eid);
      if (!ent) return res.status(404).json({ error: 'No USC config for entity ' + eid + ' in ' + q.year });
      const buf = await renderStatementPdf(ent, q);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fileNameFor(ent, q) + '"');
      res.send(buf);
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Download All — both statements as a single zip.
  app.get('/api/admin/usc-capital/statements.zip', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const q = resolveQuarter((req.query && req.query.as_of) || '');
      const data = buildData(ctx, q);
      const buf = await buildZip(data);
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename="USC Capital Statements ' + q.fileDate + '.zip"');
      res.send(buf);
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Run report: generate both statements, save to Workpapers, and file them to
  // SharePoint (if Graph is configured). Returns the figures + save status.
  app.post('/api/admin/usc-capital/run', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const q = resolveQuarter((req.body && req.body.as_of) || '');
      const who = (req.user && (req.user.email || req.user.name)) || 'system';
      const data = buildData(ctx, q);
      const pdfs = [];
      for (const ent of data.entities) {
        const buf = await renderStatementPdf(ent, q);
        const filename = fileNameFor(ent, q);
        const savedTo = saveToWorkpapers(ctx, ent.entity_id, q, filename, buf, who);
        pdfs.push({ entity_id: ent.entity_id, short_name: ent.short_name, filename, buf, saved_to: savedTo });
      }
      const makeSp = !req.body || req.body.file_to_sharepoint !== false;
      const sharepoint = makeSp ? await uploadToSharePoint(pdfs, q) : { uploaded: false, reason: 'skipped' };
      res.json({
        quarter: q.label, as_of: q.end,
        statements: data.entities.map((e) => ({
          entity_id: e.entity_id, short_name: e.short_name, legal_name: e.legal_name,
          committed: e.committed, ownership_pct: e.ownership_pct, beginning: e.beginning,
          contribution: e.contribution, distribution: e.distribution, unrealized: e.unrealized,
          net_loss: e.netLoss, ending: e.ending,
        })),
        files: pdfs.map((p) => ({ entity_id: p.entity_id, filename: p.filename, saved_to: p.saved_to })),
        sharepoint, warnings: data.warnings,
      });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });
}

module.exports = {
  resolveQuarter, ensureConfig, getConfig, buildData, renderStatementPdf,
  buildZip, uploadToSharePoint, fileNameFor, registerUscCapitalRoutes,
};
