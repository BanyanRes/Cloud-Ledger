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
// saveToWorkpapers; plus an email composer and an optional Microsoft Graph draft.

const path = require('path');
const fs = require('fs');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

// Fixed CC distribution list and sender signature, lifted verbatim from the
// administrator's cover email (Irvin Bermudez, 2Q2026) with Jimmy's signature.
const MAIL_TO = ['piaasstatements@bnymellon.com'];
const MAIL_CC = [
  'brian.j.johnson@usc.edu', 'real.assets@usc.edu', 'rohanesi@usc.edu',
  'usclps@bnymellon.com', 'bnylpaasconnect@canoeintelligence.com',
  'dsdata@burgiss.com', 'uscconnect@canoeintelligence.com',
  'JYun@banyanres.com', 'bbrosseau@banyanres.com', 'jrucker@banyanres.com',
  'WMyers@banyanres.com',
];
const CONFIDENTIALITY = 'IMPORTANT: This message is confidential. It may also be '
  + 'privileged or otherwise protected by work product immunity or other legal '
  + 'rules. If you have received it by mistake, please let us know by e-mail reply '
  + 'and delete it from your system; you may not copy this message or disclose its '
  + 'contents to anyone. The integrity and security of this message cannot be '
  + 'guaranteed on the Internet.';
const SIG_LINES = ['Jimmy Yun', 'Chief Accounting Officer', 'M: 714.321.6021',
  '2381 Rosecrans Ave Suite 425, El Segundo, CA 90245'];

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

// ── Email composer (cover email, administrator format, Jimmy's signature) ─────
function composeEmail(data, pdfs) {
  const q = data.quarter;
  const subject = 'USC Capital Statements for CLIP and Silsbee ' + q.label;
  const sigTextLines = [...SIG_LINES, 'banyanres.com'];
  const bodyText = [
    'Hello,', '',
    'Attached please find ' + q.ordinal + ' QTR ' + q.year + ' USC capital statements for CLIP and Silsbee.', '',
    'Please let us know if you have any questions.', '',
    'Thanks', '',
    ...sigTextLines, '',
    CONFIDENTIALITY, '',
  ].join('\n');
  const sigHtml = SIG_LINES.map((l) => escapeHtml(l)).join('<br>')
    + '<br><a href="https://banyanres.com">banyanres.com</a>';
  const bodyHtml = [
    '<p>Hello,</p>',
    '<p>Attached please find ' + q.ordinal + ' QTR ' + q.year + ' USC capital statements for CLIP and Silsbee.</p>',
    '<p>Please let us know if you have any questions.</p>',
    '<p>Thanks</p>',
    '<p>' + sigHtml + '</p>',
    '<p>' + escapeHtml(CONFIDENTIALITY) + '</p>',
  ].join('');
  return {
    to: MAIL_TO, cc: MAIL_CC, subject, bodyText, bodyHtml,
    attachments: pdfs.map((p) => ({ filename: p.filename, mime: 'application/pdf', b64: p.buf.toString('base64') })),
  };
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Optional: create the Outlook draft in the user's mailbox via MS Graph ─────
// App-only (client credentials). Requires MS_TENANT_ID / MS_CLIENT_ID /
// MS_CLIENT_SECRET and the Graph application permission Mail.ReadWrite (admin
// consent). USC_DRAFT_MAILBOX selects the mailbox (default JYun@banyanres.com).
// Degrades gracefully when unconfigured: the statements are still generated,
// saved and returned with the composed email.
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

async function createOutlookDraft(email) {
  const mailbox = process.env.USC_DRAFT_MAILBOX || 'JYun@banyanres.com';
  let token;
  try { token = await graphToken(); } catch (e) { return { created: false, reason: e.message }; }
  if (!token) return { created: false, reason: 'Microsoft Graph not configured (set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, optionally USC_DRAFT_MAILBOX).' };
  const base = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(mailbox);
  const payload = {
    subject: email.subject,
    body: { contentType: 'HTML', content: email.bodyHtml },
    toRecipients: email.to.map((a) => ({ emailAddress: { address: a } })),
    ccRecipients: email.cc.map((a) => ({ emailAddress: { address: a } })),
  };
  const cr = await fetch(base + '/messages', { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
  if (!cr.ok) return { created: false, reason: 'create message ' + cr.status + ': ' + (await cr.text()).slice(0, 300) };
  const msg = await cr.json();
  for (const att of email.attachments) {
    const ar = await fetch(base + '/messages/' + msg.id + '/attachments', {
      method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ '@odata.type': '#microsoft.graph.fileAttachment', name: att.filename, contentType: att.mime, contentBytes: att.b64 }),
    });
    if (!ar.ok) return { created: true, partial: true, webLink: msg.webLink, mailbox, reason: 'attachment "' + att.filename + '" failed ' + ar.status };
  }
  return { created: true, id: msg.id, webLink: msg.webLink, mailbox };
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

  // Run report: generate both statements, save to Workpapers, compose the email,
  // and create the Outlook draft (if Graph is configured). Returns everything so
  // the page can preview/download and the draft can be opened.
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
      const email = composeEmail(data, pdfs);
      const makeDraft = !req.body || req.body.create_draft !== false;
      const draft = makeDraft ? await createOutlookDraft(email) : { created: false, reason: 'skipped' };
      res.json({
        quarter: q.label, as_of: q.end,
        statements: data.entities.map((e) => ({
          entity_id: e.entity_id, short_name: e.short_name, legal_name: e.legal_name,
          committed: e.committed, ownership_pct: e.ownership_pct, beginning: e.beginning,
          contribution: e.contribution, distribution: e.distribution, unrealized: e.unrealized,
          net_loss: e.netLoss, ending: e.ending,
        })),
        files: pdfs.map((p) => ({ entity_id: p.entity_id, filename: p.filename, saved_to: p.saved_to, b64: p.buf.toString('base64') })),
        email: { to: email.to, cc: email.cc, subject: email.subject, body_text: email.bodyText },
        draft, warnings: data.warnings,
      });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });
}

module.exports = {
  resolveQuarter, ensureConfig, getConfig, buildData, renderStatementPdf,
  composeEmail, createOutlookDraft, fileNameFor, registerUscCapitalRoutes,
};
