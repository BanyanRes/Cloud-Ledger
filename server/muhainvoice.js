// ─── Administration: Muha Invoice (contractor semi-monthly invoices) ─────────
//
// Muha K. Kim is an independent contractor (Independent Contractor Services
// Agreement with Banyan Residential, LLC, effective 2026-11-11). Under the
// agreement the contractor invoices a fixed USD 3,000/month retainer in two
// semi-monthly installments (1st-15th and 16th-month end), plus a USD 50/month
// laptop allowance (USD 25 per half-month, 30 months) and a USD 19.99 wire fee
// per invoice. A partial billing period is prorated by working days (Mon-Fri)
// in effect / working days in that billing period.
//
// This module:
//   • builds the invoice schedule from the stored terms (muha_invoice_config),
//   • renders each invoice PDF (pdf-lib), matching the approved invoice layout,
//   • emails each invoice to the Bill.com inbox (Resend) N calendar days before
//     its invoice date (default 7), via an in-process scheduler, and
//   • logs every send in muha_invoice_sends so an invoice can never go out twice.
//
// Invoice date = last day of the billing period (the 15th or month end).
// Send date    = invoice date - lead_days calendar days, evaluated in Pacific time.

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const DEFAULTS = {
  contractor_name: 'Muha K. Kim',
  contractor_address: 'Bldg. 614 Apt# 202\nPyung-chon-daero 253 bungil 18\nAnyang-si Dongan-gu, Gyunggi-do\nSouth Korea',
  bill_to: 'Banyan Residential, LLC\nAttn: Jimmy Yun, Chief Accounting Officer\n2381 Rosecrans Ave, Suite 425\nEl Segundo, CA 90245\njyun@banyanres.com',
  effective_date: '2026-11-11',
  termination_date: '',
  monthly_fee: 3000,
  monthly_allowance: 50,
  allowance_months: 30,
  wire_fee: 19.99,
  invoice_prefix: 'MK',
  send_to: 'banyanresidential@bill.com',
  cc: 'jyun@banyanres.com',
  lead_days: 7,
  auto_send: 1,
};

// ── date helpers (all dates are plain YYYY-MM-DD strings, computed in UTC) ────
const D = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const S = (dt) => dt.toISOString().slice(0, 10);
const addDays = (s, n) => { const x = D(s); x.setUTCDate(x.getUTCDate() + n); return S(x); };
const eom = (y, m) => S(new Date(Date.UTC(y, m, 0))); // m is 1-based
function workdays(a, b) { // inclusive, Mon-Fri
  if (!a || !b || a > b) return 0;
  let n = 0;
  for (let x = D(a); S(x) <= b; x.setUTCDate(x.getUTCDate() + 1)) { const w = x.getUTCDay(); if (w !== 0 && w !== 6) n++; }
  return n;
}
const pretty = (s) => { const [y, m, d] = s.split('-').map(Number); return MON[m - 1] + ' ' + d + ', ' + y; };
const mdy = (s) => { const [y, m, d] = s.split('-'); return m + '/' + d + '/' + y; };
function pacificToday() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Los_Angeles', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', hour12: false }).formatToParts(new Date());
  const g = (t) => (parts.find((p) => p.type === t) || {}).value;
  return { date: g('year') + '-' + g('month') + '-' + g('day'), hour: Number(g('hour')) % 24 };
}

// ── storage ───────────────────────────────────────────────────────────────────
function ensureTables(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS muha_invoice_config (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      data TEXT NOT NULL,
      updated_by TEXT,
      updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS muha_invoice_sends (
      invoice_no TEXT PRIMARY KEY,
      status TEXT NOT NULL,          -- sending | sent | failed | skipped
      amount REAL,
      sent_to TEXT,
      send_trigger TEXT,             -- auto | manual
      error TEXT,
      sent_by TEXT,
      sent_at TEXT DEFAULT (datetime('now'))
    );`);
  const row = db.prepare('SELECT id FROM muha_invoice_config WHERE id = 1').get();
  if (!row) db.prepare('INSERT INTO muha_invoice_config (id, data, updated_by) VALUES (1, ?, ?)').run(JSON.stringify(DEFAULTS), 'seed');
  // One-time: Jimmy changed the send lead time from 5 to 7 calendar days
  // (2026-10-29). Applied once (tracked in data._migrations) so a later edit on
  // the settings page is never overwritten.
  {
    const cur = db.prepare('SELECT data FROM muha_invoice_config WHERE id = 1').get();
    let d = {}; try { d = JSON.parse(cur.data); } catch (e) { d = {}; }
    const done = Array.isArray(d._migrations) ? d._migrations : [];
    if (!done.includes('lead_days_7')) {
      d.lead_days = 7; d._migrations = done.concat('lead_days_7');
      db.prepare("UPDATE muha_invoice_config SET data = ?, updated_by = 'migration lead_days_7', updated_at = datetime('now') WHERE id = 1").run(JSON.stringify(d));
    }
  }
  // MK-2026-11A was sent to Bill.com outside CloudLedger before this page went
  // live (per Jimmy, 2026-10-29). Record it as sent so the scheduler never sends
  // it again. INSERT OR IGNORE: a no-op once the row exists.
  db.prepare("INSERT OR IGNORE INTO muha_invoice_sends (invoice_no, status, amount, sent_to, send_trigger, sent_by) "
    + "VALUES ('MK-2026-11A', 'sent', 477.49, 'banyanresidential@bill.com', 'external', 'sent outside CloudLedger')").run();
}

function getConfig(db) {
  ensureTables(db);
  const row = db.prepare('SELECT data, updated_at, updated_by FROM muha_invoice_config WHERE id = 1').get();
  let data = {};
  try { data = JSON.parse(row.data); } catch (e) { data = {}; }
  return Object.assign({}, DEFAULTS, data, { updated_at: row.updated_at, updated_by: row.updated_by });
}

function saveConfig(db, patch, who) {
  const cur = getConfig(db);
  const next = {};
  for (const k of Object.keys(DEFAULTS)) next[k] = (patch && Object.prototype.hasOwnProperty.call(patch, k)) ? patch[k] : cur[k];
  ['monthly_fee', 'monthly_allowance', 'wire_fee'].forEach((k) => { next[k] = r2(next[k]); });
  ['allowance_months', 'lead_days'].forEach((k) => { next[k] = Math.max(0, Math.round(Number(next[k]) || 0)); });
  next.auto_send = next.auto_send ? 1 : 0;
  if (Array.isArray(cur._migrations)) next._migrations = cur._migrations; // keep one-time migration markers
  if (!isDate(next.effective_date)) throw new Error('Effective date must be YYYY-MM-DD');
  if (next.termination_date && !isDate(next.termination_date)) throw new Error('Termination date must be YYYY-MM-DD or blank');
  if (!/^[^@\s]+@[^@\s]+$/.test(String(next.send_to || ''))) throw new Error('Send-to must be an email address');
  db.prepare("UPDATE muha_invoice_config SET data = ?, updated_by = ?, updated_at = datetime('now') WHERE id = 1").run(JSON.stringify(next), who || 'system');
  return getConfig(db);
}

// ── schedule ──────────────────────────────────────────────────────────────────
// One row per half-month billing period from the effective-date month until the
// termination date (or `throughMonths` months ahead when open-ended).
function buildSchedule(cfg, opts) {
  const o = opts || {};
  const eff = cfg.effective_date;
  const end = cfg.termination_date || '2099-12-31';
  const [ey, em] = eff.split('-').map(Number);
  const months = o.months || 30;
  const rows = [];
  for (let i = 0; i < months; i++) {
    const y = ey + Math.floor((em - 1 + i) / 12), m = ((em - 1 + i) % 12) + 1;
    const mm = String(m).padStart(2, '0');
    const halves = [[y + '-' + mm + '-01', y + '-' + mm + '-15', 'A', 1], [y + '-' + mm + '-16', eom(y, m), 'B', 2]];
    for (const [ps, pe, suf, half] of halves) {
      const as = ps > eff ? ps : eff;
      const ae = pe < end ? pe : end;
      if (as > ae) continue; // agreement not in effect at all this period
      const wdActive = workdays(as, ae);
      const wdPeriod = workdays(ps, pe);
      const full = as === ps && ae === pe;
      const allowMonth = i + 1;
      const fee = full ? r2(cfg.monthly_fee / 2) : r2(cfg.monthly_fee / 2 * wdActive / wdPeriod);
      const allowBase = allowMonth <= cfg.allowance_months ? cfg.monthly_allowance / 2 : 0;
      const allowance = full ? r2(allowBase) : r2(allowBase * wdActive / wdPeriod);
      const wire = (fee + allowance) > 0 ? r2(cfg.wire_fee) : 0;
      const invoiceDate = pe;
      rows.push({
        invoice_no: cfg.invoice_prefix + '-' + y + '-' + mm + suf,
        half, period_start: ps, period_end: pe, active_start: as, active_end: ae,
        wd_active: wdActive, wd_period: wdPeriod, prorated: !full,
        allowance_month: allowMonth <= cfg.allowance_months ? allowMonth : null,
        fee, allowance, wire_fee: wire, total: r2(fee + allowance + wire),
        invoice_date: invoiceDate, send_date: addDays(invoiceDate, -cfg.lead_days),
      });
      if (pe >= end) return rows;
    }
  }
  return rows;
}

function findInvoice(cfg, invoiceNo) {
  return buildSchedule(cfg, { months: 120 }).find((r) => r.invoice_no === invoiceNo) || null;
}

// ── PDF (matches the approved Excel/PDF invoice layout) ───────────────────────
async function renderInvoicePdf(cfg, inv) {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
  const pdf = await PDFDocument.create();
  pdf.setTitle('Invoice ' + inv.invoice_no);
  pdf.setAuthor(cfg.contractor_name);
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const W = 612, H = 792, L = 64, R = 548;
  const navy = rgb(0x1f / 255, 0x38 / 255, 0x64 / 255);
  const grey = rgb(0.6, 0.6, 0.6);
  const page = pdf.addPage([W, H]);
  const txt = (s, x, y, f, size, color) => page.drawText(String(s), { x, y, size, font: f, color: color || rgb(0, 0, 0) });
  const rt = (s, xR, y, f, size, color) => txt(s, xR - f.widthOfTextAtSize(String(s), size), y, f, size, color);
  const line = (x0, x1, y, th, color) => page.drawLine({ start: { x: x0, y }, end: { x: x1, y }, thickness: th || 0.5, color: color || grey });
  const money = (v) => '$' + (Number(v) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const wrap = (s, f, size, width) => {
    const words = String(s).split(' '); const out = []; let cur = '';
    for (const w of words) { const t = cur ? cur + ' ' + w : w; if (f.widthOfTextAtSize(t, size) > width && cur) { out.push(cur); cur = w; } else cur = t; }
    if (cur) out.push(cur); return out;
  };

  let y = H - 80;
  txt('INVOICE', L, y, bold, 24, navy);
  y -= 34;
  // Contractor block (left) + invoice meta (right)
  txt(cfg.contractor_name, L, y, bold, 11);
  const labX = 420, valR = R;
  rt('Invoice No.', labX, y, bold, 10); rt(inv.invoice_no, valR, y, reg, 10);
  let ly = y - 15;
  for (const a of String(cfg.contractor_address || '').split('\n').filter(Boolean)) { txt(a, L, ly, reg, 10); ly -= 14; }
  rt('Invoice Date', labX, y - 15, bold, 10); rt(mdy(inv.invoice_date), valR, y - 15, reg, 10);

  y = Math.min(ly, y - 30) - 16;
  txt('BILL TO', L, y, bold, 10, navy);
  txt('SERVICE PERIOD', 320, y, bold, 10, navy);
  txt(pretty(inv.active_start) + ' \u2013 ' + pretty(inv.active_end), 320, y - 15, reg, 10);
  let by = y - 15;
  for (const a of String(cfg.bill_to || '').split('\n').filter(Boolean)) { txt(a, L, by, reg, 10); by -= 14; }

  // Line items
  y = by - 18;
  const colPeriodR = 470;
  page.drawRectangle({ x: L, y: y - 4, width: R - L, height: 16, color: navy });
  txt('DESCRIPTION', L + 3, y, bold, 10, rgb(1, 1, 1));
  rt('PERIOD', colPeriodR, y, bold, 10, rgb(1, 1, 1));
  rt('AMOUNT', R - 3, y, bold, 10, rgb(1, 1, 1));
  y -= 22;
  const periodLbl = (inv.half === 1 ? '1st half ' : '2nd half ') + MON[Number(inv.period_start.slice(5, 7)) - 1] + ' ' + inv.period_start.slice(0, 4);
  const pr = inv.prorated ? ' (prorated: ' + inv.wd_active + ' of ' + inv.wd_period + ' working days)' : '';
  const items = [
    ['Accounting consulting services \u2014 semi-monthly installment of fixed monthly retainer' + pr, inv.fee],
  ];
  if (inv.allowance_month) items.push(['Equipment allowance \u2014 contractor-owned laptop lease (month ' + inv.allowance_month + ' of ' + cfg.allowance_months + ')' + pr, inv.allowance]);
  if (inv.wire_fee) items.push(['Wire transfer fee', inv.wire_fee]);
  for (const [desc, amt] of items) {
    const lines = wrap(desc, reg, 10, 270);
    lines.forEach((l, i) => txt(l, L + 3, y - i * 13, reg, 10));
    rt(periodLbl, colPeriodR, y, reg, 10);
    rt(money(amt), R - 3, y, reg, 10);
    y -= lines.length * 13 + 6;
    line(L, R, y + 2, 0.5);
    y -= 12;
  }
  y -= 14;
  rt('TOTAL DUE', colPeriodR, y, bold, 12);
  rt(money(inv.total), R - 3, y, bold, 12);
  line(colPeriodR + 2, R, y + 15, 1.2, rgb(0, 0, 0));
  line(colPeriodR + 2, R, y - 4, 0.6, rgb(0, 0, 0)); line(colPeriodR + 2, R, y - 6, 0.6, rgb(0, 0, 0));

  y -= 40;
  txt('NOTES', L, y, bold, 10, navy); y -= 15;
  const notes = [
    'Services provided as an independent contractor under the Independent Contractor Services Agreement between Banyan Residential, LLC and Muha K Kim, effective ' + pretty(cfg.effective_date).replace(/^(\w+) /, (m0, mo) => ({ Jan: 'January ', Feb: 'February ', Mar: 'March ', Apr: 'April ', May: 'May ', Jun: 'June ', Jul: 'July ', Aug: 'August ', Sep: 'September ', Oct: 'October ', Nov: 'November ', Dec: 'December ' }[mo] || m0)) + '.',
    'Services performed in the Republic of Korea.',
  ];
  for (const n of notes) { for (const l of wrap(n, reg, 8.5, R - L)) { txt(l, L, y, reg, 8.5); y -= 11; } y -= 6; }
  return Buffer.from(await pdf.save());
}

const fileNameFor = (inv) => 'Invoice_' + inv.invoice_no + '.pdf';

// ── email ─────────────────────────────────────────────────────────────────────
function senderAddress(fromSetting, displayName) {
  // Keep the verified sending address; replace only the display name.
  const m = String(fromSetting || '').match(/<([^>]+)>/);
  const addr = m ? m[1] : String(fromSetting || '').trim();
  return displayName + ' <' + addr + '>';
}

async function emailInvoice(ctx, cfg, inv, pdf) {
  const apiKey = ctx.getResendKey && ctx.getResendKey();
  if (!apiKey) return { ok: false, reason: 'RESEND_API_KEY is not configured on the server' };
  const cc = String(cfg.cc || '').split(/[,;\s]+/).filter((e) => /@/.test(e));
  const body = {
    from: senderAddress(ctx.getFromEmail && ctx.getFromEmail(), cfg.contractor_name),
    to: [cfg.send_to],
    subject: 'Invoice ' + inv.invoice_no + ' \u2014 ' + cfg.contractor_name + ' \u2014 $' + inv.total.toFixed(2),
    html: '<p>Please find attached invoice <b>' + inv.invoice_no + '</b> from ' + cfg.contractor_name + ' to Banyan Residential, LLC.</p>'
      + '<p>Service period: ' + pretty(inv.active_start) + ' \u2013 ' + pretty(inv.active_end) + '<br>Invoice date: ' + pretty(inv.invoice_date)
      + '<br>Amount: $' + inv.total.toLocaleString('en-US', { minimumFractionDigits: 2 }) + '</p>',
    attachments: [{ filename: fileNameFor(inv), content: pdf.toString('base64') }],
  };
  if (cc.length) body.cc = cc;
  try {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST', headers: { Authorization: 'Bearer ' + apiKey, 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
    if (!r.ok) return { ok: false, reason: 'Resend ' + r.status + ': ' + (await r.text()).slice(0, 300) };
    return { ok: true };
  } catch (e) { return { ok: false, reason: e.message }; }
}

function getSends(db) {
  ensureTables(db);
  const m = {};
  for (const r of db.prepare('SELECT * FROM muha_invoice_sends').all()) m[r.invoice_no] = r;
  return m;
}

// Claim an invoice for sending (atomic). Returns false if it was already sent,
// is in flight, or was skipped. A failed send can be retried.
function claim(db, invoiceNo, trigger, who) {
  const tx = db.transaction(() => {
    const cur = db.prepare('SELECT status FROM muha_invoice_sends WHERE invoice_no = ?').get(invoiceNo);
    if (cur && cur.status !== 'failed') return false;
    db.prepare("INSERT INTO muha_invoice_sends (invoice_no, status, send_trigger, sent_by, sent_at) VALUES (?, 'sending', ?, ?, datetime('now')) "
      + "ON CONFLICT(invoice_no) DO UPDATE SET status='sending', send_trigger=excluded.send_trigger, sent_by=excluded.sent_by, error=NULL, sent_at=datetime('now')")
      .run(invoiceNo, trigger, who);
    return true;
  });
  return tx();
}

async function sendOne(ctx, cfg, inv, trigger, who) {
  const { db } = ctx;
  if (!claim(db, inv.invoice_no, trigger, who)) return { ok: false, reason: 'Invoice ' + inv.invoice_no + ' was already sent or skipped' };
  try {
    const pdf = await renderInvoicePdf(cfg, inv);
    const res = await emailInvoice(ctx, cfg, inv, pdf);
    db.prepare("UPDATE muha_invoice_sends SET status = ?, amount = ?, sent_to = ?, error = ?, sent_at = datetime('now') WHERE invoice_no = ?")
      .run(res.ok ? 'sent' : 'failed', inv.total, cfg.send_to, res.ok ? null : res.reason, inv.invoice_no);
    console.log('[muha-invoice] ' + trigger + ' send ' + inv.invoice_no + ' -> ' + cfg.send_to + ': ' + (res.ok ? 'sent' : 'FAILED ' + res.reason));
    return res;
  } catch (e) {
    db.prepare("UPDATE muha_invoice_sends SET status = 'failed', error = ? WHERE invoice_no = ?").run(e.message, inv.invoice_no);
    return { ok: false, reason: e.message };
  }
}

// ── scheduler ─────────────────────────────────────────────────────────────────
// Runs every 30 minutes. On or after an invoice's send date (Pacific time, from
// 7am), sends it once. Never auto-sends an invoice whose invoice date is more
// than 7 days in the past (protects against a burst after a long outage or a
// back-dated config change — those are sent manually from the page instead).
async function tick(ctx) {
  const { db } = ctx;
  try {
    const cfg = getConfig(db);
    if (!cfg.auto_send) return;
    const now = pacificToday();
    if (now.hour < 7) return;
    const sends = getSends(db);
    const due = buildSchedule(cfg, { months: 60 }).filter((r) => r.send_date <= now.date
      && r.invoice_date >= addDays(now.date, -7)
      && (!sends[r.invoice_no] || sends[r.invoice_no].status === 'failed') && r.total > 0);
    for (const inv of due) {
      // a failed auto-send retries at most once per tick; manual resend is always available
      await sendOne(ctx, cfg, inv, 'auto', 'scheduler');
    }
  } catch (e) { console.error('[muha-invoice] scheduler error: ' + e.message); }
}

function startScheduler(ctx) {
  if (process.env.MUHA_INVOICE_SCHEDULER === 'off') { console.log('[muha-invoice] scheduler disabled by env'); return; }
  setTimeout(() => tick(ctx), 60 * 1000);
  setInterval(() => tick(ctx), 30 * 60 * 1000);
}

// ── routes ────────────────────────────────────────────────────────────────────
function registerMuhaInvoiceRoutes(app, ctx) {
  const { db, auth, requireRole } = ctx;
  ensureTables(db);
  const who = (req) => (req.user && (req.user.email || req.user.name)) || 'system';

  app.get('/api/admin/muha-invoice', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    try {
      const cfg = getConfig(db);
      const sends = getSends(db);
      const today = pacificToday().date;
      const rows = buildSchedule(cfg, { months: 30 }).map((r) => {
        const s = sends[r.invoice_no];
        let status = s ? s.status : (r.send_date <= today ? 'due' : 'scheduled');
        return Object.assign({}, r, { status, sent_at: s && s.sent_at, sent_to: s && s.sent_to, trigger: s && s.send_trigger, error: s && s.error, sent_by: s && s.sent_by });
      });
      res.json({ config: cfg, today, email_configured: !!(ctx.getResendKey && ctx.getResendKey()), rows });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  app.put('/api/admin/muha-invoice/config', auth, requireRole('Admin'), (req, res) => {
    try { res.json({ config: saveConfig(db, req.body || {}, who(req)) }); }
    catch (e) { res.status(400).json({ error: e.message }); }
  });

  app.get('/api/admin/muha-invoice/:no/pdf', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const cfg = getConfig(db);
      const inv = findInvoice(cfg, req.params.no);
      if (!inv) return res.status(404).json({ error: 'No invoice ' + req.params.no });
      const buf = await renderInvoicePdf(cfg, inv);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fileNameFor(inv) + '"');
      res.send(buf);
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Send now (manual). Re-sending an already-sent invoice requires force=true.
  app.post('/api/admin/muha-invoice/:no/send', auth, requireRole('Admin'), async (req, res) => {
    try {
      const cfg = getConfig(db);
      const inv = findInvoice(cfg, req.params.no);
      if (!inv) return res.status(404).json({ error: 'No invoice ' + req.params.no });
      if (req.body && req.body.force) db.prepare("UPDATE muha_invoice_sends SET status='failed', error='reset for manual resend' WHERE invoice_no = ?").run(inv.invoice_no);
      const out = await sendOne(ctx, cfg, inv, 'manual', who(req));
      if (!out.ok) return res.status(400).json({ error: out.reason });
      res.json({ ok: true, invoice_no: inv.invoice_no });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Skip (don't auto-send) or un-skip an invoice.
  app.post('/api/admin/muha-invoice/:no/skip', auth, requireRole('Admin'), (req, res) => {
    try {
      const skip = !(req.body && req.body.skip === false);
      if (skip) {
        const cur = db.prepare('SELECT status FROM muha_invoice_sends WHERE invoice_no = ?').get(req.params.no);
        if (cur && (cur.status === 'sent' || cur.status === 'sending')) return res.status(400).json({ error: 'Already sent' });
        db.prepare("INSERT INTO muha_invoice_sends (invoice_no, status, sent_by, sent_at) VALUES (?, 'skipped', ?, datetime('now')) "
          + "ON CONFLICT(invoice_no) DO UPDATE SET status='skipped', sent_by=excluded.sent_by, sent_at=datetime('now')").run(req.params.no, who(req));
      } else {
        db.prepare("DELETE FROM muha_invoice_sends WHERE invoice_no = ? AND status = 'skipped'").run(req.params.no);
      }
      res.json({ ok: true });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  startScheduler(ctx);
}

module.exports = { DEFAULTS, buildSchedule, renderInvoicePdf, workdays, registerMuhaInvoiceRoutes, _tick: tick };
