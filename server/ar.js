// ═══════════════════════════════════════════════════════════════════════════
// Accounts Receivable — customer invoicing engine
//
// Phase 1 (customers) lives in index.js. This module adds everything else:
//   • per-entity AR settings (remit-to block, default A/R account, prefix)
//   • recurring invoice templates + generation
//   • one-off invoices, each posting an accrual JE (Dr A/R, Cr Revenue)
//   • invoice PDF (pdf-lib), stored under entity_files/Invoices/<year>
//   • review-then-send email via Resend, with the PDF attached
//   • cash receipts (Dr Bank, Cr A/R) with partial-payment support
//   • A/R aging that ties to the GL A/R account balance (recon_diff)
//
// Design notes:
//   - Accrual model: the JE posts when the invoice is created, not when sent.
//   - One A/R account per invoice, defaulting to the entity's "Accounts
//     Receivable" account discovered from the COA (12000 on every live entity
//     today; the 11000 default in the original schema is not used).
//   - Voiding a DRAFT invoice deletes its JE (nothing has left the building).
//     Voiding a SENT invoice posts a reversing JE, since the customer holds the
//     original. Reclasses elsewhere in CloudLedger prefer in-place edits, but
//     an issued invoice is an external document: it gets reversed, not edited.
//   - Aging is invoice-driven but reconciled against the GL A/R balance, so any
//     mismatch surfaces as recon_diff instead of silently disagreeing with the TB.
// ═══════════════════════════════════════════════════════════════════════════
'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { PDFDocument } = require('pdf-lib');

const MONEY = (n) => {
  const v = Number(n) || 0;
  return (v < 0 ? '-' : '') + '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const ymd = (s) => { const p = String(s).split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); };
const addDays = (dateStr, days) => {
  const dt = new Date(ymd(dateStr));
  dt.setUTCDate(dt.getUTCDate() + (Number(days) || 0));
  return dt.toISOString().slice(0, 10);
};
const daysBetween = (from, to) => {
  if (!isDate(from) || !isDate(to)) return 0;
  return Math.round((ymd(to) - ymd(from)) / 86400000);
};
const todayStr = () => new Date().toISOString().slice(0, 10);

// ─── schema (additive; the core AR tables are created in index.js) ──────────
function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS ar_settings (
      entity_id INTEGER PRIMARY KEY REFERENCES entities(id) ON DELETE CASCADE,
      remit_to TEXT,
      bill_from TEXT,
      default_ar_account TEXT,
      invoice_prefix TEXT,
      footer_note TEXT,
      reply_to TEXT,
      updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS ar_receipts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      invoice_id INTEGER NOT NULL REFERENCES ar_invoices(id) ON DELETE CASCADE,
      entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
      date TEXT NOT NULL,
      amount REAL NOT NULL,
      bank_account_code TEXT NOT NULL,
      memo TEXT,
      je_id INTEGER REFERENCES journal_entries(id) ON DELETE SET NULL,
      created_by TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_ar_receipts_inv ON ar_receipts(invoice_id);
    CREATE INDEX IF NOT EXISTS idx_ar_receipts_ent ON ar_receipts(entity_id, date);
  `);
  // ar_invoices gained void bookkeeping after the original schema shipped.
  const cols = db.prepare('PRAGMA table_info(ar_invoices)').all().map(c => c.name);
  if (!cols.includes('void_je_id')) db.exec('ALTER TABLE ar_invoices ADD COLUMN void_je_id INTEGER');
  if (!cols.includes('voided_at')) db.exec('ALTER TABLE ar_invoices ADD COLUMN voided_at TEXT');
  // origin: 'native' = a CloudLedger invoice with its own accrual JE; 'opening'
  // = a legacy item imported from the prior system's A/R aging detail, GL-backed
  // (no accrual JE — the balance already sits on the control account) and aged
  // by its own invoice/due dates so the legacy A/R reads as a real subledger.
  if (!cols.includes('origin')) db.exec("ALTER TABLE ar_invoices ADD COLUMN origin TEXT DEFAULT 'native'");
  // aging_date: the date the aging clock runs from. NULL for native invoices
  // (they age by due date, unchanged). Opening items set it to the legacy GL
  // posting date so the report matches the prior system's posting-date aging.
  if (!cols.includes('aging_date')) db.exec('ALTER TABLE ar_invoices ADD COLUMN aging_date TEXT');
  console.log('[db] AR invoicing schema ready');
}

// ─── helpers ────────────────────────────────────────────────────────────────
function getSettings(db, eid) {
  const row = db.prepare('SELECT * FROM ar_settings WHERE entity_id = ?').get(eid);
  return row || { entity_id: +eid, remit_to: null, bill_from: null, default_ar_account: null, invoice_prefix: null, footer_note: null, reply_to: null };
}

// Discover the entity's A/R control account from its chart of accounts. Every
// live entity uses 12000 "Accounts Receivable"; "- Other" and note/interest
// receivable variants are skipped so a secondary receivable never wins.
function defaultArAccount(db, eid) {
  const s = getSettings(db, eid);
  if (s.default_ar_account) return s.default_ar_account;
  const rows = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ? ORDER BY code').all(eid);
  const clean = rows.find(a => /accounts?\s*receivable/i.test(a.name || '') && !/other|note|interest/i.test(a.name || ''));
  if (clean) return clean.code;
  const any = rows.find(a => /receivab/i.test(a.name || ''));
  return any ? any.code : null;
}

function nextInvoiceNum(db, eid, dateStr) {
  const prefix = (getSettings(db, eid).invoice_prefix || 'INV').replace(/[^A-Za-z0-9_-]/g, '') || 'INV';
  const yr = String(dateStr).slice(0, 4);
  const rows = db.prepare('SELECT invoice_num FROM ar_invoices WHERE entity_id = ? AND invoice_num LIKE ?').all(eid, prefix + '-' + yr + '-%');
  let max = 0;
  for (const r of rows) { const m = /-(\d+)$/.exec(r.invoice_num || ''); if (m) max = Math.max(max, +m[1]); }
  return prefix + '-' + yr + '-' + String(max + 1).padStart(4, '0');
}

// Post a balanced journal entry. Mirrors the /entries endpoint's insert shape.
function postJE(db, eid, date, memo, lines, who) {
  let dr = 0, cr = 0;
  for (const l of lines) { dr += Number(l.debit || 0); cr += Number(l.credit || 0); }
  if (Math.abs(dr - cr) > 0.005) throw new Error('Unbalanced JE: Dr ' + dr.toFixed(2) + ' vs Cr ' + cr.toFixed(2));
  const num = (db.prepare('SELECT COALESCE(MAX(entry_num),0) AS m FROM journal_entries WHERE entity_id = ?').get(eid).m || 0) + 1;
  const r = db.prepare('INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by) VALUES (?,?,?,?,?)')
    .run(eid, num, date, memo, who || 'ar-module');
  const ins = db.prepare('INSERT INTO journal_lines (entry_id, account_code, debit, credit, description, project_id, class_id, location_id) VALUES (?,?,?,?,?,?,?,?)');
  for (const l of lines) {
    ins.run(r.lastInsertRowid, String(l.account_code), r2(l.debit || 0), r2(l.credit || 0),
      l.description || '', l.project_id || null, l.class_id || null, l.location_id || null);
  }
  return { id: r.lastInsertRowid, entry_num: num };
}

function deleteJE(db, jeId) {
  if (!jeId) return;
  db.prepare('DELETE FROM journal_lines WHERE entry_id = ?').run(jeId);
  db.prepare('DELETE FROM journal_entries WHERE id = ?').run(jeId);
}

function invoiceWithLines(db, eid, id) {
  const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(id, eid);
  if (!inv) return null;
  inv.lines = db.prepare('SELECT * FROM ar_invoice_lines WHERE invoice_id = ? ORDER BY sort, id').all(id);
  inv.receipts = db.prepare('SELECT * FROM ar_receipts WHERE invoice_id = ? ORDER BY date, id').all(id);
  inv.paid_amount = r2(inv.receipts.reduce((s, x) => s + Number(x.amount || 0), 0));
  inv.open_amount = r2(Number(inv.total || 0) - inv.paid_amount);
  return inv;
}

// Normalize + validate an incoming line array against the entity's COA.
function normalizeLines(db, eid, raw) {
  if (!Array.isArray(raw) || raw.length === 0) throw new Error('At least one invoice line is required');
  const codes = new Set(db.prepare('SELECT code FROM accounts WHERE entity_id = ?').all(eid).map(a => String(a.code)));
  const out = [];
  raw.forEach((l, i) => {
    const desc = String(l.description || '').trim();
    if (!desc) throw new Error('Line ' + (i + 1) + ': description required');
    const code = String(l.revenue_account_code || '').trim();
    if (!code) throw new Error('Line ' + (i + 1) + ': revenue account required');
    if (!codes.has(code)) throw new Error('Line ' + (i + 1) + ': account ' + code + ' is not in this entity\'s chart of accounts');
    const qty = Number(l.qty);
    const rate = Number(l.rate);
    if (!Number.isFinite(qty) || !Number.isFinite(rate)) throw new Error('Line ' + (i + 1) + ': qty and rate must be numbers');
    const amount = r2(qty * rate);
    if (Math.abs(amount) < 0.005) throw new Error('Line ' + (i + 1) + ': amount is zero');
    out.push({ description: desc, qty, rate, amount, revenue_account_code: code, class_id: l.class_id || null, location_id: l.location_id || null, sort: i });
  });
  return out;
}

// Create an invoice plus its accrual JE inside one transaction.
function createInvoice(db, eid, body, who) {
  const cust = db.prepare('SELECT * FROM ar_customers WHERE id = ? AND entity_id = ?').get(body.customer_id, eid);
  if (!cust) throw new Error('Customer not found for this entity');
  const invoice_date = isDate(body.invoice_date) ? body.invoice_date : todayStr();
  const terms = Number.isFinite(+body.terms_days) ? +body.terms_days : (cust.terms_days == null ? 30 : cust.terms_days);
  const due_date = isDate(body.due_date) ? body.due_date : addDays(invoice_date, terms);
  const arCode = String(body.ar_account_code || defaultArAccount(db, eid) || '').trim();
  if (!arCode) throw new Error('No Accounts Receivable account found in this entity\'s chart of accounts. Add one (e.g. 12000) or set a default in A/R Settings.');
  const lines = normalizeLines(db, eid, body.lines);
  const total = r2(lines.reduce((s, l) => s + l.amount, 0));
  const memo = String(body.memo || '').trim() || null;

  return db.transaction(() => {
    const num = nextInvoiceNum(db, eid, invoice_date);
    const ins = db.prepare('INSERT INTO ar_invoices '
      + '(entity_id, customer_id, template_id, invoice_num, invoice_date, due_date, '
      + 'customer_name, customer_email, customer_address, memo, subtotal, total, '
      + "ar_account_code, status, created_by) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'draft',?)")
      .run(eid, cust.id, body.template_id || null, num, invoice_date, due_date,
        cust.name, cust.email || null, cust.address || null, memo, total, total, arCode, who || null);
    const invId = ins.lastInsertRowid;
    const insL = db.prepare('INSERT INTO ar_invoice_lines (invoice_id, description, qty, rate, amount, revenue_account_code, class_id, location_id, sort) VALUES (?,?,?,?,?,?,?,?,?)');
    for (const l of lines) insL.run(invId, l.description, l.qty, l.rate, l.amount, l.revenue_account_code, l.class_id, l.location_id, l.sort);

    const jeLines = [{ account_code: arCode, debit: total, credit: 0, description: 'Invoice ' + num + ' - ' + cust.name }];
    for (const l of lines) jeLines.push({ account_code: l.revenue_account_code, debit: 0, credit: l.amount, description: l.description, class_id: l.class_id, location_id: l.location_id });
    const je = postJE(db, eid, invoice_date, 'AR Invoice ' + num + ' - ' + cust.name + (memo ? ' - ' + memo : ''), jeLines, who);
    db.prepare('UPDATE ar_invoices SET je_id = ? WHERE id = ?').run(je.id, invId);
    return invId;
  })();
}

function advanceNextRun(current, frequency, dayOfMonth) {
  const base = isDate(current) ? current : todayStr();
  const parts = base.split('-').map(Number);
  const step = frequency === 'quarterly' ? 3 : frequency === 'annual' ? 12 : 1;
  const target = new Date(Date.UTC(parts[0], (parts[1] - 1) + step, 1));
  const yy = target.getUTCFullYear(), mm = target.getUTCMonth() + 1;
  const last = new Date(Date.UTC(yy, mm, 0)).getUTCDate();
  const day = Math.min(Math.max(Number(dayOfMonth) || 1, 1), last);
  return yy + '-' + String(mm).padStart(2, '0') + '-' + String(day).padStart(2, '0');
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ─── invoice PDF ────────────────────────────────────────────────────────────
async function buildInvoicePdf({ entityName, settings, invoice }) {
  const { rgb, StandardFonts } = require('pdf-lib');
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const NAVY = rgb(0.13, 0.21, 0.36);
  const GREY = rgb(0.45, 0.45, 0.45);
  const RULE = rgb(0.78, 0.78, 0.78);
  const FAINT = rgb(0.65, 0.65, 0.65);
  const L = 56, R = 556;
  let page = doc.addPage([612, 792]);
  let y = 736;
  const draw = (t, x, yy, f, size, color) =>
    page.drawText(String(t == null ? '' : t), { x, y: yy, size: size || 10, font: f || font, color: color || rgb(0.1, 0.1, 0.1) });
  const drawR = (t, xRight, yy, f, size, color) => {
    const ff = f || font, sz = size || 10, s = String(t == null ? '' : t);
    draw(s, xRight - ff.widthOfTextAtSize(s, sz), yy, ff, sz, color);
  };
  const rule = (yy, thickness, color) => page.drawLine({ start: { x: L, y: yy }, end: { x: R, y: yy }, thickness: thickness || 0.75, color: color || RULE });
  const wrap = (text, f, size, maxW) => {
    const words = String(text || '').split(/\s+/);
    const out = []; let cur = '';
    for (const w of words) {
      const t = cur ? cur + ' ' + w : w;
      if (f.widthOfTextAtSize(t, size) > maxW && cur) { out.push(cur); cur = w; } else cur = t;
    }
    if (cur) out.push(cur);
    return out.length ? out : [''];
  };

  // Header
  draw(entityName, L, y, bold, 16, NAVY);
  drawR('INVOICE', R, y + 1, bold, 22, NAVY);
  y -= 12;
  rule(y, 1.5, NAVY);
  y -= 18;

  // Bill-from block (left) and invoice meta (right)
  const fromLines = String(settings.bill_from || '').split(/\r?\n/).filter(Boolean);
  let metaY = y, fromY = y;
  const meta = [['Invoice #', invoice.invoice_num], ['Invoice date', invoice.invoice_date], ['Due date', invoice.due_date || '-']];
  for (const kv of meta) { draw(kv[0], 388, metaY, bold, 9, GREY); drawR(kv[1], R, metaY, font, 10); metaY -= 14; }
  for (const line of fromLines.slice(0, 6)) { draw(line, L, fromY, font, 9, GREY); fromY -= 12; }
  y = Math.min(metaY, fromY) - 12;

  // Bill to
  draw('BILL TO', L, y, bold, 9, GREY); y -= 14;
  draw(invoice.customer_name || '', L, y, bold, 11); y -= 13;
  for (const line of String(invoice.customer_address || '').split(/\r?\n/).filter(Boolean).slice(0, 5)) { draw(line, L, y, font, 9); y -= 11; }
  if (invoice.customer_email) { draw(invoice.customer_email, L, y, font, 9, GREY); y -= 11; }
  y -= 10;
  if (invoice.memo) { draw('Re: ' + invoice.memo, L, y, font, 10); y -= 18; }

  // Line table
  const COL_D = L, COL_Q = 372, COL_R = 452, COL_A = R;
  rule(y + 4, 0.75, RULE);
  y -= 10;
  draw('DESCRIPTION', COL_D, y, bold, 9, GREY);
  drawR('QTY', COL_Q, y, bold, 9, GREY);
  drawR('RATE', COL_R, y, bold, 9, GREY);
  drawR('AMOUNT', COL_A, y, bold, 9, GREY);
  y -= 6;
  rule(y, 0.75, RULE);
  y -= 16;

  for (const l of invoice.lines) {
    const wrapped = wrap(l.description, font, 10, 300);
    if (y - wrapped.length * 12 < 150) { page = doc.addPage([612, 792]); y = 736; }
    wrapped.forEach((w, i) => draw(w, COL_D, y - i * 12, font, 10));
    drawR(Number(l.qty) === 1 ? '1' : String(+Number(l.qty).toFixed(4)), COL_Q, y, font, 10);
    drawR(MONEY(l.rate), COL_R, y, font, 10);
    drawR(MONEY(l.amount), COL_A, y, font, 10);
    y -= wrapped.length * 12 + 6;
  }

  y -= 4;
  rule(y, 0.75, RULE);
  y -= 18;
  draw('Total Due', 372, y, bold, 12, NAVY);
  drawR(MONEY(invoice.total), R, y, bold, 12, NAVY);
  y -= 8;
  rule(y, 1.5, NAVY);
  y -= 22;

  if (Number(invoice.paid_amount || 0) > 0.005) {
    drawR('Payments received: ' + MONEY(-invoice.paid_amount), R, y, font, 10, GREY); y -= 14;
    drawR('Balance due: ' + MONEY(invoice.open_amount), R, y, bold, 11); y -= 18;
  }

  // Footer
  const footY = 96;
  if (settings.remit_to) {
    draw('REMIT TO', L, footY + 34, bold, 9, GREY);
    String(settings.remit_to).split(/\r?\n/).filter(Boolean).slice(0, 3)
      .forEach((line, i) => draw(line, L, footY + 22 - i * 11, font, 9));
  }
  if (settings.footer_note) {
    wrap(settings.footer_note, font, 8, R - L).slice(0, 3)
      .forEach((line, i) => draw(line, L, footY - 34 - i * 10, font, 8, GREY));
  }
  draw('Generated by CloudLedger', L, 52, font, 8, FAINT);
  drawR(invoice.invoice_num, R, 52, font, 8, FAINT);

  return Buffer.from(await doc.save());
}

// Persist a generated PDF as an entity_files row (same shape as a manual upload)
// so it shows up in the Workpapers tree under Invoices/<year>.
function savePdf(db, workpapersDir, eid, folderPath, originalName, buffer, who) {
  const entityDir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(entityDir, { recursive: true });
  const insFolder = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)');
  let acc = '';
  for (const part of folderPath.split('/').filter(Boolean)) {
    acc = acc ? acc + '/' + part : part;
    try { insFolder.run(eid, acc, who || 'ar-module'); } catch (_) { /* non-fatal */ }
  }
  // Replace any prior copy of the same invoice PDF rather than piling up dupes.
  const dupes = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name=?').all(eid, folderPath, originalName);
  for (const d of dupes) {
    try { fs.unlinkSync(path.join(entityDir, d.stored_filename)); } catch (_) { /* already gone */ }
    db.prepare('DELETE FROM entity_files WHERE id=?').run(d.id);
  }
  const stored = crypto.randomBytes(16).toString('hex') + '.pdf';
  fs.writeFileSync(path.join(entityDir, stored), buffer);
  const r = db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)')
    .run(eid, folderPath, stored, originalName, buffer.length, 'application/pdf', who || 'ar-module');
  return r.lastInsertRowid;
}

async function sendInvoiceEmail({ apiKey, from, replyTo, to, cc, subject, html, filename, pdf }) {
  if (!apiKey) return { ok: false, reason: 'RESEND_API_KEY is not configured on the server' };
  const body = { from: from, to: [to], subject: subject, html: html, attachments: [{ filename: filename, content: pdf.toString('base64') }] };
  if (replyTo) body.reply_to = replyTo;
  if (cc) body.cc = Array.isArray(cc) ? cc : [cc];
  try {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); console.error('[ar] Resend error ' + r.status + ': ' + t); return { ok: false, status: r.status, reason: t }; }
    return { ok: true };
  } catch (e) {
    console.error('[ar] email send failed: ' + e.message);
    return { ok: false, reason: e.message };
  }
}

function invoiceEmailHtml({ entityName, invoice, settings }) {
  const rows = invoice.lines.map(l =>
    '<tr><td style="padding:6px 0;color:#334155;font-size:13px">' + escapeHtml(l.description) + '</td>'
    + '<td style="padding:6px 0;text-align:right;color:#334155;font-size:13px">' + MONEY(l.amount) + '</td></tr>').join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;padding:24px">'
    + '<h2 style="color:#1e3a5f;margin:0 0 4px">Invoice ' + escapeHtml(invoice.invoice_num) + '</h2>'
    + '<p style="color:#64748b;font-size:13px;margin:0 0 18px">from ' + escapeHtml(entityName) + '</p>'
    + '<p style="color:#334155;font-size:14px;line-height:1.5">Hello' + (invoice.customer_name ? ' ' + escapeHtml(invoice.customer_name) : '') + ',</p>'
    + '<p style="color:#334155;font-size:14px;line-height:1.5">Please find invoice <strong>' + escapeHtml(invoice.invoice_num)
    + '</strong> attached, dated ' + escapeHtml(invoice.invoice_date) + (invoice.due_date ? ' and due ' + escapeHtml(invoice.due_date) : '') + '.</p>'
    + '<table style="width:100%;border-collapse:collapse;margin:16px 0">' + rows
    + '<tr><td style="padding:10px 0;border-top:1px solid #cbd5e1;font-weight:bold;color:#1e3a5f;font-size:14px">Total Due</td>'
    + '<td style="padding:10px 0;border-top:1px solid #cbd5e1;text-align:right;font-weight:bold;color:#1e3a5f;font-size:14px">' + MONEY(invoice.total) + '</td></tr></table>'
    + (settings.remit_to ? '<p style="color:#64748b;font-size:12px;line-height:1.5"><strong>Remit to:</strong><br>' + escapeHtml(settings.remit_to).replace(/\n/g, '<br>') + '</p>' : '')
    + (settings.footer_note ? '<p style="color:#94a3b8;font-size:11px;line-height:1.5">' + escapeHtml(settings.footer_note) + '</p>' : '')
    + '<p style="color:#94a3b8;font-size:11px;margin-top:20px">Questions about this invoice? Reply to this email.</p>'
    + '</div>';
}

// ─── A/R aging (invoice-driven, reconciled to the GL A/R balance) ───────────
// Two-source model, mirroring the A/P aging report:
//   • CloudLedger invoices age normally into current / 1-30 / 31-60 / 61-90 / 90+.
//   • Everything else on the A/R control account — legacy Intacct imports and any
//     manual JE — is NOT aged. It is summed into a single "GL" column and listed
//     as un-aged GL entries, to be cleared by journal entry over time.
// The split is by source, not by date: a JE is "CloudLedger" only if it is the
// accrual/void JE of a non-void invoice (or a receipt JE). This survives someone
// backdating a real invoice and needs no per-entity cutover date. By construction
// aged buckets + GL column == GL control balance, so recon_diff is ~0.
function buildAging(db, eid, asOf) {
  const invoices = db.prepare("SELECT * FROM ar_invoices WHERE entity_id = ? AND status != 'void' AND invoice_date <= ? ORDER BY customer_name, invoice_date, invoice_num").all(eid, asOf);
  const recByInv = new Map();
  for (const r of db.prepare('SELECT invoice_id, SUM(amount) AS paid FROM ar_receipts WHERE entity_id = ? AND date <= ? GROUP BY invoice_id').all(eid, asOf)) {
    recByInv.set(r.invoice_id, Number(r.paid || 0));
  }
  const BUCKETS = ['current', 'd1_30', 'd31_60', 'd61_90', 'd90_plus'];
  const zero = () => ({ current: 0, d1_30: 0, d31_60: 0, d61_90: 0, d90_plus: 0, gl: 0, total: 0 });
  const byCustomer = new Map();
  const detail = [];
  for (const inv of invoices) {
    const open = r2(Number(inv.total || 0) - (recByInv.get(inv.id) || 0));
    if (Math.abs(open) < 0.005) continue;
    const agingRef = inv.aging_date || inv.due_date;
    const past = agingRef ? daysBetween(agingRef, asOf) : 0;
    const bucket = past <= 0 ? 'current' : past <= 30 ? 'd1_30' : past <= 60 ? 'd31_60' : past <= 90 ? 'd61_90' : 'd90_plus';
    const key = inv.customer_name || '(no customer)';
    if (!byCustomer.has(key)) byCustomer.set(key, Object.assign({ customer: key }, zero()));
    const row = byCustomer.get(key);
    row[bucket] = r2(row[bucket] + open);
    row.total = r2(row.total + open);
    detail.push({
      invoice_id: inv.id, invoice_num: inv.invoice_num, customer: key, invoice_date: inv.invoice_date,
      due_date: inv.due_date, status: inv.status, total: r2(inv.total), paid: r2(recByInv.get(inv.id) || 0),
      open: open, days_past_due: Math.max(past, 0), bucket: bucket, je_id: inv.je_id,
    });
  }
  const rows = Array.from(byCustomer.values()).sort((a, b) => b.total - a.total);
  const totals = zero();
  for (const r of rows) { for (const b of BUCKETS) totals[b] = r2(totals[b] + r[b]); totals.total = r2(totals.total + r.total); }

  // GL tie-out: net debit balance of every A/R account these invoices touch
  // (plus the entity default), so the report always ties to the trial balance.
  const codes = new Set(invoices.map(i => i.ar_account_code).filter(Boolean));
  const def = defaultArAccount(db, eid);
  if (def) codes.add(def);
  let glBalance = 0;
  const codeList = Array.from(codes);
  if (codeList.length) {
    const ph = codeList.map(() => '?').join(',');
    const row = db.prepare('SELECT COALESCE(SUM(jl.debit - jl.credit), 0) AS bal FROM journal_lines jl '
      + 'JOIN journal_entries je ON je.id = jl.entry_id '
      + 'WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code IN (' + ph + ')').get(eid, asOf, ...codeList);
    glBalance = r2(row.bal || 0);
  }

  // ── GL column: A/R control activity NOT produced by this module. Every JE id
  // the AR module owns (invoice accruals, void reversals, receipt JEs) is
  // excluded; the remaining net movement on the control account(s) is the legacy
  // / manually-booked balance that we surface un-aged.
  const glRows = [];
  let glTotal = 0;
  const bucketsTotal = totals.total;
  // If a legacy opening subledger has been imported, the control balance is now
  // itemized into aged buckets, so the GL column becomes the residual not yet on
  // the subledger: residual = control balance - buckets.
  //
  // That residual is only meaningful in ONE direction. A POSITIVE residual is
  // legitimate: control-account money the subledger has not detailed yet. A
  // NEGATIVE residual means the subledger claims more open A/R than the control
  // account holds, which is never legitimate — the wrong detail was imported
  // (wrong file or wrong entity), or items are double-counted. Netting that into
  // the GL column would make buckets + residual == control balance by
  // construction, so recon_diff would be identically zero and the tie-out would
  // report a pass regardless of what the subledger contained. So a negative
  // residual is deliberately NOT absorbed: glTotal stays 0, the report totals
  // what the subledger actually says, and the overclaim surfaces in recon_diff.
  // A POSITIVE residual IS still absorbed, so recon_diff stays zero there by
  // construction: a green tie-out means "no overclaim", NOT "fully itemized". It
  // is therefore reported separately as opening_residual, and the report labels a
  // partially-itemized subledger amber rather than green.
  let openingResidual = 0, openingOverclaim = 0;
  const hasOpening = db.prepare("SELECT COUNT(*) AS c FROM ar_invoices WHERE entity_id = ? AND origin = 'opening'").get(eid).c > 0;
  if (hasOpening) {
    const residual = r2(glBalance - bucketsTotal);
    if (residual >= 0.005) {
      glTotal = residual;
      openingResidual = residual;
      glRows.push({ entry_id: null, entry_num: null, date: asOf, memo: 'Un-itemized GL balance (not yet on the subledger)', amount: residual });
    } else if (residual <= -0.005) {
      openingOverclaim = r2(-residual);
    }
  } else if (codeList.length) {
    // Own only the JEs of NON-void invoices, plus receipt JEs. A void invoice's
    // accrual and its reversal are deliberately left un-owned so both fall into
    // the GL column and net to zero there — otherwise the accrual would be hidden
    // (owned) while its reversal showed, throwing the tie-out off by the invoice.
    const ownedRows = db.prepare(
      "SELECT je_id AS id FROM ar_invoices WHERE entity_id = ? AND status != 'void' AND je_id IS NOT NULL "
      + "UNION SELECT je_id FROM ar_receipts WHERE entity_id = ? AND je_id IS NOT NULL").all(eid, eid);
    const owned = new Set(ownedRows.map(r => r.id));
    const ph = codeList.map(() => '?').join(',');
    // Roll un-aged control activity up to the JE, so each row reads like a GL
    // line the way the A/P report presents its imported entries.
    const lines = db.prepare(
      'SELECT je.id AS entry_id, je.entry_num, je.date, je.memo, '
      + 'SUM(jl.debit - jl.credit) AS amount '
      + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
      + 'WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code IN (' + ph + ') '
      + 'GROUP BY je.id ORDER BY je.date, je.entry_num').all(eid, asOf, ...codeList);
    for (const l of lines) {
      if (owned.has(l.entry_id)) continue;
      const amt = r2(l.amount || 0);
      if (Math.abs(amt) < 0.005) continue;
      glRows.push({ entry_id: l.entry_id, entry_num: l.entry_num, date: l.date, memo: l.memo || 'GL detail import', amount: amt });
      glTotal = r2(glTotal + amt);
    }
  }
  totals.gl = glTotal;
  totals.total = r2(bucketsTotal + glTotal);

  return {
    as_of: asOf, ar_accounts: codeList, ar_account: def || (codeList[0] || null),
    rows: rows, totals: totals, detail: detail,
    gl_rows: glRows, gl_total: glTotal,
    gl_ar_balance: glBalance, recon_diff: r2(glBalance - totals.total),
    opening_residual: openingResidual, opening_overclaim: openingOverclaim,
  };
}

// ─── legacy opening-balance subledger + external cash application ────────────
// Import a prior-system A/R aging detail as GL-backed "opening" open items: one
// ar_invoice per document, aging by its own dates, with NO accrual JE (the
// balance already sits on the control account from the legacy GL import).
// Idempotent per entity: existing opening items with no receipts are replaced.
function importOpeningItems(db, eid, items, opts = {}) {
  const arCode = defaultArAccount(db, eid) || '12000';

  // ── guard: the detail may not claim more open A/R than the control account
  // actually holds. A grand total above the GL balance means the wrong entity is
  // selected or the wrong file was picked — the subledger would itemize money
  // that is not on the control account. Under-claiming is legitimate (a partial
  // itemization) and surfaces as the un-itemized residual in the aging report.
  // Checked BEFORE any existing opening items are deleted, so a rejected import
  // leaves the current subledger intact.
  const itemsTotal = r2(items.reduce((s, it) => s + r2(it.amount), 0));
  // Never skip the guard just because no as_of was supplied — a caller that omits
  // it is exactly the caller least likely to have checked. But the basis matters:
  //   • With an as_of, compare to the control balance AT that date. Exact.
  //   • Without one, compare to the PEAK balance the control account has ever
  //     held. Two wrong bases were tried first and both produce false rejections:
  //     max(document date in file) reads the account before the legacy opening JE
  //     posts (document dates precede the posting date, which is why opening items
  //     age by posting date), and the current balance is lower than the file
  //     whenever cash has been applied since cutover — which broke re-importing
  //     the original aging. The peak is immune to both and still catches a
  //     wrong-entity file, which overshoots at every date.
  let guardBasis;
  if (isDate(opts.as_of)) {
    guardBasis = r2((db.prepare('SELECT COALESCE(SUM(jl.debit - jl.credit), 0) AS bal FROM journal_lines jl '
      + 'JOIN journal_entries je ON je.id = jl.entry_id '
      + 'WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code = ?').get(eid, opts.as_of, arCode) || {}).bal || 0);
  } else {
    const daily = db.prepare('SELECT je.date AS d, SUM(jl.debit - jl.credit) AS amt FROM journal_lines jl '
      + 'JOIN journal_entries je ON je.id = jl.entry_id '
      + 'WHERE je.entity_id = ? AND jl.account_code = ? GROUP BY je.date ORDER BY je.date').all(eid, arCode);
    let run = 0; guardBasis = 0;
    for (const row of daily) { run = r2(run + Number(row.amt || 0)); if (run > guardBasis) guardBasis = run; }
  }
  {
    const bal = guardBasis;
    const over = r2(itemsTotal - bal);
    if (over > 0.01 && !opts.allow_over_gl) {
      const ent = (db.prepare('SELECT name FROM entities WHERE id = ?').get(eid) || {}).name || ('entity ' + eid);
      throw new Error('This detail totals ' + itemsTotal.toFixed(2) + ', but account ' + arCode + ' on ' + ent
        + ' holds only ' + bal.toFixed(2) + (isDate(opts.as_of) ? ' as of ' + opts.as_of : ' at its highest')
        + ' — the detail is ' + over.toFixed(2)
        + ' MORE than the GL. That usually means the wrong entity is selected or the file belongs to another'
        + ' entity. Nothing was imported. Re-send with allow_over_gl:true to override.');
    }
  }

  const existing = db.prepare("SELECT id FROM ar_invoices WHERE entity_id = ? AND origin = 'opening'").all(eid).map(r => r.id);
  if (existing.length) {
    const ph = existing.map(() => '?').join(',');
    const withRec = db.prepare('SELECT COUNT(*) AS c FROM ar_receipts WHERE invoice_id IN (' + ph + ')').get(...existing).c;
    if (withRec > 0 && !opts.force) throw new Error('Opening items already have ' + withRec + ' applied receipt(s); pass force:true to replace');
    db.prepare('DELETE FROM ar_receipts WHERE invoice_id IN (' + ph + ')').run(...existing);
    db.prepare('DELETE FROM ar_invoice_lines WHERE invoice_id IN (' + ph + ')').run(...existing);
    const del = db.prepare('DELETE FROM ar_invoices WHERE id = ?');
    for (const id of existing) del.run(id);
  }
  const findCust = db.prepare('SELECT id, name FROM ar_customers WHERE entity_id = ? AND name = ?');
  const insCust = db.prepare('INSERT INTO ar_customers (entity_id, name) VALUES (?, ?)');
  const used = new Set(db.prepare('SELECT invoice_num FROM ar_invoices WHERE entity_id = ?').all(eid).map(r => r.invoice_num));
  const insInv = db.prepare('INSERT INTO ar_invoices (entity_id, customer_id, invoice_num, invoice_date, due_date, aging_date, customer_name, subtotal, total, ar_account_code, status, sent_at, origin, created_by) '
    + "VALUES (?,?,?,?,?,?,?,?,?,?,'sent',?,'opening',?)");
  let inserted = 0, total = 0; const out = [];
  db.transaction(() => {
    for (const it of items) {
      const name = String(it.customer_name || '').trim() || '(no customer)';
      let cust = findCust.get(eid, name);
      if (!cust) { const r = insCust.run(eid, name); cust = { id: r.lastInsertRowid, name }; }
      let num = String(it.document_no || it.invoice_num || '').trim() || ('OPEN-' + (inserted + 1));
      if (used.has(num)) { let n = 2; while (used.has(num + '-' + n)) n++; num = num + '-' + n; }
      used.add(num);
      const amt = r2(it.amount);
      const invDate = isDate(it.invoice_date) ? it.invoice_date : (isDate(it.due_date) ? it.due_date : todayStr());
      const dueDate = isDate(it.due_date) ? it.due_date : invDate;
      const agingDate = isDate(it.posting_date) ? it.posting_date : invDate;
      insInv.run(eid, cust.id, num, invDate, dueDate, agingDate, name, amt, amt, arCode, invDate, opts.who || 'opening-import');
      inserted++; total = r2(total + amt); out.push({ invoice_num: num, customer: name, amount: amt });
    }
  })();
  return { inserted, total, ar_account: arCode, items: out };
}

// Apply a cash receipt to an invoice WITHOUT posting a new JE — used when the
// cash side is already booked elsewhere (e.g. a bank-transaction post whose JE
// debits the bank and credits the A/R control account). Records the subledger
// allocation and flips the invoice to paid when fully applied.
function recordArReceipt(db, o) {
  const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(o.invoice_id, o.entity_id);
  if (!inv) throw new Error('Invoice ' + o.invoice_id + ' not found for entity ' + o.entity_id);
  if (inv.status === 'void') throw new Error('Invoice ' + inv.invoice_num + ' is void');
  const amt = r2(o.amount);
  if (!(amt > 0.005)) throw new Error('Receipt amount must be positive');
  const priorPaid = r2(db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id = ?').get(o.invoice_id).p);
  if (Number(inv.total || 0) >= 0 && (priorPaid + amt) - r2(inv.total) > 0.005) {
    throw new Error('Receipt of ' + amt.toFixed(2) + ' exceeds the open balance of ' + r2(Number(inv.total || 0) - priorPaid).toFixed(2) + ' on ' + inv.invoice_num);
  }
  db.prepare('INSERT INTO ar_receipts (invoice_id, entity_id, date, amount, bank_account_code, memo, je_id, created_by) VALUES (?,?,?,?,?,?,?,?)')
    .run(o.invoice_id, o.entity_id, o.date, amt, o.bank_account_code || '', o.memo || null, o.je_id || null, o.created_by || null);
  const paid = r2(db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id = ?').get(o.invoice_id).p);
  if (paid >= r2(inv.total) - 0.005) db.prepare("UPDATE ar_invoices SET status='paid', paid_at=? WHERE id=?").run(o.date, o.invoice_id);
  return { invoice_num: inv.invoice_num, paid, open: r2(Number(inv.total || 0) - paid) };
}

// ═══ route registration ════════════════════════════════════════════════════
function registerArRoutes(app, ctx) {
  const db = ctx.db;
  const { auth, requireEntityAccess, requireRole, workpapersDir, verifyToken } = ctx;
  const writers = [auth, requireEntityAccess(), requireRole('Admin', 'Accountant')];
  const readers = [auth, requireEntityAccess()];
  const who = (req) => (req.user && (req.user.name || req.user.email)) || 'ar-module';
  const entName = (eid) => (db.prepare('SELECT name FROM entities WHERE id = ?').get(eid) || {}).name || 'Entity';
  const fail = (res, e) => res.status(400).json({ error: e.message || String(e) });

  ensureSchema(db);

  // ── settings ──
  app.get('/api/entities/:eid/ar/settings', ...readers, (req, res) => {
    const s = getSettings(db, req.params.eid);
    res.json(Object.assign({}, s, { resolved_ar_account: defaultArAccount(db, req.params.eid), email_configured: !!ctx.getResendKey() }));
  });

  app.put('/api/entities/:eid/ar/settings', ...writers, (req, res) => {
    const eid = req.params.eid, b = req.body || {};
    const cur = getSettings(db, eid);
    const val = (k) => (b[k] !== undefined ? (String(b[k] || '').trim() || null) : cur[k]);
    db.prepare('INSERT INTO ar_settings (entity_id, remit_to, bill_from, default_ar_account, invoice_prefix, footer_note, reply_to, updated_at) '
      + "VALUES (?,?,?,?,?,?,?, datetime('now')) "
      + 'ON CONFLICT(entity_id) DO UPDATE SET remit_to=excluded.remit_to, bill_from=excluded.bill_from, '
      + 'default_ar_account=excluded.default_ar_account, invoice_prefix=excluded.invoice_prefix, '
      + "footer_note=excluded.footer_note, reply_to=excluded.reply_to, updated_at=datetime('now')")
      .run(eid, val('remit_to'), val('bill_from'), val('default_ar_account'), val('invoice_prefix'), val('footer_note'), val('reply_to'));
    res.json(Object.assign({}, getSettings(db, eid), { resolved_ar_account: defaultArAccount(db, eid) }));
  });

  // ── recurring templates ──
  app.get('/api/entities/:eid/ar/templates', ...readers, (req, res) => {
    const rows = db.prepare('SELECT t.*, c.name AS customer_name, c.email AS customer_email '
      + 'FROM ar_invoice_templates t JOIN ar_customers c ON c.id = t.customer_id '
      + 'WHERE t.entity_id = ? ORDER BY t.active DESC, c.name').all(req.params.eid);
    const lineStmt = db.prepare('SELECT * FROM ar_template_lines WHERE template_id = ? ORDER BY sort, id');
    for (const t of rows) {
      t.lines = lineStmt.all(t.id);
      t.amount = r2(t.lines.reduce((s, l) => s + Number(l.qty || 0) * Number(l.rate || 0), 0));
    }
    res.json(rows);
  });

  app.post('/api/entities/:eid/ar/templates', ...writers, (req, res) => {
    try {
      const eid = req.params.eid, b = req.body || {};
      const cust = db.prepare('SELECT * FROM ar_customers WHERE id = ? AND entity_id = ?').get(b.customer_id, eid);
      if (!cust) throw new Error('Customer not found for this entity');
      const lines = normalizeLines(db, eid, b.lines);
      const freq = ['monthly', 'quarterly', 'annual'].includes(b.frequency) ? b.frequency : 'monthly';
      const dom = Math.min(Math.max(Number(b.day_of_month) || 1, 1), 28);
      const arCode = String(b.ar_account_code || defaultArAccount(db, eid) || '').trim();
      if (!arCode) throw new Error('No Accounts Receivable account found for this entity');
      const id = db.transaction(() => {
        const r = db.prepare('INSERT INTO ar_invoice_templates (entity_id, customer_id, memo, frequency, day_of_month, next_run, ar_account_code, active) VALUES (?,?,?,?,?,?,?,1)')
          .run(eid, cust.id, String(b.memo || '').trim() || null, freq, dom, isDate(b.next_run) ? b.next_run : null, arCode);
        const insL = db.prepare('INSERT INTO ar_template_lines (template_id, description, qty, rate, revenue_account_code, class_id, location_id, sort) VALUES (?,?,?,?,?,?,?,?)');
        for (const l of lines) insL.run(r.lastInsertRowid, l.description, l.qty, l.rate, l.revenue_account_code, l.class_id, l.location_id, l.sort);
        return r.lastInsertRowid;
      })();
      res.json({ id: id });
    } catch (e) { fail(res, e); }
  });

  app.patch('/api/entities/:eid/ar/templates/:id', ...writers, (req, res) => {
    try {
      const eid = req.params.eid, b = req.body || {};
      const t = db.prepare('SELECT * FROM ar_invoice_templates WHERE id = ? AND entity_id = ?').get(req.params.id, eid);
      if (!t) return res.status(404).json({ error: 'Not found' });
      const memo = b.memo !== undefined ? (String(b.memo || '').trim() || null) : t.memo;
      const freq = ['monthly', 'quarterly', 'annual'].includes(b.frequency) ? b.frequency : t.frequency;
      const dom = b.day_of_month !== undefined ? Math.min(Math.max(Number(b.day_of_month) || 1, 1), 28) : t.day_of_month;
      const nextRun = b.next_run !== undefined ? (isDate(b.next_run) ? b.next_run : null) : t.next_run;
      const active = b.active !== undefined ? (b.active ? 1 : 0) : t.active;
      const arCode = b.ar_account_code !== undefined ? (String(b.ar_account_code || '').trim() || t.ar_account_code) : t.ar_account_code;
      db.transaction(() => {
        db.prepare('UPDATE ar_invoice_templates SET memo=?, frequency=?, day_of_month=?, next_run=?, ar_account_code=?, active=? WHERE id=? AND entity_id=?')
          .run(memo, freq, dom, nextRun, arCode, active, t.id, eid);
        if (Array.isArray(b.lines)) {
          const lines = normalizeLines(db, eid, b.lines);
          db.prepare('DELETE FROM ar_template_lines WHERE template_id = ?').run(t.id);
          const insL = db.prepare('INSERT INTO ar_template_lines (template_id, description, qty, rate, revenue_account_code, class_id, location_id, sort) VALUES (?,?,?,?,?,?,?,?)');
          for (const l of lines) insL.run(t.id, l.description, l.qty, l.rate, l.revenue_account_code, l.class_id, l.location_id, l.sort);
        }
      })();
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/entities/:eid/ar/templates/:id', ...writers, (req, res) => {
    db.prepare('DELETE FROM ar_invoice_templates WHERE id = ? AND entity_id = ?').run(req.params.id, req.params.eid);
    res.json({ success: true });
  });

  // Generate an invoice from a template ("run now"); also advances next_run.
  app.post('/api/entities/:eid/ar/templates/:id/generate', ...writers, (req, res) => {
    try {
      const eid = req.params.eid;
      const t = db.prepare('SELECT * FROM ar_invoice_templates WHERE id = ? AND entity_id = ?').get(req.params.id, eid);
      if (!t) return res.status(404).json({ error: 'Not found' });
      const tLines = db.prepare('SELECT * FROM ar_template_lines WHERE template_id = ? ORDER BY sort, id').all(t.id);
      if (!tLines.length) throw new Error('This template has no lines');
      const invoice_date = isDate(req.body && req.body.invoice_date) ? req.body.invoice_date
        : (isDate(t.next_run) ? t.next_run : todayStr());
      const id = createInvoice(db, eid, {
        customer_id: t.customer_id, template_id: t.id, invoice_date: invoice_date, memo: t.memo,
        ar_account_code: t.ar_account_code,
        lines: tLines.map(l => ({ description: l.description, qty: l.qty, rate: l.rate, revenue_account_code: l.revenue_account_code, class_id: l.class_id, location_id: l.location_id })),
      }, who(req));
      db.prepare('UPDATE ar_invoice_templates SET next_run = ? WHERE id = ?').run(advanceNextRun(invoice_date, t.frequency, t.day_of_month), t.id);
      res.json(invoiceWithLines(db, eid, id));
    } catch (e) { fail(res, e); }
  });

  // ── invoices ──
  app.get('/api/entities/:eid/ar/invoices', ...readers, (req, res) => {
    const eid = req.params.eid;
    const where = ['entity_id = ?']; const args = [eid];
    if (req.query.status) { where.push('status = ?'); args.push(req.query.status); }
    if (isDate(req.query.from)) { where.push('invoice_date >= ?'); args.push(req.query.from); }
    if (isDate(req.query.to)) { where.push('invoice_date <= ?'); args.push(req.query.to); }
    const rows = db.prepare('SELECT * FROM ar_invoices WHERE ' + where.join(' AND ') + ' ORDER BY invoice_date DESC, id DESC').all(...args);
    const paid = new Map(db.prepare('SELECT invoice_id, SUM(amount) AS p FROM ar_receipts WHERE entity_id = ? GROUP BY invoice_id').all(eid).map(r => [r.invoice_id, Number(r.p || 0)]));
    const now = todayStr();
    for (const r of rows) {
      r.paid_amount = r2(paid.get(r.id) || 0);
      r.open_amount = r2(Number(r.total || 0) - r.paid_amount);
      r.overdue = r.status !== 'paid' && r.status !== 'void' && !!r.due_date && r.due_date < now && r.open_amount > 0.005;
    }
    res.json(rows);
  });

  app.get('/api/entities/:eid/ar/invoices/:id', ...readers, (req, res) => {
    const inv = invoiceWithLines(db, req.params.eid, req.params.id);
    if (!inv) return res.status(404).json({ error: 'Not found' });
    res.json(inv);
  });

  app.post('/api/entities/:eid/ar/invoices', ...writers, (req, res) => {
    try {
      const id = createInvoice(db, req.params.eid, req.body || {}, who(req));
      res.json(invoiceWithLines(db, req.params.eid, id));
    } catch (e) { fail(res, e); }
  });

  // Draft invoices can be edited in place (the JE is rebuilt); issued ones cannot.
  app.patch('/api/entities/:eid/ar/invoices/:id', ...writers, (req, res) => {
    try {
      const eid = req.params.eid, b = req.body || {};
      const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(req.params.id, eid);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      if (inv.status !== 'draft') throw new Error('Only draft invoices can be edited. Void it and issue a new one instead.');
      const invoice_date = isDate(b.invoice_date) ? b.invoice_date : inv.invoice_date;
      const due_date = isDate(b.due_date) ? b.due_date : inv.due_date;
      const memo = b.memo !== undefined ? (String(b.memo || '').trim() || null) : inv.memo;
      const arCode = String(b.ar_account_code || inv.ar_account_code).trim();
      const lines = Array.isArray(b.lines) ? normalizeLines(db, eid, b.lines)
        : db.prepare('SELECT * FROM ar_invoice_lines WHERE invoice_id = ? ORDER BY sort, id').all(inv.id);
      const total = r2(lines.reduce((s, l) => s + Number(l.amount != null ? l.amount : Number(l.qty) * Number(l.rate)), 0));
      db.transaction(() => {
        db.prepare('UPDATE ar_invoices SET invoice_date=?, due_date=?, memo=?, ar_account_code=?, subtotal=?, total=? WHERE id=?')
          .run(invoice_date, due_date, memo, arCode, total, total, inv.id);
        if (Array.isArray(b.lines)) {
          db.prepare('DELETE FROM ar_invoice_lines WHERE invoice_id = ?').run(inv.id);
          const insL = db.prepare('INSERT INTO ar_invoice_lines (invoice_id, description, qty, rate, amount, revenue_account_code, class_id, location_id, sort) VALUES (?,?,?,?,?,?,?,?,?)');
          for (const l of lines) insL.run(inv.id, l.description, l.qty, l.rate, l.amount, l.revenue_account_code, l.class_id, l.location_id, l.sort);
        }
        // Rebuild the accrual JE so the GL always matches the invoice.
        deleteJE(db, inv.je_id);
        const jeLines = [{ account_code: arCode, debit: total, credit: 0, description: 'Invoice ' + inv.invoice_num + ' - ' + inv.customer_name }];
        for (const l of lines) jeLines.push({ account_code: l.revenue_account_code, debit: 0, credit: r2(l.amount != null ? l.amount : Number(l.qty) * Number(l.rate)), description: l.description, class_id: l.class_id, location_id: l.location_id });
        const je = postJE(db, eid, invoice_date, 'AR Invoice ' + inv.invoice_num + ' - ' + inv.customer_name + (memo ? ' - ' + memo : ''), jeLines, who(req));
        db.prepare('UPDATE ar_invoices SET je_id = ? WHERE id = ?').run(je.id, inv.id);
      })();
      res.json(invoiceWithLines(db, eid, inv.id));
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/entities/:eid/ar/invoices/:id', ...writers, (req, res) => {
    const eid = req.params.eid;
    const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(req.params.id, eid);
    if (!inv) return res.status(404).json({ error: 'Not found' });
    if (inv.status !== 'draft') return res.status(409).json({ error: 'Only draft invoices can be deleted. Use Void for issued invoices.' });
    const paid = db.prepare('SELECT COUNT(*) AS n FROM ar_receipts WHERE invoice_id = ?').get(inv.id).n;
    if (paid) return res.status(409).json({ error: 'Invoice has payments recorded; remove them first.' });
    db.transaction(() => {
      deleteJE(db, inv.je_id);
      db.prepare('DELETE FROM ar_invoices WHERE id = ?').run(inv.id);
    })();
    res.json({ success: true });
  });

  // Void: draft deletes its JE; issued posts a reversing JE.
  app.post('/api/entities/:eid/ar/invoices/:id/void', ...writers, (req, res) => {
    try {
      const eid = req.params.eid;
      const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(req.params.id, eid);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      if (inv.status === 'void') return res.json({ success: true, already: true });
      const paid = db.prepare('SELECT COUNT(*) AS n FROM ar_receipts WHERE invoice_id = ?').get(inv.id).n;
      if (paid) throw new Error('Invoice has payments recorded; delete the receipts before voiding.');
      const voidDate = isDate(req.body && req.body.date) ? req.body.date : todayStr();
      const lines = db.prepare('SELECT * FROM ar_invoice_lines WHERE invoice_id = ? ORDER BY sort, id').all(inv.id);
      db.transaction(() => {
        if (inv.status === 'draft') {
          deleteJE(db, inv.je_id);
          db.prepare("UPDATE ar_invoices SET status='void', je_id=NULL, voided_at=datetime('now') WHERE id=?").run(inv.id);
        } else {
          const jeLines = [{ account_code: inv.ar_account_code, debit: 0, credit: r2(inv.total), description: 'Void invoice ' + inv.invoice_num }];
          for (const l of lines) jeLines.push({ account_code: l.revenue_account_code, debit: r2(l.amount), credit: 0, description: 'Void: ' + l.description, class_id: l.class_id, location_id: l.location_id });
          const je = postJE(db, eid, voidDate, 'Void AR Invoice ' + inv.invoice_num + ' - ' + inv.customer_name, jeLines, who(req));
          db.prepare("UPDATE ar_invoices SET status='void', void_je_id=?, voided_at=datetime('now') WHERE id=?").run(je.id, inv.id);
        }
      })();
      res.json(invoiceWithLines(db, eid, inv.id));
    } catch (e) { fail(res, e); }
  });

  // ── PDF ──
  async function pdfBufferFor(eid, id) {
    const inv = invoiceWithLines(db, eid, id);
    if (!inv) return null;
    return buildInvoicePdf({ entityName: entName(eid), settings: getSettings(db, eid), invoice: inv });
  }

  // Inline view. Token via query param so a plain <a href> works, matching the
  // entity-files download route.
  app.get('/api/entities/:eid/ar/invoices/:id/pdf', async (req, res) => {
    try {
      if (!req.query.token) return res.status(401).json({ error: 'Token required' });
      verifyToken(req.query.token);
    } catch (e) { return res.status(401).json({ error: 'Invalid token' }); }
    try {
      const inv = invoiceWithLines(db, req.params.eid, req.params.id);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      const buf = await pdfBufferFor(req.params.eid, req.params.id);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'inline; filename="' + inv.invoice_num + '.pdf"');
      res.send(buf);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Save a copy into Workpapers > Invoices/<year> without emailing.
  app.post('/api/entities/:eid/ar/invoices/:id/save-pdf', ...writers, async (req, res) => {
    try {
      const eid = req.params.eid;
      const inv = invoiceWithLines(db, eid, req.params.id);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      const buf = await pdfBufferFor(eid, inv.id);
      const fileId = savePdf(db, workpapersDir, eid, 'Invoices/' + String(inv.invoice_date).slice(0, 4), inv.invoice_num + '.pdf', buf, who(req));
      db.prepare('UPDATE ar_invoices SET pdf_file_id = ? WHERE id = ?').run(fileId, inv.id);
      res.json({ success: true, pdf_file_id: fileId });
    } catch (e) { fail(res, e); }
  });

  // ── send ──
  app.post('/api/entities/:eid/ar/invoices/:id/send', ...writers, async (req, res) => {
    try {
      const eid = req.params.eid;
      const inv = invoiceWithLines(db, eid, req.params.id);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      if (inv.status === 'void') throw new Error('This invoice is void');
      const to = String((req.body && req.body.to) || inv.customer_email || '').trim();
      if (!to) throw new Error('No recipient email. Add one on the customer record or pass one in.');
      const settings = getSettings(db, eid);
      const buf = await pdfBufferFor(eid, inv.id);
      const fileId = savePdf(db, workpapersDir, eid, 'Invoices/' + String(inv.invoice_date).slice(0, 4), inv.invoice_num + '.pdf', buf, who(req));
      db.prepare('UPDATE ar_invoices SET pdf_file_id = ? WHERE id = ?').run(fileId, inv.id);
      const name = entName(eid);
      const result = await sendInvoiceEmail({
        apiKey: ctx.getResendKey(),
        from: ctx.getFromEmail(),
        replyTo: settings.reply_to || null,
        to: to,
        cc: (req.body && req.body.cc) || null,
        subject: 'Invoice ' + inv.invoice_num + ' from ' + name,
        html: invoiceEmailHtml({ entityName: name, invoice: inv, settings: settings }),
        filename: inv.invoice_num + '.pdf',
        pdf: buf,
      });
      if (!result.ok) return res.status(502).json({ error: 'Email not sent: ' + (result.reason || 'unknown error'), pdf_saved: true, pdf_file_id: fileId });
      db.prepare("UPDATE ar_invoices SET status = CASE WHEN status='draft' THEN 'sent' ELSE status END, sent_at = datetime('now') WHERE id = ?").run(inv.id);
      res.json({ success: true, sent_to: to, pdf_file_id: fileId, invoice: invoiceWithLines(db, eid, inv.id) });
    } catch (e) { fail(res, e); }
  });

  // Mark as sent without emailing (invoice delivered by other means).
  app.post('/api/entities/:eid/ar/invoices/:id/mark-sent', ...writers, (req, res) => {
    const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ? AND entity_id = ?').get(req.params.id, req.params.eid);
    if (!inv) return res.status(404).json({ error: 'Not found' });
    if (inv.status !== 'draft') return res.status(409).json({ error: 'Only draft invoices can be marked sent' });
    db.prepare("UPDATE ar_invoices SET status='sent', sent_at=datetime('now') WHERE id=?").run(inv.id);
    res.json({ success: true });
  });

  // ── receipts (cash application) ──
  app.post('/api/entities/:eid/ar/invoices/:id/receipts', ...writers, (req, res) => {
    try {
      const eid = req.params.eid, b = req.body || {};
      const inv = invoiceWithLines(db, eid, req.params.id);
      if (!inv) return res.status(404).json({ error: 'Not found' });
      if (inv.status === 'void') throw new Error('This invoice is void');
      const date = isDate(b.date) ? b.date : todayStr();
      const bank = String(b.bank_account_code || '').trim();
      if (!bank) throw new Error('Bank account is required');
      const bankAcct = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ? AND code = ?').get(eid, bank);
      if (!bankAcct) throw new Error('Bank account ' + bank + ' is not in this entity\'s chart of accounts');
      const amount = r2(b.amount != null ? b.amount : inv.open_amount);
      if (!(amount > 0.005)) throw new Error('Payment amount must be positive');
      if (amount - inv.open_amount > 0.005) throw new Error('Payment of ' + MONEY(amount) + ' exceeds the open balance of ' + MONEY(inv.open_amount));
      const memo = String(b.memo || '').trim() || null;
      db.transaction(() => {
        const je = postJE(db, eid, date, 'AR Receipt - Invoice ' + inv.invoice_num + ' - ' + inv.customer_name + (memo ? ' - ' + memo : ''), [
          { account_code: bank, debit: amount, credit: 0, description: 'Payment on invoice ' + inv.invoice_num },
          { account_code: inv.ar_account_code, debit: 0, credit: amount, description: 'Payment on invoice ' + inv.invoice_num + ' - ' + inv.customer_name },
        ], who(req));
        db.prepare('INSERT INTO ar_receipts (invoice_id, entity_id, date, amount, bank_account_code, memo, je_id, created_by) VALUES (?,?,?,?,?,?,?,?)')
          .run(inv.id, eid, date, amount, bank, memo, je.id, who(req));
        const paid = r2(db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id = ?').get(inv.id).p);
        if (paid >= r2(inv.total) - 0.005) db.prepare("UPDATE ar_invoices SET status='paid', paid_at=? WHERE id=?").run(date, inv.id);
      })();
      res.json(invoiceWithLines(db, eid, inv.id));
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/entities/:eid/ar/invoices/:id/receipts/:rid', ...writers, (req, res) => {
    const eid = req.params.eid;
    const rec = db.prepare('SELECT * FROM ar_receipts WHERE id = ? AND invoice_id = ? AND entity_id = ?').get(req.params.rid, req.params.id, eid);
    if (!rec) return res.status(404).json({ error: 'Not found' });
    db.transaction(() => {
      deleteJE(db, rec.je_id);
      db.prepare('DELETE FROM ar_receipts WHERE id = ?').run(rec.id);
      const inv = db.prepare('SELECT * FROM ar_invoices WHERE id = ?').get(rec.invoice_id);
      const paid = r2(db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id = ?').get(rec.invoice_id).p);
      if (inv && inv.status === 'paid' && paid < r2(inv.total) - 0.005) {
        db.prepare("UPDATE ar_invoices SET status = CASE WHEN sent_at IS NULL THEN 'draft' ELSE 'sent' END, paid_at = NULL WHERE id = ?").run(inv.id);
      }
    })();
    res.json({ success: true });
  });

  // ── opening subledger import + AR cash-application picker feed ──
  app.post('/api/entities/:eid/ar/opening-import', ...writers, (req, res) => {
    try {
      const items = Array.isArray(req.body && req.body.items) ? req.body.items : null;
      if (!items || !items.length) throw new Error('items[] is required');
      const b = req.body || {};
      res.json(importOpeningItems(db, req.params.eid, items, {
        who: who(req), force: !!b.force,
        as_of: b.as_of, allow_over_gl: !!b.allow_over_gl,
      }));
    } catch (e) { fail(res, e); }
  });

  app.get('/api/entities/:eid/ar/opening', ...readers, (req, res) => {
    const r = db.prepare("SELECT COUNT(*) AS n, COALESCE(SUM(total),0) AS total FROM ar_invoices WHERE entity_id = ? AND origin = 'opening'").get(req.params.eid);
    res.json({ count: r.n, total: r2(r.total) });
  });

  // Open invoices a bank deposit can be applied to; flags exact-amount matches
  // so the coding screen can auto-suggest a single invoice for the deposit.
  app.get('/api/entities/:eid/ar/open-invoices', ...readers, (req, res) => {
    const eid = req.params.eid;
    const amount = req.query.amount != null && req.query.amount !== '' ? r2(req.query.amount) : null;
    const invs = db.prepare("SELECT * FROM ar_invoices WHERE entity_id = ? AND status != 'void' ORDER BY customer_name, invoice_date, invoice_num").all(eid);
    const paidBy = new Map();
    for (const r of db.prepare('SELECT invoice_id, SUM(amount) AS paid FROM ar_receipts WHERE entity_id = ? GROUP BY invoice_id').all(eid)) paidBy.set(r.invoice_id, Number(r.paid || 0));
    const today = todayStr(); const open = [];
    for (const inv of invs) {
      const openAmt = r2(Number(inv.total || 0) - (paidBy.get(inv.id) || 0));
      if (Math.abs(openAmt) < 0.005) continue;
      const agingRef = inv.aging_date || inv.due_date;
      const past = agingRef ? daysBetween(agingRef, today) : 0;
      const bucket = past <= 0 ? 'current' : past <= 30 ? 'd1_30' : past <= 60 ? 'd31_60' : past <= 90 ? 'd61_90' : 'd90_plus';
      open.push({ id: inv.id, invoice_num: inv.invoice_num, customer: inv.customer_name, invoice_date: inv.invoice_date,
        due_date: inv.due_date, open: openAmt, bucket, days_past_due: Math.max(past, 0),
        exact_match: amount != null && Math.abs(openAmt - amount) < 0.005 });
    }
    res.json({ amount, invoices: open, exact_matches: amount != null ? open.filter(o => o.exact_match) : [] });
  });

  // ── aging ──
  app.get('/api/entities/:eid/ar/aging', ...readers, (req, res) => {
    const asOf = isDate(req.query.as_of) ? req.query.as_of : todayStr();
    res.json(buildAging(db, req.params.eid, asOf));
  });
}

module.exports = {
  registerArRoutes, buildInvoicePdf, buildAging, defaultArAccount, ensureSchema,
  importOpeningItems, recordArReceipt,
  // exported for tests
  createInvoice, nextInvoiceNum, advanceNextRun, postJE, invoiceWithLines,
};
