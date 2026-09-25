// ─── Monthly Closing Workpapers (CLA leadsheet format, Banyan Residential) ────
//
// One monthly workbook reproducing CLA's numbered close leadsheets, each backed
// by the SUBSTANTIVE schedule CLA uses to prove the balance — not a GL dump:
//
//   • Cash            — per bank account: register balance per GL and the
//                       reconciled balance per bank rec (statement ± outstanding
//                       checks / deposits in transit), with the difference.
//   • Receivables     — A/R Aging (by customer / invoice, from CL's invoice
//                       subledger, tied to the GL) + A/R Summary: per-customer
//                       aging, prior month vs current month, with the difference.
//   • Payables        — A/P Aging (by vendor, Bill.com-synced bills, tied to the
//                       GL) + A/P Summary, prior vs current, with the difference.
//   • Prepaids        — per prepaid account, the amortization schedule (CLA
//                       layout: Balance P0 then Additions / Expenses / Balance for
//                       each month of the year), from the prepaid register.
//   • Fixed Assets    — the Fixed Asset Depreciation Schedule (per asset: life,
//                       in-service, cost, monthly depr, accum beg → monthly
//                       columns → accum end, NBV), grouped by asset account.
//   • Credit Cards    — statement-date reconciliation rolled to month end.
//   • Intercompany    — Tie Out: our balance vs the counterparty's own ledger.
//   • Equity          — Equity Rollforward: prior year end → monthly activity →
//                       ending, by equity account.
//   • Other Assets / Investments / Debt / Other Liabilities — a schedule per
//                       category: Beginning (prior month) → Activity → Ending.
//   • Summary         — exceptions, the Assets = Liab + Equity + NI tie, and a
//                       current-month vs prior-month balance comparison for
//                       every account (prior month straight from CloudLedger).
//
// Leadsheets link by formula to their schedules (CLA's convention), so a
// schedule that disagrees with the GL surfaces as a flagged difference rather
// than being papered over. Literals are atomic GL/subledger figures, register
// fields, and blank blue input cells for data CloudLedger does not hold (bank
// and card statement balances).
//
// Registers: cla_prepaid_items and cla_fixed_assets (per entity), seeded once
// for Banyan Residential from CLA's own schedules (./seeds/cla_banyan_seed.js)
// and maintained through GET/PUT .../registers.
//
// Reuses the GL layer of ./monthlyclose (buildData, resolveMonth), the A/R
// aging of ./ar (buildAging) and the A/P aging passed in ctx (buildApAging).
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { buildData, resolveMonth } = require('./monthlyclose');
const { buildAging } = require('./ar');
const SEED = require('./seeds/cla_banyan_seed');

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// ─── Display constants (CLA style) ────────────────────────────────────────────
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const DATEFMT = 'mm-dd-yy';
const PCT = '0.0%';
const FONT = 'Calibri';
const YELLOW = 'FFFFFF00';
const RED = 'FFFF0000';
const BLUE = 'FF0000FF';
const LINK = 'FF0563C1';
const HDRFILL = 'FFDDEBF7';
const SECTFILL = 'FFBDD7EE';
const GRPFILL = 'FFF2F2F2';
const ENTRY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: YELLOW } };
const INPUT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const F = (o = {}) => Object.assign({ name: FONT, size: 11 }, o);
const THIN = { style: 'thin' };
const DBL = { style: 'double' };
const box = { top: THIN, bottom: THIN, left: THIN, right: THIN };

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const short = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return mo + '/' + dd + '/' + String(y).slice(2); };
const spell = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return MONTHS[mo - 1] + ' ' + dd + ', ' + y; };
const asDate = (d) => { if (!d) return null; const [y, mo, dd] = String(d).split('-').map(Number); return new Date(Date.UTC(y, mo - 1, dd)); };
const qn = (s) => "'" + String(s).replace(/'/g, "''") + "'";
const fmt = (n) => '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const colL = (n) => { let s = ''; while (n > 0) { const k = (n - 1) % 26; s = String.fromCharCode(65 + k) + s; n = Math.floor((n - 1) / 26); } return s; };
const monthEndOf = (y, mo) => new Date(Date.UTC(y, mo, 0)).toISOString().slice(0, 10); // mo = 1..12
const ym = (d) => String(d || '').slice(0, 7);

function sheetNameFor(code, used) {
  let base = String(code).replace(/[\[\]\:\*\?\/\\]/g, '_').slice(0, 31) || 'acct';
  let name = base, i = 1;
  while (used.has(name.toLowerCase())) { const suf = '_' + (++i); name = base.slice(0, 31 - suf.length) + suf; }
  used.add(name.toLowerCase());
  return name;
}
function reserveName(name, used) { used.add(String(name).toLowerCase()); return name; }

// CLA title block: entity, "<X> LEADSHEET", Month/Year Ended entry cells + legend.
function titleBlock(ws, entityName, leadTitle, m, legendCol) {
  ws.getCell('C1').value = entityName; ws.getCell('C1').font = F({ size: 16, bold: true });
  ws.getCell('C2').value = leadTitle; ws.getCell('C2').font = F({ size: 12, bold: true });
  ws.getCell('C3').value = 'Month Ended:'; ws.getCell('C3').font = F();
  const d3 = ws.getCell('D3');
  d3.value = asDate(m.end); d3.numFmt = DATEFMT; d3.font = F({ bold: true, color: { argb: RED } });
  d3.fill = ENTRY_FILL; d3.alignment = { horizontal: 'center' }; d3.border = box;
  ws.getCell('C4').value = 'Year Ended:'; ws.getCell('C4').font = F();
  const d4 = ws.getCell('D4');
  d4.value = { formula: 'DATE(YEAR($D$3),12,31)' }; d4.numFmt = DATEFMT; d4.font = F({ bold: true });
  d4.alignment = { horizontal: 'center' }; d4.border = box;
  const lc = legendCol || 'G';
  ws.getCell(lc + '3').value = '= Entry Cell'; ws.getCell(lc + '3').font = F();
  const sw = ws.getCell(lc + '2'); sw.fill = ENTRY_FILL; sw.border = box;
}

function hdr(ws, rowNum, startCol, titles) {
  titles.forEach((t, i) => {
    const c = ws.getRow(rowNum).getCell(startCol + i);
    c.value = t; c.font = F({ bold: true });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDRFILL } };
    c.alignment = { horizontal: 'center', wrapText: true }; c.border = box;
  });
}
function hdrDates(ws, rowNum, startCol, dates) {
  dates.forEach((d, i) => {
    const c = ws.getRow(rowNum).getCell(startCol + i);
    c.value = asDate(d); c.numFmt = DATEFMT; c.font = F({ bold: true });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDRFILL } };
    c.alignment = { horizontal: 'center' }; c.border = box;
  });
}
function wpLink(ws, addr, sheet, text) {
  const c = ws.getCell(addr);
  if (sheet) c.value = { formula: 'HYPERLINK("#" & "' + qn(sheet) + '!A1","' + String(text).replace(/"/g, '""') + '")' };
  c.font = F({ bold: true, underline: true, color: { argb: LINK } });
  c.alignment = { horizontal: 'center' };
  return c;
}
function backLink(ws, addr, tab) {
  const c = ws.getCell(addr);
  c.value = { formula: 'HYPERLINK("#" & "' + qn(tab) + '!A1","Back to leadsheet")' };
  c.font = F({ underline: true, color: { argb: LINK } });
}
function tabHead(ws, title, entityName, subtitle, backTab) {
  ws.getCell('A1').value = title; ws.getCell('A1').font = F({ size: 12, bold: true });
  ws.getCell('A2').value = entityName; ws.getCell('A2').font = F({ bold: true });
  ws.getCell('A3').value = subtitle; ws.getCell('A3').font = F({ italic: true });
  if (backTab) backLink(ws, 'F1', backTab);
}
function num(ws, addr, v, o = {}) {
  const c = ws.getCell(addr); c.numFmt = o.fmt || ACCT; c.font = F(o.font || {});
  if (v && typeof v === 'object' && v.formula) c.value = v; else c.value = (v == null ? '' : v);
  if (o.border) c.border = o.border; if (o.fill) c.fill = o.fill;
  return c;
}
function txt(ws, addr, v, o = {}) { const c = ws.getCell(addr); c.value = v; c.font = F(o.font || {}); if (o.fmt) c.numFmt = o.fmt; if (o.align) c.alignment = o.align; return c; }
function blueInput(ws, addr) { const c = ws.getCell(addr); c.numFmt = ACCT; c.font = F({ color: { argb: BLUE } }); c.fill = INPUT_FILL; c.border = box; return c; }
function totalCell(ws, addr, formula) { const c = ws.getCell(addr); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL }; c.value = formula ? { formula } : 0; return c; }
const ref = (tab, addr) => qn(tab) + '!' + addr.replace(/([A-Z]+)(\d+)/, '$$$1$$$2');

// ─── Registers (prepaid items, fixed assets) ──────────────────────────────────
function ensureRegisterSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS cla_prepaid_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER NOT NULL, account_code TEXT NOT NULL,
      date_paid TEXT, vendor TEXT, expense_account TEXT, description TEXT, start_date TEXT, end_date TEXT,
      monthly REAL DEFAULT 0, opening_balance REAL DEFAULT 0, premium REAL, sort_order INTEGER DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now')));
    CREATE INDEX IF NOT EXISTS idx_cla_prepaid_ent ON cla_prepaid_items(entity_id);
    CREATE TABLE IF NOT EXISTS cla_fixed_assets (
      id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER NOT NULL, asset_account TEXT NOT NULL, dep_account TEXT,
      description TEXT, life_years REAL, in_service TEXT, cost REAL DEFAULT 0, accum_dep_beg REAL DEFAULT 0,
      sort_order INTEGER DEFAULT 0, created_at TEXT DEFAULT (datetime('now')));
    CREATE INDEX IF NOT EXISTS idx_cla_fa_ent ON cla_fixed_assets(entity_id);
  `);
}
function ensureSeed(db, ent) {
  if (!SEED.entityMatch.test(String(ent.name || '').trim())) return;
  const np = db.prepare('SELECT COUNT(*) c FROM cla_prepaid_items WHERE entity_id = ?').get(ent.id).c;
  if (!np) {
    const ins = db.prepare('INSERT INTO cla_prepaid_items (entity_id, account_code, date_paid, vendor, expense_account, description, start_date, end_date, monthly, opening_balance, premium, sort_order) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)');
    SEED.prepaid.forEach((p, i) => ins.run(ent.id, p.account_code, p.date_paid, p.vendor, p.expense_account, p.description, p.start_date, p.end_date, p.monthly || 0, p.opening_balance || 0, p.premium == null ? null : p.premium, i));
  }
  const nf = db.prepare('SELECT COUNT(*) c FROM cla_fixed_assets WHERE entity_id = ?').get(ent.id).c;
  if (!nf) {
    const ins = db.prepare('INSERT INTO cla_fixed_assets (entity_id, asset_account, dep_account, description, life_years, in_service, cost, accum_dep_beg, sort_order) VALUES (?,?,?,?,?,?,?,?,?)');
    SEED.fixedAssets.forEach((a, i) => ins.run(ent.id, a.asset_account, a.dep_account, a.description, a.life_years, a.in_service, a.cost || 0, a.accum_dep_beg || 0, i));
  }
}
function loadRegisters(db, eid) {
  return {
    prepaid: db.prepare('SELECT * FROM cla_prepaid_items WHERE entity_id = ? ORDER BY account_code, sort_order, id').all(eid),
    fixedAssets: db.prepare('SELECT * FROM cla_fixed_assets WHERE entity_id = ? ORDER BY asset_account, sort_order, id').all(eid),
  };
}
function replaceRegisters(db, eid, body) {
  const tx = db.transaction(() => {
    if (Array.isArray(body.prepaid)) {
      db.prepare('DELETE FROM cla_prepaid_items WHERE entity_id = ?').run(eid);
      const ins = db.prepare('INSERT INTO cla_prepaid_items (entity_id, account_code, date_paid, vendor, expense_account, description, start_date, end_date, monthly, opening_balance, premium, sort_order) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)');
      body.prepaid.forEach((p, i) => ins.run(eid, String(p.account_code || ''), p.date_paid || null, p.vendor || '', p.expense_account == null ? null : String(p.expense_account), p.description || '', p.start_date || null, p.end_date || null, Number(p.monthly) || 0, Number(p.opening_balance) || 0, (p.premium == null || p.premium === '') ? null : Number(p.premium), i));
    }
    if (Array.isArray(body.fixed_assets)) {
      db.prepare('DELETE FROM cla_fixed_assets WHERE entity_id = ?').run(eid);
      const ins = db.prepare('INSERT INTO cla_fixed_assets (entity_id, asset_account, dep_account, description, life_years, in_service, cost, accum_dep_beg, sort_order) VALUES (?,?,?,?,?,?,?,?,?)');
      body.fixed_assets.forEach((a, i) => ins.run(eid, String(a.asset_account || ''), a.dep_account == null ? null : String(a.dep_account), a.description || '', Number(a.life_years) || 0, a.in_service || null, Number(a.cost) || 0, Number(a.accum_dep_beg) || 0, i));
    }
  });
  tx();
}

// ─── Category set (CLA order) ─────────────────────────────────────────────────
const CATS = [
  { key: 'cash', tab: 'Cash Leadsheet', lead: 'CASH LEADSHEET', total: 'Total Cash', type: 'Asset', label: 'Cash' },
  { key: 'ar', tab: 'AR Leadsheet', lead: 'ACCOUNTS RECEIVABLE LEADSHEET', total: 'Total Receivables', type: 'Asset', label: 'Accounts Receivable' },
  { key: 'prepaid', tab: 'Prepaid Leadsheet', lead: 'PREPAID EXPENSES LEAD SHEET', total: 'Total Prepaid Expenses', type: 'Asset', label: 'Prepaid Expenses' },
  { key: 'fixed', tab: 'Fixed Assets Leadsheet', lead: 'FIXED ASSETS LEADSHEET', total: 'Total Fixed Assets', type: 'Asset', label: 'Fixed Assets' },
  { key: 'otherassets', tab: 'Other Assets Leadsheet', lead: 'OTHER ASSETS LEADSHEET', total: 'Total Other Assets', type: 'Asset', label: 'Other Assets', sched: 'Other Assets Schedule' },
  { key: 'invest', tab: 'Investments Leadsheet', lead: 'INVESTMENTS LEADSHEET', total: 'Total Investments', type: 'Asset', label: 'Investments', sched: 'Investments Schedule' },
  { key: 'interco', tab: 'Intercompany Leadsheet', lead: 'INTERCOMPANY LEADSHEET', total: 'Total Intercompany', type: 'Asset', label: 'Intercompany' },
  { key: 'ap', tab: 'AP Leadsheet', lead: 'ACCOUNTS PAYABLE LEADSHEET', total: 'Total Payables', type: 'Liability', label: 'Accounts Payable' },
  { key: 'cc', tab: 'Credit Cards Leadsheet', lead: 'CREDIT CARDS LEADSHEET', total: 'Total Credit Cards', type: 'Liability', label: 'Credit Cards' },
  { key: 'debt', tab: 'Debt Leadsheet', lead: 'DEBT LEADSHEET', total: 'Total Debt', type: 'Liability', label: 'Debt', sched: 'Debt Schedule' },
  { key: 'otherliab', tab: 'Other Liabilities Leadsheet', lead: 'OTHER LIABILITIES LEADSHEET', total: 'Total Other Liabilities', type: 'Liability', label: 'Other Liabilities', sched: 'Other Liabilities Schedule' },
  { key: 'equity', tab: 'Equity Leadsheet', lead: 'EQUITY LEADSHEET', total: 'Total Equity', type: 'Equity', label: 'Equity' },
];
const catOf = (key) => CATS.find((c) => c.key === key);

// Split fixed-asset accounts into asset vs accumulated-depreciation and pair them,
// preferring the register's asset→dep mapping, then a name match.
function splitFixed(rows, fixedAssets) {
  const isDep = (a) => /accumulated\s+(dep|amort)|acc\.?\s*dep|acc\s+depreciation|amortization/i.test(a.name) || a.end < 0;
  const assets = rows.filter((a) => !isDep(a));
  const deps = rows.filter((a) => isDep(a));
  const depFor = new Map();
  for (const fa of fixedAssets || []) if (fa.asset_account && fa.dep_account && !depFor.has(String(fa.asset_account))) depFor.set(String(fa.asset_account), String(fa.dep_account));
  const sig = (n) => String(n).toLowerCase().replace(/accumulated|depreciation|amortization|acc\.?|dep\.?|expense|:|\-/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter((w) => w.length >= 4);
  const pairs = []; const usedDep = new Set();
  for (const a of assets) {
    let dep = null;
    const mapped = depFor.get(String(a.code));
    const mi = mapped ? deps.findIndex((d, i) => !usedDep.has(i) && String(d.code) === mapped) : -1;
    if (mi >= 0) { usedDep.add(mi); dep = deps[mi]; }
    else {
      const words = sig(a.name); let best = null, bestScore = 0;
      deps.forEach((d, i) => { if (usedDep.has(i)) return; const dw = sig(d.name); let s = 0; for (const w of words) if (dw.includes(w)) s += w.length; if (s > bestScore) { bestScore = s; best = i; } });
      if (best != null && bestScore >= 4) { usedDep.add(best); dep = deps[best]; }
    }
    pairs.push({ asset: a, dep });
  }
  deps.forEach((d, i) => { if (!usedDep.has(i)) pairs.push({ asset: null, dep: d }); });
  return pairs;
}

// ─── Data ─────────────────────────────────────────────────────────────────────
function pickPrimary(rows, re, codes) {
  return rows.find((a) => codes.includes(String(a.code))) || rows.find((a) => re.test(String(a.name || ''))) || rows[0] || null;
}
function fyMonthEnds(m) { const out = []; for (let k = 1; k <= m.monthNum; k++) out.push(monthEndOf(Number(m.year), k)); return out; }

// JS mirror of the prepaid schedule (drives the Summary flags; the workbook
// carries the same logic as live formulas). Amortization starts the month after
// the policy start month and runs until the balance is exhausted.
function prepaidSchedule(items, m) {
  const ends = fyMonthEnds(m);
  return items.map((it) => {
    let bal = r2(it.opening_balance || 0);
    const cols = [];
    const amortFrom = it.start_date ? (() => { const y = Number(it.start_date.slice(0, 4)), mo = Number(it.start_date.slice(5, 7)); return mo === 12 ? monthEndOf(y + 1, 1) : monthEndOf(y, mo + 1); })() : null;
    for (const e of ends) {
      const add = (it.premium != null && it.date_paid && ym(it.date_paid) === ym(e)) ? r2(it.premium) : 0;
      const can = amortFrom && e >= amortFrom && (bal + add) > 0.005 && (it.monthly || 0) > 0;
      const exp = can ? -Math.min(r2(it.monthly), r2(bal + add)) : 0;
      bal = r2(bal + add + exp);
      cols.push({ end: e, add, exp: r2(exp), bal });
    }
    return { item: it, cols, ending: bal };
  });
}
// JS mirror of the fixed-asset schedule (straight-line, monthly, from the
// in-service month, capped at cost).
// EDATE with Excel's day clamping; end of life = EDATE(in_service, months) - 1 day.
function edate(d, months) {
  const [y, mo, dd] = String(d).split('-').map(Number);
  const total = mo - 1 + months; const ty = y + Math.floor(total / 12); const tm = ((total % 12) + 12) % 12;
  const dim = new Date(Date.UTC(ty, tm + 1, 0)).getUTCDate();
  return new Date(Date.UTC(ty, tm, Math.min(dd, dim)));
}
function endOfLife(a) {
  const life = Number(a.life_years) || 0; if (!a.in_service || life <= 0) return null;
  const t = edate(a.in_service, Math.round(life * 12)); t.setUTCDate(t.getUTCDate() - 1);
  return t.toISOString().slice(0, 10);
}
function fixedSchedule(assets, m) {
  const ends = fyMonthEnds(m);
  return assets.map((a) => {
    const life = Number(a.life_years) || 0, cost = r2(a.cost), beg = r2(a.accum_dep_beg);
    const monthly = life > 0 ? r2(cost / (life * 12)) : 0;
    const eol = endOfLife(a);
    let accum = beg; const dep = [];
    for (const e of ends) {
      const d = (a.in_service && e >= a.in_service && (!eol || e <= eol) && monthly > 0) ? Math.min(monthly, Math.max(0, r2(cost - accum))) : 0;
      accum = r2(accum + d); dep.push(r2(d));
    }
    return { asset: a, monthly, dep, accumEnd: accum, nbv: r2(cost - accum), endOfLife: eol };
  });
}

function buildClaData(ctx, m, eid) {
  const { db, computeBalances } = ctx;
  const base = buildData(ctx, m, eid);
  ensureRegisterSchema(db); ensureSeed(db, base.entity);
  const reg = loadRegisters(db, eid);
  const byCat = {}; for (const a of base.acctRows) (byCat[a.cat] = byCat[a.cat] || []).push(a);
  const flags = base.flags.slice();

  // A/R aging (current + prior month), from the invoice subledger.
  let arCur = null, arPrior = null;
  try { arCur = buildAging(db, eid, m.end); arPrior = buildAging(db, eid, m.beg); } catch (e) { arCur = null; arPrior = null; }
  const arPrimary = (byCat.ar || []).find((a) => arCur && String(a.code) === String(arCur.ar_account)) || pickPrimary(byCat.ar || [], /^accounts receivable$/i, ['12000', '120000', '11000']);
  if (arCur && Math.abs(arCur.recon_diff || 0) >= 0.01) flags.push({ severity: 'exception', wp: 'AR Aging', message: 'A/R aging does not tie to the GL: aging total ' + fmt(arCur.totals.total) + ' vs GL ' + fmt(arCur.gl_ar_balance) + ' (off by ' + fmt(arCur.recon_diff) + ').' });

  // A/P aging (current + prior month): Bill.com bills by vendor + un-aged GL.
  const apPrimary = pickPrimary(byCat.ap || [], /^accounts payable$/i, ['20000', '202000']);
  let apCur = null, apPrior = null;
  if (apPrimary && typeof ctx.buildApAging === 'function') {
    try { apCur = ctx.buildApAging(eid, m.end, apPrimary.code); apPrior = ctx.buildApAging(eid, m.beg, apPrimary.code); } catch (e) { apCur = null; apPrior = null; }
  }
  if (apCur && Math.abs(apCur.recon_diff || 0) >= 0.01) flags.push({ severity: 'exception', wp: 'AP Aging', message: 'A/P aging does not tie to the GL: aging total ' + fmt(apCur.grand_total.total) + ' vs GL ' + fmt(apCur.gl_balance) + ' (off by ' + fmt(apCur.recon_diff) + ').' });

  // Prepaid schedules vs GL.
  const prepaidByAcct = new Map();
  for (const it of reg.prepaid) { const k = String(it.account_code); if (!prepaidByAcct.has(k)) prepaidByAcct.set(k, []); prepaidByAcct.get(k).push(it); }
  const prepaidSched = new Map();
  for (const a of byCat.prepaid || []) {
    const items = prepaidByAcct.get(String(a.code)) || [];
    const sched = prepaidSchedule(items, m);
    prepaidSched.set(String(a.code), sched);
    const tot = r2(sched.reduce((s, x) => s + x.ending, 0));
    if (!items.length && Math.abs(a.end) >= 0.01) flags.push({ severity: 'review', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': GL balance ' + fmt(a.end) + ' but no items in the prepaid register — add the policies (Registers) so the amortization schedule supports it.' });
    else if (items.length && Math.abs(tot - a.end) >= 0.01) flags.push({ severity: 'exception', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': amortization schedule ' + fmt(tot) + ' does not agree to the GL ' + fmt(a.end) + ' (off by ' + fmt(tot - a.end) + ').' });
  }
  // Fixed-asset schedule vs GL.
  const faSched = fixedSchedule(reg.fixedAssets, m);
  const fixedPairs = splitFixed(byCat.fixed || [], reg.fixedAssets);
  for (const p of fixedPairs) {
    if (p.asset) {
      const rows = faSched.filter((x) => String(x.asset.asset_account) === String(p.asset.code));
      const cost = r2(rows.reduce((s, x) => s + r2(x.asset.cost), 0));
      if (!rows.length && Math.abs(p.asset.end) >= 0.01) flags.push({ severity: 'review', wp: p.asset.code + ' ' + p.asset.name, message: p.asset.code + ' ' + p.asset.name + ': GL cost ' + fmt(p.asset.end) + ' but no assets in the fixed-asset register for this account.' });
      else if (rows.length && Math.abs(cost - p.asset.end) >= 0.01) flags.push({ severity: 'exception', wp: p.asset.code + ' ' + p.asset.name, message: p.asset.code + ' ' + p.asset.name + ': fixed-asset schedule cost ' + fmt(cost) + ' does not agree to the GL ' + fmt(p.asset.end) + ' (off by ' + fmt(cost - p.asset.end) + ').' });
    }
    if (p.dep) {
      const rows = faSched.filter((x) => String(x.asset.dep_account) === String(p.dep.code));
      const accum = r2(rows.reduce((s, x) => s + x.accumEnd, 0));
      if (rows.length && Math.abs(-accum - p.dep.end) >= 0.01) flags.push({ severity: 'exception', wp: p.dep.code + ' ' + p.dep.name, message: p.dep.code + ' ' + p.dep.name + ': schedule accumulated depreciation ' + fmt(-accum) + ' does not agree to the GL ' + fmt(p.dep.end) + ' (off by ' + fmt(-accum - p.dep.end) + ').' });
    }
  }

  // Equity: prior year end balances + monthly activity through the report month.
  const pye = (Number(m.year) - 1) + '-12-31';
  const pyeBal = new Map();
  for (const r of (computeBalances(eid, { as_of: pye, close_pl_before: m.yearStart }) || [])) pyeBal.set(String(r.code), r2(r.balance));
  const eqMonthly = new Map(); // code -> [activity m1..mn]
  const eqCodes = new Set((byCat.equity || []).map((a) => String(a.code)));
  if (eqCodes.size) {
    const rows = db.prepare(`SELECT jl.account_code code, CAST(strftime('%m', je.date) AS INTEGER) mo, SUM(jl.credit - jl.debit) amt
      FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
      WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? GROUP BY jl.account_code, mo`).all(eid, m.yearStart, m.end);
    for (const r of rows) { const c = String(r.code); if (!eqCodes.has(c)) continue; if (!eqMonthly.has(c)) eqMonthly.set(c, new Array(m.monthNum).fill(0)); if (r.mo >= 1 && r.mo <= m.monthNum) eqMonthly.get(c)[r.mo - 1] = r2(r.amt); }
  }

  return Object.assign({}, base, { flags, byCat, arCur, arPrior, arPrimary, apCur, apPrior, apPrimary, reg, prepaidByAcct, prepaidSched, faSched, fixedPairs, pye, pyeBal, eqMonthly });
}

// ─── Supporting tabs ──────────────────────────────────────────────────────────

// Cash: register balance per GL + reconciled balance per bank rec (no GL detail).
function buildCashRecTab(wb, a, en, m, used, leadTab) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 52; ws.getColumn('C').width = 18; ws.getColumn('D').width = 3; ws.getColumn('E').width = 40;
  tabHead(ws, a.code + ' - ' + a.name, en, 'Bank reconciliation — month ended ' + spell(m.end), leadTab);
  txt(ws, 'B5', 'Register balance per GL — ' + short(m.end), { font: { bold: true } });
  num(ws, 'C5', a.end, { font: { bold: true }, border: { bottom: THIN } });
  txt(ws, 'B7', 'Balance per bank statement — ' + short(m.end) + ' (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C7');
  txt(ws, 'B8', 'Less: outstanding checks / payments not yet cleared (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C8');
  txt(ws, 'B9', 'Add: deposits in transit (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C9');
  txt(ws, 'B10', 'Reconciled balance per bank rec', { font: { bold: true } });
  totalCell(ws, 'C10', 'C7-C8+C9');
  txt(ws, 'B12', 'Difference (register per GL − reconciled per bank rec)', { font: { italic: true } });
  num(ws, 'C12', { formula: 'C5-C10' }, { font: { italic: true } });
  txt(ws, 'E7', 'Paste / attach the bank statement and the bank rec support here.', { font: { italic: true, color: { argb: 'FF7F7F7F' } } });
  return { sheet, registerRef: ref(sheet, 'C5'), clearedRef: ref(sheet, 'C10') };
}

// Credit card: statement-date reconciliation rolled to month end.
function buildCcRecTab(wb, a, en, m, used, leadTab) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 56; ws.getColumn('C').width = 18; ws.getColumn('D').width = 3; ws.getColumn('E').width = 40;
  tabHead(ws, a.code + ' - ' + a.name, en, 'Credit card reconciliation — month ended ' + spell(m.end), leadTab);
  txt(ws, 'B5', 'Statement end date (enter)', { font: { color: { argb: BLUE } } }); { const c = blueInput(ws, 'C5'); c.numFmt = DATEFMT; }
  txt(ws, 'B6', 'Month end date'); num(ws, 'C6', asDate(m.end), { fmt: DATEFMT });
  txt(ws, 'B8', 'Statement ending balance as of statement date (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C8');
  txt(ws, 'B9', 'Register balance as of statement date (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C9');
  txt(ws, 'B10', 'Uncleared transactions as of statement date (enter)', { font: { color: { argb: BLUE } } }); blueInput(ws, 'C10');
  txt(ws, 'B11', 'Reconciled balance as of statement date', { font: { bold: true } }); totalCell(ws, 'C11', 'C9+C10');
  txt(ws, 'B12', 'Difference (statement − reconciled)', { font: { italic: true } }); num(ws, 'C12', { formula: 'C8-C11' }, { font: { italic: true } });
  txt(ws, 'B14', 'Roll forward to month end', { font: { bold: true } });
  txt(ws, 'B15', 'Register balance per GL — ' + short(m.end) + ' (period end)', { font: { bold: true } });
  num(ws, 'C15', a.end, { font: { bold: true }, border: { top: THIN, bottom: DBL } });
  txt(ws, 'E8', 'Paste / attach the credit card statement here.', { font: { italic: true, color: { argb: 'FF7F7F7F' } } });
  return { sheet, clearedRef: ref(sheet, 'C8'), registerRef: ref(sheet, 'C11'), endingRef: ref(sheet, 'C15') };
}

// A/R Aging + A/R Summary. Returns { agingTab, totalRef, summaryTab }.
function buildArTabs(wb, data, en, m, used, leadTab) {
  const { arCur, arPrior } = data;
  const agingTab = reserveName('AR Aging', used);
  const ws = wb.addWorksheet(agingTab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 34; ws.getColumn('C').width = 22; ws.getColumn('D').width = 12; ws.getColumn('E').width = 12; ws.getColumn('F').width = 10;
  ['G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach((c) => ws.getColumn(c).width = 14);
  tabHead(ws, 'A/R Aging Detail', en, 'As of ' + short(m.end) + ' — open invoices by customer, aged from the due date (CloudLedger invoice subledger)', leadTab);
  const HR = 5;
  hdr(ws, HR, 2, ['Customer', 'Invoice #', 'Invoice Date', 'Due Date', 'Days Past Due', 'Current', '1-30', '31-60', '61-90', '90+', 'GL (un-aged)', 'Amount']);
  let r = HR + 1; const first = r;
  const bcol = { current: 'G', d1_30: 'H', d31_60: 'I', d61_90: 'J', d90_plus: 'K' };
  if (arCur) {
    const det = arCur.detail.slice().sort((x, y) => String(x.customer).localeCompare(String(y.customer)) || String(x.invoice_date || '').localeCompare(String(y.invoice_date || '')));
    for (const d of det) {
      txt(ws, 'B' + r, d.customer); txt(ws, 'C' + r, d.invoice_num || '');
      num(ws, 'D' + r, asDate(d.invoice_date), { fmt: DATEFMT }); num(ws, 'E' + r, asDate(d.due_date), { fmt: DATEFMT });
      num(ws, 'F' + r, d.days_past_due, { fmt: '0' });
      num(ws, bcol[d.bucket] + r, d.open);
      num(ws, 'M' + r, { formula: 'SUM(G' + r + ':L' + r + ')' });
      r++;
    }
    for (const g of arCur.gl_rows || []) {
      txt(ws, 'B' + r, (g.memo || 'GL entry') + (g.entry_num ? ' (JE ' + g.entry_num + ')' : ''), { font: { italic: true } });
      num(ws, 'D' + r, asDate(g.date), { fmt: DATEFMT });
      num(ws, 'L' + r, g.amount); num(ws, 'M' + r, { formula: 'SUM(G' + r + ':L' + r + ')' });
      r++;
    }
  }
  const last = r - 1;
  txt(ws, 'B' + r, 'TOTAL', { font: { bold: true } });
  ['G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach((c) => totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null));
  const totRow = r; r += 2;
  txt(ws, 'B' + r, 'Balance per GL — A/R account(s) ' + (arCur ? arCur.ar_accounts.join(', ') : '') + ' — ' + short(m.end), { font: { bold: true } });
  num(ws, 'M' + r, arCur ? arCur.gl_ar_balance : 0, { font: { bold: true } }); const glRow = r; r++;
  txt(ws, 'B' + r, 'Difference (aging − GL)', { font: { italic: true } }); num(ws, 'M' + r, { formula: 'M' + totRow + '-M' + glRow }, { font: { italic: true } });
  const totalRef = ref(agingTab, 'M' + totRow);

  const summaryTab = reserveName('AR Summary', used);
  const su = wb.addWorksheet(summaryTab, { views: [{ showGridLines: false }] });
  const mapRows = (ag) => (ag ? ag.rows.map((x) => ({ name: x.customer, current: x.current, b1: x.d1_30, b2: x.d31_60, b3: x.d61_90, b4: x.d90_plus, total: x.total })) : []);
  buildAgingSummary(su, en, m, 'Customer', 'A/R Aging Summary — prior month vs current month',
    { rows: mapRows(arPrior), gl: arPrior ? r2(arPrior.gl_total) : 0 }, { rows: mapRows(arCur), gl: arCur ? r2(arCur.gl_total) : 0 },
    ['Current', '1-30', '31-60', '61-90', '90+'], leadTab);
  return { agingTab, totalRef, summaryTab };
}

// A/P Aging + A/P Summary. Returns { agingTab, totalRef, summaryTab }.
function buildApTabs(wb, data, en, m, used, leadTab) {
  const { apCur, apPrior } = data;
  const agingTab = reserveName('AP Aging', used);
  const ws = wb.addWorksheet(agingTab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 8; ws.getColumn('D').width = 16; ws.getColumn('E').width = 32; ws.getColumn('F').width = 12; ws.getColumn('G').width = 10;
  ['H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach((c) => ws.getColumn(c).width = 14);
  tabHead(ws, 'A/P Aging Detail', en, 'As of ' + short(m.end) + ' — open bills by vendor, aged from the bill date (CloudLedger / Bill.com)', leadTab);
  const HR = 5;
  hdr(ws, HR, 2, ['Date', 'Type', 'Num', 'Vendor', 'Due Date', 'Past Due (days)', 'Current', '1-30', '31-60', '61-90', '91+', 'GL (un-aged)', 'Amount']);
  let r = HR + 1;
  const bcol = { current: 'H', d1_30: 'I', d31_60: 'J', d61_90: 'K', d91_plus: 'L' };
  const subtotalRows = [];
  if (apCur) {
    for (const g of apCur.vendors || []) {
      const gFirst = r;
      for (const b of g.rows) {
        num(ws, 'B' + r, asDate(b.date), { fmt: DATEFMT }); txt(ws, 'C' + r, b.type || 'Bill'); txt(ws, 'D' + r, b.num || '');
        txt(ws, 'E' + r, b.vendor); num(ws, 'F' + r, asDate(b.due_date), { fmt: DATEFMT }); num(ws, 'G' + r, b.past_due_days, { fmt: '0' });
        num(ws, bcol[b.bucket] + r, b.amount); num(ws, 'N' + r, { formula: 'SUM(H' + r + ':M' + r + ')' });
        r++;
      }
      const gLast = r - 1;
      txt(ws, 'B' + r, 'Total ' + g.vendor, { font: { bold: true } });
      ['H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach((c) => { const cell = num(ws, c + r, gLast >= gFirst ? { formula: 'SUM(' + c + gFirst + ':' + c + gLast + ')' } : 0, { font: { bold: true } }); cell.border = { top: THIN }; });
      subtotalRows.push(r); r++;
    }
    for (const g of apCur.gl_rows || []) {
      num(ws, 'B' + r, asDate(g.date), { fmt: DATEFMT }); txt(ws, 'C' + r, 'GL'); txt(ws, 'D' + r, g.entry_num || '');
      txt(ws, 'E' + r, g.memo || 'GL entry', { font: { italic: true } });
      num(ws, 'M' + r, g.amount); num(ws, 'N' + r, { formula: 'SUM(H' + r + ':M' + r + ')' });
      subtotalRows.push(r); r++;
    }
  }
  txt(ws, 'B' + r, 'TOTAL', { font: { bold: true } });
  ['H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach((c) => totalCell(ws, c + r, subtotalRows.length ? subtotalRows.map((sr) => c + sr).join('+') : null));
  const totRow = r; r += 2;
  txt(ws, 'B' + r, 'Balance per GL — A/P account ' + (apCur ? apCur.ap_account : '') + ' — ' + short(m.end), { font: { bold: true } });
  num(ws, 'N' + r, apCur ? r2(apCur.gl_balance) : 0, { font: { bold: true } }); const glRow = r; r++;
  txt(ws, 'B' + r, 'Difference (aging − GL)', { font: { italic: true } }); num(ws, 'N' + r, { formula: 'N' + totRow + '-N' + glRow }, { font: { italic: true } });
  const totalRef = ref(agingTab, 'N' + totRow);

  const summaryTab = reserveName('AP Summary', used);
  const su = wb.addWorksheet(summaryTab, { views: [{ showGridLines: false }] });
  const rowsOf = (ap) => (ap ? (ap.vendors || []).map((g) => ({ name: g.vendor, current: g.subtotal.current, b1: g.subtotal.d1_30, b2: g.subtotal.d31_60, b3: g.subtotal.d61_90, b4: g.subtotal.d91_plus, total: g.subtotal.total })) : []);
  buildAgingSummary(su, en, m, 'Vendor', 'A/P Aging Summary — prior month vs current month',
    { rows: rowsOf(apPrior), gl: apPrior ? r2(apPrior.gl_total) : 0 }, { rows: rowsOf(apCur), gl: apCur ? r2(apCur.gl_total) : 0 },
    ['Current', '1-30', '31-60', '61-90', '91+'], leadTab);
  return { agingTab, totalRef, summaryTab };
}

// Prior-vs-current aging summary (CLA "Summary" tab): left block prior month,
// right block current month, Difference column.
function buildAgingSummary(ws, en, m, nameLabel, title, prior, cur, bucketLabels, leadTab) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 34;
  for (const c of ['C', 'D', 'E', 'F', 'G', 'H', 'I']) ws.getColumn(c).width = 13;
  ws.getColumn('J').width = 3;
  for (const c of ['K', 'L', 'M', 'N', 'O', 'P', 'Q']) ws.getColumn(c).width = 13;
  ws.getColumn('R').width = 3; ws.getColumn('S').width = 14;
  tabHead(ws, title, en, 'Prior month as of ' + short(m.beg) + '  |  Current month as of ' + short(m.end), leadTab);
  const R0 = 5;
  const sec = (addr, t) => { const c = ws.getCell(addr); c.value = t; c.font = F({ bold: true }); c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTFILL } }; };
  sec('C' + R0, 'Prior month — as of ' + short(m.beg)); sec('K' + R0, 'Current month — as of ' + short(m.end));
  const HR = R0 + 1;
  hdr(ws, HR, 2, [nameLabel].concat(bucketLabels, ['GL (un-aged)', 'Total']));
  hdr(ws, HR, 11, bucketLabels.concat(['GL (un-aged)', 'Total']));
  hdr(ws, HR, 19, ['Difference']);
  const names = new Map();
  for (const x of prior.rows) names.set(x.name, { p: x, c: null });
  for (const x of cur.rows) { if (!names.has(x.name)) names.set(x.name, { p: null, c: x }); else names.get(x.name).c = x; }
  const list = Array.from(names.entries()).sort((a, b) => (Math.abs((b[1].c || {}).total || 0) - Math.abs((a[1].c || {}).total || 0)) || a[0].localeCompare(b[0]));
  let r = HR + 1; const first = r;
  const put = (row, col0, x) => {
    const vals = x ? [x.current, x.b1, x.b2, x.b3, x.b4] : [0, 0, 0, 0, 0];
    vals.forEach((v, i) => num(ws, colL(col0 + i) + row, v));
    num(ws, colL(col0 + 5) + row, 0);
    num(ws, colL(col0 + 6) + row, { formula: 'SUM(' + colL(col0) + row + ':' + colL(col0 + 5) + row + ')' });
  };
  for (const [name, v] of list) {
    txt(ws, 'B' + r, name); put(r, 3, v.p); put(r, 11, v.c);
    num(ws, 'S' + r, { formula: 'Q' + r + '-I' + r });
    r++;
  }
  txt(ws, 'B' + r, 'Un-aged GL entries', { font: { italic: true } });
  [3, 4, 5, 6, 7].forEach((c) => num(ws, colL(c) + r, 0)); num(ws, 'H' + r, prior.gl); num(ws, 'I' + r, { formula: 'SUM(C' + r + ':H' + r + ')' });
  [11, 12, 13, 14, 15].forEach((c) => num(ws, colL(c) + r, 0)); num(ws, 'P' + r, cur.gl); num(ws, 'Q' + r, { formula: 'SUM(K' + r + ':P' + r + ')' });
  num(ws, 'S' + r, { formula: 'Q' + r + '-I' + r }); r++;
  const last = r - 1;
  txt(ws, 'B' + r, 'Grand totals', { font: { bold: true } });
  for (const c of ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'S']) totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null);
}

// Prepaid amortization schedule for one prepaid account (CLA layout).
function buildPrepaidTab(wb, a, items, en, m, used, leadTab) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  const n = m.monthNum; const ends = fyMonthEnds(m); const pye = (Number(m.year) - 1) + '-12-31';
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 11; ws.getColumn('C').width = 20; ws.getColumn('D').width = 10; ws.getColumn('E').width = 40;
  ws.getColumn('F').width = 11; ws.getColumn('G').width = 11; ws.getColumn('H').width = 12; ws.getColumn('I').width = 13;
  for (let k = 1; k <= n; k++) for (let j = 0; j < 3; j++) ws.getColumn(colL(10 + 3 * (k - 1) + j)).width = 13;
  tabHead(ws, a.code + ' - ' + a.name, en, 'Prepaid amortization schedule — fiscal ' + m.year + ' through ' + spell(m.end), leadTab);
  txt(ws, 'F10', 'Period covered', { font: { bold: true } });
  hdrDates(ws, 10, 9, [pye]);
  for (let k = 1; k <= n; k++) hdrDates(ws, 10, 12 + 3 * (k - 1), [ends[k - 1]]);
  const heads = ['Date Paid', 'Vendor', 'Expense Acct', 'Description', 'Start Date', 'End Date', 'Monthly', 'Balance P0'];
  for (let k = 1; k <= n; k++) heads.push('Additions P' + k, 'Expenses P' + k, 'Balance P' + k);
  hdr(ws, 11, 2, heads);
  let r = 12; const first = r;
  for (const it of items) {
    num(ws, 'B' + r, asDate(it.date_paid), { fmt: DATEFMT }); txt(ws, 'C' + r, it.vendor || ''); txt(ws, 'D' + r, it.expense_account || '');
    txt(ws, 'E' + r, it.description || ''); num(ws, 'F' + r, asDate(it.start_date), { fmt: DATEFMT }); num(ws, 'G' + r, asDate(it.end_date), { fmt: DATEFMT });
    num(ws, 'H' + r, r2(it.monthly)); num(ws, 'I' + r, r2(it.opening_balance));
    for (let k = 1; k <= n; k++) {
      const addC = colL(10 + 3 * (k - 1)), expC = colL(11 + 3 * (k - 1)), balC = colL(12 + 3 * (k - 1));
      const prev = k === 1 ? 'I' + r : colL(12 + 3 * (k - 2)) + r;
      const add = (it.premium != null && it.date_paid && ym(it.date_paid) === ym(ends[k - 1])) ? r2(it.premium) : 0;
      num(ws, addC + r, add);
      num(ws, expC + r, { formula: 'IF(AND($F' + r + '<>"",$H' + r + '>0,' + balC + '$10>=EOMONTH($F' + r + ',1),(' + prev + '+' + addC + r + ')>0.005),-MIN($H' + r + ',' + prev + '+' + addC + r + '),0)' });
      num(ws, balC + r, { formula: prev + '+' + addC + r + '+' + expC + r });
    }
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, 'Total ' + a.name, { font: { bold: true } });
  const totCols = ['H', 'I']; for (let k = 1; k <= n; k++) totCols.push(colL(10 + 3 * (k - 1)), colL(11 + 3 * (k - 1)), colL(12 + 3 * (k - 1)));
  for (const c of totCols) totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null);
  const totRow = r; const balN = colL(12 + 3 * (n - 1)); const expN = colL(11 + 3 * (n - 1));
  r += 2;
  txt(ws, 'E' + r, 'Balance per GL — ' + a.code + ' — ' + short(m.end), { font: { bold: true } }); num(ws, balN + r, a.end, { font: { bold: true } }); const glRow = r; r++;
  txt(ws, 'E' + r, 'Difference (schedule − GL)', { font: { italic: true } }); num(ws, balN + r, { formula: balN + totRow + '-' + balN + glRow }, { font: { italic: true } }); r += 2;
  txt(ws, 'B' + r, 'Journal entry — amortization for ' + MONTHS[m.monthNum - 1] + ' ' + m.year, { font: { bold: true } }); r++;
  hdr(ws, r, 2, ['Date', 'Account', 'Debit', 'Credit']); r++;
  const expAcct = items.length ? (items[0].expense_account || '') : '';
  num(ws, 'B' + r, asDate(m.end), { fmt: DATEFMT }); txt(ws, 'C' + r, String(expAcct) + ' — expense'); num(ws, 'D' + r, { formula: '-' + expN + totRow }); r++;
  txt(ws, 'C' + r, a.code + ' — ' + a.name); num(ws, 'E' + r, { formula: '-' + expN + totRow }); r++;
  return { sheet, endRef: ref(sheet, balN + totRow) };
}

// Fixed Asset Depreciation Schedule (one tab, grouped by asset account).
// Returns { tab, refs: Map(assetCode -> {costRef}, depCode -> {accumRef}) }.
function buildFixedSchedule(wb, data, en, m, used, leadTab) {
  const tab = reserveName('Fixed Asset Schedule', used);
  const ws = wb.addWorksheet(tab, { views: [{ showGridLines: false }] });
  const n = m.monthNum; const ends = fyMonthEnds(m); const pye = (Number(m.year) - 1) + '-12-31';
  const MC0 = 9; // first month column (I)
  const accumC = colL(MC0 + n), nbvC = colL(MC0 + n + 1), lastMC = colL(MC0 + n - 1);
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 46; ws.getColumn('C').width = 9; ws.getColumn('D').width = 11; ws.getColumn('E').width = 11;
  ws.getColumn('F').width = 13; ws.getColumn('G').width = 11; ws.getColumn('H').width = 13;
  for (let k = 0; k < n + 2; k++) ws.getColumn(colL(MC0 + k)).width = 12;
  tabHead(ws, 'Fixed Asset Depreciation Schedule', en, 'Fiscal ' + m.year + ' through ' + spell(m.end) + ' — straight-line, monthly, from the in-service month', leadTab);
  txt(ws, 'H7', 'Accumulated', { font: { bold: true }, align: { horizontal: 'center' } });
  txt(ws, colL(MC0) + '7', 'Depreciation / Amortization Expense', { font: { bold: true } });
  txt(ws, accumC + '7', 'Accumulated', { font: { bold: true }, align: { horizontal: 'center' } });
  txt(ws, 'H8', 'as of ' + short(pye), { font: { italic: true }, align: { horizontal: 'center' } });
  hdr(ws, 9, 2, ['Description', 'Useful Life (Yrs)', 'Date In Service', 'Depr/Amort End Date', 'Cost', 'Monthly Depr/Amort', 'Depr/Amort Beg Balance']);
  hdrDates(ws, 9, MC0, ends);
  hdr(ws, 9, MC0 + n, ['Depr/Amort End Balance', 'NBV']);
  const out = new Map();
  let r = 10;
  for (const p of data.fixedPairs) {
    const acct = p.asset || p.dep; if (!acct) continue;
    const assets = data.reg.fixedAssets.filter((x) => p.asset ? String(x.asset_account) === String(p.asset.code) : String(x.dep_account) === String(p.dep.code));
    const gh = ws.getCell('B' + r); gh.value = acct.code + '  ' + acct.name + (p.asset && p.dep ? '   (accumulated: ' + p.dep.code + ' ' + p.dep.name + ')' : ''); gh.font = F({ bold: true });
    for (let c = 2; c <= MC0 + n + 1; c++) ws.getRow(r).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRPFILL } };
    r++;
    const gFirst = r;
    for (const fa of assets) {
      txt(ws, 'B' + r, fa.description || ''); num(ws, 'C' + r, Number(fa.life_years) || 0, { fmt: '0.00' });
      num(ws, 'D' + r, asDate(fa.in_service), { fmt: DATEFMT });
      num(ws, 'E' + r, { formula: 'IF(AND($D' + r + '<>"",$C' + r + '>0),EDATE($D' + r + ',ROUND($C' + r + '*12,0))-1,"")' }, { fmt: DATEFMT });
      num(ws, 'F' + r, r2(fa.cost));
      num(ws, 'G' + r, { formula: 'IF($C' + r + '>0,ROUND($F' + r + '/($C' + r + '*12),2),0)' });
      num(ws, 'H' + r, r2(fa.accum_dep_beg));
      for (let k = 1; k <= n; k++) {
        const mc = colL(MC0 + k - 1);
        const priorSum = k === 1 ? '0' : 'SUM($' + colL(MC0) + r + ':' + colL(MC0 + k - 2) + r + ')';
        num(ws, mc + r, { formula: 'IF(AND($D' + r + '<>"",' + mc + '$9>=$D' + r + ',OR($E' + r + '="",' + mc + '$9<=$E' + r + ')),MIN($G' + r + ',MAX(0,$F' + r + '-$H' + r + '-' + priorSum + ')),0)' });
      }
      num(ws, accumC + r, { formula: '$H' + r + '+SUM(' + colL(MC0) + r + ':' + lastMC + r + ')' });
      num(ws, nbvC + r, { formula: '$F' + r + '-' + accumC + r });
      r++;
    }
    const gLast = r - 1;
    txt(ws, 'B' + r, 'Total ' + acct.name, { font: { bold: true } });
    const cols = ['F', 'H']; for (let k = 0; k < n; k++) cols.push(colL(MC0 + k)); cols.push(accumC, nbvC);
    for (const c of cols) totalCell(ws, c + r, gLast >= gFirst ? 'SUM(' + c + gFirst + ':' + c + gLast + ')' : null);
    if (p.asset) out.set(String(p.asset.code), { costRef: ref(tab, 'F' + r), monthRef: ref(tab, lastMC + r) });
    if (p.dep) out.set(String(p.dep.code), { accumRef: ref(tab, accumC + r), monthRef: ref(tab, lastMC + r) });
    r += 2;
  }
  txt(ws, 'B' + r, 'Agreement to the general ledger — ' + short(m.end), { font: { bold: true, size: 12 } }); r++;
  hdr(ws, r, 2, ['Account', 'Per schedule', 'Per GL', 'Difference']); r++;
  for (const p of data.fixedPairs) {
    if (p.asset) { const o = out.get(String(p.asset.code)); txt(ws, 'B' + r, p.asset.code + ' ' + p.asset.name + ' (cost)'); num(ws, 'C' + r, o ? { formula: o.costRef } : 0); num(ws, 'D' + r, p.asset.end); num(ws, 'E' + r, { formula: 'C' + r + '-D' + r }); r++; }
    if (p.dep) { const o = out.get(String(p.dep.code)); txt(ws, 'B' + r, p.dep.code + ' ' + p.dep.name + ' (accumulated)'); num(ws, 'C' + r, o ? { formula: '-' + o.accumRef } : 0); num(ws, 'D' + r, p.dep.end); num(ws, 'E' + r, { formula: 'C' + r + '-D' + r }); r++; }
  }
  r++;
  txt(ws, 'B' + r, 'Journal entry — depreciation / amortization for ' + MONTHS[m.monthNum - 1] + ' ' + m.year, { font: { bold: true } }); r++;
  hdr(ws, r, 2, ['Account', 'Debit', 'Credit']); r++;
  for (const p of data.fixedPairs) {
    if (!p.dep) continue;
    const o = out.get(String(p.dep.code)); if (!o) continue;
    txt(ws, 'B' + r, 'Depreciation / amortization expense — ' + (p.asset ? p.asset.name : p.dep.name)); num(ws, 'C' + r, { formula: o.monthRef }); r++;
    txt(ws, 'B' + r, '     ' + p.dep.code + ' ' + p.dep.name); num(ws, 'D' + r, { formula: o.monthRef }); r++;
  }
  return { tab, refs: out };
}

// Category schedule: Beginning (prior month end) → Activity → Ending, one row per account.
function buildCategorySchedule(wb, cd, rows, en, m, used, leadTab) {
  const tab = reserveName(cd.sched || (cd.label + ' Schedule'), used);
  const ws = wb.addWorksheet(tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 44; ws.getColumn('D').width = 16; ws.getColumn('E').width = 16; ws.getColumn('F').width = 16; ws.getColumn('G').width = 36;
  tabHead(ws, cd.label + ' — Schedule', en, 'Roll-forward by account — ' + short(m.beg) + ' to ' + short(m.end) + ' (per general ledger)', leadTab);
  const HR = 5;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Beginning ' + short(m.beg), 'Activity (net)', 'Ending ' + short(m.end), 'Comments']);
  let r = HR + 1; const first = r; const refs = new Map();
  for (const a of rows) {
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    num(ws, 'D' + r, a.begin); num(ws, 'E' + r, a.activity); num(ws, 'F' + r, { formula: 'D' + r + '+E' + r });
    refs.set(String(a.code), { sheet: tab, endRef: ref(tab, 'F' + r) });
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  for (const c of ['D', 'E', 'F']) totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null);
  return { tab, refs };
}

// Roll block appended to a tab (for secondary A/R and A/P accounts). Returns refs.
function rollBlock(ws, startRow, title, rows, m, tab) {
  let r = startRow; const refs = new Map();
  txt(ws, 'B' + r, title, { font: { bold: true, size: 12 } }); r++;
  hdr(ws, r, 2, ['Account No.', 'Account Name', 'Beginning ' + short(m.beg), 'Activity (net)', 'Ending ' + short(m.end)]); r++;
  for (const a of rows) {
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    num(ws, 'D' + r, a.begin); num(ws, 'E' + r, a.activity); num(ws, 'F' + r, { formula: 'D' + r + '+E' + r });
    refs.set(String(a.code), { sheet: tab, endRef: ref(tab, 'F' + r) }); r++;
  }
  return { refs, nextRow: r };
}

// Intercompany Tie Out: our balance (Dr = receivable) vs the counterparty's ledger.
function buildTieOut(wb, rows, en, m, used, leadTab) {
  const tab = reserveName('Tie Out', used);
  const ws = wb.addWorksheet(tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 36; ws.getColumn('D').width = 32; ws.getColumn('E').width = 17; ws.getColumn('F').width = 17; ws.getColumn('G').width = 15; ws.getColumn('H').width = 40;
  tabHead(ws, 'Intercompany Tie Out', en, 'Month ended ' + spell(m.end) + ' — Debit = receivable, Credit = payable; counterparty balances per their own CloudLedger ledger', leadTab);
  const HR = 5;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Counterparty Entity', 'Our Balance', 'Per Counterparty GL', 'Variance', 'Comments']);
  let r = HR + 1; const first = r; const refs = new Map();
  for (const a of rows) {
    const ours = a.type === 'Asset' ? a.end : -a.end;
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    txt(ws, 'D' + r, a.cp ? a.cp.name : '');
    num(ws, 'E' + r, ours);
    if (a.cp) { num(ws, 'F' + r, a.cpNet || 0); num(ws, 'G' + r, { formula: 'E' + r + '-F' + r }); }
    else { num(ws, 'F' + r, ''); num(ws, 'G' + r, ''); if (Math.abs(a.end) >= 0.01) txt(ws, 'H' + r, 'No matching CloudLedger entity — confirm the counterparty balance manually', { font: { italic: true, color: { argb: 'FF9C4221' } } }); }
    refs.set(String(a.code), { sheet: tab, ourRef: ref(tab, 'E' + r), cpRef: a.cp ? ref(tab, 'F' + r) : null });
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, 'Total Intercompany', { font: { bold: true } });
  for (const c of ['E', 'F', 'G']) totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null);
  return { tab, refs };
}

// Equity Rollforward: prior year end → monthly activity → ending, by account.
function buildEquityRollforward(wb, rows, data, en, m, used, leadTab, niRef) {
  const tab = reserveName('Equity Rollforward', used);
  const ws = wb.addWorksheet(tab, { views: [{ showGridLines: false }] });
  const n = m.monthNum; const endC = colL(5 + n);
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 40; ws.getColumn('D').width = 16;
  for (let k = 0; k < n; k++) ws.getColumn(colL(5 + k)).width = 14;
  ws.getColumn(endC).width = 16;
  tabHead(ws, 'Equity Rollforward', en, 'Fiscal ' + m.year + ' through ' + spell(m.end) + ' — by equity account, per general ledger', leadTab);
  const HR = 5;
  const heads = ['Account No.', 'Account Name', 'Prior Year End ' + short(data.pye)];
  for (let k = 0; k < n; k++) heads.push(MON3[k] + ' ' + m.year);
  heads.push('Ending ' + short(m.end));
  hdr(ws, HR, 2, heads);
  let r = HR + 1; const first = r; const refs = new Map();
  for (const a of rows) {
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    num(ws, 'D' + r, data.pyeBal.get(String(a.code)) || 0);
    const act = data.eqMonthly.get(String(a.code)) || new Array(n).fill(0);
    for (let k = 0; k < n; k++) num(ws, colL(5 + k) + r, act[k] || 0);
    num(ws, endC + r, { formula: 'D' + r + '+SUM(' + colL(5) + r + ':' + colL(4 + n) + r + ')' });
    refs.set(String(a.code), { sheet: tab, endRef: ref(tab, endC + r) });
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, 'Total equity accounts', { font: { bold: true } });
  const cols = ['D']; for (let k = 0; k < n; k++) cols.push(colL(5 + k)); cols.push(endC);
  for (const c of cols) totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null);
  const totRow = r; r += 2;
  txt(ws, 'C' + r, 'Net income (loss) — fiscal YTD (per Income Statement)'); num(ws, endC + r, { formula: niRef }); const niRow = r; r++;
  txt(ws, 'C' + r, 'Total equity including current-year earnings', { font: { bold: true } }); totalCell(ws, endC + r, endC + totRow + '+' + endC + niRow);
  return { tab, refs };
}

function buildIncomeStatement(wb, en, m, ni, used) {
  const tab = reserveName('Income Statement', used);
  const pl = wb.addWorksheet(tab, { views: [{ showGridLines: false }] });
  pl.getColumn('A').width = 3.4; pl.getColumn('B').width = 14; pl.getColumn('C').width = 52; pl.getColumn('D').width = 16;
  pl.getCell('C1').value = en; pl.getCell('C1').font = F({ size: 16, bold: true });
  pl.getCell('C2').value = 'STATEMENT OF OPERATIONS — FISCAL YEAR TO DATE'; pl.getCell('C2').font = F({ size: 12, bold: true });
  pl.getCell('C3').value = m.yearStart + ' to ' + short(m.end); pl.getCell('C3').font = F({ italic: true });
  const PHR = 5; hdr(pl, PHR, 2, ['Code', 'Account', 'Amount']);
  let pr = PHR + 1;
  txt(pl, 'B' + pr, 'REVENUE', { font: { bold: true } }); pr++;
  const revFirst = pr;
  for (const x of ni.end.revenue) { txt(pl, 'B' + pr, x.code); txt(pl, 'C' + pr, x.name); num(pl, 'D' + pr, x.amt); pr++; }
  const revLast = pr - 1;
  txt(pl, 'C' + pr, 'Total revenue', { font: { bold: true } });
  const totRev = 'D' + pr; { const c = num(pl, totRev, revLast >= revFirst ? { formula: 'SUM(D' + revFirst + ':D' + revLast + ')' } : 0, { font: { bold: true } }); c.border = { top: THIN }; }
  pr += 2;
  txt(pl, 'B' + pr, 'EXPENSES', { font: { bold: true } }); pr++;
  const expFirst = pr;
  for (const x of ni.end.expense) { txt(pl, 'B' + pr, x.code); txt(pl, 'C' + pr, x.name); num(pl, 'D' + pr, x.amt); pr++; }
  const expLast = pr - 1;
  txt(pl, 'C' + pr, 'Total expenses', { font: { bold: true } });
  const totExp = 'D' + pr; { const c = num(pl, totExp, expLast >= expFirst ? { formula: 'SUM(D' + expFirst + ':D' + expLast + ')' } : 0, { font: { bold: true } }); c.border = { top: THIN }; }
  pr += 2;
  txt(pl, 'C' + pr, 'NET INCOME (LOSS) — fiscal YTD', { font: { bold: true } });
  totalCell(pl, 'D' + pr, totRev + '-' + totExp);
  return ref(tab, 'D' + pr);
}

// ─── Leadsheets (CLA columns) ─────────────────────────────────────────────────
function simpleLeadsheet(ws, cd, en, m, rows, refByCode, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 40;
  ws.getColumn('D').width = 16; ws.getColumn('E').width = 16; ws.getColumn('F').width = 40;
  titleBlock(ws, en, cd.lead, m, 'G');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(String(a.code)) || {};
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    wpLink(ws, 'D' + r, rr.sheet, rr.sheet || '');
    num(ws, 'E' + r, rr.endRef ? { formula: rr.endRef } : a.end);
    leadCell.set(String(a.code), ref(cd.tab, 'E' + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  totalCell(ws, 'E' + r, last >= first ? 'SUBTOTAL(109,E' + first + ':E' + last + ')' : null);
  return ref(cd.tab, 'E' + r);
}
function cashLeadsheet(ws, cd, en, m, rows, refByCode, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 38;
  ws.getColumn('D').width = 12; ws.getColumn('E').width = 16; ws.getColumn('F').width = 16; ws.getColumn('G').width = 18; ws.getColumn('H').width = 30;
  titleBlock(ws, en, cd.lead, m, 'H');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet / Location', 'Cleared Balance', 'Register Balance', 'Bank Statement Ref', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(String(a.code)) || {};
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    num(ws, 'E' + r, rr.clearedRef ? { formula: rr.clearedRef } : '');
    num(ws, 'F' + r, rr.registerRef ? { formula: rr.registerRef } : a.end);
    txt(ws, 'G' + r, 'Bank Statement');
    leadCell.set(String(a.code), ref(cd.tab, 'F' + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  totalCell(ws, 'E' + r, last >= first ? 'SUBTOTAL(109,E' + first + ':E' + last + ')' : null);
  totalCell(ws, 'F' + r, last >= first ? 'SUBTOTAL(109,F' + first + ':F' + last + ')' : null);
  return ref(cd.tab, 'F' + r);
}
function ccLeadsheet(ws, cd, en, m, rows, refByCode, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 34;
  ws.getColumn('D').width = 12; ws.getColumn('E').width = 15; ws.getColumn('F').width = 15; ws.getColumn('G').width = 15; ws.getColumn('H').width = 16; ws.getColumn('I').width = 26;
  titleBlock(ws, en, cd.lead, m, 'I');
  txt(ws, 'E6', '(As of Statement Date)', { font: { italic: true, size: 9 } }); txt(ws, 'G6', '(As of Period-End)', { font: { italic: true, size: 9 } });
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet / Location', 'Cleared Balance', 'Register Balance', 'Ending Balance', 'CC Statement Ref', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(String(a.code)) || {};
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    num(ws, 'E' + r, rr.clearedRef ? { formula: rr.clearedRef } : '');
    num(ws, 'F' + r, rr.registerRef ? { formula: rr.registerRef } : '');
    num(ws, 'G' + r, rr.endingRef ? { formula: rr.endingRef } : a.end);
    txt(ws, 'H' + r, 'CC Statement');
    leadCell.set(String(a.code), ref(cd.tab, 'G' + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  for (const c of ['E', 'F', 'G']) totalCell(ws, c + r, last >= first ? 'SUBTOTAL(109,' + c + first + ':' + c + last + ')' : null);
  return ref(cd.tab, 'G' + r); // period-end (GL) total ties to the balance sheet
}
function fixedLeadsheet(ws, cd, en, m, pairs, faRefs, faTab, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 30; ws.getColumn('D').width = 15;
  ws.getColumn('E').width = 12; ws.getColumn('F').width = 30; ws.getColumn('G').width = 15; ws.getColumn('H').width = 15; ws.getColumn('I').width = 26;
  titleBlock(ws, en, cd.lead, m, 'J');
  txt(ws, 'G3', 'Capitalization Policy:'); txt(ws, 'H3', '>$1,000 for single item');
  const sB = ws.getCell('B7'); sB.value = 'Fixed Asset Account Breakdown'; sB.font = F({ bold: true }); sB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTFILL } };
  const sF = ws.getCell('E7'); sF.value = 'Depreciation Account Breakdown'; sF.font = F({ bold: true }); sF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTFILL } };
  hdr(ws, 8, 2, ['Asset Acct. #', 'Asset Acct. Name', 'Asset Cost', 'Depreciation Account #', 'Depreciation Account Name', 'Deprec. Balance', 'Net Asset Amount', 'Comments']);
  let r = 9; const first = r;
  for (const p of pairs) {
    if (p.asset) {
      txt(ws, 'B' + r, p.asset.code); wpLink(ws, 'C' + r, faTab, p.asset.name); ws.getCell('C' + r).alignment = { horizontal: 'left' };
      const o = faRefs.get(String(p.asset.code));
      num(ws, 'D' + r, o && o.costRef ? { formula: o.costRef } : p.asset.end);
      leadCell.set(String(p.asset.code), ref(cd.tab, 'D' + r));
    } else num(ws, 'D' + r, 0);
    if (p.dep) {
      txt(ws, 'E' + r, p.dep.code); txt(ws, 'F' + r, p.dep.name);
      const o = faRefs.get(String(p.dep.code));
      num(ws, 'G' + r, o && o.accumRef ? { formula: '-' + o.accumRef } : p.dep.end);
      leadCell.set(String(p.dep.code), ref(cd.tab, 'G' + r));
    } else num(ws, 'G' + r, 0);
    num(ws, 'H' + r, { formula: 'D' + r + '+G' + r });
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  for (const c of ['D', 'G', 'H']) totalCell(ws, c + r, last >= first ? 'SUBTOTAL(109,' + c + first + ':' + c + last + ')' : null);
  return ref(cd.tab, 'H' + r);
}
function otherAssetsLeadsheet(ws, cd, en, m, rows, refByCode, projects, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 44; ws.getColumn('D').width = 16; ws.getColumn('E').width = 16; ws.getColumn('F').width = 30;
  titleBlock(ws, en, cd.lead, m, 'G');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance\nDr (Cr)', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(String(a.code)) || {};
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    wpLink(ws, 'D' + r, rr.sheet, rr.sheet || '');
    num(ws, 'E' + r, rr.endRef ? { formula: rr.endRef } : a.end);
    leadCell.set(String(a.code), ref(cd.tab, 'E' + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  totalCell(ws, 'E' + r, last >= first ? 'SUM(E' + first + ':E' + last + ')' : null);
  if (projects && projects.length) {
    let pr = r + 3;
    txt(ws, 'C' + pr, 'Project Name', { font: { bold: true } }); txt(ws, 'D' + pr, 'Project Code', { font: { bold: true } }); pr++;
    for (const p of projects) { txt(ws, 'C' + pr, p.name); txt(ws, 'D' + pr, p.code); pr++; }
  }
  return ref(cd.tab, 'E' + r);
}
function intercoLeadsheet(ws, cd, en, m, rows, refByCode, leadCell) {
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 34; ws.getColumn('D').width = 22;
  ws.getColumn('E').width = 12; ws.getColumn('F').width = 15; ws.getColumn('G').width = 15; ws.getColumn('H').width = 15; ws.getColumn('I').width = 26;
  titleBlock(ws, en, cd.lead, m, 'I');
  txt(ws, 'C5', 'Entity Name:'); txt(ws, 'D5', en, { font: { bold: true } });
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Entity Name', 'Account Name', 'Worksheet', 'Balance', 'Other Entity Bal', 'Variance', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(String(a.code)) || {};
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } });
    txt(ws, 'C' + r, a.cp ? a.cp.name : String(a.name).replace(/^due\s+(to|from)\s+/i, ''));
    txt(ws, 'D' + r, a.name);
    wpLink(ws, 'E' + r, rr.sheet, rr.sheet || '');
    num(ws, 'F' + r, rr.ourRef ? { formula: rr.ourRef } : (a.type === 'Asset' ? a.end : -a.end));
    if (rr.cpRef) { num(ws, 'G' + r, { formula: rr.cpRef }); num(ws, 'H' + r, { formula: 'F' + r + '-G' + r }); }
    else { num(ws, 'G' + r, ''); num(ws, 'H' + r, ''); if (Math.abs(a.end) >= 0.01) txt(ws, 'I' + r, 'No matching CL entity — confirm manually', { font: { italic: true, color: { argb: 'FF9C4221' } } }); }
    leadCell.set(String(a.code), ref(cd.tab, 'F' + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  for (const c of ['F', 'G', 'H']) totalCell(ws, c + r, last >= first ? 'SUBTOTAL(109,' + c + first + ':' + c + last + ')' : null);
  return ref(cd.tab, 'F' + r);
}

// ─── Summary ──────────────────────────────────────────────────────────────────
function buildSummary(su, en, m, data, catTie, niRef, leadCell) {
  su.getColumn('A').width = 3.4; su.getColumn('B').width = 14; su.getColumn('C').width = 48; su.getColumn('D').width = 18; su.getColumn('E').width = 18; su.getColumn('F').width = 18; su.getColumn('G').width = 11;
  su.getCell('C1').value = en; su.getCell('C1').font = F({ size: 16, bold: true });
  su.getCell('C2').value = 'MONTHLY CLOSING WORKPAPERS — SUMMARY'; su.getCell('C2').font = F({ size: 12, bold: true });
  su.getCell('C3').value = 'Month ended ' + spell(m.end); su.getCell('C3').font = F({ italic: true });
  let sr = 5;
  const flags = data.flags || [];
  if (!flags.length) {
    su.mergeCells('B' + sr + ':G' + sr);
    const c = su.getCell('B' + sr); c.value = '✓  All checks passed — every schedule agrees to the general ledger and the balance sheet ties.'; c.font = F({ bold: true, color: { argb: 'FF1E7A34' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF7EE' } }; su.getRow(sr).height = 22; sr += 2;
  } else {
    su.mergeCells('B' + sr + ':G' + sr);
    const c = su.getCell('B' + sr); c.value = '⚠  ' + flags.length + ' item' + (flags.length > 1 ? 's' : '') + ' require review'; c.font = F({ bold: true, color: { argb: 'FF9C4221' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDECEA' } }; su.getRow(sr).height = 22; sr += 2;
    txt(su, 'B' + sr, 'Exceptions — Review Required', { font: { bold: true, size: 12 } }); sr++;
    for (const f of flags) {
      su.mergeCells('B' + sr + ':G' + sr);
      const cc = su.getCell('B' + sr);
      cc.value = (f.severity === 'exception' ? '✖ ' : '⚠ ') + f.message;
      cc.font = F({ color: { argb: f.severity === 'exception' ? 'FFB3261E' : 'FF9C4221' } });
      cc.alignment = { wrapText: true, vertical: 'top' };
      su.getRow(sr).height = Math.min(220, 14 * Math.max(1, Math.ceil(String(f.message).length / 110)) + 4); sr++;
    }
    sr++;
  }
  txt(su, 'B' + sr, 'Balance-sheet tie', { font: { bold: true, size: 12 } }); sr++;
  const assetRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Asset').map((cd) => catTie[cd.key]);
  const liabRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Liability').map((cd) => catTie[cd.key]);
  const eqRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Equity').map((cd) => catTie[cd.key]);
  const sumOf = (refs) => refs.length ? refs.join('+') : '0';
  const aRow = sr; txt(su, 'C' + sr, 'Total assets'); num(su, 'D' + sr, { formula: sumOf(assetRefs) }); sr++;
  const lRow = sr; txt(su, 'C' + sr, 'Total liabilities'); num(su, 'D' + sr, { formula: sumOf(liabRefs) }); sr++;
  const eRow = sr; txt(su, 'C' + sr, 'Total members’ equity'); num(su, 'D' + sr, { formula: sumOf(eqRefs) }); sr++;
  const nRow = sr; txt(su, 'C' + sr, 'Net income (loss) — fiscal YTD (per Income Statement)'); num(su, 'D' + sr, { formula: niRef }); sr++;
  txt(su, 'C' + sr, 'Assets − (Liabilities + Equity + Net income)  —  should be $0.00', { font: { bold: true } });
  { const c = num(su, 'D' + sr, { formula: 'D' + aRow + '-(D' + lRow + '+D' + eRow + '+D' + nRow + ')' }, { font: { bold: true } }); c.border = { top: THIN }; }
  sr += 3;

  txt(su, 'B' + sr, 'Balance comparison — current month vs prior month', { font: { bold: true, size: 12 } }); sr++;
  txt(su, 'B' + sr, 'Prior-month balances are taken directly from CloudLedger as of ' + short(m.beg) + '; current-month balances link to the leadsheets.', { font: { italic: true, color: { argb: 'FF7F7F7F' } } }); sr++;
  hdr(su, sr, 2, ['Account No.', 'Account Name', 'Prior Month ' + short(m.beg), 'Current Month ' + short(m.end), 'Change', 'Change %']); sr++;
  const priorOf = (a) => (a.cat === 'interco' ? (a.type === 'Asset' ? a.begin : -a.begin) : a.begin);
  for (const cd of CATS) {
    const rows = data.byCat[cd.key] || []; if (!rows.length) continue;
    const hcell = su.getCell('B' + sr); hcell.value = cd.label; hcell.font = F({ bold: true });
    for (let c = 2; c <= 7; c++) su.getRow(sr).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRPFILL } };
    sr++;
    const gFirst = sr;
    for (const a of rows) {
      txt(su, 'B' + sr, a.code, { align: { horizontal: 'left' } }); txt(su, 'C' + sr, a.name);
      num(su, 'D' + sr, priorOf(a));
      const lc = leadCell.get(String(a.code));
      num(su, 'E' + sr, lc ? { formula: lc } : (a.cat === 'interco' ? (a.type === 'Asset' ? a.end : -a.end) : a.end));
      num(su, 'F' + sr, { formula: 'E' + sr + '-D' + sr });
      num(su, 'G' + sr, { formula: 'IF(ABS(D' + sr + ')<0.005,"",F' + sr + '/ABS(D' + sr + '))' }, { fmt: PCT });
      sr++;
    }
    const gLast = sr - 1;
    txt(su, 'C' + sr, 'Total ' + cd.label, { font: { bold: true } });
    for (const c of ['D', 'E', 'F']) { const cell = num(su, c + sr, { formula: 'SUM(' + c + gFirst + ':' + c + gLast + ')' }, { font: { bold: true } }); cell.border = { top: THIN }; }
    num(su, 'G' + sr, { formula: 'IF(ABS(D' + sr + ')<0.005,"",F' + sr + '/ABS(D' + sr + '))' }, { fmt: PCT, font: { bold: true } });
    sr += 2;
  }
}

// ─── Workbook ─────────────────────────────────────────────────────────────────
function buildWorkbook(data) {
  const { entity, month: m, byCat } = data;
  const en = entity.name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  const used = new Set(['summary']);
  for (const cd of CATS) used.add(cd.tab.toLowerCase());
  const catTie = {}; const leadCell = new Map();
  const newLead = (cd) => wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });

  if ((byCat.cash || []).length) {
    const cd = catOf('cash'); const ws = newLead(cd); const refs = new Map();
    for (const a of byCat.cash) refs.set(String(a.code), buildCashRecTab(wb, a, en, m, used, cd.tab));
    catTie.cash = cashLeadsheet(ws, cd, en, m, byCat.cash, refs, leadCell);
  }
  if ((byCat.ar || []).length) {
    const cd = catOf('ar'); const ws = newLead(cd); const refs = new Map();
    const t = buildArTabs(wb, data, en, m, used, cd.tab);
    const others = byCat.ar.filter((a) => !data.arPrimary || String(a.code) !== String(data.arPrimary.code));
    if (data.arPrimary) refs.set(String(data.arPrimary.code), { sheet: t.agingTab, endRef: t.totalRef });
    if (others.length) { const ag = wb.getWorksheet(t.agingTab); const rb = rollBlock(ag, ag.rowCount + 3, 'Other receivable accounts (per general ledger)', others, m, t.agingTab); for (const [k, v] of rb.refs) refs.set(k, v); }
    catTie.ar = simpleLeadsheet(ws, cd, en, m, byCat.ar, refs, leadCell);
  }
  if ((byCat.prepaid || []).length) {
    const cd = catOf('prepaid'); const ws = newLead(cd); const refs = new Map();
    for (const a of byCat.prepaid) refs.set(String(a.code), buildPrepaidTab(wb, a, data.prepaidByAcct.get(String(a.code)) || [], en, m, used, cd.tab));
    catTie.prepaid = simpleLeadsheet(ws, cd, en, m, byCat.prepaid, refs, leadCell);
  }
  if ((byCat.fixed || []).length) {
    const cd = catOf('fixed'); const ws = newLead(cd);
    const fa = buildFixedSchedule(wb, data, en, m, used, cd.tab);
    catTie.fixed = fixedLeadsheet(ws, cd, en, m, data.fixedPairs, fa.refs, fa.tab, leadCell);
  }
  for (const key of ['otherassets', 'invest']) {
    const rows = byCat[key] || []; if (!rows.length) continue;
    const cd = catOf(key); const ws = newLead(cd);
    const sc = buildCategorySchedule(wb, cd, rows, en, m, used, cd.tab);
    catTie[key] = key === 'otherassets' ? otherAssetsLeadsheet(ws, cd, en, m, rows, sc.refs, data.projects || [], leadCell) : simpleLeadsheet(ws, cd, en, m, rows, sc.refs, leadCell);
  }
  if ((byCat.interco || []).length) {
    const cd = catOf('interco'); const ws = newLead(cd);
    const t = buildTieOut(wb, byCat.interco, en, m, used, cd.tab);
    catTie.interco = intercoLeadsheet(ws, cd, en, m, byCat.interco, t.refs, leadCell);
  }
  if ((byCat.ap || []).length) {
    const cd = catOf('ap'); const ws = newLead(cd); const refs = new Map();
    const t = buildApTabs(wb, data, en, m, used, cd.tab);
    const others = byCat.ap.filter((a) => !data.apPrimary || String(a.code) !== String(data.apPrimary.code));
    if (data.apPrimary) refs.set(String(data.apPrimary.code), { sheet: t.agingTab, endRef: t.totalRef });
    if (others.length) { const ag = wb.getWorksheet(t.agingTab); const rb = rollBlock(ag, ag.rowCount + 3, 'Accrued and other payable accounts (per general ledger)', others, m, t.agingTab); for (const [k, v] of rb.refs) refs.set(k, v); }
    catTie.ap = simpleLeadsheet(ws, cd, en, m, byCat.ap, refs, leadCell);
  }
  if ((byCat.cc || []).length) {
    const cd = catOf('cc'); const ws = newLead(cd); const refs = new Map();
    for (const a of byCat.cc) refs.set(String(a.code), buildCcRecTab(wb, a, en, m, used, cd.tab));
    catTie.cc = ccLeadsheet(ws, cd, en, m, byCat.cc, refs, leadCell);
  }
  for (const key of ['debt', 'otherliab']) {
    const rows = byCat[key] || []; if (!rows.length) continue;
    const cd = catOf(key); const ws = newLead(cd);
    const sc = buildCategorySchedule(wb, cd, rows, en, m, used, cd.tab);
    catTie[key] = simpleLeadsheet(ws, cd, en, m, rows, sc.refs, leadCell);
  }
  const niRef = buildIncomeStatement(wb, en, m, data.ni, used);
  if ((byCat.equity || []).length) {
    const cd = catOf('equity'); const ws = newLead(cd);
    const eq = buildEquityRollforward(wb, byCat.equity, data, en, m, used, cd.tab, niRef);
    catTie.equity = simpleLeadsheet(ws, cd, en, m, byCat.equity, eq.refs, leadCell);
  }
  buildSummary(su, en, m, data, catTie, niRef, leadCell);
  return wb;
}

// ─── Persistence ──────────────────────────────────────────────────────────────
const folderFor = (m) => 'Workpapers/Monthly Closing Workpapers/' + m.year;
const fileNameFor = (m) => 'Monthly_Closing_Workpapers_' + m.label + '.xlsx';

function saveToWorkpapers(ctx, eid, m, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(m), original = fileNameFor(m);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid)); fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function registerClaMonthlyCloseRoutes(app, ctx) {
  const { db, auth, requireEntityAccess, requireRole } = ctx;
  ensureRegisterSchema(db);

  app.post('/api/workpapers/cla-monthly-close/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const m = resolveMonth((req.body && req.body.month_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildClaData(ctx, m, eid);
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, m, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-ClaClose-Summary', JSON.stringify({
          month: m.label, month_name: m.monthName + ' ' + m.year,
          saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          accounts: data.acctRows.length,
          ties: data.ties, exceptions: (data.flags || []).length, flags: (data.flags || []),
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });

  // Registers behind the prepaid and fixed-asset schedules.
  app.get('/api/workpapers/cla-monthly-close/:entity_id/registers', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
        if (!ent) return res.status(404).json({ error: 'Entity not found' });
        ensureSeed(db, ent);
        const reg = loadRegisters(db, eid);
        res.json({ entity_id: eid, prepaid: reg.prepaid, fixed_assets: reg.fixedAssets });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });
  app.put('/api/workpapers/cla-monthly-close/:entity_id/registers', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        replaceRegisters(db, eid, req.body || {});
        const reg = loadRegisters(db, eid);
        res.json({ ok: true, entity_id: eid, prepaid: reg.prepaid, fixed_assets: reg.fixedAssets });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });
}

module.exports = { registerClaMonthlyCloseRoutes, buildWorkbook, buildClaData, prepaidSchedule, fixedSchedule };
