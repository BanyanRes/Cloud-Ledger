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
  if (v && typeof v === 'object' && v.formula) c.value = v; else c.value = (v == null || v === '' ? null : v);
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
  // Intercompany is intentionally omitted: intercompany accounts are reconciled in
  // the Intercompany module for all entities, not on this monthly close workbook.
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

// Fiscal-YTD GL detail for a set of accounts, bucketed by account code, each line
// carrying its posting month (1..n). Used to pinpoint the transaction behind a
// schedule-vs-GL difference on the Summary.
function glYtd(db, eid, from, to, codes) {
  const set = new Set([...codes].map(String));
  if (!set.size) return new Map();
  const rows = db.prepare(`SELECT je.date date, je.entry_num entry_num, je.doc_number doc_number,
      je.vendor vendor, je.memo memo, jl.account_code code, jl.debit debit, jl.credit credit, jl.description description,
      CAST(strftime('%m', je.date) AS INTEGER) mo
    FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ?
    ORDER BY jl.account_code, je.date, je.entry_num, jl.id`).all(eid, from, to);
  const byAcct = new Map();
  for (const r of rows) {
    const c = String(r.code); if (!set.has(c)) continue;
    if (!byAcct.has(c)) byAcct.set(c, []);
    byAcct.get(c).push({
      date: r.date, num: (r.entry_num != null ? String(r.entry_num) : '') || String(r.doc_number || ''),
      vendor: r.vendor || '', memo: r.memo || r.description || '', debit: r2(r.debit || 0), credit: r2(r.credit || 0), mo: r.mo,
    });
  }
  return byAcct;
}

// Other-asset balances as of a date, per account, split by project. Banyan carries
// the project (e.g. Van Buren, Apache) on journal_lines.project_id -> dim_projects
// (Bill.com's Department maps to it), NOT on the location dimension. Signed
// debit-natural (assets). Returns the project name list and a Map(code ->
// Map(project -> balance)); '(No project)' collects any untagged lines so the row
// still ties to the GL, but is dropped from the columns when nothing is untagged.
function oaBalancesByProject(db, eid, asOf, codes) {
  const set = codes.map(String);
  if (!set.length) return { projects: [], byAcct: new Map(), hasNoProject: false };
  const ph = set.map(() => '?').join(',');
  const rows = db.prepare(`SELECT jl.account_code code, COALESCE(dp.name, '(No project)') proj, SUM(jl.debit - jl.credit) amt
    FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN dim_projects dp ON dp.id = jl.project_id
    WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code IN (${ph})
    GROUP BY jl.account_code, proj`).all(eid, asOf, ...set);
  const byAcct = new Map(); const projSet = new Set(); let hasNoProject = false;
  for (const r of rows) {
    const c = String(r.code), proj = r.proj || '(No project)';
    if (!byAcct.has(c)) byAcct.set(c, new Map());
    byAcct.get(c).set(proj, r2((byAcct.get(c).get(proj) || 0) + r.amt));
    if (proj !== '(No project)') projSet.add(proj);
    else if (Math.abs(r2(r.amt)) >= 0.005) hasNoProject = true;
  }
  const projects = Array.from(projSet).sort((a, b) => a.localeCompare(b));
  return { projects, byAcct, hasNoProject };
}

// Offset-letter sequence for the AP recon (a..z then A..Z, skipping x/X which
// flags beginning-balance items), then double letters.
function apLetterSeq(n) {
  const S = 'abcdefghijklmnopqrstuvwyzABCDEFGHIJKLMNOPQRSTUVWYZ'.split('');
  if (n < S.length) return S[n];
  const a = Math.floor(n / S.length) - 1, b = n % S.length;
  return (a >= 0 && a < S.length ? S[a] : 'z') + S[b];
}
function apGroupSum(rows, keyFn) { const m = new Map(); for (const x of rows) { const k = keyFn(x) || '(unnamed)'; m.set(k, r2((m.get(k) || 0) + (x.amount || 0))); } return m; }

// AP Recon engine (same symmetric offset matcher CLA/Weaver use, adapted for a
// monthly close on any A/P account): inception-to-date on the A/P account, so all
// paid bills cancel against their payments and every remaining unmatched credit
// is a genuinely open invoice, which should equal the GL A/P balance and the
// Bill.com open-invoice report. Beginning balance is 0 (nothing before inception).
function buildApReconData(db, eid, apCode, asOf, bcfg) {
  const r2c = (n) => Math.round((Number(n) || 0) * 100);
  const begin = 0;
  const lines = db.prepare(
    'SELECT je.id entry_id, je.date date, je.entry_num entry_num, je.doc_number doc_number, je.vendor vendor, je.memo memo, jl.id line_id, jl.debit debit, jl.credit credit, jl.description description '
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + 'WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code = ? '
    + 'ORDER BY je.date, je.entry_num, jl.id'
  ).all(eid, asOf, String(apCode));
  const offStmt = db.prepare('SELECT jl.account_code code, a.name name FROM journal_lines jl LEFT JOIN accounts a ON a.entity_id = ? AND a.code = jl.account_code WHERE jl.entry_id = ? AND jl.account_code <> ?');
  const offCache = new Map();
  const offsetFor = (entryId) => {
    if (offCache.has(entryId)) return offCache.get(entryId);
    const rs = offStmt.all(eid, entryId, String(apCode));
    let v; if (!rs.length) v = ''; else if (rs.length === 1) v = rs[0].code + (rs[0].name ? ' ' + rs[0].name : ''); else v = '-Split-';
    offCache.set(entryId, v); return v;
  };
  const vmap = new Map();
  for (const l of lines) { const cr = r2(l.credit || 0); if (cr > 0 && l.vendor && String(l.vendor).trim()) { const k = r2c(cr); if (!vmap.has(k)) vmap.set(k, { vendor: String(l.vendor).trim(), invoice: l.doc_number != null ? String(l.doc_number) : '', date: l.date, num: l.entry_num != null ? String(l.entry_num) : '', memo: l.memo || l.description || '' }); } }
  const vlookup = (amt, row) => {
    const hit = vmap.get(r2c(amt)); if (hit) return hit;
    const mm = String((row.memo || row.description) || '').match(/^Bill\s*-\s*([^:]+):/i);
    return { vendor: mm ? mm[1].trim() : (row.vendor || ''), invoice: row.doc_number != null ? String(row.doc_number) : '', date: row.date, num: row.entry_num != null ? String(row.entry_num) : '', memo: row.memo || row.description || '' };
  };
  const open = []; const tag = new Array(lines.length).fill(null); let pair = 0;
  const findOpp = (amtc, sign) => {
    const want = -sign, pos = [];
    for (let i = 0; i < open.length; i++) if (open[i].sign === want) pos.push(i);
    for (const i of pos) if (open[i].amtc === amtc) return [i];
    for (let a = 0; a < pos.length; a++) for (let b = a + 1; b < pos.length; b++) if (open[pos[a]].amtc + open[pos[b]].amtc === amtc) return [pos[a], pos[b]];
    for (let a = 0; a < pos.length; a++) for (let b = a + 1; b < pos.length; b++) for (let c = b + 1; c < pos.length; c++) if (open[pos[a]].amtc + open[pos[b]].amtc + open[pos[c]].amtc === amtc) return [pos[a], pos[b], pos[c]];
    return null;
  };
  lines.forEach((l, i) => {
    const dr = r2(l.debit || 0), cr = r2(l.credit || 0);
    const sign = cr > 0 ? 1 : -1, amtc = r2c(cr > 0 ? cr : dr);
    if (amtc === 0) { tag[i] = ''; return; }
    const mm = findOpp(amtc, sign);
    if (mm) { const lt = apLetterSeq(pair++); tag[i] = lt; mm.forEach((p) => (tag[open[p].idx] = lt)); mm.sort((a, b) => b - a).forEach((p) => open.splice(p, 1)); }
    else open.push({ amtc, sign, idx: i });
  });
  const openBillIdx = []; let xTotal = 0;
  for (const o of open) { if (o.sign < 0) { tag[o.idx] = 'X'; xTotal = r2(xTotal + o.amtc / 100); } else { tag[o.idx] = ''; openBillIdx.push(o.idx); } }
  let bal = begin;
  const glRows = lines.map((l, i) => { const dr = r2(l.debit || 0), cr = r2(l.credit || 0); bal = r2(bal + cr - dr); return { date: l.date, num: (l.entry_num != null ? String(l.entry_num) : '') || String(l.doc_number || ''), vendor: l.vendor || '', offset: offsetFor(l.entry_id), memo: l.description || l.memo || '', debit: dr, credit: cr, balance: bal, letter: tag[i] }; });
  const openBills = openBillIdx.map((i) => { const l = lines[i], amt = r2(l.credit || 0), v = vlookup(amt, l); return { vendor: v.vendor, invoice: v.invoice, date: v.date || l.date, num: v.num, amount: amt, memo: v.memo }; });
  openBills.sort((a, b) => b.amount - a.amount);
  const openTotal = r2(openBills.reduce((a, x) => a + x.amount, 0));
  let billcom = null, billcomSource = 'none';
  try {
    if (bcfg && bcfg.ap_aging_lines_json) {
      const arr = JSON.parse(bcfg.ap_aging_lines_json) || [];
      const norm = arr.map((x) => ({ vendor: x.vendor || '', invoice: x.invoice_number || '', date: x.bill_date || '', amount: r2(x.amount || 0) })).filter((x) => Math.abs(x.amount) >= 0.005);
      if (norm.length) { billcom = { asOf: bcfg.ap_aging_as_of || '', lines: norm, total: r2(norm.reduce((a, x) => a + x.amount, 0)) }; billcomSource = 'aging'; }
    }
  } catch (e) { billcom = null; }
  const billcomLines = billcom ? billcom.lines : [];
  const billcomTotal = billcom ? billcom.total : null;
  const billcomAsOf = billcom ? billcom.asOf : asOf;
  const glByV = apGroupSum(openBills, (x) => String(x.vendor || '').trim());
  const bcByV = apGroupSum(billcomLines, (x) => String(x.vendor || '').trim());
  const vnames = new Set([...glByV.keys(), ...bcByV.keys()]);
  const byVendor = [...vnames].map((k) => { const gl = glByV.get(k) || 0, bc = bcByV.get(k) || 0; return { vendor: k, gl, billcom: bc, diff: r2(bc - gl) }; }).sort((a, b) => b.gl - a.gl);
  const glBal = r2(begin + glRows.reduce((a, r) => a + r.credit - r.debit, 0));
  return { apCode: String(apCode), gl: glBal, begin, ending: glBal, totalDebit: r2(glRows.reduce((a, r) => a + r.debit, 0)), totalCredit: r2(glRows.reduce((a, r) => a + r.credit, 0)), glRows, openBills, openTotal, xTotal, billcomLines, billcomTotal, billcomAsOf, billcomSource, byVendor };
}

function buildClaData(ctx, m, eid) {
  const { db, computeBalances } = ctx;
  const base = buildData(ctx, m, eid);
  ensureRegisterSchema(db); ensureSeed(db, base.entity);
  const reg = loadRegisters(db, eid);
  // Categorization overrides on top of the generic rules: any account carried in
  // the fixed-asset register (asset or accumulated side) is a fixed asset, and a
  // security deposit is an other asset, as on CLA's leadsheets.
  const regFixed = new Set();
  for (const fa of reg.fixedAssets) { if (fa.asset_account) regFixed.add(String(fa.asset_account)); if (fa.dep_account) regFixed.add(String(fa.dep_account)); }
  for (const a of base.acctRows) {
    if (regFixed.has(String(a.code))) a.cat = 'fixed';
    else if (a.cat === 'fixed' && /security\s+deposit/i.test(String(a.name || ''))) a.cat = 'otherassets';
  }
  const byCat = {}; for (const a of base.acctRows) (byCat[a.cat] = byCat[a.cat] || []).push(a);

  // The Summary lists discrepancies between the GL and each supporting schedule,
  // and for every discrepancy names the GL transaction(s) that caused it. The
  // generic layer's intercompany and balance-sheet-tie flags are intentionally
  // NOT carried here — intercompany is reconciled in the Intercompany module.
  const discrepancies = [];
  const pushDisc = (d) => { d.causes = d.causes || []; discrepancies.push(d); };

  // Prior year end balances (shared by the discrepancy openings and the equity
  // rollforward), and fiscal-YTD GL detail for the schedule accounts.
  const pye = (Number(m.year) - 1) + '-12-31';
  const pyeBal = new Map();
  for (const r of (computeBalances(eid, { as_of: pye, close_pl_before: m.yearStart }) || [])) pyeBal.set(String(r.code), r2(r.balance));

  // A/R aging (current + prior month), from the invoice subledger.
  let arCur = null, arPrior = null;
  try { arCur = buildAging(db, eid, m.end); arPrior = buildAging(db, eid, m.beg); } catch (e) { arCur = null; arPrior = null; }
  const arPrimary = (byCat.ar || []).find((a) => arCur && String(a.code) === String(arCur.ar_account)) || pickPrimary(byCat.ar || [], /^accounts receivable$/i, ['12000', '120000', '11000']);
  if (arCur && Math.abs(arCur.recon_diff || 0) >= 0.01) {
    pushDisc({ code: (arCur.ar_accounts || []).join(', '), name: 'Accounts Receivable (aging)', schedName: 'AR Aging', schedBal: r2(arCur.totals.total), glBal: r2(arCur.gl_ar_balance), diff: r2(arCur.totals.total - arCur.gl_ar_balance), causes: [], note: 'A/R subledger (aging) does not reconcile to the GL control account; the residual below is unexplained — review the AR Aging detail and its un-aged GL entries.' });
  }

  // A/P aging (current + prior month): Bill.com bills by vendor + un-aged GL.
  const apPrimary = pickPrimary(byCat.ap || [], /^accounts payable$/i, ['20000', '202000']);
  let apCur = null, apPrior = null;
  if (apPrimary && typeof ctx.buildApAging === 'function') {
    try { apCur = ctx.buildApAging(eid, m.end, apPrimary.code); apPrior = ctx.buildApAging(eid, m.beg, apPrimary.code); } catch (e) { apCur = null; apPrior = null; }
  }
  if (apCur && Math.abs(apCur.recon_diff || 0) >= 0.01) {
    pushDisc({ code: apPrimary ? apPrimary.code : (apCur.ap_account || ''), name: 'Accounts Payable (aging)', schedName: 'AP Aging', schedBal: r2(apCur.grand_total.total), glBal: r2(apCur.gl_balance), diff: r2(apCur.grand_total.total - apCur.gl_balance), causes: [], note: 'A/P subledger (aging) does not reconcile to the GL control account; the residual below is unexplained — review the AP Aging detail and its un-aged GL entries.' });
  }

  // Prepaid + fixed-asset schedules. Pull fiscal-YTD GL detail for those accounts
  // so a schedule-vs-GL difference can name the exact posting behind it.
  const prepaidByAcct = new Map();
  for (const it of reg.prepaid) { const k = String(it.account_code); if (!prepaidByAcct.has(k)) prepaidByAcct.set(k, []); prepaidByAcct.get(k).push(it); }
  const prepaidSched = new Map();
  const faSched = fixedSchedule(reg.fixedAssets, m);
  const fixedPairs = splitFixed(byCat.fixed || [], reg.fixedAssets);
  const detailCodes = new Set();
  for (const a of byCat.prepaid || []) detailCodes.add(String(a.code));
  for (const p of fixedPairs) { if (p.asset) detailCodes.add(String(p.asset.code)); if (p.dep) detailCodes.add(String(p.dep.code)); }
  const ytd = glYtd(db, eid, m.yearStart, m.end, detailCodes);
  const monthlyActivity = (lines, natural) => { const a = new Array(m.monthNum).fill(0); for (const l of (lines || [])) { const s = natural === 'debit' ? r2(l.debit - l.credit) : r2(l.credit - l.debit); if (l.mo >= 1 && l.mo <= m.monthNum) a[l.mo - 1] = r2(a[l.mo - 1] + s); } return a; };
  const monthLines = (lines, k) => (lines || []).filter((l) => l.mo === k);

  // Prepaid amortization schedule vs GL.
  for (const a of byCat.prepaid || []) {
    const items = prepaidByAcct.get(String(a.code)) || [];
    const sched = prepaidSchedule(items, m);
    prepaidSched.set(String(a.code), sched);
    const tot = r2(sched.reduce((s, x) => s + x.ending, 0));
    if (!items.length && Math.abs(a.end) >= 0.01) {
      pushDisc({ code: a.code, name: a.name, schedName: 'Prepaid amortization', schedBal: 0, glBal: r2(a.end), diff: r2(-a.end), causes: [], note: 'No items in the prepaid register for this account — add the policy(ies) on the Registers screen so the amortization schedule supports the GL balance.' });
    } else if (items.length && Math.abs(tot - a.end) >= 0.01) {
      const lines = ytd.get(String(a.code)) || [];
      const glAct = monthlyActivity(lines, 'debit');
      const schedAct = new Array(m.monthNum).fill(0);
      for (const s of sched) s.cols.forEach((c, k) => { schedAct[k] = r2(schedAct[k] + c.add + c.exp); });
      const causes = [];
      const openSched = r2(items.reduce((s, it) => s + (it.opening_balance || 0), 0));
      const openGl = r2(pyeBal.get(String(a.code)) || 0);
      if (Math.abs(openSched - openGl) >= 0.01) causes.push({ date: pye, num: '', memo: 'Opening balance at ' + short(pye), amount: r2(openGl - openSched), note: 'Schedule P0 ' + fmt(openSched) + ' vs GL ' + fmt(openGl) });
      for (let k = 0; k < m.monthNum; k++) {
        if (Math.abs(glAct[k] - schedAct[k]) < 0.01) continue;
        const ls = monthLines(lines, k + 1);
        for (const l of ls) causes.push({ date: l.date, num: l.num, memo: l.memo || l.vendor || 'GL entry', amount: r2(l.debit - l.credit), note: MON3[k] + ' — GL posted ' + fmt(l.debit - l.credit) });
        if (Math.abs(schedAct[k]) >= 0.005) causes.push({ date: monthEndOf(Number(m.year), k + 1), num: '', memo: 'Less: schedule-expected amortization for ' + MON3[k], amount: r2(-schedAct[k]), note: 'What the amortization schedule modeled for ' + MON3[k] });
        else if (!ls.length) causes.push({ date: monthEndOf(Number(m.year), k + 1), num: '', memo: 'No prepaid activity posted', amount: 0, note: MON3[k] + ': schedule expected ' + fmt(schedAct[k]) });
      }
      pushDisc({ code: a.code, name: a.name, schedName: 'Prepaid amortization', schedBal: tot, glBal: r2(a.end), diff: r2(tot - a.end), causes, note: causes.length ? '' : 'Difference not isolated to a single month — review the schedule inputs (start date / monthly amount).' });
    }
  }

  // Fixed-asset schedule vs GL (cost and accumulated depreciation).
  for (const p of fixedPairs) {
    if (p.asset) {
      const rows = faSched.filter((x) => String(x.asset.asset_account) === String(p.asset.code));
      const cost = r2(rows.reduce((s, x) => s + r2(x.asset.cost), 0));
      if (!rows.length && Math.abs(p.asset.end) >= 0.01) {
        pushDisc({ code: p.asset.code, name: p.asset.name + ' (cost)', schedName: 'Fixed asset schedule', schedBal: 0, glBal: r2(p.asset.end), diff: r2(-p.asset.end), causes: [], note: 'No assets in the fixed-asset register for this account — add them on the Registers screen.' });
      } else if (rows.length && Math.abs(cost - p.asset.end) >= 0.01) {
        const lines = ytd.get(String(p.asset.code)) || [];
        const causes = lines.map((l) => ({ date: l.date, num: l.num, memo: l.memo || l.vendor || 'GL entry', amount: r2(l.debit - l.credit), note: 'Cost account activity this year' }));
        const openGl = r2(pyeBal.get(String(p.asset.code)) || 0);
        const openDiff = r2(openGl - cost);
        if (Math.abs(openDiff) >= 0.005) causes.push({ date: pye, num: '', memo: 'Cost on the books at ' + short(pye) + ' vs register', amount: openDiff, note: 'Register cost ' + fmt(cost) + ' vs GL cost at ' + short(pye) + ' ' + fmt(openGl) + ' (predates the fiscal year)' });
        pushDisc({ code: p.asset.code, name: p.asset.name + ' (cost)', schedName: 'Fixed asset schedule', schedBal: cost, glBal: r2(p.asset.end), diff: r2(cost - p.asset.end), causes, note: '' });
      }
    }
    if (p.dep) {
      const rows = faSched.filter((x) => String(x.asset.dep_account) === String(p.dep.code));
      if (rows.length) {
        const accum = r2(rows.reduce((s, x) => s + x.accumEnd, 0)); // positive magnitude
        const schedBal = r2(-accum);                                // credit balance
        if (Math.abs(schedBal - p.dep.end) >= 0.01) {
          const lines = ytd.get(String(p.dep.code)) || [];
          const glAct = monthlyActivity(lines, 'credit'); // accum dep is credit-natural
          const schedAct = new Array(m.monthNum).fill(0);
          for (const x of rows) x.dep.forEach((d, k) => { schedAct[k] = r2(schedAct[k] + d); });
          const causes = [];
          const openSched = r2(rows.reduce((s, x) => s + r2(x.asset.accum_dep_beg), 0)); // debit magnitude
          const openGl = r2(pyeBal.get(String(p.dep.code)) || 0);                        // credit (negative)
          if (Math.abs(r2(-openSched) - openGl) >= 0.01) causes.push({ date: pye, num: '', memo: 'Opening accumulated depreciation at ' + short(pye), amount: r2(openGl - r2(-openSched)), note: 'Schedule beg ' + fmt(-openSched) + ' vs GL ' + fmt(openGl) });
          for (let k = 0; k < m.monthNum; k++) {
            if (Math.abs(glAct[k] - schedAct[k]) < 0.01) continue;
            const ls = monthLines(lines, k + 1);
            for (const l of ls) causes.push({ date: l.date, num: l.num, memo: l.memo || 'Depreciation', amount: r2(l.debit - l.credit), note: MON3[k] + ' — GL posted ' + fmt(l.credit - l.debit) + ' depreciation' });
            if (Math.abs(schedAct[k]) >= 0.005) causes.push({ date: monthEndOf(Number(m.year), k + 1), num: '', memo: 'Less: schedule-expected depreciation for ' + MON3[k], amount: r2(schedAct[k]), note: 'What the depreciation schedule modeled for ' + MON3[k] });
            else if (!ls.length) causes.push({ date: monthEndOf(Number(m.year), k + 1), num: '', memo: 'No depreciation posted', amount: 0, note: MON3[k] + ': schedule expected ' + fmt(schedAct[k]) });
          }
          pushDisc({ code: p.dep.code, name: p.dep.name + ' (accumulated)', schedName: 'Fixed asset schedule', schedBal, glBal: r2(p.dep.end), diff: r2(schedBal - p.dep.end), causes, note: causes.length ? '' : 'Difference not isolated to a single month — review the depreciation schedule inputs.' });
        }
      }
    }
  }

  // Equity: monthly activity through the report month (prior year end balances
  // already loaded into pyeBal above).
  const eqMonthly = new Map(); // code -> [activity m1..mn]
  const eqCodes = new Set((byCat.equity || []).map((a) => String(a.code)));
  if (eqCodes.size) {
    const rows = db.prepare(`SELECT jl.account_code code, CAST(strftime('%m', je.date) AS INTEGER) mo, SUM(jl.credit - jl.debit) amt
      FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
      WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? GROUP BY jl.account_code, mo`).all(eid, m.yearStart, m.end);
    for (const r of rows) { const c = String(r.code); if (!eqCodes.has(c)) continue; if (!eqMonthly.has(c)) eqMonthly.set(c, new Array(m.monthNum).fill(0)); if (r.mo >= 1 && r.mo <= m.monthNum) eqMonthly.get(c)[r.mo - 1] = r2(r.amt); }
  }

  // Other Assets by project (GL location dimension), as of month end.
  const oaByProject = oaBalancesByProject(db, eid, m.end, (byCat.otherassets || []).map((a) => String(a.code)));

  // AP Recon: Bill.com A/P Detail vs GL open invoices (inception-to-date offsets).
  let apRecon = null;
  if (apPrimary) {
    try {
      const bcfg = db.prepare('SELECT ap_aging_lines_json, ap_aging_as_of FROM billcom_config WHERE entity_id = ?').get(eid);
      apRecon = buildApReconData(db, eid, apPrimary.code, m.end, bcfg);
    } catch (e) { apRecon = null; }
  }
  if (apRecon) {
    if (apRecon.billcomSource === 'none') {
      pushDisc({ code: apPrimary.code, name: 'Accounts Payable — Bill.com recon', schedName: 'AP Recon', schedBal: null, glBal: r2(apRecon.openTotal), diff: null, causes: [], note: 'No Bill.com A/P Detail report has been uploaded, so the GL A/P of ' + fmt(apRecon.openTotal) + ' is not yet independently verified against Bill.com. Upload the Bill.com A/P Detail (Open Items) report as of ' + short(m.end) + ' (card below the report), then regenerate.' });
    } else if (Math.abs(r2((apRecon.billcomTotal || 0) - apRecon.openTotal)) >= 0.01) {
      const causes = (apRecon.byVendor || []).filter((v) => Math.abs(v.diff) >= 0.01).map((v) => ({ date: '', num: '', memo: v.vendor, amount: r2(v.gl - v.billcom), note: 'Bill.com ' + fmt(v.billcom) + ' vs GL ' + fmt(v.gl) }));
      pushDisc({ code: apPrimary.code, name: 'Accounts Payable — Bill.com recon', schedName: 'AP Recon', schedBal: r2(apRecon.billcomTotal), glBal: r2(apRecon.openTotal), diff: r2((apRecon.billcomTotal || 0) - apRecon.openTotal), causes, note: causes.length ? 'Open A/P per Bill.com does not agree to the GL by vendor.' : '' });
    }
  }

  // Make every discrepancy's listed transactions foot to the difference: append a
  // residual line so the causes sum exactly to (GL - schedule).
  for (const d of discrepancies) {
    if (d.schedBal == null || d.glBal == null || d.diff == null) continue;
    const bridge = r2(d.glBal - d.schedBal);
    const identified = r2((d.causes || []).reduce((s, c) => s + (Number(c.amount) || 0), 0));
    const resid = r2(bridge - identified);
    if (Math.abs(resid) >= 0.01) d.causes.push({ date: '', num: '', memo: 'Unexplained / other difference', amount: resid, note: 'Not traced to a specific transaction' });
  }

  // Flags (for the summary banner count and the X-header / client card) are derived
  // from the discrepancy list — schedule-vs-GL differences only.
  const flags = discrepancies.map((d) => ({
    severity: 'exception',
    wp: (d.code ? d.code + ' ' : '') + d.name,
    message: (d.code ? d.code + ' ' : '') + d.name + (d.diff != null && d.schedBal != null
      ? (': ' + d.schedName + ' ' + fmt(d.schedBal) + ' vs GL ' + fmt(d.glBal) + ' (off by ' + fmt(d.diff) + ')')
      : (': ' + (d.note || 'review'))),
  }));

  return Object.assign({}, base, { flags, discrepancies, byCat, arCur, arPrior, arPrimary, apCur, apPrior, apPrimary, apRecon, reg, prepaidByAcct, prepaidSched, faSched, fixedPairs, pye, pyeBal, eqMonthly, oaByProject });
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
      num(ws, expC + r, { formula: 'IF(OR($F' + r + '="",$H' + r + '<=0),0,IF(AND(' + balC + '$10>=EOMONTH($F' + r + ',1),(' + prev + '+' + addC + r + ')>0.005),-MIN($H' + r + ',' + prev + '+' + addC + r + '),0))' });
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

// AP Recon (CLA/Weaver format): 4 tabs proving the Bill.com open-invoice report
// ties to the GL A/P. Sits alongside the AP Aging + AP Summary. Returns nothing
// (its own tabs; the AP leadsheet still links to the AP Aging total).
function buildApReconTabs(wb, recon, en, m, used) {
  const apAcct = recon.apCode;
  // 1. AP Recon (summary) — Per Bill.com vs Per GL by vendor, tying to the A/P account.
  const ap = wb.addWorksheet(reserveName('AP Recon', used), { views: [{ showGridLines: false }] });
  ap.getColumn('A').width = 3.4; ap.getColumn('B').width = 44; ap.getColumn('C').width = 16; ap.getColumn('D').width = 16; ap.getColumn('E').width = 16;
  tabHead(ap, 'Accounts Payable Reconciliation — Bill.com to GL (' + apAcct + ')', en, 'As of ' + short(m.end));
  const noBc = recon.billcomSource === 'none';
  hdr(ap, 6, 2, ['Vendor', 'Per Bill.com', 'Per GL', 'Difference']);
  let ar = 7; const first = ar;
  for (const v of (recon.byVendor || [])) {
    txt(ap, 'B' + ar, v.vendor);
    if (!noBc) num(ap, 'C' + ar, v.billcom);
    num(ap, 'D' + ar, v.gl);
    if (!noBc) num(ap, 'E' + ar, { formula: 'C' + ar + '-D' + ar });
    ar++;
  }
  const last = ar - 1;
  txt(ap, 'B' + ar, 'Total accounts payable', { font: { bold: true } });
  if (!noBc) totalCell(ap, 'C' + ar, last >= first ? 'SUM(C' + first + ':C' + last + ')' : null);
  totalCell(ap, 'D' + ar, last >= first ? 'SUM(D' + first + ':D' + last + ')' : null);
  if (!noBc) totalCell(ap, 'E' + ar, 'C' + ar + '-D' + ar);
  ar += 2;
  txt(ap, 'B' + ar, 'Accounts payable per general ledger (' + apAcct + ')'); num(ap, 'D' + ar, recon.gl); const glRow = ar; ar++;
  txt(ap, 'B' + ar, 'Open invoices remaining per GL detail (offsets applied)'); num(ap, 'D' + ar, recon.openTotal); const openRow = ar; ar++;
  txt(ap, 'B' + ar, 'Difference', { font: { bold: true } }); num(ap, 'D' + ar, { formula: 'D' + glRow + '-D' + openRow }, { font: { bold: true }, border: { top: THIN } }); ar += 2;
  ap.mergeCells('B' + ar + ':E' + ar);
  txt(ap, 'B' + ar, noBc
    ? ('Per Bill.com is blank because no Bill.com A/P Detail report has been uploaded, so the GL A/P is not yet independently verified. Upload the Bill.com A/P Detail (Open Items) report as of ' + short(m.end) + ' and regenerate.')
    : ('Per Bill.com = the uploaded Bill.com A/P Detail as of ' + short(recon.billcomAsOf || m.end) + '. Per GL = open invoices remaining on account ' + apAcct + ' after offsets, which ties to the balance sheet.'),
    { font: { italic: true, color: { argb: 'FF7F7F7F' } } });

  // 2. Bill.com AP Detail — open invoices.
  const bc = wb.addWorksheet(reserveName('Bill.com AP Detail', used), { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  bc.getColumn('A').width = 3.4; bc.getColumn('B').width = 40; bc.getColumn('C').width = 18; bc.getColumn('D').width = 14; bc.getColumn('E').width = 16;
  tabHead(bc, 'Bill.com A/P Detail — Open Invoices', en, 'As of ' + short(recon.billcomAsOf || m.end));
  hdr(bc, 5, 2, ['Vendor', 'Invoice #', 'Bill Date', 'Amount']);
  let bcr = 6; const bcFirst = bcr;
  for (const b of (recon.billcomLines || [])) { txt(bc, 'B' + bcr, b.vendor); txt(bc, 'C' + bcr, b.invoice || ''); txt(bc, 'D' + bcr, b.date || ''); num(bc, 'E' + bcr, b.amount); bcr++; }
  const bcLast = bcr - 1;
  if (noBc) { bc.mergeCells('B' + bcr + ':E' + bcr); txt(bc, 'B' + bcr, 'No Bill.com A/P Detail report uploaded. Export the Bill.com A/P Detail (Open Items) report as of ' + short(m.end) + ' and upload it (card under the report) to complete this reconciliation.', { font: { italic: true, color: { argb: 'FF9C4221' } } }); }
  else { txt(bc, 'B' + bcr, 'Total open invoices per Bill.com', { font: { bold: true } }); totalCell(bc, 'E' + bcr, bcLast >= bcFirst ? 'SUM(E' + bcFirst + ':E' + bcLast + ')' : null); }

  // 3. CL AP Detail — open invoices per GL (unmatched credits after offsets).
  const cld = wb.addWorksheet(reserveName('CL AP Detail', used), { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  cld.getColumn('A').width = 3.4; cld.getColumn('B').width = 40; cld.getColumn('C').width = 18; cld.getColumn('D').width = 14; cld.getColumn('E').width = 10; cld.getColumn('F').width = 16;
  tabHead(cld, 'Accounts Payable Detail per GL — Open Invoices', en, 'As of ' + short(m.end));
  hdr(cld, 5, 2, ['Vendor', 'Invoice #', 'Bill Date', 'JE #', 'Amount']);
  let clr = 6; const clFirst = clr;
  for (const b of (recon.openBills || [])) { txt(cld, 'B' + clr, b.vendor); txt(cld, 'C' + clr, b.invoice || ''); txt(cld, 'D' + clr, b.date || ''); txt(cld, 'E' + clr, b.num || ''); num(cld, 'F' + clr, b.amount); clr++; }
  const clLast = clr - 1;
  txt(cld, 'B' + clr, 'Total open invoices per GL', { font: { bold: true } }); totalCell(cld, 'F' + clr, clLast >= clFirst ? 'SUM(F' + clFirst + ':F' + clLast + ')' : null); clr += 2;
  cld.mergeCells('B' + clr + ':F' + clr);
  txt(cld, 'B' + clr, 'The credits on account ' + apAcct + ' left uncancelled after the GL offset analysis. Ties to the A/P balance and to the Bill.com A/P Detail.', { font: { italic: true, color: { argb: 'FF7F7F7F' } } });
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
// Other Assets leadsheet as an account × project matrix: every Other Assets GL
// account down the rows, one column per project (dim_projects, e.g. Van Buren /
// Apache) so the balances are summarized by project and each row's Total ties to
// the GL. A "(No project)" column is shown only when some line is genuinely
// untagged. Project columns and the grand total foot with SUM formulas.
function otherAssetsLeadsheet(ws, cd, en, m, rows, oaByProject, leadCell) {
  const projects = (oaByProject && oaByProject.projects) || [];
  const byAcct = (oaByProject && oaByProject.byAcct) || new Map();
  const hasNoProj = !!(oaByProject && oaByProject.hasNoProject);
  const cols = projects.concat(hasNoProj ? ['(No project)'] : []); // untagged column only when needed
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 40;
  const firstProjCol = 4; // column D
  for (let i = 0; i < cols.length; i++) ws.getColumn(colL(firstProjCol + i)).width = 15;
  const totalCol = colL(firstProjCol + cols.length);
  const commentCol = colL(firstProjCol + cols.length + 1);
  ws.getColumn(totalCol).width = 16; ws.getColumn(commentCol).width = 30;
  titleBlock(ws, en, cd.lead, m, commentCol);
  txt(ws, 'C5', 'Balances by project as of ' + short(m.end) + ' (Dr positive / Cr in parentheses)', { font: { italic: true, color: { argb: 'FF7F7F7F' } } });
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name'].concat(cols, ['Total', 'Comments']));
  let r = HR + 1; const first = r;
  for (const a of rows) {
    txt(ws, 'B' + r, a.code, { align: { horizontal: 'left' } }); txt(ws, 'C' + r, a.name);
    const projMap = byAcct.get(String(a.code)) || new Map();
    let allocated = 0;
    for (let i = 0; i < cols.length; i++) {
      const v = r2(projMap.get(cols[i]) || 0);
      if (cols[i] !== '(No project)') allocated = r2(allocated + v);
      num(ws, colL(firstProjCol + i) + r, v);
    }
    // When a "(No project)" column is present, force it to the remainder so the
    // row Total ties to the GL end balance even if some lines carry no project.
    if (hasNoProj) {
      const noProjIdx = cols.length - 1;
      const explicitNoProj = r2((projMap.get('(No project)') || 0));
      const impliedNoProj = r2(a.end - allocated);
      num(ws, colL(firstProjCol + noProjIdx) + r, Math.abs(explicitNoProj) >= 0.005 ? explicitNoProj : impliedNoProj);
    }
    num(ws, totalCol + r, { formula: 'SUM(' + colL(firstProjCol) + r + ':' + colL(firstProjCol + cols.length - 1) + r + ')' });
    leadCell.set(String(a.code), ref(cd.tab, totalCol + r));
    r++;
  }
  const last = r - 1;
  txt(ws, 'B' + r, cd.total, { font: { bold: true } });
  for (let i = 0; i <= cols.length; i++) { const c = colL(firstProjCol + i); totalCell(ws, c + r, last >= first ? 'SUM(' + c + first + ':' + c + last + ')' : null); }
  return ref(cd.tab, totalCol + r);
}
// ─── Summary ──────────────────────────────────────────────────────
// A list of the discrepancies between the GL and each supporting schedule, and,
// under each, the specific GL transaction(s) that caused it. No balance-sheet tie
// and no prior-vs-current comparison — just what disagrees and why.
function buildSummary(su, en, m, data) {
  su.getColumn('A').width = 3.4; su.getColumn('B').width = 13; su.getColumn('C').width = 46; su.getColumn('D').width = 30; su.getColumn('E').width = 18; su.getColumn('F').width = 18; su.getColumn('G').width = 18; su.getColumn('H').width = 40;
  su.getCell('C1').value = en; su.getCell('C1').font = F({ size: 16, bold: true });
  su.getCell('C2').value = 'MONTHLY CLOSING WORKPAPERS — SUMMARY'; su.getCell('C2').font = F({ size: 12, bold: true });
  su.getCell('C3').value = 'Month ended ' + spell(m.end); su.getCell('C3').font = F({ italic: true });
  const discs = data.discrepancies || [];
  let sr = 5;
  if (!discs.length) {
    su.mergeCells('B' + sr + ':H' + sr);
    const c = su.getCell('B' + sr); c.value = '✓  No discrepancies — every supporting schedule agrees to the general ledger.'; c.font = F({ bold: true, color: { argb: 'FF1E7A34' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF7EE' } }; su.getRow(sr).height = 22;
    return;
  }
  su.mergeCells('B' + sr + ':H' + sr);
  const b = su.getCell('B' + sr); b.value = '⚠  ' + discs.length + ' discrepanc' + (discs.length > 1 ? 'ies' : 'y') + ' between the GL and the supporting schedules'; b.font = F({ bold: true, color: { argb: 'FF9C4221' } });
  b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDECEA' } }; su.getRow(sr).height = 22; sr += 2;
  txt(su, 'B' + sr, 'Discrepancies — GL vs Supporting Schedule', { font: { bold: true, size: 12 } }); sr++;
  txt(su, 'B' + sr, 'Each discrepancy shows the account, the supporting schedule, the two balances and the difference; the transactions that caused it are listed beneath.', { font: { italic: true, color: { argb: 'FF7F7F7F' } } }); sr += 2;

  for (const d of discs) {
    const hc = su.getCell('B' + sr); hc.value = (d.code ? d.code + '  ' : '') + d.name; hc.font = F({ bold: true });
    for (let c = 2; c <= 8; c++) su.getRow(sr).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRPFILL } };
    txt(su, 'E' + sr, 'Per ' + d.schedName, { font: { bold: true }, align: { horizontal: 'right' } });
    txt(su, 'F' + sr, 'Per GL', { font: { bold: true }, align: { horizontal: 'right' } });
    txt(su, 'G' + sr, 'Difference', { font: { bold: true }, align: { horizontal: 'right' } });
    sr++;
    if (d.schedBal != null) num(su, 'E' + sr, d.schedBal); else txt(su, 'E' + sr, 'n/a', { align: { horizontal: 'right' } });
    num(su, 'F' + sr, d.glBal);
    if (d.diff != null) num(su, 'G' + sr, d.diff, { font: { bold: true, color: { argb: 'FFB3261E' } } });
    sr++;
    if (d.note) { su.mergeCells('C' + sr + ':H' + sr); txt(su, 'C' + sr, d.note, { font: { italic: true, color: { argb: 'FF9C4221' } }, align: { wrapText: true, vertical: 'top' } }); su.getRow(sr).height = Math.min(90, 14 * Math.max(1, Math.ceil(String(d.note).length / 120)) + 4); sr++; }
    if (d.causes && d.causes.length) {
      txt(su, 'C' + sr, 'Transaction(s) that caused the discrepancy:', { font: { bold: true, italic: true } }); sr++;
      hdr(su, sr, 3, ['Date', 'JE #', 'Memo / Vendor', 'Amount', 'Note']); sr++;
      for (const cz of d.causes) {
        if (cz.date) num(su, 'C' + sr, asDate(String(cz.date).slice(0, 10)), { fmt: DATEFMT }); else txt(su, 'C' + sr, '');
        txt(su, 'D' + sr, cz.num || '');
        txt(su, 'E' + sr, cz.memo || '');
        num(su, 'F' + sr, cz.amount);
        if (cz.note) { su.mergeCells('G' + sr + ':H' + sr); txt(su, 'G' + sr, cz.note, { align: { wrapText: true } }); }
        sr++;
      }
      const causeSum = r2((d.causes || []).reduce((s, c) => s + (Number(c.amount) || 0), 0));
      txt(su, 'C' + sr, 'Total of transactions above (moves Per schedule to Per GL)', { font: { bold: true } });
      num(su, 'F' + sr, causeSum, { font: { bold: true }, border: { top: THIN } });
      txt(su, 'G' + sr, '(= Per GL − Per schedule)', { font: { italic: true } });
      sr++;
    }
    sr += 1;
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
  // Other Assets is an account × project matrix (no roll-forward schedule);
  // Investments keeps the roll-forward schedule.
  if ((byCat.otherassets || []).length) {
    const cd = catOf('otherassets'); const ws = newLead(cd);
    catTie.otherassets = otherAssetsLeadsheet(ws, cd, en, m, byCat.otherassets, data.oaByProject, leadCell);
  }
  if ((byCat.invest || []).length) {
    const cd = catOf('invest'); const ws = newLead(cd);
    const sc = buildCategorySchedule(wb, cd, byCat.invest, en, m, used, cd.tab);
    catTie.invest = simpleLeadsheet(ws, cd, en, m, byCat.invest, sc.refs, leadCell);
  }
  if ((byCat.ap || []).length) {
    const cd = catOf('ap'); const ws = newLead(cd); const refs = new Map();
    const t = buildApTabs(wb, data, en, m, used, cd.tab);
    const others = byCat.ap.filter((a) => !data.apPrimary || String(a.code) !== String(data.apPrimary.code));
    if (data.apPrimary) refs.set(String(data.apPrimary.code), { sheet: t.agingTab, endRef: t.totalRef });
    if (others.length) { const ag = wb.getWorksheet(t.agingTab); const rb = rollBlock(ag, ag.rowCount + 3, 'Accrued and other payable accounts (per general ledger)', others, m, t.agingTab); for (const [k, v] of rb.refs) refs.set(k, v); }
    catTie.ap = simpleLeadsheet(ws, cd, en, m, byCat.ap, refs, leadCell);
    // Bill.com-to-GL AP Recon (4 tabs) alongside the AP Aging + AP Summary.
    if (data.apRecon) buildApReconTabs(wb, data.apRecon, en, m, used);
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
  buildSummary(su, en, m, data);
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
