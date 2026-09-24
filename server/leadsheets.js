// ─── Month-End Leadsheets (monthly balance-sheet close package) ──────────────
//
// Reproduces the entity's monthly FloQast close package — one workbook per
// leadsheet, in the same layout — but built entirely from the CloudLedger GL:
//
//   01 Cash and Cash Equivalents   Cash Leadsheet + one tab per bank account
//                                  (statement vs register, per the CL bank rec)
//   03 Accounts Receivable         AR Leadsheet + A/R aging (ar.js buildAging)
//   05 Prepaid Expenses            12-month prepaid schedule + expense JE
//   06 Other Assets                one "prior balance + current month" tab each
//   07 Fixed Assets                Fixed Asset Schedule (straight-line) + tie
//   08 Accounts Payable            A/P aging built from the GL (apaging.js)
//   09 Accrued Expenses            12-month accrual schedule + accrual JE
//   11 Other Liabilities           same schedule (security deposits etc.)
//   12 Debt                        one roll-forward tab per loan, CP / LTP split
//   14 Intercompany                one tab per due to/from + counterparty Tie Out
//   15 Equity Rollforward          monthly roll-forward + YTD NI by location
//
// Conventions (same as the FloQast originals):
//   • Leadsheet balances are Dr (Cr): assets positive, liabilities/equity negative.
//   • Each leadsheet row links to its tab by HYPERLINK() and pulls its balance
//     from that tab by formula; totals are SUBTOTAL()s. The FloQast "FQ Anchor"
//     column is replaced by the CL trial-balance figure and a Variance.
//   • Blue cells are entry cells (useful lives, start/end dates, CP split, …).
//   • Every formula also carries its computed result, so the file reads
//     correctly before Excel recalculates.
//
// Accounts are assigned to leadsheets by the Banyan chart-of-accounts ranges
// (categorize()), so every balance-sheet account with activity lands on exactly
// one leadsheet. Silsbee's profile only sets the display name and the asset ->
// accumulated-depreciation mapping. Filed under Workpapers › Month-End
// Leadsheets › YYYY › MM Month; the route returns all workbooks as one .zip.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const { buildAging } = require('./ar');
const { counterpartyNet, resolveCounterparty } = require('./quarterlyclose');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const nz = (n) => Math.abs(Number(n) || 0) >= 0.005;
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
const pad = (n) => String(n).padStart(2, '0');
const fmt = (n) => (Number(n) < 0 ? '-' : '') + '$' + Math.abs(Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ── Period ────────────────────────────────────────────────────────────────────
function resolveMonth(input) {
  const m = String(input || '').match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) throw new Error('month_end must be a month in YYYY-MM (or YYYY-MM-DD) form');
  const y = Number(m[1]), mo = Number(m[2]);
  if (mo < 1 || mo > 12) throw new Error('month_end must be a valid month');
  const eom = (yy, mm) => new Date(Date.UTC(yy, mm, 0)).toISOString().slice(0, 10);
  return {
    y, mo, label: y + '-' + pad(mo), monthName: MONTHS[mo - 1],
    start: y + '-' + pad(mo) + '-01', end: eom(y, mo), prior: eom(y, mo - 1),
    yearStart: y + '-01-01', priorYE: (y - 1) + '-12-31', yearEnd: y + '-12-31',
    fyEnds: MONTHS.map((_, i) => eom(y, i + 1)),
  };
}
const dt = (s) => { const [y, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const serial = (s) => Math.round(dt(s).getTime() / 86400000) + 25569; // Excel date serial
const mdy = (s) => { const [y, m, d] = String(s).split('-'); return m + '/' + d + '/' + y; };

// ── Ledger access ─────────────────────────────────────────────────────────────
function makeLedger(db, eid) {
  const accts = db.prepare('SELECT code, name, type, bank_acct FROM accounts WHERE entity_id = ?').all(eid);
  const info = new Map(accts.map((a) => [String(a.code), a]));
  const balStmt = db.prepare('SELECT jl.account_code AS code, SUM(jl.debit) AS td, SUM(jl.credit) AS tc FROM journal_lines jl '
    + 'JOIN journal_entries je ON je.id = jl.entry_id WHERE je.entity_id = ? AND je.date <= ? GROUP BY jl.account_code');
  const cache = new Map();
  const balances = (asOf) => {
    if (!cache.has(asOf)) {
      const mm = new Map();
      for (const r of balStmt.all(eid, asOf)) mm.set(String(r.code), r2((r.td || 0) - (r.tc || 0)));
      cache.set(asOf, mm);
    }
    return cache.get(asOf);
  };
  const linesStmt = db.prepare('SELECT je.id AS entry_id, je.date, je.entry_num, je.doc_number, je.vendor, je.memo, '
    + 'jl.id AS line_id, jl.account_code, jl.debit, jl.credit, jl.description, dl.name AS location_name '
    + 'FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id LEFT JOIN dim_locations dl ON dl.id = jl.location_id '
    + 'WHERE je.entity_id = ? AND jl.account_code = ? AND je.date >= ? AND je.date <= ? ORDER BY je.date, je.entry_num, jl.id');
  const offStmt = db.prepare('SELECT jl.account_code AS code, a.name AS name, jl.debit, jl.credit FROM journal_lines jl '
    + 'LEFT JOIN accounts a ON a.entity_id = ? AND a.code = jl.account_code WHERE jl.entry_id = ? AND jl.account_code <> ?');
  const everStmt = db.prepare('SELECT DISTINCT jl.account_code AS code FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id '
    + 'WHERE je.entity_id = ? AND je.date <= ?');
  return {
    eid, info,
    dr: (code, asOf) => balances(asOf).get(String(code)) || 0,           // Dr(Cr) signed balance
    lines: (code, from, to) => linesStmt.all(eid, String(code), from || '0000-01-01', to)
      .map((l) => Object.assign(l, { dr: r2((l.debit || 0) - (l.credit || 0)) })),
    offsets: (entryId, code) => offStmt.all(eid, entryId, String(code)),
    activeCodes: (asOf) => everStmt.all(eid, asOf).map((r) => String(r.code)),
  };
}
const isBS = (t) => t === 'Asset' || t === 'Liability' || t === 'Equity';
const natSign = (type) => (type === 'Asset' ? 1 : -1); // natural balance = Dr(Cr) * natSign

// Banyan chart-of-accounts ranges -> leadsheet. First match wins.
function categorize(a) {
  const c = parseInt(String(a.code), 10) || 0, n = String(a.name || '').toLowerCase(), t = a.type;
  const inR = (lo, hi) => c >= lo && c <= hi;
  if (t === 'Asset') {
    if (Number(a.bank_acct) === 1 || inR(10000, 10999)) return 'cash';
    if (inR(12000, 12009) || /accounts receivable|allowance for doubtful/.test(n)) return 'ar';
    if (inR(13000, 13099) || /^prepaid/.test(n)) return 'prepaid';
    if (inR(15000, 16999)) return 'fixed';
    if (inR(18000, 18999) || /\bdue from\b/.test(n)) return 'interco';
    return 'otherAssets';
  }
  if (t === 'Liability') {
    if (c === 20000 || /^accounts payable$/.test(n)) return 'ap';
    if (inR(23000, 23999) || /\bdue to\b/.test(n)) return 'interco';
    if (inR(22000, 22999) || inR(25000, 25999) || /interest payable|loans? payable|notes? payable|line of credit|mortgage/.test(n)) return 'debt';
    if (inR(24000, 24999) || /security deposit/.test(n)) return 'otherLiab';
    return 'accrued';
  }
  return 'equity';
}

// Per-entity presentation profile. Silsbee: FloQast header name + the
// asset -> accumulated depreciation account mapping from its leadsheet.
function profileFor(ent) {
  const nm = String(ent.name || ''), code = String(ent.code || '');
  if ((/silsbee/i.test(nm) && /property owner/i.test(nm)) || code === 'CLRSILSB2') {
    return {
      displayName: 'County Line Rail Silsbee LLC',
      depMap: { 15210: '16500', 15150: '16160', 15165: '16160', 15175: '16160', 15200: '16160', 15220: '16160' },
    };
  }
  return { displayName: nm, depMap: {} };
}

// ── Workbook styling (Calibri, FloQast look) ──────────────────────────────────
const F = (o = {}) => Object.assign({ name: 'Calibri', size: 11 }, o);
const ACC = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const DATEF = 'mm-dd-yy';
const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const ENTRY = solid('FFDCE6F1');   // FloQast "Entry Cell" blue
const YELLOW = solid('FFFFFF00');
const NAVY = solid('FF2E334E');
const GREENH = solid('FFEBF1DE');
const REDH = solid('FFF2DCDB');
const RECH = solid('FFCEDEEF');
const THIN = { style: 'thin' }, MED = { style: 'medium' }, DBL = { style: 'double' };
const BOX = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const LINKF = F({ bold: true, italic: true, color: { argb: 'FF0563C1' } });
const REDF = F({ color: { argb: 'FFFF0000' } });
const qs = (name) => "'" + String(name).replace(/'/g, "''") + "'!";
const colL = (n) => { let s = ''; while (n > 0) { const k = (n - 1) % 26; s = String.fromCharCode(65 + k) + s; n = Math.floor((n - 1) / 26); } return s; };
const fv = (formula, result) => ({ formula, result });
const hyper = (sheet, cell, text) => fv('HYPERLINK("#' + qs(sheet).replace(/"/g, '""') + cell + '","' + String(text).replace(/"/g, '""') + '")', String(text));

function put(ws, ref, value, o = {}) {
  const c = ws.getCell(ref);
  c.value = value;
  c.font = o.font || F(o.fontOpts || {});
  if (o.fill) c.fill = o.fill;
  if (o.numFmt) c.numFmt = o.numFmt;
  if (o.border) c.border = o.border;
  if (o.align) c.alignment = typeof o.align === 'string' ? { horizontal: o.align } : o.align;
  return c;
}
const money = (ws, ref, value, o = {}) => put(ws, ref, value, Object.assign({ numFmt: ACC }, o));
const dateCell = (ws, ref, s, o = {}) => put(ws, ref, s ? dt(s) : null, Object.assign({ numFmt: DATEF, align: 'center' }, o));

function newWb() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  return wb;
}

// Leadsheet title block: entity / title / Month Ended / Year Ended.
function titleBlock(ws, P, title) {
  put(ws, 'C1', P.display, { fontOpts: { bold: true, size: 14 } });
  put(ws, 'C2', title, { fontOpts: { bold: true, size: 12 } });
  put(ws, 'C3', 'Month Ended:');
  dateCell(ws, 'D3', P.m.end, { fontOpts: { bold: true }, fill: ENTRY, border: BOX });
  put(ws, 'C4', 'Year Ended:');
  dateCell(ws, 'D4', P.m.yearEnd, { fontOpts: { bold: true }, fill: ENTRY, border: BOX });
  ws.getColumn(1).width = 3.4;
}

// Generic leadsheet table from row 7. cols: [{ key, header, width, kind }]
// kind: code | name | link | money | text. rows: [{ key: value }]. Money
// columns flagged total:true get a SUBTOTAL(109) on the total row.
function leadTable(ws, cols, rows, totalLabel, startRow = 7) {
  const colOf = {};
  cols.forEach((c, i) => {
    const col = 2 + i; colOf[c.key] = colL(col);
    ws.getColumn(col).width = c.width;
    const cell = ws.getRow(startRow).getCell(col);
    cell.value = c.header; cell.font = F({ bold: !!c.boldHdr }); cell.border = { bottom: THIN };
    cell.alignment = { horizontal: i < 2 ? 'left' : 'center', vertical: 'bottom', wrapText: true };
  });
  let r = startRow + 1; const first = r;
  for (const row of rows) {
    cols.forEach((c, i) => {
      const cell = ws.getRow(r).getCell(2 + i);
      const v = row[c.key];
      cell.value = v === undefined ? null : v;
      if (c.kind === 'code') { cell.fill = ENTRY; cell.border = BOX; cell.alignment = { horizontal: 'center' }; cell.font = F(); }
      else if (c.kind === 'name') { cell.fill = ENTRY; cell.border = BOX; cell.font = F(); }
      else if (c.kind === 'link') { cell.font = LINKF; cell.alignment = { horizontal: 'center' }; }
      else if (c.kind === 'money') { cell.numFmt = ACC; cell.font = row['_red_' + c.key] ? REDF : F(); if (c.entry) { cell.fill = ENTRY; cell.border = BOX; } }
      else { cell.font = row['_red_' + c.key] ? REDF : F(); cell.alignment = { horizontal: c.align || 'left', wrapText: false }; }
    });
    r += 1;
  }
  const last = r - 1;
  put(ws, 'B' + r, totalLabel, { border: { top: THIN } });
  cols.forEach((c, i) => {
    if (i === 0) return;
    const L = colL(2 + i), cell = ws.getRow(r).getCell(2 + i);
    cell.border = { top: THIN };
    if (c.kind === 'money' && c.total) {
      const sum = r2(rows.reduce((s, x) => s + (Number(resultOf(x[c.key])) || 0), 0));
      cell.value = rows.length ? fv('SUBTOTAL(109,' + L + first + ':' + L + last + ')', sum) : 0;
      cell.numFmt = ACC; cell.font = F({ bold: true });
    }
  });
  return { first, last, totalRow: r, colOf };
}
const resultOf = (v) => (v && typeof v === 'object' && 'formula' in v) ? v.result : v;

// "Back to leadsheet" link + tab title used on every support tab.
function tabTitle(ws, title, lsName, linkRef = 'E1', o = {}) {
  put(ws, 'A1', title, { fontOpts: Object.assign({ bold: true, italic: true }, o.titleFont || {}) });
  put(ws, linkRef, hyper(lsName, 'A1', 'Back to leadsheet'), { font: LINKF });
}

// Variance helper: formula text + result, red when off.
function varianceCells(row, key, aRef, bRef, a, b) {
  const v = r2((Number(a) || 0) - (Number(b) || 0));
  row[key] = fv(aRef + '-' + bRef, v);
  if (nz(v)) row['_red_' + key] = true;
  return v;
}

// ─── Support-tab builders ─────────────────────────────────────────────────────

// Prior balance + current-month activity -> ending, with the month's GL lines.
// style: 'oa' (Other Assets), 'ic' (Intercompany), 'debt', 'basic'. Amounts on
// the tab are on the account's natural side; returns the ending-balance ref.
function rollTab(wb, P, a, lsName, style) {
  const ws = wb.addWorksheet(String(a.code), { views: [{ showGridLines: false }] });
  const sgn = natSign(a.type);
  const beg = r2(P.L.dr(a.code, P.m.prior) * sgn);
  const lines = P.L.lines(a.code, P.m.start, P.m.end).filter((l) => nz(l.dr));
  tabTitle(ws, a.code + ' - ' + a.name, lsName, 'E1', { titleFont: { color: { argb: 'FF000000' } } });
  [['A', 27], ['B', 15], ['C', 16], ['D', 12], ['E', 30], ['F', 18], ['G', 55]].forEach(([c, w]) => (ws.getColumn(c).width = w));
  const L = {
    oa: ['Prior Invoices through', 'Current Month', 'Total Account Bal at'],
    ic: ['Balance at', 'Current Month Changes', 'ME Balance at'],
    debt: ['Beginning Balance', 'Additions', 'Repayments', 'Ending Balance'],
    basic: ['Beginning Balance', P.m.monthName + ' ' + P.m.y + ' activity', 'Ending Balance'],
  }[style];
  const detHdr = style === 'debt' ? 9 : 8;
  const first = detHdr + 1, last = detHdr + Math.max(lines.length, 1);
  const rng = 'B' + first + ':B' + last;
  const act = r2(lines.reduce((s, l) => s + l.dr * sgn, 0));
  let r = 3;
  put(ws, 'A' + r, L[0], { fontOpts: { bold: style !== 'ic' } }); dateCell(ws, 'B' + r, P.m.prior, { fontOpts: { bold: true } });
  money(ws, 'C' + r, beg); r++;
  if (style === 'debt') {
    const add = r2(lines.filter((l) => l.dr * sgn > 0).reduce((s, l) => s + l.dr * sgn, 0));
    put(ws, 'A' + r, L[1]); money(ws, 'C' + r, fv('SUMIF(' + rng + ',">0")', add)); r++;
    put(ws, 'A' + r, L[2]); money(ws, 'C' + r, fv('SUMIF(' + rng + ',"<0")', r2(act - add)), { border: { bottom: THIN } }); r++;
  } else {
    put(ws, 'A' + r, L[1]); money(ws, 'C' + r, fv('SUM(' + rng + ')', act), { border: { bottom: THIN } }); r++;
  }
  const endRow = r, end = r2(beg + act);
  put(ws, 'A' + r, L[L.length - 1], { fontOpts: { bold: true } }); dateCell(ws, 'B' + r, P.m.end, { fontOpts: { bold: true } });
  money(ws, 'C' + r, fv('SUM(C3:C' + (r - 1) + ')', end), { fontOpts: { bold: true }, fill: ENTRY });
  ['Date', 'Amount', 'JE #', 'Vendor / Payee', 'Invoice / Doc #', 'Memo'].forEach((h, i) =>
    put(ws, colL(1 + i) + detHdr, h, { fontOpts: { bold: true }, border: { bottom: THIN }, align: i === 1 ? 'right' : 'left' }));
  let rr = first;
  for (const l of lines) {
    dateCell(ws, 'A' + rr, l.date, { align: 'left' });
    money(ws, 'B' + rr, r2(l.dr * sgn));
    put(ws, 'C' + rr, l.entry_num != null ? 'JE-' + pad4(l.entry_num) : '');
    put(ws, 'D' + rr, l.vendor || '');
    put(ws, 'E' + rr, l.doc_number != null ? String(l.doc_number) : '');
    put(ws, 'F' + rr, l.description || l.memo || '');
    rr++;
  }
  if (!lines.length) put(ws, 'A' + first, 'No activity in ' + P.m.monthName + ' ' + P.m.y + '.', { fontOpts: { italic: true, color: { argb: 'FF808080' } } });
  if (Math.abs(r2(end - P.L.dr(a.code, P.m.end) * sgn)) >= 0.01) {
    P.flags.push({ severity: 'exception', wp: a.code, message: a.code + ' ' + a.name + ': roll-forward ' + fmt(end) + ' does not agree to the GL ' + fmt(P.L.dr(a.code, P.m.end) * sgn) + '.' });
  }
  return { sheet: String(a.code), ref: 'C' + endRow, value: end, sgn };
}
const pad4 = (n) => String(n).padStart(4, '0');
// Leadsheet reference to a roll-forward tab's ending balance, in Dr (Cr) terms.
const signedRef = (t) => (t.sgn < 0 ? '-' : '') + qs(t.sheet) + t.ref;

// 12-month schedule (FloQast prepaid / accrual layout). Columns per month:
// debits (Additions / Payments), credits (Expenses / Accruals), Balance. All
// Dr(Cr) signed, so a prepaid shows positive and an accrual negative, exactly
// like the originals. Rows are one per vendor. Returns the balance at month end.
function scheduleTab(wb, P, a, lsName, kind) {
  const ws = wb.addWorksheet(String(a.code), { views: [{ showGridLines: false, state: 'frozen', xSplit: kind === 'prepaid' ? 8 : 6, ySplit: 10 }] });
  const isPre = kind === 'prepaid';
  const drLbl = isPre ? 'Additions' : 'Payments', crLbl = isPre ? 'Expenses' : 'Accruals';
  const all = P.L.lines(a.code, null, P.m.end).filter((l) => nz(l.dr));
  // Row key = vendor. Untagged lines take the only vendor if there is exactly
  // one, else the vendor named in the memo, else an "untagged" row.
  const vendors = [...new Set(all.map((l) => String(l.vendor || '').trim()).filter(Boolean))];
  const keyOf = (l) => {
    const v = String(l.vendor || '').trim(); if (v) return v;
    const memo = String((l.memo || '') + ' ' + (l.description || '')).toLowerCase();
    const hit = vendors.find((x) => { const w = x.toLowerCase().split(/[\s,.]+/).find((t) => t.length >= 3); return w && memo.includes(w); });
    if (hit) return hit;
    if (vendors.length === 1) return vendors[0];
    return '(Not tagged to a vendor)';
  };
  const groups = new Map();
  const g = (k) => { if (!groups.has(k)) groups.set(k, { key: k, p0: 0, dr: Array(12).fill(0), cr: Array(12).fill(0), lastDr: null, lastCr: null, offs: new Map() }); return groups.get(k); };
  for (const l of all) {
    const x = g(keyOf(l));
    if (l.date <= P.m.priorYE) { x.p0 = r2(x.p0 + l.dr); }
    else {
      const k = Number(l.date.slice(5, 7)) - 1;
      if (l.dr > 0) x.dr[k] = r2(x.dr[k] + l.dr); else x.cr[k] = r2(x.cr[k] + l.dr);
    }
    if (l.dr > 0) x.lastDr = l; else x.lastCr = l;
    if (l.dr < 0) { // expense (prepaid) / accrual side: remember the offset account
      for (const o of P.L.offsets(l.entry_id, a.code)) if ((o.debit || 0) > 0) x.offs.set(o.code, (x.offs.get(o.code) || 0) + o.debit);
    }
  }
  const rows = [...groups.values()].filter((x) => nz(x.p0) || x.dr.some(nz) || x.cr.some(nz))
    .sort((p, q) => String(p.key).localeCompare(String(q.key)));
  // Layout
  tabTitle(ws, a.code + ' - ' + a.name, lsName, 'E1');
  put(ws, 'G1', ' Entry Cell', { fill: ENTRY, border: BOX });
  put(ws, 'A3', 'INSTRUCTIONS', { fontOpts: { bold: true, italic: true } });
  put(ws, 'A4', '1.'); put(ws, 'B4', 'Balances and monthly activity are pulled from the CloudLedger general ledger for account ' + a.code + ', one row per vendor.');
  put(ws, 'A5', '2.'); put(ws, 'B5', isPre ? 'Key the policy start / end dates in the blue cells; the monthly amount is the latest ' + P.m.monthName + ' expense.' : 'Review each vendor balance for anything that should have been paid or reversed.');
  put(ws, 'A6', '3.'); put(ws, 'B6', 'The Journal Entry below is the current-month entry as posted in CloudLedger.');
  const fixedCols = isPre
    ? [['Date Paid', 11.4], ['Vendor', 26.4], ['Expense Acct', 19.9], ['Description', 27.1], ['Start Date', 11.1], ['End Date', 10.9], ['Monthly', 11.9]]
    : [['Date Last Paid', 11.4], ['Vendor', 26.4], ['Expense Acct / Asset Acct', 19.9], ['Description', 28.1], ['Monthly Accrual', 11.1]];
  ws.getColumn(1).width = 3.7;
  const HDR = 10, first = 11;
  fixedCols.forEach(([h, w], i) => { ws.getColumn(2 + i).width = w; put(ws, colL(2 + i) + HDR, h, { fontOpts: { bold: true }, border: { bottom: MED }, align: { horizontal: 'center', wrapText: true } }); });
  const p0Col = 2 + fixedCols.length; // Balance P0
  const balCol = (k) => p0Col + 3 * k;      // k = 0..12
  put(ws, colL(p0Col) + HDR, 'Balance P0', { fontOpts: { bold: true }, border: { bottom: MED }, align: 'center' });
  dateCell(ws, colL(p0Col) + (HDR - 1), P.m.priorYE, { fontOpts: { bold: true }, border: { top: MED } });
  ws.getColumn(p0Col).width = 13.9;
  for (let k = 1; k <= 12; k++) {
    const b = balCol(k);
    [[b - 2, drLbl + ' P' + k, 14.3], [b - 1, crLbl + ' P' + k, 14.0], [b, 'Balance P' + k, 13.6]].forEach(([c, h, w]) => {
      ws.getColumn(c).width = w; put(ws, colL(c) + HDR, h, { fontOpts: { bold: true }, border: { bottom: MED }, align: 'center' });
    });
    dateCell(ws, colL(b) + (HDR - 1), P.m.fyEnds[k - 1], { fontOpts: { bold: true }, border: { top: MED } });
  }
  let r = first;
  const colTotals = {};
  const addT = (c, v) => { colTotals[c] = r2((colTotals[c] || 0) + v); };
  for (const x of rows) {
    const lastDr = x.lastDr, lastCr = x.lastCr;
    const off = [...x.offs.entries()].sort((p, q) => q[1] - p[1])[0];
    const offName = off ? off[0] + ((P.L.info.get(off[0]) || {}).name ? ' ' + P.L.info.get(off[0]).name : '') : '';
    const monthly = r2(Math.abs(x.cr[P.m.mo - 1] || 0)) || null;
    const ent = { fill: ENTRY, border: BOX };
    if (isPre) {
      dateCell(ws, 'B' + r, lastDr ? lastDr.date : null, ent);
      put(ws, 'C' + r, x.key, ent); put(ws, 'D' + r, offName, ent);
      put(ws, 'E' + r, lastDr ? (lastDr.memo || lastDr.description || '') : '', ent);
      dateCell(ws, 'F' + r, null, ent); dateCell(ws, 'G' + r, null, ent);
      money(ws, 'H' + r, monthly, ent);
    } else {
      dateCell(ws, 'B' + r, lastDr ? lastDr.date : null, ent);
      put(ws, 'C' + r, x.key, ent); put(ws, 'D' + r, offName, ent);
      put(ws, 'E' + r, lastCr ? (lastCr.memo || lastCr.description || '') : (lastDr ? (lastDr.memo || '') : ''), ent);
      money(ws, 'F' + r, monthly, ent);
    }
    money(ws, colL(p0Col) + r, x.p0, ent); addT(p0Col, x.p0);
    let bal = x.p0;
    for (let k = 1; k <= 12; k++) {
      const b = balCol(k), future = k > P.m.mo;
      const d = future ? 0 : x.dr[k - 1], c = future ? 0 : x.cr[k - 1];
      money(ws, colL(b - 2) + r, d, ent); money(ws, colL(b - 1) + r, c, ent);
      bal = r2(bal + d + c);
      money(ws, colL(b) + r, fv('SUM(' + colL(b - 3) + r + ':' + colL(b - 1) + r + ')', bal));
      addT(b - 2, d); addT(b - 1, c); addT(b, bal);
    }
    r++;
  }
  if (!rows.length) { put(ws, 'C' + r, 'No balance or activity this year.', { fontOpts: { italic: true } }); r++; }
  const tot = r, last = r - 1;
  put(ws, 'B' + tot, isPre ? 'Total Prepaid Expenses' : 'Total ' + a.name, { border: { top: MED } });
  for (let c = p0Col; c <= balCol(12); c++) {
    const isBal = (c - p0Col) % 3 === 0;
    const k = (c - p0Col) / 3;
    const cell = money(ws, colL(c) + tot, fv('SUBTOTAL(109,' + colL(c) + first + ':' + colL(c) + last + ')', colTotals[c] || 0),
      { border: isBal ? { top: MED, bottom: MED } : { top: MED } });
    if (isBal && k === P.m.mo) { cell.fill = YELLOW; cell.font = F({ bold: true }); }
  }
  const endRef = colL(balCol(P.m.mo)) + tot, endVal = colTotals[balCol(P.m.mo)] || 0;

  // Current-month journal entry, as posted: the side that moves the schedule
  // (prepaid: expense credits; accrual: accrual credits) and its offsets.
  const jeRows = new Map();
  let crTot = 0;
  for (const l of all.filter((l) => l.date >= P.m.start && l.date <= P.m.end && l.dr < 0)) {
    crTot = r2(crTot - l.dr);
    for (const o of P.L.offsets(l.entry_id, a.code)) if ((o.debit || 0) > 0) jeRows.set(o.code, { name: o.name || '', amt: r2(((jeRows.get(o.code) || {}).amt || 0) + o.debit) });
  }
  let jr = tot + 2;
  put(ws, 'B' + jr, 'Journal Entry', { fontOpts: { bold: true }, border: { top: THIN, bottom: THIN } });
  put(ws, 'F' + jr, 'Variance');
  const jt = jr + 1;
  put(ws, 'B' + jt, 'Total', { fontOpts: { bold: true } });
  const jh = jt + 2;
  ['Date', 'Account', 'Debit', 'Credit'].forEach((h, i) => put(ws, colL(2 + i) + jh, h, { fontOpts: { bold: true }, border: { bottom: MED } }));
  let jl = jh + 1;
  const jFirst = jl;
  put(ws, 'B' + jl, fv(qs(lsName) + 'D3', serial(P.m.end)), { numFmt: DATEF });
  put(ws, 'C' + jl, a.code + ' - ' + a.name); money(ws, 'E' + jl, crTot); jl++;
  let drTot = 0;
  for (const [code, v] of jeRows) { put(ws, 'C' + jl, code + ' - ' + v.name); money(ws, 'D' + jl, v.amt); drTot = r2(drTot + v.amt); jl++; }
  const jLast = jl - 1;
  money(ws, 'D' + jt, fv('SUM(D' + jFirst + ':D' + jLast + ')', drTot), { fontOpts: { bold: true }, border: { top: THIN, bottom: DBL } });
  money(ws, 'E' + jt, fv('SUM(E' + jFirst + ':E' + jLast + ')', crTot), { fontOpts: { bold: true }, border: { top: THIN, bottom: DBL } });
  money(ws, 'F' + jt, fv('D' + jt + '-E' + jt, r2(drTot - crTot)), { fontOpts: { bold: true } });
  if (!nz(crTot)) put(ws, 'C' + jFirst, a.code + ' - ' + a.name + ' (no ' + (isPre ? 'expense' : 'accrual') + ' posted in ' + P.m.monthName + ')');

  const gl = P.L.dr(a.code, P.m.end);
  if (Math.abs(r2(endVal - gl)) >= 0.01) P.flags.push({ severity: 'exception', wp: a.code, message: a.code + ' ' + a.name + ': schedule balance ' + fmt(endVal) + ' does not agree to the GL ' + fmt(gl) + '.' });
  return { sheet: String(a.code), ref: endRef, value: endVal };
}

// ─── 01 Cash ──────────────────────────────────────────────────────────────────
function buildCash(P) {
  const wb = newWb(), LS = 'Cash Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'CASH LEAD SHEET');
  const recStmt = P.db.prepare('SELECT * FROM reconciliations WHERE entity_id = ? AND account_code = ? AND statement_date >= ? AND statement_date <= ? ORDER BY statement_date DESC, id DESC LIMIT 1');
  const rows = [];
  P.byCat.cash.forEach((a, i) => {
    const lsRow = 8 + i;
    const ws = wb.addWorksheet(String(a.code), { views: [{ showGridLines: false }] });
    [['A', 24], ['B', 18], ['C', 14], ['D', 50]].forEach(([c, w]) => (ws.getColumn(c).width = w));
    put(ws, 'A1', fv(qs(LS) + 'C' + lsRow, a.name), { fontOpts: { bold: true, italic: true, size: 16, color: { argb: 'FF0070C0' } } });
    put(ws, 'D2', hyper(LS, 'A1', 'Back to leadsheet'), { font: LINKF });
    const reg = P.L.dr(a.code, P.m.end);
    const rec = recStmt.get(P.eid, a.code, P.m.start, P.m.end);
    const activeYr = P.L.lines(a.code, P.m.yearStart, P.m.end).some((l) => nz(l.dr));
    put(ws, 'A4', 'Statement Balance', { fill: solid('FFFFFFFF') });
    money(ws, 'B4', rec ? r2(rec.statement_balance) : (nz(reg) || activeYr ? null : 0), { fill: ENTRY });
    put(ws, 'A5', 'Register Balance');
    money(ws, 'B5', reg, { fill: ENTRY });
    let comment = '', recDiff = null, bankRef = '';
    if (rec) {
      // Items open at the statement date: on/before it and not cleared by this or
      // an earlier reconciliation.
      const lines = P.db.prepare('SELECT je.id AS entry_id, je.entry_num, je.date, je.memo, jl.debit, jl.credit, '
        + '(SELECT COUNT(*) FROM journal_lines x WHERE x.entry_id = je.id AND x.id <= jl.id) - 1 AS line_index '
        + 'FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id '
        + 'WHERE je.entity_id = ? AND jl.account_code = ? AND je.date <= ? ORDER BY je.date, je.id, jl.id').all(P.eid, a.code, rec.statement_date);
      const cleared = new Set(P.db.prepare('SELECT ci.entry_id, ci.line_index FROM cleared_items ci LEFT JOIN reconciliations r ON r.id = ci.reconciliation_id '
        + 'WHERE ci.entity_id = ? AND ci.account_code = ? AND (r.statement_date IS NULL OR r.statement_date <= ?)').all(P.eid, a.code, rec.statement_date)
        .map((c) => c.entry_id + '-' + c.line_index));
      const open = lines.filter((l) => !cleared.has(l.entry_id + '-' + l.line_index)).map((l) => Object.assign(l, { amt: r2((l.debit || 0) - (l.credit || 0)) })).filter((l) => nz(l.amt));
      const dit = open.filter((l) => l.amt > 0), chk = open.filter((l) => l.amt < 0);
      const regAtStmt = P.L.dr(a.code, rec.statement_date);
      put(ws, 'A8', 'Reconciliation Report', { fontOpts: { bold: true, size: 14 }, fill: RECH });
      ['B8', 'C8', 'D8'].forEach((c) => (ws.getCell(c).fill = RECH));
      put(ws, 'A9', 'As Of ' + mdy(rec.statement_date) + ', completed in CloudLedger ' + String(rec.completed_at || '').slice(0, 10) + ' by ' + (rec.completed_by || ''), { fontOpts: { italic: true, size: 9 } });
      let r = 11;
      const sumRef = (arr, startRow) => arr.length ? 'SUM(B' + startRow + ':B' + (startRow + arr.length - 1) + ')' : '0';
      const ditStart = 21, chkStart = ditStart + Math.max(dit.length, 1) + 3;
      const ditTot = r2(dit.reduce((s, l) => s + l.amt, 0)), chkTot = r2(chk.reduce((s, l) => s + l.amt, 0));
      const stmt = r2(rec.statement_balance), adj = r2(stmt + ditTot + chkTot);
      put(ws, 'A' + r, 'Statement Ending Balance', { fontOpts: { bold: true } }); money(ws, 'B' + r, fv('B4', stmt)); r++;
      put(ws, 'A' + r, 'Deposits in Transit'); money(ws, 'B' + r, fv(sumRef(dit, ditStart), ditTot)); r++;
      put(ws, 'A' + r, 'Outstanding Checks and Charges'); money(ws, 'B' + r, fv(sumRef(chk, chkStart), chkTot), { border: { bottom: THIN } }); r++;
      put(ws, 'A' + r, 'Adjusted Bank Balance', { fontOpts: { bold: true } }); money(ws, 'B' + r, fv('SUM(B11:B13)', adj), { fontOpts: { bold: true } }); r += 2;
      put(ws, 'A' + r, 'Book Balance at ' + mdy(rec.statement_date)); money(ws, 'B' + r, regAtStmt); r++;
      recDiff = r2(adj - regAtStmt);
      put(ws, 'A' + r, 'Difference', { fontOpts: { bold: true } }); money(ws, 'B' + r, fv('B14-B16', recDiff), { fontOpts: { bold: true, color: { argb: nz(recDiff) ? 'FFFF0000' : 'FF000000' } }, border: { top: THIN, bottom: DBL } });
      const list = (title, arr, start) => {
        put(ws, 'A' + (start - 2), title, { fontOpts: { bold: true, size: 12 } });
        ['Date', 'Amount', 'JE #', 'Memo'].forEach((h, k) => put(ws, colL(1 + k) + (start - 1), h, { fontOpts: { bold: true }, fill: RECH }));
        if (!arr.length) put(ws, 'A' + start, 'None', { fontOpts: { italic: true } });
        arr.forEach((l, k) => { dateCell(ws, 'A' + (start + k), l.date, { align: 'left' }); money(ws, 'B' + (start + k), l.amt); put(ws, 'C' + (start + k), 'JE-' + pad4(l.entry_num)); put(ws, 'D' + (start + k), l.memo || ''); });
      };
      list('Deposits in Transit', dit, ditStart);
      list('Outstanding Checks and Charges', chk, chkStart);
      bankRef = 'Rec ' + mdy(rec.statement_date);
      if (rec.statement_date !== P.m.end) comment = '- Statement dated ' + mdy(rec.statement_date);
      if (nz(recDiff)) P.flags.push({ severity: 'exception', wp: '01 Cash', message: a.code + ' ' + a.name + ': the ' + mdy(rec.statement_date) + ' bank rec is off by ' + fmt(recDiff) + '.' });
    } else if (nz(reg) || activeYr) {
      put(ws, 'A8', 'No bank reconciliation has been completed in CloudLedger for ' + P.m.monthName + ' ' + P.m.y + '. Key the statement balance in B4.', { fontOpts: { italic: true, color: { argb: 'FFFF0000' } } });
      comment = '- No ' + P.m.monthName + ' bank rec in CL';
      P.flags.push({ severity: 'review', wp: '01 Cash', message: a.code + ' ' + a.name + ' (' + fmt(reg) + '): no bank reconciliation completed in CloudLedger for ' + P.m.monthName + ' ' + P.m.y + '.' });
    } else {
      put(ws, 'A11', '*** Zero balance, no activity in ' + P.m.y, { fontOpts: { bold: true, italic: true } });
      comment = '- Zero balance';
    }
    rows.push({
      code: Number(a.code) || a.code, name: a.name, link: hyper(String(a.code), 'A1', String(a.code)),
      cleared: fv(qs(String(a.code)) + 'B4', rec ? r2(rec.statement_balance) : (nz(reg) || activeYr ? 0 : 0)),
      register: fv(qs(String(a.code)) + 'B5', reg),
      recdiff: recDiff == null ? null : fv(qs(String(a.code)) + 'B17', recDiff),
      _red_recdiff: nz(recDiff),
      bank: bankRef, comments: comment, _red_comments: /No .* bank rec/.test(comment),
    });
  });
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 11.7, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 37.6, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 13, kind: 'link' },
    { key: 'cleared', header: 'Cleared Balance', width: 17.3, kind: 'money', total: true },
    { key: 'register', header: 'Register Balance', width: 17.3, kind: 'money', total: true },
    { key: 'recdiff', header: 'Rec Difference', width: 14, kind: 'money', total: true },
    { key: 'bank', header: 'Bank Statement', width: 17.9, kind: 'text', align: 'center' },
    { key: 'comments', header: 'Comments', width: 42.4, kind: 'text' },
  ], rows, 'Total Cash');
  const tot = r2(P.byCat.cash.reduce((s, a) => s + P.L.dr(a.code, P.m.end), 0));
  return { wb, summary: { balance: tot, gl: tot } };
}

// ─── 03 Accounts Receivable ───────────────────────────────────────────────────
function buildAR(P) {
  const wb = newWb(), LS = 'AR Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'ACCOUNTS RECEIVABLE LEADSHEET');
  let aging = null;
  try { aging = buildAging(P.db, P.eid, P.m.end); } catch (e) { aging = null; }
  const agingCodes = new Set(aging ? (aging.ar_accounts || []).map(String).concat(aging.ar_account ? [String(aging.ar_account)] : []) : []);
  let agingDone = false;
  const rows = [];
  for (const a of P.byCat.ar) {
    const gl = P.L.dr(a.code, P.m.end);
    const row = { code: Number(a.code) || a.code, name: a.name, gl };
    if (aging && agingCodes.has(String(a.code)) && !agingDone) {
      agingDone = true;
      const t = agingTabAR(wb, P, a, LS, aging);
      row.link = hyper(String(a.code), 'A1', String(a.code));
      row.balance = fv(qs(String(a.code)) + t.ref, t.value);
      const glAging = r2(aging.gl_ar_balance);
      varianceCells(row, 'variance', colL(5) + (8 + rows.length), colL(6) + (8 + rows.length), t.value, agingCodes.size > 1 ? glAging : gl);
      if (agingCodes.size > 1) { row.gl = glAging; row.comments = '- Aging covers ' + [...agingCodes].join(', '); }
      if (nz(aging.opening_residual)) row.comments = '- ' + fmt(aging.opening_residual) + ' not itemized on the subledger';
      if (nz(r2(t.value - row.gl))) P.flags.push({ severity: 'exception', wp: '03 AR', message: 'A/R aging ' + fmt(t.value) + ' does not agree to GL ' + a.code + ' ' + fmt(row.gl) + '.' });
    } else {
      const t = rollTab(wb, P, a, LS, 'basic');
      row.link = hyper(String(a.code), 'A1', String(a.code));
      row.balance = fv(signedRef(t), r2(t.value * t.sgn));
      varianceCells(row, 'variance', 'E' + (8 + rows.length), 'F' + (8 + rows.length), t.value * t.sgn, gl);
    }
    rows.push(row);
  }
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 15.4, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 37, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 14.3, kind: 'link' },
    { key: 'balance', header: 'Balance', width: 18.4, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 16, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 13, kind: 'money', total: true },
    { key: 'comments', header: 'Comments', width: 36, kind: 'text' },
  ], rows, 'Total Receivables');
  return { wb, summary: sumRows(rows) };
}
const sumRows = (rows) => ({ balance: r2(rows.reduce((s, r) => s + (Number(resultOf(r.balance)) || 0), 0)), gl: r2(rows.reduce((s, r) => s + (Number(r.gl) || 0), 0)) });

// A/R aging tab — the CL "A/R Aging Detail — built from GL" layout.
function agingTabAR(wb, P, a, LS, ag) {
  const ws = wb.addWorksheet(String(a.code), { views: [{ showGridLines: false }] });
  put(ws, 'A1', a.code + ' - ' + a.name, { fontOpts: { bold: true, italic: true } });
  put(ws, 'G4', hyper(LS, 'A1', 'Back to leadsheet'), { font: LINKF });
  [35, 18, 14, 14, 13, 16, 16, 16, 16, 16, 16, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  put(ws, 'A10', P.ent.name, { fontOpts: { bold: true } });
  put(ws, 'A11', 'A/R Aging Detail — built from GL ' + (ag.ar_account || a.code), { fontOpts: { italic: true } });
  put(ws, 'A12', 'As of ' + ag.as_of, { fontOpts: { italic: true } });
  const head = ['Customer', 'Invoice #', 'Invoice date', 'Due date', 'Days past due', 'Current', '1-30', '31-60', '61-90', '90+', 'GL', 'Amount'];
  head.forEach((h, i) => put(ws, colL(1 + i) + 14, h, { fontOpts: { bold: true }, border: i >= 5 ? { bottom: THIN } : undefined }));
  const BK = ['current', 'd1_30', 'd31_60', 'd61_90', 'd90_plus'];
  let r = 15; const subRows = [];
  for (const cu of ag.rows) {
    const f = r;
    for (const x of ag.detail.filter((d) => d.customer === cu.customer)) {
      put(ws, 'A' + r, cu.customer); put(ws, 'B' + r, String(x.invoice_num || '')); put(ws, 'C' + r, x.invoice_date || ''); put(ws, 'D' + r, x.due_date || '');
      put(ws, 'E' + r, x.days_past_due);
      BK.forEach((b, i) => { if (x.bucket === b) money(ws, colL(6 + i) + r, r2(x.open)); });
      money(ws, 'L' + r, r2(x.open)); r++;
    }
    put(ws, 'A' + r, 'Total ' + cu.customer, { fontOpts: { bold: true } });
    ['F', 'G', 'H', 'I', 'J', 'L'].forEach((c, i) => money(ws, c + r, fv('SUM(' + c + f + ':' + c + (r - 1) + ')', r2(i < 5 ? cu[BK[i]] : cu.total)), { fontOpts: { bold: true }, border: { top: THIN } }));
    subRows.push(r); r++;
  }
  if ((ag.gl_rows || []).length) {
    r++; put(ws, 'A' + r, 'GL ENTRIES (imported / manual — not aged)', { fontOpts: { bold: true } }); r++;
    const f = r;
    for (const g of ag.gl_rows) { put(ws, 'A' + r, g.memo || 'GL detail import'); put(ws, 'B' + r, g.entry_num != null ? 'JE-' + pad4(g.entry_num) : ''); put(ws, 'C' + r, g.date || ''); money(ws, 'K' + r, r2(g.amount)); money(ws, 'L' + r, r2(g.amount)); r++; }
    put(ws, 'A' + r, 'Total GL Entries', { fontOpts: { bold: true } });
    ['K', 'L'].forEach((c) => money(ws, c + r, fv('SUM(' + c + f + ':' + c + (r - 1) + ')', r2(ag.gl_total)), { fontOpts: { bold: true }, border: { top: THIN } }));
    subRows.push(r); r++;
  }
  const t = ag.totals, tot = r;
  put(ws, 'A' + tot, 'TOTAL', { fontOpts: { bold: true } });
  [['F', t.current], ['G', t.d1_30], ['H', t.d31_60], ['I', t.d61_90], ['J', t.d90_plus], ['K', t.gl], ['L', t.total]].forEach(([c, v]) =>
    money(ws, c + tot, subRows.length ? fv('SUM(' + subRows.map((x) => c + x).join(',') + ')', r2(v)) : 0, { fontOpts: { bold: true }, border: { bottom: DBL } }));
  put(ws, 'A' + (tot + 1), 'Reconciliation vs GL ' + (ag.ar_account || a.code) + ' (' + fmt(ag.gl_ar_balance) + ')');
  money(ws, 'L' + (tot + 1), fv('L' + tot + '-' + r2(ag.gl_ar_balance), r2(t.total - ag.gl_ar_balance)));
  put(ws, 'F6', 'AR balance'); money(ws, 'G6', fv('L' + tot, r2(t.total)), { fill: ENTRY });
  return { ref: 'G6', value: r2(t.total) };
}

// ─── 05 Prepaid / 09 Accrued / 11 Other Liabilities (schedule leadsheets) ─────
function buildSchedulePackage(P, cat, LS, title, kind, totalLabel) {
  const wb = newWb();
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, title);
  const rows = [];
  for (const a of P.byCat[cat]) {
    const gl = P.L.dr(a.code, P.m.end);
    const t = scheduleTab(wb, P, a, LS, kind);
    const row = { code: Number(a.code) || a.code, name: a.name, link: hyper(String(a.code), 'A1', String(a.code)), balance: fv(qs(t.sheet) + t.ref, t.value), gl };
    varianceCells(row, 'variance', 'E' + (8 + rows.length), 'F' + (8 + rows.length), t.value, gl);
    if (!nz(gl) && !nz(t.value)) row.comments = '- Zero balance';
    rows.push(row);
  }
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 11.9, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 34.4, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 15.4, kind: 'link' },
    { key: 'balance', header: 'Balance', width: 16.1, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 16, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 13, kind: 'money', total: true },
    { key: 'comments', header: 'Comments', width: 40, kind: 'text' },
  ], rows, totalLabel);
  return { wb, summary: sumRows(rows) };
}

// ─── 06 Other Assets / 12 Debt / 14 Intercompany (roll-forward leadsheets) ────
function buildOtherAssets(P) {
  const wb = newWb(), LS = 'Other Assets Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'OTHER ASSETS LEAD SHEET');
  const rows = [];
  for (const a of P.byCat.otherAssets) {
    const gl = P.L.dr(a.code, P.m.end);
    const t = rollTab(wb, P, a, LS, 'oa');
    const row = { code: Number(a.code) || a.code, name: a.name, link: hyper(t.sheet, 'A1', t.sheet), balance: fv(qs(t.sheet) + t.ref, t.value), gl };
    varianceCells(row, 'variance', 'E' + (8 + rows.length), 'F' + (8 + rows.length), t.value, gl);
    if (!nz(gl)) row.comments = '- Zero balance at month end';
    rows.push(row);
  }
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 14.4, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 40.4, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 13, kind: 'link' },
    { key: 'balance', header: 'Balance\nDr (Cr)', width: 16.1, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 16, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 13, kind: 'money', total: true },
    { key: 'comments', header: 'Comments', width: 60, kind: 'text' },
  ], rows, 'Total Other Assets');
  const note = 8 + rows.length + 2;
  put(ls, 'B' + note, '* Direct and indirect costs incurred during the period related to the real estate development are capitalized; each tab shows the prior balance and the current month GL detail.', { fontOpts: { italic: true, size: 9 } });
  return { wb, summary: sumRows(rows) };
}

function buildDebt(P) {
  const wb = newWb(), LS = 'Loan Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'DEBT LEADSHEET');
  put(ls, 'G3', ' Entry Cell', { fill: ENTRY, border: BOX });
  put(ls, 'H6', 'Current Portion (CP) / Long Term Portion (LTP) Breakout', { fontOpts: { bold: true } });
  const rows = [];
  P.byCat.debt.forEach((a) => {
    const r = 8 + rows.length;
    const gl = P.L.dr(a.code, P.m.end);
    const t = rollTab(wb, P, a, LS, 'debt');
    const bal = r2(t.value * t.sgn);
    const row = { code: Number(a.code) || a.code, name: a.name, link: hyper(t.sheet, 'A1', t.sheet), balance: fv(signedRef(t), bal), gl };
    varianceCells(row, 'variance', 'E' + r, 'F' + r, bal, gl);
    // Short-term / interest payable accounts default fully current.
    const cur = /short[- ]term|interest payable/i.test(a.name) ? bal : 0;
    row.cp = cur; row.ltp = fv('E' + r + '-H' + r, r2(bal - cur));
    rows.push(row);
  });
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 11.9, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 38, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 13, kind: 'link' },
    { key: 'balance', header: 'Balance', width: 16.1, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 16, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 13, kind: 'money', total: true },
    { key: 'cp', header: 'CP LTD Balance', width: 16, kind: 'money', total: true, entry: true },
    { key: 'ltp', header: 'LTP LTD Balance', width: 16, kind: 'money', total: true },
    { key: 'comments', header: 'Comments', width: 40, kind: 'text' },
  ], rows, 'Total Loans');
  return { wb, summary: sumRows(rows) };
}

function selfPattern(name) {
  const stop = new Set(['the', 'a', 'an', 'llc', 'lp', 'inc', 'fund', 'holdings', 'holding', 'company', 'co', 'partners', 'property', 'owner', 'clr', 'county', 'line', 'rail', 'of', 'and', 'i', 'ii', 'iii']);
  const words = String(name || '').replace(/[.,&]/g, ' ').split(/\s+/).filter((w) => w.length >= 3 && !stop.has(w.toLowerCase()));
  const w = words.sort((p, q) => q.length - p.length)[0];
  return w ? new RegExp('\\b' + w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i') : null;
}

function buildIntercompany(P) {
  const wb = newWb(), LS = 'Intercompany Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'INTERCOMPANY LEADSHEET');
  put(ls, 'C5', 'Entity:'); put(ls, 'D5', P.ent.name, { fontOpts: { bold: true } });
  const tie = wb.addWorksheet('Tie Out', { views: [{ showGridLines: false }] });
  const entities = P.db.prepare('SELECT id, name FROM entities').all();
  const selfRe = selfPattern(P.ent.name);
  const rows = [], groups = new Map();
  for (const a of P.byCat.interco) {
    const gl = P.L.dr(a.code, P.m.end);
    const t = rollTab(wb, P, a, LS, 'ic');
    const bal = r2(t.value * t.sgn);
    const cp = resolveCounterparty(entities, a.name, P.eid);
    const key = cp ? 'e' + cp.id : 'n' + a.code;
    if (!groups.has(key)) groups.set(key, { cp, accts: [], net: 0 });
    const grp = groups.get(key); grp.accts.push(a.code); grp.net = r2(grp.net + bal);
    const row = { code: Number(a.code) || a.code, name: a.name, link: hyper(t.sheet, 'A1', t.sheet), balance: fv(signedRef(t), bal), gl, cpname: cp ? cp.name : '' };
    varianceCells(row, 'variance', 'E' + (8 + rows.length), 'F' + (8 + rows.length), bal, gl);
    rows.push(row);
  }
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 11.9, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 40, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 13, kind: 'link' },
    { key: 'balance', header: 'Balance\nDr (Cr)', width: 16.1, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 16, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 13, kind: 'money', total: true },
    { key: 'cpname', header: 'Counterparty (CL entity)', width: 34, kind: 'text' },
    { key: 'comments', header: 'Comments', width: 40, kind: 'text' },
  ], rows, 'Total Intercompany');
  put(ls, 'D' + (8 + rows.length + 2), hyper('Tie Out', 'A1', 'Tie Out'), { font: LINKF });
  put(ls, 'E' + (8 + rows.length + 2), '← each counterparty tied to its own CloudLedger ledger', { fontOpts: { italic: true, size: 9 } });

  // Tie Out: our net position per counterparty vs what the counterparty's own
  // ledger shows it owes us.
  tabTitle(tie, 'Intercompany Tie Out', LS, 'D1');
  put(tie, 'A3', 'Debit = receivable, Credit = payable. The counterparty column is read live from that entity\'s own CloudLedger ledger (accounts naming ' + (selfRe ? selfRe.source.replace(/\\b/g, '') : P.ent.name) + ').', { fontOpts: { italic: true, size: 9 } });
  const H = ['Counterparty', 'Our Accounts', 'Per ' + P.display + ' Dr (Cr)', 'Per Counterparty Ledger Dr (Cr)', 'Difference', 'Status'];
  [34, 22, 20, 22, 16, 60].forEach((w, i) => (tie.getColumn(i + 1).width = w));
  H.forEach((h, i) => put(tie, colL(1 + i) + 5, h, { fontOpts: { bold: true }, border: { bottom: THIN }, align: { horizontal: i >= 2 && i <= 4 ? 'center' : 'left', wrapText: true } }));
  let r = 6;
  for (const grp of groups.values()) {
    let cpNet = null, status, sev = null;
    if (grp.cp && selfRe) {
      const mir = counterpartyNet(P.computeBalances, grp.cp.id, selfRe, P.m.end);
      if (mir.legs.length) {
        cpNet = mir.net; // what they owe us (their payable +)
        const diff = r2(grp.net - cpNet);
        status = nz(diff) ? 'Does not mirror — ' + mir.legs.map((l) => l.code + ' ' + l.name).join('; ') : 'Mirrors ' + mir.legs.map((l) => l.code).join(', ');
        if (nz(diff)) sev = 'exception';
      } else {
        status = nz(grp.net) ? 'No account naming us on ' + grp.cp.name + '\'s ledger (may be tracked by location/class) — confirm manually' : 'Zero balance';
        if (nz(grp.net)) sev = 'review';
      }
    } else {
      status = nz(grp.net) ? 'Counterparty is not a CloudLedger entity — confirm the balance to their statement' : 'Zero balance';
      if (nz(grp.net)) sev = 'review';
    }
    put(tie, 'A' + r, grp.cp ? grp.cp.name : (P.L.info.get(String(grp.accts[0])) || {}).name || grp.accts[0]);
    put(tie, 'B' + r, grp.accts.join(', '));
    const ourF = grp.accts.map((c) => { const i = rows.findIndex((x) => String(x.code) === String(c)); return qs(LS) + 'E' + (8 + i); }).join('+');
    money(tie, 'C' + r, fv(ourF, grp.net));
    money(tie, 'D' + r, cpNet);
    if (cpNet != null) money(tie, 'E' + r, fv('C' + r + '-D' + r, r2(grp.net - cpNet)), { font: nz(grp.net - cpNet) ? REDF : F() });
    put(tie, 'F' + r, status, { font: sev === 'exception' ? REDF : F() });
    if (sev) P.flags.push({ severity: sev, wp: '14 Intercompany', message: (grp.cp ? grp.cp.name : 'Accounts ' + grp.accts.join(', ')) + ': ' + fmt(grp.net) + ' — ' + status + '.' });
    r++;
  }
  put(tie, 'A' + r, 'Total', { fontOpts: { bold: true }, border: { top: THIN } });
  money(tie, 'C' + r, fv('SUM(C6:C' + (r - 1) + ')', r2([...groups.values()].reduce((s, g2) => s + g2.net, 0))), { fontOpts: { bold: true }, border: { top: THIN } });
  return { wb, summary: sumRows(rows) };
}

// ─── 07 Fixed Assets ──────────────────────────────────────────────────────────
function defaultLife(name) {
  const n = String(name || '').toLowerCase();
  if (/land improvement/.test(n)) return 20;
  if (/\bland\b/.test(n)) return 0;
  if (/crossing/.test(n)) return 20;
  if (/track|railway|\brail\b/.test(n)) return 30;
  if (/building/.test(n)) return 39;
  if (/equipment|machinery/.test(n)) return 7;
  if (/fixture|furniture|vehicle|computer/.test(n)) return 5;
  if (/improvement/.test(n)) return 15;
  return null;
}
const isDepAcct = (a) => /depreciation|amortization|acc(um)?\.?\s*dep/i.test(a.name) || (parseInt(a.code, 10) >= 16000 && parseInt(a.code, 10) <= 16999);

function buildFixedAssets(P) {
  const wb = newWb(), LS = 'Fixed Assets Leadsheet', FS = 'Fixed Asset Schedule';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  const ws = wb.addWorksheet(FS, { views: [{ showGridLines: false, state: 'frozen', xSplit: 10, ySplit: 9 }] });
  titleBlock(ls, P, 'FIXED ASSETS LEAD SHEET');
  put(ls, 'H3', 'Capitalization Policy:'); put(ls, 'I3', ' >$5,000 for single item', { fontOpts: { bold: true }, fill: ENTRY, border: BOX });
  const assets = P.byCat.fixed.filter((a) => !isDepAcct(a));
  const deps = P.byCat.fixed.filter(isDepAcct);
  // asset -> accumulated depreciation account
  const depFor = (a, life) => {
    if (P.profile.depMap[a.code]) return deps.find((d) => d.code === String(P.profile.depMap[a.code])) || null;
    if (!life || !deps.length) return null;
    const toks = String(a.name).toLowerCase().split(/[^a-z]+/).filter((t) => t.length >= 4 && !['land', 'asset', 'assets'].includes(t));
    const hit = deps.find((d) => toks.some((t) => String(d.name).toLowerCase().includes(t)));
    return hit || deps.slice().sort((p, q) => Math.abs(P.L.dr(q.code, P.m.end)) - Math.abs(P.L.dr(p.code, P.m.end)))[0];
  };
  // Schedule layout (FloQast columns): B Description C Asset # D Life E In service
  // F End date G Cost H Monthly I Beg accum  K..V months  X End accum  Y NBV
  [['A', 7.7], ['B', 38.1], ['C', 13.4], ['D', 8.7], ['E', 11.6], ['F', 11.6], ['G', 16.3], ['H', 12.4], ['I', 13.9], ['J', 2], ['W', 2], ['X', 15.6], ['Y', 15.6]].forEach(([c, w]) => (ws.getColumn(c).width = w));
  for (let k = 0; k < 12; k++) ws.getColumn(11 + k).width = 11.9;
  put(ws, 'A1', 'Fixed Asset Depreciation Schedule', { fontOpts: { bold: true, italic: true } });
  put(ws, 'H1', ' Entry Cell', { fill: ENTRY, border: BOX });
  put(ws, 'D2', hyper(LS, 'A1', 'Back to Leadsheet'), { font: LINKF });
  put(ws, 'A3', 'INSTRUCTIONS', { fontOpts: { bold: true, italic: true } });
  put(ws, 'A4', '1.'); put(ws, 'B4', 'Cost lines are the CloudLedger GL entries on each fixed-asset account (one row per journal entry).');
  put(ws, 'A5', '2.'); put(ws, 'B5', 'Useful life and in-service date are entry cells (defaults: straight-line, service starts the month after the entry date).');
  put(ws, 'A6', '3.'); put(ws, 'B6', 'Accumulated depreciation per this schedule is tied to the GL on the leadsheet.');
  put(ws, 'I7', 'Accumulated', { fontOpts: { bold: true } }); put(ws, 'X7', 'Accumulated', { fontOpts: { bold: true } });
  [['C8', 'Asset'], ['D8', 'Useful'], ['E8', 'Date In'], ['F8', 'Depr/Amort'], ['H8', 'Monthly'], ['I8', 'Depr/Amort'], ['K8', 'Depreciation/Amortization Expense'], ['X8', 'Depr/Amort']]
    .forEach(([c, v]) => put(ws, c, v, { fontOpts: { bold: true } }));
  [['B9', 'Description'], ['C9', '#'], ['D9', 'Life (Yrs)'], ['E9', 'Service'], ['F9', 'End Date'], ['G9', 'Cost'], ['H9', 'Depr/Amort'], ['I9', 'Beg Balance'], ['X9', 'End Balance'], ['Y9', 'NBV']]
    .forEach(([c, v]) => put(ws, c, v, { fontOpts: { bold: true }, border: { bottom: THIN } }));
  for (let k = 0; k < 12; k++) put(ws, colL(11 + k) + 9, fv('EOMONTH(' + qs(LS) + '$D$4,' + (k - 11) + ')', serial(P.m.fyEnds[k])), { fontOpts: { bold: true }, border: { bottom: THIN }, numFmt: 'mmm-yy', align: 'center' });
  const PYE = 'DATE(' + (P.m.y - 1) + ',12,31)';
  let r = 11;
  const lsRows = [];
  const monthsThru = (inSvc, thru) => { const [y1, m1] = inSvc.split('-').map(Number); const [y2, m2] = thru.split('-').map(Number); return (y2 - y1) * 12 + (m2 - m1) + 1; };
  for (const a of assets) {
    const lines = P.L.lines(a.code, null, P.m.end).filter((l) => nz(l.dr));
    const byEntry = new Map();
    for (const l of lines) { const e = byEntry.get(l.entry_id) || { date: l.date, memo: l.memo || l.description || '', amt: 0 }; e.amt = r2(e.amt + l.dr); byEntry.set(l.entry_id, e); }
    const items = [...byEntry.values()].filter((e) => nz(e.amt));
    let life = defaultLife(a.name);
    if (life == null) { life = 0; P.flags.push({ severity: 'review', wp: '07 Fixed Assets', message: a.code + ' ' + a.name + ': no default useful life — key it on the Fixed Asset Schedule.' }); }
    const dep = depFor(a, life);
    const hdr = r;
    put(ws, 'A' + hdr, fv(qs(LS) + '$B$' + (9 + lsRows.length), Number(a.code) || a.code), { font: F({ color: { argb: 'FFFFFFFF' }, bold: true }), fill: NAVY });
    put(ws, 'B' + hdr, fv(qs(LS) + '$C$' + (9 + lsRows.length), a.name), { font: F({ color: { argb: 'FFFFFFFF' }, bold: true }), fill: NAVY });
    r++;
    const first = r;
    const sums = { G: 0, I: 0, X: 0, Y: 0, m: Array(12).fill(0) };
    for (const it of items) {
      const [yy, mm, dd] = it.date.split('-').map(Number);
      const svc = dd === 1 ? it.date : new Date(Date.UTC(yy, mm, 1)).toISOString().slice(0, 10);
      const ent = { fill: ENTRY, border: { bottom: THIN } };
      put(ws, 'B' + r, it.memo, ent); put(ws, 'C' + r, '', ent);
      put(ws, 'D' + r, life, ent); dateCell(ws, 'E' + r, svc, ent);
      const endD = life ? new Date(Date.UTC(Number(svc.slice(0, 4)) + life, Number(svc.slice(5, 7)) - 1, Number(svc.slice(8, 10)) - 1)).toISOString().slice(0, 10) : null;
      put(ws, 'F' + r, fv('IF($D' + r + '=0,"",EDATE($E' + r + ',$D' + r + '*12)-1)', endD ? serial(endD) : ''), { numFmt: DATEF, align: 'center' });
      money(ws, 'G' + r, it.amt, ent);
      const monthly = life ? r2(it.amt / (life * 12)) : 0;
      money(ws, 'H' + r, fv('IF($D' + r + '=0,0,ROUND($G' + r + '/($D' + r + '*12),2))', monthly));
      const begM = life && svc <= P.m.priorYE ? Math.min(life * 12, monthsThru(svc, P.m.priorYE)) : 0;
      const beg = life ? r2(Math.sign(it.amt) * Math.min(Math.abs(it.amt), Math.abs(monthly * begM))) : 0;
      money(ws, 'I' + r, fv('IF(OR($D' + r + '=0,$E' + r + '>' + PYE + '),0,MIN(ABS($G' + r + '),ABS($H' + r + '*MIN($D' + r + '*12,(YEAR(' + PYE + ')-YEAR($E' + r + '))*12+MONTH(' + PYE + ')-MONTH($E' + r + ')+1)))*SIGN($G' + r + '))', beg), ent);
      let accum = beg;
      for (let k = 0; k < 12; k++) {
        const C = colL(11 + k), me = P.m.fyEnds[k], mStart = me.slice(0, 8) + '01';
        const on = life && me <= P.m.end && me >= svc && !(endD && mStart > endD);
        const v = on ? monthly : 0;
        money(ws, C + r, fv('IF(OR(' + C + '$9>' + qs(LS) + '$D$3,$D' + r + '=0,' + C + '$9<$E' + r + ',EOMONTH(' + C + '$9,-1)+1>$F' + r + '),0,$H' + r + ')', v));
        accum = r2(accum + v); sums.m[k] = r2(sums.m[k] + v);
      }
      money(ws, 'X' + r, fv('I' + r + '+SUM(K' + r + ':V' + r + ')', accum));
      money(ws, 'Y' + r, fv('G' + r + '-X' + r, r2(it.amt - accum)));
      sums.G = r2(sums.G + it.amt); sums.I = r2(sums.I + beg); sums.X = r2(sums.X + accum); sums.Y = r2(sums.Y + it.amt - accum);
      r++;
    }
    if (!items.length) { put(ws, 'B' + r, 'No cost on the GL.', { fontOpts: { italic: true } }); r++; }
    const tot = r;
    put(ws, 'B' + tot, fv('"Total " & $B' + hdr, 'Total ' + a.name), { fontOpts: { bold: true }, border: { top: THIN, bottom: THIN } });
    const sumF = (c, v) => money(ws, c + tot, fv('SUM(' + c + (hdr) + ':' + c + (tot - 1) + ')', v), { fontOpts: { bold: true }, border: { top: THIN, bottom: THIN } });
    sumF('G', sums.G); sumF('I', sums.I); for (let k = 0; k < 12; k++) sumF(colL(11 + k), sums.m[k]); sumF('X', sums.X); sumF('Y', sums.Y);
    r += 2;
    lsRows.push({ a, dep, costRef: 'G' + tot, cost: sums.G, depRef: 'X' + tot, accum: sums.X, gl: P.L.dr(a.code, P.m.end) });
  }
  // Leadsheet
  put(ls, 'B7', 'Fixed Asset Account Breakdown', { fontOpts: { bold: true }, fill: GREENH, border: { top: MED, bottom: MED } });
  ['C7', 'D7', 'E7', 'F7'].forEach((c) => { ls.getCell(c).fill = GREENH; ls.getCell(c).border = { top: MED, bottom: MED }; });
  put(ls, 'G7', 'Depreciation Account Breakdown', { fontOpts: { bold: true }, fill: REDH, border: { top: MED, bottom: MED } });
  ['H7', 'I7'].forEach((c) => { ls.getCell(c).fill = REDH; ls.getCell(c).border = { top: MED, bottom: MED }; });
  const rows = lsRows.map((x, i) => {
    const r0 = 9 + i;
    const row = {
      code: Number(x.a.code) || x.a.code, name: x.a.name,
      cost: fv(qs(FS) + x.costRef, x.cost), gl: x.gl,
      depcode: x.dep ? (Number(x.dep.code) || x.dep.code) : '', depname: x.dep ? x.dep.name : (x.accum ? '' : '— not depreciated'),
      dep: fv('-' + qs(FS) + x.depRef, r2(-x.accum)),
      net: fv('D' + r0 + '+I' + r0, r2(x.cost - x.accum)),
    };
    varianceCells(row, 'variance', 'D' + r0, 'E' + r0, x.cost, x.gl);
    return row;
  });
  const lt = leadTable(ls, [
    { key: 'code', header: 'Asset Acct. #', width: 14.4, kind: 'code' },
    { key: 'name', header: 'Asset Acct. Name', width: 38.1, kind: 'name' },
    { key: 'cost', header: 'Asset Cost', width: 15.4, kind: 'money', total: true },
    { key: 'gl', header: 'Asset GL Balance', width: 16.1, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 12, kind: 'money', total: true },
    { key: 'depcode', header: 'Depreciation Account #', width: 16, kind: 'text', align: 'center' },
    { key: 'depname', header: 'Depreciation Account Name', width: 42.7, kind: 'text' },
    { key: 'dep', header: 'Deprec. Balance (per schedule)', width: 19.3, kind: 'money', total: true },
    { key: 'net', header: 'Net Asset Amount', width: 16.9, kind: 'money', total: true },
  ], rows, 'Total Fixed Assets', 8);
  // Accumulated depreciation by account: schedule vs GL.
  let r2r = lt.totalRow + 3;
  put(ls, 'B' + r2r, 'Accumulated Depreciation — Schedule vs GL', { fontOpts: { bold: true }, fill: REDH, border: { top: MED, bottom: MED } });
  ['C', 'D', 'E', 'F'].forEach((c) => { ls.getCell(c + r2r).fill = REDH; ls.getCell(c + r2r).border = { top: MED, bottom: MED }; });
  r2r++;
  ['Account #', 'Account Name', 'Per Schedule', 'GL Balance', 'Variance'].forEach((h, i) => put(ls, colL(2 + i) + r2r, h, { border: { bottom: THIN }, align: i < 2 ? 'left' : 'center' }));
  r2r++;
  const depFirst = r2r;
  let depSch = 0, depGl = 0;
  for (const d of deps) {
    const sch = r2(-lsRows.filter((x) => x.dep && x.dep.code === d.code).reduce((s, x) => s + x.accum, 0));
    const gl = P.L.dr(d.code, P.m.end);
    put(ls, 'B' + r2r, Number(d.code) || d.code, { fill: ENTRY, border: BOX, align: 'center' });
    put(ls, 'C' + r2r, d.name, { fill: ENTRY, border: BOX });
    money(ls, 'D' + r2r, fv('SUMIF($G$' + lt.first + ':$G$' + lt.last + ',B' + r2r + ',$I$' + lt.first + ':$I$' + lt.last + ')', sch));
    money(ls, 'E' + r2r, gl);
    const v = r2(sch - gl);
    money(ls, 'F' + r2r, fv('D' + r2r + '-E' + r2r, v), { font: nz(v) ? REDF : F() });
    if (nz(v)) P.flags.push({ severity: 'review', wp: '07 Fixed Assets', message: d.code + ' ' + d.name + ': schedule depreciation ' + fmt(sch) + ' vs GL ' + fmt(gl) + ' (off by ' + fmt(v) + ') — check useful lives / in-service dates or the depreciation entry.' });
    depSch = r2(depSch + sch); depGl = r2(depGl + gl); r2r++;
  }
  put(ls, 'B' + r2r, 'Total Depreciation', { border: { top: THIN, bottom: MED } });
  ['D', 'E', 'F'].forEach((c, i) => money(ls, c + r2r, deps.length ? fv('SUM(' + c + depFirst + ':' + c + (r2r - 1) + ')', [depSch, depGl, r2(depSch - depGl)][i]) : 0, { fontOpts: { bold: true }, border: { top: THIN, bottom: MED } }));
  // Package summary: net book value per the schedule vs the GL (the difference
  // is the depreciation variance above).
  const bal = r2(P.byCat.fixed.reduce((s, a) => s + P.L.dr(a.code, P.m.end), 0));
  return { wb, summary: { balance: r2(lsRows.reduce((s, x) => s + x.cost, 0) + depSch), gl: bal } };
}

// ─── 08 Accounts Payable ──────────────────────────────────────────────────────
function buildAP(P) {
  const wb = newWb(), LS = 'Accounts Payable Leadsheet';
  const ls = wb.addWorksheet(LS, { views: [{ showGridLines: false }] });
  titleBlock(ls, P, 'ACCOUNTS PAYABLE LEAD SHEET');
  const rows = [];
  for (const a of P.byCat.ap) {
    const gl = P.L.dr(a.code, P.m.end);
    const ag = P.apAging(P.eid, P.m.end, a.code);
    const t = agingTabAP(wb, P, a, LS, ag);
    const row = { code: Number(a.code) || a.code, name: a.name, link: hyper(t.sheet, 'A1', t.sheet), balance: fv('-' + qs(t.sheet) + t.ref, r2(-t.value)), gl };
    varianceCells(row, 'variance', 'E' + (8 + rows.length), 'F' + (8 + rows.length), -t.value, gl);
    if (nz(ag.recon_diff)) P.flags.push({ severity: 'exception', wp: '08 AP', message: 'A/P aging for ' + a.code + ' is off from the GL by ' + fmt(ag.recon_diff) + '.' });
    if (ag.gl_entry_count) row.comments = '- ' + fmt(ag.grand_total.gl) + ' carried as un-aged GL entries';
    rows.push(row);
  }
  leadTable(ls, [
    { key: 'code', header: 'Account No.', width: 13, kind: 'code' },
    { key: 'name', header: 'Account Name', width: 32.9, kind: 'name' },
    { key: 'link', header: 'Worksheet', width: 11.7, kind: 'link' },
    { key: 'balance', header: 'Balance', width: 15.3, kind: 'money', total: true },
    { key: 'gl', header: 'GL Balance', width: 15.3, kind: 'money', total: true },
    { key: 'variance', header: 'Variance', width: 12, kind: 'money', total: true },
    { key: 'comments', header: 'Comments', width: 40, kind: 'text' },
  ], rows, 'Total Payables');
  return { wb, summary: sumRows(rows) };
}

function agingTabAP(wb, P, a, LS, ag) {
  const ws = wb.addWorksheet(String(a.code), { views: [{ showGridLines: false }] });
  put(ws, 'A1', a.code + ' - ' + a.name, { fontOpts: { bold: true, italic: true } });
  put(ws, 'E4', hyper(LS, 'A1', 'Back to leadsheet'), { font: LINKF });
  put(ws, 'G4', '*Aging Report should be based on GL posting date', { fontOpts: { italic: true } });
  [2.7, 12, 8, 22, 34, 12, 11, 14, 14, 14, 14, 14, 14, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  put(ws, 'B11', P.ent.name, { fontOpts: { bold: true } });
  put(ws, 'B12', 'A/P Aging Detail — built from GL ' + ag.ap_account, { fontOpts: { italic: true } });
  put(ws, 'B13', 'As of ' + ag.as_of, { fontOpts: { italic: true } });
  const head = ['Date', 'Type', 'Num', 'Vendor', 'Due Date', 'Past Due (days)', 'Current', '1-30', '31-60', '61-90', '91+', 'GL', 'Amount'];
  head.forEach((h, i) => put(ws, colL(2 + i) + 15, h, { fontOpts: { bold: true }, border: i >= 6 ? { bottom: THIN } : undefined }));
  const BK = ['current', 'd1_30', 'd31_60', 'd61_90', 'd91_plus'];
  let r = 16; const subRows = [];
  for (const v of ag.vendors) {
    const f = r;
    for (const x of v.rows) {
      put(ws, 'B' + r, x.date); put(ws, 'C' + r, x.type || 'Bill'); put(ws, 'D' + r, String(x.num || '')); put(ws, 'E' + r, v.vendor);
      put(ws, 'F' + r, x.due_date || x.date); put(ws, 'G' + r, x.past_due_days);
      BK.forEach((b, i) => { if (x.bucket === b) money(ws, colL(8 + i) + r, r2(x.amount)); });
      money(ws, 'N' + r, r2(x.amount)); r++;
    }
    put(ws, 'B' + r, 'Total ' + v.vendor, { fontOpts: { bold: true } });
    ['H', 'I', 'J', 'K', 'L', 'N'].forEach((c, i) => money(ws, c + r, fv('SUM(' + c + f + ':' + c + (r - 1) + ')', r2(i < 5 ? v.subtotal[BK[i]] : v.subtotal.total)), { fontOpts: { bold: true }, border: { top: THIN } }));
    subRows.push(r); r++;
  }
  if ((ag.gl_rows || []).length) {
    r++; put(ws, 'B' + r, 'GL ENTRIES (imported / manual — not aged)', { fontOpts: { bold: true } }); r++;
    const f = r;
    for (const g of ag.gl_rows) { put(ws, 'B' + r, g.date || ''); put(ws, 'C' + r, 'Journal'); put(ws, 'D' + r, g.entry_num != null ? 'JE-' + pad4(g.entry_num) : ''); put(ws, 'E' + r, g.memo || g.description || ''); money(ws, 'M' + r, r2(g.amount)); money(ws, 'N' + r, r2(g.amount)); r++; }
    put(ws, 'B' + r, 'Total GL Entries', { fontOpts: { bold: true } });
    ['M', 'N'].forEach((c) => money(ws, c + r, fv('SUM(' + c + f + ':' + c + (r - 1) + ')', r2(ag.gl_total)), { fontOpts: { bold: true }, border: { top: THIN } }));
    subRows.push(r); r++;
  }
  const t = ag.grand_total, tot = r;
  put(ws, 'B' + tot, 'TOTAL', { fontOpts: { bold: true } });
  [['H', t.current], ['I', t.d1_30], ['J', t.d31_60], ['K', t.d61_90], ['L', t.d91_plus], ['M', t.gl], ['N', t.total]].forEach(([c, v]) =>
    money(ws, c + tot, subRows.length ? fv('SUM(' + subRows.map((x) => c + x).join(',') + ')', r2(v)) : 0, { fontOpts: { bold: true }, border: { bottom: DBL } }));
  put(ws, 'B' + (tot + 1), 'Reconciliation vs GL ' + ag.ap_account + ' (' + fmt(ag.gl_balance) + ')');
  money(ws, 'N' + (tot + 1), fv('N' + tot + '-' + r2(ag.gl_balance), r2(t.total - ag.gl_balance)));
  money(ws, 'E6', fv('N' + tot, r2(t.total)), { fill: ENTRY });
  return { sheet: String(a.code), ref: 'E6', value: r2(t.total) };
}

// ─── 15 Equity Rollforward ────────────────────────────────────────────────────
function buildEquity(P) {
  const wb = newWb(), ER = 'Equity Rollforward', NI = 'YTD NI';
  const ws = wb.addWorksheet(ER, { views: [{ showGridLines: false, state: 'frozen', xSplit: 6, ySplit: 13 }] });
  const ni = wb.addWorksheet(NI, { views: [{ showGridLines: false }] });
  const db = P.db, eid = P.eid, m = P.m;
  put(ws, 'C1', P.display, { fontOpts: { bold: true } });
  put(ws, 'C2', 'EQUITY ROLLFORWARD', { fontOpts: { bold: true } });
  put(ws, 'C3', 'Month Ended:'); dateCell(ws, 'D3', m.end, { fontOpts: { bold: true }, fill: ENTRY });
  put(ws, 'C4', 'Year Ended:'); dateCell(ws, 'D4', m.yearEnd, { fontOpts: { bold: true }, fill: ENTRY });
  put(ws, 'A6', 'INSTRUCTIONS', { fontOpts: { bold: true } });
  put(ws, 'A7', '1.'); put(ws, 'B7', 'Balances are Dr (Cr): capital shows negative, a net loss positive. Pulled from the CloudLedger GL.');
  put(ws, 'A8', '2.'); put(ws, 'B8', 'The Equity Closing Entry column moves prior-period net income into retained earnings (nets to $0).');
  put(ws, 'A9', '3.'); put(ws, 'B9', 'Ending equity is tied to the balance sheet (assets less liabilities) at the bottom.');
  [['A', 3.1], ['B', 64], ['C', 12], ['D', 17], ['E', 17], ['F', 17]].forEach(([c, w]) => (ws.getColumn(c).width = w));
  for (let k = 0; k < 13; k++) ws.getColumn(7 + k).width = 15.4;
  ws.getColumn(19).width = 19.1;
  [['D12', 'Prior Year End'], ['E12', 'Equity\nClosing Entry\n(should net to $0)'], ['F12', 'Opening Equity'], ['G12', 'Current Year']].forEach(([c, v]) =>
    put(ws, c, v, { fontOpts: { bold: true }, align: { horizontal: 'center', wrapText: true } }));
  ws.getRow(12).height = 45;
  put(ws, 'D13', fv('EOMONTH($D$4,-12)', serial(m.priorYE)), { fontOpts: { bold: true }, numFmt: DATEF, align: 'center' });
  put(ws, 'E13', fv('D13+1', serial(m.yearStart)), { fontOpts: { bold: true }, numFmt: DATEF, align: 'center' });
  put(ws, 'F13', fv('D13+1', serial(m.yearStart)), { fontOpts: { bold: true }, numFmt: DATEF, align: 'center' });
  for (let k = 0; k < 12; k++) put(ws, colL(7 + k) + 13, fv('EOMONTH($D$4,' + (k - 11) + ')', serial(m.fyEnds[k])), { fontOpts: { bold: true }, numFmt: DATEF, align: 'center', fill: k + 1 === m.mo ? YELLOW : undefined });
  put(ws, 'S13', 'YTD', { fontOpts: { bold: true }, align: 'center' });

  const eq = P.byCat.equity;
  const pl = db.prepare("SELECT COALESCE(SUM(jl.debit - jl.credit),0) AS v FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id "
    + "JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code WHERE je.entity_id = ? AND je.date <= ? AND a.type IN ('Revenue','Expense')");
  const priorPL = r2(pl.get(eid, m.priorYE).v);
  const actStmt = db.prepare("SELECT jl.account_code AS code, CAST(substr(je.date,6,2) AS INTEGER) AS mm, SUM(jl.debit - jl.credit) AS amt FROM journal_lines jl "
    + "JOIN journal_entries je ON je.id = jl.entry_id JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code "
    + "WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? AND a.type = 'Equity' GROUP BY jl.account_code, mm");
  const act = new Map();
  for (const x of actStmt.all(eid, m.yearStart, m.end)) { if (!act.has(String(x.code))) act.set(String(x.code), Array(12).fill(0)); act.get(String(x.code))[x.mm - 1] = r2(x.amt); }
  let reIdx = eq.findIndex((a) => String(a.code) === '39000');
  if (reIdx < 0) reIdx = eq.findIndex((a) => /retained earnings/i.test(a.name));
  const lines = eq.map((a) => ({ label: a.code + ' - ' + a.name, d: P.L.dr(a.code, m.priorYE), e: 0, act: act.get(String(a.code)) || Array(12).fill(0) }));
  if (reIdx < 0) { lines.push({ label: 'Retained Earnings', d: 0, e: 0, act: Array(12).fill(0) }); reIdx = lines.length - 1; }
  lines[reIdx].e = priorPL;

  put(ws, 'B15', 'Partnership / LLC', { fontOpts: { bold: true } });
  put(ws, 'B16', 'Beginning Partners\' (Capital)/Deficit:');
  let r = 17; const first = r;
  const moCols = (k) => colL(7 + k);
  const totals = { d: 0, e: 0, f: 0, m: Array(12).fill(0), s: 0 };
  for (const x of lines) {
    put(ws, 'B' + r, x.label, { fill: ENTRY });
    money(ws, 'D' + r, x.d, { fill: ENTRY });
    money(ws, 'E' + r, x.e);
    money(ws, 'F' + r, fv('SUM(D' + r + ':E' + r + ')', r2(x.d + x.e)));
    let ytd = r2(x.d + x.e);
    for (let k = 0; k < 12; k++) {
      const v = k < m.mo ? x.act[k] : null;
      money(ws, moCols(k) + r, v, { fill: k < m.mo ? ENTRY : undefined });
      if (v) { ytd = r2(ytd + v); totals.m[k] = r2(totals.m[k] + v); }
    }
    money(ws, 'S' + r, fv('SUM(F' + r + ':R' + r + ')', ytd));
    totals.d = r2(totals.d + x.d); totals.e = r2(totals.e + x.e); totals.f = r2(totals.f + x.d + x.e); totals.s = r2(totals.s + ytd);
    r++;
  }
  // Prior-period net (income)/loss, closed to retained earnings by the E column.
  put(ws, 'B' + r, 'Prior Period Net (Income)/Loss', { fill: ENTRY });
  money(ws, 'D' + r, priorPL, { fill: ENTRY }); money(ws, 'E' + r, fv('-D' + r, -priorPL)); money(ws, 'F' + r, fv('SUM(D' + r + ':E' + r + ')', 0));
  money(ws, 'S' + r, fv('SUM(F' + r + ':R' + r + ')', 0));
  totals.d = r2(totals.d + priorPL); totals.e = r2(totals.e - priorPL);
  r++;
  const last = r - 1, begRow = r;
  put(ws, 'B' + begRow, 'Beginning Partners\' (Capital)/Deficit', { fontOpts: { bold: true } });
  const sumCol = (c, v) => money(ws, c + begRow, fv('SUM(' + c + first + ':' + c + last + ')', v), { fontOpts: { bold: true }, border: { top: THIN } });
  sumCol('D', totals.d); sumCol('E', totals.e); sumCol('F', totals.f);
  for (let k = 0; k < 12; k++) sumCol(moCols(k), totals.m[k]);
  sumCol('S', totals.s);
  // Net income by month (Dr/Cr: a loss is positive) from the YTD NI tab.
  const niRows = db.prepare("SELECT CAST(substr(je.date,6,2) AS INTEGER) AS mm, COALESCE(dl.name,'(No location)') AS loc, SUM(jl.credit - jl.debit) AS ni "
    + "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code "
    + "LEFT JOIN dim_locations dl ON dl.id = jl.location_id WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? AND a.type IN ('Revenue','Expense') "
    + 'GROUP BY mm, loc').all(eid, m.yearStart, m.end);
  const locs = [...new Set(niRows.map((x) => x.loc))].sort();
  const niBy = new Map(locs.map((l) => [l, Array(12).fill(0)]));
  for (const x of niRows) niBy.get(x.loc)[x.mm - 1] = r2(x.ni);
  const niTot = Array(12).fill(0).map((_, k) => r2(locs.reduce((s, l) => s + niBy.get(l)[k], 0)));
  const niTotRow = 3 + Math.max(locs.length, 1) + 1;
  const niR = begRow + 2;
  put(ws, 'B' + niR, 'Current Period Net (Income)/Loss', { fill: ENTRY });
  let niYtd = 0;
  for (let k = 0; k < 12; k++) {
    const v = k < m.mo ? r2(-niTot[k]) : null;
    money(ws, moCols(k) + niR, k < m.mo ? fv('-' + qs(NI) + colL(2 + k) + niTotRow, v) : null, { fill: k < m.mo ? ENTRY : undefined });
    if (v) niYtd = r2(niYtd + v);
  }
  money(ws, 'S' + niR, fv('SUM(F' + niR + ':R' + niR + ')', niYtd));
  // Ending (equity)/deficit by month, cumulative.
  const endR = niR + 2;
  put(ws, 'B' + endR, 'Ending (Equity)/Deficit', { fontOpts: { bold: true } });
  let run = totals.f;
  for (let k = 0; k < 12; k++) {
    const C = moCols(k);
    if (k < m.mo) {
      run = r2(run + totals.m[k] + r2(-niTot[k]));
      money(ws, C + endR, fv('$F$' + begRow + '+SUM($G$' + begRow + ':' + C + begRow + ')+SUM($G$' + niR + ':' + C + niR + ')', run), { fontOpts: { bold: true }, border: { top: THIN, bottom: DBL }, fill: k + 1 === m.mo ? YELLOW : undefined });
    }
  }
  const endEq = run;
  money(ws, 'S' + endR, fv(moCols(m.mo - 1) + endR, endEq), { fontOpts: { bold: true }, border: { top: THIN, bottom: DBL } });
  // Tie to the balance sheet: equity = -(assets + liabilities) in Dr/Cr terms.
  let bsAL = 0;
  for (const cat of Object.keys(P.byCat)) if (cat !== 'equity') for (const a of P.byCat[cat]) bsAL = r2(bsAL + P.L.dr(a.code, m.end));
  const tieR = endR + 2;
  put(ws, 'B' + tieR, 'Net assets per balance sheet (assets less liabilities), shown Dr (Cr)');
  money(ws, 'S' + tieR, r2(-bsAL));
  put(ws, 'B' + (tieR + 1), 'Variance', { fontOpts: { bold: true } });
  const v = r2(endEq - (-bsAL));
  money(ws, 'S' + (tieR + 1), fv('S' + endR + '-S' + tieR, v), { fontOpts: { bold: true, color: { argb: nz(v) ? 'FFFF0000' : 'FF000000' } } });
  if (nz(v)) P.flags.push({ severity: 'exception', wp: '15 Equity', message: 'Equity roll-forward ending ' + fmt(endEq) + ' does not agree to net assets ' + fmt(-bsAL) + ' (off by ' + fmt(v) + ') — the trial balance may not balance.' });

  // YTD NI tab (net income, income positive, by location and month).
  [['A', 40]].forEach(([c, w]) => (ni.getColumn(c).width = w));
  for (let k = 0; k < 13; k++) ni.getColumn(2 + k).width = 13;
  put(ni, 'A1', 'Location', { fontOpts: { bold: true } });
  MON3.forEach((mn, k) => put(ni, colL(2 + k) + '1', mn, { fontOpts: { bold: true }, align: 'center' }));
  put(ni, 'N1', 'Total', { fontOpts: { bold: true }, align: 'center' });
  let rr = 3;
  for (const l of (locs.length ? locs : ['(No P&L activity)'])) {
    put(ni, 'A' + rr, l);
    const arr = niBy.get(l) || Array(12).fill(0);
    for (let k = 0; k < 12; k++) money(ni, colL(2 + k) + rr, k < m.mo ? arr[k] : null);
    money(ni, 'N' + rr, fv('SUM(B' + rr + ':M' + rr + ')', r2(arr.reduce((s, x) => s + x, 0)))); rr++;
  }
  const tr = niTotRow;
  for (let k = 0; k < 13; k++) money(ni, colL(2 + k) + tr, fv('SUM(' + colL(2 + k) + '3:' + colL(2 + k) + (tr - 1) + ')', k < 12 ? niTot[k] : r2(niTot.reduce((s, x) => s + x, 0))), { fontOpts: { bold: true }, border: { top: THIN }, fill: k === 12 ? YELLOW : undefined });
  put(ni, 'A' + (tr + 2), 'Check (should be 0)');
  for (let k = 0; k < m.mo; k++) money(ni, colL(2 + k) + (tr + 2), fv(colL(2 + k) + tr + '+' + qs(ER) + moCols(k) + niR, 0));
  return { wb, summary: { balance: endEq, gl: r2(-bsAL) } };
}

// ─── Package ──────────────────────────────────────────────────────────────────
const BOOKS = [
  { n: '01', slug: 'Cash_and_Cash_Equivalents', cat: 'cash', title: 'Cash and Cash Equivalents', build: buildCash },
  { n: '03', slug: 'Accounts_Receivable_Leadsheet', cat: 'ar', title: 'Accounts Receivable', build: buildAR },
  { n: '05', slug: 'Prepaid_Expenses_Leadsheet', cat: 'prepaid', title: 'Prepaid Expenses', build: (P) => buildSchedulePackage(P, 'prepaid', 'Prepaid Expenses Leadsheet', 'PREPAID EXPENSES LEADSHEET', 'prepaid', 'Total Prepaid Expenses') },
  { n: '06', slug: 'Other_Assets_Leadsheet', cat: 'otherAssets', title: 'Other Assets', build: buildOtherAssets },
  { n: '07', slug: 'Fixed_Asset_Leadsheet', cat: 'fixed', title: 'Fixed Assets', build: buildFixedAssets },
  { n: '08', slug: 'Accounts_Payable_Leadsheet', cat: 'ap', title: 'Accounts Payable', build: buildAP },
  { n: '09', slug: 'Accrued_Expenses_Leadsheet', cat: 'accrued', title: 'Accrued Expenses', build: (P) => buildSchedulePackage(P, 'accrued', 'Accrued Expenses Leadsheet', 'ACCRUED LIABILITIES LEADSHEET', 'accrued', 'Total Accrued Expenses') },
  { n: '11', slug: 'Other_Liabilities_Leadsheet', cat: 'otherLiab', title: 'Other Liabilities', build: (P) => buildSchedulePackage(P, 'otherLiab', 'Other Liabilities Leadsheet', 'OTHER LIABILITIES LEADSHEET', 'accrued', 'Total Other Liabilities') },
  { n: '12', slug: 'Debt_Leadsheet', cat: 'debt', title: 'Debt', build: buildDebt },
  { n: '14', slug: 'Intercompany_Standard', cat: 'interco', title: 'Intercompany', build: buildIntercompany },
  { n: '15', slug: 'Equity_Rollforward', cat: 'equity', title: 'Equity Rollforward', build: buildEquity },
];

async function buildPackage(ctx, eid, monthEnd) {
  const { db } = ctx;
  const m = resolveMonth(monthEnd);
  const ent = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(eid);
  if (!ent) throw new Error('Entity ' + eid + ' not found');
  const L = makeLedger(db, eid);
  const profile = profileFor(ent);
  const byCat = { cash: [], ar: [], prepaid: [], otherAssets: [], fixed: [], ap: [], accrued: [], otherLiab: [], debt: [], interco: [], equity: [] };
  const flags = [];
  for (const code of L.activeCodes(m.end)) {
    const a = L.info.get(code);
    if (!a) { flags.push({ severity: 'exception', wp: 'Coverage', message: 'Account ' + code + ' has GL activity but is not in the chart of accounts, so it is on no leadsheet.' }); continue; }
    if (!isBS(a.type)) continue;
    byCat[categorize(a)].push({ code: String(a.code), name: a.name || String(a.code), type: a.type, bank_acct: a.bank_acct });
  }
  for (const k of Object.keys(byCat)) byCat[k].sort((p, q) => String(p.code).localeCompare(String(q.code), undefined, { numeric: true }));
  const P = { db, eid, ent, m, L, profile, display: profile.displayName, byCat, flags, computeBalances: ctx.computeBalances, apAging: ctx.apAging };
  const outputs = [], skipped = [];
  for (const b of BOOKS) {
    if (!byCat[b.cat].length && b.cat !== 'equity') { skipped.push(b.title); continue; }
    const { wb, summary } = await b.build(P);
    const buf = Buffer.from(await wb.xlsx.writeBuffer());
    outputs.push({ file: b.n + '_' + b.slug + '_CL_' + m.label + '.xlsx', title: b.title, accounts: byCat[b.cat].length, buf,
      balance: summary.balance, gl: summary.gl, variance: r2((summary.balance || 0) - (summary.gl || 0)) });
  }
  // Trial balance check: every BS account + cumulative P&L must net to zero.
  let tb = 0;
  for (const code of L.activeCodes(m.end)) tb = r2(tb + L.dr(code, m.end));
  if (nz(tb)) flags.push({ severity: 'exception', wp: 'Coverage', message: 'The trial balance does not balance at ' + mdy(m.end) + ' (debits exceed credits by ' + fmt(tb) + ').' });
  const covered = Object.values(byCat).reduce((s, x) => s + x.length, 0);
  return { m, ent, display: profile.displayName, outputs, skipped, flags, covered };
}

// ── Persistence + route ───────────────────────────────────────────────────────
const folderFor = (m) => 'Workpapers/Month-End Leadsheets/' + m.y + '/' + pad(m.mo) + ' ' + m.monthName;

function saveToWorkpapers(ctx, eid, m, file, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(m);
  const parts = folder.split('/');
  const ins = db.prepare("INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, file);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid)); fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + file.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare("INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, file, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return prior.length;
}

function registerLeadsheetRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/leadsheets/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const pkg = await buildPackage(ctx, eid, (req.body && req.body.month_end) || '');
        const zip = new JSZip();
        let replaced = 0;
        for (const o of pkg.outputs) { zip.file(o.file, o.buf); replaced += saveToWorkpapers(ctx, eid, pkg.m, o.file, o.buf, who); }
        const zbuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
        const zname = String(pkg.display).replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '') + '_Leadsheets_CL_' + pkg.m.label + '.zip';
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename="' + zname + '"');
        res.setHeader('X-Leadsheets-Summary', JSON.stringify({
          month: pkg.m.label, month_name: pkg.m.monthName + ' ' + pkg.m.y, saved_to: folderFor(pkg.m), replaced,
          accounts: pkg.covered, skipped: pkg.skipped,
          books: pkg.outputs.map((o) => ({ file: o.file, title: o.title, accounts: o.accounts, balance: o.balance, gl: o.gl, variance: o.variance })),
          flags: pkg.flags,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(zbuf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerLeadsheetRoutes, buildPackage, resolveMonth, categorize };
