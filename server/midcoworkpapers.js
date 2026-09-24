// ─── CLRFI Midco I workpapers (CLA lead-sheet format) ─────────────────────────
//
// One monthly workbook that reproduces CLA's four FloQast balance-sheet
// workpapers for CLRFI Midco I (CL entity 70), tab-for-tab and in CLA's layout,
// built from the CloudLedger ledger instead of pasted screenshots:
//
//   01  Cash Leadsheet         + one tab per bank account (statement vs register,
//                                with the CL bank reconciliation behind it)
//   06  Other Assets Leadsheet + one tab per account (13100 Interest Reserve):
//                                prior balance + current-period GL activity
//   12  Loan Leadsheet         + one tab per loan (25063 BOT): prior balance +
//                                funds received, CP/LTP breakout
//   15  Equity Rollforward     + YTD NI: CLA's monthly equity roll-forward
//
// Conventions (CLA's): every balance is shown debit-positive — credits in
// parentheses — so the loan and equity read (x). Blue cells are entry cells.
// "Prior period" is the prior quarter end (Midco reports quarterly, so the
// current period is the quarter to date), matching CLA's June package.
//
// Every subtotal, every lead-sheet figure and every tie is a live formula
// pointing at its supporting tab; the only literals are ledger amounts (opening
// balances, GL lines, monthly net income) and the bank statement balance from
// the CL reconciliation. CLA's FloQast-only columns (FQ Anchor) carry a CL tie
// instead (reconciliation difference / variance to GL).
//
// Built like the other CL workpapers (buildData -> buildWorkbook ->
// saveToWorkpapers) and registered by index.js. ctx = { db, auth,
// requireEntityAccess, requireRole, workpapersDir, computeBalances }.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const DISPLAY_NAME = 'CLRFI Midco I';
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const pad2 = (n) => String(n).padStart(2, '0');
const eom = (y, m) => new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10); // m is 1-based
const nextDay = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10); };
const xlDate = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const shortDate = (s) => { const [y, m, d] = s.split('-').map(Number); return m + '/' + d + '/' + String(y).slice(2); };

function isMidcoEntity(ent) {
  if (!ent) return false;
  return Number(ent.id) === 70 || /cl[rf]{2}i\s*midco\s*i\b/i.test(String(ent.name || ''));
}

// Any date in the month -> that month end; the roll-forward starts at the prior
// quarter end.
function resolvePeriod(input) {
  const m = String(input || '').match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) throw new Error('month_end must be a date in YYYY-MM-DD form');
  const y = Number(m[1]), mo = Number(m[2]);
  if (mo < 1 || mo > 12) throw new Error('month_end must be a valid date');
  const qStart = Math.floor((mo - 1) / 3) * 3 + 1;
  const rollBeg = eom(y, qStart - 1); // qStart-1 = 0 -> Dec 31 of the prior year
  return {
    y, mo, end: eom(y, mo), rollBeg, rollFrom: nextDay(rollBeg),
    ys: y + '-01-01', ye: y + '-12-31', pye: (y - 1) + '-12-31',
    year: String(y), label: y + '-' + pad2(mo), monthFolder: pad2(mo) + ' - ' + MONTHS[mo - 1],
  };
}

// ── Ledger reads ─────────────────────────────────────────────────────────────
// Debit-positive balance per account as of a date (CL's computeBalances, so the
// figures are the ones the CL Balance Sheet shows).
function drBalances(computeBalances, eid, asOf) {
  const m = new Map();
  for (const r of computeBalances(eid, { as_of: asOf }) || []) {
    const isDr = r.type === 'Asset' || r.type === 'Expense';
    m.set(String(r.code), r2(isDr ? r.balance : -r.balance));
  }
  return m;
}

function glLines(db, eid, from, to, codes) {
  if (!codes.length) return [];
  return db.prepare(`
    SELECT je.id AS entry_id, je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number,
           je.vendor AS vendor, je.memo AS memo, jl.id AS line_id, jl.account_code AS code,
           jl.debit AS debit, jl.credit AS credit, jl.description AS description
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ?
      AND jl.account_code IN (${codes.map(() => '?').join(',')})
    ORDER BY je.date, je.entry_num, jl.id
  `).all(eid, from, to, ...codes).map((r) => ({
    date: r.date, num: String(r.entry_num || r.doc_number || ''), name: r.vendor || '',
    memo: String(r.description || r.memo || '').trim(), code: String(r.code),
    amount: r2((r.debit || 0) - (r.credit || 0)),
  }));
}

// Monthly P&L (debit-positive: income negative) by month and location.
function monthlyPnl(db, eid, from, to) {
  return db.prepare(`
    SELECT substr(je.date, 1, 7) AS ym, dl.name AS loc, SUM(jl.debit) AS td, SUM(jl.credit) AS tc
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? AND a.type IN ('Revenue', 'Expense')
    GROUP BY ym, loc
  `).all(eid, from, to).map((r) => ({ ym: r.ym, loc: r.loc || '', dr: r2((r.td || 0) - (r.tc || 0)) }));
}

// Bank reconciliation for an account as of the month end: the CL reconciliation
// dated exactly the month end, plus the items still uncleared at that date
// (deposits in transit / outstanding checks), mirroring CL's reconciliation report.
function bankRec(db, eid, code, end) {
  const rec = db.prepare('SELECT * FROM reconciliations WHERE entity_id = ? AND account_code = ? AND statement_date <= ? ORDER BY statement_date DESC, id DESC LIMIT 1').get(eid, code, end);
  if (!rec || rec.statement_date !== end) return { rec: null, latest: rec ? rec.statement_date : null, deposits: [], checks: [] };
  const lines = db.prepare(`
    SELECT je.id AS entry_id, je.entry_num, je.date, je.memo, je.vendor, jl.debit, jl.credit,
           (SELECT COUNT(*) FROM journal_lines x WHERE x.entry_id = je.id AND x.id <= jl.id) - 1 AS line_index
    FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id
    WHERE je.entity_id = ? AND jl.account_code = ? AND je.date <= ?
    ORDER BY je.date, je.id, jl.id
  `).all(eid, code, end);
  const cleared = new Set(db.prepare(`
    SELECT ci.entry_id, ci.line_index FROM cleared_items ci
    LEFT JOIN reconciliations r ON r.id = ci.reconciliation_id
    WHERE ci.entity_id = ? AND ci.account_code = ? AND (ci.reconciliation_id IS NULL OR r.statement_date <= ?)
  `).all(eid, code, end).map((c) => c.entry_id + '-' + c.line_index));
  const open = lines.filter((l) => !cleared.has(l.entry_id + '-' + l.line_index)).map((l) => ({
    date: l.date, num: String(l.entry_num || ''), memo: String(l.vendor || l.memo || '').trim(),
    amount: r2((l.debit || 0) - (l.credit || 0)),
  })).filter((l) => Math.abs(l.amount) >= 0.005);
  return { rec, latest: rec.statement_date, deposits: open.filter((l) => l.amount > 0), checks: open.filter((l) => l.amount < 0) };
}

// ── Data ─────────────────────────────────────────────────────────────────────
const CASH = (a) => a.type === 'Asset' && (Number(a.bank_acct) === 1 || /^10\d{3}$/.test(a.code));
const OTHER = (a) => a.type === 'Asset' && (/^13\d{3}$/.test(a.code) || /interest reserve/i.test(a.name));
const LOAN = (a) => a.type === 'Liability' && (/^2[25]\d{3}$/.test(a.code) || /\bloans?\b|notes? payable/i.test(a.name));

function buildData(ctx, per, eid) {
  const { db, computeBalances } = ctx;
  const ent = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(eid);
  const accts = db.prepare('SELECT code, name, type, bank_acct FROM accounts WHERE entity_id = ?').all(eid)
    .map((a) => ({ code: String(a.code), name: a.name || '', type: a.type, bank_acct: a.bank_acct }));
  const bEnd = drBalances(computeBalances, eid, per.end);
  const bBeg = drBalances(computeBalances, eid, per.rollBeg);
  const bPye = drBalances(computeBalances, eid, per.pye);
  const bal = (m, c) => m.get(c) || 0;
  const flags = [];

  // Activity in the year (for cash) / the roll period (others), to keep accounts
  // that moved but ended at zero.
  const ytdLines = glLines(db, eid, per.ys, per.end, accts.map((a) => a.code));
  const movedYtd = new Set(ytdLines.map((l) => l.code));
  const movedRoll = new Set(ytdLines.filter((l) => l.date >= per.rollFrom).map((l) => l.code));
  const pick = (test, moved) => accts.filter(test)
    .filter((a) => Math.abs(bal(bEnd, a.code)) >= 0.005 || Math.abs(bal(bBeg, a.code)) >= 0.005 || moved.has(a.code))
    .sort((x, y) => x.code.localeCompare(y.code));

  // 01 Cash
  const cash = pick(CASH, movedYtd).map((a) => {
    const br = bankRec(db, eid, a.code, per.end);
    const book = bal(bEnd, a.code);
    const stmt = br.rec ? r2(br.rec.statement_balance) : null;
    const dit = r2(br.deposits.reduce((s, l) => s + l.amount, 0));
    const oc = r2(br.checks.reduce((s, l) => s + l.amount, 0));
    const diff = stmt == null ? null : r2(stmt + dit + oc - book);
    if (!br.rec) {
      flags.push({ severity: 'review', wp: 'Cash', message: a.code + ' ' + a.name + ': no CloudLedger bank reconciliation dated ' + shortDate(per.end)
        + (br.latest ? ' (latest is ' + shortDate(br.latest) + ')' : '') + ' — key the bank statement balance on tab ' + a.code + '.' });
    } else if (Math.abs(diff) >= 0.005) {
      flags.push({ severity: 'exception', wp: 'Cash', message: a.code + ' ' + a.name + ': adjusted bank balance differs from the register by ' + diff.toFixed(2) + '.' });
    }
    return Object.assign({}, a, { book, stmt, diff, rec: br.rec, deposits: br.deposits, checks: br.checks });
  });

  // 06 Other assets, 12 Loans: prior-quarter-end balance + the period's GL lines.
  const rollDetail = (a) => {
    const lines = ytdLines.filter((l) => l.code === a.code && l.date >= per.rollFrom);
    const beg = bal(bBeg, a.code), end = bal(bEnd, a.code);
    const act = r2(lines.reduce((s, l) => s + l.amount, 0));
    if (Math.abs(r2(beg + act - end)) >= 0.005) {
      flags.push({ severity: 'exception', wp: a.code, message: a.code + ' ' + a.name + ': prior balance + period activity (' + r2(beg + act).toFixed(2) + ') does not equal the GL balance (' + end.toFixed(2) + ').' });
    }
    return Object.assign({}, a, { beg, end, lines });
  };
  const other = pick(OTHER, movedRoll).map(rollDetail);
  const loans = pick(LOAN, movedRoll).map(rollDetail);

  // 15 Equity roll-forward.
  const eqAccts = accts.filter((a) => a.type === 'Equity');
  const eqActivity = new Map(); // code -> [12 months]
  for (const l of ytdLines) {
    const a = eqAccts.find((x) => x.code === l.code); if (!a) continue;
    const mi = Number(l.date.slice(5, 7)) - 1;
    if (!eqActivity.has(a.code)) eqActivity.set(a.code, Array(12).fill(0));
    eqActivity.get(a.code)[mi] = r2(eqActivity.get(a.code)[mi] + l.amount);
  }
  const kind = (a) => (/distribut|draw|withdraw/i.test(a.name) ? 'draws' : (/contribut|capital/i.test(a.name) ? 'contrib' : 'other'));
  const equity = {
    beginning: eqAccts.filter((a) => Math.abs(bal(bPye, a.code)) >= 0.005 || /retained earnings/i.test(a.name) || a.code === '39000')
      .sort((x, y) => x.code.localeCompare(y.code)).map((a) => Object.assign({}, a, { pye: bal(bPye, a.code) })),
    contrib: [], draws: [], other: [],
  };
  for (const a of eqAccts.slice().sort((x, y) => x.code.localeCompare(y.code))) {
    const months = eqActivity.get(a.code);
    if (months && months.some((v) => Math.abs(v) >= 0.005)) equity[kind(a)].push(Object.assign({}, a, { months }));
  }
  // Unclosed P&L before the year (debit-positive) = CLA's "Prior Period Net (Income)/Loss".
  const priorNi = r2(monthlyPnl(db, eid, '0000-01-01', per.pye).reduce((s, r) => s + r.dr, 0));
  // The prior period's income closes into Retained Earnings — keep a row for it
  // even if the ledger has no RE account yet.
  if (Math.abs(priorNi) >= 0.005 && !equity.beginning.some((a) => a.code === '39000' || /retained earnings/i.test(a.name))) {
    equity.beginning.push({ code: '39000', name: 'Retained Earnings', type: 'Equity', pye: 0 });
  }
  const pnl = monthlyPnl(db, eid, per.ys, per.end);
  const niMonths = Array(12).fill(0);
  const locMap = new Map();
  for (const r of pnl) {
    const mi = Number(r.ym.slice(5, 7)) - 1;
    niMonths[mi] = r2(niMonths[mi] + r.dr);
    const loc = r.loc || DISPLAY_NAME;
    if (!locMap.has(loc)) locMap.set(loc, Array(12).fill(0));
    locMap.get(loc)[mi] = r2(locMap.get(loc)[mi] - r.dr); // YTD NI tab shows income positive
  }
  const niByLoc = [...locMap.entries()].sort((x, y) => x[0].localeCompare(y[0])).map(([loc, months]) => ({ loc, months }));
  if (!niByLoc.length) niByLoc.push({ loc: DISPLAY_NAME, months: Array(12).fill(0) });
  // GL total equity incl. current-year earnings, debit-positive.
  let eqGl = 0;
  for (const a of eqAccts) eqGl += bal(bEnd, a.code);
  for (const a of accts.filter((x) => x.type === 'Revenue' || x.type === 'Expense')) eqGl += bal(bEnd, a.code);
  eqGl = r2(eqGl);
  const eqRoll = r2(equity.beginning.reduce((s, a) => s + a.pye, 0) + priorNi
    + ['contrib', 'draws', 'other'].reduce((s, k) => s + equity[k].reduce((t, a) => t + a.months.reduce((u, v) => u + v, 0), 0), 0)
    + niMonths.reduce((s, v) => s + v, 0));
  if (Math.abs(r2(eqRoll - eqGl)) >= 0.005) flags.push({ severity: 'exception', wp: 'Equity Rollforward', message: 'Equity roll-forward ending (' + eqRoll.toFixed(2) + ') does not equal GL total equity (' + eqGl.toFixed(2) + ').' });

  // Balance-sheet accounts these four workpapers do not cover (CLA supports them
  // in its other numbered workpapers) — listed on the Index for completeness.
  const covered = new Set([...cash, ...other, ...loans].map((a) => a.code));
  const uncovered = accts.filter((a) => (a.type === 'Asset' || a.type === 'Liability') && !covered.has(a.code) && Math.abs(bal(bEnd, a.code)) >= 0.005)
    .sort((x, y) => x.code.localeCompare(y.code)).map((a) => Object.assign({}, a, { end: bal(bEnd, a.code) }));

  const sumEnd = (list) => r2(list.reduce((s, a) => s + bal(bEnd, a.code), 0));
  return {
    per, entity_name: ent ? ent.name : ('entity ' + eid), flags,
    cash, other, loans, equity, priorNi, niMonths, niByLoc, uncovered,
    ties: { cash: sumEnd(cash), other_assets: sumEnd(other), loans: sumEnd(loans), equity: eqGl },
  };
}

// ── Workbook styling (CLA's template) ────────────────────────────────────────
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const DATE = 'm/d/yyyy';
const C = {
  entry: 'FFDCE6F1',     // light blue entry cell
  link: 'FF244062',      // navy hyperlink text
  red: 'FFFF0000',
  gray: 'FFD9D9D9',
  green: 'FFE2EFDA',     // "Current Year" band
  greenL: 'FFEBF1DE', greenM: 'FFD7E4BC',
  purpleL: 'FFE5E0EC', purpleM: 'FFCCC0DA',
  orangeL: 'FFFDE9D9',
  pink: 'FFE6B9B8',      // CP / LTP breakout band
  rec: 'FFCEDEEF',       // reconciliation-report band
  yellow: 'FFFFFF00',
};
const font = (o = {}) => Object.assign({ name: 'Calibri', size: 11 }, o);
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const THIN = { style: 'thin' }, MED = { style: 'medium' };
const f = (formula, result) => ({ formula, result });
const hyper = (sheet, ref, text) => f('HYPERLINK("#\'' + sheet + '\'!' + ref + '","' + String(text).replace(/"/g, '""') + '")', String(text));
const col = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

function put(ws, ref, value, o = {}) {
  const c = ws.getCell(ref);
  if (value !== undefined) c.value = value;
  c.font = font(o.font || {});
  if (o.fmt) c.numFmt = o.fmt;
  if (o.fill) c.fill = fill(o.fill);
  if (o.align) c.alignment = Object.assign({}, c.alignment || {}, { horizontal: o.align });
  if (o.wrap) c.alignment = Object.assign({}, c.alignment || {}, { wrapText: true, vertical: 'bottom' });
  if (o.border) c.border = o.border;
  return c;
}
function edge(ws, r1, c1, r2_, c2, st) {
  for (let r = r1; r <= r2_; r++) for (let c = c1; c <= c2; c++) {
    const cell = ws.getCell(r, c); const b = Object.assign({}, cell.border || {});
    if (r === r1) b.top = st; if (r === r2_) b.bottom = st; if (c === c1) b.left = st; if (c === c2) b.right = st;
    cell.border = b;
  }
}
const rule = (ws, r, c1, c2, side, st) => { for (let c = c1; c <= c2; c++) { const cell = ws.getCell(r, c); cell.border = Object.assign({}, cell.border || {}, { [side]: st }); } };
const boxAll = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const widths = (ws, map) => { for (const [k, w] of Object.entries(map)) ws.getColumn(k).width = w; };

// Title block every CLA lead sheet opens with.
function leadHeader(ws, title, per) {
  put(ws, 'C1', DISPLAY_NAME, { font: { bold: true, size: 14 } }); ws.getRow(1).height = 18.75;
  put(ws, 'C2', title, { font: { bold: true, size: 12 } }); ws.getRow(2).height = 15.75;
  put(ws, 'C3', 'Month Ended:');
  put(ws, 'D3', xlDate(per.end), { font: { bold: true }, fmt: DATE, fill: C.entry, align: 'center', border: boxAll });
  put(ws, 'C4', 'Year Ended:');
  put(ws, 'D4', xlDate(per.ye), { font: { bold: true }, fmt: DATE, fill: C.entry, align: 'center', border: boxAll });
}
function headerRow(ws, r, c1, labels, aligns) {
  labels.forEach((t, i) => put(ws, col(c1 + i) + r, t, { font: { bold: true }, align: (aligns && aligns[i]) || undefined, wrap: /\n/.test(t) }));
  rule(ws, r, c1, c1 + labels.length - 1, 'bottom', THIN);
}

// ── 01 Cash ──────────────────────────────────────────────────────────────────
function buildCash(wb, d) {
  const ws = wb.addWorksheet('Cash Leadsheet', { views: [{ showGridLines: false }] });
  widths(ws, { A: 3.4, B: 11.7, C: 37.6, D: 13, E: 17.3, F: 17.3, G: 14, H: 17.9, I: 24.6 });
  leadHeader(ws, 'CASH LEAD SHEET', d.per);
  headerRow(ws, 7, 2, ['Account No.', 'Account Name', 'Worksheet', 'Cleared Balance', 'Register Balance', 'Rec Difference', 'Bank Rec', 'Comments'],
    [null, null, 'center', 'center', 'center', 'left', 'center', 'center']);
  const tabs = [];
  let r = 8;
  for (const a of d.cash) {
    put(ws, 'B' + r, Number(a.code), { fill: C.entry, align: 'center', border: { top: THIN, bottom: THIN, right: THIN } });
    put(ws, 'C' + r, a.name, { fill: C.entry, align: 'left', fmt: '@', border: boxAll });
    put(ws, 'D' + r, hyper(a.code, 'A1', a.code), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } }, align: 'center' });
    put(ws, 'E' + r, f("'" + a.code + "'!B4", a.stmt == null ? 0 : a.stmt), { fmt: ACCT, fill: C.entry, align: 'right' });
    put(ws, 'F' + r, f("'" + a.code + "'!B5", a.book), { fmt: ACCT, fill: C.entry, align: 'right' });
    tabs.push({ a, row: r });
    r++;
  }
  // CLA leaves one empty entry row at the foot of the table.
  put(ws, 'B' + r, null, { fill: C.entry, border: { top: THIN, bottom: THIN, right: THIN } });
  put(ws, 'C' + r, null, { fill: C.entry, border: boxAll });
  put(ws, 'E' + r, null, { fill: C.entry, fmt: ACCT }); put(ws, 'F' + r, null, { fill: C.entry, fmt: ACCT });
  const tot = r + 1;
  put(ws, 'B' + tot, 'Total Cash', { font: { bold: true } });
  put(ws, 'E' + tot, f('SUM(E8:E' + r + ')', r2(d.cash.reduce((s, a) => s + (a.stmt || 0), 0))), { font: { bold: true }, fmt: ACCT });
  put(ws, 'F' + tot, f('SUM(F8:F' + r + ')', d.ties.cash), { font: { bold: true }, fmt: ACCT });
  put(ws, 'G' + tot, f('SUM(G8:G' + r + ')', r2(d.cash.reduce((s, a) => s + (a.diff || 0), 0))), { font: { bold: true, italic: true, color: { argb: C.red } }, fmt: ACCT });
  rule(ws, tot, 2, 9, 'top', THIN);
  edge(ws, 7, 2, tot, 9, MED);

  // One tab per bank account; then back-fill the lead-sheet cells that point at it.
  for (const { a, row } of tabs) {
    const ref = buildCashTab(wb, d, a, row);
    put(ws, 'G' + row, f("'" + a.code + "'!" + ref.diff, a.diff == null ? 0 : a.diff), { font: { bold: true, italic: true, color: { argb: C.red } }, fmt: ACCT });
    const label = a.rec ? ('Rec ' + shortDate(a.rec.statement_date)) : 'Not reconciled';
    put(ws, 'H' + row, hyper(a.code, ref.rec, label), { font: { bold: true, italic: true, underline: true, color: { argb: a.rec ? C.link : C.red } }, align: 'center' });
    put(ws, 'I' + row, null, { font: { color: { argb: C.red } }, align: 'left' });
  }
  return { ws, totalRow: tot };
}

function buildCashTab(wb, d, a, leadRow) {
  const ws = wb.addWorksheet(a.code, { views: [{ showGridLines: false }] });
  widths(ws, { A: 19.9, B: 17.6, C: 12, D: 44, E: 10, F: 10, G: 16 });
  put(ws, 'A1', f("'Cash Leadsheet'!C" + leadRow, a.name), { font: { bold: true, italic: true, size: 16, color: { argb: 'FF0070C0' } }, fmt: '@' });
  ws.getRow(1).height = 21;
  put(ws, 'D2', hyper('Cash Leadsheet', 'A1', 'Back to Leadsheet'), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } } });
  put(ws, 'A4', 'Statement Balance');
  put(ws, 'B4', a.stmt == null ? null : a.stmt, { fmt: ACCT, fill: C.entry });
  put(ws, 'A5', 'Register Balance');
  put(ws, 'B5', a.book, { fmt: ACCT, fill: C.entry });

  // Reconciliation report block (CLA pastes the bank rec here).
  const V = (o = {}) => Object.assign({ name: 'Verdana', size: 8 }, o);
  const vput = (ref, v, o = {}) => { const c = put(ws, ref, v, o); c.font = V(o.vfont || {}); return c; };
  vput('B8', 'Reconciliation Report', { vfont: { bold: true, size: 14 }, fill: C.rec, align: 'center' }); ws.mergeCells('B8:G8');
  vput('B9', 'As Of ' + shortDate(d.per.end).replace(/^(\d+)\/(\d+)\/(\d+)$/, (m0, mm, dd, yy) => pad2(mm) + '/' + pad2(dd) + '/20' + yy), { vfont: { bold: true }, fill: C.rec, align: 'center' }); ws.mergeCells('B9:G9');
  vput('B10', 'Account: ' + a.name, { vfont: { bold: true }, fill: C.rec, align: 'center' }); ws.mergeCells('B10:G10');
  const lines = [
    [12, 'Statement Ending Balance'], [13, 'Deposits in Transit'], [14, 'Outstanding Checks and Charges'], [15, 'Adjusted Bank Balance'],
    [17, 'Book Balance'], [18, 'Adjustments*'], [19, 'Adjusted Book Balance'], [21, 'Difference'],
  ];
  for (const [rr, t] of lines) vput('B' + rr, t, { vfont: { bold: true } });

  // Detail lists first so the summary can SUM them.
  let r = 24;
  const list = (title, items) => {
    vput('B' + r, title, { vfont: { bold: true, size: 12 } }); r++;
    ['Date', 'Num', 'Name / Memo', '', '', 'Amount'].forEach((h, i) => { if (h || i === 3 || i === 4) vput(col(2 + i) + r, h, { vfont: { bold: true }, fill: C.rec, align: i === 5 ? 'right' : 'left' }); });
    r++;
    const first = r;
    for (const it of items) {
      vput('B' + r, xlDate(it.date), { fmt: 'mm/dd/yyyy', align: 'left' });
      vput('C' + r, it.num, { align: 'left' });
      vput('D' + r, it.memo);
      vput('G' + r, it.amount, { fmt: '#,##0.00;(#,##0.00)', align: 'right' });
      r++;
    }
    if (!items.length) { vput('D' + r, 'None', { vfont: { italic: true } }); r++; }
    vput('B' + r, 'Total ' + title, { vfont: { bold: true } });
    const tot = r;
    vput('G' + tot, items.length ? f('SUM(G' + first + ':G' + (r - 1) + ')', r2(items.reduce((s, x) => s + x.amount, 0))) : 0, { vfont: { bold: true }, fmt: '#,##0.00;(#,##0.00)', border: { top: THIN } });
    r += 2;
    return tot;
  };
  const ditTot = list('Deposits in Transit', a.deposits);
  const ocTot = list('Outstanding Checks and Charges', a.checks);

  const NUM = '#,##0.00;(#,##0.00)';
  const dit = r2(a.deposits.reduce((s, x) => s + x.amount, 0)), oc = r2(a.checks.reduce((s, x) => s + x.amount, 0));
  const adjBank = r2((a.stmt || 0) + dit + oc);
  vput('G12', f('B4', a.stmt || 0), { vfont: { bold: true }, fmt: NUM, align: 'right' });
  vput('G13', f('G' + ditTot, dit), { vfont: { bold: true }, fmt: NUM, align: 'right' });
  vput('G14', f('G' + ocTot, oc), { vfont: { bold: true }, fmt: NUM, align: 'right' });
  vput('G15', f('SUM(G12:G14)', adjBank), { vfont: { bold: true }, fmt: NUM, align: 'right', border: { top: THIN } });
  vput('G17', f('B5', a.book), { vfont: { bold: true }, fmt: NUM, align: 'right' });
  vput('G18', 0, { vfont: { bold: true }, fmt: NUM, align: 'right' });
  vput('G19', f('SUM(G17:G18)', a.book), { vfont: { bold: true }, fmt: NUM, align: 'right', border: { top: THIN } });
  vput('G21', f('G15-G19', r2(adjBank - a.book)), { vfont: { bold: true, color: { argb: C.red } }, fmt: NUM, align: 'right', border: { top: THIN, bottom: { style: 'double' } } });
  vput('B22', a.rec
    ? ('Reconciled in CloudLedger by ' + (a.rec.completed_by || '—') + (a.rec.completed_at ? ' on ' + String(a.rec.completed_at).slice(0, 10) : '') + '. Items shown are the register lines still uncleared at ' + shortDate(d.per.end) + '.')
    : ('No CloudLedger bank reconciliation is dated ' + shortDate(d.per.end) + ' — key the bank statement balance in B4.'),
  { vfont: { italic: true, color: { argb: a.rec ? 'FF595959' : C.red } } });
  return { diff: 'G21', rec: 'B8' };
}

// ── 06 Other Assets ──────────────────────────────────────────────────────────
function buildOther(wb, d) {
  const ws = wb.addWorksheet('Other Assets Leadsheet', { views: [{ showGridLines: false }] });
  widths(ws, { A: 3.4, B: 14.4, C: 38.1, D: 19.7, E: 16.1, F: 13, G: 22.1 });
  leadHeader(ws, 'OTHER ASSETS LEAD SHEET', d.per);
  ws.getRow(7).height = 28.8;
  headerRow(ws, 7, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance\nDr (Cr)', 'GL Variance', 'Comments'], [null, null, 'center', 'center', 'left', 'center']);
  let r = 8;
  for (const a of d.other) {
    const ref = buildRollTab(wb, d, a, 'Other Assets Leadsheet');
    put(ws, 'B' + r, Number(a.code), { fill: C.entry, align: 'center', border: { top: THIN, bottom: THIN, right: THIN } });
    put(ws, 'C' + r, a.name, { fill: C.entry, align: 'left', border: { top: THIN, bottom: THIN, left: THIN } });
    put(ws, 'D' + r, hyper(a.code, 'A1', a.code), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } }, align: 'center' });
    put(ws, 'E' + r, f("'" + a.code + "'!" + ref.total, a.end), { fmt: ACCT, fill: C.entry, border: boxAll });
    put(ws, 'F' + r, f("'" + a.code + "'!" + ref.variance, 0), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT, align: 'left' });
    r++;
  }
  put(ws, 'B' + r, null, { fill: C.entry, border: { top: THIN, bottom: THIN, right: THIN } });
  put(ws, 'C' + r, null, { fill: C.entry, border: { top: THIN, bottom: THIN, left: THIN } });
  put(ws, 'E' + r, null, { fill: C.entry, fmt: ACCT, border: boxAll });
  const tot = r + 1;
  put(ws, 'B' + tot, 'Total Other Assets', { font: { bold: true } });
  put(ws, 'E' + tot, f('SUM(E8:E' + r + ')', d.ties.other_assets), { font: { bold: true }, fmt: ACCT });
  put(ws, 'F' + tot, f('SUM(F8:F' + r + ')', 0), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT });
  rule(ws, tot, 2, 7, 'top', THIN);
  edge(ws, 7, 2, tot, 7, MED);
  put(ws, 'B' + (tot + 2), '* Direct and indirect costs incurred during the period related to the real estate development have been capitalized. All bills are stored in the Bill.com account.', { font: { bold: true } });
  return { ws, totalRow: tot };
}

// Prior-quarter-end balance + the period's GL lines (other assets).
function buildRollTab(wb, d, a, leadName) {
  const ws = wb.addWorksheet(a.code, { views: [{ showGridLines: false }] });
  widths(ws, { A: 30, B: 16, C: 16, D: 30, E: 70 });
  const reserve = /interest reserve/i.test(a.name);
  put(ws, 'A1', a.code + ' - ' + a.name, { font: { bold: true, italic: true, size: 16, color: { argb: 'FF366092' } } });
  ws.getRow(1).height = 21;
  put(ws, 'D2', hyper(leadName, 'A1', 'Back to Leadsheet'), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } } });
  const act = r2(a.lines.reduce((s, l) => s + l.amount, 0));
  put(ws, 'A3', reserve ? 'Prior Invoices through' : 'Prior Period Balance at', { font: { bold: true } });
  put(ws, 'B3', xlDate(d.per.rollBeg), { font: { bold: true }, fmt: DATE, align: 'center' });
  put(ws, 'C3', a.beg, { fmt: ACCT });
  put(ws, 'A4', 'Current Period');
  put(ws, 'C4', f('C9', act), { fmt: ACCT, fill: C.purpleL, border: { bottom: THIN } });
  put(ws, 'A5', 'Total Account Bal at', { font: { bold: true } });
  put(ws, 'B5', f("'" + leadName + "'!D3", xlDate(d.per.end)), { font: { bold: true }, fmt: DATE, align: 'center' });
  put(ws, 'C5', f('SUM(C3:C4)', r2(a.beg + act)), { font: { bold: true }, fmt: ACCT, fill: C.entry });
  put(ws, 'A6', 'Per General Ledger', { font: { italic: true } });
  put(ws, 'C6', a.end, { fmt: ACCT });
  put(ws, 'A7', 'Variance', { font: { italic: true } });
  put(ws, 'C7', f('C5-C6', r2(a.beg + act - a.end)), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT });
  put(ws, 'A9', reserve ? 'Midco I Interest Paid from Interest Reserve' : 'Current Period Activity', { font: { bold: true, italic: true } });
  headerRow(ws, 11, 1, ['Date', 'Num', 'Amount', 'Name', 'Memo / Description'], ['left', 'left', 'right', 'left', 'left']);
  let r = 12;
  for (const l of a.lines) {
    put(ws, 'A' + r, xlDate(l.date), { fmt: DATE, align: 'left' });
    put(ws, 'B' + r, l.num, { align: 'left' });
    put(ws, 'C' + r, l.amount, { fmt: ACCT, fill: C.purpleL });
    put(ws, 'D' + r, l.name);
    put(ws, 'E' + r, l.memo);
    r++;
  }
  if (!a.lines.length) { put(ws, 'D' + r, 'No activity in the period', { font: { italic: true } }); r++; }
  put(ws, 'A' + r, 'Total', { font: { bold: true } });
  put(ws, 'C' + r, a.lines.length ? f('SUM(C12:C' + (r - 1) + ')', act) : 0, { font: { bold: true }, fmt: ACCT, border: { top: THIN } });
  put(ws, 'C9', f('C' + r, act), { fmt: ACCT, fill: C.purpleL });
  return { total: 'C5', variance: 'C7' };
}

// ── 12 Debt ──────────────────────────────────────────────────────────────────
function buildLoans(wb, d) {
  const ws = wb.addWorksheet('Loan Leadsheet', { views: [{ showGridLines: false, zoomScale: 85 }] });
  widths(ws, { A: 3, B: 12.6, C: 31.5, D: 13.5, E: 17, F: 14.1, G: 16, H: 14, I: 17, J: 14, K: 60 });
  leadHeader(ws, 'DEBT LEADSHEET', d.per);
  put(ws, 'F3', null, { fill: C.entry, border: boxAll, align: 'left' });
  put(ws, 'G3', '= Entry Cell');
  put(ws, 'G6', 'Current Portion (CP) / Long Term Portion (LTP) Breakouts', { font: { bold: true, size: 12 }, fill: C.pink, align: 'center' });
  ws.mergeCells('G6:J6');
  edge(ws, 6, 7, 6, 10, MED);
  headerRow(ws, 7, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance', 'GL Variance', 'CP LTD Balance', 'CP Support Ref', 'LTP LTD Balance', 'LTP Support Ref', 'Comments'],
    [null, null, 'center', 'center', 'left', 'left', 'left', 'left', 'left', 'center']);
  let r = 8;
  for (const a of d.loans) {
    const ref = buildLoanTab(wb, d, a);
    put(ws, 'B' + r, Number(a.code), { fill: C.entry, align: 'center', border: { top: THIN, bottom: THIN, left: MED, right: THIN } });
    put(ws, 'C' + r, a.name, { fill: C.entry, align: 'left', border: boxAll });
    put(ws, 'D' + r, hyper(a.code, 'D2', a.code), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } }, align: 'center' });
    put(ws, 'E' + r, f("'" + a.code + "'!" + ref.total, a.end), { fmt: ACCT });
    put(ws, 'F' + r, f("'" + a.code + "'!" + ref.variance, 0), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT, align: 'left' });
    put(ws, 'G' + r, 0, { fmt: ACCT, fill: C.entry, align: 'left' });
    put(ws, 'H' + r, null, { font: { bold: true, color: { argb: C.red } }, fill: C.entry, border: { top: THIN, bottom: THIN }, align: 'left' });
    put(ws, 'I' + r, f('E' + r + '-G' + r, a.end), { fmt: ACCT, align: 'left' });
    put(ws, 'J' + r, null, { font: { bold: true, color: { argb: C.red } }, fill: C.entry, border: { top: THIN, bottom: THIN }, align: 'left' });
    put(ws, 'K' + r, null, { align: 'left' });
    r++;
  }
  put(ws, 'B' + r, null, { fill: C.entry, border: { bottom: THIN, left: MED, right: THIN } });
  put(ws, 'C' + r, null, { fill: C.entry, border: { bottom: THIN, left: THIN, right: THIN } });
  put(ws, 'H' + r, null, { fill: C.entry, border: { bottom: THIN } });
  put(ws, 'J' + r, null, { fill: C.entry, border: { bottom: THIN } });
  const tot = r + 1;
  put(ws, 'B' + tot, 'Total Loans', { font: { bold: true }, align: 'center' });
  put(ws, 'E' + tot, f('SUM(E8:E' + r + ')', d.ties.loans), { font: { bold: true }, fmt: ACCT });
  put(ws, 'F' + tot, f('SUM(F8:F' + r + ')', 0), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT });
  put(ws, 'G' + tot, f('SUM(G8:G' + r + ')', 0), { font: { bold: true }, fmt: ACCT });
  put(ws, 'I' + tot, f('SUM(I8:I' + r + ')', d.ties.loans), { font: { bold: true }, fmt: ACCT });
  rule(ws, tot, 2, 11, 'top', THIN);
  edge(ws, 7, 2, tot, 11, MED);
  edge(ws, 7, 7, tot, 10, MED); // CP / LTP breakout box
  return { ws, totalRow: tot };
}

function buildLoanTab(wb, d, a) {
  const ws = wb.addWorksheet(a.code, { views: [{ showGridLines: false }] });
  widths(ws, { A: 34.1, B: 17, C: 12, D: 18, E: 15.2, F: 4, G: 12, H: 12, I: 75, J: 17 });
  put(ws, 'A1', a.code + ' - ' + a.name, { font: { bold: true, italic: true, size: 14, color: { argb: 'FF17365D' } } });
  ws.getRow(1).height = 18.75;
  put(ws, 'D1', null, { fill: C.entry, border: boxAll });
  put(ws, 'E1', '= Entry Cell');
  put(ws, 'D2', hyper('Loan Leadsheet', 'A1', 'Back to Leadsheet'), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } } });
  const act = r2(a.lines.reduce((s, l) => s + l.amount, 0));
  put(ws, 'A3', 'Prior Period Loan Amount:');
  put(ws, 'B3', f('B11', a.beg), { fmt: ACCT, fill: C.greenL, border: boxAll, align: 'center' });
  put(ws, 'A4', 'Dev Funds Received in period');
  put(ws, 'B4', f('J11', act), { fmt: ACCT, fill: C.purpleL, border: boxAll, align: 'center' });
  put(ws, 'A5', 'Variance', { font: { italic: true }, align: 'center' });
  put(ws, 'B5', 0, { fmt: ACCT, fill: C.orangeL, border: boxAll, align: 'center' });
  put(ws, 'B6', f('SUM(B3:B5)', r2(a.beg + act)), { font: { bold: true }, fmt: ACCT, fill: C.entry, border: boxAll, align: 'center' });
  put(ws, 'A7', 'Per General Ledger', { font: { italic: true } });
  put(ws, 'B7', a.end, { fmt: ACCT, align: 'center' });
  put(ws, 'A8', 'Difference', { font: { italic: true } });
  put(ws, 'B8', f('B6-B7', r2(a.beg + act - a.end)), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT, align: 'center' });

  const AR = (o = {}) => Object.assign({ name: 'Arial', size: 11 }, o);
  put(ws, 'A11', 'Beginning Balance', { fill: C.greenM }).font = AR({ bold: true });
  put(ws, 'B11', a.beg, { fmt: ACCT, fill: C.greenL }).font = AR({ size: 10 });
  put(ws, 'C11', 'at ' + shortDate(d.per.rollBeg), { font: { italic: true, color: { argb: 'FF595959' } } });
  put(ws, 'G11', 'Funds Received by Entities Using the Loan Funds', { fill: C.purpleM }).font = AR({ bold: true });
  for (const c of ['H', 'I']) put(ws, c + '11', null, { fill: C.purpleM });
  headerRow(ws, 13, 7, ['Date', 'Num', 'Description', 'Amount'], ['left', 'left', 'left', 'right']);
  let r = 14;
  for (const l of a.lines) {
    put(ws, 'G' + r, xlDate(l.date), { fmt: DATE, align: 'left' });
    put(ws, 'H' + r, l.num, { align: 'left' });
    put(ws, 'I' + r, [l.name, l.memo].filter(Boolean).join(' — '));
    put(ws, 'J' + r, l.amount, { fmt: ACCT, fill: C.purpleL });
    r++;
  }
  if (!a.lines.length) { put(ws, 'I' + r, 'No activity in the period', { font: { italic: true } }); r++; }
  put(ws, 'J11', a.lines.length ? f('SUM(J14:J' + (r - 1) + ')', act) : 0, { fmt: ACCT, fill: C.purpleL });
  return { total: 'B6', variance: 'B8' };
}

// ── 15 Equity ────────────────────────────────────────────────────────────────
function buildEquity(wb, d) {
  const ws = wb.addWorksheet('Equity Rollforward', { views: [{ state: 'frozen', xSplit: 6, ySplit: 13, showGridLines: false, zoomScale: 80 }] });
  widths(ws, { A: 3.1, B: 52, C: 14.6, D: 17, E: 17, F: 17, S: 19.1 });
  for (let c = 7; c <= 18; c++) ws.getColumn(c).width = 15.4;
  leadHeader(ws, 'EQUITY ROLLFORWARD', d.per);
  const M = d.per.mo; // months with actuals
  const MC = (i) => col(7 + i); // month i (0-based) -> G..R
  rule(ws, 10, 1, 19, 'bottom', THIN);

  // Header band.
  const GH = { font: { bold: true }, fill: C.gray, align: 'center' };
  put(ws, 'C12', 'Ownership %', GH); put(ws, 'D12', 'Prior Year End', GH);
  put(ws, 'E12', 'Equity\nClosing Entry\n(should net to $0)', Object.assign({ wrap: true }, GH)); put(ws, 'F12', 'Opening Equity', GH);
  for (const c of ['C', 'D', 'E', 'F']) ws.getCell(c + '12').alignment = { horizontal: 'center', vertical: 'bottom', wrapText: true };
  ws.getRow(12).height = 45;
  put(ws, 'G12', 'Current Year', { font: { bold: true }, fill: C.green, align: 'center' });
  ws.mergeCells('G12:R12'); edge(ws, 12, 7, 12, 18, THIN);
  const DH = { font: { bold: true }, fmt: DATE, align: 'center', border: boxAll };
  const pyeD = xlDate(d.per.pye);
  put(ws, 'D13', f('EOMONTH($D$4,-12)', pyeD), Object.assign({ fill: C.gray }, DH));
  put(ws, 'E13', f('D13+1', xlDate(d.per.ys)), Object.assign({ fill: C.gray }, DH));
  put(ws, 'F13', f('D13+1', xlDate(d.per.ys)), Object.assign({ fill: C.gray }, DH));
  for (let i = 0; i < 12; i++) put(ws, MC(i) + '13', f('EOMONTH($D$4,' + (i - 11) + ')', xlDate(eom(d.per.y, i + 1))), DH);
  put(ws, 'S13', 'YTD', { font: { bold: true }, fill: C.gray, align: 'center', border: boxAll });

  let r = 15;
  const label = (t, o = {}) => { put(ws, 'B' + r, t, { font: Object.assign({ color: { argb: 'FF000000' } }, o) }); };
  const blank = () => { r++; };
  const sumRange = (c, a, b) => (b >= a ? 'SUM(' + c + a + ':' + c + b + ')' : '0');
  const totals = {}; // block -> row

  label('Sole Proprietor', { bold: true }); r += 2;
  label('Limited Liability Company (LLC) & S-Corp LLC', { bold: true }); r += 2;
  label('Partnership', { bold: true }); r++;

  // Beginning capital: each equity account at the prior year end; the prior
  // period's unclosed net income closes into Retained Earnings (column E).
  label("Beginning Partners' (Capital)/Deficit:", { italic: true }); r++;
  const begFirst = r;
  const reIdx = Math.max(0, d.equity.beginning.findIndex((a) => a.code === '39000' || /retained earnings/i.test(a.name)));
  const niRowPlaceholder = []; // cells that need the prior-NI row number
  d.equity.beginning.forEach((a, i) => {
    put(ws, 'B' + r, a.code + ' - ' + a.name, { fill: C.entry, align: 'left', border: { top: THIN, left: THIN, right: THIN } });
    put(ws, 'D' + r, a.pye, { fmt: ACCT, fill: C.entry, border: { top: THIN, left: THIN, right: THIN } });
    if (i === reIdx) niRowPlaceholder.push('E' + r);
    else put(ws, 'E' + r, 0, { fmt: ACCT, border: { top: THIN, left: THIN, right: THIN } });
    put(ws, 'F' + r, f('SUM(D' + r + ':E' + r + ')', i === reIdx ? r2(a.pye + d.priorNi) : a.pye), { fmt: ACCT, fill: C.gray, border: { top: THIN, left: THIN, right: THIN } });
    put(ws, 'S' + r, f('F' + r, i === reIdx ? r2(a.pye + d.priorNi) : a.pye), { fmt: ACCT, align: 'right' });
    r++;
  });
  if (!d.equity.beginning.length) { put(ws, 'B' + r, null, { fill: C.entry, border: boxAll }); r++; }
  const begLast = r - 1;
  const begTot = r;
  label("Beginning Partners' (Capital)/Deficit", { bold: true });
  const begSum = r2(d.equity.beginning.reduce((s, a) => s + a.pye, 0));
  for (const [c, v] of [['D', begSum], ['E', d.priorNi], ['F', r2(begSum + d.priorNi)], ['S', r2(begSum + d.priorNi)]]) {
    put(ws, c + r, f(sumRange(c, begFirst, begLast), v), { font: { bold: true }, fmt: ACCT, fill: c === 'F' ? C.gray : undefined, border: { top: THIN } });
  }
  totals.beg = r; r += 2;

  // Activity blocks: monthly GL activity per equity account.
  const block = (title, totalTitle, rows, key) => {
    label(title, { italic: true }); r++;
    const first = r;
    const items = rows.length ? rows : [null];
    for (const a of items) {
      put(ws, 'B' + r, a ? (a.code + ' - ' + a.name) : null, { fill: C.entry, align: 'left', border: { top: THIN, left: THIN, right: THIN } });
      put(ws, 'F' + r, null, { fill: C.gray, fmt: ACCT });
      for (let i = 0; i < 12; i++) {
        const v = a && i < M ? a.months[i] : null;
        put(ws, MC(i) + r, v != null && Math.abs(v) >= 0.005 ? v : null, { fmt: ACCT, fill: C.entry, border: { top: THIN, left: THIN, right: THIN } });
      }
      put(ws, 'S' + r, f('SUM(G' + r + ':R' + r + ')', a ? r2(a.months.slice(0, M).reduce((s, v) => s + v, 0)) : 0), { fmt: ACCT });
      r++;
    }
    const last = r - 1;
    label(totalTitle, { bold: true });
    for (let i = 0; i < 12; i++) {
      const v = r2(rows.reduce((s, a) => s + (i < M ? a.months[i] : 0), 0));
      put(ws, MC(i) + r, f(sumRange(MC(i), first, last), v), { font: { bold: true }, fmt: ACCT, border: { top: THIN } });
    }
    put(ws, 'F' + r, null, { fill: C.gray, border: { top: THIN } });
    put(ws, 'S' + r, f(sumRange('S', first, last), r2(rows.reduce((s, a) => s + a.months.slice(0, M).reduce((t, v) => t + v, 0), 0))), { font: { bold: true }, fmt: ACCT, border: { top: THIN } });
    totals[key] = r; r += 2;
  };
  block('(Contributions):', 'Total (Contributions)', d.equity.contrib, 'contrib');

  // Prior period net (income)/loss and its closing entry.
  label('Prior Period Net (Income)/Loss', { italic: true });
  const niRow = r;
  put(ws, 'D' + r, d.priorNi, { fmt: ACCT, fill: C.entry });
  put(ws, 'E' + r, f('-D' + r, -d.priorNi), { fmt: ACCT });
  put(ws, 'F' + r, f('SUM(D' + r + ':E' + r + ')', 0), { fmt: ACCT, fill: C.gray });
  r++;
  label('Total Net (Income)/Loss', { bold: true });
  for (const [c, v] of [['D', d.priorNi], ['E', -d.priorNi], ['F', 0]]) put(ws, c + r, f(c + niRow, v), { font: { bold: true }, fmt: ACCT, fill: c === 'F' ? C.gray : undefined, border: { top: THIN } });
  totals.ni = r; r += 2;
  for (const ref of niRowPlaceholder) put(ws, ref, f('D' + niRow, d.priorNi), { fmt: ACCT, border: { top: THIN, left: THIN, right: THIN } });

  block('Draws', 'Total Draws', d.equity.draws, 'draws');
  if (d.equity.other.length) block('Other Equity Activity', 'Total Other Equity Activity', d.equity.other, 'other');

  label('S-Corporation', { bold: true }); r += 2;
  label('C-Corporation', { bold: true }); r += 2;

  // Monthly roll: beginning -> activity -> net (income)/loss -> ending.
  const actKeys = ['contrib', 'draws', 'other'].filter((k) => totals[k]);
  const monthAct = (i) => r2(actKeys.reduce((s, k) => s + d.equity[k].reduce((t, a) => t + (i < M ? a.months[i] : 0), 0), 0));
  const opening = r2(begSum + d.priorNi);
  const begRow = r, niCurRow = r + 2, endRow = r + 4;
  label('Current Period Beginning (Equity)/Deficit', { bold: true });
  // D/E show the capital accounts before and after the closing entry (CLA's
  // layout); F adds the prior-period income row, which nets to zero.
  put(ws, 'D' + begRow, f('D' + totals.beg, begSum), { fmt: ACCT, border: { top: THIN, bottom: THIN } });
  put(ws, 'E' + begRow, f('E' + totals.beg, d.priorNi), { fmt: ACCT, border: { top: THIN, bottom: THIN } });
  put(ws, 'F' + begRow, f('F' + totals.beg + '+F' + totals.ni, opening), { fmt: ACCT, fill: C.gray, border: { top: THIN, bottom: THIN } });
  let run = opening;
  const endVals = [];
  for (let i = 0; i < 12; i++) {
    const begV = run;
    put(ws, MC(i) + begRow, f(i === 0 ? 'F' + begRow : MC(i - 1) + endRow, begV), { fmt: ACCT, border: { top: THIN, bottom: THIN } });
    run = r2(run + monthAct(i) + (i < M ? d.niMonths[i] : 0));
    endVals.push(run);
  }
  put(ws, 'S' + begRow, f('F' + begRow, opening), { fmt: ACCT, border: { top: THIN, bottom: THIN } });

  put(ws, 'B' + niCurRow, 'Current Period Net (Income)/Loss', { fill: C.entry, align: 'left', border: boxAll });
  for (let i = 0; i < 12; i++) put(ws, MC(i) + niCurRow, i < M ? d.niMonths[i] : null, { fmt: ACCT, fill: C.entry, border: boxAll });
  const ytdNi = r2(d.niMonths.slice(0, M).reduce((s, v) => s + v, 0));
  put(ws, 'S' + niCurRow, f('SUM(G' + niCurRow + ':R' + niCurRow + ')', ytdNi), { fmt: ACCT });

  put(ws, 'B' + endRow, 'Ending (Equity)/Deficit', { font: { bold: true } });
  for (let i = 0; i < 12; i++) {
    const parts = [MC(i) + begRow, ...actKeys.map((k) => MC(i) + totals[k]), MC(i) + niCurRow];
    put(ws, MC(i) + endRow, f(parts.join('+'), endVals[i]), { fmt: ACCT, border: { top: THIN, bottom: THIN } });
  }
  put(ws, 'F' + endRow, null, { fill: C.gray, border: { top: THIN, bottom: THIN } });
  const sParts = ['S' + begRow, ...actKeys.map((k) => 'S' + totals[k]), 'S' + niCurRow];
  put(ws, 'S' + endRow, f(sParts.join('+'), endVals[M - 1]), { font: { bold: true }, fmt: ACCT, fill: C.yellow, border: { top: THIN, bottom: THIN } });

  // Tie to the general ledger.
  const glRow = endRow + 2;
  put(ws, 'B' + glRow, 'Total equity per general ledger at ' + shortDate(d.per.end) + ' (incl. current-year net income)', { font: { italic: true } });
  put(ws, 'S' + glRow, d.ties.equity, { fmt: ACCT });
  put(ws, 'B' + (glRow + 1), 'Variance', { font: { italic: true } });
  put(ws, 'S' + (glRow + 1), f('S' + endRow + '-S' + glRow, r2(endVals[M - 1] - d.ties.equity)), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT });

  // Frame + the gray Opening Equity column, as CLA draws it.
  const last = glRow + 2;
  for (let rr = 14; rr <= last; rr++) {
    const fc = ws.getCell('F' + rr); if (!fc.fill || fc.fill.type !== 'pattern') fc.fill = fill(C.gray);
    ws.getCell('B' + rr).border = Object.assign({}, ws.getCell('B' + rr).border || {}, { left: THIN });
    ws.getCell('S' + rr).border = Object.assign({}, ws.getCell('S' + rr).border || {}, { right: THIN });
  }
  rule(ws, last, 2, 19, 'bottom', THIN);
  return { ws, endRow, niCurRow, ytdCell: 'S' + endRow };
}

function buildYtdNi(wb, d, eq) {
  const ws = wb.addWorksheet('YTD NI', { views: [{ showGridLines: false }] });
  widths(ws, { A: 31.7 });
  for (let c = 2; c <= 14; c++) ws.getColumn(c).width = 13.7;
  const M = d.per.mo;
  put(ws, 'A1', 'Location', { font: { bold: true, color: { argb: 'FF000000' } }, align: 'center', border: { bottom: THIN } });
  MON3.forEach((m, i) => put(ws, col(2 + i) + '1', m, { font: { bold: true, color: { argb: 'FF000000' } }, align: 'center', border: { bottom: THIN, left: i === 0 ? THIN : undefined } }));
  put(ws, 'N1', 'Total', { font: { bold: true, color: { argb: 'FF000000' } }, align: 'center', border: { bottom: THIN, left: THIN } });
  let r = 3;
  const first = r;
  for (const l of d.niByLoc) {
    put(ws, 'A' + r, l.loc, { align: 'right' });
    for (let i = 0; i < 12; i++) put(ws, col(2 + i) + r, i < M && Math.abs(l.months[i]) >= 0.005 ? l.months[i] : null, { fmt: ACCT, border: i === 0 ? { left: THIN } : undefined });
    put(ws, 'N' + r, f('SUM(B' + r + ':M' + r + ')', r2(l.months.slice(0, M).reduce((s, v) => s + v, 0))), { fmt: ACCT, border: { left: THIN } });
    r++;
  }
  const last = r - 1;
  for (let c = 2; c <= 14; c++) ws.getCell(col(c) + r).border = { bottom: THIN, left: (c === 2 || c === 14) ? THIN : undefined };
  r++;
  const tot = r;
  for (let i = 0; i < 12; i++) {
    const v = i < M ? r2(-d.niMonths[i]) : 0;
    put(ws, col(2 + i) + tot, f('SUM(' + col(2 + i) + first + ':' + col(2 + i) + last + ')', v), { font: { bold: true }, fmt: ACCT, border: i === 0 ? { left: THIN } : undefined });
  }
  put(ws, 'N' + tot, f('SUM(N' + first + ':N' + last + ')', r2(-d.niMonths.slice(0, M).reduce((s, v) => s + v, 0))), { font: { bold: true }, fmt: ACCT, fill: C.yellow, border: { left: THIN } });
  const chk = tot + 2;
  put(ws, 'A' + chk, 'Check', { font: { italic: true, size: 8 }, align: 'right' });
  for (let i = 0; i < 12; i++) {
    const mc = col(7 + i);
    put(ws, col(2 + i) + chk, f(col(2 + i) + tot + "+'Equity Rollforward'!" + mc + eq.niCurRow, 0), { fmt: ACCT });
  }
  return ws;
}

// ── Index (first tab) ────────────────────────────────────────────────────────
function buildIndex(ws, d, refs) {
  widths(ws, { A: 3.4, B: 8, C: 34, D: 20, E: 20, F: 16, G: 50 });
  leadHeader(ws, 'WORKPAPER INDEX', d.per);
  headerRow(ws, 7, 2, ['Ref', 'Workpaper', 'Lead Sheet Balance', 'Per General Ledger', 'Variance', 'Accounts'], ['center', null, 'center', 'center', 'center', 'left']);
  const rows = [
    ['01', 'Cash and Cash Equivalents', 'Cash Leadsheet', "'Cash Leadsheet'!F" + refs.cash, d.ties.cash, d.cash],
    ['06', 'Other Assets', 'Other Assets Leadsheet', "'Other Assets Leadsheet'!E" + refs.other, d.ties.other_assets, d.other],
    ['12', 'Debt', 'Loan Leadsheet', "'Loan Leadsheet'!E" + refs.loans, d.ties.loans, d.loans],
    ['15', 'Equity Rollforward', 'Equity Rollforward', "'Equity Rollforward'!" + refs.equity, d.ties.equity, null],
  ];
  let r = 8;
  for (const [ref, name, sheet, formula, gl, list] of rows) {
    put(ws, 'B' + r, ref, { align: 'center', fill: C.entry, border: boxAll });
    put(ws, 'C' + r, hyper(sheet, 'A1', name), { font: { bold: true, italic: true, underline: true, color: { argb: C.link } } });
    put(ws, 'D' + r, f(formula, gl), { fmt: ACCT });
    put(ws, 'E' + r, gl, { fmt: ACCT });
    put(ws, 'F' + r, f('D' + r + '-E' + r, 0), { font: { bold: true, color: { argb: C.red } }, fmt: ACCT });
    put(ws, 'G' + r, list ? (list.map((a) => a.code).join(', ') || '—') : 'All equity accounts + current-year net income');
    r++;
  }
  edge(ws, 7, 2, r - 1, 7, MED);
  put(ws, 'B' + (r + 1), 'Balances are shown debit-positive (credits in parentheses), as on CLA’s lead sheets. Blue cells are entry cells; every other figure links to its supporting tab.', { font: { italic: true, color: { argb: 'FF595959' } } });

  r += 3;
  const fl = d.flags || [];
  put(ws, 'B' + r, fl.length ? ('Review items (' + fl.length + ')') : 'No review items — every lead sheet ties to the general ledger and every bank account is reconciled.', { font: { bold: true, color: { argb: fl.length ? 'FFC00000' : 'FF008000' } } });
  r++;
  for (const x of fl) {
    put(ws, 'B' + r, x.severity === 'exception' ? '✖' : '△', { font: { color: { argb: x.severity === 'exception' ? 'FFC00000' : 'FFB8860B' } }, align: 'center' });
    put(ws, 'C' + r, x.message);
    r++;
  }
  if (d.uncovered.length) {
    r++;
    put(ws, 'B' + r, 'Balance-sheet accounts supported by other workpapers (not part of these four)', { font: { bold: true } }); r++;
    headerRow(ws, r, 2, ['Acct', 'Account Name', 'GL Balance'], ['center', null, 'center']); r++;
    for (const a of d.uncovered) {
      put(ws, 'B' + r, Number(a.code) || a.code, { align: 'center' });
      put(ws, 'C' + r, a.name);
      put(ws, 'D' + r, a.end, { fmt: ACCT });
      r++;
    }
  }
}

function buildWorkbook(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const idx = wb.addWorksheet('Index', { views: [{ showGridLines: false }] });
  const cash = buildCash(wb, d);
  const other = buildOther(wb, d);
  const loans = buildLoans(wb, d);
  const eq = buildEquity(wb, d);
  buildYtdNi(wb, d, eq);
  buildIndex(idx, d, { cash: cash.totalRow, other: other.totalRow, loans: loans.totalRow, equity: eq.ytdCell });
  return wb;
}

// ── Persistence ──────────────────────────────────────────────────────────────
const folderFor = (per) => 'Workpapers/Midco Workpapers/' + per.year + '/' + per.monthFolder;
const fileNameFor = (per) => 'CLRFI_Midco_I_Workpapers_' + per.label + '.xlsx';

function saveToWorkpapers(ctx, eid, per, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(per), original = fileNameFor(per);
  const parts = folder.split('/');
  const ins = db.prepare("INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid)); fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare("INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function registerMidcoWorkpapersRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/midco/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const ent = ctx.db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
        if (!isMidcoEntity(ent)) return res.status(400).json({ error: 'The Midco workpapers are set up for CLRFI Midco I only.' });
        const per = resolvePeriod((req.body && (req.body.month_end || req.body.quarter_end)) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, per, eid);
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, per, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Midco-Summary', JSON.stringify({
          period: MONTHS[per.mo - 1] + ' ' + per.y, month_end: per.end, prior_period: per.rollBeg,
          saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          ties: data.ties, accounts: { cash: data.cash.length, other: data.other.length, loans: data.loans.length },
          flags: data.flags,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerMidcoWorkpapersRoutes, resolvePeriod, buildData, buildWorkbook, isMidcoEntity };
