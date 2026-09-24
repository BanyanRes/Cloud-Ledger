// ─── Monthly Closing Workpaper (CLA leadsheet format) ─────────────────────────
//
// One monthly workbook that supports EVERY balance-sheet account of an entity,
// fully GL-derived, laid out in the CLA numbered-leadsheet style. For each
// balance-sheet category (Cash, Accounts Receivable, Prepaid Expenses, Fixed
// Assets, Other Assets, Investments, Intercompany, Accounts Payable, Credit
// Cards, Debt, Other Liabilities, Equity) the workbook holds:
//
//   • A "<Category> Leadsheet" summary tab in the exact CLA display style —
//     gridlines off, Calibri, the entity/title block with yellow "Month Ended:"
//     and "Year Ended:" entry cells and an "= Entry Cell" legend, then a lead
//     table: Account No. | Account Name | Worksheet (HYPERLINK to the account's
//     supporting tab) | Balance (a reference to that tab) | FQ Anchor
//     ("#fq-<code>") | Comments, footed with =SUBTOTAL(109, …). Balances use the
//     accounting number format _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_).
//   • A supporting roll-forward tab per account (named by account code, the
//     HYPERLINK / #fq target): Beginning (prior month end per GL) + the month's
//     GL activity = Ending. Cash adds a blue "per bank statement" input and an
//     unreconciled-difference line.
//
// Design rules (per the workpaper convention): no hard-coded derived amounts —
// each ending balance is Beginning + SUM(activity), each leadsheet balance is a
// reference to its supporting tab, every total is =SUBTOTAL/…, and the Assets =
// Liabilities + Equity tie is a live formula. The only literals are the atomic GL
// line amounts and the prior-month-end opening balances (straight from the
// ledger), plus blank blue input cells. The chart of accounts is read at run
// time, so the generator is entity-agnostic.
//
// Built like the other workpapers (buildData -> buildWorkbook -> saveToWorkpapers)
// and registered by index.js. ctx = { db, auth, requireEntityAccess, requireRole,
// workpapersDir, computeBalances }.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const pad2 = (n) => String(n).padStart(2, '0');
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Resolve any date to the month end that contains it, plus the prior month end.
function resolveMonth(dateInput) {
  const m = String(dateInput || '').match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) throw new Error('month_end must be a date in YYYY-MM-DD form');
  const y = Number(m[1]), mo = Number(m[2]);
  if (mo < 1 || mo > 12) throw new Error('month_end must be a valid date');
  const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  const end = y + '-' + pad2(mo) + '-' + pad2(last);
  // Beginning = prior month end (last day of the month before this one).
  const beg = new Date(Date.UTC(y, mo - 1, 0)).toISOString().slice(0, 10);
  return {
    end, beg,
    year: String(y), monthNum: mo, monthName: MONTHS[mo - 1],
    label: y + '-' + pad2(mo),
    yearStart: y + '-01-01',
    yearEnd: y + '-12-31',
  };
}

// ── GL transaction detail for the month, per account, signed the way the balance
//    moves (debit-natural accounts positive on a debit, etc.).
function glMonth(db, eid, from, to) {
  const rows = db.prepare(`
    SELECT je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number,
           je.vendor AS vendor, je.memo AS memo,
           jl.account_code AS account_code, a.type AS account_type,
           jl.debit AS debit, jl.credit AS credit, jl.description AS description,
           dc.name AS class_name, dl.name AS location_name
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    LEFT JOIN dim_classes dc ON dc.id = jl.class_id
    LEFT JOIN dim_locations dl ON dl.id = jl.location_id
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ?
    ORDER BY jl.account_code, je.date, je.entry_num, jl.id
  `).all(eid, from, to);
  const byAcct = new Map();
  for (const r of rows) {
    const isDr = r.account_type === 'Asset' || r.account_type === 'Expense';
    const signed = r2(isDr ? (r.debit - r.credit) : (r.credit - r.debit));
    if (!byAcct.has(r.account_code)) byAcct.set(r.account_code, []);
    byAcct.get(r.account_code).push({
      date: r.date, num: r.entry_num || r.doc_number || '',
      payee: r.vendor || '', memo: r.memo || r.description || '',
      class_name: r.class_name || '', location_name: r.location_name || '',
      signed,
    });
  }
  return byAcct;
}

function balMap(computeBalances, eid, asOf, closeBefore) {
  const rows = computeBalances(eid, { as_of: asOf, close_pl_before: closeBefore }) || [];
  const m = new Map();
  for (const r of rows) m.set(String(r.code), { name: r.name, type: r.type, balance: r2(r.balance), bank_acct: r.bank_acct });
  return m;
}

function pnl(db, eid, from, to) {
  const rows = db.prepare(`
    SELECT jl.account_code code, a.name name, a.type type,
           SUM(jl.debit) td, SUM(jl.credit) tc
    FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? AND a.type IN ('Revenue','Expense')
    GROUP BY jl.account_code
  `).all(eid, from, to);
  const revenue = [], expense = [];
  for (const r of rows) {
    if (r.type === 'Revenue') { const amt = r2((r.tc || 0) - (r.td || 0)); if (Math.abs(amt) >= 0.005) revenue.push({ code: r.code, name: r.name || '', amt }); }
    else { const amt = r2((r.td || 0) - (r.tc || 0)); if (Math.abs(amt) >= 0.005) expense.push({ code: r.code, name: r.name || '', amt }); }
  }
  revenue.sort((a, b) => String(a.code).localeCompare(String(b.code)));
  expense.sort((a, b) => String(a.code).localeCompare(String(b.code)));
  const totRev = r2(revenue.reduce((s, x) => s + x.amt, 0));
  const totExp = r2(expense.reduce((s, x) => s + x.amt, 0));
  return { revenue, expense, totRev, totExp, net: r2(totRev - totExp) };
}

// ── Account categorization (generic, first match wins), mapped to the CLA
//    numbered-leadsheet set. ────────────────────────────────────────────────
function categoryOf(code, name, type, bank) {
  const c = String(code || ''), n = String(name || '').toLowerCase();
  const isInterco = /\bdue\s+(to|from)\b/.test(n) || /intercompan/.test(n);
  if (isInterco) return 'interco';
  if (type === 'Asset') {
    if (/^10/.test(c) || bank === 1 || /\bcash\b|checking|savings|money market|operating account/.test(n)) return 'cash';
    if (/receivable/.test(n)) return 'ar';
    if (/^prepaid\b|prepaid /.test(n)) return 'prepaid';
    if (/fixed asset|equipment|vehicle|building|furniture|leasehold|accumulated (dep|amort)|land\b|improvement/.test(n) || /^1[56]/.test(c)) return 'fixed';
    if (/^investment\b/.test(n) || /^19/.test(c)) return 'invest';
    return 'otherassets';
  }
  if (type === 'Liability') {
    if (/credit card/.test(n)) return 'cc';
    if (/loan|notes?\s+payable|line of credit|mortgage|bond/.test(n) || /^25/.test(c)) return 'debt';
    if (/accounts payable|\bpayable\b|accrued/.test(n) || /^2[01]/.test(c)) return 'ap';
    return 'otherliab';
  }
  if (type === 'Equity') return 'equity';
  return 'otherassets';
}
// Ordered category list — one CLA-styled leadsheet tab per non-empty category.
const CAT_TABS = [
  { key: 'cash', tab: 'Cash Leadsheet', lead: 'CASH LEADSHEET', total: 'Total Cash' },
  { key: 'ar', tab: 'AR Leadsheet', lead: 'ACCOUNTS RECEIVABLE LEADSHEET', total: 'Total Receivables' },
  { key: 'prepaid', tab: 'Prepaid Leadsheet', lead: 'PREPAID EXPENSES LEAD SHEET', total: 'Total Prepaid Expenses' },
  { key: 'fixed', tab: 'Fixed Assets Leadsheet', lead: 'FIXED ASSETS LEADSHEET', total: 'Total Fixed Assets' },
  { key: 'otherassets', tab: 'Other Assets Leadsheet', lead: 'OTHER ASSETS LEADSHEET', total: 'Total Other Assets' },
  { key: 'invest', tab: 'Investments Leadsheet', lead: 'INVESTMENTS LEADSHEET', total: 'Total Investments' },
  { key: 'interco', tab: 'Intercompany Leadsheet', lead: 'INTERCOMPANY LEADSHEET', total: 'Total Intercompany' },
  { key: 'ap', tab: 'AP Leadsheet', lead: 'ACCOUNTS PAYABLE LEADSHEET', total: 'Total Payables' },
  { key: 'cc', tab: 'Credit Cards Leadsheet', lead: 'CREDIT CARDS LEADSHEET', total: 'Total Credit Cards' },
  { key: 'debt', tab: 'Debt Leadsheet', lead: 'DEBT LEADSHEET', total: 'Total Debt' },
  { key: 'otherliab', tab: 'Other Liabilities Leadsheet', lead: 'OTHER LIABILITIES LEADSHEET', total: 'Total Other Liabilities' },
  { key: 'equity', tab: 'Equity Leadsheet', lead: 'EQUITY LEADSHEET', total: 'Total Equity' },
];

// A token to find this entity on a counterparty's ledger ("Odyssey Holdings LLC"
// -> /odyssey/), using the first distinctive word of the entity name.
function selfToken(entityName) {
  const stop = new Set(['the', 'a', 'an', 'llc', 'lp', 'inc', 'l.p.', 'l.l.c.', 'fund', 'holdings', 'holding', 'company', 'co', 'partners']);
  const words = String(entityName || '').replace(/[.,]/g, ' ').split(/\s+/).filter(Boolean);
  const w = words.find((x) => !stop.has(x.toLowerCase())) || words[0] || '';
  return w ? new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') : null;
}

// Net amount a counterparty entity owes TO this entity per the counterparty's own
// GL: their liabilities referencing us count positive, their assets negative.
function counterpartyNet(computeBalances, cpEid, selfRe, asOf) {
  const rows = computeBalances(cpEid, { as_of: asOf }) || [];
  let net = 0; const legs = [];
  for (const r of rows) {
    if (!selfRe.test(String(r.name || ''))) continue;
    if (r.type === 'Liability') { net = r2(net + (r.balance || 0)); legs.push({ code: String(r.code), name: r.name, amt: r2(r.balance) }); }
    else if (r.type === 'Asset') { net = r2(net - (r.balance || 0)); legs.push({ code: String(r.code), name: r.name, amt: r2(-(r.balance || 0)) }); }
  }
  return { net: r2(net), legs };
}

// Resolve the counterparty NAME embedded in a "Due to/from X" account to a CL
// entity. Returns { id, name } or null.
function resolveCounterparty(entities, acctName, selfId) {
  const m = String(acctName || '').match(/\bdue\s+(?:to|from)\s+(.+)$/i);
  if (!m) return null;
  let x = m[1].replace(/\b(llc|l\.l\.c\.|lp|l\.p\.|inc\.?|holdings?|management|the)\b/gi, ' ').replace(/[^A-Za-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!x || x.length < 3) return null;
  const toks = x.split(' ').filter((t) => t.length >= 3);
  if (!toks.length) return null;
  let best = null, bestScore = 0;
  for (const e of entities) {
    if (e.id === selfId) continue;
    const en = String(e.name || '').toLowerCase();
    let score = 0;
    for (const t of toks) if (en.includes(t)) score += t.length;
    if (score > bestScore) { bestScore = score; best = e; }
  }
  if (best && bestScore >= Math.max(4, toks[0].length)) return { id: best.id, name: best.name };
  return null;
}

// ── Build all the data the workbook needs. ───────────────────────────────────
function buildData(ctx, m, eid) {
  const { db, computeBalances } = ctx;
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  if (!ent) throw new Error('Entity ' + eid + ' not found');
  const entities = db.prepare('SELECT id, name FROM entities').all();
  const accounts = db.prepare('SELECT code, name, type, bank_acct FROM accounts WHERE entity_id = ?').all(eid);
  const acctInfo = new Map();
  for (const a of accounts) acctInfo.set(String(a.code), a);

  const begBS = balMap(computeBalances, eid, m.beg, m.yearStart);
  const endBS = balMap(computeBalances, eid, m.end, m.yearStart);
  const lines = glMonth(db, eid, m.beg === m.end ? m.end : addDay(m.beg), m.end);

  // Union of BS accounts touched (nonzero begin, end, or activity).
  const codes = new Set();
  for (const [c, v] of begBS) if (isBS(v.type) && Math.abs(v.balance) >= 0.005) codes.add(c);
  for (const [c, v] of endBS) if (isBS(v.type) && Math.abs(v.balance) >= 0.005) codes.add(c);
  for (const [c, arr] of lines) { const info = acctInfo.get(c); if (info && isBS(info.type) && arr.some((x) => Math.abs(x.signed) >= 0.005)) codes.add(c); }

  const acctRows = [];
  for (const c of codes) {
    const info = acctInfo.get(c) || {};
    const type = info.type || (endBS.get(c) || begBS.get(c) || {}).type || 'Asset';
    const name = info.name || (endBS.get(c) || begBS.get(c) || {}).name || c;
    const begin = r2((begBS.get(c) || {}).balance || 0);
    const end = r2((endBS.get(c) || {}).balance || 0);
    const acctLines = (lines.get(c) || []).filter((x) => Math.abs(x.signed) >= 0.005);
    const activity = r2(acctLines.reduce((s, x) => s + x.signed, 0));
    const cat = categoryOf(c, name, type, info.bank_acct);
    acctRows.push({ code: c, name, type, begin, end, activity, lines: acctLines, cat, bank_acct: info.bank_acct });
  }
  acctRows.sort((a, b) => (a.type === b.type ? String(a.code).localeCompare(String(b.code)) : typeOrder(a.type) - typeOrder(b.type)));

  // Intercompany counterparty resolution + mirror.
  const selfRe = selfToken(ent.name);
  for (const a of acctRows) {
    if (a.cat !== 'interco') continue;
    const cp = resolveCounterparty(entities, a.name, eid);
    a.cp = cp;
    if (cp && selfRe) {
      const mir = counterpartyNet(computeBalances, cp.id, selfRe, m.end);
      a.cpNet = mir.net;
      a.cpLegs = mir.legs;
    }
  }

  // Period P&L (fiscal-YTD) for the equity net-income line.
  const niEnd = pnl(db, eid, m.yearStart, m.end);
  const niBegVal = m.beg < m.yearStart ? 0 : pnl(db, eid, m.yearStart, m.beg).net;

  // Ties / flags (for the Summary banner).
  const totA = r2(acctRows.filter((a) => a.type === 'Asset').reduce((s, a) => s + a.end, 0));
  const totL = r2(acctRows.filter((a) => a.type === 'Liability').reduce((s, a) => s + a.end, 0));
  const totE = r2(acctRows.filter((a) => a.type === 'Equity').reduce((s, a) => s + a.end, 0));
  const imbalance = r2(totA - (totL + totE + niEnd.net));
  const flags = [];
  if (Math.abs(imbalance) >= 0.01) {
    flags.push({ severity: 'exception', wp: 'Lead Sheet', message: 'Balance sheet does not tie: Assets ' + fmt(totA) + ' ≠ Liabilities ' + fmt(totL) + ' + Equity ' + fmt(totE) + ' + Net income ' + fmt(niEnd.net) + ' (off by ' + fmt(imbalance) + ').' });
  }
  for (const a of acctRows) {
    const rollDiff = r2(a.begin + a.activity - a.end);
    if (Math.abs(rollDiff) >= 0.01) {
      flags.push({ severity: 'exception', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': roll-forward does not tie — beginning ' + fmt(a.begin) + ' + activity ' + fmt(a.activity) + ' = ' + fmt(a.begin + a.activity) + ', but GL ending is ' + fmt(a.end) + ' (off by ' + fmt(rollDiff) + ').' });
    }
  }
  for (const a of acctRows) {
    if (a.cat !== 'interco') continue;
    if (!a.cp) {
      if (Math.abs(a.end) >= 0.01) flags.push({ severity: 'review', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': ' + fmt(a.end) + ' outstanding — no matching CL entity found to tie against; confirm the counterparty balance manually.' });
      continue;
    }
    const ourNet = a.type === 'Asset' ? a.end : -a.end;
    const diff = r2(ourNet - (a.cpNet || 0));
    if (Math.abs(diff) >= 0.01) {
      flags.push({ severity: 'exception', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': ' + ent.name + ' shows ' + fmt(ourNet) + ' owed by ' + a.cp.name + ', but ' + a.cp.name + '’s ledger shows ' + fmt(a.cpNet || 0) + ' (off by ' + fmt(diff) + ').' });
    }
  }

  return {
    entity: ent, month: m, acctRows,
    ni: { end: niEnd, begVal: niBegVal },
    ties: { total_assets: totA, total_liabilities: totL, total_equity: totE, net_income: niEnd.net, imbalance },
    flags,
  };
}

function addDay(d) { const [y, m, dd] = d.split('-').map(Number); return new Date(Date.UTC(y, m - 1, dd + 1)).toISOString().slice(0, 10); }
function isBS(t) { return t === 'Asset' || t === 'Liability' || t === 'Equity'; }
function typeOrder(t) { return t === 'Asset' ? 0 : t === 'Liability' ? 1 : t === 'Equity' ? 2 : 3; }
const fmt = (n) => '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ─── Workbook (CLA leadsheet display) ─────────────────────────────────────────
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'; // accounting number format
const DATEFMT = 'mm-dd-yy';
const FONT = 'Calibri';
const YELLOW = 'FFFFFF00';
const RED = 'FFFF0000';
const HDRFILL = 'FFDDEBF7'; // pale blue table header (Table Style 1-ish)
const F = (o = {}) => Object.assign({ name: FONT, size: 11 }, o);
const THIN = { style: 'thin' };
const DBL = { style: 'double' };
const ENTRY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: YELLOW } };
const INPUT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const short = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return mo + '/' + dd + '/' + String(y).slice(2); };
const spell = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return MONTHS[mo - 1] + ' ' + dd + ', ' + y; };
const asDate = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return new Date(Date.UTC(y, mo - 1, dd)); };
const qn = (s) => "'" + String(s).replace(/'/g, "''") + "'"; // quoted sheet name for a formula ref

// A safe, unique worksheet name from an account code (Excel: <=31 chars, no []:*?/\).
function sheetNameFor(code, used) {
  let base = String(code).replace(/[\[\]\:\*\?\/\\]/g, '_').slice(0, 31) || 'acct';
  let name = base, i = 1;
  while (used.has(name.toLowerCase())) { const suf = '_' + (++i); name = base.slice(0, 31 - suf.length) + suf; }
  used.add(name.toLowerCase());
  return name;
}

// The CLA leadsheet title block: entity, "<X> LEADSHEET", Month/Year Ended entry
// cells and an "= Entry Cell" legend. Returns nothing; keys everything off D3/D4.
function titleBlock(ws, entityName, leadTitle, m, legendCol) {
  ws.getCell('C1').value = entityName; ws.getCell('C1').font = F({ size: 16, bold: true });
  ws.getCell('C2').value = leadTitle; ws.getCell('C2').font = F({ size: 12, bold: true });
  ws.getCell('C3').value = 'Month Ended:'; ws.getCell('C3').font = F();
  const d3 = ws.getCell('D3');
  d3.value = asDate(m.end); d3.numFmt = DATEFMT; d3.font = F({ bold: true, color: { argb: RED } });
  d3.fill = ENTRY_FILL; d3.alignment = { horizontal: 'center' };
  d3.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  ws.getCell('C4').value = 'Year Ended:'; ws.getCell('C4').font = F();
  const d4 = ws.getCell('D4');
  d4.value = { formula: 'DATE(YEAR($D$3),12,31)' }; d4.numFmt = DATEFMT; d4.font = F({ bold: true });
  d4.alignment = { horizontal: 'center' };
  d4.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  const lc = legendCol || 'G';
  const leg = ws.getCell(lc + '3'); leg.value = '= Entry Cell'; leg.font = F();
  const sw = ws.getCell(lc + '2'); sw.fill = ENTRY_FILL; sw.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
}

// One CLA-style category Leadsheet tab. `rowsForCat` are acctRows; `ref` maps
// code -> { sheet, endRow } for the supporting-tab balance link. Returns the
// cell (e.g. "'Cash Leadsheet'!$E$13") holding this category's SUBTOTAL, so the
// balance-sheet tie can reference it.
function buildLeadsheet(wb, cd, entityName, m, rowsForCat, ref) {
  const interco = cd.key === 'interco';
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4;
  ws.getColumn('B').width = 13; ws.getColumn('C').width = 34; ws.getColumn('D').width = 16;
  ws.getColumn('E').width = 16; ws.getColumn('F').width = 15; ws.getColumn('G').width = 40;
  if (interco) { ws.getColumn('G').width = 17; ws.getColumn('H').width = 17; ws.getColumn('I').width = 40; }

  titleBlock(ws, entityName, cd.lead, m, interco ? 'I' : 'G');

  // Header row (row 7).
  const HR = 7;
  const cols = interco
    ? ['Account No.', 'Account Name', 'Worksheet', 'Balance', 'FQ Anchor', 'Other Entity Bal', 'Variance', 'Comments']
    : ['Account No.', 'Account Name', 'Worksheet', 'Balance', 'FQ Anchor', 'Comments'];
  cols.forEach((t, i) => {
    const c = ws.getRow(HR).getCell(i + 2); // start at column B
    c.value = t; c.font = F({ bold: true });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDRFILL } };
    c.alignment = { horizontal: 'center', wrapText: true };
    c.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  });

  let r = HR + 1;
  const first = r;
  for (const a of rowsForCat) {
    const rr = ref.get(a.code) || {};
    const sheet = rr.sheet;
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.name; ws.getCell('C' + r).font = F(); ws.getCell('C' + r).alignment = { horizontal: 'left' };
    const wcell = ws.getCell('D' + r);
    if (sheet) { wcell.value = { formula: 'HYPERLINK("#' + qn(sheet) + '!A1","\u2192 WP")' }; }
    else wcell.value = '';
    wcell.font = F({ bold: true, underline: true, color: { argb: 'FF0563C1' } }); wcell.alignment = { horizontal: 'center' };
    const bcell = ws.getCell('E' + r); bcell.numFmt = ACCT; bcell.font = F();
    if (sheet && rr.endRow) bcell.value = { formula: qn(sheet) + '!$E$' + rr.endRow };
    else bcell.value = a.end;
    const fq = ws.getCell('F' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    if (interco) {
      const oc = ws.getCell('G' + r); oc.numFmt = ACCT; oc.font = F();
      const vc = ws.getCell('H' + r); vc.numFmt = ACCT; vc.font = F();
      if (a.cp) {
        oc.value = a.cpNet || 0;
        // Our net receivable on this account (asset +, liability -) less the cp's net.
        const ourRef = (a.type === 'Asset' ? '' : '-') + 'E' + r;
        vc.value = { formula: ourRef + '-G' + r };
      } else { oc.value = ''; vc.value = ''; }
    }
    r++;
  }
  const last = r - 1;

  // Total row (SUBTOTAL so it ignores any filtered rows, matching the CLA sheets).
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  const totCell = ws.getCell('E' + r); totCell.numFmt = ACCT; totCell.font = F({ bold: true });
  totCell.border = { top: THIN, bottom: DBL };
  if (last >= first) totCell.value = { formula: 'SUBTOTAL(109,E' + first + ':E' + last + ')' };
  else totCell.value = 0;
  if (interco) {
    const ov = ws.getCell('G' + r); ov.numFmt = ACCT; ov.font = F({ bold: true }); ov.border = { top: THIN, bottom: DBL };
    ov.value = last >= first ? { formula: 'SUBTOTAL(109,G' + first + ':G' + last + ')' } : 0;
    const vv = ws.getCell('H' + r); vv.numFmt = ACCT; vv.font = F({ bold: true }); vv.border = { top: THIN, bottom: DBL };
    vv.value = last >= first ? { formula: 'SUBTOTAL(109,H' + first + ':H' + last + ')' } : 0;
  }
  return qn(cd.tab) + '!$E$' + r;
}

// One supporting roll-forward tab per account (the HYPERLINK / #fq target).
// Records ref[code] = { sheet, endRow }. Returns nothing.
function buildAccountTab(wb, a, entityName, m, ref, used) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 12; ws.getColumn('B').width = 12; ws.getColumn('C').width = 24;
  ws.getColumn('D').width = 52; ws.getColumn('E').width = 16;
  // FQ anchor marker (top-left) so "#fq-<code>" on the leadsheet is meaningful.
  ws.getCell('A1').value = '#fq-' + a.code; ws.getCell('A1').font = F({ size: 8, italic: true, color: { argb: 'FFBFBFBF' } });
  ws.getCell('C1').value = entityName; ws.getCell('C1').font = F({ size: 12, bold: true });
  ws.getCell('C2').value = a.code + '  \u2014  ' + a.name; ws.getCell('C2').font = F({ bold: true });
  ws.getCell('C3').value = 'Roll-forward \u2014 month ended ' + spell(m.end); ws.getCell('C3').font = F({ italic: true });

  const HR = 5;
  ['Date', 'Num', 'Payee', 'Description / Memo', 'Amount'].forEach((t, i) => {
    const c = ws.getRow(HR).getCell(i + 1);
    c.value = t; c.font = F({ bold: true });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDRFILL } };
    c.alignment = { horizontal: 'center' };
    c.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  });
  let r = HR + 1;
  // Beginning balance (prior month end, per GL).
  ws.getCell('D' + r).value = 'Beginning balance \u2014 ' + short(m.beg) + ' (per GL)'; ws.getCell('D' + r).font = F({ italic: true });
  const begRow = r; const bc = ws.getCell('E' + r); bc.value = a.begin; bc.numFmt = ACCT; bc.font = F();
  r++;
  const firstLine = r;
  for (const ln of a.lines) {
    ws.getCell('A' + r).value = ln.date; ws.getCell('A' + r).font = F();
    ws.getCell('B' + r).value = ln.num; ws.getCell('B' + r).font = F();
    ws.getCell('C' + r).value = ln.payee || ln.location_name || ln.class_name || ''; ws.getCell('C' + r).font = F();
    ws.getCell('D' + r).value = ln.memo; ws.getCell('D' + r).font = F();
    const ec = ws.getCell('E' + r); ec.value = ln.signed; ec.numFmt = ACCT; ec.font = F();
    r++;
  }
  const lastLine = r - 1;
  ws.getCell('D' + r).value = 'Total activity for the month'; ws.getCell('D' + r).font = F({ bold: true });
  const actRow = r; const ac = ws.getCell('E' + r); ac.numFmt = ACCT; ac.font = F({ bold: true }); ac.border = { top: THIN };
  ac.value = lastLine >= firstLine ? { formula: 'SUM(E' + firstLine + ':E' + lastLine + ')' } : 0;
  r++;
  ws.getCell('D' + r).value = 'Ending balance \u2014 ' + short(m.end); ws.getCell('D' + r).font = F({ bold: true });
  const endRow = r; const enc = ws.getCell('E' + r); enc.numFmt = ACCT; enc.font = F({ bold: true }); enc.border = { top: THIN, bottom: DBL };
  enc.value = { formula: 'E' + begRow + '+E' + actRow };
  r++;
  ref.set(a.code, { sheet, endRow });

  // Cash: bank-rec input + unreconciled difference.
  if (a.cat === 'cash') {
    ws.getCell('D' + r).value = 'Per bank statement (enter)'; ws.getCell('D' + r).font = F({ color: { argb: 'FF0000FF' } });
    const brc = ws.getCell('E' + r); brc.numFmt = ACCT; brc.font = F({ color: { argb: 'FF0000FF' } }); brc.fill = INPUT_FILL; brc.border = { top: THIN };
    const bankRow = r; r++;
    ws.getCell('D' + r).value = 'Unreconciled difference (GL \u2212 bank)'; ws.getCell('D' + r).font = F({ italic: true });
    const dc = ws.getCell('E' + r); dc.numFmt = ACCT; dc.font = F({ italic: true });
    dc.value = { formula: 'E' + endRow + '-E' + bankRow };
    r++;
  }
}

function buildWorkbook(data) {
  const { entity, month: m, acctRows, ni } = data;
  const en = entity.name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  // Summary first (tab order = creation order); populated after totals are known.
  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });

  // Supporting account tabs first so leadsheet balances can link to them.
  const ref = new Map(); const used = new Set(['summary']);
  const byCat = {};
  for (const a of acctRows) (byCat[a.cat] = byCat[a.cat] || []).push(a);
  for (const cd of CAT_TABS) used.add(cd.tab.toLowerCase());
  used.add('income statement');
  for (const a of acctRows) buildAccountTab(wb, a, en, m, ref, used);

  // Category leadsheet tabs (only where the category has accounts), recording the
  // total cell of each so the Summary balance-sheet tie can reference it.
  const catTotalCell = {};
  const catType = {
    cash: 'Asset', ar: 'Asset', prepaid: 'Asset', fixed: 'Asset', otherassets: 'Asset', invest: 'Asset',
    interco: 'mixed', ap: 'Liability', cc: 'Liability', debt: 'Liability', otherliab: 'Liability', equity: 'Equity',
  };
  for (const cd of CAT_TABS) {
    const rowsForCat = byCat[cd.key] || [];
    if (!rowsForCat.length) continue;
    catTotalCell[cd.key] = buildLeadsheet(wb, cd, en, m, rowsForCat, ref);
  }

  // Income Statement (fiscal-YTD) supporting tab, for the equity net-income line.
  const pl = wb.addWorksheet('Income Statement', { views: [{ showGridLines: false }] });
  pl.getColumn('A').width = 3.4; pl.getColumn('B').width = 14; pl.getColumn('C').width = 52; pl.getColumn('D').width = 16;
  pl.getCell('C1').value = en; pl.getCell('C1').font = F({ size: 16, bold: true });
  pl.getCell('C2').value = 'STATEMENT OF OPERATIONS \u2014 FISCAL YEAR TO DATE'; pl.getCell('C2').font = F({ size: 12, bold: true });
  pl.getCell('C3').value = m.yearStart + ' to ' + short(m.end); pl.getCell('C3').font = F({ italic: true });
  const PHR = 5;
  ['Code', 'Account', 'Amount'].forEach((t, i) => {
    const c = pl.getRow(PHR).getCell(i + 2);
    c.value = t; c.font = F({ bold: true }); c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDRFILL } };
    c.alignment = { horizontal: 'center' }; c.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  });
  let pr = PHR + 1;
  pl.getCell('B' + pr).value = 'REVENUE'; pl.getCell('B' + pr).font = F({ bold: true }); pr++;
  const revFirst = pr;
  for (const x of ni.end.revenue) { pl.getCell('B' + pr).value = x.code; pl.getCell('B' + pr).font = F(); pl.getCell('C' + pr).value = x.name; pl.getCell('C' + pr).font = F(); const c = pl.getCell('D' + pr); c.value = x.amt; c.numFmt = ACCT; c.font = F(); pr++; }
  const revLast = pr - 1;
  pl.getCell('C' + pr).value = 'Total revenue'; pl.getCell('C' + pr).font = F({ bold: true });
  const totRevCell = 'D' + pr; { const c = pl.getCell(totRevCell); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN }; c.value = revLast >= revFirst ? { formula: 'SUM(D' + revFirst + ':D' + revLast + ')' } : 0; }
  pr += 2;
  pl.getCell('B' + pr).value = 'EXPENSES'; pl.getCell('B' + pr).font = F({ bold: true }); pr++;
  const expFirst = pr;
  for (const x of ni.end.expense) { pl.getCell('B' + pr).value = x.code; pl.getCell('B' + pr).font = F(); pl.getCell('C' + pr).value = x.name; pl.getCell('C' + pr).font = F(); const c = pl.getCell('D' + pr); c.value = x.amt; c.numFmt = ACCT; c.font = F(); pr++; }
  const expLast = pr - 1;
  pl.getCell('C' + pr).value = 'Total expenses'; pl.getCell('C' + pr).font = F({ bold: true });
  const totExpCell = 'D' + pr; { const c = pl.getCell(totExpCell); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN }; c.value = expLast >= expFirst ? { formula: 'SUM(D' + expFirst + ':D' + expLast + ')' } : 0; }
  pr += 2;
  pl.getCell('C' + pr).value = 'NET INCOME (LOSS) \u2014 fiscal YTD'; pl.getCell('C' + pr).font = F({ bold: true });
  const niEndCell = 'D' + pr; { const c = pl.getCell(niEndCell); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL }; c.value = { formula: totRevCell + '-' + totExpCell }; }
  pr += 2;
  pl.getCell('C' + pr).value = 'Net income (loss) \u2014 fiscal YTD through ' + short(m.beg) + ' (beginning of month, per GL)'; pl.getCell('C' + pr).font = F({ italic: true });
  const niBegCell = 'D' + pr; { const c = pl.getCell(niBegCell); c.value = ni.begVal; c.numFmt = ACCT; c.font = F({ italic: true }); }
  const NI_END_REF = qn('Income Statement') + '!$' + niEndCell.replace(/(\d+)/, '$$$1');
  const NI_BEG_REF = qn('Income Statement') + '!$' + niBegCell.replace(/(\d+)/, '$$$1');

  // ── Summary / balance-sheet tie + exceptions. ────────────────────────────
  su.getColumn('A').width = 3.4; su.getColumn('B').width = 60; su.getColumn('C').width = 20; su.getColumn('D').width = 20;
  su.getCell('C1').value = en; su.getCell('C1').font = F({ size: 16, bold: true });
  su.getCell('C2').value = 'MONTHLY CLOSING WORKPAPER \u2014 SUMMARY'; su.getCell('C2').font = F({ size: 12, bold: true });
  su.getCell('C3').value = 'Month ended ' + spell(m.end); su.getCell('C3').font = F({ italic: true });
  let sr = 5;
  const flags = data.flags || [];
  if (!flags.length) {
    su.mergeCells('B' + sr + ':D' + sr);
    const c = su.getCell('B' + sr); c.value = '\u2713  All checks passed \u2014 the balance sheet ties and every account rolls forward from the general ledger.'; c.font = F({ bold: true, color: { argb: 'FF1E7A34' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF7EE' } }; c.alignment = { wrapText: true }; su.getRow(sr).height = 22; sr += 2;
  } else {
    su.mergeCells('B' + sr + ':D' + sr);
    const c = su.getCell('B' + sr); c.value = '\u26A0  ' + flags.length + ' item' + (flags.length > 1 ? 's' : '') + ' require review'; c.font = F({ bold: true, color: { argb: 'FF9C4221' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDECEA' } }; su.getRow(sr).height = 22; sr += 2;
    su.getCell('B' + sr).value = 'Exceptions \u2014 Review Required'; su.getCell('B' + sr).font = F({ bold: true, size: 12 }); sr++;
    for (const f of flags) {
      su.mergeCells('B' + sr + ':D' + sr);
      const cc = su.getCell('B' + sr);
      cc.value = (f.severity === 'exception' ? '\u2716 ' : '\u26A0 ') + f.message;
      cc.font = F({ color: { argb: f.severity === 'exception' ? 'FFB3261E' : 'FF9C4221' } });
      cc.alignment = { wrapText: true, vertical: 'top' };
      su.getRow(sr).height = Math.min(220, 14 * Math.max(1, Math.ceil(String(f.message).length / 100)) + 4);
      sr++;
    }
    sr++;
  }
  su.getCell('B' + sr).value = 'Balance-sheet tie'; su.getCell('B' + sr).font = F({ bold: true, size: 12 }); sr++;
  // Sum the category leadsheet totals by balance-sheet type.
  const sumTypeRefs = (ty) => CAT_TABS.filter((cd) => catTotalCell[cd.key] && (catType[cd.key] === ty || (ty === 'Asset' && cd.key === 'interco'))).map((cd) => catTotalCell[cd.key]);
  const assetRefs = CAT_TABS.filter((cd) => catTotalCell[cd.key] && catType[cd.key] === 'Asset').map((cd) => catTotalCell[cd.key]);
  // Intercompany can be asset- or liability-side; add its net to assets for the tie
  // (a net payable simply shows negative).
  if (catTotalCell.interco) assetRefs.push(catTotalCell.interco);
  const liabRefs = CAT_TABS.filter((cd) => catTotalCell[cd.key] && catType[cd.key] === 'Liability').map((cd) => catTotalCell[cd.key]);
  const eqRefs = CAT_TABS.filter((cd) => catTotalCell[cd.key] && catType[cd.key] === 'Equity').map((cd) => catTotalCell[cd.key]);
  const sumOf = (refs) => refs.length ? refs.join('+') : '0';
  const assetsRow = sr; su.getCell('B' + sr).value = 'Total assets'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(assetRefs) }; } sr++;
  const liabRow = sr; su.getCell('B' + sr).value = 'Total liabilities'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(liabRefs) }; } sr++;
  const eqRow = sr; su.getCell('B' + sr).value = 'Total members\u2019 equity'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(eqRefs) }; } sr++;
  const niRow = sr; su.getCell('B' + sr).value = 'Net income (loss) \u2014 fiscal YTD (per Income Statement)'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: NI_END_REF }; } sr++;
  su.getCell('B' + sr).value = 'Assets \u2212 (Liabilities + Equity + Net income)  \u2014  should be $0.00'; su.getCell('B' + sr).font = F({ bold: true });
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN }; c.value = { formula: 'C' + assetsRow + '-(C' + liabRow + '+C' + eqRow + '+C' + niRow + ')' }; }
  void sumTypeRefs; // (kept for clarity; assets/liab/eq computed explicitly above)

  return wb;
}

// ── Persistence. ─────────────────────────────────────────────────────────────
const folderFor = (m) => 'Workpapers/Monthly Closing/' + m.year;
const fileNameFor = (m) => 'Monthly_Closing_Workpaper_' + m.label + '.xlsx';

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

function registerMonthlyCloseRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/monthly-close/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const m = resolveMonth((req.body && req.body.month_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, m, eid);
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, m, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Monthclose-Summary', JSON.stringify({
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
}

module.exports = { registerMonthlyCloseRoutes, resolveMonth, buildData, buildWorkbook };
