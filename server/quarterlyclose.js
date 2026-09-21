// ─── Quarterly Closing Workpaper (balance-sheet account support) ───────────────
//
// A single quarterly workbook that supports EVERY balance-sheet account of an
// entity, fully GL-derived, built as a set of supporting schedules with a Lead
// Sheet that LINKS to them. Design rules (per the workpaper convention):
//
//   • No hard-coded amounts. Every derived figure is an Excel formula:
//       – each account's Ending  = Beginning + SUM(its month GL lines)
//       – each Lead Sheet cell   = a reference to the supporting-tab cell
//       – every subtotal / total = SUM() over the rows above
//       – the Assets = Liab + Equity tie = a cell-reference formula
//     The only literals are the atomic GL line amounts and the prior-month-end
//     opening balances (both sourced straight from the ledger), plus blank blue
//     input cells (e.g. "per bank statement").
//   • Generic: the chart of accounts is read at run time, so this works for any
//     entity. Accounts are grouped into supporting tabs by kind (Cash,
//     Intercompany, Investments, Debt, Equity, Other).
//   • Substantive: intercompany due-to/from balances are tied to the
//     counterparty entity's own GL (mirror check), and any imbalance or missing
//     mirror is flagged on the Summary tab.
//
// Built like the CLRF workpapers (buildData -> buildWorkbook -> saveToWorkpapers)
// and registered by index.js. ctx = { db, auth, requireEntityAccess, requireRole,
// workpapersDir, computeBalances }.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Resolve any date to the quarter end that contains it, plus the prior quarter end.
function resolveQuarter(dateInput) {
  const m = String(dateInput || '').match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const y = Number(m[1]), mo = Number(m[2]);
  if (mo < 1 || mo > 12) throw new Error('quarter_end must be a valid date');
  const q = Math.ceil(mo / 3);        // 1..4 — the quarter containing the picked date
  const qEndMonth = q * 3;            // 3, 6, 9 or 12
  const last = new Date(Date.UTC(y, qEndMonth, 0)).getUTCDate();
  const end = y + '-' + String(qEndMonth).padStart(2, '0') + '-' + String(last).padStart(2, '0');
  // Beginning = prior quarter end (last day of the month before this quarter starts).
  const beg = new Date(Date.UTC(y, qEndMonth - 3, 0)).toISOString().slice(0, 10);
  return {
    end, beg,
    year: String(y), quarter: q, quarterLabel: 'Q' + q,
    label: y + '-Q' + q,
    yearStart: y + '-01-01',
  };
}

// ── GL transaction detail for the month, per account, on the account's natural
// side, with the amount signed the way the balance moves.
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

// ── Account categorization (generic, first match wins). ──────────────────────
function categoryOf(code, name, type, bank) {
  const c = String(code || ''), n = String(name || '').toLowerCase();
  if (type === 'Asset' && (/^10/.test(c) || bank === 1)) return 'cash';
  if (/\bdue\s+(to|from)\b/.test(n) || /intercompan/.test(n)) return 'interco';
  if (/^investment\b/.test(n) || /^19/.test(c)) return 'invest';
  if (type === 'Liability' && (/loan|notes?\s+payable|line of credit|mortgage/.test(n) || /^25/.test(c))) return 'debt';
  if (type === 'Equity') return 'equity';
  return 'other';
}
const CAT_TABS = [
  { key: 'cash', tab: 'Cash', title: 'Cash & Bank Accounts' },
  { key: 'interco', tab: 'Intercompany', title: 'Intercompany (Due To / From)' },
  { key: 'invest', tab: 'Investments', title: 'Investments' },
  { key: 'debt', tab: 'Debt', title: 'Loans & Notes Payable' },
  { key: 'equity', tab: 'Equity', title: 'Members’ Equity' },
  { key: 'other', tab: 'Other BS Accounts', title: 'Other Balance-Sheet Accounts' },
];

// A token to find this entity on a counterparty's ledger ("Odyssey Holdings LLC"
// -> /odyssey/). Uses the first distinctive word of the entity name.
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

// Try to resolve the counterparty NAME embedded in a "Due to/from X" account to
// a CL entity. Returns { id, name } or null.
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
  // Require a reasonably strong match (the longest token, or most of the name).
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
  const lines = glMonth(db, eid, m.beg === m.end ? m.end : addDay(m.beg), m.end); // month activity (exclude beginning day carry)

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
      a.cpNet = mir.net;      // amount cp owes us per their GL
      a.cpLegs = mir.legs;
    }
  }

  // Period P&L (fiscal-YTD) for the equity net-income line.
  const niEnd = pnl(db, eid, m.yearStart, m.end);
  const niBegVal = m.beg < m.yearStart ? 0 : pnl(db, eid, m.yearStart, m.beg).net;

  // Ties / flags (computed in JS for the Summary banner + header).
  const totA = r2(acctRows.filter((a) => a.type === 'Asset').reduce((s, a) => s + a.end, 0));
  const totL = r2(acctRows.filter((a) => a.type === 'Liability').reduce((s, a) => s + a.end, 0));
  const totE = r2(acctRows.filter((a) => a.type === 'Equity').reduce((s, a) => s + a.end, 0));
  const imbalance = r2(totA - (totL + totE + niEnd.net));
  const flags = [];
  if (Math.abs(imbalance) >= 0.01) {
    flags.push({ severity: 'exception', wp: 'Lead Sheet', message: 'Balance sheet does not tie: Assets ' + fmt(totA) + ' ≠ Liabilities ' + fmt(totL) + ' + Equity ' + fmt(totE) + ' + Net income ' + fmt(niEnd.net) + ' (off by ' + fmt(imbalance) + ').' });
  }
  // Per-account roll-forward integrity: begin + activity must equal end.
  for (const a of acctRows) {
    const rollDiff = r2(a.begin + a.activity - a.end);
    if (Math.abs(rollDiff) >= 0.01) {
      flags.push({ severity: 'exception', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': roll-forward does not tie — beginning ' + fmt(a.begin) + ' + activity ' + fmt(a.activity) + ' = ' + fmt(a.begin + a.activity) + ', but GL ending is ' + fmt(a.end) + ' (off by ' + fmt(rollDiff) + ').' });
    }
  }
  // Intercompany mirror ties.
  for (const a of acctRows) {
    if (a.cat !== 'interco') continue;
    if (!a.cp) {
      if (Math.abs(a.end) >= 0.01) flags.push({ severity: 'review', wp: a.code + ' ' + a.name, message: a.code + ' ' + a.name + ': ' + fmt(a.end) + ' outstanding — no matching CL entity found to tie against; confirm the counterparty balance manually.' });
      continue;
    }
    // Our net receivable from this cp on THIS account: asset +, liability -.
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

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const NAVY = 'FF1F3864';
const HDR = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const SMALLI = F({ size: 9, italic: true });
const THIN = { style: 'thin' };
const DBL = { style: 'double' };
const BLUE = F({ color: { argb: 'FF0000FF' } });
const INPUTFILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const spellDate = (end) => { const [y, mo, d] = String(end).split('-').map(Number); return MONTHS[mo - 1] + ' ' + d + ', ' + y; };
const short = (end) => { const [y, mo, d] = String(end).split('-').map(Number); return mo + '/' + d + '/' + String(y).slice(2); };
const qn = (s) => "'" + String(s).replace(/'/g, "''") + "'"; // quoted sheet name for a formula ref

function hdrRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true };
  });
  if (widths) widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
}
function titleBlock(ws, name, subtitle, dateLine) {
  ws.getCell('A1').value = name; ws.getCell('A1').font = F({ size: 12, bold: true });
  ws.getCell('A2').value = subtitle; ws.getCell('A2').font = F({ bold: true });
  ws.getCell('A3').value = dateLine; ws.getCell('A3').font = SMALLI;
}
const setMoney = (ws, ref, v, o = {}) => { const c = ws.getCell(ref); c.value = v; c.numFmt = MONEY; c.font = F(o); return c; };
const setFormula = (ws, ref, formula, o = {}) => { const c = ws.getCell(ref); c.value = { formula }; c.numFmt = MONEY; c.font = F(o); return c; };

function buildWorkbook(data) {
  const { entity, month: m, acctRows, ni } = data;
  const en = entity.name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  // Create Summary and Lead Sheet FIRST so they are the first two tabs; they are
  // populated later once the supporting-tab cell references are known. (ExcelJS
  // tab order follows worksheet creation order.)
  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  const ls = wb.addWorksheet('Lead Sheet', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });

  // ── Build the supporting tabs first, recording each account's Beginning and
  //    Ending cell references so the Lead Sheet and P&L can link to them.
  const ref = new Map(); // code -> { tab, begCell, endCell }
  const byCat = {};
  for (const a of acctRows) { (byCat[a.cat] = byCat[a.cat] || []).push(a); }

  // Group the interco/counterparty ties as we go.
  const built = [];
  for (const cd of CAT_TABS) {
    const rowsForCat = byCat[cd.key] || [];
    if (!rowsForCat.length) continue;
    built.push(cd);
    const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
    titleBlock(ws, en, cd.title, 'Roll-forward for the quarter ended ' + spellDate(m.end));
    const isCash = cd.key === 'cash';
    hdrRow(ws, 5, ['Date', 'Num', 'Payee', 'Description / Memo', 'Amount'], [12, 14, 26, 50, 16]);
    let r = 6;
    for (const a of rowsForCat) {
      // Section header.
      ws.mergeCells('A' + r + ':E' + r);
      const hc = ws.getCell('A' + r); hc.value = a.code + '  —  ' + a.name; hc.font = F({ bold: true });
      hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEFF4' } };
      r++;
      // Beginning balance (sourced from the GL at the prior month end).
      ws.getCell('D' + r).value = 'Beginning balance — ' + short(m.beg) + ' (per GL)'; ws.getCell('D' + r).font = F({ italic: true });
      const begCell = cd.tab + '!$E$' + r; const begLocal = 'E' + r;
      setMoney(ws, begLocal, a.begin);
      r++;
      const firstLine = r;
      for (const ln of a.lines) {
        ws.getCell('A' + r).value = ln.date; ws.getCell('A' + r).font = F();
        ws.getCell('B' + r).value = ln.num; ws.getCell('B' + r).font = F();
        ws.getCell('C' + r).value = ln.payee || ln.location_name || ln.class_name || ''; ws.getCell('C' + r).font = F();
        ws.getCell('D' + r).value = ln.memo; ws.getCell('D' + r).font = F();
        setMoney(ws, 'E' + r, ln.signed);
        r++;
      }
      const lastLine = r - 1;
      // Total activity = SUM of the month's lines (0 if none).
      ws.getCell('D' + r).value = 'Total activity for the quarter'; ws.getCell('D' + r).font = F({ bold: true });
      const actCell = 'E' + r;
      if (lastLine >= firstLine) setFormula(ws, actCell, 'SUM(E' + firstLine + ':E' + lastLine + ')', { bold: true }).border = { top: THIN };
      else { setMoney(ws, actCell, 0, { bold: true }).border = { top: THIN }; }
      r++;
      // Ending balance = Beginning + activity (formula — never typed).
      ws.getCell('D' + r).value = 'Ending balance — ' + short(m.end); ws.getCell('D' + r).font = F({ bold: true });
      const endLocal = 'E' + r; const endCell = cd.tab + '!$E$' + r;
      setFormula(ws, endLocal, begLocal + '+' + actCell, { bold: true }).border = { top: THIN, bottom: DBL };
      r++;
      ref.set(a.code, { tab: cd.tab, begCell, endCell });

      // Cash: bank-rec input + unreconciled difference.
      if (isCash) {
        ws.getCell('D' + r).value = 'Per bank statement (enter)'; ws.getCell('D' + r).font = BLUE;
        const bankCell = 'E' + r; const bc = ws.getCell(bankCell); bc.numFmt = MONEY; bc.font = BLUE; bc.fill = INPUTFILL; bc.border = { top: THIN };
        r++;
        ws.getCell('D' + r).value = 'Unreconciled difference (GL − bank)'; ws.getCell('D' + r).font = F({ italic: true });
        setFormula(ws, 'E' + r, endLocal + '-' + bankCell, { italic: true });
        r++;
      }
      r++; // spacer
    }

    // Intercompany: counterparty mirror reconciliation, grouped by counterparty.
    if (cd.key === 'interco') {
      r++;
      ws.mergeCells('A' + r + ':E' + r);
      const hh = ws.getCell('A' + r); hh.value = 'Counterparty mirror reconciliation (tie to the related entity’s ledger)'; hh.font = F({ bold: true, size: 11 });
      r += 1;
      hdrRow(ws, r, ['Account', 'Counterparty', en + ' shows', 'Counterparty shows', 'Difference'], [16, 30, 18, 18, 16]);
      r++;
      for (const a of rowsForCat) {
        const rr = ref.get(a.code);
        ws.getCell('A' + r).value = a.code; ws.getCell('A' + r).font = F();
        ws.getCell('B' + r).value = a.cp ? a.cp.name : '(no CL entity — confirm manually)'; ws.getCell('B' + r).font = F();
        // Our net receivable on this account: asset +, liability -.
        const ourSign = a.type === 'Asset' ? '' : '-';
        setFormula(ws, 'C' + r, ourSign + qn(rr.tab) + '!$E$' + rr.endCell.split('$E$')[1]);
        if (a.cp) {
          setMoney(ws, 'D' + r, a.cpNet || 0);
          setFormula(ws, 'E' + r, 'C' + r + '-D' + r);
        } else {
          ws.getCell('D' + r).value = '—';
          ws.getCell('E' + r).value = '—';
        }
        r++;
      }
      ws.getCell('B' + (r + 1)).value = 'A receivable on one ledger should be the mirror payable on the other; a non-zero difference is flagged on the Summary tab.'; ws.getCell('B' + (r + 1)).font = SMALLI;
    }
  }

  // ── P&L (fiscal-year-to-date) supporting tab, for the equity net-income line.
  const pl = wb.addWorksheet('Income Statement', { views: [{ showGridLines: false }] });
  titleBlock(pl, en, 'Statement of Operations — fiscal year to date', m.yearStart + ' to ' + short(m.end));
  hdrRow(pl, 5, ['Code', 'Account', 'Amount'], [14, 52, 18]);
  let pr = 6;
  pl.getCell('A' + pr).value = 'REVENUE'; pl.getCell('A' + pr).font = F({ bold: true }); pr++;
  const revFirst = pr;
  for (const x of ni.end.revenue) { pl.getCell('A' + pr).value = x.code; pl.getCell('A' + pr).font = F(); pl.getCell('B' + pr).value = x.name; pl.getCell('B' + pr).font = F(); setMoney(pl, 'C' + pr, x.amt); pr++; }
  const revLast = pr - 1;
  pl.getCell('B' + pr).value = 'Total revenue'; pl.getCell('B' + pr).font = F({ bold: true });
  const totRevCell = 'C' + pr;
  if (revLast >= revFirst) setFormula(pl, totRevCell, 'SUM(C' + revFirst + ':C' + revLast + ')', { bold: true }).border = { top: THIN }; else setMoney(pl, totRevCell, 0, { bold: true });
  pr += 2;
  pl.getCell('A' + pr).value = 'EXPENSES'; pl.getCell('A' + pr).font = F({ bold: true }); pr++;
  const expFirst = pr;
  for (const x of ni.end.expense) { pl.getCell('A' + pr).value = x.code; pl.getCell('A' + pr).font = F(); pl.getCell('B' + pr).value = x.name; pl.getCell('B' + pr).font = F(); setMoney(pl, 'C' + pr, x.amt); pr++; }
  const expLast = pr - 1;
  pl.getCell('B' + pr).value = 'Total expenses'; pl.getCell('B' + pr).font = F({ bold: true });
  const totExpCell = 'C' + pr;
  if (expLast >= expFirst) setFormula(pl, totExpCell, 'SUM(C' + expFirst + ':C' + expLast + ')', { bold: true }).border = { top: THIN }; else setMoney(pl, totExpCell, 0, { bold: true });
  pr += 2;
  pl.getCell('B' + pr).value = 'NET INCOME (LOSS) — fiscal YTD'; pl.getCell('B' + pr).font = F({ bold: true });
  const niEndCell = 'C' + pr;
  setFormula(pl, niEndCell, totRevCell + '-' + totExpCell, { bold: true }).border = { top: THIN, bottom: DBL };
  pr += 2;
  pl.getCell('B' + pr).value = 'Net income (loss) — fiscal YTD through ' + short(m.beg) + ' (beginning of quarter, per GL)'; pl.getCell('B' + pr).font = F({ italic: true });
  const niBegCell = 'C' + pr;
  setMoney(pl, niBegCell, ni.begVal, { italic: true });
  const NI_END_REF = qn('Income Statement') + '!$' + niEndCell.replace(/(\d+)/, '$$$1'); // 'Income Statement'!$C$row
  const NI_BEG_REF = qn('Income Statement') + '!$' + niBegCell.replace(/(\d+)/, '$$$1');

  // ── Lead Sheet (the balance-sheet summary) — LINKS to the supporting tabs.
  ls.getCell('A1').value = en; ls.getCell('A1').font = F({ size: 13, bold: true });
  ls.getCell('A2').value = 'Quarterly Closing Workpaper — Balance Sheet Lead Schedule'; ls.getCell('A2').font = F({ bold: true });
  ls.getCell('A3').value = 'Quarter ended ' + spellDate(m.end); ls.getCell('A3').font = SMALLI;
  ls.getCell('A4').value = 'Every figure below links to its supporting schedule; subtotals and the balance-sheet tie are live formulas.'; ls.getCell('A4').font = SMALLI;
  hdrRow(ls, 6, ['Code', 'Account', 'W/P', 'Beginning ' + short(m.beg), 'Activity', 'Ending ' + short(m.end)], [14, 46, 16, 18, 16, 18]);
  let lr = 7;
  const subtotalCells = { Asset: [], Liability: [], Equity: [] };
  const groupDefs = [['Asset', 'ASSETS'], ['Liability', 'LIABILITIES'], ['Equity', 'MEMBERS’ EQUITY']];
  const totalRefs = {};
  for (const [ty, label] of groupDefs) {
    ls.getCell('A' + lr).value = label; ls.getCell('A' + lr).font = F({ bold: true, size: 11 });
    ls.getCell('A' + lr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEFF4' } };
    lr++;
    const first = lr;
    for (const a of acctRows.filter((x) => x.type === ty)) {
      const rr = ref.get(a.code);
      ls.getCell('A' + lr).value = a.code; ls.getCell('A' + lr).font = F();
      ls.getCell('B' + lr).value = a.name; ls.getCell('B' + lr).font = F();
      ls.getCell('C' + lr).value = rr ? rr.tab : ''; ls.getCell('C' + lr).font = SMALLI;
      if (rr) {
        setFormula(ls, 'D' + lr, qn(rr.tab) + '!' + rr.begCell.split('!')[1]);
        setFormula(ls, 'F' + lr, qn(rr.tab) + '!' + rr.endCell.split('!')[1]);
      } else { setMoney(ls, 'D' + lr, a.begin); setMoney(ls, 'F' + lr, a.end); }
      setFormula(ls, 'E' + lr, 'F' + lr + '-D' + lr);
      lr++;
    }
    // Equity: add the fiscal-YTD net income line, linked to the Income Statement.
    if (ty === 'Equity') {
      ls.getCell('B' + lr).value = 'Net income (loss) — fiscal YTD (per Income Statement)'; ls.getCell('B' + lr).font = F();
      ls.getCell('C' + lr).value = 'Income Statement'; ls.getCell('C' + lr).font = SMALLI;
      setFormula(ls, 'D' + lr, NI_BEG_REF);
      setFormula(ls, 'F' + lr, NI_END_REF);
      setFormula(ls, 'E' + lr, 'F' + lr + '-D' + lr);
      lr++;
    }
    const last = lr - 1;
    ls.getCell('B' + lr).value = 'Total ' + label.replace(/’/g, "'"); ls.getCell('B' + lr).font = F({ bold: true });
    // Guard empty groups: SUM(col{first}:col{first-1}) is a self-including range
    // Excel reads as a circular reference. Write 0 when the group has no rows.
    for (const col of ['D', 'E', 'F']) {
      if (last >= first) setFormula(ls, col + lr, 'SUM(' + col + first + ':' + col + last + ')', { bold: true }).border = { top: THIN, bottom: DBL };
      else setMoney(ls, col + lr, 0, { bold: true }).border = { top: THIN, bottom: DBL };
    }
    totalRefs[ty] = lr;
    lr++; lr++;
  }
  // Balance-sheet tie: Assets = Liabilities + Equity(incl. net income).
  ls.getCell('B' + lr).value = 'Balance-sheet check: Assets − (Liabilities + Equity)'; ls.getCell('B' + lr).font = F({ bold: true });
  setFormula(ls, 'F' + lr, 'F' + totalRefs.Asset + '-(F' + totalRefs.Liability + '+F' + totalRefs.Equity + ')', { bold: true });
  ls.getCell('D' + lr).value = 'should be $0.00'; ls.getCell('D' + lr).font = SMALLI;
  lr += 2;
  ls.getCell('A' + lr).value = 'W/P references point to the supporting schedule tabs, where each ending balance rolls from the beginning balance plus the quarter’s general-ledger activity.'; ls.getCell('A' + lr).font = SMALLI;

  // ── Summary / Exceptions tab (first tab). ────────────────────────────────
  titleBlock(su, en, 'Quarterly Closing Workpaper — Summary', 'Quarter ended ' + spellDate(m.end));
  su.getColumn(1).width = 3; su.getColumn(2).width = 60; su.getColumn(3).width = 20; su.getColumn(4).width = 20;
  let sr = 5;
  const flags = data.flags || [];
  if (!flags.length) {
    su.mergeCells('B' + sr + ':D' + sr);
    const c = su.getCell('B' + sr); c.value = '✓  All checks passed — the balance sheet ties and every account rolls forward from the general ledger.'; c.font = F({ bold: true, color: { argb: 'FF1E7A34' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF7EE' } }; c.alignment = { wrapText: true }; su.getRow(sr).height = 22; sr += 2;
  } else {
    su.mergeCells('B' + sr + ':D' + sr);
    const c = su.getCell('B' + sr); c.value = '⚠  ' + flags.length + ' item' + (flags.length > 1 ? 's' : '') + ' require review'; c.font = F({ bold: true, color: { argb: 'FF9C4221' } });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDECEA' } }; su.getRow(sr).height = 22; sr += 2;
    su.getCell('B' + sr).value = 'Exceptions — Review Required'; su.getCell('B' + sr).font = F({ bold: true, size: 11 }); sr++;
    for (const f of flags) {
      su.mergeCells('B' + sr + ':D' + sr);
      const cc = su.getCell('B' + sr);
      cc.value = (f.severity === 'exception' ? '✖ ' : '⚠ ') + f.message;
      cc.font = F({ color: { argb: f.severity === 'exception' ? 'FFB3261E' : 'FF9C4221' } });
      cc.alignment = { wrapText: true, vertical: 'top' };
      su.getRow(sr).height = Math.min(220, 14 * Math.max(1, Math.ceil(String(f.message).length / 100)) + 4);
      sr++;
    }
    sr++;
  }
  su.getCell('B' + sr).value = 'Balance-sheet tie'; su.getCell('B' + sr).font = F({ bold: true, size: 11 }); sr++;
  const t = data.ties;
  const rowsSummary = [
    ['Total assets', 'F' + totalRefs.Asset],
    ['Total liabilities', 'F' + totalRefs.Liability],
    ['Total members’ equity', 'F' + totalRefs.Equity],
    ['Net income (loss) — fiscal YTD', niEndCell, 'Income Statement'],
  ];
  for (const [lab, cellRef, tab] of rowsSummary) {
    su.getCell('B' + sr).value = lab; su.getCell('B' + sr).font = F();
    setFormula(su, 'C' + sr, qn(tab || 'Lead Sheet') + '!$' + cellRef.replace(/(\d+)/, '$$$1'));
    sr++;
  }
  // The Lead Sheet check row sits 2 rows below the Equity total row (blank + check).
  su.getCell('B' + sr).value = 'Assets − (Liabilities + Equity)  —  should be $0.00'; su.getCell('B' + sr).font = F({ bold: true });
  setFormula(su, 'C' + sr, qn('Lead Sheet') + '!$F$' + (totalRefs.Equity + 2), { bold: true });

  return wb;
}

// ── Persistence. ─────────────────────────────────────────────────────────────
const folderFor = (m) => 'Workpapers/Quarterly Closing/' + m.year + '/' + m.quarterLabel;
const fileNameFor = (m) => 'Quarterly_Closing_Workpaper_' + m.label + '.xlsx';

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

function registerQuarterlyCloseRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/quarterly-close/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const m = resolveQuarter((req.body && req.body.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, m, eid);
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, m, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Quarterclose-Summary', JSON.stringify({
          quarter: m.label, quarter_name: m.quarterLabel + ' ' + m.year,
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

module.exports = { registerQuarterlyCloseRoutes, resolveQuarter, buildData, buildWorkbook };
