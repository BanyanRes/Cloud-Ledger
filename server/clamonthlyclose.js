// ─── Monthly Close — CLA Leadsheets (high-fidelity replica) ───────────────────
//
// A separate monthly workpaper that reproduces CLA's numbered close leadsheets
// account-for-account, in their exact column layout and display, for an entity
// (built for Banyan Residential LLC, entity-agnostic in the engine). One workbook
// with a per-category Leadsheet tab in the CLA style plus the account-appropriate
// supporting schedules:
//
//   • Cash          — Cleared / Register / Bank Statement Ref columns; a per-bank
//                     reconciliation tab (register roll from GL + blue "per bank
//                     statement" input + unreconciled difference).
//   • Receivables   — Balance leadsheet; aging schedule support (blue buckets that
//                     tie to the GL balance).
//   • Prepaids      — Balance leadsheet; amortization/roll support.
//   • Fixed Assets  — Asset-cost / Accumulated-depreciation / Net columns keyed to
//                     one "Fixed Asset Schedule" tab (cost & accum-dep roll-forward
//                     per asset group) + the depreciation JE block.
//   • Other Assets  — Balance Dr (Cr) leadsheet with a per-account detail tab and a
//                     "Grand Total" foot the leadsheet links to; project-code table.
//   • Intercompany  — Entity / Balance / Other Entity Bal / Variance tie-out, each
//                     Due to/from account tied to the counterparty's own GL.
//   • Payables / Credit Cards / Debt / Other Liab / Investments / Equity — the CLA
//                     leadsheet style, GL-derived, with per-account roll support.
//
// Design (per the workpaper convention): no hard-coded derived amounts — every
// leadsheet balance is a formula reference to its supporting tab, every ending is
// Beginning + SUM(GL activity), every total is =SUBTOTAL/SUM, and the
// Assets = Liabilities + Equity + Net income tie is a live formula. The only
// literals are atomic GL line amounts, prior-month opening balances, and blank
// blue input cells for data CloudLedger does not hold (bank/CC statement balances,
// aging buckets). FQ Anchor columns ("#fq-<code>") match CLA for FinQuery linking.
//
// Reuses the GL data layer of ./monthlyclose (buildData, resolveMonth) and is
// registered by index.js. ctx = { db, auth, requireEntityAccess, requireRole,
// workpapersDir, computeBalances }.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { buildData, resolveMonth } = require('./monthlyclose');

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// ─── Display constants (CLA style) ────────────────────────────────────────────
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const DATEFMT = 'mm-dd-yy';
const FONT = 'Calibri';
const YELLOW = 'FFFFFF00';
const RED = 'FFFF0000';
const BLUE = 'FF0000FF';
const LINK = 'FF0563C1';
const HDRFILL = 'FFDDEBF7';   // pale blue table header
const SECTFILL = 'FFBDD7EE';  // section ribbon
const ENTRY_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: YELLOW } };
const INPUT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const F = (o = {}) => Object.assign({ name: FONT, size: 11 }, o);
const THIN = { style: 'thin' };
const DBL = { style: 'double' };
const box = { top: THIN, bottom: THIN, left: THIN, right: THIN };

const short = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return mo + '/' + dd + '/' + String(y).slice(2); };
const spell = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return MONTHS[mo - 1] + ' ' + dd + ', ' + y; };
const asDate = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return new Date(Date.UTC(y, mo - 1, dd)); };
const qn = (s) => "'" + String(s).replace(/'/g, "''") + "'";
const fmt = (n) => '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function sheetNameFor(code, used) {
  let base = String(code).replace(/[\[\]\:\*\?\/\\]/g, '_').slice(0, 31) || 'acct';
  let name = base, i = 1;
  while (used.has(name.toLowerCase())) { const suf = '_' + (++i); name = base.slice(0, 31 - suf.length) + suf; }
  used.add(name.toLowerCase());
  return name;
}

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

// Worksheet hyperlink cell (CLA shows the account number as the link text).
function wpLink(ws, addr, sheet, text) {
  const c = ws.getCell(addr);
  if (sheet) c.value = { formula: 'HYPERLINK("#" & ' + '"' + qn(sheet) + '!A1"' + ',"' + String(text).replace(/"/g, '""') + '")' };
  c.font = F({ bold: true, underline: true, color: { argb: LINK } });
  c.alignment = { horizontal: 'center' };
  return c;
}

// ─── Supporting-tab builders ──────────────────────────────────────────────────

// FQ anchor marker + entity/account heading at the top of a supporting tab.
function tabHead(ws, a, entityName, m, subtitle) {
  ws.getCell('A1').value = '#fq-' + a.code; ws.getCell('A1').font = F({ size: 8, italic: true, color: { argb: 'FFBFBFBF' } });
  ws.getCell('C1').value = entityName; ws.getCell('C1').font = F({ size: 12, bold: true });
  ws.getCell('C2').value = a.code + '  —  ' + a.name; ws.getCell('C2').font = F({ bold: true });
  ws.getCell('C3').value = subtitle; ws.getCell('C3').font = F({ italic: true });
}

// Generic per-account roll-forward tab. Returns { sheet, endRef, endLabelRow }.
// `endLabel` lets Other Assets foot the roll as "Grand Total" (CLA SUMIF target).
function buildRollTab(wb, a, entityName, m, used, opts = {}) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 12; ws.getColumn('B').width = 12; ws.getColumn('C').width = 26;
  ws.getColumn('D').width = 54; ws.getColumn('E').width = 16;
  tabHead(ws, a, entityName, m, 'Roll-forward — month ended ' + spell(m.end));
  const HR = 5;
  hdr(ws, HR, 1, ['Date', 'Num', 'Payee', 'Description / Memo', 'Amount']);
  let r = HR + 1;
  ws.getCell('D' + r).value = 'Beginning balance — ' + short(m.beg) + ' (per GL)'; ws.getCell('D' + r).font = F({ italic: true });
  const begRow = r; const bc = ws.getCell('E' + r); bc.value = a.begin; bc.numFmt = ACCT; bc.font = F(); r++;
  const firstLine = r;
  for (const ln of (a.lines || [])) {
    ws.getCell('A' + r).value = ln.date; ws.getCell('A' + r).font = F();
    ws.getCell('B' + r).value = ln.num || ln.doc_number || ''; ws.getCell('B' + r).font = F();
    ws.getCell('C' + r).value = ln.payee || ln.vendor || ln.location_name || ln.class_name || ''; ws.getCell('C' + r).font = F();
    ws.getCell('D' + r).value = ln.memo || ln.description || ''; ws.getCell('D' + r).font = F();
    const ec = ws.getCell('E' + r); ec.value = ln.signed; ec.numFmt = ACCT; ec.font = F(); r++;
  }
  const lastLine = r - 1;
  ws.getCell('D' + r).value = 'Total activity for the month'; ws.getCell('D' + r).font = F({ bold: true });
  const actRow = r; const ac = ws.getCell('E' + r); ac.numFmt = ACCT; ac.font = F({ bold: true }); ac.border = { top: THIN };
  ac.value = lastLine >= firstLine ? { formula: 'SUM(E' + firstLine + ':E' + lastLine + ')' } : 0; r++;
  const endLabel = opts.endLabel || ('Ending balance — ' + short(m.end));
  ws.getCell('A' + r).value = opts.grandTotal ? 'Grand Total' : ''; ws.getCell('A' + r).font = F({ bold: true });
  ws.getCell('D' + r).value = endLabel; ws.getCell('D' + r).font = F({ bold: true });
  const endRow = r; const enc = ws.getCell('E' + r); enc.numFmt = ACCT; enc.font = F({ bold: true }); enc.border = { top: THIN, bottom: DBL };
  enc.value = { formula: 'E' + begRow + '+E' + actRow };
  return { sheet, endRef: qn(sheet) + '!$E$' + endRow, endRow };
}

// Cash bank-reconciliation tab. Register (GL roll) + blue per-bank-statement
// input + unreconciled difference. Returns { sheet, registerRef, clearedRef }.
function buildCashRecTab(wb, a, entityName, m, used) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 12; ws.getColumn('B').width = 12; ws.getColumn('C').width = 26;
  ws.getColumn('D').width = 54; ws.getColumn('E').width = 16; ws.getColumn('F').width = 16;
  tabHead(ws, a, entityName, m, 'Bank reconciliation — month ended ' + spell(m.end));
  const HR = 5;
  hdr(ws, HR, 1, ['Date', 'Num', 'Payee', 'Description / Memo', 'Amount']);
  let r = HR + 1;
  ws.getCell('D' + r).value = 'Beginning balance — ' + short(m.beg) + ' (per GL)'; ws.getCell('D' + r).font = F({ italic: true });
  const begRow = r; const bc = ws.getCell('E' + r); bc.value = a.begin; bc.numFmt = ACCT; bc.font = F(); r++;
  const firstLine = r;
  for (const ln of (a.lines || [])) {
    ws.getCell('A' + r).value = ln.date; ws.getCell('A' + r).font = F();
    ws.getCell('B' + r).value = ln.num || ln.doc_number || ''; ws.getCell('B' + r).font = F();
    ws.getCell('C' + r).value = ln.payee || ln.vendor || ln.location_name || ''; ws.getCell('C' + r).font = F();
    ws.getCell('D' + r).value = ln.memo || ln.description || ''; ws.getCell('D' + r).font = F();
    const ec = ws.getCell('E' + r); ec.value = ln.signed; ec.numFmt = ACCT; ec.font = F(); r++;
  }
  const lastLine = r - 1;
  ws.getCell('D' + r).value = 'Total activity for the month'; ws.getCell('D' + r).font = F({ bold: true });
  const actRow = r; const ac = ws.getCell('E' + r); ac.numFmt = ACCT; ac.font = F({ bold: true }); ac.border = { top: THIN };
  ac.value = lastLine >= firstLine ? { formula: 'SUM(E' + firstLine + ':E' + lastLine + ')' } : 0; r++;
  ws.getCell('D' + r).value = 'Register balance (per GL) — ' + short(m.end); ws.getCell('D' + r).font = F({ bold: true });
  const regRow = r; const rc = ws.getCell('E' + r); rc.numFmt = ACCT; rc.font = F({ bold: true }); rc.border = { top: THIN, bottom: DBL };
  rc.value = { formula: 'E' + begRow + '+E' + actRow }; r += 2;
  // Blue reconciliation block.
  ws.getCell('C' + r).value = 'Reconciliation to bank statement'; ws.getCell('C' + r).font = F({ bold: true }); r++;
  ws.getCell('D' + r).value = 'Balance per bank statement (enter)'; ws.getCell('D' + r).font = F({ color: { argb: BLUE } });
  const clrRow = r; const cc = ws.getCell('E' + r); cc.numFmt = ACCT; cc.font = F({ color: { argb: BLUE } }); cc.fill = INPUT_FILL; cc.border = box; r++;
  ws.getCell('D' + r).value = 'Less: outstanding checks (enter)'; ws.getCell('D' + r).font = F({ color: { argb: BLUE } });
  const ockRow = r; const oc = ws.getCell('E' + r); oc.numFmt = ACCT; oc.font = F({ color: { argb: BLUE } }); oc.fill = INPUT_FILL; oc.border = box; r++;
  ws.getCell('D' + r).value = 'Add: deposits in transit (enter)'; ws.getCell('D' + r).font = F({ color: { argb: BLUE } });
  const ditRow = r; const dc = ws.getCell('E' + r); dc.numFmt = ACCT; dc.font = F({ color: { argb: BLUE } }); dc.fill = INPUT_FILL; dc.border = box; r++;
  ws.getCell('D' + r).value = 'Cleared / adjusted bank balance'; ws.getCell('D' + r).font = F({ bold: true });
  const adjRow = r; const acj = ws.getCell('E' + r); acj.numFmt = ACCT; acj.font = F({ bold: true }); acj.border = { top: THIN };
  acj.value = { formula: 'E' + clrRow + '-E' + ockRow + '+E' + ditRow }; r++;
  ws.getCell('D' + r).value = 'Unreconciled difference (register − adjusted bank)'; ws.getCell('D' + r).font = F({ italic: true });
  const df = ws.getCell('E' + r); df.numFmt = ACCT; df.font = F({ italic: true }); df.value = { formula: 'E' + regRow + '-E' + adjRow };
  return { sheet, registerRef: qn(sheet) + '!$E$' + regRow, clearedRef: qn(sheet) + '!$E$' + adjRow };
}

// Credit-card reconciliation tab. Register (GL roll, period-end) + blue statement
// (cleared) + blue ending. Returns { sheet, registerRef, clearedRef, endingRef }.
function buildCcRecTab(wb, a, entityName, m, used) {
  const sheet = sheetNameFor(a.code, used);
  const ws = wb.addWorksheet(sheet, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 12; ws.getColumn('B').width = 12; ws.getColumn('C').width = 26;
  ws.getColumn('D').width = 54; ws.getColumn('E').width = 16;
  tabHead(ws, a, entityName, m, 'Credit card reconciliation — month ended ' + spell(m.end));
  const HR = 5;
  hdr(ws, HR, 1, ['Date', 'Num', 'Payee', 'Description / Memo', 'Amount']);
  let r = HR + 1;
  ws.getCell('D' + r).value = 'Beginning balance — ' + short(m.beg) + ' (per GL)'; ws.getCell('D' + r).font = F({ italic: true });
  const begRow = r; const bc = ws.getCell('E' + r); bc.value = a.begin; bc.numFmt = ACCT; bc.font = F(); r++;
  const firstLine = r;
  for (const ln of (a.lines || [])) {
    ws.getCell('A' + r).value = ln.date; ws.getCell('A' + r).font = F();
    ws.getCell('B' + r).value = ln.num || ln.doc_number || ''; ws.getCell('B' + r).font = F();
    ws.getCell('C' + r).value = ln.payee || ln.vendor || ''; ws.getCell('C' + r).font = F();
    ws.getCell('D' + r).value = ln.memo || ln.description || ''; ws.getCell('D' + r).font = F();
    const ec = ws.getCell('E' + r); ec.value = ln.signed; ec.numFmt = ACCT; ec.font = F(); r++;
  }
  const lastLine = r - 1;
  ws.getCell('D' + r).value = 'Total activity for the month'; ws.getCell('D' + r).font = F({ bold: true });
  const actRow = r; const ac = ws.getCell('E' + r); ac.numFmt = ACCT; ac.font = F({ bold: true }); ac.border = { top: THIN };
  ac.value = lastLine >= firstLine ? { formula: 'SUM(E' + firstLine + ':E' + lastLine + ')' } : 0; r++;
  ws.getCell('D' + r).value = 'Register balance (per GL, period-end) — ' + short(m.end); ws.getCell('D' + r).font = F({ bold: true });
  const regRow = r; const rc = ws.getCell('E' + r); rc.numFmt = ACCT; rc.font = F({ bold: true }); rc.border = { top: THIN, bottom: DBL };
  rc.value = { formula: 'E' + begRow + '+E' + actRow }; r += 2;
  ws.getCell('D' + r).value = 'Balance per statement, as of statement date (enter)'; ws.getCell('D' + r).font = F({ color: { argb: BLUE } });
  const clrRow = r; const cc = ws.getCell('E' + r); cc.numFmt = ACCT; cc.font = F({ color: { argb: BLUE } }); cc.fill = INPUT_FILL; cc.border = box; r++;
  ws.getCell('D' + r).value = 'Ending balance, as of period-end (enter, if different from register)'; ws.getCell('D' + r).font = F({ color: { argb: BLUE } });
  const endRow = r; const en2 = ws.getCell('E' + r); en2.numFmt = ACCT; en2.font = F({ color: { argb: BLUE } }); en2.fill = INPUT_FILL; en2.border = box;
  return { sheet, registerRef: qn(sheet) + '!$E$' + regRow, clearedRef: qn(sheet) + '!$E$' + clrRow, endingRef: qn(sheet) + '!$E$' + endRow };
}

// One "Fixed Asset Schedule" tab: cost & accumulated-depreciation roll-forward per
// asset group. Returns Map code -> { costRef, endRow } for the leadsheet links.
function buildFixedSchedule(wb, fixedRows, entityName, m) {
  const ws = wb.addWorksheet('Fixed Asset Schedule', { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 40; ws.getColumn('C').width = 16;
  ws.getColumn('D').width = 16; ws.getColumn('E').width = 16; ws.getColumn('F').width = 40;
  ws.getCell('C1').value = entityName; ws.getCell('C1').font = F({ size: 16, bold: true });
  ws.getCell('C2').value = 'FIXED ASSET SCHEDULE'; ws.getCell('C2').font = F({ size: 12, bold: true });
  ws.getCell('C3').value = 'Month ended ' + spell(m.end); ws.getCell('C3').font = F({ italic: true });
  const HR = 6;
  hdr(ws, HR, 2, ['Account / Asset Group', 'Beginning ' + short(m.beg), 'Activity', 'Ending ' + short(m.end), 'Comments']);
  let r = HR + 1;
  const ref = new Map();
  for (const a of fixedRows) {
    ws.getCell('B' + r).value = a.code + ' — ' + a.name; ws.getCell('B' + r).font = F({ bold: true });
    const bc = ws.getCell('C' + r); bc.value = a.begin; bc.numFmt = ACCT; bc.font = F();
    const ac = ws.getCell('D' + r); ac.value = a.activity; ac.numFmt = ACCT; ac.font = F();
    const ec = ws.getCell('E' + r); ec.numFmt = ACCT; ec.font = F({ bold: true }); ec.value = { formula: 'C' + r + '+D' + r }; ec.border = { top: THIN };
    ref.set(a.code, { costRef: qn('Fixed Asset Schedule') + '!$E$' + r, endRow: r });
    r++;
  }
  return ref;
}

// ─── Leadsheet builders (exact CLA columns) ───────────────────────────────────

// Generic simple leadsheet: Account No. | Account Name | Worksheet | Balance |
// FQ Anchor | Comments. `balanceRef(a)` returns the supporting-tab cell formula.
// Returns the SUBTOTAL tie cell (e.g. "'AP Leadsheet'!$E$15").
function simpleLeadsheet(wb, cd, entityName, m, rows, refByCode) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 40;
  ws.getColumn('D').width = 14; ws.getColumn('E').width = 16; ws.getColumn('F').width = 14; ws.getColumn('G').width = 40;
  titleBlock(ws, entityName, cd.lead, m, 'G');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance', 'FQ Anchor', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(a.code) || {};
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.name; ws.getCell('C' + r).font = F();
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    const bcell = ws.getCell('E' + r); bcell.numFmt = ACCT; bcell.font = F();
    bcell.value = rr.endRef ? { formula: rr.endRef } : a.end;
    const fq = ws.getCell('F' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  const tot = ws.getCell('E' + r); tot.numFmt = ACCT; tot.font = F({ bold: true }); tot.border = { top: THIN, bottom: DBL };
  tot.value = last >= first ? { formula: 'SUBTOTAL(109,E' + first + ':E' + last + ')' } : 0;
  return qn(cd.tab) + '!$E$' + r;
}

// Cash leadsheet: Cleared / Register / Bank Statement Ref. Tie = Register total.
function cashLeadsheet(wb, cd, entityName, m, rows, refByCode) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 38;
  ws.getColumn('D').width = 12; ws.getColumn('E').width = 16; ws.getColumn('F').width = 16;
  ws.getColumn('G').width = 14; ws.getColumn('H').width = 18; ws.getColumn('I').width = 30;
  titleBlock(ws, entityName, cd.lead, m, 'I');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet / Location', 'Cleared Balance', 'Register Balance', 'FQ Anchor', 'Bank Statement Ref', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(a.code) || {};
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.name; ws.getCell('C' + r).font = F();
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    const clr = ws.getCell('E' + r); clr.numFmt = ACCT; clr.font = F(); clr.value = rr.clearedRef ? { formula: rr.clearedRef } : '';
    const reg = ws.getCell('F' + r); reg.numFmt = ACCT; reg.font = F(); reg.value = rr.registerRef ? { formula: rr.registerRef } : a.end;
    const fq = ws.getCell('G' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    ws.getCell('H' + r).value = 'Bank Statement'; ws.getCell('H' + r).font = F();
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  const tclr = ws.getCell('E' + r); tclr.numFmt = ACCT; tclr.font = F({ bold: true }); tclr.border = { top: THIN, bottom: DBL };
  tclr.value = last >= first ? { formula: 'SUBTOTAL(109,E' + first + ':E' + last + ')' } : 0;
  const treg = ws.getCell('F' + r); treg.numFmt = ACCT; treg.font = F({ bold: true }); treg.border = { top: THIN, bottom: DBL };
  treg.value = last >= first ? { formula: 'SUBTOTAL(109,F' + first + ':F' + last + ')' } : 0;
  return qn(cd.tab) + '!$F$' + r; // Register total ties to the balance sheet.
}

// Credit cards leadsheet: Cleared (stmt date) / Register (period-end) / Ending.
// Tie = Register total.
function ccLeadsheet(wb, cd, entityName, m, rows, refByCode) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 34;
  ws.getColumn('D').width = 12; ws.getColumn('E').width = 15; ws.getColumn('F').width = 15; ws.getColumn('G').width = 15;
  ws.getColumn('H').width = 14; ws.getColumn('I').width = 16; ws.getColumn('J').width = 26;
  titleBlock(ws, entityName, cd.lead, m, 'J');
  ws.getCell('E6').value = '(As of Statement Date)'; ws.getCell('E6').font = F({ italic: true, size: 9 });
  ws.getCell('G6').value = '(As of Period-End)'; ws.getCell('G6').font = F({ italic: true, size: 9 });
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet / Location', 'Cleared Balance', 'Register Balance', 'Ending Balance', 'FQ Anchor', 'CC Statement Ref', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(a.code) || {};
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.name; ws.getCell('C' + r).font = F();
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    const clr = ws.getCell('E' + r); clr.numFmt = ACCT; clr.font = F(); clr.value = rr.clearedRef ? { formula: rr.clearedRef } : '';
    const reg = ws.getCell('F' + r); reg.numFmt = ACCT; reg.font = F(); reg.value = rr.registerRef ? { formula: rr.registerRef } : a.end;
    const end = ws.getCell('G' + r); end.numFmt = ACCT; end.font = F(); end.value = rr.endingRef ? { formula: 'IF(' + rr.endingRef + '=0,' + rr.registerRef + ',' + rr.endingRef + ')' } : (rr.registerRef ? { formula: rr.registerRef } : a.end);
    const fq = ws.getCell('H' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  ['E', 'F', 'G'].forEach((col) => {
    const c = ws.getCell(col + r); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL };
    c.value = last >= first ? { formula: 'SUBTOTAL(109,' + col + first + ':' + col + last + ')' } : 0;
  });
  return qn(cd.tab) + '!$F$' + r;
}

// Fixed assets leadsheet: two-section (asset breakdown / depreciation breakdown),
// Net = Cost + Deprec. Tie = Net total.
function fixedLeadsheet(wb, cd, entityName, m, pairs, fixedRef) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 30; ws.getColumn('D').width = 15;
  ws.getColumn('E').width = 12; ws.getColumn('F').width = 12; ws.getColumn('G').width = 30; ws.getColumn('H').width = 15;
  ws.getColumn('I').width = 12; ws.getColumn('J').width = 15; ws.getColumn('K').width = 26;
  titleBlock(ws, entityName, cd.lead, m, 'K');
  ws.getCell('H3').value = 'Capitalization Policy:'; ws.getCell('H3').font = F();
  ws.getCell('I3').value = '>$1,000 for single item'; ws.getCell('I3').font = F();
  // Section ribbon (row 7) + header (row 8).
  const sB = ws.getCell('B7'); sB.value = 'Fixed Asset Account Breakdown'; sB.font = F({ bold: true }); sB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTFILL } };
  const sF = ws.getCell('F7'); sF.value = 'Depreciation Account Breakdown'; sF.font = F({ bold: true }); sF.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SECTFILL } };
  hdr(ws, 8, 2, ['Asset Acct. #', 'Asset Acct. Name', 'Asset Cost', 'Asset FQ Anchor', 'Depreciation Account #', 'Depreciation Account Name', 'Deprec. Balance', 'Deprec. FQ Anchor', 'Net Asset Amount', 'Comments']);
  let r = 9; const first = r;
  for (const p of pairs) {
    const asset = p.asset, dep = p.dep;
    if (asset) {
      ws.getCell('B' + r).value = asset.code; ws.getCell('B' + r).font = F();
      ws.getCell('C' + r).value = asset.name; ws.getCell('C' + r).font = F();
      const dcell = ws.getCell('D' + r); dcell.numFmt = ACCT; dcell.font = F();
      const ar = fixedRef.get(asset.code); dcell.value = ar ? { formula: ar.costRef } : asset.end;
      ws.getCell('E' + r).value = '#fq-' + asset.code; ws.getCell('E' + r).font = F({ bold: true, color: { argb: RED } });
    } else { ws.getCell('D' + r).value = 0; ws.getCell('D' + r).numFmt = ACCT; ws.getCell('D' + r).font = F(); }
    if (dep) {
      ws.getCell('F' + r).value = dep.code; ws.getCell('F' + r).font = F();
      ws.getCell('G' + r).value = dep.name; ws.getCell('G' + r).font = F();
      const hcell = ws.getCell('H' + r); hcell.numFmt = ACCT; hcell.font = F();
      const dr = fixedRef.get(dep.code); hcell.value = dr ? { formula: dr.costRef } : dep.end;
      ws.getCell('I' + r).value = '#fq-' + dep.code; ws.getCell('I' + r).font = F({ bold: true, color: { argb: RED } });
    } else { ws.getCell('H' + r).value = 0; ws.getCell('H' + r).numFmt = ACCT; ws.getCell('H' + r).font = F(); }
    const nc = ws.getCell('J' + r); nc.numFmt = ACCT; nc.font = F(); nc.value = { formula: 'D' + r + '+H' + r };
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  ['D', 'H', 'J'].forEach((col) => {
    const c = ws.getCell(col + r); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL };
    c.value = last >= first ? { formula: 'SUBTOTAL(109,' + col + first + ':' + col + last + ')' } : 0;
  });
  // Depreciation JE hint block.
  const jr = r + 2;
  ws.getCell('C' + jr).value = 'To record depreciation / amortization expense (enter this period):'; ws.getCell('C' + jr).font = F({ italic: true });
  hdr(ws, jr + 1, 3, ['Account', 'Dr', 'Cr']);
  ws.getCell('C' + (jr + 2)).value = 'Depreciation Expense'; ws.getCell('C' + (jr + 2)).font = F();
  const dr1 = ws.getCell('D' + (jr + 2)); dr1.numFmt = ACCT; dr1.fill = INPUT_FILL; dr1.font = F({ color: { argb: BLUE } }); dr1.border = box;
  ws.getCell('C' + (jr + 3)).value = 'Accumulated Depreciation'; ws.getCell('C' + (jr + 3)).font = F();
  const cr1 = ws.getCell('E' + (jr + 3)); cr1.numFmt = ACCT; cr1.fill = INPUT_FILL; cr1.font = F({ color: { argb: BLUE } }); cr1.border = box;
  return qn(cd.tab) + '!$J$' + r; // Net total ties to the balance sheet.
}

// Other assets leadsheet: Balance Dr (Cr) linked to each detail tab "Grand Total".
function otherAssetsLeadsheet(wb, cd, entityName, m, rows, refByCode, projects) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 13; ws.getColumn('C').width = 44;
  ws.getColumn('D').width = 14; ws.getColumn('E').width = 16; ws.getColumn('F').width = 14; ws.getColumn('G').width = 30;
  titleBlock(ws, entityName, cd.lead, m, 'G');
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Account Name', 'Worksheet', 'Balance\nDr (Cr)', 'FQ Anchor', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(a.code) || {};
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.name; ws.getCell('C' + r).font = F();
    wpLink(ws, 'D' + r, rr.sheet, a.code);
    const bcell = ws.getCell('E' + r); bcell.numFmt = ACCT; bcell.font = F();
    bcell.value = rr.sheet ? { formula: 'SUMIF(' + qn(rr.sheet) + '!A:A,"Grand Total",' + qn(rr.sheet) + '!E:E)' } : a.end;
    const fq = ws.getCell('F' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  const tot = ws.getCell('E' + r); tot.numFmt = ACCT; tot.font = F({ bold: true }); tot.border = { top: THIN, bottom: DBL };
  tot.value = last >= first ? { formula: 'SUM(E' + first + ':E' + last + ')' } : 0;
  // Project-code reference table (as CLA carries).
  if (projects && projects.length) {
    let pr = r + 3;
    ws.getCell('C' + pr).value = 'Project Name'; ws.getCell('C' + pr).font = F({ bold: true });
    ws.getCell('D' + pr).value = 'Project Code'; ws.getCell('D' + pr).font = F({ bold: true }); pr++;
    for (const p of projects) { ws.getCell('C' + pr).value = p.name; ws.getCell('C' + pr).font = F(); ws.getCell('D' + pr).value = p.code; ws.getCell('D' + pr).font = F(); pr++; }
  }
  return qn(cd.tab) + '!$E$' + r;
}

// Intercompany leadsheet: Entity / Balance / Other Entity Bal / Variance tie-out.
function intercoLeadsheet(wb, cd, entityName, m, rows, refByCode) {
  const ws = wb.addWorksheet(cd.tab, { views: [{ showGridLines: false }] });
  ws.getColumn('A').width = 3.4; ws.getColumn('B').width = 12; ws.getColumn('C').width = 34; ws.getColumn('D').width = 22;
  ws.getColumn('E').width = 12; ws.getColumn('F').width = 15; ws.getColumn('G').width = 12; ws.getColumn('H').width = 15; ws.getColumn('I').width = 12; ws.getColumn('J').width = 26;
  titleBlock(ws, entityName, cd.lead, m, 'J');
  ws.getCell('C5').value = 'Entity Name:'; ws.getCell('C5').font = F();
  ws.getCell('D5').value = entityName; ws.getCell('D5').font = F({ bold: true });
  const HR = 7;
  hdr(ws, HR, 2, ['Account No.', 'Entity Name', 'Account Name', 'Worksheet', 'Balance', 'FQ Anchor', 'Other Entity Bal', 'Variance', 'Comments']);
  let r = HR + 1; const first = r;
  for (const a of rows) {
    const rr = refByCode.get(a.code) || {};
    ws.getCell('B' + r).value = a.code; ws.getCell('B' + r).font = F(); ws.getCell('B' + r).alignment = { horizontal: 'left' };
    ws.getCell('C' + r).value = a.cp ? a.cp.name : a.name.replace(/^due\s+(to|from)\s+/i, ''); ws.getCell('C' + r).font = F();
    ws.getCell('D' + r).value = a.name; ws.getCell('D' + r).font = F();
    wpLink(ws, 'E' + r, rr.sheet, a.code);
    const bcell = ws.getCell('F' + r); bcell.numFmt = ACCT; bcell.font = F(); bcell.value = rr.endRef ? { formula: rr.endRef } : a.end;
    const fq = ws.getCell('G' + r); fq.value = '#fq-' + a.code; fq.font = F({ bold: true, color: { argb: RED } });
    const oc = ws.getCell('H' + r); oc.numFmt = ACCT; oc.font = F();
    if (a.cp) oc.value = a.cpNet || 0; else { oc.value = ''; }
    const vc = ws.getCell('I' + r); vc.numFmt = ACCT; vc.font = F();
    // Our net receivable (asset +, liability −) plus the counterparty's net.
    const ourSigned = (a.type === 'Asset' ? '' : '-') + 'F' + r;
    if (a.cp) vc.value = { formula: 'IFERROR(' + ourSigned + '+H' + r + ',"")' }; else vc.value = '';
    if (!a.cp && Math.abs(a.end) >= 0.01) { const cm = ws.getCell('J' + r); cm.value = 'No matching CL entity — confirm manually'; cm.font = F({ italic: true, color: { argb: 'FF9C4221' } }); }
    r++;
  }
  const last = r - 1;
  ws.getCell('B' + r).value = cd.total; ws.getCell('B' + r).font = F({ bold: true });
  ['F', 'H', 'I'].forEach((col) => {
    const c = ws.getCell(col + r); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL };
    c.value = last >= first ? { formula: 'SUBTOTAL(109,' + col + first + ':' + col + last + ')' } : 0;
  });
  return qn(cd.tab) + '!$F$' + r;
}

// ─── Category set (CLA order) ─────────────────────────────────────────────────
const CATS = [
  { key: 'cash', tab: 'Cash Leadsheet', lead: 'CASH LEADSHEET', total: 'Total Cash', type: 'Asset' },
  { key: 'ar', tab: 'AR Leadsheet', lead: 'ACCOUNTS RECEIVABLE LEADSHEET', total: 'Total Receivables', type: 'Asset' },
  { key: 'prepaid', tab: 'Prepaid Leadsheet', lead: 'PREPAID EXPENSES LEAD SHEET', total: 'Total Prepaid Expenses', type: 'Asset' },
  { key: 'fixed', tab: 'Fixed Assets Leadsheet', lead: 'FIXED ASSETS LEADSHEET', total: 'Total Fixed Assets', type: 'Asset' },
  { key: 'otherassets', tab: 'Other Assets Leadsheet', lead: 'OTHER ASSETS LEADSHEET', total: 'Total Other Assets', type: 'Asset' },
  { key: 'invest', tab: 'Investments Leadsheet', lead: 'INVESTMENTS LEADSHEET', total: 'Total Investments', type: 'Asset' },
  { key: 'interco', tab: 'Intercompany Leadsheet', lead: 'INTERCOMPANY LEADSHEET', total: 'Total Intercompany', type: 'Asset' },
  { key: 'ap', tab: 'AP Leadsheet', lead: 'ACCOUNTS PAYABLE LEADSHEET', total: 'Total Payables', type: 'Liability' },
  { key: 'cc', tab: 'Credit Cards Leadsheet', lead: 'CREDIT CARDS LEADSHEET', total: 'Total Credit Cards', type: 'Liability' },
  { key: 'debt', tab: 'Debt Leadsheet', lead: 'DEBT LEADSHEET', total: 'Total Debt', type: 'Liability' },
  { key: 'otherliab', tab: 'Other Liabilities Leadsheet', lead: 'OTHER LIABILITIES LEADSHEET', total: 'Total Other Liabilities', type: 'Liability' },
  { key: 'equity', tab: 'Equity Leadsheet', lead: 'EQUITY LEADSHEET', total: 'Total Equity', type: 'Equity' },
];

// Split fixed-asset accounts into asset vs accumulated-depreciation and pair them.
function splitFixed(rows) {
  const isDep = (a) => /accumulated\s+(dep|amort)|acc\.?\s*dep|acc\s+depreciation|amortization/i.test(a.name) || a.end < 0;
  const assets = rows.filter((a) => !isDep(a));
  const deps = rows.filter((a) => isDep(a));
  const sig = (n) => String(n).toLowerCase().replace(/accumulated|depreciation|amortization|acc\.?|dep\.?|expense|:|\-/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter((w) => w.length >= 4);
  const pairs = []; const usedDep = new Set();
  for (const a of assets) {
    const words = sig(a.name); let best = null, bestScore = 0;
    deps.forEach((d, i) => {
      if (usedDep.has(i)) return;
      const dw = sig(d.name); let s = 0; for (const w of words) if (dw.includes(w)) s += w.length;
      if (s > bestScore) { bestScore = s; best = i; }
    });
    const dep = (best != null && bestScore >= 4) ? (usedDep.add(best), deps[best]) : null;
    pairs.push({ asset: a, dep });
  }
  deps.forEach((d, i) => { if (!usedDep.has(i)) pairs.push({ asset: null, dep: d }); });
  return pairs;
}

// ─── Workbook ─────────────────────────────────────────────────────────────────
function buildWorkbook(data) {
  const { entity, month: m, acctRows, ni } = data;
  const en = entity.name;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  const su = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
  const used = new Set(['summary', 'fixed asset schedule', 'income statement', 'tie out']);
  for (const cd of CATS) used.add(cd.tab.toLowerCase());

  const byCat = {};
  for (const a of acctRows) (byCat[a.cat] = byCat[a.cat] || []).push(a);

  const catTie = {}; // key -> tie cell formula string

  // ── Assets ──
  // Cash: bank-rec tabs then leadsheet.
  if ((byCat.cash || []).length) {
    const refByCode = new Map();
    for (const a of byCat.cash) { const rr = buildCashRecTab(wb, a, en, m, used); refByCode.set(a.code, rr); }
    catTie.cash = cashLeadsheet(wb, CATS[0], en, m, byCat.cash, refByCode);
  }
  // AR, Prepaid, Investments, Other Liab, Debt, Equity, AP: generic roll tabs + simple leadsheet.
  const simpleKeys = ['ar', 'prepaid', 'invest', 'ap', 'debt', 'otherliab', 'equity'];
  for (const key of simpleKeys) {
    const rows = byCat[key] || []; if (!rows.length) continue;
    const cd = CATS.find((c) => c.key === key);
    const refByCode = new Map();
    for (const a of rows) { const rr = buildRollTab(wb, a, en, m, used); refByCode.set(a.code, rr); }
    catTie[key] = simpleLeadsheet(wb, cd, en, m, rows, refByCode);
  }
  // Other Assets: detail tabs with a "Grand Total" foot + leadsheet SUMIF.
  if ((byCat.otherassets || []).length) {
    const refByCode = new Map();
    for (const a of byCat.otherassets) { const rr = buildRollTab(wb, a, en, m, used, { grandTotal: true, endLabel: 'Grand Total — ' + short(m.end) }); refByCode.set(a.code, rr); }
    catTie.otherassets = otherAssetsLeadsheet(wb, CATS[4], en, m, byCat.otherassets, refByCode, data.projects || []);
  }
  // Fixed assets: one schedule tab + two-section leadsheet.
  if ((byCat.fixed || []).length) {
    const fixedRef = buildFixedSchedule(wb, byCat.fixed, en, m);
    const pairs = splitFixed(byCat.fixed);
    catTie.fixed = fixedLeadsheet(wb, CATS[3], en, m, pairs, fixedRef);
  }
  // Intercompany: roll tabs + tie-out leadsheet.
  if ((byCat.interco || []).length) {
    const refByCode = new Map();
    for (const a of byCat.interco) { const rr = buildRollTab(wb, a, en, m, used); refByCode.set(a.code, rr); }
    catTie.interco = intercoLeadsheet(wb, CATS[6], en, m, byCat.interco, refByCode);
  }
  // Credit cards: cc-rec tabs + leadsheet.
  if ((byCat.cc || []).length) {
    const refByCode = new Map();
    for (const a of byCat.cc) { const rr = buildCcRecTab(wb, a, en, m, used); refByCode.set(a.code, rr); }
    catTie.cc = ccLeadsheet(wb, CATS[8], en, m, byCat.cc, refByCode);
  }

  // ── Income Statement (fiscal-YTD) supporting tab for the equity NI line. ──
  const niRef = buildIncomeStatement(wb, en, m, ni);

  // ── Summary / balance-sheet tie + exceptions. ──
  buildSummary(su, en, m, data, catTie, niRef);
  return wb;
}

function buildIncomeStatement(wb, en, m, ni) {
  const pl = wb.addWorksheet('Income Statement', { views: [{ showGridLines: false }] });
  pl.getColumn('A').width = 3.4; pl.getColumn('B').width = 14; pl.getColumn('C').width = 52; pl.getColumn('D').width = 16;
  pl.getCell('C1').value = en; pl.getCell('C1').font = F({ size: 16, bold: true });
  pl.getCell('C2').value = 'STATEMENT OF OPERATIONS — FISCAL YEAR TO DATE'; pl.getCell('C2').font = F({ size: 12, bold: true });
  pl.getCell('C3').value = m.yearStart + ' to ' + short(m.end); pl.getCell('C3').font = F({ italic: true });
  const PHR = 5; hdr(pl, PHR, 2, ['Code', 'Account', 'Amount']);
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
  pl.getCell('C' + pr).value = 'NET INCOME (LOSS) — fiscal YTD'; pl.getCell('C' + pr).font = F({ bold: true });
  const niEndCell = 'D' + pr; { const c = pl.getCell(niEndCell); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN, bottom: DBL }; c.value = { formula: totRevCell + '-' + totExpCell }; }
  return qn('Income Statement') + '!$D$' + pr.toString();
}

function buildSummary(su, en, m, data, catTie, niRef) {
  su.getColumn('A').width = 3.4; su.getColumn('B').width = 62; su.getColumn('C').width = 20; su.getColumn('D').width = 20;
  su.getCell('C1').value = en; su.getCell('C1').font = F({ size: 16, bold: true });
  su.getCell('C2').value = 'MONTHLY CLOSE (CLA LEADSHEETS) — SUMMARY'; su.getCell('C2').font = F({ size: 12, bold: true });
  su.getCell('C3').value = 'Month ended ' + spell(m.end); su.getCell('C3').font = F({ italic: true });
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
    su.getCell('B' + sr).value = 'Exceptions — Review Required'; su.getCell('B' + sr).font = F({ bold: true, size: 12 }); sr++;
    for (const f of flags) {
      su.mergeCells('B' + sr + ':D' + sr);
      const cc = su.getCell('B' + sr);
      cc.value = (f.severity === 'exception' ? '✖ ' : '⚠ ') + f.message;
      cc.font = F({ color: { argb: f.severity === 'exception' ? 'FFB3261E' : 'FF9C4221' } });
      cc.alignment = { wrapText: true, vertical: 'top' };
      su.getRow(sr).height = Math.min(220, 14 * Math.max(1, Math.ceil(String(f.message).length / 100)) + 4); sr++;
    }
    sr++;
  }
  su.getCell('B' + sr).value = 'Balance-sheet tie'; su.getCell('B' + sr).font = F({ bold: true, size: 12 }); sr++;
  const assetRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Asset').map((cd) => catTie[cd.key]);
  const liabRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Liability').map((cd) => catTie[cd.key]);
  const eqRefs = CATS.filter((cd) => catTie[cd.key] && cd.type === 'Equity').map((cd) => catTie[cd.key]);
  const sumOf = (refs) => refs.length ? refs.join('+') : '0';
  const aRow = sr; su.getCell('B' + sr).value = 'Total assets'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(assetRefs) }; } sr++;
  const lRow = sr; su.getCell('B' + sr).value = 'Total liabilities'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(liabRefs) }; } sr++;
  const eRow = sr; su.getCell('B' + sr).value = 'Total members’ equity'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: sumOf(eqRefs) }; } sr++;
  const nRow = sr; su.getCell('B' + sr).value = 'Net income (loss) — fiscal YTD (per Income Statement)'; su.getCell('B' + sr).font = F();
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F(); c.value = { formula: niRef }; } sr++;
  su.getCell('B' + sr).value = 'Assets − (Liabilities + Equity + Net income)  —  should be $0.00'; su.getCell('B' + sr).font = F({ bold: true });
  { const c = su.getCell('C' + sr); c.numFmt = ACCT; c.font = F({ bold: true }); c.border = { top: THIN }; c.value = { formula: 'C' + aRow + '-(C' + lRow + '+C' + eRow + '+C' + nRow + ')' }; }
}

// ─── Persistence ──────────────────────────────────────────────────────────────
const folderFor = (m) => 'Workpapers/Monthly Close - CLA/' + m.year;
const fileNameFor = (m) => 'Monthly_Close_CLA_' + m.label + '.xlsx';

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
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/cla-monthly-close/:entity_id/generate', auth, requireEntityAccess('entity_id'),
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
}

module.exports = { registerClaMonthlyCloseRoutes, buildWorkbook };
