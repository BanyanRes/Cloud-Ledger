// ═══════════════════════════════════════════════════════════════════════════
// je_xlsx — render a single journal entry to a styled .xlsx with LIVE formulas.
//
// The debit and credit column totals are real =SUM() formulas over the line
// rows, and a "Difference" check is =SUM(debits)-SUM(credits) (0.00 on a
// balanced entry) so the workbook stays analyzable and a reviewer can add or
// change a line and watch the totals move. Amounts are written as real numbers
// with the same accounting number format the financial-statement export uses.
//
// Input `je` is the object the entry-detail route returns:
//   { entity: {name, code}, entry_num, date, memo, doc_number, vendor,
//     lines: [{ account_code, account_name, debit, credit, description,
//               project_name, project_code, location_name, class_name }] }
// ═══════════════════════════════════════════════════════════════════════════
const ExcelJS = require('exceljs');

const MONEY_FMT = '#,##0.00;(#,##0.00);"-"';
const num = (v) => (v == null || v === '' ? 0 : Number(v));

function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Does this entry carry any dimension tag on any line? Only then is the
// Dimension column drawn, so a plain two-line JE stays narrow.
function hasDims(lines) {
  return (lines || []).some(l => l.project_name || l.project_code || l.location_name || l.class_name);
}
function dimText(l) {
  if (l.project_name || l.project_code) return 'Project — ' + (l.project_code && l.project_code !== l.project_name ? l.project_code + ' — ' + (l.project_name || '') : (l.project_name || l.project_code));
  if (l.location_name) return 'Location — ' + l.location_name;
  if (l.class_name) return 'Class — ' + l.class_name;
  return '';
}

async function buildEntryWorkbook(je) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true; // so the SUM formulas show a value on open

  const jeNo = 'JE-' + String(je.entry_num == null ? '' : je.entry_num).padStart(4, '0');
  const ws = wb.addWorksheet(jeNo);
  ws.properties.defaultRowHeight = 15;

  const lines = je.lines || [];
  const showDims = hasDims(lines);
  // Columns: Account | [Dimension] | Description | Debit | Credit
  const COL = { acct: 1 };
  let c = 2;
  if (showDims) { COL.dim = c++; }
  COL.desc = c++; COL.debit = c++; COL.credit = c++;
  const lastCol = COL.credit;

  // ── Header block ──────────────────────────────────────────────────────────
  let r = 1;
  const put = (row, col, val, opts = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = val;
    if (opts.bold || opts.size) cell.font = { bold: !!opts.bold, size: opts.size || 11 };
    if (opts.align) cell.alignment = { horizontal: opts.align };
    if (opts.numFmt) cell.numFmt = opts.numFmt;
    return cell;
  };
  const mergeAcross = (row) => { try { ws.mergeCells(row, 1, row, lastCol); } catch (_) {} };

  put(r, 1, (je.entity && je.entity.name) || 'Journal Entry', { bold: true, size: 14, align: 'center' }); mergeAcross(r); r++;
  put(r, 1, 'Journal Entry ' + jeNo, { bold: true, size: 12, align: 'center' }); mergeAcross(r); r++;
  r++; // spacer

  const meta = [['Date', je.date || ''], ['Doc / Invoice #', je.doc_number || ''], ['Vendor', je.vendor || ''], ['Memo', je.memo || '']];
  for (const [label, val] of meta) {
    if (val === '' && (label === 'Doc / Invoice #' || label === 'Vendor')) continue; // hide empty optionals
    put(r, 1, label, { bold: true });
    ws.getCell(r, 2).value = val;
    try { ws.mergeCells(r, 2, r, lastCol); } catch (_) {}
    r++;
  }
  r++; // spacer

  // ── Column headers ────────────────────────────────────────────────────────
  const headerRow = r;
  put(r, COL.acct, 'Account', { bold: true });
  if (showDims) put(r, COL.dim, 'Dimension', { bold: true });
  put(r, COL.desc, 'Description', { bold: true });
  put(r, COL.debit, 'Debit', { bold: true, align: 'right' });
  put(r, COL.credit, 'Credit', { bold: true, align: 'right' });
  for (let cc = 1; cc <= lastCol; cc++) ws.getCell(r, cc).border = { bottom: { style: 'thin' } };
  r++;

  // ── Line rows ─────────────────────────────────────────────────────────────
  const firstLineRow = r;
  for (const l of lines) {
    const acctLabel = (l.account_code || '') + (l.account_name ? ' — ' + l.account_name : '');
    put(r, COL.acct, acctLabel);
    if (showDims) put(r, COL.dim, dimText(l));
    put(r, COL.desc, l.description || '');
    const d = num(l.debit), cr = num(l.credit);
    const dc = ws.getCell(r, COL.debit); dc.numFmt = MONEY_FMT; dc.alignment = { horizontal: 'right' };
    if (d) dc.value = d;
    const cc2 = ws.getCell(r, COL.credit); cc2.numFmt = MONEY_FMT; cc2.alignment = { horizontal: 'right' };
    if (cr) cc2.value = cr;
    r++;
  }
  const lastLineRow = r - 1;

  // ── Totals row: live SUM formulas ────────────────────────────────────────
  const totalRow = r;
  put(r, COL.acct, 'TOTAL', { bold: true, align: 'right' });
  if (COL.desc !== COL.acct) put(r, COL.desc, '', { bold: true });
  // Merge the label across everything left of the Debit column so "TOTAL" sits
  // right against the figures.
  try { ws.mergeCells(totalRow, 1, totalRow, COL.debit - 1); } catch (_) {}
  ws.getCell(totalRow, 1).value = 'TOTAL';
  ws.getCell(totalRow, 1).font = { bold: true };
  ws.getCell(totalRow, 1).alignment = { horizontal: 'right' };

  const dLetter = colLetter(COL.debit), cLetter = colLetter(COL.credit);
  const sumDr = lines.reduce((s, l) => s + num(l.debit), 0);
  const sumCr = lines.reduce((s, l) => s + num(l.credit), 0);
  const dCell = ws.getCell(totalRow, COL.debit);
  const cCell = ws.getCell(totalRow, COL.credit);
  if (lastLineRow >= firstLineRow) {
    dCell.value = { formula: 'SUM(' + dLetter + firstLineRow + ':' + dLetter + lastLineRow + ')', result: sumDr };
    cCell.value = { formula: 'SUM(' + cLetter + firstLineRow + ':' + cLetter + lastLineRow + ')', result: sumCr };
  } else {
    dCell.value = sumDr; cCell.value = sumCr;
  }
  for (const cell of [dCell, cCell]) { cell.numFmt = MONEY_FMT; cell.font = { bold: true }; cell.alignment = { horizontal: 'right' }; cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } }; }
  r++;

  // ── Difference check (debits − credits): 0.00 when balanced ──────────────
  const diffRow = r;
  try { ws.mergeCells(diffRow, 1, diffRow, COL.debit - 1); } catch (_) {}
  ws.getCell(diffRow, 1).value = 'Difference (Debits − Credits)';
  ws.getCell(diffRow, 1).font = { italic: true };
  ws.getCell(diffRow, 1).alignment = { horizontal: 'right' };
  const diffCell = ws.getCell(diffRow, COL.debit);
  diffCell.value = { formula: dLetter + totalRow + '-' + cLetter + totalRow, result: sumDr - sumCr };
  diffCell.numFmt = MONEY_FMT; diffCell.alignment = { horizontal: 'right' }; diffCell.font = { italic: true };
  r++;

  // ── Column widths ─────────────────────────────────────────────────────────
  ws.getColumn(COL.acct).width = 42;
  if (showDims) ws.getColumn(COL.dim).width = 26;
  ws.getColumn(COL.desc).width = 34;
  ws.getColumn(COL.debit).width = 16;
  ws.getColumn(COL.credit).width = 16;

  return wb.xlsx.writeBuffer();
}

module.exports = { buildEntryWorkbook, MONEY_FMT };
