// ═══════════════════════════════════════════════════════════════════════════
// xlsxStyledReport — build a styled .xlsx from the array-of-arrays a client
// report already assembles, plus the live-formula list it already builds.
//
// Why this exists: every client-side Export Excel goes through SheetJS
// (`XLSX.writeFile`), and the community build of SheetJS cannot write cell
// styles at all — no borders, therefore no underlines. CLA (Dennis Arada,
// 8/19/2026) asked for the three underlines an accountant expects on a detail
// report: a single rule under the last amount in each GL account, a rule under
// each subtotal, and a double rule under the grand total. None of them are
// expressible on the SheetJS path.
//
// Rather than port a report's row-building logic to the server, the client keeps
// building `rows` and `formulas` exactly as it does today and additionally says
// WHICH rows are subtotals and WHICH is the grand total. This module turns that
// into a real workbook with ExcelJS (already a server dependency, and the same
// library the Trailing-12-Months export uses).
//
// Contract — all row/column indices are 0-BASED, matching the client's arrays:
//   rows          : any[][]                    the sheet contents
//   formulas      : [{ r, c, f }]              live formulas (f without '=')
//   numFmt        : string                     default '#,##0.00;(#,##0.00)'
//   plainCols     : number[]                   columns never money-formatted
//   sheetName     : string                     default 'Report'
//   style: {
//     titleRows          : number[]   bold 13pt (entity name, report title)
//     metaRows           : number[]   italic 10pt grey (period, project)
//     headerRows         : number[]   bold + bottom rule across amountCols
//     boldRows           : number[]   bold, no rule
//     underlineRows      : number[]   thin bottom border on amountCols
//     doubleUnderlineRows: number[]   double bottom border on amountCols
//     amountCols         : number[]   which columns the rules are drawn under
//     indentRows         : [{ r, n }] left indent for a label column
//     alignCols          : { [col]: 'left'|'center'|'right' }  horizontal align
//                                     applied to every data cell in that column
//   }
//
// A row may appear in both boldRows and underlineRows; both apply. Rules are
// drawn only under `amountCols`, never under the label columns, because that is
// how a printed statement reads: the rule belongs to the number it totals.
// ═══════════════════════════════════════════════════════════════════════════
const ExcelJS = require('exceljs');

const DEFAULT_NUMFMT = '#,##0.00;(#,##0.00)';

// Excel's own column-width unit is roughly "characters", so measure the widest
// rendered value per column the same way the SheetJS path did, and cap it.
function autoWidths(rows) {
  const len = (v) => (typeof v === 'number'
    ? v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).length
    : String(v == null ? '' : v).length);
  const nCols = rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  const out = [];
  for (let c = 0; c < nCols; c++) {
    let w = 8;
    for (const r of rows) if (r && r[c] != null && r[c] !== '') w = Math.max(w, len(r[c]));
    out.push(Math.min(w + 2, 60));
  }
  return out;
}

function buildStyledWorkbookBuffer(spec) {
  const rows = Array.isArray(spec && spec.rows) ? spec.rows : [];
  const st = (spec && spec.style) || {};
  const numFmt = (spec && spec.numFmt) || DEFAULT_NUMFMT;
  const plain = new Set(spec && Array.isArray(spec.plainCols) ? spec.plainCols : []);
  const set = (a) => new Set(Array.isArray(a) ? a : []);
  const titleRows = set(st.titleRows);
  const metaRows = set(st.metaRows);
  const headerRows = set(st.headerRows);
  const boldRows = set(st.boldRows);
  const underlineRows = set(st.underlineRows);
  const doubleRows = set(st.doubleUnderlineRows);
  const amountCols = Array.isArray(st.amountCols) ? st.amountCols : [];
  const indentByRow = new Map((Array.isArray(st.indentRows) ? st.indentRows : []).map(x => [x.r, x.n]));
  const alignCols = (st.alignCols && typeof st.alignCols === 'object') ? st.alignCols : {};

  const wb = new ExcelJS.Workbook();
  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;   // so live formulas show a value on open
  const ws = wb.addWorksheet((spec && spec.sheetName) || 'Report');

  // Values first, then styling, so a style pass can never be undone by a later
  // value write (ExcelJS keeps them on the same cell object).
  rows.forEach((row, r) => {
    if (!row) return;
    row.forEach((v, c) => {
      if (v === '' || v == null) return;
      const cell = ws.getCell(r + 1, c + 1);
      cell.value = v;
      if (typeof v === 'number' && !plain.has(c)) cell.numFmt = numFmt;
    });
  });

  // Live formulas. A cached result is kept when the client sent one, so the
  // sheet reads correctly even before Excel recalculates.
  for (const g of (Array.isArray(spec && spec.formulas) ? spec.formulas : [])) {
    if (!g || !g.f) continue;
    const cell = ws.getCell(g.r + 1, g.c + 1);
    const prev = cell.value;
    const cached = (typeof prev === 'number') ? prev : undefined;
    cell.value = cached === undefined ? { formula: g.f } : { formula: g.f, result: cached };
    if (!plain.has(g.c)) cell.numFmt = numFmt;
  }

  const nCols = rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  const font = (r, c, f) => { const cell = ws.getCell(r + 1, c + 1); cell.font = { ...(cell.font || {}), ...f }; };
  const ruleRow = (r, style) => {
    // Draw under the amount columns only; fall back to the whole row when the
    // caller did not say which columns hold amounts.
    const cols = amountCols.length ? amountCols : Array.from({ length: nCols }, (_, i) => i);
    for (const c of cols) {
      const cell = ws.getCell(r + 1, c + 1);
      cell.border = { ...(cell.border || {}), bottom: { style } };
    }
  };

  rows.forEach((row, r) => {
    if (!row) return;
    const n = row.length || nCols;
    if (titleRows.has(r)) for (let c = 0; c < n; c++) font(r, c, { bold: true, size: 13 });
    if (metaRows.has(r)) for (let c = 0; c < n; c++) font(r, c, { italic: true, size: 10, color: { argb: 'FF444444' } });
    if (headerRows.has(r)) {
      for (let c = 0; c < n; c++) font(r, c, { bold: true });
      const cols = amountCols.length ? amountCols : Array.from({ length: n }, (_, i) => i);
      for (const c of cols) {
        const cell = ws.getCell(r + 1, c + 1);
        cell.border = { ...(cell.border || {}), bottom: { style: 'thin' } };
        cell.alignment = { ...(cell.alignment || {}), horizontal: 'right' };
      }
      // Header labels follow alignCols when given, so a text column's heading
      // sits over its values instead of defaulting to right.
      for (const key of Object.keys(alignCols)) {
        const c = Number(key);
        if (!Number.isInteger(c) || c < 0 || c >= n) continue;
        const cell = ws.getCell(r + 1, c + 1);
        cell.alignment = { ...(cell.alignment || {}), horizontal: alignCols[key] };
      }
    }
    if (boldRows.has(r) || underlineRows.has(r) || doubleRows.has(r)) {
      for (let c = 0; c < n; c++) font(r, c, { bold: true });
    }
    if (underlineRows.has(r)) ruleRow(r, 'thin');
    if (doubleRows.has(r)) ruleRow(r, 'double');
    if (indentByRow.has(r)) {
      const cell = ws.getCell(r + 1, 1);
      cell.alignment = { ...(cell.alignment || {}), indent: indentByRow.get(r) };
    }
    // Per-column horizontal alignment on data rows only. Header rows keep their
    // own right-alignment on amount columns; title/meta rows are left as-is.
    if (!titleRows.has(r) && !metaRows.has(r) && !headerRows.has(r)) {
      for (const key of Object.keys(alignCols)) {
        const c = Number(key);
        if (!Number.isInteger(c) || c < 0 || c >= n) continue;
        const cell = ws.getCell(r + 1, c + 1);
        cell.alignment = { ...(cell.alignment || {}), horizontal: alignCols[key] };
      }
    }
  });

  const widths = Array.isArray(spec && spec.colWidths) && spec.colWidths.length
    ? spec.colWidths : autoWidths(rows);
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // Freeze under the last header row so long details scroll with their headings.
  const lastHeader = Math.max(-1, ...[...headerRows]);
  if (lastHeader >= 0) ws.views = [{ state: 'frozen', ySplit: lastHeader + 1 }];

  return wb.xlsx.writeBuffer();
}

module.exports = { buildStyledWorkbookBuffer, DEFAULT_NUMFMT };
