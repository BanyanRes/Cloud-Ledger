// ═══════════════════════════════════════════════════════════════════════════
// xlsxToPdf — pure-Node conversion of a single worksheet to a PDF (no
// LibreOffice, Railway-safe). Reads a sheet with SheetJS, renders it with
// pdf-lib as an auto-sized table on landscape Letter pages, paginating by rows.
//
// Used by the Financial Statements engine so an uploaded .xlsx requisition
// report (e.g. the "Budget to Actual" sheet) can be merged into the package
// the same way a PDF requisition report is. The resulting PDF has a real text
// layer, so downstream stripInvoiceLogPages() works unchanged.
// ═══════════════════════════════════════════════════════════════════════════
const XLSX = require('xlsx');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// Landscape Letter.
const LP = { w: 792, h: 612, mL: 30, mR: 30, mT: 40, mB: 34 };
const BODY_FONT = 7;       // pt for body rows
const ROW_PAD = 3;         // vertical padding within a row
const CELL_PAD = 4;        // horizontal padding within a cell
const MIN_COL_W = 14;      // never collapse a column narrower than this

// A cell value looks numeric if, after stripping accounting punctuation, it is
// a number. We right-align those. Percentages and parenthesized negatives count.
function isNumericDisplay(s) {
  if (s == null) return false;
  const t = String(s).trim();
  if (!t) return false;
  if (t === '-' || t === '\u2013') return true; // dash placeholder counts as numeric
  const cleaned = t.replace(/[$,%()\s]/g, '').replace(/^-/, '');
  return cleaned !== '' && !isNaN(Number(cleaned));
}

// Read a worksheet into a dense 2D array of display strings, honoring the
// workbook's number formats (SheetJS `w` = formatted text) and cached values.
// Blank cells become ''. Leading/trailing all-blank rows/cols are trimmed.
function sheetToGrid(ws) {
  const ref = ws['!ref'];
  if (!ref) return { rows: [], nCols: 0 };
  const range = XLSX.utils.decode_range(ref);
  const grid = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      let text = '';
      if (cell) {
        if (cell.w != null) text = String(cell.w);
        else if (cell.v != null) {
          if (cell.t === 'd' && cell.v instanceof Date) {
            const d = cell.v;
            text = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
          } else text = String(cell.v);
        }
      }
      row.push(text);
    }
    grid.push(row);
  }
  let nCols = grid.reduce((m, row) => Math.max(m, row.length), 0);
  const colHasContent = new Array(nCols).fill(false);
  for (const row of grid) for (let c = 0; c < nCols; c++) if ((row[c] || '').trim()) colHasContent[c] = true;
  let lastCol = -1;
  for (let c = 0; c < nCols; c++) if (colHasContent[c]) lastCol = c;
  nCols = lastCol + 1;
  const trimmed = grid.map(row => row.slice(0, nCols));
  while (trimmed.length && trimmed[trimmed.length - 1].every(v => !(v || '').trim())) trimmed.pop();
  let firstCol = nCols;
  for (let c = 0; c < nCols; c++) if (colHasContent[c]) { firstCol = c; break; }
  if (firstCol > 0 && firstCol < nCols) {
    return { rows: trimmed.map(row => row.slice(firstCol)), nCols: nCols - firstCol };
  }
  return { rows: trimmed, nCols };
}

// Compute column widths from content, then scale to fit the printable width.
function computeColWidths(rows, nCols, font, bold, fontSize) {
  const widths = new Array(nCols).fill(MIN_COL_W);
  for (let ri = 0; ri < rows.length; ri++) {
    const row = rows[ri];
    for (let c = 0; c < nCols; c++) {
      const txt = row[c] || '';
      if (!txt) continue;
      const f = ri < 2 ? bold : font; // header-ish rows use bold metrics
      let w;
      try { w = f.widthOfTextAtSize(txt, fontSize); } catch { w = txt.length * fontSize * 0.5; }
      widths[c] = Math.max(widths[c], w + CELL_PAD * 2);
    }
  }
  const printable = LP.w - LP.mL - LP.mR;
  const total = widths.reduce((a, b) => a + b, 0);
  if (total > printable) {
    const scale = printable / total;
    for (let c = 0; c < nCols; c++) widths[c] = Math.max(MIN_COL_W * 0.6, widths[c] * scale);
  }
  return widths;
}

const strW = (font, s, sz) => { try { return font.widthOfTextAtSize(s, sz); } catch { return s.length * sz * 0.5; } };

// Word-wrap a string into lines that each fit maxWidth at fontSize. A single
// word longer than maxWidth is hard-broken by character so nothing overflows.
function wrapText(str, font, fontSize, maxWidth) {
  const s = String(str || '');
  if (!s) return [''];
  if (strW(font, s, fontSize) <= maxWidth) return [s];
  const words = s.split(/\s+/);
  const lines = [];
  let cur = '';
  const pushHardBroken = (word) => {
    let chunk = '';
    for (const ch of word) {
      if (strW(font, chunk + ch, fontSize) > maxWidth && chunk) { lines.push(chunk); chunk = ch; }
      else chunk += ch;
    }
    return chunk;
  };
  for (const w of words) {
    const trial = cur ? cur + ' ' + w : w;
    if (strW(font, trial, fontSize) <= maxWidth) { cur = trial; continue; }
    if (cur) { lines.push(cur); cur = ''; }
    if (strW(font, w, fontSize) > maxWidth) { cur = pushHardBroken(w); }
    else cur = w;
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

// Render one worksheet (SheetJS worksheet object) to PDF bytes.
// opts: { title } — optional heading drawn atop the first page.
async function worksheetToPdfBytes(ws, opts = {}) {
  const { rows, nCols } = sheetToGrid(ws);
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  if (!rows.length || !nCols) {
    const page = pdf.addPage([LP.w, LP.h]);
    page.drawText(opts.title || 'Requisition Report', { x: LP.mL, y: LP.h - LP.mT, size: 12, font: bold });
    page.drawText('(worksheet contained no data)', { x: LP.mL, y: LP.h - LP.mT - 20, size: 9, font: reg, color: rgb(0.4, 0.4, 0.4) });
    return await pdf.save({ useObjectStreams: false });
  }

  const fontSize = nCols > 12 ? BODY_FONT : BODY_FONT + 0.5;
  const widths = computeColWidths(rows, nCols, reg, bold, fontSize);
  const lineH = fontSize + 1.5;                    // leading between wrapped lines
  const colX = [LP.mL];
  for (let c = 0; c < nCols; c++) colX.push(colX[c] + widths[c]);

  // Pre-wrap every cell so we know each row's height (tallest wrapped cell).
  // Numeric cells are never wrapped (they always fit their auto-sized column).
  const wrapped = [];   // wrapped[ri][c] = { lines, numeric, font }
  const rowHeights = [];
  for (let ri = 0; ri < rows.length; ri++) {
    const row = rows[ri];
    const isHeader = ri < 2;
    const firstLabel = (row.find(v => (v || '').trim()) || '').toLowerCase();
    const isTotalRow = /(^total|costs? total|project costs|grand total|balance remaining$)/.test(firstLabel);
    const cellFont = isHeader ? bold : (isTotalRow ? bold : reg);
    const cells = [];
    let maxLines = 1;
    for (let c = 0; c < nCols; c++) {
      const raw = row[c] || '';
      const numeric = !isHeader && isNumericDisplay(raw);
      const lines = (!raw) ? [''] : (numeric ? [raw] : wrapText(raw, cellFont, fontSize, widths[c] - CELL_PAD * 2));
      if (lines.length > maxLines) maxLines = lines.length;
      cells.push({ lines, numeric, font: cellFont });
    }
    wrapped.push({ cells, isHeader, isTotalRow });
    rowHeights.push(maxLines * lineH + ROW_PAD * 2);
  }

  let page = null, y = 0;
  const newPage = (withTitle) => {
    page = pdf.addPage([LP.w, LP.h]);
    y = LP.h - LP.mT;
    if (withTitle && opts.title) {
      page.drawText(opts.title, { x: LP.mL, y, size: 11, font: bold });
      y -= 18;
    }
  };
  newPage(true);

  for (let ri = 0; ri < wrapped.length; ri++) {
    const rowH = rowHeights[ri];
    if (y - rowH < LP.mB) newPage(false);
    const { cells, isTotalRow } = wrapped[ri];
    for (let c = 0; c < nCols; c++) {
      const { lines, numeric, font } = cells[c];
      if (lines.length === 1 && !lines[0]) continue;
      for (let li = 0; li < lines.length; li++) {
        const txt = lines[li];
        if (!txt) continue;
        const tw = strW(font, txt, fontSize);
        const x = numeric ? colX[c] + widths[c] - CELL_PAD - tw : colX[c] + CELL_PAD;
        const ly = y - ROW_PAD - fontSize - li * lineH;
        page.drawText(txt, { x, y: ly, size: fontSize, font, color: rgb(0.1, 0.1, 0.1) });
      }
    }
    if (ri === 1 || isTotalRow) {
      page.drawLine({ start: { x: LP.mL, y: y - rowH + 1 }, end: { x: colX[nCols], y: y - rowH + 1 }, thickness: 0.4, color: rgb(0.55, 0.55, 0.55) });
    }
    y -= rowH;
  }
  return await pdf.save({ useObjectStreams: false });
}

// Convert a specific sheet of an .xlsx buffer to PDF bytes. If sheetName is not
// found, falls back to case-insensitive match, else the first sheet.
// Returns { bytes, sheetUsed, availableSheets }.
async function xlsxSheetToPdf(xlsxBuffer, sheetName, opts = {}) {
  const wb = XLSX.read(xlsxBuffer, { type: 'buffer', cellDates: true, cellNF: true, cellText: true });
  let name = sheetName;
  if (!name || !wb.Sheets[name]) {
    const found = wb.SheetNames.find(n => n.toLowerCase() === String(sheetName || '').toLowerCase());
    name = found || wb.SheetNames[0];
  }
  const ws = wb.Sheets[name];
  const bytes = await worksheetToPdfBytes(ws, { title: opts.title || name });
  return { bytes, sheetUsed: name, availableSheets: wb.SheetNames };
}

// Sniff whether a buffer is a ZIP-based OOXML (.xlsx) file (PK signature).
function looksLikeXlsx(buf, originalName) {
  if (originalName && /\.xlsx?$/i.test(originalName)) return true;
  if (!buf || buf.length < 4) return false;
  return buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 0x03 || buf[2] === 0x05 || buf[2] === 0x07);
}

module.exports = { xlsxSheetToPdf, worksheetToPdfBytes, sheetToGrid, looksLikeXlsx };
