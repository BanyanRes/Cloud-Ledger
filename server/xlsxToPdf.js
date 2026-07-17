// ═══════════════════════════════════════════════════════════════════════════
// xlsxToPdf — pure-Node conversion of a single worksheet to a PDF (no
// LibreOffice, Railway-safe). Reads a sheet with SheetJS and renders it with
// pdf-lib, reproducing the sheet's own layout as faithfully as a pure-Node
// renderer can: native column widths (!cols), merged cells (!merges), per-cell
// bold and horizontal alignment (cell styles), and the workbook's number
// formats (SheetJS `w` = formatted text). Paginates by rows onto landscape
// Letter pages and scales to fit only when the sheet is wider than the page.
//
// Used by the Financial Statements engine so an uploaded .xlsx requisition
// report's "Budget to Actual" sheet can be merged into the package faithfully
// (we pull the report and convert it, rather than re-formatting it). The
// resulting PDF has a real text layer, so downstream stripInvoiceLogPages()
// works unchanged.
// ═══════════════════════════════════════════════════════════════════════════
const XLSX = require('xlsx');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// Landscape Letter.
const LP = { w: 792, h: 612, mL: 30, mR: 30, mT: 40, mB: 34 };
const BODY_FONT = 7;       // pt for body rows
const ROW_PAD = 3;         // vertical padding within a row
const CELL_PAD = 4;        // horizontal padding within a cell
const MIN_COL_W = 10;      // never collapse a column narrower than this

// Excel column width (in "characters" of the default font) → points. Excel's
// width unit is roughly the width of the '0' glyph; the common conversion is
// px = round(width * 7 + 5), and pt = px * 72/96. We use that so a sheet whose
// columns were sized in Excel keeps its proportions.
function excelWidthToPoints(w) {
  if (w == null) return null;
  const px = Math.round(w * 7 + 5);
  return px * 72 / 96;
}

// A cell value looks numeric if, after stripping accounting punctuation, it is
// a number. We right-align those (unless a style says otherwise). Percentages
// and parenthesized negatives count.
function isNumericDisplay(s) {
  if (s == null) return false;
  const t = String(s).trim();
  if (!t) return false;
  if (t === '-' || t === '\u2013') return true; // dash placeholder counts as numeric
  const cleaned = t.replace(/[$,%()\s]/g, '').replace(/^-/, '');
  return cleaned !== '' && !isNaN(Number(cleaned));
}

// Read a worksheet into a structured grid honoring number formats and cached
// values, plus the sheet's native column widths, merged ranges, and per-cell
// style hints (bold, horizontal alignment). Blank leading/trailing rows/cols
// are trimmed; the same trim offset is applied to widths and merges so
// everything stays aligned.
//   returns { rows, nCols, colWidths, merges, styles }
//     rows[r][c]     = display string
//     colWidths[c]   = points (or null → auto)
//     merges         = [{ r, c, rs, cs }] (top-left row/col + row/col span),
//                      already shifted into the trimmed coordinate space
//     styles[r][c]   = { bold, align } | undefined
function sheetToGrid(ws) {
  const ref = ws['!ref'];
  if (!ref) return { rows: [], nCols: 0, colWidths: [], merges: [], styles: [] };
  const range = XLSX.utils.decode_range(ref);
  const grid = [];
  const styles = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    const styleRow = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      let text = '';
      let st;
      if (cell) {
        if (cell.w != null) text = String(cell.w);
        else if (cell.v != null) {
          if (cell.t === 'd' && cell.v instanceof Date) {
            const d = cell.v;
            text = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
          } else text = String(cell.v);
        }
        // Cell style (present when the workbook was read with cellStyles: true).
        const s = cell.s;
        if (s) {
          const bold = !!(s.font && s.font.bold);
          const align = s.alignment && s.alignment.horizontal; // 'left'|'center'|'right'|undefined
          if (bold || align) st = { bold, align };
        }
      }
      row.push(text);
      styleRow.push(st);
    }
    grid.push(row);
    styles.push(styleRow);
  }

  // Native column widths (points), indexed from range.s.c.
  const colsMeta = ws['!cols'] || [];
  const rawWidths = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const meta = colsMeta[c];
    let pts = null;
    if (meta) {
      if (meta.hidden) pts = 0;
      else if (meta.wpx != null) pts = meta.wpx * 72 / 96;
      else if (meta.wch != null) pts = excelWidthToPoints(meta.wch);
      else if (meta.width != null) pts = excelWidthToPoints(meta.width);
    }
    rawWidths.push(pts);
  }

  // Trim trailing/leading empty columns and trailing empty rows, keeping widths,
  // styles, and merges aligned.
  let nCols = grid.reduce((m, row) => Math.max(m, row.length), 0);
  const colHasContent = new Array(nCols).fill(false);
  for (const row of grid) for (let c = 0; c < nCols; c++) if ((row[c] || '').trim()) colHasContent[c] = true;
  let lastCol = -1;
  for (let c = 0; c < nCols; c++) if (colHasContent[c]) lastCol = c;
  nCols = lastCol + 1;
  let firstCol = nCols;
  for (let c = 0; c < nCols; c++) if (colHasContent[c]) { firstCol = c; break; }
  if (firstCol === nCols) firstCol = 0; // all-empty guard

  const sliceRow = row => row.slice(firstCol, nCols);
  let trimmed = grid.map(sliceRow);
  let trimmedStyles = styles.map(sliceRow);
  // Drop trailing all-blank rows (keep interior blanks — they are layout).
  while (trimmed.length && trimmed[trimmed.length - 1].every(v => !(v || '').trim())) {
    trimmed.pop(); trimmedStyles.pop();
  }
  const outNCols = nCols - firstCol;
  const colWidths = rawWidths.slice(firstCol, nCols);

  // Merged ranges → trimmed coordinate space; drop any fully outside the kept area.
  const merges = [];
  for (const m of (ws['!merges'] || [])) {
    const r0 = m.s.r - range.s.r, c0 = m.s.c - range.s.c - firstCol;
    const r1 = m.e.r - range.s.r, c1 = m.e.c - range.s.c - firstCol;
    if (c1 < 0 || c0 >= outNCols) continue;
    if (r0 >= trimmed.length) continue;
    merges.push({
      r: Math.max(0, r0), c: Math.max(0, c0),
      rs: Math.min(r1, trimmed.length - 1) - Math.max(0, r0) + 1,
      cs: Math.min(c1, outNCols - 1) - Math.max(0, c0) + 1,
    });
  }

  return { rows: trimmed, nCols: outNCols, colWidths, merges, styles: trimmedStyles };
}

// Resolve natural column widths in points: use the sheet's native width where
// it has one; otherwise size from content. NO scaling here — the caller decides
// a single uniform fit-to-page scale from the resulting totals so the whole
// sheet lands on one page (columns AND rows scaled by the same factor).
function resolveColWidths(rows, nCols, colWidths, styles, font, bold, fontSize) {
  const widths = new Array(nCols).fill(0);
  for (let c = 0; c < nCols; c++) {
    if (colWidths[c] != null) { widths[c] = Math.max(colWidths[c] === 0 ? 0 : MIN_COL_W, colWidths[c]); continue; }
    // Content-based fallback for columns Excel didn't size.
    let w = MIN_COL_W;
    for (let ri = 0; ri < rows.length; ri++) {
      const txt = rows[ri][c] || '';
      if (!txt) continue;
      const f = (styles[ri] && styles[ri][c] && styles[ri][c].bold) || ri < 2 ? bold : font;
      let tw;
      try { tw = f.widthOfTextAtSize(txt, fontSize); } catch { tw = txt.length * fontSize * 0.5; }
      w = Math.max(w, tw + CELL_PAD * 2);
    }
    widths[c] = w;
  }
  return widths;
}

const strW = (font, s, sz) => { try { return font.widthOfTextAtSize(s, sz); } catch { return s.length * sz * 0.5; } };

// Word-wrap a string into lines that each fit maxWidth at fontSize. A single
// word longer than maxWidth is hard-broken by character so nothing overflows.
function wrapText(str, font, fontSize, maxWidth) {
  const s = String(str || '');
  if (!s) return [''];
  if (maxWidth <= 0) return [s];
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

// Render one worksheet (SheetJS worksheet object) to PDF bytes, reproducing its
// column widths, merges, bold, and alignment.
//
// The whole sheet is fit onto ONE landscape page (like Excel's "Fit Sheet on
// One Page" print option): a single uniform scale derived from the natural
// content width AND height is applied to column widths, row heights, and font
// size together, so the requisition report's own layout is preserved rather than
// reflowed across pages. opts:
//   title    — a heading line drawn above the sheet. Off by default in
//              single-page mode (the sheet carries its own title block); pass a
//              string only if you want an injected caption.
//   paginate — legacy row-pagination mode (kept for callers that want it); when
//              false (default) the sheet is fit onto a single page.
async function worksheetToPdfBytes(ws, opts = {}) {
  const { rows, nCols, colWidths, merges, styles } = sheetToGrid(ws);
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  if (!rows.length || !nCols) {
    const page = pdf.addPage([LP.w, LP.h]);
    page.drawText(opts.title || 'Requisition Report', { x: LP.mL, y: LP.h - LP.mT, size: 12, font: bold });
    page.drawText('(worksheet contained no data)', { x: LP.mL, y: LP.h - LP.mT - 20, size: 9, font: reg, color: rgb(0.4, 0.4, 0.4) });
    return await pdf.save({ useObjectStreams: false });
  }

  const printableW = LP.w - LP.mL - LP.mR;
  const printableH = LP.h - LP.mT - LP.mB;
  const BASE_FONT = nCols > 12 ? BODY_FONT : BODY_FONT + 0.5;
  const MIN_FONT = 4.2;          // legibility floor; below this we accept overflow
  const titleGap = opts.title ? 16 : 0;

  // Natural (unscaled) column widths and total content width.
  const natWidths = resolveColWidths(rows, nCols, colWidths, styles, reg, bold, BASE_FONT);
  const natTotalW = natWidths.reduce((a, b) => a + b, 0) || 1;

  // Merge lookup (independent of scale) so height estimation honors colspans.
  const covered = new Set();
  const anchorSpan = new Map();
  for (const m of merges) {
    for (let dr = 0; dr < m.rs; dr++) {
      for (let dc = 0; dc < m.cs; dc++) {
        if (dr === 0 && dc === 0) continue;
        covered.add((m.r + dr) + ',' + (m.c + dc));
      }
    }
    anchorSpan.set(m.r + ',' + m.c, m.cs);
  }

  // Given a uniform scale, compute scaled widths/font and the wrapped layout +
  // total height. Wrapping (hence row height) depends on the scale, so this is
  // recomputed as we converge on a scale that fits both width and height.
  const layoutAt = (scale) => {
    const fontSize = Math.max(MIN_FONT, BASE_FONT * scale);
    const lineH = fontSize + 1.5 * scale;
    const rowPad = ROW_PAD * scale;
    const cellPad = CELL_PAD * scale;
    const widths = natWidths.map(w => w * scale);
    const colX = [LP.mL];
    for (let c = 0; c < nCols; c++) colX.push(colX[c] + widths[c]);
    const mergedWidth = (r, c) => {
      const cs = anchorSpan.get(r + ',' + c) || 1;
      let w = 0;
      for (let k = 0; k < cs && (c + k) < nCols; k++) w += widths[c + k];
      return { w, cs };
    };
    const wrapped = [];
    const rowHeights = [];
    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const isHeaderish = ri < 2;
      const firstLabel = (row.find(v => (v || '').trim()) || '').toLowerCase();
      const isTotalRow = /(^total|costs? total|project costs|grand total|balance remaining$)/.test(firstLabel);
      const cells = [];
      let maxLines = 1;
      for (let c = 0; c < nCols; c++) {
        if (covered.has(ri + ',' + c)) { cells.push(null); continue; }
        const raw = row[c] || '';
        const st = styles[ri] && styles[ri][c];
        const cellFont = (st && st.bold) || isHeaderish || isTotalRow ? bold : reg;
        const numeric = isNumericDisplay(raw) && !(st && st.align);
        const align = (st && st.align) || (numeric ? 'right' : 'left');
        const { w: cw } = mergedWidth(ri, c);
        const avail = cw - cellPad * 2;
        const lines = (!raw) ? [''] : (numeric ? [raw] : wrapText(raw, cellFont, fontSize, avail));
        if (lines.length > maxLines) maxLines = lines.length;
        cells.push({ lines, align, font: cellFont });
      }
      wrapped.push({ cells, isTotalRow, headerish: isHeaderish });
      rowHeights.push(maxLines * lineH + rowPad * 2);
    }
    const totalH = rowHeights.reduce((a, b) => a + b, 0);
    return { fontSize, lineH, rowPad, cellPad, widths, colX, mergedWidth, wrapped, rowHeights, totalH };
  };

  // Fit-to-one-page scale. Start from the width constraint (never upscale past
  // 1×), then, if the resulting height still overflows, tighten the scale by the
  // height ratio and recompute once. One extra pass converges because shrinking
  // the font only ever reduces the number of wrapped lines (never increases it),
  // so the second layout's height is a safe upper bound for the final scale.
  let scale = Math.min(1, printableW / natTotalW);
  let L = layoutAt(scale);
  if (L.totalH > printableH) {
    scale = scale * (printableH / L.totalH);
    L = layoutAt(scale);
    // A tiny safety margin in case rounding leaves it a hair over.
    let guard = 0;
    while (L.totalH > printableH && guard++ < 4) {
      scale *= 0.98;
      L = layoutAt(scale);
    }
  }

  const { fontSize, lineH, rowPad, cellPad, colX, mergedWidth, wrapped, rowHeights } = L;

  const page = pdf.addPage([LP.w, LP.h]);
  let y = LP.h - LP.mT;
  if (opts.title) {
    page.drawText(String(opts.title), { x: LP.mL, y, size: Math.max(8, 11 * scale), font: bold });
    y -= titleGap;
  }

  for (let ri = 0; ri < wrapped.length; ri++) {
    const rowH = rowHeights[ri];
    const { cells, isTotalRow } = wrapped[ri];
    for (let c = 0; c < nCols; c++) {
      const cell = cells[c];
      if (!cell) continue; // covered by a merge anchor
      const { lines, align, font } = cell;
      if (lines.length === 1 && !lines[0]) continue;
      const { w: cw } = mergedWidth(ri, c);
      for (let li = 0; li < lines.length; li++) {
        const txt = lines[li];
        if (!txt) continue;
        const tw = strW(font, txt, fontSize);
        let x;
        if (align === 'right') x = colX[c] + cw - cellPad - tw;
        else if (align === 'center') x = colX[c] + (cw - tw) / 2;
        else x = colX[c] + cellPad;
        const ly = y - rowPad - fontSize - li * lineH;
        page.drawText(txt, { x, y: ly, size: fontSize, font, color: rgb(0.1, 0.1, 0.1) });
      }
    }
    // Keep the light rule under the header band and under total rows.
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
  // cellStyles: true so we can read bold/alignment and !cols widths faithfully.
  const wb = XLSX.read(xlsxBuffer, { type: 'buffer', cellDates: true, cellNF: true, cellText: true, cellStyles: true });
  let name = sheetName;
  if (!name || !wb.Sheets[name]) {
    const found = wb.SheetNames.find(n => n.toLowerCase() === String(sheetName || '').toLowerCase());
    name = found || wb.SheetNames[0];
  }
  const ws = wb.Sheets[name];
  // Pass through only an explicitly-provided title; the requisition sheet carries
  // its own header block, so we don't inject the sheet name as a caption.
  const bytes = await worksheetToPdfBytes(ws, { title: opts.title });
  return { bytes, sheetUsed: name, availableSheets: wb.SheetNames };
}

// Sniff whether a buffer is a ZIP-based OOXML (.xlsx) file (PK signature).
function looksLikeXlsx(buf, originalName) {
  if (originalName && /\.xlsx?$/i.test(originalName)) return true;
  if (!buf || buf.length < 4) return false;
  return buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 0x03 || buf[2] === 0x05 || buf[2] === 0x07);
}

module.exports = { xlsxSheetToPdf, worksheetToPdfBytes, sheetToGrid, looksLikeXlsx };
