// ═══════════════════════════════════════════════════════════════════════════
// xlsxToPdf — pure-Node conversion of a single worksheet to a PDF (no
// LibreOffice, Railway-safe). Reads a sheet with SheetJS and renders it with
// pdf-lib. It prints the sheet's own values as-is — native column widths
// (!cols), merged cells (!merges), each cell's own bold and horizontal
// alignment, and the workbook's number formats (SheetJS `w` = formatted text) —
// and fits the whole sheet onto ONE landscape page.
//
// It reproduces the sheet's OWN cell borders (read from the OOXML, since the
// community SheetJS build doesn't surface them) so the main table reads as a
// table and the reconciliation block's underlines appear exactly where the
// workbook drew bottom borders. It still infers NO formatting from row position
// or labels — bold, fills, and border edges all come straight from the file, so
// nothing spurious is added.
//
// Used by the Financial Statements engine so an uploaded .xlsx requisition
// report's "Budget to Actual" sheet can be merged into the package faithfully
// (we pull the report and convert it, rather than re-formatting it). The
// resulting PDF has a real text layer, so downstream stripInvoiceLogPages()
// works unchanged.
// ═══════════════════════════════════════════════════════════════════════════
const XLSX = require('xlsx');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { readSheetFills } = require('./xlsxFills');
const { readSheetBorders } = require('./xlsxBorders');
const { readSheetHeaderCenter } = require('./xlsxHeader');

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
//   returns { rows, nCols, colWidths, merges, styles, fills, borders }
//     rows[r][c]     = display string
//     colWidths[c]   = points (or null → auto)
//     merges         = [{ r, c, rs, cs }] (top-left row/col + row/col span),
//                      already shifted into the trimmed coordinate space
//     styles[r][c]   = { bold, align } | undefined
//     fills[r][c]    = "RRGGBB" | undefined  (solid cell background, if any)
//     borders[r][c]  = { top, bottom, left, right } | undefined (drawn edges)
// `absFills`/`absBorders` (optional) are Maps in ABSOLUTE sheet coordinates
// (from xlsxFills.readSheetFills / xlsxBorders.readSheetBorders); each is
// shifted into the trimmed space so the renderer can paint backgrounds and draw
// border edges per cell.
function sheetToGrid(ws, absFills, absBorders) {
  const ref = ws['!ref'];
  if (!ref) return { rows: [], nCols: 0, colWidths: [], merges: [], styles: [], fills: [], borders: [] };
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

  // Cell fills → trimmed coordinate space. absFills keys are ABSOLUTE sheet
  // coords "r,c" (0-based from the sheet origin, i.e. including range.s); the
  // grid we build starts at range.s and is then left-trimmed by firstCol and
  // bottom-trimmed by the popped blank rows, so a fill at absolute (R,C) lands
  // at grid row (R - range.s.r) and col (C - range.s.c - firstCol).
  const fills = trimmed.map(row => new Array(row.length));
  if (absFills && absFills.size) {
    for (const [key, color] of absFills) {
      const [ar, ac] = key.split(',').map(Number);
      const gr = ar - range.s.r;
      const gc = ac - range.s.c - firstCol;
      if (gr < 0 || gr >= fills.length) continue;
      if (gc < 0 || gc >= outNCols) continue;
      fills[gr][gc] = color;
    }
  }

  // Cell borders → trimmed coordinate space, exactly like fills. absBorders keys
  // are ABSOLUTE sheet coords "r,c" → { top, bottom, left, right } (from
  // xlsxBorders.readSheetBorders). Each entry lands at grid row (R - range.s.r)
  // and col (C - range.s.c - firstCol), so the drawn edges line up with the
  // same cells the values do.
  const borders = trimmed.map(row => new Array(row.length));
  if (absBorders && absBorders.size) {
    for (const [key, edges] of absBorders) {
      const [ar, ac] = key.split(',').map(Number);
      const gr = ar - range.s.r;
      const gc = ac - range.s.c - firstCol;
      if (gr < 0 || gr >= borders.length) continue;
      if (gc < 0 || gc >= outNCols) continue;
      borders[gr][gc] = edges;
    }
  }

  return { rows: trimmed, nCols: outNCols, colWidths, merges, styles: trimmedStyles, fills, borders };
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
//   fills    — optional Map "r,c"→"RRGGBB" (absolute sheet coords) of solid cell
//              backgrounds to reproduce (from xlsxFills.readSheetFills).
//   borders  — optional Map "r,c"→{top,bottom,left,right} (absolute sheet coords)
//              of drawn cell-border edges to reproduce (from
//              xlsxBorders.readSheetBorders).
async function worksheetToPdfBytes(ws, opts = {}) {
  const { rows, nCols, colWidths, merges, styles, fills, borders } = sheetToGrid(ws, opts.fills, opts.borders);
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  // Centered heading (entity / report name / period date) drawn at the top of
  // EVERY page. opts.heading is an array of { text, size, bold } lines, built by
  // xlsxSheetToPdf from the sheet's own print header with the package period
  // date substituted. When absent, no heading is drawn (legacy behavior).
  const headingLines = (Array.isArray(opts.heading) ? opts.heading : [])
    .filter(h => h && h.text)
    .map(h => ({ text: String(h.text), size: h.size || 11, font: h.bold === false ? reg : bold }));
  const HEADING_GAP = 12;            // gap between the heading block and the table
  const HEADING_TOP_PAD = 2;         // small pad above the first heading line
  const headingH = headingLines.length
    ? headingLines.reduce((a, h) => a + h.size + 4, 0) + HEADING_GAP + HEADING_TOP_PAD
    : 0;
  const drawHeading = (page) => {
    if (!headingH) return;
    let hy = LP.h - LP.mT - HEADING_TOP_PAD;
    for (const h of headingLines) {
      hy -= h.size + 4;
      const tw = strW(h.font, h.text, h.size);
      page.drawText(h.text, { x: (LP.w - tw) / 2, y: hy, size: h.size, font: h.font, color: rgb(0.1, 0.1, 0.1) });
    }
  };
  // Top of the table area on each page (below the heading block, if any).
  const contentTop = LP.h - LP.mT - headingH;

  if (!rows.length || !nCols) {
    const page = pdf.addPage([LP.w, LP.h]);
    drawHeading(page);
    if (!headingH) page.drawText(opts.title || 'Requisition Report', { x: LP.mL, y: LP.h - LP.mT, size: 12, font: bold });
    page.drawText('(worksheet contained no data)', { x: LP.mL, y: contentTop - 14, size: 9, font: reg, color: rgb(0.4, 0.4, 0.4) });
    return await pdf.save({ useObjectStreams: false });
  }

  const printableW = LP.w - LP.mL - LP.mR;
  // Height available for the TABLE on each page = below the centered heading
  // block down to the bottom margin. Reserving headingH here keeps the heading
  // from overlapping the first rows and makes the fit-to-page scale correct.
  const printableH = contentTop - LP.mB;
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
      const cells = [];
      let maxLines = 1;
      for (let c = 0; c < nCols; c++) {
        if (covered.has(ri + ',' + c)) { cells.push(null); continue; }
        const raw = row[c] || '';
        const st = styles[ri] && styles[ri][c];
        // Use ONLY the cell's own reported bold — never infer bold from the
        // row's position or label. Inferring formatting is what produced the
        // spurious underlines/bolding in the reconciliation block; the goal is
        // to print the sheet's values as-is with no added formatting.
        const cellFont = (st && st.bold) ? bold : reg;
        const numeric = isNumericDisplay(raw) && !(st && st.align);
        const align = (st && st.align) || (numeric ? 'right' : 'left');
        const { w: cw } = mergedWidth(ri, c);
        const avail = cw - cellPad * 2;
        const lines = (!raw) ? [''] : (numeric ? [raw] : wrapText(raw, cellFont, fontSize, avail));
        if (lines.length > maxLines) maxLines = lines.length;
        cells.push({ lines, align, font: cellFont });
      }
      wrapped.push({ cells });
      rowHeights.push(maxLines * lineH + rowPad * 2);
    }
    const totalH = rowHeights.reduce((a, b) => a + b, 0);
    return { fontSize, lineH, rowPad, cellPad, widths, colX, mergedWidth, wrapped, rowHeights, totalH };
  };

  // How many pages a given layout actually needs, mirroring the render loop's
  // greedy row placement exactly (a row that would cross the bottom margin moves
  // wholesale to the next page). Using the real page count — not a height ratio —
  // is what lets us reliably avoid orphaning a couple of rows onto a nearly-empty
  // trailing page, since greedy placement leaves whitespace a ratio can't see.
  const pagesNeeded = (rowHeights) => {
    let pages = 1, yy = contentTop;
    for (const rowH of rowHeights) {
      if (yy - rowH < LP.mB && yy < contentTop) { pages++; yy = contentTop; }
      yy -= rowH;
    }
    return pages;
  };

  // Fit scale. Start from the width constraint (never upscale past 1×), then
  // scan scales downward to the legibility floor and pick the one that yields
  // the FEWEST pages (at the largest scale achieving that count). This packs the
  // report tightly — cropping it to its print area removed columns, which raised
  // the width-fit scale and had otherwise spilled the Budget-to-Actual
  // reconciliation's last rows onto an almost-empty extra page.
  const widthScale = Math.min(1, printableW / natTotalW);
  const minScale = Math.min(widthScale, MIN_FONT / BASE_FONT);
  let scale = widthScale;
  let L = layoutAt(widthScale);
  let bestPages = pagesNeeded(L.rowHeights);
  for (let s = widthScale; s >= minScale - 1e-6; s -= 0.02) {
    const cand = layoutAt(s);
    const p = pagesNeeded(cand.rowHeights);
    if (p < bestPages) { bestPages = p; scale = s; L = cand; }  // fewer pages → adopt (largest such scale)
  }

  const { fontSize, lineH, rowPad, cellPad, colX, mergedWidth, wrapped, rowHeights } = L;

  // Every page gets the centered heading drawn at its top; the table starts at
  // contentTop (below the heading).
  const newPage = () => { const pg = pdf.addPage([LP.w, LP.h]); drawHeading(pg); return pg; };
  let page = newPage();
  let y = contentTop;
  if (opts.title && !headingLines.length) {
    page.drawText(String(opts.title), { x: LP.mL, y, size: Math.max(8, 11 * scale), font: bold });
    y -= titleGap;
  }

  for (let ri = 0; ri < wrapped.length; ri++) {
    const rowH = rowHeights[ri];
    // Paginate instead of clipping: when a row would fall below the bottom
    // margin and at least one row is already on this page, spill onto a new
    // page. (A single row taller than the whole page still prints, to avoid an
    // infinite loop.) This is what keeps the Budget-to-Actual reconciliation
    // block, which sits at the bottom of a tall sheet, from being cut off.
    if (y - rowH < LP.mB && y < contentTop) { page = newPage(); y = contentTop; }
    const { cells } = wrapped[ri];
    // Pass 1: paint cell background fills BEFORE any text, so colored bands sit
    // behind their values. A merged region is painted once, across the merged
    // width, using the anchor cell's fill. Cells covered by a merge are skipped
    // (their color, if any, is the anchor's). Fills come straight from the
    // workbook — we don't invent them.
    if (fills && fills.length) {
      for (let c = 0; c < nCols; c++) {
        if (covered.has(ri + ',' + c)) continue;
        const hex = fills[ri] && fills[ri][c];
        if (!hex) continue;
        const { w: cw } = mergedWidth(ri, c);
        if (cw <= 0) continue;
        const col = rgb(
          parseInt(hex.slice(0, 2), 16) / 255,
          parseInt(hex.slice(2, 4), 16) / 255,
          parseInt(hex.slice(4, 6), 16) / 255,
        );
        page.drawRectangle({ x: colX[c], y: y - rowH, width: cw, height: rowH, color: col });
      }
    }
    // Pass 2: draw cell BORDER edges from the workbook (behind text, above
    // fills). Each present edge is a thin gray hairline spanning the cell's
    // merged width; a merged region draws its edges once across the full span.
    // These come straight from the sheet's own borders — the main table's
    // gridlines and the reconciliation block's underlines both appear exactly
    // where Excel drew them, with nothing inferred from row position.
    if (borders && borders.length) {
      const EDGE = rgb(0.55, 0.55, 0.55);
      const EW = Math.max(0.3, 0.5 * scale);
      const yTop = y;
      const yBot = y - rowH;
      for (let c = 0; c < nCols; c++) {
        if (covered.has(ri + ',' + c)) continue;
        const b = borders[ri] && borders[ri][c];
        if (!b) continue;
        const { w: cw } = mergedWidth(ri, c);
        if (cw <= 0) continue;
        const xL = colX[c];
        const xR = colX[c] + cw;
        if (b.top) page.drawLine({ start: { x: xL, y: yTop }, end: { x: xR, y: yTop }, thickness: EW, color: EDGE });
        if (b.bottom) page.drawLine({ start: { x: xL, y: yBot }, end: { x: xR, y: yBot }, thickness: EW, color: EDGE });
        if (b.left) page.drawLine({ start: { x: xL, y: yBot }, end: { x: xL, y: yTop }, thickness: EW, color: EDGE });
        if (b.right) page.drawLine({ start: { x: xR, y: yBot }, end: { x: xR, y: yTop }, thickness: EW, color: EDGE });
      }
    }
    // Pass 3: text on top of any fills and borders.
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
    // We draw ONLY the sheet's own values, fills, and borders — never rules or
    // underlines inferred from a row's position or label. The requisition
    // report's table gridlines and its reconciliation underlines are the
    // workbook's own cell borders (reproduced above), so nothing here is
    // invented.
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
  // Crop to the sheet's PRINT AREA when one is defined (opt-out via
  // opts.printArea === false). This makes the embedded report show exactly the
  // printable page(s): it drops the workbook's own left/right heading cells that
  // sit ABOVE the print area, and anything BELOW it (e.g. the percentage-of-
  // completion summary under the reconciliation). Falls back to the full used
  // range when no print area is defined.
  if (opts.printArea !== false) {
    const pa = resolvePrintArea(wb, name);
    if (pa) ws['!ref'] = pa;
  }
  // Never print the workbook's own heading block (PROJECT FUNDING REQUISITION,
  // Project Address, Project Entity, Application Period, Requisition #) that
  // sits above the column headings — it is redundant with the centered heading
  // we render. Clamp the crop's TOP edge down to the column-heading row even
  // when the print area or a page break reaches up into that block (Jimmy,
  // 2026-08-30).
  if (opts.dropAboveHeader !== false) clampTopToHeader(ws);
  // Read solid cell fills straight from the OOXML (SheetJS community build does
  // not surface fills), so the rendered page reproduces the report's header
  // band, subtotal rows, and Date cell. Degrades to no fills on any parse error.
  let fills;
  try { fills = await readSheetFills(xlsxBuffer, name); } catch { fills = undefined; }
  // Read cell borders the same way, so the rendered page reproduces the report's
  // own table gridlines and the reconciliation block's underlines. Degrades to
  // no borders on any parse error.
  let borders;
  try { borders = await readSheetBorders(xlsxBuffer, name); } catch { borders = undefined; }
  // Build the centered heading (entity / report name / period date) from the
  // sheet's own print header, substituting the caller's package period date for
  // the file's (often stale) date line. Degrades to no heading on any error.
  let heading;
  try { heading = await buildHeading(xlsxBuffer, name, opts); } catch { heading = null; }
  // Pass through only an explicitly-provided title; the requisition sheet's own
  // heading is now rendered from its print header (see `heading`).
  const bytes = await worksheetToPdfBytes(ws, { title: opts.title, fills, borders, heading });
  return { bytes, sheetUsed: name, availableSheets: wb.SheetNames };
}

// Resolve a sheet's PRINT AREA (an A1 range string like "B6:K108") from the
// workbook's _xlnm.Print_Area defined names, which SheetJS surfaces on
// wb.Workbook.Names. Matches by sheet index (localSheetId) or by the sheet name
// embedded in the ref. Multi-area print ranges take the first area. Returns null
// when no usable print area is defined.
function resolvePrintArea(wb, sheetName) {
  try {
    const names = (wb.Workbook && wb.Workbook.Names) || [];
    const idx = wb.SheetNames.indexOf(sheetName);
    for (const n of names) {
      if (!n || n.Name !== '_xlnm.Print_Area') continue;
      const ref = String(n.Ref || '');
      const bang = ref.lastIndexOf('!');
      if (bang < 0) continue;
      const sheetPart = ref.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
      if (!(n.Sheet === idx || sheetPart === sheetName)) continue;
      const a1 = ref.slice(bang + 1).split(',')[0].replace(/\$/g, '').trim();
      if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(a1)) return a1;
      if (/^[A-Z]+\d+$/.test(a1)) return a1 + ':' + a1;
    }
  } catch (e) { /* fall through */ }
  return null;
}

// Move the crop's TOP edge down to the report's column-heading row, dropping
// everything above it — the workbook's own heading block (PROJECT FUNDING
// REQUISITION / Project Address / Project Entity / Application Period /
// Requisition #), which is redundant with the centered heading we draw. This
// runs regardless of the print area or page break, so those rows never print
// even when the print area reaches up into them.
//
// The column-heading row is found by its labels (Account Name / Yardi Code /
// COA#). When the header is drawn on two rows (e.g. "High Point Original" above
// "Approved Budget"), the row ABOVE the label row is kept too, so the top line
// of every column header survives. If no header row is recognized, the crop is
// left unchanged.
function clampTopToHeader(ws) {
  const ref = ws['!ref'];
  if (!ref) return;
  let range;
  try { range = XLSX.utils.decode_range(ref); } catch { return; }
  const ANCHORS = ['account name', 'yardi code', 'coa#', 'coa #'];
  // Tokens that identify a column-header row (used to pull in a stacked header's
  // upper line). Chosen to NOT appear in the heading block above — so
  // "Application Period" and "PROJECT FUNDING REQUISITION" are never mistaken
  // for header rows.
  const HDR_TOKENS = ['contingency', 'reallocation', 'approved budget', 'lender', 'previous application',
    'payment this', 'total complete', 'draw to date', 'inception to', 'incurred to', 'percent drawn', 'balance remaining'];
  const rowText = (r) => {
    let s = '';
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && (cell.w != null || cell.v != null)) s += ' ' + String(cell.w != null ? cell.w : cell.v);
    }
    return s.toLowerCase();
  };
  for (let r = range.s.r; r <= range.e.r; r++) {
    if (!ANCHORS.some(a => rowText(r).includes(a))) continue;
    // Found the label row; walk up while rows still read as header rows (cap 3).
    let top = r;
    for (let up = r - 1, steps = 0; up >= range.s.r && steps < 3; up--, steps++) {
      if (!HDR_TOKENS.some(t => rowText(up).includes(t))) break;
      top = up;
    }
    if (top > range.s.r) { range.s.r = top; ws['!ref'] = XLSX.utils.encode_range(range); }
    return;
  }
}

// Build the centered heading lines for the report: the sheet's own print-header
// center gives the entity name (line 1) and report name (line 2); the date line
// is the caller's package period (opts.headingDate) when provided, else the
// header's own third line. Fallbacks: entity → opts.headingEntity, report → the
// sheet name. Returns an array of { text, size, bold } or null when there is
// nothing to show.
async function buildHeading(xlsxBuffer, name, opts) {
  const lines = await readSheetHeaderCenter(xlsxBuffer, name); // [] on failure
  const entity = (lines[0] || opts.headingEntity || '').trim();
  let report = (lines[1] || name || '').trim();
  // Show the phase number (read from the report's filename) in the heading.
  // Only appended when the sheet's own report line does not already name a
  // phase, so a header that already reads "... Phase 2B" is left untouched
  // (Jimmy, 2026-08-31).
  if (opts.headingPhase && !/\bphase\b/i.test(report)) {
    report = report ? (report + ' — Phase ' + opts.headingPhase) : ('Phase ' + opts.headingPhase);
  }
  const date = (opts.headingDate || lines[2] || '').trim();
  const out = [];
  if (entity) out.push({ text: entity, size: 12, bold: true });
  if (report) out.push({ text: report, size: 11, bold: true });
  if (date) out.push({ text: date, size: 11, bold: true });
  return out.length ? out : null;
}

// Sniff whether a buffer is a ZIP-based OOXML (.xlsx) file (PK signature).
function looksLikeXlsx(buf, originalName) {
  if (originalName && /\.xlsx?$/i.test(originalName)) return true;
  if (!buf || buf.length < 4) return false;
  return buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 0x03 || buf[2] === 0x05 || buf[2] === 0x07);
}

module.exports = { xlsxSheetToPdf, worksheetToPdfBytes, sheetToGrid, looksLikeXlsx, clampTopToHeader };
