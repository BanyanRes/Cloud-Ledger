// ═══════════════════════════════════════════════════════════════════════════
// xlsxBorders — read cell BORDER edges straight from the .xlsx OOXML, because
// the community build of SheetJS does not surface borders in `cell.s`. This is
// the border sibling of xlsxFills.js and works the same way: open the workbook
// as a ZIP, parse the <borders> palette and <cellXfs> from styles.xml, then walk
// the sheet's cells mapping each cell's style index (s=) → cellXfs → borderId →
// which edges (top/bottom/left/right) are drawn.
//
// The point is to reproduce the requisition report's OWN gridlines and
// underlines faithfully — the main Budget-to-Actual table reads as a table
// because Excel drew box borders on those cells, and the reconciliation block
// has underlines only where the workbook actually has bottom borders. We draw
// exactly those edges and infer nothing from row position or labels (position
// inference is what previously produced spurious underlines).
//
// Returns, for a given sheet, a Map keyed "r,c" (0-based, ABSOLUTE sheet
// coordinates) → { top, bottom, left, right } booleans (only edges that are
// present are set true; a cell with no borders is omitted entirely). Border
// COLOR and weight are not reproduced — every present edge is drawn as a thin
// gray hairline, which is all the report needs to read as a table.
// ═══════════════════════════════════════════════════════════════════════════
const JSZip = require('jszip');

// A border edge counts as "present" when it has a style attribute that isn't
// "none". Excel writes e.g. <bottom style="thin"><color .../></bottom> for a
// drawn edge and <bottom/> (or omits it) for none.
function edgePresent(borderXml, edgeName) {
  if (!borderXml) return false;
  // Match <edge .../> or <edge ...>...</edge>; capture its attributes.
  const re = new RegExp('<' + edgeName + '\\b([^>]*?)(?:/>|>)');
  const m = borderXml.match(re);
  if (!m) return false;
  const style = (m[1].match(/style="([^"]*)"/) || [])[1];
  return !!style && style !== 'none';
}

// Build the borders palette (array indexed by borderId) from styles.xml. Each
// entry is { top, bottom, left, right } booleans.
function parseBordersPalette(stylesXml) {
  const block = (stylesXml.match(/<borders[\s\S]*?<\/borders>/) || [])[0] || '';
  const borders = block.match(/<border\b[^>]*?(?:\/>|>[\s\S]*?<\/border>)/g) || [];
  return borders.map(b => ({
    top: edgePresent(b, 'top'),
    bottom: edgePresent(b, 'bottom'),
    left: edgePresent(b, 'left'),
    right: edgePresent(b, 'right'),
  }));
}

// cellXfs: array indexed by style index (the cell's s= attr) → borderId,
// honoring applyBorder. When applyBorder="0" Excel does not apply the border;
// we treat a missing applyBorder as "apply" when a non-zero borderId is present
// (matching how these sheets are authored), and borderId 0 (the default empty
// border) as no border.
function parseCellXfsBorderIds(stylesXml) {
  const block = (stylesXml.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [])[0] || '';
  const xfs = block.match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
  return xfs.map(xf => {
    const borderId = (xf.match(/borderId="(\d+)"/) || [])[1];
    const applyBorder = (xf.match(/applyBorder="(\d)"/) || [])[1];
    if (borderId == null) return null;
    if (applyBorder === '0') return null;
    return Number(borderId);
  });
}

// Resolve the worksheet part path (e.g. "xl/worksheets/sheet3.xml") for a given
// sheet NAME using workbook.xml + its rels. (Same logic as xlsxFills.)
async function resolveSheetPath(zip, sheetName) {
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetTags = wbXml.match(/<sheet\b[^>]*\/>/g) || [];
  let rid = null;
  for (const t of sheetTags) {
    const nm = (t.match(/name="([^"]*)"/) || [])[1];
    if (nm === sheetName) { rid = (t.match(/r:id="([^"]*)"/) || [])[1]; break; }
  }
  if (!rid) return null;
  const relTags = relsXml.match(/<Relationship\b[^>]*\/>/g) || [];
  for (const rt of relTags) {
    const id = (rt.match(/Id="([^"]*)"/) || [])[1];
    if (id === rid) {
      let target = (rt.match(/Target="([^"]*)"/) || [])[1];
      if (!target) return null;
      target = target.replace(/^\//, '');
      if (!target.startsWith('xl/')) target = 'xl/' + target.replace(/^\.\//, '');
      return target;
    }
  }
  return null;
}

// "AB12" → { r: 11, c: 27 } (0-based).
function decodeA1(a1) {
  const m = a1.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: col - 1 };
}

// Main entry. Given the raw xlsx buffer and a sheet name, return a Map
// "r,c" → { top, bottom, left, right } of drawn cell-border edges in absolute
// sheet coordinates. Any parsing failure returns an empty Map (renderer draws no
// borders), so a malformed workbook degrades to the previous no-border behavior
// rather than throwing. Cells whose border has no edges at all are omitted.
async function readSheetBorders(xlsxBuffer, sheetName) {
  const empty = new Map();
  try {
    const zip = await JSZip.loadAsync(xlsxBuffer);
    const stylesFile = zip.file('xl/styles.xml');
    if (!stylesFile) return empty;
    const stylesXml = await stylesFile.async('string');
    const palette = parseBordersPalette(stylesXml);
    const xfBorderIds = parseCellXfsBorderIds(stylesXml);

    const sheetPath = await resolveSheetPath(zip, sheetName);
    if (!sheetPath || !zip.file(sheetPath)) return empty;
    const sheetXml = await zip.file(sheetPath).async('string');

    const out = new Map();
    // Walk cells: <c r="A1" s="12" .../>. Only cells with an s= can carry a border.
    const cellRe = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*?\bs="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
    let m;
    while ((m = cellRe.exec(sheetXml)) !== null) {
      const addr = m[1];
      const sIdx = Number(m[2]);
      const borderId = xfBorderIds[sIdx];
      if (borderId == null) continue;
      const entry = palette[borderId];
      if (!entry) continue;
      if (!entry.top && !entry.bottom && !entry.left && !entry.right) continue;
      const { r, c } = decodeA1(addr);
      out.set(r + ',' + c, { top: entry.top, bottom: entry.bottom, left: entry.left, right: entry.right });
    }
    return out;
  } catch (e) {
    return empty;
  }
}

module.exports = { readSheetBorders };
