// ═══════════════════════════════════════════════════════════════════════════
// xlsxHeader — read a worksheet's PRINT HEADER (the centered title block Excel
// prints at the top of every page) straight from the .xlsx OOXML. SheetJS's
// community build does not surface headers/footers, so this is the header
// sibling of xlsxFills.js / xlsxBorders.js: open the workbook as a ZIP, find the
// sheet part, and pull the <oddHeader> center section.
//
// Requisition workbooks (development AND rail assets) carry their report title
// as a centered print header — entity name / report name / period date — rather
// than as body cells. The Financial-Statements engine renders only the sheet's
// print AREA (which starts below the workbook's own left/right heading cells and
// ends at the reconciliation), so without reading the print header the embedded
// report would have no title at all. This reader returns those center lines so
// the converter can draw a clean centered heading, substituting the package's
// own period date for the (often stale) date line in the file.
//
// Returns an array of trimmed, non-empty text lines for the sheet's odd-page
// header CENTER section (e.g. ["Braker PropCo, LLC", "Budget to Actual",
// "May 31, 2026"]). Any parse failure returns [] so the caller falls back to an
// injected heading rather than throwing.
// ═══════════════════════════════════════════════════════════════════════════
const JSZip = require('jszip');

// Resolve the worksheet part path (e.g. "xl/worksheets/sheet20.xml") for a given
// sheet NAME via workbook.xml + its rels. (Same logic as xlsxBorders.js.)
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

// Unescape the handful of XML entities that appear in header text (JSZip returns
// the raw XML, so "&amp;C" is literal in the string).
function xmlUnescape(s) {
  return String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// Pull the CENTER (&C) section out of an Excel header/footer format string and
// strip its formatting codes, leaving the visible text lines. Excel encodes the
// three regions as &L…&C…&R (any order, any absent); a string with no region
// marker is treated as center. Format codes removed: &"font,style", &<size>,
// and single-letter toggles (&B bold, &I italic, &U underline, &K<color>, etc.).
// Line breaks inside a region are CR/LF.
function centerLinesFromHeader(headerStr) {
  if (!headerStr) return [];
  const s = xmlUnescape(headerStr);
  // Split into regions. Find &L / &C / &R markers (not &&, the literal ampersand).
  const markerRe = /&([LCR])/g;
  const regions = { L: '', C: '', R: '' };
  let m, last = null, lastIdx = 0, sawMarker = false;
  const flush = (endIdx) => { if (last) regions[last] += s.slice(lastIdx, endIdx); };
  while ((m = markerRe.exec(s)) !== null) {
    // Guard against "&&L" (escaped ampersand then letter): require the char
    // before the & not be another & that pairs off. Simple check: if the
    // previous char is '&', skip.
    if (m.index > 0 && s[m.index - 1] === '&') continue;
    sawMarker = true;
    flush(m.index);
    last = m[1];
    lastIdx = markerRe.lastIndex;
  }
  if (sawMarker) flush(s.length);
  let center = sawMarker ? regions.C : s; // no markers → whole string is center
  // Strip formatting codes from the center section.
  center = center
    .replace(/&"[^"]*"/g, '')   // font: &"Calibri,Bold"
    .replace(/&\d+/g, '')        // size: &14
    .replace(/&K[0-9A-Fa-f]{6}/g, '') // color: &KFF0000
    .replace(/&[A-Za-z]/g, '')   // toggles: &B &I &U &S &P &N &D &T &Z &F &A &G &X &Y
    .replace(/&&/g, '&');        // literal ampersand
  return center
    .split(/\r\n|\r|\n/)
    .map(x => x.trim())
    .filter(Boolean);
}

// Main entry. Given the raw xlsx buffer and a sheet name, return the sheet's
// odd-page print-header CENTER lines (array of strings), or [] on any failure.
async function readSheetHeaderCenter(xlsxBuffer, sheetName) {
  try {
    const zip = await JSZip.loadAsync(xlsxBuffer);
    const sheetPath = await resolveSheetPath(zip, sheetName);
    if (!sheetPath || !zip.file(sheetPath)) return [];
    const xml = await zip.file(sheetPath).async('string');
    const block = (xml.match(/<headerFooter[\s\S]*?<\/headerFooter>/) || [])[0] || '';
    if (!block) return [];
    // Prefer the odd (default) header; fall back to the first-page header.
    const odd = (block.match(/<oddHeader>([\s\S]*?)<\/oddHeader>/) || [])[1];
    const first = (block.match(/<firstHeader>([\s\S]*?)<\/firstHeader>/) || [])[1];
    const lines = centerLinesFromHeader(odd);
    if (lines.length) return lines;
    return centerLinesFromHeader(first);
  } catch (e) {
    return [];
  }
}

module.exports = { readSheetHeaderCenter, centerLinesFromHeader };
