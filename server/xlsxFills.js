// ═══════════════════════════════════════════════════════════════════════════
// xlsxFills — read solid cell FILL (background) colors straight from the .xlsx
// OOXML, because the community build of SheetJS does not surface fills in
// `cell.s`. We open the workbook as a ZIP (jszip), resolve the sheet's part
// name, then map each cell's style index (s=) → cellXfs → fills palette →
// concrete RGB, resolving theme colors (theme1.xml) and applying tint the same
// way Excel does. Font colors and borders are intentionally ignored — only the
// solid background fill is returned, which is all the requisition report needs
// to reproduce its header band, subtotal rows, and the Date cell.
//
// Returns, for a given sheet, a Map keyed "r,c" (0-based, ABSOLUTE sheet
// coordinates) → "RRGGBB". Cells with no fill, or a non-solid / white / "none"
// pattern, are omitted (so the renderer draws nothing for them). White is
// dropped on purpose: painting white rectangles would just add noise and can
// mask the page background.
// ═══════════════════════════════════════════════════════════════════════════
const JSZip = require('jszip');

// Standard Office theme color order as referenced by <fgColor theme="N"/>.
// NOTE the well-known Excel quirk: in the *theme index space* used by fills,
// slots 0/1 and 2/3 are swapped relative to the clrScheme document order
// (dk1/lt1 → lt1/dk1). We store them already swapped so theme="0" = window
// background (lt1) and theme="1" = window text (dk1), matching Excel.
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

function parseThemeColors(themeXml) {
  const out = {};
  if (!themeXml) return out;
  const scheme = (themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/) || [])[0];
  if (!scheme) return out;
  // clrScheme document order: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
  const names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  for (const nm of names) {
    const block = (scheme.match(new RegExp('<a:' + nm + '>[\\s\\S]*?</a:' + nm + '>')) || [])[0];
    if (!block) continue;
    const srgb = (block.match(/srgbClr val="([0-9A-Fa-f]{6})"/) || [])[1];
    const sys = (block.match(/sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/) || [])[1];
    out[nm] = (srgb || sys || '000000').toUpperCase();
  }
  return out;
}

// Resolve a fill's <fgColor .../> attributes to an "RRGGBB" string (no alpha),
// or null when it can't be resolved. Handles rgb="AARRGGBB", theme="N" (+tint),
// and indexed (a small subset — falls back to null for exotic indices).
function resolveColor(attrs, themeColors) {
  if (!attrs) return null;
  const rgb = (attrs.match(/rgb="([0-9A-Fa-f]{6,8})"/) || [])[1];
  const themeIdx = (attrs.match(/theme="(\d+)"/) || [])[1];
  const tintStr = (attrs.match(/tint="(-?[0-9.]+)"/) || [])[1];
  let hex = null;
  if (rgb) {
    hex = rgb.length === 8 ? rgb.slice(2) : rgb; // strip alpha
  } else if (themeIdx != null) {
    const slot = THEME_SLOTS[Number(themeIdx)];
    hex = slot && themeColors[slot];
  }
  if (!hex) return null;
  hex = hex.toUpperCase();
  const tint = tintStr != null ? parseFloat(tintStr) : 0;
  if (tint) hex = applyTint(hex, tint);
  return hex;
}

// Excel tint: convert to HSL luminance and lighten (tint>0) toward white or
// darken (tint<0) toward black. This matches Excel's documented algorithm
// closely enough for background bands to look right.
function applyTint(hex, tint) {
  let r = parseInt(hex.slice(0, 2), 16);
  let g = parseInt(hex.slice(2, 4), 16);
  let b = parseInt(hex.slice(4, 6), 16);
  // sRGB → HSL
  const rf = r / 255, gf = g / 255, bf = b / 255;
  const max = Math.max(rf, gf, bf), min = Math.min(rf, gf, bf);
  let h = 0, s = 0; const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === rf) h = (gf - bf) / d + (gf < bf ? 6 : 0);
    else if (max === gf) h = (bf - rf) / d + 2;
    else h = (rf - gf) / d + 4;
    h /= 6;
  }
  let lum = l;
  if (tint < 0) lum = lum * (1 + tint);
  else lum = lum * (1 - tint) + tint;
  // HSL → sRGB
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let nr, ng, nb;
  if (s === 0) { nr = ng = nb = lum; }
  else {
    const q = lum < 0.5 ? lum * (1 + s) : lum + s - lum * s;
    const p = 2 * lum - q;
    nr = hue2rgb(p, q, h + 1 / 3);
    ng = hue2rgb(p, q, h);
    nb = hue2rgb(p, q, h - 1 / 3);
  }
  const toHex = v => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0').toUpperCase();
  return toHex(nr) + toHex(ng) + toHex(nb);
}

// Build the fills palette (array indexed by fillId) from styles.xml. Each entry
// is { pattern, color } where color is "RRGGBB" or null. Only solid fills carry
// a color; "none"/"gray125" resolve to null so they're treated as no-fill.
function parseFillsPalette(stylesXml, themeColors) {
  const fillsBlock = (stylesXml.match(/<fills[\s\S]*?<\/fills>/) || [])[0] || '';
  const fills = fillsBlock.match(/<fill>[\s\S]*?<\/fill>/g) || [];
  return fills.map(fl => {
    const pattern = (fl.match(/patternType="([^"]*)"/) || [])[1] || 'none';
    if (pattern !== 'solid') return { pattern, color: null };
    const fg = (fl.match(/<fgColor([^\/>]*)\/?>/) || [])[1] || '';
    return { pattern, color: resolveColor(fg, themeColors) };
  });
}

// cellXfs: array indexed by style index (the cell's s= attr) → fillId, honoring
// applyFill. When applyFill="0" (or absent on a style that inherits), Excel does
// not apply the fill; we treat missing applyFill conservatively as "apply" only
// when a non-default fillId is present, which matches how these sheets are authored.
function parseCellXfsFillIds(stylesXml) {
  const block = (stylesXml.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [])[0] || '';
  const xfs = block.match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
  return xfs.map(xf => {
    const fillId = (xf.match(/fillId="(\d+)"/) || [])[1];
    const applyFill = (xf.match(/applyFill="(\d)"/) || [])[1];
    if (fillId == null) return null;
    const id = Number(fillId);
    if (applyFill === '0') return null;
    return id;
  });
}

// Resolve the worksheet part path (e.g. "xl/worksheets/sheet3.xml") for a given
// sheet NAME using workbook.xml + its rels.
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

// Main entry. Given the raw xlsx buffer and a sheet name, return a Map
// "r,c" → "RRGGBB" of solid, non-white cell fills in absolute sheet coordinates.
// Any parsing failure returns an empty Map (renderer then draws no fills), so a
// malformed workbook degrades to the previous no-color behavior rather than
// throwing.
async function readSheetFills(xlsxBuffer, sheetName) {
  const empty = new Map();
  try {
    const zip = await JSZip.loadAsync(xlsxBuffer);
    const stylesFile = zip.file('xl/styles.xml');
    if (!stylesFile) return empty;
    const stylesXml = await stylesFile.async('string');
    const themeFile = zip.file('xl/theme/theme1.xml');
    const themeXml = themeFile ? await themeFile.async('string') : '';
    const themeColors = parseThemeColors(themeXml);
    const palette = parseFillsPalette(stylesXml, themeColors);
    const xfFillIds = parseCellXfsFillIds(stylesXml);

    const sheetPath = await resolveSheetPath(zip, sheetName);
    if (!sheetPath || !zip.file(sheetPath)) return empty;
    const sheetXml = await zip.file(sheetPath).async('string');

    const out = new Map();
    // Walk cells: <c r="A1" s="12" .../>. Only cells with an s= can carry a fill.
    const cellRe = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*?\bs="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
    let m;
    while ((m = cellRe.exec(sheetXml)) !== null) {
      const addr = m[1];
      const sIdx = Number(m[2]);
      const fillId = xfFillIds[sIdx];
      if (fillId == null) continue;
      const entry = palette[fillId];
      if (!entry || !entry.color) continue;
      const color = entry.color;
      if (color === 'FFFFFF') continue; // skip white — nothing to paint
      const { r, c } = decodeA1(addr);
      out.set(r + ',' + c, color);
    }
    return out;
  } catch (e) {
    return empty;
  }
}

// "AB12" → { r: 11, c: 27 } (0-based).
function decodeA1(a1) {
  const m = a1.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: col - 1 };
}

module.exports = { readSheetFills, _internal: { applyTint, resolveColor, parseThemeColors } };
