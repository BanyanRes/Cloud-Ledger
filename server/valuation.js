// --- CLRF Valuation Summary workpaper -----------------------------------------
//
// Quarterly workpaper for County Line Rail Fund I, LP (entity 40). Rather than
// rebuild the appraiser's 34-tab valuation workbook from scratch, this generator
// takes the PRIOR quarter's valuation file (saved under the entity's Workpapers/
// Valuation/<Qn YYYY>/ folder) as a template and injects the GL-derived figures
// for the target quarter, leaving every appraisal input carried forward. It then
// saves the result into a new Workpapers/Valuation/<target quarter>/ folder so the
// next quarter's run can find it.
//
// GL figures injected (all as of the target quarter-end):
//   1. Trial balance -> the appraiser's SOI tab is replaced with the FULL CLRF
//        trial balance (code | name | balance) and renamed "TB". Summary's Book
//        Carrying Value column links directly to the four investment accounts'
//        TB rows: D12=TB!C<121011 CLIP>, D13=TB!C<121041 Silsbee>,
//        D14=TB!C<121021 Buna>, D15=TB!C<121031 SRN>, D16 sums them. Rows are
//        resolved dynamically from where each account lands in the sorted TB.
//   2. CLIP development cost -> a self-contained "CLIP GL Dev Costs" tab listing
//        16 GL accounts from CLIP Property Owner (entity 54), footed to a total,
//        which Summary!H12 references.
//   3. Sales-comparison stabilization discount (Sales Approach CLIP!I68) is
//        re-solved as the plug so that Summary!J12 (CLIP total valuation) holds
//        at the appraiser's carried-forward conclusion while H12 rises to the GL
//        dev cost. This mirrors the Q1 2026 methodology exactly.
//   4. Summary!C8 valuation date -> the target quarter-end.
//
// The workbook is a zip full of formula cells whose cached <v> values exceljs/
// openpyxl would strip on a normal re-save (~1,000 external-link cells reference
// absent '[93]'/'[97]'/'[99]' workbooks and hold data only in their cache). So we
// patch the raw sheet XML with JSZip, rewriting only the cells we change and
// leaving every other byte -- and cache -- intact. No LibreOffice recalc.

const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');

const CLRF_ENTITY_ID = 40;
const CLIP_ENTITY_ID = 54;

// Appraiser's concluded CLIP total valuation (Summary!J12). Held constant across
// quarters while H12 (dev cost) is refreshed and the stabilization discount plugs
// the difference. Used as the anchor when the template's own J12 cache is not a
// clean number (in the distributed workbook it is #VALUE!, because H12 pulls a
// broken requisition link).
const CLIP_CONCLUDED_VALUATION = 149431786.26;

// The four book-carrying-value (investment) accounts on CLRF. The TB tab now
// carries the WHOLE CLRF trial balance; these four accounts are the ones the
// Summary schedule pulls into its Book Carrying Value column (D) via direct
// =TB!C<row> references. Summary rows: D12=CLIP, D13=Silsbee, D14=Buna, D15=SRN.
const BCV = {
  '121031': { label: 'SRN', summaryRow: 15 },
  '121011': { label: 'CLIP', summaryRow: 12 },
  '121041': { label: 'Silsbee', summaryRow: 13 },
  '121021': { label: 'Buna', summaryRow: 14 },
};

// CLIP development-cost account set (mirrors server/devcosts.js). Two groups.
const LONG_TERM_INVESTMENTS = [
  ['11010', 'Acquisitions Costs (Land Purchase)'],
  ['11040', 'Land Contract Payments'],
  ['11050', 'Other Land Costs'],
  ['11211', 'Future Expansion Project Costs'],
  ['11230', 'Other Construction Costs'],
];
const OTHER_ASSETS = [
  ['11970', 'Other Legal - Legal'],
  ['12013', 'Civil Engineering Plans'],
  ['12115', 'A&E'],
  ['12230', 'Professional Services - Accounting'],
  ['12315', 'Appraisal'],
  ['12321', 'Construction Period Interest'],
  ['12343', 'Loan Fees'],
  ['12381', 'Acquisition Fee'],
  ['12596', 'Closing Costs'],
  ['12720', 'Travel - Other Development Costs'],
  ['12913', 'Development Fee'],
];

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// -- Quarter helpers ----------------------------------------------------------
const QENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
function resolveQuarter(quarterEnd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(quarterEnd || '')) {
    throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  }
  const [y, m, d] = quarterEnd.split('-').map(Number);
  if (!QENDS[m] || d !== QENDS[m]) {
    throw new Error('quarter_end must be a quarter end (03-31, 06-30, 09-30, 12-31). Received ' + quarterEnd);
  }
  const q = m / 3;
  return { label: 'Q' + q + ' ' + y, year: String(y), quarter: 'Q' + q, q: q, end: quarterEnd };
}
const valFolder = (qtr) => 'Workpapers/Valuation/' + qtr.quarter + ' ' + qtr.year;
const valFileName = (qtr) => 'CLRF ' + qtr.quarter + ' ' + qtr.year + ' Valuation_dist.xlsx';

// Excel 1900 serial (1900 leap bug: base 1899-12-30).
function excelSerial(ymd) {
  const [y, m, d] = ymd.split('-').map(Number);
  const base = Date.UTC(1899, 11, 30);
  return Math.round((Date.UTC(y, m - 1, d) - base) / 86400000);
}

// -- Locate the prior-quarter valuation file to use as the template -----------
function priorQuarterOf(qtr) {
  const y = Number(qtr.year);
  return qtr.q === 1 ? { quarter: 'Q4', year: String(y - 1) }
    : { quarter: 'Q' + (qtr.q - 1), year: qtr.year };
}
function findTemplate(ctx, eid, qtr) {
  const { db, workpapersDir } = ctx;
  const prior = priorQuarterOf(qtr);
  const priorFolder = 'Workpapers/Valuation/' + prior.quarter + ' ' + prior.year;
  let row = db.prepare(
    'SELECT * FROM entity_files WHERE entity_id=? AND folder_path=? '
    + "AND original_name LIKE '%.xlsx' ORDER BY id DESC LIMIT 1"
  ).get(eid, priorFolder);
  if (!row) {
    row = db.prepare(
      "SELECT * FROM entity_files WHERE entity_id=? AND folder_path LIKE 'Workpapers/Valuation/%' "
      + "AND original_name LIKE '%.xlsx' ORDER BY id DESC LIMIT 1"
    ).get(eid);
  }
  if (!row) return null;
  return Object.assign({}, row, {
    source_folder: row.folder_path,
    abs_path: path.join(workpapersDir, String(eid), row.stored_filename),
  });
}

// -- Save generated workbook into the target quarter's folder -----------------
function saveToWorkpapers(ctx, eid, qtr, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = valFolder(qtr);
  const original = valFileName(qtr);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id=? AND folder_path=? '
    + 'AND original_name=?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id=?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_'
    + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, '
    + "uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

// -- XML cell helpers ---------------------------------------------------------
function replaceCell(xml, ref, newCell) {
  const re = new RegExp('<c r="' + ref + '"(?:[^>]*)(?:/>|>[\\s\\S]*?</c>)');
  const m = xml.match(re);
  if (!m) throw new Error('cell ' + ref + ' not found in sheet XML');
  if ((m[0].match(/<c r=/g) || []).length !== 1) throw new Error('overran cell ' + ref);
  return xml.slice(0, m.index) + newCell + xml.slice(m.index + m[0].length);
}
function styleOf(xml, ref) {
  const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*\\bs="(\\d+)"'));
  return m ? m[1] : null;
}
function numFromCell(xml, ref) {
  const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*>(?:<f[^>]*>[\\s\\S]*?</f>|<f[^>]*/>)?<v>([^<]+)</v>'));
  if (!m) throw new Error('no cached value for ' + ref);
  return r2(Number(m[1]));
}
// Like numFromCell but never throws: returns null when the cell is missing, has
// no cached value, or the cache is an error string (e.g. "#VALUE!") that does not
// parse to a finite number.
function numFromCellSafe(xml, ref) {
  const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*>(?:<f[^>]*>[\\s\\S]*?</f>|<f[^>]*/>)?<v>([^<]+)</v>'));
  if (!m) return null;
  const n = Number(m[1]);
  return isFinite(n) ? n : null;
}
const numCell = (ref, s, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><v>' + v + '</v></c>';
const fCell = (ref, s, f, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><f>' + xmlEsc(f) + '</f><v>' + v + '</v></c>';

// -- Build the "CLIP GL Dev Costs" worksheet XML ------------------------------
function buildDevSheetXml(dev) {
  const sC = (ref, txt) => '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(txt) + '</t></is></c>';
  const nC = (ref, v) => '<c r="' + ref + '"><v>' + v + '</v></c>';
  const fC = (ref, f, v) => '<c r="' + ref + '"><f>' + xmlEsc(f) + '</f><v>' + v + '</v></c>';
  const rows = [];
  rows.push([1, [sC('A1', 'County Line Industrial Park LLC (CLIP Property Owner)')]]);
  rows.push([2, [sC('A2', 'Development Costs per General Ledger')]]);
  rows.push([3, [sC('A3', 'As of ' + dev.as_of)]]);
  rows.push([5, [sC('A5', 'GL Acct'), sC('B5', 'Account Name'), sC('C5', 'Balance')]]);
  rows.push([6, [sC('A6', 'Long Term Investments')]]);
  let r = 7; const ltiStart = r;
  for (const pair of LONG_TERM_INVESTMENTS) { rows.push([r, [sC('A' + r, pair[0]), sC('B' + r, pair[1]), nC('C' + r, r2(dev.balances[pair[0]] || 0))]]); r++; }
  const ltiEnd = r - 1;
  rows.push([r, [sC('B' + r, 'Total Long Term Investments'), fC('C' + r, 'SUM(C' + ltiStart + ':C' + ltiEnd + ')', r2(dev.ltiTotal))]]);
  const ltiTot = r; r += 2;
  rows.push([r, [sC('A' + r, 'Other Assets')]]); r++;
  const oaStart = r;
  for (const pair of OTHER_ASSETS) { rows.push([r, [sC('A' + r, pair[0]), sC('B' + r, pair[1]), nC('C' + r, r2(dev.balances[pair[0]] || 0))]]); r++; }
  const oaEnd = r - 1;
  rows.push([r, [sC('B' + r, 'Total Other Assets'), fC('C' + r, 'SUM(C' + oaStart + ':C' + oaEnd + ')', r2(dev.oaTotal))]]);
  const oaTot = r; r += 2;
  rows.push([r, [sC('B' + r, 'Total Development Costs'), fC('C' + r, 'C' + ltiTot + '+C' + oaTot, r2(dev.ltiTotal + dev.oaTotal))]]);
  const grand = r; r += 2;
  rows.push([r, [sC('B' + r, 'Source: CloudLedger - CLIP Property Owner (entity 54) trial balance as of ' + dev.as_of
    + '. Excludes placed-in-service fixed assets (15xxx/16xxx), cash, AR, prepaids, and intercompany.')]]);
  if (grand !== 28) throw new Error('dev sheet grand-total row moved to ' + grand + ' (Summary!H12 expects C28)');
  const rowXml = rows.map((pair) => '<row r="' + pair[0] + '">' + pair[1].join('') + '</row>').join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="42" customWidth="1"/>'
    + '<col min="3" max="3" width="18" customWidth="1"/></cols>'
    + '<sheetData>' + rowXml + '</sheetData></worksheet>';
}

// -- Build the "TB" worksheet XML (full CLRF trial balance) -------------------
// Columns: A = GL account code, B = account name, C = balance. Rows are sorted
// by account code. Returns { xml, rowByCode } where rowByCode maps each account
// code to its 1-based worksheet row, so the Summary schedule can point its Book
// Carrying Value cells at =TB!C<row> for the four investment accounts.
function buildTbSheetXml(tb) {
  const sC = (ref, txt) => '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(txt) + '</t></is></c>';
  const nC = (ref, v) => '<c r="' + ref + '"><v>' + v + '</v></c>';
  const fC = (ref, f, v) => '<c r="' + ref + '"><f>' + xmlEsc(f) + '</f><v>' + v + '</v></c>';
  const rows = [];
  rows.push([1, [sC('A1', 'County Line Rail Fund I, LP (CLRF)')]]);
  rows.push([2, [sC('A2', 'Trial Balance')]]);
  rows.push([3, [sC('A3', 'As of ' + tb.as_of)]]);
  rows.push([5, [sC('A5', 'GL Acct'), sC('B5', 'Account Name'), sC('C5', 'Balance')]]);
  const rowByCode = {};
  let r = 6; const dataStart = r; let total = 0;
  for (const acct of tb.accounts) {
    rowByCode[String(acct.code)] = r;
    total = r2(total + r2(acct.balance));
    rows.push([r, [sC('A' + r, String(acct.code)), sC('B' + r, String(acct.name || '')), nC('C' + r, r2(acct.balance))]]);
    r++;
  }
  const dataEnd = r - 1;
  r += 1;
  rows.push([r, [sC('B' + r, 'Total (net)'), fC('C' + r, 'SUM(C' + dataStart + ':C' + dataEnd + ')', total)]]);
  const rowXml = rows.map((pair) => '<row r="' + pair[0] + '">' + pair[1].join('') + '</row>').join('');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="46" customWidth="1"/>'
    + '<col min="3" max="3" width="20" customWidth="1"/></cols>'
    + '<sheetData>' + rowXml + '</sheetData></worksheet>';
  return { xml: xml, rowByCode: rowByCode, total: total };
}

// -- Core transform: template bytes + GL data + quarter -> new bytes ----------
async function transform(templateBuf, gl, qtr) {
  const zip = await JSZip.loadAsync(templateBuf);

  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const defs = [];
  const defRe = /<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"/g;
  let dm;
  while ((dm = defRe.exec(wbXml)) !== null) defs.push({ name: dm[1], rid: dm[2] });
  const rid2t = {};
  const relRe = /<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"/g;
  let rm;
  while ((rm = relRe.exec(relsXml)) !== null) rid2t[rm[1]] = rm[2];
  const name2path = {};
  for (const d of defs) if (rid2t[d.rid]) name2path[d.name] = 'xl/' + rid2t[d.rid].replace(/^\//, '');

  const summaryDef = defs.find((d) => d.name.trim() === 'Summary');
  if (!summaryDef) throw new Error('template missing Summary sheet');
  const P = {
    summary: name2path[summaryDef.name],
    soi: name2path['SOI'],
    sales: name2path['Sales Approach CLIP'],
    equip: name2path['CLR Equipment Roster'],
  };
  for (const k of Object.keys(P)) if (!P[k]) throw new Error('template missing sheet for ' + k);

  // (1) Replace the SOI tab with the full CLRF trial balance and rename it "TB".
  // We overwrite the SOI sheet's XML part in place (keeping its rId/sheetId, so no
  // relationship surgery is needed) with a freshly built trial-balance grid, then
  // rename the sheet in workbook.xml from "SOI" to "TB". The four investment
  // accounts' TB rows are captured so Summary D12-15 can reference them directly.
  const tbBuilt = buildTbSheetXml(gl.tb);
  zip.file(P.soi, tbBuilt.xml);
  // Rename the sheet definition SOI -> TB in workbook.xml. wbXml2 becomes the
  // working copy that all later workbook.xml edits build on (e.g. the dev sheet).
  let wbXml2 = wbXml;
  const soiSheetRe = /(<sheet name=")SOI("[^>]*\/>)/;
  if (!soiSheetRe.test(wbXml2)) throw new Error('could not find SOI sheet definition to rename to TB');
  wbXml2 = wbXml2.replace(soiSheetRe, '$1TB$2');
  // Also repoint any defined-name references to the old sheet name (the template
  // carries _xlnm.Print_Area and _xlnm.Print_Titles as "SOI!$A$1:$K$18" /
  // "SOI!$2:$5"). Left unchanged, these reference a sheet that no longer exists
  // and Excel flags the workbook as corrupt on open. The sheet name has no
  // special characters, so references are the unquoted token "SOI!"; there is no
  // quoted 'SOI'! form to worry about. Rewrite the token to "TB!".
  wbXml2 = wbXml2.replace(/(?<![A-Za-z0-9_'])SOI!/g, 'TB!');
  zip.file('xl/workbook.xml', wbXml2);
  const bcvTotal = r2(Object.keys(BCV).reduce((a, c) => a + r2(gl.bcv[c]), 0));

  // (2) CLIP GL Dev Costs tab (add or replace).
  const devTotal = r2(gl.dev.ltiTotal + gl.dev.oaTotal);
  const devSheet = buildDevSheetXml(gl.dev);
  const devSheetName = 'CLIP GL Dev Costs';
  if (name2path[devSheetName]) {
    zip.file(name2path[devSheetName], devSheet);
  } else {
    const nums = Object.values(name2path)
      .map((p) => Number((p.match(/sheet(\d+)\.xml$/) || [])[1])).filter(Boolean);
    const devNum = Math.max.apply(null, nums) + 1;
    const rids = [];
    const ridRe = /Id="rId(\d+)"/g; let im;
    while ((im = ridRe.exec(relsXml)) !== null) rids.push(Number(im[1]));
    const devRid = 'rId' + (Math.max.apply(null, rids) + 1);
    const sids = [];
    const sidRe = /sheetId="(\d+)"/g; let sm;
    while ((sm = sidRe.exec(wbXml)) !== null) sids.push(Number(sm[1]));
    const devSid = Math.max.apply(null, sids) + 1;
    const newPath = 'xl/worksheets/sheet' + devNum + '.xml';
    zip.file(newPath, devSheet);
    // Build on wbXml2 (which already carries the SOI->TB rename) so we don't
    // clobber it when appending the dev sheet.
    zip.file('xl/workbook.xml', wbXml2.replace('</sheets>',
      '<sheet name="' + devSheetName + '" sheetId="' + devSid + '" r:id="' + devRid + '"/></sheets>'));
    zip.file('xl/_rels/workbook.xml.rels', relsXml.replace('</Relationships>',
      '<Relationship Id="' + devRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
      + 'Target="worksheets/sheet' + devNum + '.xml"/></Relationships>'));
    const ctXml = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ctXml.replace('</Types>',
      '<Override PartName="/' + newPath + '" '
      + 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'));
    name2path[devSheetName] = newPath;
  }

  // (3+4) Summary + Sales Approach chain and the C8 date.
  let sales = await zip.file(P.sales).async('string');
  const I67 = numFromCell(sales, 'I67');
  let summary = await zip.file(P.summary).async('string');
  const equipXml = await zip.file(P.equip).async('string');
  const equip = numFromCell(equipXml, 'C11');

  // Target CLIP total valuation (J12). We hold this at the appraiser's concluded
  // total while H12 rises to the GL dev cost, solving the stabilization discount
  // as the plug (the Q1 2026 methodology). We must NOT read this from the
  // template's J12 cache: in the distributed workbook H12 pulls the requisition
  // link 'CLIP Dev Costs'!I40, which resolves to #VALUE! (broken external ref),
  // and that error cascades into J12 -> its cache is "#VALUE!", which parses to
  // NaN/0 and blows up the plug. So anchor to the fixed concluded value, and only
  // trust the template's J12 when it is a clean finite number.
  const rawJ12 = numFromCellSafe(summary, 'J12');
  const targetJ12 = (rawJ12 !== null && isFinite(rawJ12) && rawJ12 > 0)
    ? r2(rawJ12) : CLIP_CONCLUDED_VALUATION;

  // Book Carrying Value column D links directly to the TB tab: each investment
  // account's balance cell (=TB!C<row>) feeds its Summary row (D12=CLIP,
  // D13=Silsbee, D14=Buna, D15=SRN, D16=SUM). Rows come from the TB grid just
  // built, so links track wherever accounts land in the sorted trial balance.
  for (const code of Object.keys(BCV)) {
    const meta = BCV[code];
    const ref = 'D' + meta.summaryRow;
    const tbRow = tbBuilt.rowByCode[code];
    const val = r2(gl.bcv[code]);
    if (!tbRow) throw new Error('investment account ' + code + ' (' + meta.label + ') not found in CLRF TB');
    summary = replaceCell(summary, ref, fCell(ref, styleOf(summary, ref), 'TB!C' + tbRow, val));
  }
  summary = replaceCell(summary, 'D16', fCell('D16', styleOf(summary, 'D16'), 'SUM(D12:D15)', bcvTotal));

  const H12 = devTotal;
  const I12 = r2(targetJ12 - H12);
  const I69 = r2(I12 - equip);
  const I68 = r2(I69 - I67);
  const G12 = I12; const K12 = targetJ12;

  const g13 = numFromCellSafe(summary, 'G13') || 0, g14 = numFromCellSafe(summary, 'G14') || 0, g15 = numFromCellSafe(summary, 'G15') || 0;
  const k13 = numFromCellSafe(summary, 'K13') || 0, k14 = numFromCellSafe(summary, 'K14') || 0, k15 = numFromCellSafe(summary, 'K15') || 0;
  const G16 = r2(G12 + g13 + g14 + g15);
  const H16 = H12, I16 = I12, J16 = targetJ12;
  const K16 = r2(K12 + k13 + k14 + k15);

  const setF = (xml, ref, formula, val) => replaceCell(xml, ref, fCell(ref, styleOf(xml, ref), formula, val));
  summary = setF(summary, 'H12', "'CLIP GL Dev Costs'!C28", H12);
  summary = setF(summary, 'G12', "'Sales Approach CLIP'!$I$69+'CLR Equipment Roster'!C11", G12);
  summary = setF(summary, 'I12', 'G12', I12);
  summary = setF(summary, 'J12', 'SUM(H12:I12)', targetJ12);
  summary = setF(summary, 'K12', 'J12', K12);
  summary = setF(summary, 'G16', 'SUM(G12:G15)', G16);
  summary = replaceCell(summary, 'H16',
    '<c r="H16" s="' + (styleOf(summary, 'H16') || '') + '"><f t="shared" ref="H16:J16" si="0">SUM(H12:H15)</f><v>' + H16 + '</v></c>');
  summary = replaceCell(summary, 'I16',
    '<c r="I16" s="' + (styleOf(summary, 'I16') || '') + '"><f t="shared" si="0"/><v>' + I16 + '</v></c>');
  summary = replaceCell(summary, 'J16',
    '<c r="J16" s="' + (styleOf(summary, 'J16') || '') + '"><f t="shared" si="0"/><v>' + J16 + '</v></c>');
  summary = setF(summary, 'K16', 'SUM(K12:K15)', K16);
  { const s = styleOf(summary, 'C8'); summary = replaceCell(summary, 'C8', numCell('C8', s, excelSerial(qtr.end))); }
  zip.file(P.summary, summary);

  { const s = styleOf(sales, 'I68'); sales = replaceCell(sales, 'I68', numCell('I68', s, I68)); }
  { const s = styleOf(sales, 'I69'); sales = replaceCell(sales, 'I69', fCell('I69', s, 'SUM(I67:I68)', I69)); }
  zip.file(P.sales, sales);

  // Drop the calculation chain. The replaced SOI sheet's formula cells (A2, J15,
  // J17 on sheetId 47) stay registered in xl/calcChain.xml, but the TB grid that
  // now occupies that sheet part has different formula cells — and stale
  // calcChain entries pointing at cells with no formulas make Excel flag the
  // workbook as corrupt on open. Removing the whole part (and its
  // [Content_Types].xml Override and workbook.xml.rels Relationship) is safe:
  // Excel rebuilds the calc chain silently. NOTE: JSZip remove() is recursive on
  // folder paths; this targets the single calcChain.xml file only.
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    const ctNow = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml',
      ctNow.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
    const wrelsNow = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    zip.file('xl/_rels/workbook.xml.rels',
      wrelsNow.replace(/<Relationship [^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
  }

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const summaryData = {
    quarter: qtr.label,
    valuation_date: qtr.end,
    book_carrying_value: {
      by_property: Object.keys(BCV).reduce((o, c) => { o[BCV[c].label] = r2(gl.bcv[c]); return o; }, {}),
      total: bcvTotal,
    },
    clip_dev_cost: { long_term_investments: r2(gl.dev.ltiTotal), other_assets: r2(gl.dev.oaTotal), total: devTotal },
    clip_valuation: {
      as_complete_plus_land_I67: I67,
      stabilization_discount_I68: I68,
      stabilization_pct: I67 ? r2((I68 / I67) * 100) : null,
      sales_as_is_I69: I69,
      equipment: equip,
      total_valuation_J12: targetJ12,
    },
    portfolio_total_K16: K16,
  };
  return { buf: outBuf, summary: summaryData };
}

// -- Gather live GL data ------------------------------------------------------
function gatherGl(ctx, qtr) {
  const { computeBalances } = ctx;
  const clrfRows = computeBalances(CLRF_ENTITY_ID, { as_of: qtr.end });
  const clrfByCode = {};
  for (const x of clrfRows) clrfByCode[String(x.code)] = x;
  const bcv = {};
  const bcvName = {};
  for (const code of Object.keys(BCV)) {
    const row = clrfByCode[code];
    bcv[code] = row ? r2(row.balance) : 0;
    bcvName[code] = row && row.name ? String(row.name) : '';
  }
  // Full CLRF trial balance for the TB tab, sorted by account code (numeric
  // where possible, then lexical). Zero-balance accounts are dropped to keep the
  // tab readable; the four investment accounts are always retained so their
  // =TB!C<row> links resolve even if a balance nets to zero at a quarter-end.
  const tbAccounts = clrfRows
    .filter((x) => BCV[String(x.code)] || Math.abs(r2(x.balance)) >= 0.005)
    .map((x) => ({ code: String(x.code), name: x.name || '', balance: r2(x.balance) }))
    .sort((a, b) => {
      const na = Number(a.code), nb = Number(b.code);
      if (isFinite(na) && isFinite(nb) && na !== nb) return na - nb;
      return a.code < b.code ? -1 : a.code > b.code ? 1 : 0;
    });
  const clipRows = computeBalances(CLIP_ENTITY_ID, { as_of: qtr.end });
  const clipByCode = {};
  for (const x of clipRows) clipByCode[String(x.code)] = x;
  const balances = {};
  let ltiTotal = 0, oaTotal = 0; const missing = [];
  for (const pair of LONG_TERM_INVESTMENTS) {
    const row = clipByCode[pair[0]]; const v = row ? r2(row.balance) : 0;
    balances[pair[0]] = v; ltiTotal = r2(ltiTotal + v); if (!row) missing.push(pair[0]);
  }
  for (const pair of OTHER_ASSETS) {
    const row = clipByCode[pair[0]]; const v = row ? r2(row.balance) : 0;
    balances[pair[0]] = v; oaTotal = r2(oaTotal + v); if (!row) missing.push(pair[0]);
  }
  return {
    bcv: bcv, bcvName: bcvName,
    tb: { as_of: qtr.end, accounts: tbAccounts },
    dev: { as_of: qtr.end, balances: balances, ltiTotal: ltiTotal, oaTotal: oaTotal, missing: missing },
  };
}

// -- Route registration -------------------------------------------------------
function registerValuationRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;

  app.post('/api/workpapers/valuation-summary/:entity_id/generate', auth,
    requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const qtr = resolveQuarter((req.body && req.body.quarter_end) || (req.query && req.query.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';

        const tpl = findTemplate(ctx, eid, qtr);
        if (!tpl) {
          return res.status(400).json({ error: 'No prior valuation workbook found under Workpapers/Valuation. '
            + "Save the prior quarter's file there first." });
        }
        if (!fs.existsSync(tpl.abs_path)) {
          return res.status(500).json({ error: 'Template row exists but file is missing on disk: ' + tpl.original_name });
        }
        const templateBuf = fs.readFileSync(tpl.abs_path);
        const gl = gatherGl(ctx, qtr);
        const result = await transform(templateBuf, gl, qtr);
        const saved = saveToWorkpapers(ctx, eid, qtr, result.buf, who);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Valuation-Summary', JSON.stringify(Object.assign({
          template_from: tpl.source_folder + '/' + tpl.original_name,
          saved_to: saved.folder_path + '/' + saved.original_name,
          replaced: saved.replaced,
          missing_dev_accounts: gl.dev.missing,
        }, result.summary)).replace(/[\r\n]/g, ' '));
        res.send(result.buf);
      } catch (e) {
        console.error('valuation-summary failed:', e);
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = {
  registerValuationRoutes: registerValuationRoutes,
  transform: transform,
  resolveQuarter: resolveQuarter,
  gatherGl: gatherGl,
  buildDevSheetXml: buildDevSheetXml,
  buildTbSheetXml: buildTbSheetXml,
  findTemplate: findTemplate,
  valFolder: valFolder,
  valFileName: valFileName,
  excelSerial: excelSerial,
  LONG_TERM_INVESTMENTS: LONG_TERM_INVESTMENTS,
  OTHER_ASSETS: OTHER_ASSETS,
  BCV: BCV,
};
