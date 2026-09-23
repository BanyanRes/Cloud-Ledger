// ─────────────────────────────────────────────────────────────────────────────
// build_subclose_template.js
//
// One-time (re-runnable) transform that turns the fund administrator's (Weaver)
// sub-close workbook — which builds its "Sub Close Summary" tab entirely from
// formulas linked to an EXTERNAL Weaver source workbook — into a SELF-CONTAINED
// workpaper whose Summary traces to in-workbook supporting schedules instead.
//
// Source of truth:  server/assets/subclose_template_src.xlsx  (Weaver original,
//                   external links + their cached values intact)
// Output:           server/assets/subclose_template.xlsx      (self-contained)
//
// The Summary references six external Weaver sheets:
//   Subsequent Close, True-Up, Closing Interest, Management Fees - New Investor,
//   Subsequent Close - OLD B4 25k  ── all with cached values embedded, so we
//                                      reproduce them verbatim as supporting tabs;
//   May 26 Distribution Checking   ── NOT cached, so we reconstruct the two
//                                      columns the Summary reads (Y and AA) from
//                                      the Summary's own cached XLOOKUP results,
//                                      keyed by investor name. Because every DC
//                                      lookup is wrapped in ROUND(...,2), the
//                                      reconstruction reproduces each result to
//                                      the penny, so the workbook ties out exactly.
//
// Re-point: ExcelJS keeps the formula text ('[1]Sheet'!Cell) even after it drops
// the external-link parts, so re-pointing is simply removing the "[1]" marker —
// '[1]True-Up'!A:A  ->  'True-Up'!A:A — which now resolves to the in-workbook tab.
// ─────────────────────────────────────────────────────────────────────────────
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

const SRC = path.join(__dirname, '..', 'assets', 'subclose_template_src.xlsx');
const OUT = path.join(__dirname, '..', 'assets', 'subclose_template.xlsx');

const decode = (s) => String(s == null ? '' : s)
  .replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&apos;/g, "'")
  .replace(/&quot;/g, '"').replace(/&amp;/g, '&');

// Sheets that have embedded cached values we reproduce verbatim.
const CACHED_SHEETS = [
  'Subsequent Close',
  'True-Up',
  'Closing Interest',
  'Management Fees - New Investor',
  'Subsequent Close - OLD B4 25k',
];
const DC_SHEET = 'May 26 Distribution Checking';

// Short header notes for the sparse supporting tabs (put in row 1 as context).
const SHEET_NOTE = {
  'Closing Interest': 'CLRF Subsequent Close — Closing Interest (fund administrator schedule, values pinned)',
  'Management Fees - New Investor': 'CLRF Subsequent Close — Management Fees, New Investor (fund administrator schedule, values pinned)',
  'Subsequent Close - OLD B4 25k': 'CLRF Subsequent Close — prior ($25k) basis variant (fund administrator schedule, value pinned)',
  'May 26 Distribution Checking': 'CLRF Subsequent Close — May 2026 Distribution Checking (reconstructed from workpaper results; col Y = contribution reallocation, col AA = cash distribution)',
};

const colOf = (a) => a.match(/^[A-Z]+/)[0];
const rowOf = (a) => +a.match(/\d+/)[0];

async function extractCache(zipBuf) {
  const zip = await JSZip.loadAsync(zipBuf);
  const el1 = await zip.file('xl/externalLinks/externalLink1.xml').async('string');
  const names = [...el1.matchAll(/<sheetName val="([^"]*)"/g)].map((m) => decode(m[1]));
  const bySheet = {};
  for (const sb of el1.matchAll(/<sheetData sheetId="(\d+)"[^>]*>([\s\S]*?)<\/sheetData>/g)) {
    const nm = names[+sb[1]];
    const cells = {};
    for (const c of sb[2].matchAll(/<cell r="([^"]*)"([^>]*)>(?:<v>([\s\S]*?)<\/v>)?<\/cell>/g)) {
      const addr = c[1].replace(/\$/g, '');
      const t = (c[2].match(/t="([^"]*)"/) || [])[1];
      const raw = c[3];
      let val;
      if (raw == null || raw === '') val = null;
      else if (t === 'str' || t === 'e') val = decode(raw);
      else if (t === 'b') val = raw === '1';
      else val = Number(raw);
      cells[addr] = val;
    }
    bySheet[nm] = cells;
  }
  return bySheet;
}

// text of an ExcelJS cell value (label/name)
function cellText(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return v.result;
    if ('text' in v) return v.text;
    return '';
  }
  return v;
}
function cellResult(cell) {
  const v = cell.value;
  if (v && typeof v === 'object' && 'result' in v) return v.result;
  if (typeof v === 'number') return v;
  return null;
}

async function main() {
  const srcBuf = fs.readFileSync(SRC);
  const cache = await extractCache(srcBuf);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const summary = wb.getWorksheet('Sub Close Summary');
  if (!summary) throw new Error('source missing Sub Close Summary');

  // ── 1. Reconstruct May 26 Distribution Checking from Summary results ──────────
  // Scan every Summary formula cell that reads DC; name is in col E (same row),
  // value is the cell's own cached result. Y-block: value = +result; AA-block:
  // formula has a leading '-', so DC!AA = -result.
  const dc = {}; // name -> { Y, AA }
  let dcCellsSeen = 0;
  summary.eachRow((row) => {
    row.eachCell((cell) => {
      const v = cell.value;
      if (!v || typeof v !== 'object' || typeof v.formula !== 'string') return;
      const f = v.formula;
      if (!f.includes(DC_SHEET)) return;
      dcCellsSeen++;
      const r = rowOf(cell.address);
      const name = cellText(summary.getCell('E' + r));
      const res = cellResult(cell);
      if (name == null || name === '' || typeof res !== 'number') return;
      const rec = (dc[name] = dc[name] || {});
      if (/!\$?Y:\$?Y/.test(f)) rec.Y = res;                 // ROUND(XLOOKUP(..,Y:Y),2)
      else if (/!\$?AA:\$?AA/.test(f)) rec.AA = -res;        // -ROUND(XLOOKUP(..,AA:AA),2)
    });
  });
  const dcNames = Object.keys(dc);

  // ── 2. Create the supporting worksheets ──────────────────────────────────────
  const created = [];
  function addSheet(name) {
    const ws = wb.addWorksheet(name, { properties: { tabColor: { argb: 'FFDDEBF7' } } });
    created.push(name);
    return ws;
  }

  // 2a. cached sheets — reproduce every cached cell at its original address
  for (const name of CACHED_SHEETS) {
    const ws = addSheet(name);
    const cells = cache[name] || {};
    let maxCol = 1;
    for (const [addr, val] of Object.entries(cells)) {
      if (val == null) continue;
      ws.getCell(addr).value = val;
      const ci = ws.getColumn(colOf(addr)).number;
      if (ci > maxCol) maxCol = ci;
    }
    // sparse sheets: add a context note on a spare row so the tab is self-describing
    if (SHEET_NOTE[name] && Object.keys(cells).length <= 6) {
      ws.getCell('A1').value = SHEET_NOTE[name];
    }
    for (let c = 1; c <= Math.min(maxCol, 40); c++) ws.getColumn(c).width = 14;
    ws.getColumn(1).width = 42;
  }

  // 2b. Distribution Checking — reconstructed A / Y / AA
  const dcWs = addSheet(DC_SHEET);
  dcWs.getCell('A1').value = SHEET_NOTE[DC_SHEET];
  dcWs.getCell('A3').value = 'Investor Name';
  dcWs.getCell('Y3').value = 'Contribution reallocation (Y)';
  dcWs.getCell('AA3').value = 'Cash distribution (AA)';
  let dr = 4;
  for (const name of dcNames) {
    dcWs.getCell('A' + dr).value = name;
    if (typeof dc[name].Y === 'number') dcWs.getCell('Y' + dr).value = dc[name].Y;
    if (typeof dc[name].AA === 'number') dcWs.getCell('AA' + dr).value = dc[name].AA;
    dr++;
  }
  dcWs.getColumn('A').width = 42; dcWs.getColumn('Y').width = 18; dcWs.getColumn('AA').width = 18;

  // ── 3. Re-point every Summary formula: drop the "[1]" external marker ─────────
  let rewired = 0;
  for (const ws of wb.worksheets) {
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        const v = cell.value;
        if (v && typeof v === 'object' && typeof v.formula === 'string' && v.formula.includes('[1]')) {
          const nf = v.formula.replace(/\[1\]/g, '');
          cell.value = { formula: nf, result: ('result' in v) ? v.result : undefined };
          rewired++;
        }
        // shared-formula masters store the text on .formula too; followers carry
        // sharedFormula — ExcelJS recomputes those from the master, so nothing else
        // to touch here.
      });
    });
  }

  await wb.xlsx.writeFile(OUT);

  console.log('DC formula cells seen:', dcCellsSeen, ' distinct investors:', dcNames.length);
  console.log('  with Y:', dcNames.filter((n) => 'Y' in dc[n]).length,
              ' with AA:', dcNames.filter((n) => 'AA' in dc[n]).length);
  console.log('supporting tabs created:', created.join(', '));
  console.log('formula cells re-pointed (had [1]):', rewired);
  console.log('written:', OUT);
}
main().catch((e) => { console.error(e); process.exit(1); });
