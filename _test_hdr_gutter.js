// Verifies the column-heading gutter fix by running the PATCHED colHeaders()
// source straight out of server/financials.js against stub page/font objects and
// reporting where each heading actually lands.
const fs = require('fs');
const { PDFDocument, StandardFonts } = require('pdf-lib');

const src = fs.readFileSync(require.resolve('./server/financials.js'), 'utf8');

// Pull the exact colHeaders method body (from its signature to the line before
// `sectionTitle(str) {`), so we test shipped code rather than a re-derivation.
const start = src.indexOf('    colHeaders(labels, hopts = {}) {');
const end = src.indexOf('    sectionTitle(str) {');
if (start < 0 || end < 0 || end < start) throw new Error('could not locate colHeaders in financials.js');
const body = src.slice(start, end).replace(/,\s*$/, '');
if (!/MIN_HDR_GUTTER/.test(body)) throw new Error('patched gutter logic NOT present in colHeaders');

const FS = { title: 11, sub: 9.5, head: 8, row: 8.5, foot: 7.5 };
const ROW_H = 12;
const HDR_TRAIL_GAP = 3 + 2 * ROW_H;

function runHeader(cols, labels, hopts, bold) {
  const drawn = [];
  const page = {
    drawText: (t, o) => drawn.push({ kind: 'text', t, x: o.x, y: o.y, w: bold.widthOfTextAtSize(t, FS.head) }),
    drawLine: (o) => drawn.push({ kind: 'line', x1: o.start.x, x2: o.end.x, y: o.start.y }),
  };
  let y = 700;
  const ensure = () => {};
  let _replaying = false, _hdrSpec = null;
  const obj = eval('({' + body + '})');
  // Bind the closure variables colHeaders reads.
  const ctx = { cols, bold, FS, ensure, page, get y() { return y; }, set y(v) { y = v; } };
  // eval in a scope where the identifiers resolve:
  const fn = new Function('cols', 'bold', 'FS', 'ensure', 'page', 'yRef', 'HDR_TRAIL_GAP', 'rgb',
    'let _replaying=false,_hdrSpec=null;let y=yRef.v;const o={' + body + '};'
    + 'const r=o.colHeaders.bind(o);return function(l,h){r(l,h);yRef.v=y;};');
  const yRef = { v: 700 };
  const call = fn(cols, bold, FS, ensure, page, yRef, HDR_TRAIL_GAP, require('pdf-lib').rgb);
  call(labels, hopts);
  return drawn;
}

(async () => {
  const pdf = await PDFDocument.create();
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const RIGHT = 612 - 54;

  const cases = [
    { name: 'Balance Sheets (Sept close, worst case)', cols: [RIGHT - 150, RIGHT - 75, RIGHT], labels: ['September 30, 2026', 'August 31, 2026', 'Change'] },
    { name: 'Balance Sheets (Odyssey, June close)', cols: [RIGHT - 150, RIGHT - 75, RIGHT], labels: ['June 30, 2026', 'December 31, 2025', 'Change'] },
    { name: 'Statements of Operations (Sept close)', cols: [RIGHT - 216, RIGHT - 144, RIGHT - 72, RIGHT], labels: ['September 30, 2026', 'August 31, 2026', 'Change', 'Year to Date'] },
    { name: 'Statements of Operations (June close)', cols: [RIGHT - 216, RIGHT - 144, RIGHT - 72, RIGHT], labels: ['June 30, 2026', 'May 31, 2026', 'Change', 'Year to Date'] },
  ];

  let worst = Infinity;
  for (const c of cases) {
    const drawn = runHeader(c.cols, c.labels, { underline: true }, bold);
    const texts = drawn.filter(d => d.kind === 'text');
    console.log('\n=== ' + c.name + ' ===');
    // Group drawn text by baseline row so we can inspect the bottom line.
    const rows = {};
    for (const t of texts) (rows[t.y] = rows[t.y] || []).push(t);
    for (const yy of Object.keys(rows).sort((a, b) => b - a)) {
      const line = rows[yy].sort((a, b) => a.x - b.x);
      console.log('  y=' + yy + '  ' + line.map(t => '"' + t.t + '" [' + t.x.toFixed(1) + '..' + (t.x + t.w).toFixed(1) + ']').join('  '));
      for (let i = 1; i < line.length; i++) {
        const gap = line[i].x - (line[i - 1].x + line[i - 1].w);
        if (gap < worst) worst = gap;
        console.log('        gap after "' + line[i - 1].t + '": ' + gap.toFixed(1) + 'pt' + (gap < 6 ? '   <-- TOO TIGHT' : ''));
      }
    }
  }
  console.log('\nSMALLEST GAP ACROSS ALL CASES: ' + worst.toFixed(1) + 'pt');
  console.log(worst >= 6 ? 'PASS' : 'FAIL');
})();
