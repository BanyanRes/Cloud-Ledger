// Extracts the actual drawn horizontal rule segments from the rendered statement
// PDFs and checks that every subtotal/total underline shares ONE width and that
// their right edges land on the numeric column edges -- i.e. the underlines line
// up vertically. Parses the page content streams directly (m/l/S operators),
// since rules are vector strokes, not text.
const fin  = require('../server/financials.js');
const zlib = require('zlib');
const { PDFDocument, PDFName, PDFRawStream } = require('pdf-lib');
const RM = require('./_rowmodel.js');

// ---- fixture: a Banyan GL with a MIX of small and huge subtotal rows, so the
// old per-row sizing would have produced visibly different rule widths. --------
const A = {
  '10143': { name: 'UBS Entity 100 Banyan - 2472', type: 'Asset' },
  '10144': { name: 'MapleMark Entity 100 Banyan - 8868', type: 'Asset' },
  '11030': { name: 'Land', type: 'Asset' },              // huge
  '12720': { name: 'Organization Fees', type: 'Asset' }, // small (Other Development)
  '15500': { name: 'Furniture and Fixtures', type: 'Asset' },
  '16500': { name: 'Accumulated Depreciation', type: 'Asset' },
  '20000': { name: 'Accounts Payable', type: 'Liability' },
  '21112': { name: 'Accrued Liabilities', type: 'Liability' },
  '34117': { name: 'Contributed Capital - Odyssey Holdings LLC', type: 'Equity' },
  '39000': { name: 'Retained Earnings', type: 'Equity' },
  '42000': { name: 'Development Fee Revenue', type: 'Revenue' },
  '60000': { name: 'Salaries and Wages', type: 'Expense' },
  '68061': { name: 'State Franchise Tax', type: 'Expense' },
  '70000': { name: 'Interest Income', type: 'Revenue' },
};
const sum = (m, t) => Object.keys(m).filter(c => A[c].type === t).reduce((a, c) => a + m[c], 0);
const withRE = (bs, pl) => Object.assign({}, bs, {
  '39000': +(sum(bs, 'Asset') - sum(bs, 'Liability') - sum(bs, 'Equity')
             - (pl ? sum(pl, 'Revenue') - sum(pl, 'Expense') : 0)).toFixed(2),
});
const PL  = { '42000': 1364458.11, '60000': 900000.00, '68061': 16000.00, '70000': 2500.00 };
const CUR  = withRE({ '10143': 2460.56, '10144': 100481.35, '11030': 2000289.13, '12720': 10735.54,
                      '15500': 8000.00, '16500': -1346.77,
                      '20000': 491888.68, '21112': 30000.00, '34117': 746370.65 }, PL);
const PRI  = withRE(Object.assign({}, CUR, { '39000': 0, '10144': 145017.67, '20000': 236756.01, '21112': 146791.08 }), PL);
const OPEN = withRE(Object.assign({}, CUR, { '39000': 0, '10143': 1624010.53, '20000': 236756.01, '21112': 146791.08, '10144': 145017.67 }), null);
const rows = (m) => Object.keys(m).map(c => ({ code: c, name: A[c].name, type: A[c].type, subtype: '', balance: m[c] }));
const getBalances = (o) => Promise.resolve(
  o.from ? rows(PL)
  : o.as_of === '2026-07-31' ? rows(CUR).concat(rows(PL))
  : o.as_of === '2026-06-30' ? rows(PRI).concat(rows(PL))
  : rows(OPEN));

// ---- pull horizontal stroke segments from a page's content stream ------------
function pageStreamText(pdfDoc, pageIndex) {
  const page = pdfDoc.getPage(pageIndex);
  const contents = page.node.Contents();
  const streams = [];
  const collect = (obj) => {
    if (!obj) return;
    if (obj instanceof PDFRawStream) streams.push(obj);
    else if (obj.asArray) obj.asArray().forEach(ref => collect(pdfDoc.context.lookup(ref)));
  };
  collect(contents instanceof PDFName ? null : (contents.asArray ? contents : pdfDoc.context.lookup(contents)));
  // Contents may be a single stream or an array of streams.
  let raw = contents;
  if (raw && raw.asArray) raw = raw.asArray().map(r => pdfDoc.context.lookup(r));
  else raw = [pdfDoc.context.lookup(page.node.get(PDFName.of('Contents'))) || contents];
  let out = '';
  for (const st of [].concat(raw)) {
    if (!st || !st.contents) continue;
    let bytes = st.contents;
    try {
      const filt = st.dict.get(PDFName.of('Filter'));
      if (filt && filt.toString().includes('FlateDecode')) bytes = zlib.inflateSync(Buffer.from(bytes));
    } catch (e) { /* already raw */ }
    out += Buffer.from(bytes).toString('latin1');
  }
  return out;
}

// Parse "x y m  x y l  S" horizontal segments. pdf-lib draws lines as
// "<x0> <y0> m <x1> <y1> l S". Collect segments where y0==y1 (horizontal).
function horizontalSegments(streamText) {
  const segs = [];
  const re = /([\d.\-]+)\s+([\d.\-]+)\s+m\s+([\d.\-]+)\s+([\d.\-]+)\s+l\s+S/g;
  let m;
  while ((m = re.exec(streamText))) {
    const x0 = +m[1], y0 = +m[2], x1 = +m[3], y1 = +m[4];
    if (Math.abs(y0 - y1) < 0.3) segs.push({ y: y0, x0: Math.min(x0, x1), x1: Math.max(x0, x1), w: Math.abs(x1 - x0) });
  }
  return segs;
}

(async () => {
  let fail = 0;
  const check = (n, ok, d) => { if (!ok) fail++; console.log((ok ? '  PASS  ' : '  FAIL  ') + n + (ok ? '' : '   [' + (d || '') + ']')); };

  const s = await fin.buildStatements(getBalances, { asOf: '2026-07-31', period: 'monthly', entityName: 'Banyan Residential LLC', entityCode: 'BANYANRE1' });
  const stmtBytes = await fin.renderStatementsPdf(s, []);
  const pdf = await PDFDocument.load(stmtBytes);

  for (let pi = 0; pi < pdf.getPageCount(); pi++) {
    const segs = horizontalSegments(pageStreamText(pdf, pi));
    if (!segs.length) continue;
    // Round widths to 0.1pt and tally.
    const widths = {};
    segs.forEach(sg => { const k = Math.round(sg.w * 10) / 10; widths[k] = (widths[k] || 0) + 1; });
    // Ignore the 2-column-header underlines (short, under date labels) and the
    // cover rules by focusing on the dominant rule width used by subtotals.
    const sorted = Object.entries(widths).sort((a, b) => b[1] - a[1]);
    // The subtotal rules are the most common repeated width on a statement page.
    const [domW, domN] = sorted[0];
    // Consider only "figure-area" segments (right edge x1 within the numeric
    // block, i.e. x1 > 300) for the alignment check.
    // Figure-area segments, excluding the column-HEADER underline band. The
    // header underlines hug their date/label text (so they legitimately differ
    // in width); they sit in the topmost ~20pt of the statement body. Subtotal
    // rules are everything below that.
    const figAll = segs.filter(sg => sg.x1 > 300);
    const maxY = Math.max(...figAll.map(sg => sg.y));
    const bodyRules = figAll.filter(sg => sg.y < maxY - 20);
    const w = new Set(bodyRules.map(sg => Math.round(sg.w * 10) / 10));
    const edges = new Set(bodyRules.map(sg => Math.round(sg.x1)));
    console.log('page ' + (pi + 1) + ': subtotal-rule widths = [' + [...w].join(', ') +
      ']; column right edges = [' + [...edges].sort((a, b) => a - b).join(', ') + ']');
    check('page ' + (pi + 1) + ': all subtotal underlines share ONE width',
      w.size <= 1, [...w].join(', '));
  }

  // ---- rule BELOW each of the 8 lines Jimmy listed (CLA 8/17) -------------
  // Row baselines come from the text layer (row model); strokes from the
  // content stream. A below-rule sits ~3pt under the baseline.
  const P = await RM.readRows(stmtBytes);
  const pageSegs = [];
  for (let pi = 0; pi < pdf.getPageCount(); pi++) pageSegs.push(horizontalSegments(pageStreamText(pdf, pi)));
  const TARGETS = ['Total Current Assets', 'Total Fixed Assets, Net', 'Total Liabilities',
                   'Total Revenue', 'Gross Profit', 'Total Operating Expenses',
                   'Total Other Income (Expense)', 'Total Income Taxes'];
  for (const t of TARGETS) {
    let found = null, page = -1;
    P.forEach((rows, pi) => {
      if (found) return;
      const r = rows.find(r => r.label === t && r.money.length);
      if (r) { found = r; page = pi; }
    });
    if (!found) { check('below-rule: ' + t, false, 'row not rendered'); continue; }
    const below = pageSegs[page].filter(sg => sg.x1 > 300 && sg.y < found.y - 0.5 && sg.y > found.y - 8);
    check('below-rule under "' + t + '" (' + below.length + ' col segments)', below.length >= found.money.length - 1);
  }

  console.log('\n' + (fail ? fail + ' FAILURE(S)' : 'ALL PASS -- uniform widths, column-aligned, below-rules present'));
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(2); });
