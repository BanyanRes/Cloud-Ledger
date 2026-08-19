// CloudLedger — content blocks for house-style Word files.
// Shorthands and the three approved composite blocks. Require these; do not
// hand-roll tables or panels.
const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle,
} = require('docx');
const H = require('./house.js');

// ── run / paragraph shorthands ──────────────────────────────────────────────
const t = (text, o = {}) => new TextRun({ text, ...o });
const b = (text) => t(text, { bold: true });
const i = (text) => t(text, { italics: true });
const m = (text) => t(text, { font: 'Consolas', size: 19 });   // inline code
const P = (kids, opts = {}) => new Paragraph({
  children: typeof kids === 'string' ? [new TextRun(kids)] : kids, ...opts,
});
const H1 = (s) => P(s, { style: 'Title1' });
const H2 = (s) => P(s, { style: 'H2' });
const H3 = (s) => P(s, { style: 'H3' });

// ── shaded code panel ───────────────────────────────────────────────────────
// One borderless single-cell table, Mono paragraphs inside. Never `\n`.
function codePanel(lines) {
  return new Table({
    columnWidths: [H.CONTENT_W],
    width: { size: H.CONTENT_W, type: WidthType.DXA },
    borders: H.borderless,
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: H.CONTENT_W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: H.C.panel },
        margins: H.cellMargin,
        children: lines.map((ln, idx) => new Paragraph({
          style: 'Mono',
          children: [new TextRun(ln === '' ? ' ' : ln)],
          spacing: { after: idx === lines.length - 1 ? 0 : 20 },
        })),
      })],
    })],
  });
}

// ── label / value block ─────────────────────────────────────────────────────
// Borderless two-column table so the values share a true left edge. Never tab
// stops. pairs = [['Date', [t('August 19, 2026')]], ...]
function metaBlock(pairs, { labelWidth = 1700 } = {}) {
  const wL = labelWidth, wR = H.CONTENT_W - labelWidth;
  return new Table({
    columnWidths: [wL, wR],
    width: { size: H.CONTENT_W, type: WidthType.DXA },
    borders: H.borderless,
    rows: pairs.map(([k, v]) => new TableRow({
      children: [
        new TableCell({
          width: { size: wL, type: WidthType.DXA },
          margins: { top: 20, bottom: 20, left: 0, right: 120 },
          children: [new Paragraph({ style: 'Meta', spacing: { after: 20 }, children: [new TextRun({ text: k, bold: true })] })],
        }),
        new TableCell({
          width: { size: wR, type: WidthType.DXA },
          margins: { top: 20, bottom: 20, left: 0, right: 0 },
          children: [new Paragraph({ style: 'Meta', spacing: { after: 20 }, children: typeof v === 'string' ? [t(v)] : v })],
        }),
      ],
    })),
  });
}

// ── bordered results table with a shaded header row ─────────────────────────
// header = ['Col A', 'Col B']; rows = [['x', [t('y')]], ...].
// widths MUST sum to H.CONTENT_W (9360); omit to split evenly.
function resultsTable(header, rows, widths) {
  if (!widths) {
    const w = Math.floor(H.CONTENT_W / header.length);
    widths = header.map((_, idx) => (idx === header.length - 1 ? H.CONTENT_W - w * (header.length - 1) : w));
  }
  const sum = widths.reduce((a, x) => a + x, 0);
  if (sum !== H.CONTENT_W) {
    throw new Error('resultsTable: column widths must sum to ' + H.CONTENT_W + ', got ' + sum);
  }
  const border = { style: BorderStyle.SINGLE, size: 4, color: H.C.hairline };
  const borders = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };
  const mkRow = (cells, isHeader) => new TableRow({
    tableHeader: isHeader,
    children: cells.map((cell, idx) => new TableCell({
      width: { size: widths[idx], type: WidthType.DXA },
      margins: H.cellMargin,
      shading: isHeader ? { type: ShadingType.CLEAR, fill: H.C.panel } : undefined,
      children: [new Paragraph({
        spacing: { after: 0 },
        children: (typeof cell === 'string' ? [t(cell, { bold: isHeader })] : cell),
      })],
    })),
  });
  return new Table({
    columnWidths: widths,
    width: { size: H.CONTENT_W, type: WidthType.DXA },
    borders,
    rows: [mkRow(header, true), ...rows.map(r => mkRow(r, false))],
  });
}

// ── sign-off ────────────────────────────────────────────────────────────────
// keepNext on every line, or it orphans onto a page of its own.
function signOff(name, closing = 'Best,') {
  return [
    P(closing, { keepNext: true, spacing: { after: 0 } }),
    P(name, { keepNext: true, spacing: { after: 240 } }),
  ];
}

module.exports = { t, b, i, m, P, H1, H2, H3, codePanel, metaBlock, resultsTable, signOff, bullet: H.bullet, rule: H.rule };
