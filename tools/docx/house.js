// CloudLedger — Word document house style.
//
// Source of truth for typography. Do not re-derive any of this per document, and
// do not hand-roll a .docx without it. See ./README.md for the one delivery rule.
const {
  Paragraph, TextRun, Footer, AlignmentType, LevelFormat, BorderStyle,
  PageNumber, convertInchesToTwip,
} = require('docx');

// US Letter. docx-js defaults to A4 — always set this.
const PAGE = {
  size: { width: 12240, height: 15840 },
  margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
};
const CONTENT_W = 9360; // content width in DXA at 1" margins

const C = {
  body: '1A1A1A',
  accent: '1F3864',
  muted: '5A5A5A',
  hairline: 'C9CFD8',
  panel: 'F4F6F8',
};

const styles = {
  default: {
    document: {
      run: { font: 'Calibri', size: 22, color: C.body },
      paragraph: { spacing: { line: 276, lineRule: 'auto', after: 160 } },
    },
  },
  paragraphStyles: [
    { id: 'Title1', name: 'Title1', basedOn: 'Normal', next: 'Normal',
      run: { font: 'Calibri', size: 32, bold: true, color: C.accent },
      paragraph: { spacing: { line: 276, lineRule: 'auto', after: 120 } } },
    { id: 'H2', name: 'H2', basedOn: 'Normal', next: 'Normal',
      run: { font: 'Calibri', size: 23, bold: true, color: C.accent },
      paragraph: { spacing: { before: 280, after: 100, line: 276, lineRule: 'auto' }, keepNext: true } },
    { id: 'H3', name: 'H3', basedOn: 'Normal', next: 'Normal',
      run: { font: 'Calibri', size: 22, bold: true, color: C.body },
      paragraph: { spacing: { before: 200, after: 80, line: 276, lineRule: 'auto' }, keepNext: true } },
    { id: 'Meta', name: 'Meta', basedOn: 'Normal', next: 'Normal',
      run: { font: 'Calibri', size: 20, color: C.muted },
      paragraph: { spacing: { after: 60, line: 276, lineRule: 'auto' } } },
    { id: 'Mono', name: 'Mono', basedOn: 'Normal', next: 'Mono',
      run: { font: 'Consolas', size: 19, color: C.body },
      paragraph: { spacing: { after: 0, line: 240, lineRule: 'auto' } } },
    // Footer text MUST be a named style: PAGE/NUMPAGES fields take their size from
    // the PARAGRAPH MARK, and run properties alone render mixed sizes in
    // soffice-generated PDFs.
    { id: 'FooterText', name: 'FooterText', basedOn: 'Normal', next: 'FooterText',
      run: { font: 'Calibri', size: 16, color: C.muted },
      paragraph: { spacing: { after: 0 }, alignment: AlignmentType.RIGHT } },
  ],
};

const numbering = {
  config: [{
    reference: 'house-bullets',
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: convertInchesToTwip(0.3), hanging: convertInchesToTwip(0.18) } } },
    }],
  }],
};

// children: array of TextRun (or a plain string)
function bullet(children, opts = {}) {
  const kids = typeof children === 'string' ? [new TextRun(children)] : children;
  return new Paragraph({
    children: kids,
    numbering: { reference: 'house-bullets', level: 0 },
    spacing: { after: 100 },
    ...opts,
  });
}

// A horizontal rule is a paragraph bottom border, never a table or underscores.
function rule() {
  return new Paragraph({
    spacing: { before: 60, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.hairline } },
  });
}

// "Page N of M", right-aligned, 8pt muted. Each fragment its own TextRun — mixing
// strings and PageNumber.* in one run drops properties on the field halves.
function footer() {
  return new Footer({
    children: [new Paragraph({
      style: 'FooterText',
      children: [
        new TextRun('Page '),
        new TextRun({ children: [PageNumber.CURRENT] }),
        new TextRun(' of '),
        new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
      ],
    })],
  });
}

const dxa = (n) => Math.round(n);
const cellMargin = { top: 80, bottom: 80, left: 120, right: 120 };
const borderless = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
  insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
};

module.exports = { PAGE, CONTENT_W, styles, numbering, C, bullet, rule, footer, dxa, cellMargin, borderless };
