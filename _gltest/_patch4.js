// Patch 4: layout infra — landscape support, repeated date-line heading,
// centered footer on all pages. (colHeaders underline handled in patch 4b.)
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + JSON.stringify(o.slice(0, 300)));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, () => n); applied++; console.log('ok:', label);
}

replace('makeLayout signature + page geometry',
`function makeLayout(pdf, fonts, meta, statementTitle) {
  const { reg, bold } = fonts;
  let page, y;
  const cols = []; // right edges for numeric columns, set per statement

  function newPage() {
    page = pdf.addPage([PAGE.w, PAGE.h]);
    y = PAGE.h - PAGE.mT;
    drawHeader();
    drawFooter();
    y -= 6;
  }
  function textC(str, size, font, yy) {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font });
  }
  function drawHeader() {
    textC(meta.entityName, FS.title, bold, PAGE.h - PAGE.mT + 22);
    textC(statementTitle, FS.sub, bold, PAGE.h - PAGE.mT + 10);
    // Sub-line: as-of date (BS) or period label — set by each statement via setSubline.
    if (layout._subline) textC(layout._subline, FS.sub, reg, PAGE.h - PAGE.mT - 2);
  }
  function drawFooter() {
    // Footnote: entity display name only — no "accompanying notes" text.
    const label = meta.entityName + ', ' + meta.longDate;
    page.drawText(label, { x: PAGE.mL, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }
  function ensure(space) { if (y - space < PAGE.mB + 8) newPage(); }`,
`function makeLayout(pdf, fonts, meta, statementTitle, opts = {}) {
  const { reg, bold } = fonts;
  // Page geometry: portrait by default, landscape when requested (opts.landscape).
  const PW = opts.landscape ? PAGE.h : PAGE.w;
  const PH = opts.landscape ? PAGE.w : PAGE.h;
  const dateLine = opts.dateLine || null; // repeated heading date-line (every page)
  let page, y;
  const cols = []; // right edges for numeric columns, set per statement

  function newPage() {
    page = pdf.addPage([PW, PH]);
    y = PH - PAGE.mT;
    drawHeader();
    drawFooter();
    // On every page: draw the repeated date-line heading + a blank space between
    // the heading and the first row (or the column headers).
    if (dateLine) { textC(dateLine, FS.sub, reg, PH - PAGE.mT - 2); y -= 14; }
    y -= 8;
  }
  function textC(str, size, font, yy) {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PW - w) / 2, y: yy, size, font });
  }
  function drawHeader() {
    textC(meta.entityName, FS.title, bold, PH - PAGE.mT + 22);
    textC(statementTitle, FS.sub, bold, PH - PAGE.mT + 10);
    // Optional extra sub-line set by a statement (rarely used now).
    if (layout._subline) textC(layout._subline, FS.sub, reg, PH - PAGE.mT - 2);
  }
  function drawFooter() {
    // Centered footer on every page: "<entity>, <date>   See Executive Summary".
    const label = meta.entityName + ', ' + meta.longDate + '   See Executive Summary';
    const w = reg.widthOfTextAtSize(label, FS.foot);
    page.drawText(label, { x: (PW - w) / 2, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }
  function ensure(space) { if (y - space < PAGE.mB + 8) newPage(); }`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nPATCH4 APPLIED', applied, 'edits.');
