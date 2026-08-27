// Patch 6: row() gains dollarPrefix support (draws "$" at the left of each
// numeric cell, value right-aligned) and column-width-aware rule spans so the
// landscape Members' Equity page underlines/rules look right.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + JSON.stringify(o.slice(0, 220)));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, () => n); applied++; console.log('ok:', label);
}

replace('row() dollarPrefix + rule spans',
`    row(label, cells, { indent = 12, boldRow = false, ruleAbove = false, ruleBelow = false, doubleBelow = false, gapAfter = 0 } = {}) {
      ensure(13);
      const font = boldRow ? bold : reg;
      if (ruleAbove) { page.drawLine({ start: { x: cols[0] - 78, y: y + 9 }, end: { x: cols[cols.length - 1], y: y + 9 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) }); }
      page.drawText(String(label), { x: PAGE.mL + indent, y, size: FS.row, font });
      cells.forEach((c, i) => {
        if (c == null || c === '') return;
        const w = font.widthOfTextAtSize(String(c), FS.row);
        page.drawText(String(c), { x: cols[i] - w, y, size: FS.row, font });
      });
      if (ruleBelow) { page.drawLine({ start: { x: cols[0] - 78, y: y - 3 }, end: { x: cols[cols.length - 1], y: y - 3 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) }); }
      if (doubleBelow) {
        page.drawLine({ start: { x: cols[0] - 78, y: y - 3 }, end: { x: cols[cols.length - 1], y: y - 3 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
        page.drawLine({ start: { x: cols[0] - 78, y: y - 5 }, end: { x: cols[cols.length - 1], y: y - 5 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
      }
      y -= 12 + gapAfter;
    },`,
`    row(label, cells, { indent = 12, boldRow = false, ruleAbove = false, ruleBelow = false, doubleBelow = false, gapAfter = 0, dollarPrefix = false } = {}) {
      ensure(13);
      const font = boldRow ? bold : reg;
      // Per-numeric-column width: distance from the previous column's right edge
      // (or a default) to this column's right edge. Used to place a left "$".
      const colWidth = (i) => (i === 0 ? 78 : Math.max(40, cols[i] - cols[i - 1]));
      const ruleLeft = cols[0] - colWidth(0) + 2;
      const ruleRight = cols[cols.length - 1];
      if (ruleAbove) { page.drawLine({ start: { x: ruleLeft, y: y + 9 }, end: { x: ruleRight, y: y + 9 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) }); }
      page.drawText(String(label), { x: PAGE.mL + indent, y, size: FS.row, font });
      cells.forEach((c, i) => {
        if (c == null || c === '') return;
        const s = String(c);
        const w = font.widthOfTextAtSize(s, FS.row);
        page.drawText(s, { x: cols[i] - w, y, size: FS.row, font });
        if (dollarPrefix) {
          // "$" anchored at the left of this column's cell box.
          const dx = cols[i] - colWidth(i) + 2;
          page.drawText('$', { x: dx, y, size: FS.row, font });
        }
      });
      if (ruleBelow) { page.drawLine({ start: { x: ruleLeft, y: y - 3 }, end: { x: ruleRight, y: y - 3 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) }); }
      if (doubleBelow) {
        page.drawLine({ start: { x: ruleLeft, y: y - 3 }, end: { x: ruleRight, y: y - 3 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
        page.drawLine({ start: { x: ruleLeft, y: y - 5 }, end: { x: ruleRight, y: y - 5 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
      }
      y -= 12 + gapAfter;
    },`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nPATCH6 APPLIED', applied, 'edits.');
