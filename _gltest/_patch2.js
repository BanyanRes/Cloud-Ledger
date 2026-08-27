// Node patcher #2 for financials.js — renderer & cover/TOC changes.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr, opts = {}) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) { if (opts.optional) { console.log('SKIP:', label); return; } throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + o.slice(0, 240)); }
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, n); applied++; console.log('ok:', label);
}

// ── A. colHeaders: underline ONLY the date/label cells, not a full-width rule.
// Also accept an optional { underlineDates } to control behavior; default keeps
// the per-cell underline (used by all statements) which reads cleaner.
replace('colHeaders underline-only-cells',
`    // Column headers (right-aligned above each numeric column).
    colHeaders(labels) {
      ensure(18);
      labels.forEach((lab, i) => {
        const parts = String(lab).split('\\n');
        parts.forEach((pl, pi) => {
          const w = bold.widthOfTextAtSize(pl, FS.head);
          page.drawText(pl, { x: cols[i] - w, y: y - pi * 9, size: FS.head, font: bold });
        });
      });
      y -= 9 * Math.max(1, ...labels.map(l => String(l).split('\\n').length));
      y -= 3;
      // thin rule under headers
      page.drawLine({ start: { x: PAGE.mL, y }, end: { x: PAGE.mR ? PAGE.w - PAGE.mR : PAGE.w, y }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
      y -= 10;
    },`,
`    // Column headers (right-aligned above each numeric column). Only the header
    // cells themselves are underlined (per the reference), not a full-width rule.
    colHeaders(labels) {
      ensure(18);
      const nLines = Math.max(1, ...labels.map(l => String(l).split('\\n').length));
      labels.forEach((lab, i) => {
        const parts = String(lab).split('\\n');
        let maxW = 0;
        parts.forEach((pl, pi) => {
          const w = bold.widthOfTextAtSize(pl, FS.head);
          if (w > maxW) maxW = w;
          page.drawText(pl, { x: cols[i] - w, y: y - pi * 9, size: FS.head, font: bold });
        });
        // underline just under this header cell, spanning the widest line of it
        const uy = y - (nLines - 1) * 9 - 3;
        page.drawLine({ start: { x: cols[i] - maxW, y: uy }, end: { x: cols[i], y: uy }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
      });
      y -= 9 * nLines;
      y -= 3;
      // Blank space between the underlined date headers and the first data line.
      y -= 12;
    },`);

// ── B. Balance-sheet subline: show "April 30, 2026 and March 31, 2026".
// The generic centered subline can't underline only the dates, so the BS block
// draws its own centered line via a small helper on the layout. Add that helper
// (drawCenteredDates) right after setSubline, then use it in the BS block.
replace('add drawCenteredDates helper',
`    _subline: null,
    setSubline(s) { this._subline = s; },
    start() { newPage(); },`,
`    _subline: null,
    setSubline(s) { this._subline = s; },
    // Draw a centered "<d1> and <d2>" line at the current cursor with ONLY the
    // two dates underlined. Used by the Balance Sheets page.
    drawCenteredDates(d1, d2) {
      const { reg: rf } = fonts;
      const sz = FS.sub;
      const conj = ' and ';
      const w1 = rf.widthOfTextAtSize(d1, sz);
      const wc = rf.widthOfTextAtSize(conj, sz);
      const w2 = rf.widthOfTextAtSize(d2, sz);
      const total = w1 + wc + w2;
      let x = (PAGE.w - total) / 2;
      page.drawText(d1, { x, y, size: sz, font: rf });
      page.drawLine({ start: { x, y: y - 2 }, end: { x: x + w1, y: y - 2 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
      x += w1;
      page.drawText(conj, { x, y, size: sz, font: rf });
      x += wc;
      page.drawText(d2, { x, y, size: sz, font: rf });
      page.drawLine({ start: { x, y: y - 2 }, end: { x: x + w2, y: y - 2 }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
      y -= 16; // blank space between the dates and the first line
    },
    start() { newPage(); },`);

// Use the two-date underlined subline on the Balance Sheets page (and add a
// space before the first line). Drop the generic parenthetical subline.
replace('BS two-date subline',
`    const L = makeLayout(pdf, fonts, m, 'Balance Sheets');
    L.setSubline(m.longDate + '   (with comparative totals as of ' + m.priorLongDate + ')');
    L.start();
    L.setCols(twoCols);
    L.colHeaders([m.longDate, m.priorLongDate]);
    L.sectionTitle('ASSETS');`,
`    const L = makeLayout(pdf, fonts, m, 'Balance Sheets');
    L.start();
    // Centered "<current date> and <prior date>" with only the dates underlined,
    // then a blank space before the columns/first line.
    L.drawCenteredDates(m.longDate, m.priorLongDate);
    L.space(4);
    L.setCols(twoCols);
    L.colHeaders([m.longDate, m.priorLongDate]);
    L.sectionTitle('ASSETS');`);

// The header's own centered subline should be blank on the BS page (we draw our
// own). setSubline defaults to null already; nothing else to change.

fs.writeFileSync(P, src, 'utf8');
console.log('\nAPPLIED', applied, 'edits.');
