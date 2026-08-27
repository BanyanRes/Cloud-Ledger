// Node patcher #3 for financials.js — cover page + separate TOC page.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + o.slice(0, 260));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, n); applied++; console.log('ok:', label);
}

// Rewrite the cover body: professional cover page (no TOC), then a SEPARATE
// Table of Contents page. Cover date = just the as-of long date (e.g.
// "April 30, 2026"). TOC labels updated: Balance Sheets (plural) and
// "Budget to Actual" instead of "Requisition Report".
replace('cover + separate TOC',
`  center(meta.entityName, 20, bold, 520);
  center('Financial Statements', 15, reg, 486);
  center(meta.periodLabel, 12, reg, 462);
  page.drawLine({ start: { x: 180, y: 448 }, end: { x: PAGE.w - 180, y: 448 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Table of Contents', 11, bold, 380);
  const toc = ['Executive Summary', 'Balance Sheet', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\\u2019 Equity', 'Requisition Report'];
  let ty = 358;
  toc.forEach(t => { center(t, 10, reg, ty); ty -= 20; });
  center('Prepared by CloudLedger \\u2014 ' + meta.longDate, 9, reg, 90, rgb(0.45, 0.45, 0.45));
  return await pdf.save();`,
`  // ── Cover page ────────────────────────────────────────────────────────────
  // Clean, professional cover: entity name, "Financial Statements", the as-of
  // date, framed by two thin rules. No table of contents here (it lives on its
  // own page that follows).
  center(meta.entityName, 22, bold, 512);
  page.drawLine({ start: { x: 150, y: 494 }, end: { x: PAGE.w - 150, y: 494 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Financial Statements', 15, reg, 470);
  center(meta.longDate, 12, reg, 448);
  page.drawLine({ start: { x: 150, y: 430 }, end: { x: PAGE.w - 150, y: 430 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });

  // ── Table of Contents page (separate) ────────────────────────────────────
  const toc2 = pdf.addPage([PAGE.w, PAGE.h]);
  const centerOn = (pg, str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    pg.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  centerOn(toc2, meta.entityName, 13, bold, PAGE.h - PAGE.mT + 10);
  centerOn(toc2, 'Table of Contents', 15, bold, PAGE.h - 150);
  toc2.drawLine({ start: { x: 180, y: PAGE.h - 168 }, end: { x: PAGE.w - 180, y: PAGE.h - 168 }, thickness: 0.6, color: rgb(0.3, 0.3, 0.3) });
  const toc = ['Executive Summary', 'Balance Sheets', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\\u2019 Equity', 'Budget to Actual'];
  let ty = PAGE.h - 210;
  toc.forEach(t => { centerOn(toc2, t, 11, reg, ty); ty -= 26; });
  return await pdf.save();`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nAPPLIED', applied, 'edits.');
