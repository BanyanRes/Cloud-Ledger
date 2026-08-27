// Patch 4b: colHeaders — make underline opt-in (default OFF) so balance-sheet
// dates are not underlined.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + JSON.stringify(o.slice(0, 200)));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, () => n); applied++; console.log('ok:', label);
}

replace('colHeaders signature',
`    colHeaders(labels) {\n      ensure(18);`,
`    colHeaders(labels, hopts = {}) {\n      ensure(18);`);

replace('colHeaders underline block',
`        // underline just under this header cell, spanning the widest line of it
        const uy = y - (nLines - 1) * 9 - 3;
        page.drawLine({ start: { x: cols[i] - maxW, y: uy }, end: { x: cols[i], y: uy }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });`,
`        // Underline each header cell only when explicitly requested. Per the
        // round-2 feedback the balance-sheet dates should NOT be underlined, so
        // the default is no underline.
        if (hopts.underline) {
          const uy = y - (nLines - 1) * 9 - 3;
          page.drawLine({ start: { x: cols[i] - maxW, y: uy }, end: { x: cols[i], y: uy }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
        }`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nPATCH4b APPLIED', applied, 'edits.');
