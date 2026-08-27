// Node patcher for financials.js — applies Jimmy's requested report changes.
// Idempotent-ish: each replace asserts its anchor exists exactly once.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';

let applied = 0;
function replace(label, oldStr, newStr, opts = {}) {
  const count = src.split(oldStr).length - 1;
  if (count === 0) {
    if (opts.optional) { console.log('SKIP (not found, optional):', label); return; }
    throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + oldStr.slice(0, 200));
  }
  if (count > 1 && !opts.all) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = opts.all ? src.split(oldStr).join(newStr) : src.replace(oldStr, newStr);
  applied++;
  console.log('ok:', label + (opts.all ? ' (x' + count + ')' : ''));
}
// normalize a JS snippet's line endings to the file's EOL
const E = s => s.replace(/\n/g, EOL);

// ── 1. Entity display-name helper (insert after the pdf-lib require block) ──
replace('entity display-name helper',
E(`const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { xlsxSheetToPdf, looksLikeXlsx } = require('./xlsxToPdf');

// \u2500\u2500 numeric helpers`),
E(`const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { xlsxSheetToPdf, looksLikeXlsx } = require('./xlsxToPdf');

// \u2500\u2500 entity display-name normalization \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
// The report presents this development entity as "County Line SRN". The GL
// entity record is named "Sabine River & Northern Railroad" (entity 37) and is
// NOT renamed \u2014 only the statement package uses the display name. Map known
// aliases; anything else passes through unchanged.
const ENTITY_DISPLAY_NAMES = [
  { match: /sabine|(county\\s*line\\s*)?srn/i, display: 'County Line SRN' },
];
function displayEntityName(name) {
  const n = String(name || '').trim();
  for (const { match, display } of ENTITY_DISPLAY_NAMES) if (match.test(n)) return display;
  return n || 'Entity';
}

// \u2500\u2500 numeric helpers`));

// ── 2. Apply the display name in the meta object ──
replace('use display name in meta',
`entityName: opts.entityName || 'Entity', asOf,`,
`entityName: displayEntityName(opts.entityName), asOf,`);

// ── 3. Balance-sheet MAP: fix groupings to mirror the reference exactly ──
// Replace the whole BS_ACCOUNT_MAP object with the corrected reference mapping.
const oldMapStart = `const BS_ACCOUNT_MAP = {`;
const oldMapEnd = `  '39000': ['Members Equity', 'Retained Earnings'],\n};`;
{
  const si = src.indexOf(oldMapStart);
  const eiRaw = src.indexOf(E(oldMapEnd));
  if (si < 0 || eiRaw < 0) throw new Error('BS_ACCOUNT_MAP anchors not found');
  const ei = eiRaw + E(oldMapEnd).length;
  const before = src.slice(0, si);
  const after = src.slice(ei);
  const newMap = E(`const BS_ACCOUNT_MAP = {
  // Current Assets
  //   Cash and Cash Equivalents
  '10162': ['Current Assets', 'Cash and Cash Equivalents'],
  '10163': ['Current Assets', 'Cash and Cash Equivalents'],
  //   Accounts Receivable, Net
  '12000': ['Current Assets', 'Accounts Receivable, Net'],
  //   Intercompany Receivable
  '18311': ['Current Assets', 'Intercompany Receivable'],
  //   Other Current Assets
  '13001': ['Current Assets', 'Other Current Assets'],
  '13100': ['Current Assets', 'Other Current Assets'],
  '18002': ['Current Assets', 'Other Current Assets'],
  // Fixed Assets, Net
  '15100': ['Fixed Assets, Net', 'Fixed Assets'],
  '15165': ['Fixed Assets, Net', 'Fixed Assets'],
  // Intangible Assets, Net  (Intangible Assets less accumulated Amortization contra)
  '11009': ['Intangible Assets, Net', 'Intangible Assets'],
  '11013': ['Intangible Assets, Net', 'Intangible Assets'],
  '11014': ['Intangible Assets, Net', 'Intangible Assets'],
  '11016': ['Intangible Assets, Net', 'Intangible Assets'],
  '16600': ['Intangible Assets, Net', 'Amortization'],
  // Investments \u2192 Long Term Investments
  '11011': ['Investments', 'Long Term Investments'],
  '11012': ['Investments', 'Long Term Investments'],
  '12021': ['Investments', 'Long Term Investments'],
  // Other Assets (capitalized development costs, etc. \u2014 mirrors reference list)
  '11713': ['Other Assets', 'Other Assets'],
  '11760': ['Other Assets', 'Other Assets'],
  '11920': ['Other Assets', 'Other Assets'],
  '12115': ['Other Assets', 'Other Assets'],
  '12127': ['Other Assets', 'Other Assets'],
  '12230': ['Other Assets', 'Other Assets'],
  '12315': ['Other Assets', 'Other Assets'],
  '12325': ['Other Assets', 'Other Assets'],
  '12343': ['Other Assets', 'Other Assets'],
  '12364': ['Other Assets', 'Other Assets'],
  '12420': ['Other Assets', 'Other Assets'],
  '12423': ['Other Assets', 'Other Assets'],
  '12596': ['Other Assets', 'Other Assets'],
  '12600': ['Other Assets', 'Other Assets'],
  '12720': ['Other Assets', 'Other Assets'],
  '12913': ['Other Assets', 'Other Assets'],
  '15160': ['Other Assets', 'Other Assets'],
  // Current Liabilities
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '21006': ['Current Liabilities', 'Other Current Liabilities'],
  '21110': ['Current Liabilities', 'Other Current Liabilities'],
  // Long Term Liabilities \u2192 Loans
  '25063': ['Long Term Liabilities', 'Loans'],
  // Members Equity (Retained Earnings & Net Income handled specially in renderer)
  '34006': ['Members Equity', 'Members Equity'],
  '34014': ['Members Equity', 'Members Equity'],
  '34165': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};`);
  src = before + newMap + after;
  applied++;
  console.log('ok: rewrote BS_ACCOUNT_MAP to reference groupings');
}

// ── 4. Statement title: "Balance Sheet" -> "Balance Sheets" everywhere ──
replace('BS title -> Balance Sheets',
`makeLayout(pdf, fonts, m, 'Balance Sheet');`,
`makeLayout(pdf, fonts, m, 'Balance Sheets');`);

// ── 5. Footer: county line SRN, remove "accompanying notes" ──
replace('footer label',
E(`  function drawFooter() {
    const label = meta.entityName + ', ' + meta.longDate + '  |  See Executive Summary and accompanying notes';
    page.drawText(label, { x: PAGE.mL, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }`),
E(`  function drawFooter() {
    // Footnote: entity display name only \u2014 no "accompanying notes" text.
    const label = meta.entityName + ', ' + meta.longDate;
    page.drawText(label, { x: PAGE.mL, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }`));

fs.writeFileSync(P, src, 'utf8');
console.log('\\nAPPLIED', applied, 'edits. EOL=' + JSON.stringify(EOL));
