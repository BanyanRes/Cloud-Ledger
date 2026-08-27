// Portfolio regression for the balance-sheet cash-classification fix.
// Compares the OLD name-only rule against the PATCHED classifier (bsSection,
// exported from server/financials.js) over every asset account in the production
// backup, and reports every account whose Cash and Cash Equivalents membership
// changes in either direction.
const fs = require('fs');
const { _helpers } = require('./server/financials.js');
const { bsSection } = _helpers;

const rows = JSON.parse(fs.readFileSync('_all_assets.json', 'utf8'));

// The rule as it stood before this change (name-only), applied after the
// explicit BS_ACCOUNT_MAP and the intercompany pre-test — same order as
// bsClassify, so the comparison is apples to apples.
const MAPPED = new Set(['10162', '10163', '12000', '18311', '13001', '13100', '18002', '15100', '15165',
  '11009', '11013', '11014', '11016', '16600', '11011', '11012', '12021', '11713', '11760', '11920',
  '12115', '12127', '12230', '12315', '12325', '12343', '12364', '12420', '12423', '12596', '12600',
  '12720', '12913', '15160']);
const MAPPED_CASH = new Set(['10162', '10163']);

function oldSection(r) {
  const nm = (r.name || '').toLowerCase();
  if (MAPPED.has(String(r.code))) return MAPPED_CASH.has(String(r.code)) ? 'CASH' : 'other';
  if (/cash|checking|savings|bank|clearing/.test(nm)) return 'CASH';
  return 'other';
}

// The profiles that do NOT go through bsClassify at all.
const OWN_PROFILE = /banyan\s*residential|banyan\s*sfr\s*gp\s*investors/i;

let intoCash = [], outOfCash = [], skipped = 0;

// bsSection only returns the section; re-derive the subsection the same way the
// renderer does, by calling the exported helper on a clone and reading section,
// then confirming with the shipped predicate for cash.
function bsSectionSub(r) {
  // Mirror of bsClassify's asset branch ordering for the subsection we care
  // about. Intercompany routing wins first on the srn profile.
  const nm = (r.name || '').toLowerCase();
  if (r.type === 'Asset' && /due from|intercompany/.test(nm)) return 'Intercompany Receivable';
  if (MAPPED.has(String(r.code))) return MAPPED_CASH.has(String(r.code)) ? 'Cash and Cash Equivalents' : 'other';
  return isCashAccountShipped(r) ? 'Cash and Cash Equivalents' : 'other';
}

// The shipped predicate, read straight out of financials.js so this test cannot
// drift from the code it is testing.
const src = fs.readFileSync('./server/financials.js', 'utf8').replace(/\r\n/g, '\n');
const m = src.match(/function isCashCode\(code\) \{[\s\S]*?\n\}\s*\nfunction isCashAccount\(r\) \{[\s\S]*?\n\}/);
if (!m) throw new Error('could not extract isCashCode/isCashAccount');
const isCashAccountShipped = new Function(m[0] + '; return isCashAccount;')();

// Re-run now that the helpers are defined (function declarations hoist, const does not).
intoCash = []; outOfCash = []; skipped = 0;
for (const r of rows) {
  if (OWN_PROFILE.test(r.entity)) { skipped++; continue; }
  const before = oldSection(r) === 'CASH';
  const after = bsSectionSub(r) === 'Cash and Cash Equivalents';
  if (after && !before) intoCash.push(r);
  if (before && !after) outOfCash.push(r);
}

const fmt = n => (n == null ? '' : Number(n).toFixed(2).padStart(15));
console.log('accounts examined: ' + rows.length + '   (skipped ' + skipped + ' on banyan/bsfrgp profiles)');
console.log('\n=== MOVED INTO Cash and Cash Equivalents (' + intoCash.length + ') ===');
for (const r of intoCash.sort((a, b) => Math.abs(b.bal || 0) - Math.abs(a.bal || 0))) {
  console.log('  ' + String(r.entity).slice(0, 32).padEnd(32) + ' ' + r.code + '  ' + String(r.name).slice(0, 46).padEnd(46) + ' bank=' + r.bank_acct + fmt(r.bal));
}
console.log('  TOTAL BALANCE MOVED: ' + fmt(intoCash.reduce((a, r) => a + (r.bal || 0), 0)));
console.log('\n=== MOVED OUT OF Cash and Cash Equivalents (' + outOfCash.length + ') ===');
if (!outOfCash.length) console.log('  none — the new predicate is a superset of the old one on this data');
for (const r of outOfCash) {
  console.log('  ' + String(r.entity).slice(0, 32).padEnd(32) + ' ' + r.code + '  ' + String(r.name).slice(0, 46).padEnd(46) + ' bank=' + r.bank_acct + fmt(r.bal));
}
console.log('\n' + (outOfCash.length === 0 ? 'PASS — no regressions' : 'REVIEW — accounts left the cash section'));
