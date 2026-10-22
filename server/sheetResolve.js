// sheetResolve - the single source of truth for locating a worksheet by its
// canonical name across the per-project requisition / lender workbooks, which
// name their tabs inconsistently ("Budget to Actual", "Buna Budget",
// "CLIP Phase 3 Budget", "Budget to Actual (OP)", phase-suffixed, etc.).
//
// Both the requisition engine (requisition_rollforward.findSheet, reading an
// ExcelJS workbook) and the financial-statement package embed
// (xlsxToPdf.xlsxSheetToPdf, reading a SheetJS workbook) resolve tabs through
// resolveSheetName, so they always find the same tab the same way.
'use strict';

const SHEET_ALIASES = {
  'prior invoice log': ['prior inv log', 'prior log', 'prior invoices'],
  'current invoice log': ['current inv log', 'current log', 'current invoices'],
  'budget to actual': ['b2a', 'budget vs actual', 'budget-to-actual', 'budget to actuals', 'buna budget'],
  'dev fee': ['development fee', 'dev fee calc', 'developer fee', 'development fee calculation', 'development fee calc'],
  'hard cost contingency table': ['hard cost contingency', 'hard contingency table', 'hard cost contingency tbl'],
  'soft cost contingency table': ['soft cost contingency', 'soft contingency table', 'soft cost contingency tbl'],
};

const norm = (s) => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
// Strip a trailing phase suffix (" P2", "-P1", " Phase 2") so Silsbee-style
// templates, which tag every tab with its phase, still resolve to the canonical
// name. Requires a separator before the token so mid-word "p" (e.g. "group 2")
// is not affected.
const stripPhase = (s) => s.replace(/[\s-]+(?:p|phase)\s*\d+$/i, '').trim();

// Resolve a canonical tab name against a list of actual sheet names. Returns the
// matching name, or undefined when none matches. This is the exact matching
// requisition_rollforward.findSheet used to carry inline - kept here verbatim so
// both callers stay in lockstep.
function resolveSheetName(names, canonicalName) {
  const list = Array.isArray(names) ? names : [];
  const want = norm(canonicalName);
  // 1. exact (case-sensitive) - mirrors ExcelJS workbook.getWorksheet(name).
  for (const n of list) { if (n === canonicalName) return n; }
  // 2. normalized exact + 3. aliases (also with the phase suffix removed)
  const aliases = SHEET_ALIASES[want] || [];
  for (const n of list) {
    const nn = norm(n);
    const nb = stripPhase(nn);
    if (nn === want || aliases.includes(nn) || nb === want || aliases.includes(nb)) return n;
  }
  // 4. Fuzzy fallback, Budget-to-Actual only. CPAs name this one tab
  //    inconsistently across project templates - "CLIP Phase 3 Budget",
  //    "Budget to Actual 2B", "Budget to Actual (OP)", "Buna Budget" - often
  //    with the phase token embedded mid-name (not a strippable suffix), so the
  //    passes above miss it. In a requisition workbook the only tab carrying the
  //    word "budget" is the B2A report (the invoice logs, contingency tables and
  //    the dev-fee tab never do), so resolve on that, preferring the name closest
  //    to the canonical. Scoped to this canonical so every other tab stays strict.
  if (want === 'budget to actual') {
    const isOther = (n) => /invoice log|inv log|contingency|dev(?:elopment)?\s*fee/.test(n);
    let best, bestScore = 0;
    for (const n of list) {
      const nn = norm(n);
      if (!/\bbudget\b/.test(nn) || isOther(nn)) continue;
      const score = /budget\s*(?:to|vs\.?|-)?\s*actual/.test(nn) ? 3 : (/\bactual\b/.test(nn) ? 2 : 1);
      if (score > bestScore) { best = n; bestScore = score; }
    }
    if (best) return best;
  }
  return undefined;
}

module.exports = { SHEET_ALIASES, resolveSheetName };
