// ═══════════════════════════════════════════════════════════════════════════
// financials.js — GL-derived financial-statement package generator.
//
// Produces the same statement set the accounting team hand-prepares for
// development entities (SRN-style): Balance Sheet, Statements of Operations,
// Statement of Cash Flows, and Statement of Changes in Members' Equity, then
// assembles a single merged PDF: cover → executive summary (uploaded) →
// GL statements (rendered here) → requisition report (uploaded).
//
// Design notes
// ------------
// * The four statements are built from balance snapshots produced by the SAME
//   query the /balances endpoint uses. We take an injected `getBalances`
//   function so this module stays testable with fixture data and reusable from
//   the live route. Signature: getBalances({ as_of, from, to, close_pl_before })
//   -> [{ code, name, type, subtype, balance, total_debit, total_credit }].
//
// * Retained-earnings vs. Net-income split (the one thing that differed from the
//   CPA's version): RE is shown FROZEN at the beginning-of-year opening balance,
//   and the full current-year YTD P&L is shown on a single Net Income (Loss)
//   line. We get the frozen RE by calling getBalances with
//   close_pl_before = Jan 1 of the statement year, then SUBTRACTING the YTD P&L
//   back out of RE so RE reflects only pre-year activity. This reproduces
//   Document 2's presentation (opening RE + YTD net income = ending equity) and
//   ties to the Statement of Changes in Members' Equity by construction.
//
// * Cash flow uses the indirect method over the YTD window, mirroring the
//   hand-prepared statement (single "For the N Months Ended" column).
// ═══════════════════════════════════════════════════════════════════════════

const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { xlsxSheetToPdf, looksLikeXlsx } = require('./xlsxToPdf');

// ── entity display-name normalization ───────────────────────────────
// The report presents this development entity as "County Line SRN". The GL
// entity record is named "Sabine River & Northern Railroad" (entity 37) and is
// NOT renamed — only the statement package uses the display name. Map known
// aliases; anything else passes through unchanged.
const ENTITY_DISPLAY_NAMES = [
  { match: /sabine|(county\s*line\s*)?srn/i, display: 'County Line SRN' },
];
function displayEntityName(name) {
  const n = String(name || '').trim();
  for (const { match, display } of ENTITY_DISPLAY_NAMES) if (match.test(n)) return display;
  return n || 'Entity';
}

// ── numeric helpers ────────────────────────────────────────────────────────
const r2 = n => Math.round((Number(n) || 0) * 100) / 100;
const isZero = n => Math.abs(Number(n) || 0) < 0.005;
// Accounting format: 1,234.56 ; negatives in parentheses; zero as a dash.
function acct(n, { dash = true, blankZero = false } = {}) {
  const v = r2(n);
  if (isZero(v)) return blankZero ? '' : (dash ? '-' : '0.00');
  const s = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? '(' + s + ')' : s;
}

// A balance row's "natural" signed balance (already computed by getBalances:
// Asset/Expense = debit-positive, others = credit-positive).
const bal = row => Number(row.balance) || 0;

// Sum a filtered set of balance rows.
const sumRows = (rows, pred) => r2(rows.filter(pred).reduce((s, r) => s + bal(r), 0));

// Net income implied by a set of balance rows (Revenue positive, Expense negative).
function netIncomeOf(rows) {
  let ni = 0;
  for (const r of rows) {
    if (r.type === 'Revenue') ni += bal(r);
    else if (r.type === 'Expense') ni -= bal(r);
  }
  return r2(ni);
}

// ── date helpers ───────────────────────────────────────────────────────────
function yearStart(asOf) { return String(asOf).slice(0, 4) + '-01-01'; }
function priorMonthEnd(asOf) {
  const d = new Date(asOf + 'T00:00:00Z');
  const first = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1));
  first.setUTCDate(0); // last day of previous month
  return first.toISOString().slice(0, 10);
}
function monthStart(asOf) { return String(asOf).slice(0, 7) + '-01'; }
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
function longDate(asOf) {
  const d = new Date(asOf + 'T00:00:00Z');
  return MONTHS[d.getUTCMonth()] + ' ' + d.getUTCDate() + ', ' + d.getUTCFullYear();
}
function monthsEndedLabel(asOf) {
  const m = parseInt(String(asOf).slice(5, 7), 10); // 1..12 → months elapsed YTD
  const word = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve'][m] || String(m);
  return 'For the ' + word + (m === 1 ? ' Month' : ' Months') + ' Ended ' + longDate(asOf);
}

// Add/subtract whole months from a month-END date, returning the month-end of
// the result. e.g. addMonthsEnd('2026-03-31', -1) → '2026-02-28'.
function addMonthsEnd(asOf, delta) {
  const d = new Date(asOf + 'T00:00:00Z');
  // Move to the first of this month, shift by delta+1 months, back up one day →
  // last day of the target month (handles 28/29/30/31 correctly).
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + delta + 1, 1));
  t.setUTCDate(0);
  return t.toISOString().slice(0, 10);
}
function dayBefore(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - 1);
  return d.toISOString().slice(0, 10);
}

// Resolve the P&L comparative windows and the balance-sheet comparative date for
// a given as-of date and period mode ('monthly' | 'quarterly' | 'annually').
//   cur  = the current period window [from,to]
//   pri  = the prior comparable period window [from,to]
//   bsPriorDate = the balance-sheet comparative column's as-of date
//   periodLabel = "For the Month/Quarter/Year Ended <date>"
//   colLabel    = short header for the P&L period columns
// The YTD column and cash-flow window are ALWAYS calendar-YTD (1/1 → asOf),
// independent of the period mode.
function resolvePeriod(asOf, period) {
  const p = (period || 'monthly').toLowerCase();
  if (p === 'quarterly') {
    const curFrom = addMonthsEnd(asOf, -2); // first day handled below
    const cur = { from: monthStart(addMonthsEnd(asOf, -2)), to: asOf };
    const priTo = dayBefore(cur.from);
    const pri = { from: monthStart(addMonthsEnd(priTo, -2)), to: priTo };
    return {
      cur, pri, bsPriorDate: priTo,
      periodLabel: 'For the Quarter Ended ' + longDate(asOf),
      colLabel: 'Quarter Ended',
    };
  }
  if (p === 'annually' || p === 'annual' || p === 'yearly') {
    const cur = { from: monthStart(addMonthsEnd(asOf, -11)), to: asOf };
    const priTo = dayBefore(cur.from);
    const pri = { from: monthStart(addMonthsEnd(priTo, -11)), to: priTo };
    return {
      cur, pri, bsPriorDate: priTo,
      periodLabel: 'For the Year Ended ' + longDate(asOf),
      colLabel: 'Year Ended',
    };
  }
  // monthly (default)
  const priorEnd = priorMonthEnd(asOf);
  return {
    cur: { from: monthStart(asOf), to: asOf },
    pri: { from: monthStart(priorEnd), to: priorEnd },
    bsPriorDate: priorEnd,
    periodLabel: 'For the Month Ended ' + longDate(asOf),
    colLabel: 'Month Ended',
  };
}

// Cash & cash-equivalents detection. Development-entity charts (SRN) name their
// operating bank accounts by account number (e.g. "SRNR x3505") with no "cash"
// in the name and no bank_acct flag set, so a name-only test misses them. The
// reliable signal on these QuickBooks-derived charts is the account-code block:
// 100xx–101xx are operating cash/bank and 107xx are Bill.com clearing accounts
// (cash-equivalent pass-throughs). We combine the code block with the name and
// bank_acct heuristics so any chart style is covered.
function isCashCode(code) {
  const c = String(code || '');
  return /^10[01]\d/.test(c) || /^107\d/.test(c);
}
function isCashAccount(r) {
  return r.type === 'Asset' && (isCashCode(r.code) || /cash|checking|savings|money market|operating acct|bank/i.test(r.name || '') || r.bank_acct);
}

// ── Balance-sheet classification ───────────────────────────────────────────
// The SRN chart has no populated `subtype`, and account codes do not fall into
// clean numeric ranges by section (e.g. 15100 Land and 15165 Railroad Track are
// Fixed Assets, but 15160 Railroad & Building Improvements is Other Assets).
// To reproduce the CPA package's exact groupings, we classify by an explicit
// two-level map: section → subsection. Each balance-sheet account is assigned a
// { section, sub } pair. Anything unmapped falls through to a name/code
// heuristic so the statement never silently drops an account.
//
// Sections (in presentation order) and their subsections mirror the reference:
//   Current Assets      → Cash and Cash Equivalents / Accounts Receivable, Net
//                         / Intercompany Receivable / Other Current Assets
//   Fixed Assets, Net   → Fixed Assets
//   Intangible Assets, Net → Intangible Assets / Amortization (contra)
//   Investments         → Long Term Investments
//   Other Assets        → Other Assets
//   Current Liabilities → Accounts Payable / Other Current Liabilities
//   Long Term Liabilities → Loans
//   Members Equity      → Members Equity / Retained Earnings
const BS_ACCOUNT_MAP = {
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
  // Investments → Long Term Investments
  '11011': ['Investments', 'Long Term Investments'],
  '11012': ['Investments', 'Long Term Investments'],
  '12021': ['Investments', 'Long Term Investments'],
  // Other Assets (capitalized development costs, etc. — mirrors reference list)
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
  // Long Term Liabilities → Loans
  '25063': ['Long Term Liabilities', 'Loans'],
  // Members Equity (Retained Earnings & Net Income handled specially in renderer)
  '34006': ['Members Equity', 'Members Equity'],
  '34014': ['Members Equity', 'Members Equity'],
  '34165': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};

// Contra accounts shown as a subtraction within their section (amortization).
const BS_CONTRA_CODES = new Set(['16600']);

function bsClassify(row) {
  const explicit = BS_ACCOUNT_MAP[String(row.code)];
  if (explicit) return { section: explicit[0], sub: explicit[1] };
  // Heuristic fallback for accounts not in the map, so nothing is dropped.
  const name = (row.name || '').toLowerCase();
  if (row.type === 'Asset') {
    if (/cash|checking|savings|bank|clearing/.test(name)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
    if (/receivable/.test(name) && /due from|intercompany/.test(name)) return { section: 'Current Assets', sub: 'Intercompany Receivable' };
    if (/receivable/.test(name)) return { section: 'Current Assets', sub: 'Accounts Receivable, Net' };
    if (/prepaid|reserve|deposit/.test(name)) return { section: 'Current Assets', sub: 'Other Current Assets' };
    return { section: 'Other Assets', sub: 'Other Assets' };
  }
  if (row.type === 'Liability') {
    if (/loan|note payable|bot|bond/.test(name)) return { section: 'Long Term Liabilities', sub: 'Loans' };
    if (/payable/.test(name)) return { section: 'Current Liabilities', sub: 'Accounts Payable' };
    return { section: 'Current Liabilities', sub: 'Other Current Liabilities' };
  }
  if (row.type === 'Equity') {
    if (/retained earning/.test(name)) return { section: 'Members Equity', sub: 'Retained Earnings' };
    return { section: 'Members Equity', sub: 'Members Equity' };
  }
  return { section: 'Other', sub: 'Other' };
}

// Back-compat shim: some callers/tests use bsSection(row) → section string.
function bsSection(row) { return bsClassify(row).section; }

// Presentation order for sections and, within each, their subsections.
const BS_ASSET_ORDER = ['Current Assets', 'Fixed Assets, Net', 'Intangible Assets, Net', 'Investments', 'Other Assets'];
const BS_LIAB_ORDER = ['Current Liabilities', 'Long Term Liabilities'];
const BS_SUB_ORDER = {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Intercompany Receivable', 'Other Current Assets'],
  'Fixed Assets, Net': ['Fixed Assets'],
  'Intangible Assets, Net': ['Intangible Assets', 'Amortization'],
  'Investments': ['Long Term Investments'],
  'Other Assets': ['Other Assets'],
  'Current Liabilities': ['Accounts Payable', 'Other Current Liabilities'],
  'Long Term Liabilities': ['Loans'],
};

// ── P&L operating-expense classification ────────────────────────────────────
// Per the CLR operating-expense restructure (Will Myers / Jimmy Yun, Jun 2026),
// the old broad P&L sections (G&A, Payroll, Utilities & Facilities, Taxes &
// Insurance) are replaced by 11 finer categories. NO GL accounts change — this
// is purely a re-grouping of how expense sub-lines roll up into subtotals, so
// the operating-expense grand total (and therefore net income) is unaffected.
//
// Management Fees is intentionally kept as its own unchanged category. Car Hire
// and other cost-of-revenue lines are handled separately as COGS and are not
// part of this map.
//
// PL_EXPENSE_MAP is an explicit code → category assignment built from the SRN
// chart; PL_EXPENSE_CATEGORY_ORDER fixes the presentation order. Any expense
// account not in the map falls through to a name heuristic so other CLR
// entities' charts still classify sensibly rather than dropping a line.
const PL_EXPENSE_MAP = {
  // Professional Services
  '63000': 'Professional Services',   // Accounting
  '63025': 'Professional Services',   // Professional Fees
  // Technology & Software
  '67300': 'Technology & Software',   // Telephone & Internet
  // Administrative & Other
  '60200': 'Administrative & Other',  // Payroll Processing Fee (bank/processing)
  '60210': 'Administrative & Other',  // Travel
  '60500': 'Administrative & Other',  // Meals
  '67100': 'Administrative & Other',  // Dues & Subscriptions
  '67150': 'Administrative & Other',  // Miscellaneous
  '67200': 'Administrative & Other',  // Office Expense
  '67400': 'Administrative & Other',  // Advertising & Marketing
  // Personnel / Payroll
  '60000': 'Personnel / Payroll',     // Salaries & Wages
  '60002': 'Personnel / Payroll',     // Payroll Taxes
  '60005': 'Personnel / Payroll',     // Health Insurance
  '60012': 'Personnel / Payroll',     // RRB Taxes - Employer Portion
  '63042': 'Personnel / Payroll',     // Offsite Staff
  // Track & Infrastructure
  '61050': 'Track & Infrastructure',  // Site/Yard Maintenance
  // Equipment & Rolling Stock
  '61000': 'Equipment & Rolling Stock', // Locomotive Rent
  '61005': 'Equipment & Rolling Stock', // Vehicle Rent
  '61053': 'Equipment & Rolling Stock', // Equipment Supplies
  '61054': 'Equipment & Rolling Stock', // Locomotive Repair
  // Fuel & Utilities
  '61150': 'Fuel & Utilities',        // Utilities
  '61152': 'Fuel & Utilities',        // Water
  '61164': 'Fuel & Utilities',        // Fuel
  // Contracted Services
  '61056': 'Contracted Services',     // Landscape Maintenance
  '61064': 'Contracted Services',     // Pest Control Services
  // Insurance
  '65000': 'Insurance',               // Insurance - Liability
  '68055': 'Insurance',               // Property Insurance
  // Taxes & Assessments
  '68000': 'Taxes & Assessments',     // Tax & License
  '68050': 'Taxes & Assessments',     // Property Tax
  '68060': 'Taxes & Assessments',     // State and Local Taxes
  // Regulatory & Compliance — new category; no SRN accounts yet
  // Management Fees (unchanged)
  '63041': 'Management Fees',         // CLRO Management Fees
};

// Presentation order for operating-expense categories (Management Fees last,
// kept separate and unchanged per the email).
const PL_EXPENSE_CATEGORY_ORDER = [
  'Professional Services',
  'Technology & Software',
  'Administrative & Other',
  'Personnel / Payroll',
  'Track & Infrastructure',
  'Equipment & Rolling Stock',
  'Fuel & Utilities',
  'Contracted Services',
  'Insurance',
  'Taxes & Assessments',
  'Regulatory & Compliance',
  'Management Fees',
];

// Classify an expense account into one of the 11 categories. Explicit map wins;
// otherwise a name heuristic keeps unmapped accounts (other CLR entities) from
// being dropped. Falls back to 'Administrative & Other' as a catch-all.
function plExpenseCategory(row) {
  const explicit = PL_EXPENSE_MAP[String(row.code)];
  if (explicit) return explicit;
  const name = (row.name || '').toLowerCase();
  if (/management fee/.test(name)) return 'Management Fees';
  if (/wage|salary|salaries|payroll|benefit|health insurance|rrb|offsite staff|overtime/.test(name)) return 'Personnel / Payroll';
  if (/accounting|legal|professional fee|engineering|consulting|environmental consult/.test(name)) return 'Professional Services';
  if (/software|subscription|telecom|telephone|internet/.test(name)) return 'Technology & Software';
  if (/fra|regulatory|compliance/.test(name)) return 'Regulatory & Compliance';
  if (/diesel|fuel|electric|water|utilit/.test(name)) return 'Fuel & Utilities';
  if (/locomotive|truck|vehicle|equipment|rolling stock/.test(name)) return 'Equipment & Rolling Stock';
  if (/track|crossing|site\/yard|yard maintenance|infrastructure/.test(name)) return 'Track & Infrastructure';
  if (/landscap|pest|contracted|janitor/.test(name)) return 'Contracted Services';
  if (/insurance/.test(name)) return 'Insurance';
  if (/property tax|state and local tax|tax & license|tax and license|assessment|other tax/.test(name)) return 'Taxes & Assessments';
  return 'Administrative & Other';
}


// ═══════════════════════════════════════════════════════════════════════════
// buildStatements — the numeric core. Pure given getBalances; no I/O.
//
// opts: { asOf, entityName, closeYtdIntoNI (default true) }
// Returns a structured object the PDF renderer consumes.
// ═══════════════════════════════════════════════════════════════════════════
async function buildStatements(getBalances, opts) {
  const asOf = opts.asOf;
  const ys = yearStart(asOf);
  const period = resolvePeriod(asOf, opts.period);
  const priorBsDate = period.bsPriorDate;

  // Snapshots:
  //  bsCur / bsPri — balance sheet as of period-end and the prior COMPARABLE
  //    period-end (prior month / quarter / year, per the toggle), with prior-year
  //    P&L closed into RE (close_pl_before = year start) so RE holds the opening
  //    balance and current-year P&L stays open on income accounts.
  //  isYtd — calendar-YTD P&L (always 1/1 → asOf), drives the YTD column and CF.
  //  isCur / isPri — P&L for the current and prior comparable PERIOD windows.
  const [bsCur, bsPri, isYtd, isCur, isPri] = await Promise.all([
    getBalances({ as_of: asOf, close_pl_before: ys }),
    getBalances({ as_of: priorBsDate, close_pl_before: yearStart(priorBsDate) }),
    getBalances({ from: ys, to: asOf }),
    getBalances({ from: period.cur.from, to: period.cur.to }),
    getBalances({ from: period.pri.from, to: period.pri.to }),
  ]);

  const niYtd = netIncomeOf(isYtd);
  // Prior BS column's net-income line = P&L for the prior year through the prior
  // comparative date (calendar-YTD basis relative to that column's own year).
  const niPriYtd = netIncomeOf(await getBalances({ from: yearStart(priorBsDate), to: priorBsDate }));

  // ── Balance Sheet ────────────────────────────────────────────────────────
  // Group asset/liability/equity rows for both columns keyed by account code.
  function bsColumn(rows) {
    const closedNI = netIncomeOf(rows); // rows still carry open current-year P&L
    // Split RE: freeze at opening balance by removing the YTD P&L that the
    // close_pl_before mechanism did NOT close (it only closed prior YEARS).
    const map = new Map();
    for (const r of rows) {
      if (r.type === 'Revenue' || r.type === 'Expense') continue; // P&L → Net Income line
      map.set(r.code, r);
    }
    return { map, ni: closedNI };
  }
  const colCur = bsColumn(bsCur);
  const colPri = bsColumn(bsPri);

  // Union of BS account codes across both columns, preserving code order.
  const bsCodes = Array.from(new Set([...colCur.map.keys(), ...colPri.map.keys()]))
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  // Rows for a single (section, subsection) pair, in account-code order. Contra
  // accounts (accumulated amortization) keep their natural sign; the renderer
  // subtracts the subsection total at the section level.
  function bsRowsForSub(section, sub, type) {
    return bsCodes
      .map(code => {
        const rc = colCur.map.get(code), rp = colPri.map.get(code);
        const ref = rc || rp;
        if (!ref || ref.type !== type) return null;
        const cls = bsClassify(ref);
        if (cls.section !== section || cls.sub !== sub) return null;
        const cur = rc ? bal(rc) : 0, pri = rp ? bal(rp) : 0;
        if (isZero(cur) && isZero(pri)) return null;
        return { code, name: ref.name, cur: r2(cur), pri: r2(pri), change: r2(cur - pri), contra: BS_CONTRA_CODES.has(String(code)) };
      })
      .filter(Boolean);
  }

  // Build a section as an ordered list of subsections, each with its own rows and
  // subtotal. A section's net total treats contra subsections as subtractions.
  function bsSectionFor(section, type) {
    const subOrder = BS_SUB_ORDER[section] || [];
    // Include any subsection that appears in the order list first, then any
    // unexpected subsections (defensive: never drop a classified account).
    const seen = new Set(subOrder);
    const extraSubs = [];
    for (const code of bsCodes) {
      const ref = colCur.map.get(code) || colPri.map.get(code);
      if (!ref || ref.type !== type) continue;
      const cls = bsClassify(ref);
      if (cls.section === section && !seen.has(cls.sub)) { seen.add(cls.sub); extraSubs.push(cls.sub); }
    }
    const subs = [...subOrder, ...extraSubs]
      .map(sub => {
        const rows = bsRowsForSub(section, sub, type);
        if (!rows.length) return null;
        const isContra = rows.every(r => r.contra);
        const subtotal = { cur: r2(rows.reduce((a, r) => a + r.cur, 0)), pri: r2(rows.reduce((a, r) => a + r.pri, 0)) };
        return { title: sub, rows, subtotal, contra: isContra };
      })
      .filter(Boolean);
    // Section net total. GL balances are already signed (contra accounts such
    // as accumulated amortization carry a natural negative balance), so we sum
    // subsection subtotals directly — no extra sign flip for contra.
    const total = subs.reduce((t, s) => ({
      cur: r2(t.cur + s.subtotal.cur),
      pri: r2(t.pri + s.subtotal.pri),
    }), { cur: 0, pri: 0 });
    return { title: section, subs, total };
  }

  const assetSections = BS_ASSET_ORDER
    .map(s => bsSectionFor(s, 'Asset'))
    .filter(s => s.subs.length);
  const liabSections = BS_LIAB_ORDER
    .map(s => bsSectionFor(s, 'Liability'))
    .filter(s => s.subs.length);

  // Equity: flat list of contributed-capital accounts (Members Equity subsection),
  // with Retained Earnings and Net Income surfaced as their own lines by the
  // renderer. equityRows here is only the contributed-capital accounts.
  function equityRowsForSub(sub) {
    return bsCodes
      .map(code => {
        const rc = colCur.map.get(code), rp = colPri.map.get(code);
        const ref = rc || rp;
        if (!ref || ref.type !== 'Equity') return null;
        const cls = bsClassify(ref);
        if (cls.sub !== sub) return null;
        const cur = rc ? bal(rc) : 0, pri = rp ? bal(rp) : 0;
        if (isZero(cur) && isZero(pri)) return null;
        return { code, name: ref.name, cur: r2(cur), pri: r2(pri), change: r2(cur - pri) };
      })
      .filter(Boolean);
  }
  const equityRows = equityRowsForSub('Members Equity');
  const retainedRows = equityRowsForSub('Retained Earnings');

  const totalAssets = { cur: r2(assetSections.reduce((s, x) => s + x.total.cur, 0)),
                        pri: r2(assetSections.reduce((s, x) => s + x.total.pri, 0)) };
  const totalLiab = { cur: r2(liabSections.reduce((s, x) => s + x.total.cur, 0)),
                      pri: r2(liabSections.reduce((s, x) => s + x.total.pri, 0)) };
  const totalContribEquity = { cur: r2(equityRows.reduce((a, r) => a + r.cur, 0)),
                               pri: r2(equityRows.reduce((a, r) => a + r.pri, 0)) };
  // Retained-earnings subsection total (frozen opening RE carried on the BS as
  // its own line, separate from the current-year Net Income line).
  const totalRetained = { cur: r2(retainedRows.reduce((a, r) => a + r.cur, 0)),
                          pri: r2(retainedRows.reduce((a, r) => a + r.pri, 0)) };
  // Net income line: current-year YTD (cur column) and prior-year-through-prior-
  // month YTD (pri column) — matches the hand-prepared "Net Income (Loss)" row.
  const niLine = { cur: niYtd, pri: niPriYtd };
  const totalEquity = { cur: r2(totalContribEquity.cur + totalRetained.cur + niLine.cur), pri: r2(totalContribEquity.pri + totalRetained.pri + niLine.pri) };
  const totalLiabEquity = { cur: r2(totalLiab.cur + totalEquity.cur), pri: r2(totalLiab.pri + totalEquity.pri) };

  // ── Statements of Operations ───────────────────────────────────────────────
  // Build a P&L line set keyed by code, with current-month / prior-month / YTD.
  function plMap(rows) { const m = new Map(); for (const r of rows) if (r.type === 'Revenue' || r.type === 'Expense') m.set(r.code, r); return m; }
  const mCur = plMap(isCur), mPri = plMap(isPri), mYtd = plMap(isYtd);
  const plCodes = Array.from(new Set([...mCur.keys(), ...mPri.keys(), ...mYtd.keys()]))
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  function plLines(pred) {
    return plCodes.map(code => {
      const ref = mYtd.get(code) || mCur.get(code) || mPri.get(code);
      if (!ref || !pred(ref)) return null;
      // For P&L display, revenue/expense both shown as positive magnitudes in
      // their sections; sign handling happens at the section-total level.
      const val = m => { const r = m.get(code); return r ? bal(r) : 0; };
      const cur = val(mCur), pri = val(mPri), ytd = val(mYtd);
      if (isZero(cur) && isZero(pri) && isZero(ytd)) return null;
      return { code, name: ref.name, cur: r2(cur), pri: r2(pri), ytd: r2(ytd), change: r2(cur - pri) };
    }).filter(Boolean);
  }

  const revenue = plLines(r => r.type === 'Revenue');
  const cogs = plLines(r => r.type === 'Expense' && /cogs|cost of goods|cost of revenue|car hire/i.test((r.subtype || '') + ' ' + (r.name || '')));
  const cogsCodes = new Set(cogs.map(l => l.code));
  const opex = plLines(r => r.type === 'Expense' && !cogsCodes.has(r.code));

  const sumCol = (lines, k) => r2(lines.reduce((s, l) => s + l[k], 0));
  const totRev = { cur: sumCol(revenue, 'cur'), pri: sumCol(revenue, 'pri'), ytd: sumCol(revenue, 'ytd') };
  const totCogs = { cur: sumCol(cogs, 'cur'), pri: sumCol(cogs, 'pri'), ytd: sumCol(cogs, 'ytd') };
  const grossProfit = { cur: r2(totRev.cur - totCogs.cur), pri: r2(totRev.pri - totCogs.pri), ytd: r2(totRev.ytd - totCogs.ytd) };
  const totOpex = { cur: sumCol(opex, 'cur'), pri: sumCol(opex, 'pri'), ytd: sumCol(opex, 'ytd') };
  const netIncome = { cur: r2(grossProfit.cur - totOpex.cur), pri: r2(grossProfit.pri - totOpex.pri), ytd: r2(grossProfit.ytd - totOpex.ytd) };

  // Group operating expenses into the 11 presentation categories (per the CLR
  // operating-expense restructure). Category subtotals sum to totOpex exactly —
  // this is purely a re-grouping, so net income is unaffected. Only categories
  // that actually have lines are emitted, in the fixed presentation order.
  const opexByCat = new Map();
  for (const l of opex) {
    const cat = plExpenseCategory(l);
    if (!opexByCat.has(cat)) opexByCat.set(cat, []);
    opexByCat.get(cat).push(l);
  }
  const opexGroups = [];
  const pushGroup = (cat, lines) => {
    lines.sort((a, b) => String(a.code).localeCompare(String(b.code)));
    opexGroups.push({
      title: cat,
      lines,
      subtotal: { cur: sumCol(lines, 'cur'), pri: sumCol(lines, 'pri'), ytd: sumCol(lines, 'ytd') },
    });
  };
  for (const cat of PL_EXPENSE_CATEGORY_ORDER) {
    const lines = opexByCat.get(cat);
    if (lines && lines.length) pushGroup(cat, lines);
    opexByCat.delete(cat);
  }
  // Safety net: any category the heuristic produced but that isn't in the fixed
  // order (shouldn't happen) is still emitted so no expense line is ever dropped.
  for (const [cat, lines] of opexByCat) {
    if (lines && lines.length) pushGroup(cat, lines);
  }

  // ── Statement of Cash Flows (indirect, YTD) ────────────────────────────────
  // Beginning balances = as of (year start − 1 day). Deltas over the YTD window.
  const bsOpen = await getBalances({ as_of: priorMonthEnd(ys), close_pl_before: ys });
  const openMap = new Map(); for (const r of bsOpen) openMap.set(r.code, r);
  const curMap = new Map(); for (const r of bsCur) if (r.type !== 'Revenue' && r.type !== 'Expense') curMap.set(r.code, r);
  const isCashRow = isCashAccount;
  const cashBeg = sumRows(bsOpen.filter(isCashRow), () => true);
  const cashEnd = sumRows(bsCur.filter(isCashRow), () => true);

  // Non-cash add-back for amortization/depreciation. Add back ONLY the portion
  // booked as a P&L expense — that is the only part that reduced net income and
  // must be restored in operating activities. On development entities (SRN),
  // depreciation is typically CAPITALIZED (Dr asset, Cr accumulated amort) with
  // no P&L expense; that is a purely non-cash reclass between two asset accounts,
  // so there is nothing to add back — both legs flow through investing where they
  // net out. Basing the add-back on the P&L expense (0 when capitalized) and
  // letting the contra flow normally in that case is what keeps the statement
  // tying; verified against live SRN GL (actual cash change reproduced exactly).
  const amortExpense = r2(sumRows(isYtd, r => r.type === 'Expense' && /amortization|depreciation/i.test(r.name)));
  const amortization = amortExpense;

  // Classify every non-cash balance-sheet account into a cash-flow bucket by a
  // single pass, so nothing is silently dropped. Each account's period delta
  // maps to its cash effect (asset increase → cash use; liab/equity increase →
  // cash source). We itemize the named lines the hand-prepared statement shows
  // and roll everything else in each section into an "other" catch-all, which
  // guarantees the sections are complete and the statement ties by construction.
  const cfBuckets = {
    ar: 0, prepaidOther: 0, ap: 0, accrued: 0, intercompany: 0, otherOperating: 0,
    capex: 0, ltInvest: 0, otherInvesting: 0,
    equityContrib: 0, debtChange: 0, otherFinancing: 0,
  };
  const cfCodes = new Set([...openMap.keys(), ...curMap.keys()]);
  for (const code of cfCodes) {
    const rc = curMap.get(code), ro = openMap.get(code);
    const ref = rc || ro;
    if (!ref) continue;
    if (ref.type === 'Revenue' || ref.type === 'Expense') continue;
    if (isCashRow(ref)) continue; // cash itself is the reconciling target
    // Skip the accumulated-amortization/depreciation contra ONLY when its move was
    // booked as a P&L expense (added back above) — otherwise it would be counted
    // twice. When depreciation is capitalized (no P&L expense), let the contra flow
    // through investing, where it nets against the capitalized asset leg.
    if (ref.type === 'Asset' && /accum|amortization|depreciation/i.test(ref.name) && !isZero(amortExpense)) continue;
    const delta = r2((rc ? bal(rc) : 0) - (ro ? bal(ro) : 0));
    if (isZero(delta)) continue;
    const nm = ref.name || '';
    if (ref.type === 'Asset') {
      const sec = bsSection(ref);
      const cashEffect = -delta; // asset up → cash down
      if (/intercompany|due from|due to/i.test(nm)) cfBuckets.intercompany += cashEffect;
      else if (sec === 'Current Assets' && /receivable/i.test(nm)) cfBuckets.ar += cashEffect;
      else if (sec === 'Current Assets') cfBuckets.prepaidOther += cashEffect;
      else if (sec === 'Fixed Assets') cfBuckets.capex += cashEffect;
      else cfBuckets.ltInvest += cashEffect; // intangible / investment / other long-term
    } else if (ref.type === 'Liability') {
      const cashEffect = delta; // liability up → cash up
      const sec = bsSection(ref);
      if (/payable/i.test(nm)) cfBuckets.ap += cashEffect;
      else if (sec === 'Long Term Liabilities') cfBuckets.debtChange += cashEffect;
      else cfBuckets.accrued += cashEffect; // accrued / other current liabilities
    } else if (ref.type === 'Equity') {
      // Equity delta includes contributions/distributions but NOT current-year
      // net income (P&L is excluded above), so the whole delta is financing.
      cfBuckets.equityContrib += delta;
    }
  }
  Object.keys(cfBuckets).forEach(k => { cfBuckets[k] = r2(cfBuckets[k]); });

  const cashFlow = {
    netIncome: niYtd,
    amortization,
    changeAR: cfBuckets.ar,
    changePrepaidOther: cfBuckets.prepaidOther,
    changeAP: cfBuckets.ap,
    changeAccrued: cfBuckets.accrued,
    changeIntercompany: cfBuckets.intercompany,
    capex: cfBuckets.capex,
    ltInvest: cfBuckets.ltInvest,
    equityContrib: cfBuckets.equityContrib,
    debtChange: cfBuckets.debtChange,
    cashBeg: r2(cashBeg), cashEnd: r2(cashEnd),
  };
  cashFlow.netOperating = r2(cashFlow.netIncome + cashFlow.amortization + cashFlow.changeAR + cashFlow.changePrepaidOther + cashFlow.changeAP + cashFlow.changeAccrued + cashFlow.changeIntercompany);
  cashFlow.netInvesting = r2(cashFlow.capex + cashFlow.ltInvest);
  cashFlow.netFinancing = r2(cashFlow.equityContrib + cashFlow.debtChange);
  cashFlow.netChange = r2(cashFlow.netOperating + cashFlow.netInvesting + cashFlow.netFinancing);
  // Tie-out: the reconciled net change vs. the actual cash movement. With
  // complete opening/closing balance sheets these agree by construction; any
  // residual (rounding, a mid-year chart change) is surfaced, not hidden.
  cashFlow.actualCashChange = r2(cashEnd - cashBeg);
  cashFlow.tieOut = r2(cashFlow.netChange - cashFlow.actualCashChange);

  // ── Statement of Changes in Members' Equity ───────────────────────────────
  // Beginning (year start) contributed equity by account + beginning RE, then
  // contributions (delta) and YTD net income → ending.
  const equityMembers = equityRows.map(r => {
    const openRow = bsOpen.find(x => x.code === r.code);
    const beg = openRow ? r2(bal(openRow)) : 0;
    return { code: r.code, name: r.name, beginning: beg, contributions: r2(r.cur - beg), distributions: 0, netIncome: 0, ending: r2(r.cur) };
  });
  // Retained earnings row (opening RE + YTD NI). Opening RE = frozen RE in the
  // BS columns is captured by the equity accounts already; but the hand-prepared
  // statement carries a distinct Retained Earnings member line.
  const reOpenRow = bsOpen.find(x => x.type === 'Equity' && /retained earning/i.test(x.name));
  const reOpen = reOpenRow ? r2(bal(reOpenRow)) : 0;
  const reMember = { code: 're', name: 'Retained Earnings', beginning: reOpen, contributions: 0, distributions: 0, netIncome: niYtd, ending: r2(reOpen + niYtd) };
  const equityStmt = [...equityMembers.filter(m => !/retained earning/i.test(m.name)), reMember];
  const equityTotals = equityStmt.reduce((t, m) => ({
    beginning: r2(t.beginning + m.beginning), contributions: r2(t.contributions + m.contributions),
    distributions: r2(t.distributions + m.distributions), netIncome: r2(t.netIncome + m.netIncome), ending: r2(t.ending + m.ending),
  }), { beginning: 0, contributions: 0, distributions: 0, netIncome: 0, ending: 0 });

  return {
    meta: { entityName: displayEntityName(opts.entityName), asOf, priorDate: priorBsDate, longDate: longDate(asOf),
            priorLongDate: longDate(priorBsDate), monthsEnded: monthsEndedLabel(asOf),
            period: (opts.period || 'monthly').toLowerCase(), periodLabel: period.periodLabel, colLabel: period.colLabel },
    balanceSheet: { assetSections, liabSections, equityRows, retainedRows, totalAssets, totalLiab, totalContribEquity, niLine, totalEquity, totalLiabEquity },
    operations: { revenue, cogs, opex, opexGroups, totRev, totCogs, grossProfit, totOpex, netIncome },
    cashFlow,
    equity: { rows: equityStmt, totals: equityTotals },
    checks: {
      balanceSheetTies: isZero(totalAssets.cur - totalLiabEquity.cur),
      balanceSheetDiff: r2(totalAssets.cur - totalLiabEquity.cur),
      cashFlowTies: isZero(cashFlow.tieOut),
      cashFlowDiff: cashFlow.tieOut,
      niAgrees: isZero(netIncome.ytd - niYtd),
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// PDF rendering — CPA-plain style (centered bold headers, thin rules, footer).
// Renders the four GL statements onto US-Letter pages with pdf-lib.
// ═══════════════════════════════════════════════════════════════════════════

const PAGE = { w: 612, h: 792, mL: 54, mR: 54, mT: 60, mB: 54 };
const FS = { title: 11, sub: 9.5, head: 8, row: 8.5, foot: 7.5 }; // point sizes

// A tiny page-layout cursor over a pdf-lib page. Handles the y-cursor, centered
// headers, a footer, right-aligned numeric columns, and automatic page breaks.
function makeLayout(pdf, fonts, meta, statementTitle, opts = {}) {
  const { reg, bold } = fonts;
  // Page geometry: portrait by default, landscape when requested (opts.landscape).
  const PW = opts.landscape ? PAGE.h : PAGE.w;
  const PH = opts.landscape ? PAGE.w : PAGE.h;
  const dateLine = opts.dateLine || null; // repeated heading date-line (every page)
  let page, y;
  const cols = []; // right edges for numeric columns, set per statement
  // Saved column-header spec so the header row repeats atop every continuation
  // page (so a reader always knows what each column is). Set by colHeaders();
  // replayed by newPage(). `_replaying` guards against re-entrancy since
  // colHeaders() calls ensure() which could otherwise trigger another newPage().
  let _hdrSpec = null;      // { labels, hopts }
  let _replaying = false;

  function newPage() {
    page = pdf.addPage([PW, PH]);
    y = PH - PAGE.mT;
    drawHeader();
    drawFooter();
    // On every page: draw the repeated date-line heading + a wider blank space
    // (~2 lines) between the heading and the column headers / first row so the
    // top of the statement doesn't look cramped.
    if (dateLine) { textC(dateLine, FS.sub, reg, PH - PAGE.mT - 2); y -= 14; }
    y -= 20;
    // Repeat the column headers on continuation pages (not the very first page,
    // where the statement body calls colHeaders() itself in the right spot).
    if (_hdrSpec && !_replaying) {
      _replaying = true;
      layout.colHeaders(_hdrSpec.labels, _hdrSpec.hopts);
      _replaying = false;
    }
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
    // Centered footer on every page: "<entity>, <date>  |  See Executive Summary".
    const label = meta.entityName + ', ' + meta.longDate + '  |  See Executive Summary';
    const w = reg.widthOfTextAtSize(label, FS.foot);
    page.drawText(label, { x: (PW - w) / 2, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }
  function ensure(space) { if (y - space < PAGE.mB + 8) newPage(); }

  const layout = {
    _subline: null,
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
    start() { newPage(); },
    get y() { return y; },
    set y(v) { y = v; },
    get page() { return page; },
    space(n) { y -= n; },
    ensure,
    // Keep-together: if the next `space` points of content won't fit on the
    // current page, break to a new page BEFORE rendering it — so a grouped block
    // (e.g. an operating-expense category with its header, lines and subtotal)
    // never splits across a page boundary. If the block is taller than a full
    // usable page it can't be kept whole, so we let it flow normally rather than
    // emitting an endless run of blank pages.
    keepTogether(space) {
      const usable = (PH - PAGE.mT) - (PAGE.mB + 8);
      if (space <= usable) ensure(space);
    },
    // Column layout: array of right-edge x positions for numeric columns.
    setCols(rightEdges) { cols.length = 0; rightEdges.forEach(e => cols.push(e)); },
    // Column headers (right-aligned above each numeric column). Only the header
    // cells themselves are underlined (per the reference), not a full-width rule.
    //
    // hopts:
    //   underline   — draw an underline beneath each header cell (default off)
    //   bottomAlign — align every multi-line label to a COMMON bottom baseline
    //                 (taller labels grow upward) so all columns' last lines sit
    //                 on one row. Off (default) = top-aligned from the cursor
    //                 downward (legacy behavior for single-line BS/Operations
    //                 headers).
    //   colBox      — when underlining, span each rule across a fixed per-column
    //                 box (right edge = cols[i], width = inter-column pitch minus
    //                 a gutter) instead of only the text width, so the rules read
    //                 as one-per-column with a narrow gap between them.
    colHeaders(labels, hopts = {}) {
      // Remember the spec so newPage() can repeat this header on continuation
      // pages. Only record on the FIRST (body-driven) call, not during replay.
      if (!_replaying) _hdrSpec = { labels, hopts };
      const LH = 9;                 // header line height
      const nLines = Math.max(1, ...labels.map(l => String(l).split('\n').length));
      // Reserve height for the tallest label block + the underline + trailing gap.
      ensure(LH * nLines + 10);
      // ALL column headers are bottom-aligned: the LAST line of every label sits
      // on one common baseline, and taller (multi-line) labels grow UPWARD from
      // it. That common baseline is `baseY`. Start the block by dropping from the
      // cursor so the tallest label's top line clears the heading above.
      const topY = y;                       // top of the header block
      const baseY = topY - (nLines - 1) * LH; // common bottom baseline for last lines
      // Per-column underline box: the inter-column pitch minus a gutter, so
      // adjacent underlines are separated. Use the smallest pitch so no two
      // boxes overlap.
      const GUTTER = 14; // blank points between adjacent underline boxes
      let pitch = Infinity;
      for (let i = 1; i < cols.length; i++) pitch = Math.min(pitch, cols[i] - cols[i - 1]);
      const boxW = Number.isFinite(pitch) ? Math.max(20, pitch - GUTTER) : null;
      // Underline sits just BELOW the last line's baseline for every column, at a
      // single common y so all column rules line up on one row at the bottom of
      // the header block (never crossing through the text).
      const uy = baseY - 3;
      labels.forEach((lab, i) => {
        const parts = String(lab).split('\n');
        let maxW = 0;
        parts.forEach(pl => { const w = bold.widthOfTextAtSize(pl, FS.head); if (w > maxW) maxW = w; });
        // Bottom-align: the last line sits on baseY; earlier lines stack above it.
        parts.forEach((pl, pi) => {
          const w = bold.widthOfTextAtSize(pl, FS.head);
          const lineY = baseY + (parts.length - 1 - pi) * LH;
          page.drawText(pl, { x: cols[i] - w, y: lineY, size: FS.head, font: bold });
        });
        if (hopts.underline) {
          // colBox → fixed per-column span (with a gutter) so the rule reads as
          // one-per-column; otherwise hug just the widest line of this label.
          const span = (hopts.colBox && boxW) ? boxW : maxW;
          page.drawLine({ start: { x: cols[i] - span, y: uy }, end: { x: cols[i], y: uy }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
        }
      });
      // Advance the cursor to just below the underline plus a small gap.
      y = uy - 8;
    },
    sectionTitle(str) {
      ensure(16);
      page.drawText(str, { x: PAGE.mL, y, size: FS.row, font: bold });
      y -= 13;
    },
    // A data row: label (optionally indented) + numeric cells (strings already formatted).
    //   valueInset — right-inset (pt) so a right-aligned value doesn't jam the
    //                column's right edge (used on the wide equity columns).
    //   colRules   — draw the subtotal/total rules PER COLUMN (each spanning that
    //                column's value box with a gutter) instead of one continuous
    //                line across all columns. Defaults ON so every statement's
    //                subtotal/total underlines sit under each number separately,
    //                never as one long line running across the whole row.
    row(label, cells, { indent = 12, boldRow = false, ruleAbove = false, ruleBelow = false, doubleBelow = false, gapAfter = 0, dollarPrefix = false, valueInset = 0, colRules = true } = {}) {
      ensure(13);
      const font = boldRow ? bold : reg;
      // Per-column rule width. For the per-column rules we want a uniform box
      // sized to the numeric columns (the inter-column pitch), NOT the wide
      // first-column default used for "$" placement — otherwise the leftmost
      // rule would run far left under the label. Use the smallest inter-column
      // pitch so no two rules overlap and each sits under just its number.
      const colWidth = (i) => (i === 0 ? 78 : Math.max(40, cols[i] - cols[i - 1]));
      let pitch = Infinity;
      for (let i = 1; i < cols.length; i++) pitch = Math.min(pitch, cols[i] - cols[i - 1]);
      const ruleBoxW = Number.isFinite(pitch) ? Math.max(40, pitch) : colWidth(0);
      const ruleLeft = cols[0] - colWidth(0) + 2;
      const ruleRight = cols[cols.length - 1];
      // Per-column rule segments: one short line under each column's value box,
      // leaving a gutter between adjacent columns so the rules read as one-per-
      // column rather than a single line drawn straight across ("일직선").
      const GUTTER = 12;
      const drawRule = (yy) => {
        if (colRules) {
          cols.forEach((cx, i) => {
            const x0 = cx - ruleBoxW + 2 + GUTTER / 2;
            const x1 = cx;
            page.drawLine({ start: { x: x0, y: yy }, end: { x: x1, y: yy }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
          });
        } else {
          page.drawLine({ start: { x: ruleLeft, y: yy }, end: { x: ruleRight, y: yy }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
        }
      };
      if (ruleAbove) drawRule(y + 9);
      page.drawText(String(label), { x: PAGE.mL + indent, y, size: FS.row, font });
      cells.forEach((c, i) => {
        if (c == null || c === '') return;
        const s = String(c);
        const w = font.widthOfTextAtSize(s, FS.row);
        page.drawText(s, { x: cols[i] - w - valueInset, y, size: FS.row, font });
        if (dollarPrefix) {
          // "$" anchored at the left of this column's cell box.
          const dx = cols[i] - colWidth(i) + 2;
          page.drawText('$', { x: dx, y, size: FS.row, font });
        }
      });
      if (ruleBelow) drawRule(y - 3);
      if (doubleBelow) { drawRule(y - 3); drawRule(y - 5); }
      y -= 12 + gapAfter;
    },
  };
  return layout;
}

// Render the four statements into a fresh PDFDocument and return its bytes.
// If an `outOffsets` array is passed, it is filled with { label, page } entries
// giving each statement's 0-based starting page index within this PDF (used to
// compute Table-of-Contents page references).
async function renderStatementsPdf(s, outOffsets) {
  const track = (label) => { if (outOffsets) outOffsets.push({ label, page: pdf.getPageCount() }); };
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fonts = { reg, bold };
  const m = s.meta;
  const money = v => acct(v, { dash: true });

  // Numeric column right-edges. Balance Sheet now has 3 columns (current, prior,
  // and a Change column at the far right); Operations uses 3.
  const RIGHT = PAGE.w - PAGE.mR;
  // Balance Sheet: current & prior dates plus a "Change" column on the far right.
  // Pack three columns into the same right band the two-column layout used.
  const bsCols = [RIGHT - 190, RIGHT - 95, RIGHT];
  const twoCols = [RIGHT - 95, RIGHT];
  const threeCols = [RIGHT - 200, RIGHT - 100, RIGHT];
  // Operations: four columns — current month, prior month, a Change column
  // (current − prior) sitting between the prior-month and Year-to-Date columns,
  // and Year to Date on the far right. Numeric columns are kept narrow (76pt
  // pitch) so the first (account-name) column is wide enough to show full names.
  const opsCols = [RIGHT - 228, RIGHT - 152, RIGHT - 76, RIGHT];

  // ── 1. Balance Sheet ───────────────────────────────────────────────────────
  {
    // Heading date-line repeats on every page (incl. continuation pages) and a
    // blank space follows it before the first row. Dates are NOT underlined.
    const L = makeLayout(pdf, fonts, m, 'Balance Sheets', { dateLine: m.longDate + ' and ' + m.priorLongDate });
    track('Balance Sheets');
    L.start();
    L.setCols(bsCols);
    // Three columns now: current date, prior date, and a "Change" column showing
    // the month-over-month movement (current − prior). Headers underlined.
    L.colHeaders([m.longDate, m.priorLongDate, 'Change'], { underline: true });
    L.sectionTitle('ASSETS');
    // A BS money triple: current, prior, and the change (cur − pri). Change is
    // computed here so every line — accounts, subtotals, section totals, and the
    // grand totals — carries a consistent month-over-month delta column.
    const bsCells = (cur, pri) => [money(cur), money(pri), money(r2(cur - pri))];
    // Two-level: section header → subsection header → accounts → subsection
    // subtotal, then a bold section total. Subsection headers/subtotals appear
    // only when a section has multiple subsections or a contra subsection;
    // otherwise the section collapses to accounts directly under its header to
    // avoid a redundant subtotal that just repeats a single account line.
    const renderBsSection = (sec, sectionTotalLabel) => {
      L.row(sec.title, [], { indent: 6, boldRow: true });
      const showSubHeaders = sec.subs.length > 1 || sec.subs.some(su => su.contra);
      for (const su of sec.subs) {
        if (showSubHeaders) L.row(su.title, [], { indent: 16 });
        const rowIndent = showSubHeaders ? 26 : 16;
        for (const r of su.rows) L.row(r.name, bsCells(r.cur, r.pri), { indent: rowIndent });
        if (showSubHeaders && su.rows.length > 1) {
          // GL balances are already signed; a contra subtotal (accumulated
          // amortization) is naturally negative and prints in parentheses.
          L.row('Total ' + su.title, bsCells(su.subtotal.cur, su.subtotal.pri), { indent: 20, ruleAbove: true });
        }
      }
      L.row(sectionTotalLabel, bsCells(sec.total.cur, sec.total.pri), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    };
    for (const sec of s.balanceSheet.assetSections) renderBsSection(sec, 'Total ' + sec.title);
    L.row('Total Assets', bsCells(s.balanceSheet.totalAssets.cur, s.balanceSheet.totalAssets.pri), { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, gapAfter: 8 });

    L.sectionTitle('LIABILITIES AND MEMBERS\u2019 EQUITY');
    for (const sec of s.balanceSheet.liabSections) renderBsSection(sec, 'Total ' + sec.title);
    L.row('Total Liabilities', bsCells(s.balanceSheet.totalLiab.cur, s.balanceSheet.totalLiab.pri), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Members\u2019 Equity', [], { indent: 6, boldRow: true });
    for (const r of s.balanceSheet.equityRows) L.row(r.name, bsCells(r.cur, r.pri), { indent: 16 });
    for (const r of (s.balanceSheet.retainedRows || [])) L.row(r.name, bsCells(r.cur, r.pri), { indent: 16 });
    L.row('Net Income (Loss)', bsCells(s.balanceSheet.niLine.cur, s.balanceSheet.niLine.pri), { indent: 16 });
    L.row('Total Members\u2019 Equity', bsCells(s.balanceSheet.totalEquity.cur, s.balanceSheet.totalEquity.pri), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Total Liabilities and Members\u2019 Equity', bsCells(s.balanceSheet.totalLiabEquity.cur, s.balanceSheet.totalLiabEquity.pri), { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }

  // ── 2. Statements of Operations ─────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statements of Operations', { dateLine: 'For the Months Ended ' + m.longDate + ' and ' + m.priorLongDate });
    track('Statements of Operations');
    L.start();
    L.setCols(opsCols);
    // Current-month column shows the current period-end date; prior-month column
    // shows the prior period-end date (per round-2 feedback). A Change column
    // (current − prior) sits between the prior-month and Year-to-Date columns.
    // Headers underlined.
    L.colHeaders([m.longDate, m.priorLongDate, 'Change', 'Year to Date'], { underline: true });
    const chg = (cur, pri) => money(r2(cur - pri));
    const line = (r, o = {}) => L.row(r.name, [money(r.cur), money(r.pri), chg(r.cur, r.pri), money(r.ytd)], { indent: 16, ...o });

    L.sectionTitle('Revenue');
    s.operations.revenue.forEach(r => line(r));
    L.row('Total Revenue', [money(s.operations.totRev.cur), money(s.operations.totRev.pri), chg(s.operations.totRev.cur, s.operations.totRev.pri), money(s.operations.totRev.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    if (s.operations.cogs.length) {
      L.sectionTitle('Cost of Revenue');
      s.operations.cogs.forEach(r => line(r));
      L.row('Total Cost of Revenue', [money(s.operations.totCogs.cur), money(s.operations.totCogs.pri), chg(s.operations.totCogs.cur, s.operations.totCogs.pri), money(s.operations.totCogs.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 4 });
      L.row('Gross Profit', [money(s.operations.grossProfit.cur), money(s.operations.grossProfit.pri), chg(s.operations.grossProfit.cur, s.operations.grossProfit.pri), money(s.operations.grossProfit.ytd)], { indent: 6, boldRow: true, gapAfter: 6 });
    }
    L.sectionTitle('Operating Expenses');
    // Grouped into the 11 presentation categories, each with its own subtotal.
    // Category subtotals sum to Total Operating Expenses exactly (pure re-group).
    // Fall back to a flat list if grouping produced nothing (defensive).
    const groups = s.operations.opexGroups && s.operations.opexGroups.length
      ? s.operations.opexGroups : null;
    if (groups) {
      for (const g of groups) {
        // Keep each category together: reserve the height of its header row +
        // all line rows + (if shown) the subtotal row, so a group like
        // "Contracted Services" never splits with its subtotal on the next page.
        const nLines = g.lines.length;
        const hasSubtotal = nLines > 1;
        const groupH = 12 /* header */ + nLines * 12 + (hasSubtotal ? 12 : 0) + 4 /* subtotal rule buffer */;
        L.keepTogether(groupH);
        L.row(g.title, [], { indent: 12, boldRow: true });
        g.lines.forEach(r => L.row(r.name, [money(r.cur), money(r.pri), chg(r.cur, r.pri), money(r.ytd)], { indent: 26 }));
        if (hasSubtotal) {
          L.row('Total ' + g.title, [money(g.subtotal.cur), money(g.subtotal.pri), chg(g.subtotal.cur, g.subtotal.pri), money(g.subtotal.ytd)], { indent: 20, ruleAbove: true });
        }
      }
    } else {
      s.operations.opex.forEach(r => line(r));
    }
    L.row('Total Operating Expenses', [money(s.operations.totOpex.cur), money(s.operations.totOpex.pri), chg(s.operations.totOpex.cur, s.operations.totOpex.pri), money(s.operations.totOpex.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Net Income (Loss)', [money(s.operations.netIncome.cur), money(s.operations.netIncome.pri), chg(s.operations.netIncome.cur, s.operations.netIncome.pri), money(s.operations.netIncome.ytd)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }

  // ── 3. Statement of Cash Flows ──────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows', { dateLine: m.monthsEnded });
    track('Statement of Cash Flows');
    L.start();
    L.setCols([RIGHT]);
    // No column heading: the single YTD column is self-evident from the
    // statement title, so no "Year to Date" label is drawn above it.
    const cf = s.cashFlow;
    L.sectionTitle('Cash Flows from Operating Activities');
    L.row('Net Income (Loss)', [money(cf.netIncome)], { indent: 16 });
    L.row('Adjustments to reconcile net income to net cash:', [], { indent: 16 });
    if (!isZero(cf.amortization)) L.row('Amortization and depreciation', [money(cf.amortization)], { indent: 28 });
    if (!isZero(cf.changeAR)) L.row('(Increase) decrease in accounts receivable', [money(cf.changeAR)], { indent: 28 });
    if (!isZero(cf.changePrepaidOther)) L.row('(Increase) decrease in prepaid and other current assets', [money(cf.changePrepaidOther)], { indent: 28 });
    if (!isZero(cf.changeIntercompany)) L.row('(Increase) decrease in intercompany balances', [money(cf.changeIntercompany)], { indent: 28 });
    if (!isZero(cf.changeAP)) L.row('Increase (decrease) in accounts payable', [money(cf.changeAP)], { indent: 28 });
    if (!isZero(cf.changeAccrued)) L.row('Increase (decrease) in accrued and other liabilities', [money(cf.changeAccrued)], { indent: 28 });
    L.row('Net Cash Provided (Used) by Operating Activities', [money(cf.netOperating)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.sectionTitle('Cash Flows from Investing Activities');
    if (!isZero(cf.capex)) L.row('Purchases of property and equipment', [money(cf.capex)], { indent: 28 });
    if (!isZero(cf.ltInvest)) L.row('Development, intangible and other long-term costs', [money(cf.ltInvest)], { indent: 28 });
    L.row('Net Cash Provided (Used) by Investing Activities', [money(cf.netInvesting)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.sectionTitle('Cash Flows from Financing Activities');
    if (!isZero(cf.equityContrib)) L.row('Member contributions (distributions), net', [money(cf.equityContrib)], { indent: 28 });
    if (!isZero(cf.debtChange)) L.row('Proceeds from (repayment of) long-term debt', [money(cf.debtChange)], { indent: 28 });
    L.row('Net Cash Provided (Used) by Financing Activities', [money(cf.netFinancing)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.row('Net Increase (Decrease) in Cash', [money(cf.netChange)], { indent: 6, boldRow: true, ruleAbove: true });
    L.row('Cash, Beginning of Period', [money(cf.cashBeg)], { indent: 6 });
    L.row('Cash, End of Period', [money(cf.cashEnd)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
    if (!isZero(cf.tieOut)) {
      L.space(6);
      L.row('Note: reconciled change differs from cash movement by ' + money(cf.tieOut) + ' (see notes).', [], { indent: 6 });
    }
  }

  // ── 4. Statement of Changes in Members' Equity ──────────────────────────────
  {
    // Landscape page mirroring the CPA reference: five columns, a Distributions
    // column shown even when all zero, and a Net Income (Loss) column wide enough
    // to keep the value on one row. Values are shown without a "$" prefix.
    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\u2019 Equity',
      { landscape: true, dateLine: m.monthsEnded });
    const LRIGHT = PAGE.h - PAGE.mR; // landscape printable right edge (PAGE.h is the long side)
    track('Statement of Changes in Members\u2019 Equity');
    L.start();
    // Column right-edges across the landscape width. Two-line headers, dates
    // shown as m/d/yyyy short form to match the reference.
    const shortMD = (long) => {
      // "April 30, 2026" -> "4/30/2026"
      const map = { January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12 };
      const mm = long.match(/^(\w+)\s+(\d+),\s+(\d+)$/);
      if (!mm) return long;
      return map[mm[1]] + '/' + mm[2] + '/' + mm[3];
    };
    const begDate = '1/1/' + String(m.asOf).slice(0, 4);
    const endDate = shortMD(m.longDate);
    // Right-edges anchored at the printable right edge and marched LEFT by a
    // fixed pitch. The first numeric column (c1) is placed far enough right that
    // its "$" prefix cell box (c1 - colWidth(0) + 2) clears the longest member
    // label — the previous layout put c1 too far left, letting long labels
    // ("Contributed Capital - County Line Rail Fund, LLP", right edge ≈ 254pt)
    // overrun into the first column. c1 = 356 puts the "$" box left at ~280pt
    // (a ~26pt gap after the longest label); pitch (~95pt) still leaves each
    // column ample room for a "$" prefix and a large right-aligned value.
    const c1 = 356;
    const PITCH = (LRIGHT - c1) / 4;
    const c2 = c1 + PITCH, c3 = c2 + PITCH, c4 = c3 + PITCH, c5 = LRIGHT;
    const eCols = [c1, c2, c3, c4, c5];
    L.setCols(eCols);
    // "Balances at" / "Equity" on the TOP two lines and the DATE on the BOTTOM
    // line, so the date reads directly above the header underline (matches the
    // CPA reference, which places the date at the bottom of the header block).
    L.colHeaders([
      'Equity Balances at\n' + begDate,
      'Contributions',
      'Distributions',
      'Net Income\n(Loss)',
      'Equity Balances at\n' + endDate,
    ], { bottomAlign: true, underline: true, colBox: true });
    // Money cell: value right-aligned to the column with a small inset so it
    // doesn't jam against the column edge. No "$" prefix (per request).
    const dollarRow = (label, vals, o = {}) => {
      L.row(label, vals.map(v => acct(v)), Object.assign({ indent: 10, dollarPrefix: false, valueInset: 4 }, o));
    };
    L.row('Member', [], { indent: 6, boldRow: true });
    for (const r of s.equity.rows) {
      dollarRow(r.name, [r.beginning, r.contributions, r.distributions, r.netIncome, r.ending], { indent: 16 });
    }
    const t = s.equity.totals;
    dollarRow('Total', [t.beginning, t.contributions, t.distributions, t.netIncome, t.ending],
      { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, colRules: true });
  }

  return await pdf.save();
}

// ═══════════════════════════════════════════════════════════════════════════
// stripInvoiceLogPages — given a requisition-report PDF (bytes), drop any page
// whose text contains a "Current Invoice Log" or "Prior Invoice Log" heading.
// Uses pdf-parse per-page via the pagerender hook to get each page's text, then
// rebuilds the PDF with pdf-lib keeping only the pages we want.
// Returns { bytes, removed: [pageIndexes], kept, total }.
// ═══════════════════════════════════════════════════════════════════════════
const INVOICE_LOG_RE = /(current|prior)\s+invoice\s+log/i;

async function stripInvoiceLogPages(pdfBytes) {
  const pdfParse = require('pdf-parse');
  // Collect per-page text. pdf-parse calls pagerender once per page in order.
  // pdf-parse ships an old pdf.js that can choke on some PDFs; if it throws we
  // fall back to keeping every page (and surface a flag) rather than failing the
  // whole package generation.
  const pageTexts = [];
  let parseFailed = false;
  try {
    await pdfParse(Buffer.from(pdfBytes), {
      pagerender: (pageData) => pageData.getTextContent().then(tc => {
        const str = tc.items.map(i => i.str).join(' ');
        pageTexts.push(str);
        return str;
      }),
    });
  } catch (e) {
    parseFailed = true;
  }

  const src = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const total = src.getPageCount();
  const removed = [];
  const keepIdx = [];
  for (let i = 0; i < total; i++) {
    const txt = pageTexts[i] || '';
    if (INVOICE_LOG_RE.test(txt)) removed.push(i);
    else keepIdx.push(i);
  }
  const textDetected = pageTexts.some(t => t && t.trim().length > 0);
  // If parsing failed, or detection found nothing (scanned/image PDF with no
  // text layer), or matched everything, keep all pages rather than silently
  // dropping — the caller can warn.
  if (parseFailed || !textDetected || keepIdx.length === total || keepIdx.length === 0) {
    return { bytes: pdfBytes, removed: [], kept: total, total, textDetected: textDetected && !parseFailed, parseFailed };
  }
  const out = await PDFDocument.create();
  const copied = await out.copyPages(src, keepIdx);
  copied.forEach(p => out.addPage(p));
  const bytes = await out.save();
  return { bytes, removed, kept: keepIdx.length, total, textDetected: true, parseFailed: false };
}

// ═══════════════════════════════════════════════════════════════════════════
// Cover page — plain, centered, matching the CPA package's first page.
// ═══════════════════════════════════════════════════════════════════════════
async function renderCoverPdf(meta, tocEntries) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const page = pdf.addPage([PAGE.w, PAGE.h]);
  const center = (str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  // ── Cover page ────────────────────────────────────────────────────────────
  center(meta.entityName, 22, bold, 512);
  page.drawLine({ start: { x: 150, y: 494 }, end: { x: PAGE.w - 150, y: 494 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Financial Statements', 15, reg, 470);
  center(meta.longDate, 12, reg, 448);
  page.drawLine({ start: { x: 150, y: 430 }, end: { x: PAGE.w - 150, y: 430 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });

  // ── Table of Contents page (separate) with page references ────────────────
  const toc2 = pdf.addPage([PAGE.w, PAGE.h]);
  const centerOn = (pg, str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    pg.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  centerOn(toc2, meta.entityName, 13, bold, PAGE.h - PAGE.mT + 10);
  centerOn(toc2, 'Table of Contents', 15, bold, PAGE.h - 150);
  toc2.drawLine({ start: { x: 180, y: PAGE.h - 168 }, end: { x: PAGE.w - 180, y: PAGE.h - 168 }, thickness: 0.6, color: rgb(0.3, 0.3, 0.3) });
  // Fall back to a label-only list if no page references were supplied.
  const entries = (tocEntries && tocEntries.length)
    ? tocEntries
    : ['Executive Summary', 'Balance Sheets', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\u2019 Equity', 'Budget to Actual'].map(label => ({ label, page: null }));
  const LX = 120, RX = PAGE.w - 120;
  let ty = PAGE.h - 210;
  const sz = 11;
  for (const e of entries) {
    const label = e.label;
    toc2.drawText(label, { x: LX, y: ty, size: sz, font: reg, color: rgb(0.1, 0.1, 0.1) });
    if (e.page != null) {
      const num = String(e.page);
      const numW = reg.widthOfTextAtSize(num, sz);
      toc2.drawText(num, { x: RX - numW, y: ty, size: sz, font: reg, color: rgb(0.1, 0.1, 0.1) });
      // Dotted leader between label and page number.
      const labW = reg.widthOfTextAtSize(label, sz);
      const dotStart = LX + labW + 6, dotEnd = RX - numW - 6;
      const dotY = ty + 2;
      for (let dx = dotStart; dx < dotEnd; dx += 4) {
        toc2.drawText('.', { x: dx, y: dotY - 2, size: sz, font: reg, color: rgb(0.5, 0.5, 0.5) });
      }
    }
    ty -= 26;
  }
  return await pdf.save();
}

// ═══════════════════════════════════════════════════════════════════════════
// generatePackage — assemble the full merged PDF:
//   cover → executive summary (uploaded) → GL statements → requisition report
//   (uploaded, with invoice-log pages stripped).
//
// args: {
//   statements,                 // result of buildStatements
//   execSummaryBytes (optional) // uploaded exec-summary PDF
//   reqReportBytes (optional)   // uploaded requisition report — PDF or .xlsx
//   reqReportName (optional)    // original filename, used to detect .xlsx
//   reqSheetName (optional)     // worksheet to extract when .xlsx (default
//                               //   "Budget to Actual", case-insensitive; falls
//                               //   back to first sheet if not present)
// }
// Returns { bytes, info: { pages, reqRemoved, reqKept, cashFlowTies, ... } }.
// ═══════════════════════════════════════════════════════════════════════════
async function generatePackage({ statements, execSummaryBytes, reqReportBytes, reqReportName, reqSheetName }) {
  const merged = await PDFDocument.create();
  const info = { sections: [], warnings: [] };

  const appendPdf = async (bytes, label) => {
    if (!bytes) return 0;
    let srcDoc;
    try { srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true }); }
    catch (e) { info.warnings.push('Could not read ' + label + ' PDF: ' + e.message); return 0; }
    const idx = srcDoc.getPageIndices();
    const pages = await merged.copyPages(srcDoc, idx);
    pages.forEach(p => merged.addPage(p));
    info.sections.push({ label, pages: pages.length });
    return pages.length;
  };

  // ── Two-phase assembly so the Table of Contents can show real page numbers. ─
  // Phase 1: build the BODY (everything after cover+TOC) into a separate doc,
  // recording the absolute start page of each TOC section. The cover + TOC are
  // two pages, so body page N (0-based within the body) is printed page N+3.
  const COVER_TOC_PAGES = 2;
  const body = await PDFDocument.create();
  const tocEntries = [];
  const appendToBody = async (bytes, label, addToc) => {
    if (!bytes) return 0;
    let srcDoc;
    try { srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true }); }
    catch (e) { info.warnings.push('Could not read ' + label + ' PDF: ' + e.message); return 0; }
    const startPage = body.getPageCount();
    const idx = srcDoc.getPageIndices();
    const pages = await body.copyPages(srcDoc, idx);
    pages.forEach(p => body.addPage(p));
    info.sections.push({ label, pages: pages.length });
    if (addToc) tocEntries.push({ label, page: startPage + COVER_TOC_PAGES + 1 });
    return pages.length;
  };

  // Executive summary (uploaded, merged as-is).
  if (execSummaryBytes) await appendToBody(execSummaryBytes, 'Executive Summary', true);
  else { info.warnings.push('No executive summary uploaded.'); }

  // GL statements — capture each statement's start page within the statements
  // PDF, then offset by where the statements PDF lands in the body.
  const stmtOffsets = [];
  const stmtBytes = await renderStatementsPdf(statements, stmtOffsets);
  const stmtBodyStart = body.getPageCount();
  await appendToBody(stmtBytes, 'Financial Statements', false);
  for (const off of stmtOffsets) {
    tocEntries.push({ label: off.label, page: stmtBodyStart + off.page + COVER_TOC_PAGES + 1 });
  }
  // 4. Requisition report (uploaded). Accepts a PDF or an .xlsx workbook. When a
  //    workbook is uploaded, extract the requested sheet (default "Budget to
  //    Actual") and render it to a PDF page first — the rendered PDF carries a
  //    real text layer, so invoice-log stripping runs on it identically.
  if (reqReportBytes) {
    let reqPdfBytes = reqReportBytes;
    let fromXlsx = false;
    if (looksLikeXlsx(reqReportBytes, reqReportName)) {
      try {
        const wantSheet = reqSheetName || 'Budget to Actual';
        // No injected title: the requisition sheet carries its own header block
        // ("Project Funding Requisition" / entity / period / report #), and the
        // converter fits the whole sheet onto one landscape page as-is.
        const conv = await xlsxSheetToPdf(reqReportBytes, wantSheet, {});
        reqPdfBytes = Buffer.from(conv.bytes);
        fromXlsx = true;
        info.reqConvertedFromXlsx = true;
        info.reqSheetUsed = conv.sheetUsed;
        info.reqAvailableSheets = conv.availableSheets;
        if (conv.sheetUsed.toLowerCase() !== wantSheet.toLowerCase()) {
          info.warnings.push('Requisition workbook had no "' + wantSheet + '" sheet; used "' + conv.sheetUsed + '" instead.');
        }
      } catch (e) {
        info.warnings.push('Could not convert requisition workbook to PDF: ' + e.message);
        reqPdfBytes = null;
      }
    }
    if (reqPdfBytes && fromXlsx) {
      // xlsx path: we already extracted exactly the requested sheet (the Budget
      // to Actual report), so there is no invoice log to strip. Append the
      // converted page directly and skip stripInvoiceLogPages — that heuristic
      // is for multi-page PDF packets and could wrongly drop the one B2A page.
      const kept = await appendToBody(reqPdfBytes, 'Budget to Actual', true);
      info.reqRemoved = [];
      info.reqKept = kept;
      info.reqTotal = kept;
    } else if (reqPdfBytes) {
      const stripped = await stripInvoiceLogPages(reqPdfBytes);
      info.reqRemoved = stripped.removed;
      info.reqKept = stripped.kept;
      info.reqTotal = stripped.total;
      if (!stripped.textDetected) info.warnings.push(stripped.parseFailed
        ? 'Requisition PDF could not be parsed for invoice-log detection; all pages were kept.'
        : 'Requisition PDF had no extractable text; invoice-log pages could not be detected and were left in.');
      await appendToBody(stripped.bytes, 'Budget to Actual', true);
    }
  } else {
    info.warnings.push('No requisition report uploaded.');
  }

  // Phase 2: render cover + TOC (with page references) and assemble the final
  // PDF as cover/TOC first, then the body.
  const coverBytes = await renderCoverPdf(statements.meta, tocEntries);
  const coverDoc = await PDFDocument.load(coverBytes, { ignoreEncryption: true });
  const coverPages = await merged.copyPages(coverDoc, coverDoc.getPageIndices());
  coverPages.forEach(p => merged.addPage(p));
  const bodyPages = await merged.copyPages(body, body.getPageIndices());
  bodyPages.forEach(p => merged.addPage(p));

  info.cashFlowTies = statements.checks.cashFlowTies;
  info.cashFlowDiff = statements.checks.cashFlowDiff;
  info.balanceSheetTies = statements.checks.balanceSheetTies;
  info.tocEntries = tocEntries;
  info.pages = merged.getPageCount();
  const bytes = await merged.save();
  return { bytes, info };
}

// ═══════════════════════════════════════════════════════════════════════════
// buildTtmPL — Trailing-12-Months P&L matrix.
//
// Produces a P&L (Statements-of-Operations grouping) with 12 monthly columns
// (oldest → newest, ending at asOf) plus a Total column that sums the 12 months.
// Pure given getBalances({from,to}); no I/O.
//   opts: { asOf, entityName }
// ═══════════════════════════════════════════════════════════════════════════
async function buildTtmPL(getBalances, opts) {
  const asOf = opts.asOf;
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const end = addMonthsEnd(asOf, -i);
    months.push({ from: monthStart(end), to: end, label: monthLabel(end), end });
  }
  const monthRows = await Promise.all(months.map(mo => getBalances({ from: mo.from, to: mo.to })));

  const maps = monthRows.map(rows => {
    const m = new Map();
    for (const r of rows) if (r.type === 'Revenue' || r.type === 'Expense') m.set(String(r.code), r);
    return m;
  });
  const refByCode = new Map();
  for (let i = maps.length - 1; i >= 0; i--) {
    for (const [code, r] of maps[i]) if (!refByCode.has(code)) refByCode.set(code, r);
  }
  const allCodes = Array.from(refByCode.keys())
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  function lineFor(code) {
    const ref = refByCode.get(code);
    const vals = maps.map(m => { const r = m.get(code); return r ? r2(bal(r)) : 0; });
    const total = r2(vals.reduce((s, v) => s + v, 0));
    if (vals.every(isZero) && isZero(total)) return null;
    return { code, name: ref.name, type: ref.type, subtype: ref.subtype, vals, total };
  }
  const allLines = allCodes.map(lineFor).filter(Boolean);

  const revenue = allLines.filter(l => l.type === 'Revenue');
  const cogs = allLines.filter(l => l.type === 'Expense' && /cogs|cost of goods|cost of revenue|car hire/i.test((l.subtype || '') + ' ' + (l.name || '')));
  const cogsCodes = new Set(cogs.map(l => l.code));
  const opex = allLines.filter(l => l.type === 'Expense' && !cogsCodes.has(l.code));

  const sumLines = lines => {
    const vals = new Array(12).fill(0);
    let total = 0;
    for (const l of lines) { for (let i = 0; i < 12; i++) vals[i] += l.vals[i]; total += l.total; }
    return { vals: vals.map(r2), total: r2(total) };
  };
  const totRev = sumLines(revenue);
  const totCogs = sumLines(cogs);
  const grossProfit = { vals: totRev.vals.map((v, i) => r2(v - totCogs.vals[i])), total: r2(totRev.total - totCogs.total) };
  const totOpex = sumLines(opex);
  const netIncome = { vals: grossProfit.vals.map((v, i) => r2(v - totOpex.vals[i])), total: r2(grossProfit.total - totOpex.total) };

  const opexByCat = new Map();
  for (const l of opex) {
    const cat = plExpenseCategory(l);
    if (!opexByCat.has(cat)) opexByCat.set(cat, []);
    opexByCat.get(cat).push(l);
  }
  const opexGroups = [];
  const pushGroup = (cat, lines) => {
    lines.sort((a, b) => String(a.code).localeCompare(String(b.code), undefined, { numeric: true }));
    opexGroups.push({ title: cat, lines, subtotal: sumLines(lines) });
  };
  for (const cat of PL_EXPENSE_CATEGORY_ORDER) {
    const lines = opexByCat.get(cat);
    if (lines && lines.length) pushGroup(cat, lines);
    opexByCat.delete(cat);
  }
  for (const [cat, lines] of opexByCat) if (lines && lines.length) pushGroup(cat, lines);

  return {
    meta: {
      entityName: opts.entityName || '',
      asOf,
      title: 'Trailing 12 Months',
      periodLabel: 'For the Trailing Twelve Months Ended ' + longDate(asOf),
      months: months.map(m => ({ from: m.from, to: m.to, label: m.label })),
      totalLabel: 'Total',
    },
    revenue, totRev,
    cogs, totCogs, grossProfit, hasCogs: cogs.length > 0,
    opex, opexGroups, totOpex,
    netIncome,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// buildTtmAnalysis — surfaces line items that need attention from the 12-month
// series. Focus is on UNFAVORABLE movements in the most recent month vs. the
// average of the prior active months:
//   • Expense up  = latest month materially above its trailing average (spending
//                   is rising) → unfavorable.
//   • Revenue down= latest month materially below its trailing average (revenue
//                   is falling) → unfavorable.
// Plus overall Net Income context: any loss months, and whether the latest month
// deteriorated vs. its trailing average. Thresholds keep noise down: a movement
// must clear both a percentage swing and a dollar floor to be flagged.
// ─────────────────────────────────────────────────────────────────────────────
function buildTtmAnalysis(ctx) {
  const { months, revenue, opex, opexGroups, netIncome } = ctx;
  const n = 12;
  const lastIdx = n - 1;
  const lastLabel = months[lastIdx] ? months[lastIdx].label : '';
  const PCT_THRESHOLD = 0.25;   // ≥25% swing vs. trailing average
  const DOLLAR_FLOOR = 500;     // and ≥ $500 movement, to ignore trivial lines

  // Trailing average of the prior months (indices 0..lastIdx-1), counting only
  // months where the line was active (non-zero), so a line that only recently
  // started isn't compared against a diluted average full of zeros.
  const trailingAvg = vals => {
    const prior = vals.slice(0, lastIdx).filter(v => Math.abs(v) >= 0.005);
    if (!prior.length) return null;
    return prior.reduce((s, v) => s + v, 0) / prior.length;
  };

  const items = [];

  // Expense lines (individual accounts) — flag latest month running hot.
  for (const l of opex) {
    const last = l.vals[lastIdx];
    const avg = trailingAvg(l.vals);
    if (avg == null || avg <= 0) continue;
    const delta = r2(last - avg);
    if (delta >= DOLLAR_FLOOR && (delta / avg) >= PCT_THRESHOLD) {
      items.push({
        severity: (delta / avg) >= 0.75 ? 'high' : 'medium',
        kind: 'expense_up',
        name: l.name,
        code: l.code,
        detail: lastLabel + ' expense of ' + fmtAbs(last) + ' is ' + pctStr(delta / avg)
          + ' above the trailing average of ' + fmtAbs(avg) + ' (up ' + fmtAbs(delta) + ').',
        last: r2(last), avg: r2(avg), delta,
      });
    }
  }

  // Revenue lines — flag latest month running cold (revenue falling).
  for (const l of revenue) {
    const last = l.vals[lastIdx];
    const avg = trailingAvg(l.vals);
    if (avg == null || avg <= 0) continue;
    const delta = r2(avg - last); // positive = shortfall
    if (delta >= DOLLAR_FLOOR && (delta / avg) >= PCT_THRESHOLD) {
      items.push({
        severity: (delta / avg) >= 0.75 ? 'high' : 'medium',
        kind: 'revenue_down',
        name: l.name,
        code: l.code,
        detail: lastLabel + ' revenue of ' + fmtAbs(last) + ' is ' + pctStr(delta / avg)
          + ' below the trailing average of ' + fmtAbs(avg) + ' (down ' + fmtAbs(delta) + ').',
        last: r2(last), avg: r2(avg), delta,
      });
    }
  }

  // Sort most material first (largest dollar movement), high severity first.
  const sevRank = { high: 0, medium: 1 };
  items.sort((a, b) => (sevRank[a.severity] - sevRank[b.severity]) || (Math.abs(b.delta) - Math.abs(a.delta)));

  // Net income context.
  const niContext = [];
  const lossMonths = [];
  for (let i = 0; i < n; i++) {
    if (netIncome.vals[i] < -0.005) lossMonths.push(months[i] ? months[i].label : ('M' + (i + 1)));
  }
  if (lossMonths.length) {
    niContext.push({
      severity: lossMonths.length >= 3 ? 'high' : 'medium',
      kind: 'net_loss_months',
      detail: 'Net loss in ' + lossMonths.length + ' of ' + n + ' months: ' + lossMonths.join(', ') + '.',
    });
  }
  const niAvg = trailingAvg(netIncome.vals);
  const niLast = netIncome.vals[lastIdx];
  if (niAvg != null) {
    const niDelta = r2(niLast - niAvg);
    if (niDelta <= -DOLLAR_FLOOR && Math.abs(niAvg) >= 0.005 && (Math.abs(niDelta) / Math.abs(niAvg)) >= PCT_THRESHOLD) {
      niContext.push({
        severity: 'medium',
        kind: 'net_income_down',
        detail: lastLabel + ' net income of ' + fmtSigned(niLast) + ' is ' + fmtAbs(Math.abs(niDelta))
          + ' below the trailing average of ' + fmtSigned(niAvg) + '.',
      });
    }
  }

  return {
    lastMonthLabel: lastLabel,
    thresholds: { pct: PCT_THRESHOLD, dollar: DOLLAR_FLOOR },
    items,          // unfavorable line-level movements
    netIncome: niContext,
    hasFindings: items.length > 0 || niContext.length > 0,
  };
}

// Format helpers for analysis narrative (magnitude with commas, and % / signed).
function fmtAbs(n) { return Math.abs(Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtSigned(n) { const v = Number(n) || 0; const t = fmtAbs(v); return v < 0 ? '(' + t + ')' : t; }
function pctStr(frac) { return Math.round((Number(frac) || 0) * 100) + '%'; }

// Short month-column label, e.g. '2026-04-30' → 'Apr 2026'.
function monthLabel(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  const ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return ABBR[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}

// ═══════════════════════════════════════════════════════════════════════════
// buildFundStatements — CLRF-style limited-partnership fund statement package.
//
// Produces the numeric model for five statements:
//   1. Statement of Assets, Liabilities, and Partners' Capital
//   2. Schedule of Investments (from the fund_investments config table)
//   3. Statement of Operations (investment-company format)
//   4. Statement of Changes in Partners' Capital (GP / LP columns)
//   5. Statement of Cash Flows (indirect)
//
// Pure given the injected data accessors; no direct DB/PDF I/O. opts:
//   { asOf, entityName,
//     getBalances(o),          // -> computeBalances(eid, o)
//     investments: [ { parent_name, name, acquisition_date, cost, fair_value, sort_order } ],
//     partnerClasses: [ { id, name, partner_type } ]  // GP/LP tag per class
//   }
//
// Fund-specific GL mapping (CLRF entity 40), verified against the Q1-2026 CPA
// statements:
//   Portfolio investments at fair value = 1201xx + 1202xx + 1218xx (unrealized)
//   Operations expense buckets: Management fees = 5101xx; Professional fees =
//     52xxxx; Broken deal costs = 5302xx; Other expenses = everything else in
//     Expense that isn't one of the above and isn't the 6101xx unrealized G/L
//     contra (which belongs to the investment fair-value mark, not operations).
// ═══════════════════════════════════════════════════════════════════════════
async function buildFundStatements(opts) {
  const asOf = opts.asOf;
  const ys = yearStart(asOf);
  const getBalances = opts.getBalances;
  const investments = (opts.investments || []).slice().sort((a, b) => (a.sort_order - b.sort_order) || (a.id - b.id));
  const partnerClasses = opts.partnerClasses || [];
  const gpClassIds = new Set(partnerClasses.filter(c => String(c.partner_type).toUpperCase() === 'GP').map(c => c.id));

  const priorBsDate = priorMonthEnd(ys); // beginning-of-year balance sheet date

  // Snapshots. Current + beginning-of-period balance sheets (prior years closed
  // into RE so RE holds the opening balance); YTD P&L drives Operations + CF.
  const [bsCur, bsBeg, isYtd] = await Promise.all([
    getBalances({ as_of: asOf, close_pl_before: ys }),
    getBalances({ as_of: priorBsDate, close_pl_before: yearStart(priorBsDate) }),
    getBalances({ from: ys, to: asOf }),
  ]);

  const byCode = rows => { const m = new Map(); for (const r of rows) m.set(String(r.code), r); return m; };
  const curMap = byCode(bsCur), begMap = byCode(bsBeg);

  // ── helpers ────────────────────────────────────────────────────────────────
  const codeStarts = (code, prefixes) => prefixes.some(p => String(code).startsWith(p));
  const sumWhere = (rows, pred) => r2(rows.filter(pred).reduce((s, r) => s + bal(r), 0));

  const niYtd = netIncomeOf(isYtd); // net investment loss for the period

  // ── 1. Statement of Assets, Liabilities & Partners' Capital ─────────────────
  // Investment accounts collapse into a single "Portfolio investments, at fair
  // value" line (cost + capitalized + unrealized), with the cost basis shown
  // parenthetically. Everything else lists individually by account.
  const INVEST_PREFIXES = ['1201', '1202', '1218'];
  const isInvest = code => codeStarts(code, INVEST_PREFIXES);
  const investCostCode = code => codeStarts(code, ['1201', '1202']); // cost (excl. unrealized mark)

  function assetRows(map) {
    const rows = [];
    let investFV = 0, investCost = 0;
    for (const r of map.values()) {
      if (r.type !== 'Asset') continue;
      if (isInvest(r.code)) { investFV += bal(r); if (investCostCode(r.code)) investCost += bal(r); continue; }
      if (isZero(bal(r))) continue;
      rows.push({ code: String(r.code), name: r.name, amount: r2(bal(r)) });
    }
    rows.sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
    return { rows, investFV: r2(investFV), investCost: r2(investCost) };
  }
  const curAssets = assetRows(curMap);

  // Assets are presented investments-first, then the rest in code order.
  const assetLines = [];
  if (!isZero(curAssets.investFV) || !isZero(curAssets.investCost)) {
    assetLines.push({ name: 'Portfolio investments, at fair value (cost ' + acct(curAssets.investCost, { dash: false }) + ')', amount: curAssets.investFV, invest: true });
  }
  for (const r of curAssets.rows) assetLines.push({ name: r.name, amount: r.amount });
  const totalAssets = r2(assetLines.reduce((s, l) => s + l.amount, 0));

  // Liabilities: list each non-zero liability account by name, code order.
  const liabLines = [];
  for (const r of curMap.values()) {
    if (r.type !== 'Liability' || isZero(bal(r))) continue;
    liabLines.push({ code: String(r.code), name: r.name, amount: r2(bal(r)) });
  }
  liabLines.sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
  const totalLiab = r2(liabLines.reduce((s, l) => s + l.amount, 0));

  // Partners' capital total = all Equity accounts + current-period net income
  // (which is still open on the income accounts in the close_pl_before snapshot).
  const totalCapital = r2(sumWhere(bsCur, r => r.type === 'Equity') + niYtd);

  // GP vs LP split of the ending capital, from per-class capital roll-forward
  // (computed below). Falls back to 0/total if no classes are tagged.
  // (populated after the changes-in-capital section.)

  // ── 3. Statement of Operations (investment-company format) ──────────────────
  const invIncome = sumWhere(isYtd, r => r.type === 'Revenue');
  const expRows = isYtd.filter(r => r.type === 'Expense');
  const isMgmtFee = code => codeStarts(code, ['5101']);
  const isProfFee = code => codeStarts(code, ['52']);
  const isBrokenDeal = code => codeStarts(code, ['5302']);
  const isUnrealizedContra = code => codeStarts(code, ['6101']); // fair-value mark, not an operating expense
  const mgmtFees = sumWhere(expRows, r => isMgmtFee(r.code));
  const profFees = sumWhere(expRows, r => isProfFee(r.code));
  const brokenDeal = sumWhere(expRows, r => isBrokenDeal(r.code));
  const otherExp = sumWhere(expRows, r => !isMgmtFee(r.code) && !isProfFee(r.code) && !isBrokenDeal(r.code) && !isUnrealizedContra(r.code));
  const totalExpenses = r2(mgmtFees + profFees + brokenDeal + otherExp);
  const netInvestmentLoss = r2(invIncome - totalExpenses);
  // Net change in unrealized appreciation flows to operations only if the fund
  // presents it there; CLRF Q1 shows net investment loss = decrease in capital
  // from operations, so we mirror that (unrealized handled within investments).
  const netOpsResult = netInvestmentLoss;

  const operations = {
    investmentIncome: [
      // No itemized investment income in the CLRF format when zero; the total line carries it.
    ],
    totalInvestmentIncome: r2(invIncome),
    expenses: [
      { name: 'Management fees', amount: mgmtFees },
      { name: 'Professional fees', amount: profFees },
      { name: 'Broken deal costs', amount: brokenDeal },
      { name: 'Other expenses', amount: otherExp },
    ].filter(e => !isZero(e.amount)),
    totalExpenses,
    netInvestmentLoss,
    netResult: netOpsResult,
  };

  // ── 4. Statement of Changes in Partners' Capital (GP / LP) ──────────────────
  // Per-class capital activity for the period, grouped GP vs LP. Each class's
  // beginning capital = equity accounts at beginning-of-year (prior closed);
  // period activity = equity-account movements over the period, split into the
  // standard fund columns by account meaning; net loss = the class's share of
  // net investment loss (its P&L movement over the period).
  //
  // Column mapping by equity account meaning (CLRF):
  //   Contributions       = 3011xx/3012xx/3013xx/3018xx positive movements
  //   Capital call refunds = negative movements of the contribution accounts
  //   Syndication costs   = 3702xx movement
  //   Net loss (ops)      = class P&L over the period (Revenue - Expense)
  // Because a single account (contributions) holds both calls and refunds net,
  // we present the net movement in the Contributions column and rely on the GL
  // sign; the fund-level totals still tie to the GL exactly.
  const CONTRIB_PREFIXES = ['3011', '3012', '3013', '3018'];
  const SYND_PREFIX = ['3702'];

  // Fund-level column totals come straight from the GL and tie exactly:
  //   beginning  = fund equity at beginning of year (prior years closed)
  //   contributions/refunds (net) = movement of the contribution accounts over the period
  //   syndication = movement of the syndication-cost account over the period
  //   netLoss     = fund net investment loss for the period
  //   ending      = beginning + contributions + syndication + netLoss (ties to totalCapital)
  const fundBegEquity = r2(sumWhere(bsBeg, r => r.type === 'Equity'));
  const contribMove = r2(sumWhere(isYtd, r => r.type === 'Equity' && codeStarts(r.code, CONTRIB_PREFIXES)));
  const syndMove = r2(sumWhere(isYtd, r => r.type === 'Equity' && codeStarts(r.code, SYND_PREFIX)));
  // Any other equity movement over the period (e.g. reclasses) so the columns
  // still foot to the true ending capital.
  const otherEqMove = r2(sumWhere(isYtd, r => r.type === 'Equity' && !codeStarts(r.code, CONTRIB_PREFIXES) && !codeStarts(r.code, SYND_PREFIX)));
  const fundTotals = {
    beginning: fundBegEquity,
    contributions: contribMove,
    syndication: syndMove,
    other: otherEqMove,
    netLoss: r2(niYtd),
    ending: r2(fundBegEquity + contribMove + syndMove + otherEqMove + niYtd),
  };

  // GP/LP split. The GL's equity/contribution accounts are not reliably tagged
  // by investor class at a point in time (a per-class as-of snapshot returns the
  // whole fund), but PERIOD MOVEMENTS are class-tagged. So we split the activity
  // columns (contributions, syndication, net loss) by summing period movements
  // for GP-tagged classes, and split the beginning balance by each group's share
  // of total capital commitments when available; otherwise beginning is all-LP.
  // Ending = beginning + activity per group, and GP+LP always re-foot to the
  // fund totals exactly (LP is computed as fund − GP).
  async function classPeriodMove(classId) {
    const pl = await getBalances({ from: ys, to: asOf, class_id: classId });
    const contrib = r2(sumWhere(pl, r => r.type === 'Equity' && codeStarts(r.code, CONTRIB_PREFIXES)));
    const synd = r2(sumWhere(pl, r => r.type === 'Equity' && codeStarts(r.code, SYND_PREFIX)));
    const other = r2(sumWhere(pl, r => r.type === 'Equity' && !codeStarts(r.code, CONTRIB_PREFIXES) && !codeStarts(r.code, SYND_PREFIX)));
    const netLoss = netIncomeOf(pl);
    return { contrib, synd, other, netLoss };
  }

  // GP activity from period movements of GP-tagged classes.
  const gpAgg = { contributions: 0, syndication: 0, other: 0, netLoss: 0 };
  let anyClassData = false;
  for (const c of partnerClasses) {
    if (!gpClassIds.has(c.id)) continue;
    const m = await classPeriodMove(c.id);
    if (!isZero(m.contrib) || !isZero(m.synd) || !isZero(m.other) || !isZero(m.netLoss)) anyClassData = true;
    gpAgg.contributions = r2(gpAgg.contributions + m.contrib);
    gpAgg.syndication = r2(gpAgg.syndication + m.synd);
    gpAgg.other = r2(gpAgg.other + m.other);
    gpAgg.netLoss = r2(gpAgg.netLoss + m.netLoss);
  }

  // Beginning-capital GP share: use GP classes' commitment share of total
  // commitments if commitments are loaded; else 0 (all LP). This is the only
  // beginning-balance signal available since per-class snapshots aren't reliable.
  let gpBegShare = 0;
  const commitments = opts.commitments || [];
  if (commitments.length && gpClassIds.size) {
    const totalCommit = commitments.reduce((s, x) => s + (Number(x.commitment_amount) || 0), 0);
    const gpCommit = commitments.filter(x => gpClassIds.has(x.class_id)).reduce((s, x) => s + (Number(x.commitment_amount) || 0), 0);
    if (totalCommit > 0) gpBegShare = gpCommit / totalCommit;
  }
  const gpBeginning = r2(fundTotals.beginning * gpBegShare);

  const groups = {
    GP: {
      beginning: gpBeginning,
      contributions: gpAgg.contributions,
      syndication: gpAgg.syndication,
      netLoss: gpAgg.netLoss,
      ending: r2(gpBeginning + gpAgg.contributions + gpAgg.syndication + gpAgg.other + gpAgg.netLoss),
    },
    LP: {},
  };
  // LP = fund − GP for every column, so the two columns always re-foot exactly.
  groups.LP = {
    beginning: r2(fundTotals.beginning - groups.GP.beginning),
    contributions: r2(fundTotals.contributions - groups.GP.contributions),
    syndication: r2(fundTotals.syndication - groups.GP.syndication),
    netLoss: r2(fundTotals.netLoss - groups.GP.netLoss),
    ending: r2(fundTotals.ending - groups.GP.ending),
  };
  const capTotals = {
    beginning: fundTotals.beginning,
    contributions: fundTotals.contributions,
    syndication: fundTotals.syndication,
    netLoss: fundTotals.netLoss,
    ending: fundTotals.ending,
  };

  // GP/LP ending split for the balance-sheet capital section.
  const capGP = groups.GP.ending;
  const capLP = groups.LP.ending;

  // ── 2. Schedule of Investments (from config) ────────────────────────────────
  // Group underlyings by parent_name (holding company); a blank parent means the
  // investment stands alone. Percentages are of total partners' capital.
  const invByParent = new Map();
  for (const inv of investments) {
    const key = inv.parent_name || '';
    if (!invByParent.has(key)) invByParent.set(key, []);
    invByParent.get(key).push(inv);
  }
  const scheduleGroups = [];
  let schedTotCost = 0, schedTotFV = 0;
  for (const [parent, list] of invByParent) {
    const rows = list.map(inv => ({
      name: inv.name,
      acquisition_date: inv.acquisition_date || '',
      cost: r2(inv.cost),
      fair_value: r2(inv.fair_value),
      pctCapital: totalCapital ? r2((inv.fair_value / totalCapital) * 100) : 0,
    }));
    const subCost = r2(rows.reduce((s, r) => s + r.cost, 0));
    const subFV = r2(rows.reduce((s, r) => s + r.fair_value, 0));
    schedTotCost = r2(schedTotCost + subCost);
    schedTotFV = r2(schedTotFV + subFV);
    scheduleGroups.push({
      parent, rows,
      subtotal: { cost: subCost, fair_value: subFV, pctCapital: totalCapital ? r2((subFV / totalCapital) * 100) : 0 },
    });
  }
  const schedule = {
    groups: scheduleGroups,
    total: { cost: schedTotCost, fair_value: schedTotFV, pctCapital: totalCapital ? r2((schedTotFV / totalCapital) * 100) : 0 },
    hasData: investments.length > 0,
  };

  // ── 5. Statement of Cash Flows (indirect) ───────────────────────────────────
  // Beginning balances at year start − 1 day; deltas over the YTD window. Cash
  // is every asset account whose name/code reads as cash & equivalents.
  const isCashCode = code => codeStarts(code, ['1002', '1005', '1072']);
  const cashCur = sumWhere(bsCur, r => r.type === 'Asset' && isCashCode(r.code));
  const cashBeg = sumWhere(bsBeg, r => r.type === 'Asset' && isCashCode(r.code));

  // Working-capital deltas: change in each non-cash, non-investment asset and
  // each liability over the period. Sign for cash-flow: asset increase uses
  // cash (negative), liability increase provides cash (positive).
  const deltaAsset = code => r2((curMap.get(code) ? bal(curMap.get(code)) : 0) - (begMap.get(code) ? bal(begMap.get(code)) : 0));
  const allAssetCodes = new Set([...curMap.keys(), ...begMap.keys()].filter(c => {
    const ref = curMap.get(c) || begMap.get(c); return ref && ref.type === 'Asset';
  }));
  const allLiabCodes = new Set([...curMap.keys(), ...begMap.keys()].filter(c => {
    const ref = curMap.get(c) || begMap.get(c); return ref && ref.type === 'Liability';
  }));

  let wcAssets = 0, wcLiab = 0, investChange = 0;
  for (const c of allAssetCodes) {
    if (isCashCode(c)) continue;
    if (isInvest(c)) { investChange = r2(investChange + deltaAsset(c)); continue; }
    wcAssets = r2(wcAssets - deltaAsset(c)); // asset increase reduces cash
  }
  for (const c of allLiabCodes) {
    const d = r2((curMap.get(c) ? bal(curMap.get(c)) : 0) - (begMap.get(c) ? bal(begMap.get(c)) : 0));
    wcLiab = r2(wcLiab + d); // liability increase provides cash
  }
  // Equity contributions/refunds/syndication over the period = financing.
  const eqCur = sumWhere(bsCur, r => r.type === 'Equity');
  const eqBegAll = sumWhere(bsBeg, r => r.type === 'Equity');
  const financingEquity = r2(eqCur - eqBegAll);

  const netOperating = r2(niYtd + wcAssets + wcLiab);
  const netInvesting = r2(-investChange); // investment asset increase = cash outflow
  const netFinancing = r2(financingEquity);
  const netChange = r2(netOperating + netInvesting + netFinancing);
  const cashTieOut = r2(cashCur - (cashBeg + netChange));

  const cashFlow = {
    netLoss: niYtd,
    changeWCAssets: wcAssets,
    changeWCLiab: wcLiab,
    netOperating,
    investChange: r2(-investChange),
    netInvesting,
    financingEquity,
    netFinancing,
    netChange,
    cashBeg: r2(cashBeg),
    cashEnd: r2(cashCur),
    tieOut: cashTieOut,
  };

  return {
    meta: {
      entityName: opts.entityName || '',
      asOf,
      longDate: longDate(asOf),
      periodLabel: 'For the Quarter Ended ' + longDate(asOf),
      title: 'Financial Statements',
    },
    assetsLiabCapital: {
      assetLines, totalAssets,
      liabLines, totalLiab,
      capGP: r2(capGP), capLP: r2(capLP), totalCapital,
      totalLiabCapital: r2(totalLiab + totalCapital),
    },
    schedule,
    operations,
    changesInCapital: { groups, totals: capTotals, hasClassData: anyClassData },
    cashFlow,
    _tie: { totalAssets, totalLiabPlusCapital: r2(totalLiab + totalCapital), niYtd },
  };
}


// ═══════════════════════════════════════════════════════════════════════════
// renderFundStatementsPdf — render the CLRF fund package to PDF bytes, using the
// same makeLayout primitives as the corporate package so the styling matches.
// Statements (each its own page sequence): Assets/Liabilities/Partners' Capital,
// Schedule of Investments, Statement of Operations, Statement of Changes in
// Partners' Capital (landscape), Statement of Cash Flows.
// If outOffsets is passed it is filled with { label, page } for the TOC.
// ═══════════════════════════════════════════════════════════════════════════
async function renderFundStatementsPdf(s, outOffsets) {
  const track = (label) => { if (outOffsets) outOffsets.push({ label, page: pdf.getPageCount() }); };
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fonts = { reg, bold };
  const m = s.meta;
  const money = v => acct(v, { dash: true });
  const RIGHT = PAGE.w - PAGE.mR;
  const oneCol = [RIGHT];

  // ── 1. Statement of Assets, Liabilities, and Partners' Capital ──────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Assets, Liabilities, and Partners\u2019 Capital', { dateLine: m.longDate });
    track('Statement of Assets, Liabilities, and Partners\u2019 Capital');
    L.start();
    L.setCols(oneCol);
    const a = s.assetsLiabCapital;
    L.sectionTitle('Assets');
    a.assetLines.forEach((l, i) => L.row(l.name, [money(l.amount)], { indent: 16, dollarPrefix: i === 0 }));
    L.row('Total assets', [money(a.totalAssets)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: true, gapAfter: 10 });

    L.sectionTitle('Liabilities and partners\u2019 capital');
    L.row('Liabilities:', [], { indent: 10, boldRow: true });
    a.liabLines.forEach((l, i) => L.row(l.name, [money(l.amount)], { indent: 20, dollarPrefix: i === 0 }));
    L.row('Total liabilities', [money(a.totalLiab)], { indent: 10, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.row('Partners\u2019 capital:', [], { indent: 10, boldRow: true });
    if (!isZero(a.capGP) || a.capGP !== a.totalCapital) L.row('General partners', [money(a.capGP)], { indent: 20 });
    L.row('Limited partners', [money(a.capLP)], { indent: 20 });
    L.row('Total partners\u2019 capital', [money(a.totalCapital)], { indent: 10, boldRow: true, ruleAbove: true, gapAfter: 8 });
    L.row('Total liabilities and partners\u2019 capital', [money(a.totalLiabCapital)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: true });
  }

  // ── 2. Schedule of Investments ──────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Schedule of Investments', { dateLine: m.longDate });
    track('Schedule of Investments');
    L.start();
    // Columns: Acquisition date | Cost | Fair Value | % of Capital
    const sCols = [RIGHT - 300, RIGHT - 180, RIGHT - 60, RIGHT];
    L.setCols(sCols);
    L.colHeaders(['Date of\nAcquisition', 'Cost', 'Fair Value', '% of\nCapital'], { bottomAlign: true, underline: true, colBox: true });
    const sch = s.schedule;
    if (!sch.hasData) {
      L.row('No investment detail configured. Add underlyings in Fund Reporting settings.', [], { indent: 10 });
    } else {
      const pct = v => (Number(v) || 0).toFixed(2);
      for (const g of sch.groups) {
        if (g.parent) L.row(g.parent, [], { indent: 6, boldRow: true });
        const rowIndent = g.parent ? 16 : 10;
        for (const r of g.rows) {
          L.row(r.name, [r.acquisition_date, money(r.cost), money(r.fair_value), pct(r.pctCapital)], { indent: rowIndent });
        }
        if (g.parent && g.rows.length > 1) {
          L.row('Total ' + g.parent, ['', money(g.subtotal.cost), money(g.subtotal.fair_value), pct(g.subtotal.pctCapital)], { indent: 10, boldRow: true, ruleAbove: true });
        }
      }
      L.row('Total investments', ['', money(sch.total.cost), money(sch.total.fair_value), pct(sch.total.pctCapital) + ' %'], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
    }
  }

  // ── 3. Statement of Operations ──────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Operations', { dateLine: m.periodLabel || ('For the Period Ended ' + m.longDate) });
    track('Statement of Operations');
    L.start();
    L.setCols(oneCol);
    const o = s.operations;
    L.sectionTitle('Investment income:');
    o.investmentIncome.forEach(r => L.row(r.name, [money(r.amount)], { indent: 16 }));
    L.row('Total investment income', [money(o.totalInvestmentIncome)], { indent: 16, boldRow: true, dollarPrefix: true, gapAfter: 6 });
    L.sectionTitle('Expenses:');
    o.expenses.forEach(r => L.row(r.name, [money(r.amount)], { indent: 16 }));
    L.row('Total expenses', [money(o.totalExpenses)], { indent: 16, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Net investment loss', [money(o.netInvestmentLoss)], { indent: 6, boldRow: true, ruleAbove: true });
    L.row('Net decrease in partners\u2019 capital resulting from operations', [money(o.netResult)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: true });
  }

  // ── 4. Statement of Changes in Partners' Capital (landscape, GP/LP/Total) ────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Partners\u2019 Capital', { landscape: true, dateLine: m.periodLabel || ('For the Period Ended ' + m.longDate) });
    const LRIGHT = PAGE.h - PAGE.mR;
    track('Statement of Changes in Partners\u2019 Capital');
    L.start();
    const c1 = 300;
    const PITCH = (LRIGHT - c1) / 2;
    const cols = [c1, c1 + PITCH, LRIGHT];
    L.setCols(cols);
    L.colHeaders(['General Partners', 'Limited Partners', 'Total'], { bottomAlign: true, underline: true, colBox: true });
    const cc = s.changesInCapital;
    const g = cc.groups, t = cc.totals;
    const rowvals = (key) => [acct(g.GP[key]), acct(g.LP[key]), acct(t[key])];
    L.row('Balance, beginning of period', rowvals('beginning'), { indent: 10, valueInset: 4 });
    L.row('Capital contributions (refunds), net', rowvals('contributions'), { indent: 10, valueInset: 4 });
    if (!isZero(t.syndication)) L.row('Syndication costs', rowvals('syndication'), { indent: 10, valueInset: 4 });
    L.row('Net decrease resulting from operations', rowvals('netLoss'), { indent: 10, valueInset: 4 });
    L.row('Balance, end of period', rowvals('ending'), { indent: 10, boldRow: true, ruleAbove: true, doubleBelow: true, valueInset: 4 });
    if (!cc.hasClassData) {
      L.space(8);
      L.row('Note: GP/LP split requires investor classes tagged in Fund Reporting settings; shown as Limited Partners until tagged.', [], { indent: 10 });
    }
  }

  // ── 5. Statement of Cash Flows ──────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows', { dateLine: m.periodLabel || ('For the Period Ended ' + m.longDate) });
    track('Statement of Cash Flows');
    L.start();
    L.setCols(oneCol);
    const cf = s.cashFlow;
    L.sectionTitle('Operating activities:');
    L.row('Net decrease in partners\u2019 capital resulting from operations', [money(cf.netLoss)], { indent: 16, dollarPrefix: true });
    L.row('Adjustments to reconcile to net cash used in operating activities:', [], { indent: 16 });
    if (!isZero(cf.investChange)) L.row('Purchases of investments, net of returns of capital', [money(cf.investChange)], { indent: 28 });
    if (!isZero(cf.changeWCAssets)) L.row('Changes in operating assets', [money(cf.changeWCAssets)], { indent: 28 });
    if (!isZero(cf.changeWCLiab)) L.row('Changes in operating liabilities', [money(cf.changeWCLiab)], { indent: 28 });
    L.row('Net cash used in operating activities', [money(r2(cf.netOperating + cf.investChange))], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.sectionTitle('Financing activities:');
    if (!isZero(cf.financingEquity)) L.row('Capital contributions, net of refunds and syndication costs', [money(cf.financingEquity)], { indent: 28 });
    L.row('Net cash provided by financing activities', [money(cf.netFinancing)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.row('Net change in cash and cash equivalents', [money(cf.netChange)], { indent: 6, boldRow: true, ruleAbove: true });
    L.row('Cash and cash equivalents, beginning of period', [money(cf.cashBeg)], { indent: 6 });
    L.row('Cash and cash equivalents, end of period', [money(cf.cashEnd)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: true });
    if (!isZero(cf.tieOut)) {
      L.space(6);
      L.row('Note: reconciled change differs from cash movement by ' + money(cf.tieOut) + '.', [], { indent: 6 });
    }
  }

  return await pdf.save();
}


module.exports = {
  buildStatements,
  buildTtmPL,
  buildFundStatements,
  renderFundStatementsPdf,
  generatePackage,
  renderStatementsPdf,
  stripInvoiceLogPages,
  // exported for unit tests / reuse
  _helpers: { acct, r2, isZero, netIncomeOf, bsSection, priorMonthEnd, yearStart, monthStart, monthsEndedLabel, longDate, resolvePeriod },
};
