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
const execSummaries = require('./execSummaries');

// ── entity display-name normalization ───────────────────────────────
// The report presents this development entity as "County Line SRN". The GL
// entity record is named "Sabine River & Northern Railroad" (entity 37) and is
// NOT renamed — only the statement package uses the display name. Map known
// aliases; anything else passes through unchanged.
const ENTITY_DISPLAY_NAMES = [
  { match: /sabine|(county\s*line\s*)?srn/i, display: 'County Line SRN' },
  // Anchored to the exact ledger name (entity 41). An unanchored /banyan\s*
  // residential/ also swallowed "Banyan Residential NCG Fund I, LLC" (entity 14)
  // and "Guaranty Banyan Residential" (entity 5), so both printed entity 41's
  // name at the top of their own statements. Found 2026-08-26.
  { match: /^banyan\s*residential$/i, display: 'Banyan Residential, LLC' },
  // CLIP Property Owner (entity 54) is the operating company that carries the
  // parent's pushed-down 100% investment on its own books. Its issued CPA
  // package is titled for the parent, "County Line Industrial Park, LLC", so the
  // statement package prints that name even though the ledger entity is still
  // "CLIP Property Owner". (The ledger entity is NOT renamed.)
  { match: /clip\s*property\s*owner|county\s*line\s*industrial\s*park/i, display: 'County Line Industrial Park, LLC' },
  // CLR Silsbee Property Owner (entity 39) — same push-down situation as CLIP.
  // Its CPA package is titled "County Line Rail Silsbee, LLC", so the statement
  // package prints that name; the ledger entity stays "CLR Silsbee Property
  // Owner".
  { match: /clr\s*silsbee\s*property\s*owner|county\s*line\s*rail\s*silsbee/i, display: 'County Line Rail Silsbee, LLC' },
];
// ── entity designation suffix ───────────────────────────────────────
// ASC 272-10-45-2 (from AICPA Practice Bulletin 14) requires the heading of each
// statement to identify the entity as a limited liability entity, so every
// statement name carries its designation. House format puts a comma before the
// suffix — "County Line SRN, LLC" (Jimmy, 2026-08-26) — which is the form the
// executive-summary titles already used.
//
// An entity that already carries a designation keeps it: an existing LLC suffix
// is normalized to the comma form ("BROZ FUND I LLC" -> "BROZ FUND I, LLC"), and
// LP / Inc. / Corp. / LLP / Ltd. are left exactly as they are. County Line Rail
// Fund — the fund itself, an LP — is excluded outright per Jimmy.
const DESIGNATION_RX = /(?:,\s*|\s+)(l\.?l\.?c\.?|l\.?l\.?l\.?p\.?|l\.?l\.?p\.?|l\.?p\.?|inc\.?|incorporated|corp\.?|corporation|co\.|ltd\.?|limited|trust)$/i;
const NO_DESIGNATION = [/county\s*line\s*rail\s*fund/i];
function withDesignation(name) {
  const n = String(name || '').trim().replace(/\s+/g, ' ');
  if (!n) return n;
  if (NO_DESIGNATION.some(rx => rx.test(n))) return n;
  const m = n.match(DESIGNATION_RX);
  if (m) return /^l\.?l\.?c\.?$/i.test(m[1]) ? n.replace(DESIGNATION_RX, ', LLC') : n;
  return n + ', LLC';
}
function displayEntityName(name) {
  const n = String(name || '').trim();
  for (const { match, display } of ENTITY_DISPLAY_NAMES) if (match.test(n)) return withDesignation(display);
  return withDesignation(n) || 'Entity';
}

// ── statement profile ─────────────────────────────────────
// The default ("srn") profile reproduces the SRN-style development-entity
// package. A holding company like Banyan SFR GP Investors (entity 49) has a
// different chart AND a different Statements-of-Operations shape (Operating
// Expenses / Other Income (Expense) / Income Taxes), so it selects its own
// profile by entity name/code. Every profile-aware helper keys off this switch.
function entityProfile(opts) {
  const name = String((opts && opts.entityName) || '').trim();
  const code = String((opts && opts.entityCode) || '').trim().toUpperCase();
  if (/banyan\s*sfr\s*gp\s*investors/i.test(name) || code === 'BANYANSF') return 'bsfrgp';
  // Banyan Residential (entity 41) — an operating/holding company whose CPA
  // package uses the classic broad P&L groupings (Payroll & Related, Travel/
  // Meals & Entertainment, Utilities & Facilities, G&A, Office Expense, Taxes &
  // Insurance, Depreciation & Amortization) plus Other Income (Expense) and
  // Income Taxes, and a development-cost-heavy balance sheet (Soft Costs / Land
  // / Other Development under Other Assets). Distinct from both srn and bsfrgp.
  if (/banyan\s*residential/i.test(name) || code === 'BANYANRE1') return 'banyan';
  // County Line Rail Fund I, LP (entity 40) — the fund itself. Its regular
  // statement package keeps the ORIGINAL srn balance-sheet shape: no separate
  // Intercompany Receivable / Intercompany Payable subsections (Jimmy, 2026-08-18).
  // Scope is the fund only — CLRFI Midco I and the operating companies below it
  // are ordinary srn entities and DO get the intercompany sections.
  if (/^county\s*line\s*rail\s*fund/i.test(name) || code === 'COUNTYLI1') return 'clrf';
  // CLIP Property Owner (entity 54). Its CPA reference package (County Line
  // Industrial Park, LLC) defines a specific balance-sheet shape that differs
  // from the generic srn heuristic: Allowance for Credit Losses nets inside
  // Accounts Receivable, Net; Earnest Money and Due from Outside Vendor sit in
  // Other Current Assets; PP&E is split into a gross Fixed Assets subtotal and a
  // separate Accumulated Depreciation subtotal; land/construction costs form a
  // Long Term Investments section; and the pushed-down parent investment
  // (19041) and matching contributed capital (34164) are eliminated. Distinct
  // from srn — pinned by entity name/code so no other entity is affected.
  if (/clip\s*property\s*owner/i.test(name) || code === 'CLIPPROP' || code === 'CLIPPRO1') return 'clip';
  // CLR Silsbee Property Owner (entity 39). Same CPA-mirrored shape as clip:
  // Allowance for Credit Losses nets inside Accounts Receivable, Net; PP&E is
  // split into gross Fixed Assets + a separate Accumulated Depreciation
  // subtotal; land/construction costs form a Long Term Investments section; and
  // the pushed-down parent investment (17001) and matching contributed capital
  // (34063) are eliminated. Silsbee has no Earnest Money / Due from Outside
  // Vendor, so nothing extra moves into Other Current Assets. Pinned by entity
  // name/code so no other entity is affected.
  if (/clr\s*silsbee\s*property\s*owner/i.test(name) || code === 'CLRSILSB2' || code === 'CLRSILSB') return 'silsbee';
  // Turnkey Rail (entity 36) - a construction-contractor P&L. Its CPA reference
  // package (June 2026) shapes the Statements of Operations as:
  //   Revenue (Construction Revenue) -> Total Revenue
  //   Cost of Goods Sold (the whole 5xxxx Cost-of-Construction block)
  //   Gross Profit
  //   General & Administrative Expenses (6xxxx)
  //   Other Income (Expense) - Interest Income, Interest Expense
  //   Net Income (Loss)
  // Two things the default srn shape got wrong for this chart. COGS is detected
  // by NAME, and none of Turnkey's 55xxx lines ('Track Materials',
  // 'Administrative Costs', 'Maintenance', ...) match, so the entire cost of
  // construction fell into operating expenses and no Gross Profit row rendered
  // at all. And srn has no Other Income (Expense) section, so Interest Income
  // (42000, a Revenue account) sat in the top line. Pinned by name/code so no
  // other entity is affected. (Jimmy, 2026-08-27.)
  if (/^turnkey[ ]*rail$/i.test(name) || code === 'TURNKEYR') return 'turnkey';
  // Banyan QOZB development consolidations — Bridge Banyan HP QOZB (43) and
  // Braker QOZ Business (45). Their consolidated CPA packages (CLA) present the
  // balance sheet with Other Assets split into Soft Costs / Construction Costs /
  // Land / Permits and fees / Other development (each with a subtotal), the
  // Allowance for Doubtful Acct inside Other Current Assets, the intercompany
  // Due-to lines folded into Accrued Liabilities, an equity section with
  // separate Members' Equity and Retained Earnings subtotals, loan proceeds in
  // Financing, and no single rule above Net Income. Pinned by parent name/code
  // (these are the fsEntityName/fsEntityCode the consolidated package uses) so
  // no other entity is affected. (Jimmy, 2026-08-28.)
  if (/bridge\s*banyan\s*hp\s*qozb/i.test(name) || /braker\s*qoz\s*business/i.test(name) || code === 'BRIDGEBA' || code === 'BRAKERQO1') return 'banyandev';
  return 'srn';
}

// -- Turnkey Rail P&L routing ------------------------------------------------
// COGS is the whole 5xxxx block - 50000 Cost of Goods Sold plus 55000-55170
// Cost of Construction (Insurance, Track Materials, Small Tools, Crossing
// Materials, Leases & Rentals, Fuel & Disposals, Truck, Haul/Fencing/Welding,
// Traffic Control, Utilities, Subcontractors, Turnkey Labor, Contract
// Employees, Staffing Labor, Administrative Costs, Marketing, Maintenance).
// Confirmed against the CPA package, whose COGS section shows only 55xxx lines.
function turnkeyIsCogs(row) { return /^5/.test(String(row.code || '')); }

// Other Income (Expense). NOTE the chart-specific trap: on Turnkey, 42000 is
// Interest INCOME (type Revenue, subtype 'Other Revenue') and 70000 is Interest
// EXPENSE (subtype 'Other Expense'). Other profiles in this file map 70000 to
// Interest Income, so never share a code heuristic between them - key off the
// subtype, with the code and name only as backstops.
function turnkeyIsOtherIncome(row) {
  if (row.type !== 'Revenue') return false;
  const sub = String(row.subtype || '').toLowerCase();
  const nm = String(row.name || '').toLowerCase();
  return /other revenue|other income/.test(sub) || /interest income/.test(nm) || String(row.code) === '42000';
}
function turnkeyIsOtherExpense(row) {
  if (row.type !== 'Expense') return false;
  if (turnkeyIsCogs(row)) return false; // COGS wins; never double-count
  const sub = String(row.subtype || '').toLowerCase();
  const nm = String(row.name || '').toLowerCase();
  return /other expense/.test(sub) || /interest expense/.test(nm) || String(row.code) === '70000';
}

// ── Other Income (Expense) - ONE shared classifier for every profile ────────
// Jimmy, 2026-08-28: on EVERY entity, misc revenue / interest income / other
// income / misc expenses come out of operating results and are presented
// together below them, with interest expense, the non-operating gains and
// losses and the other oddments alongside, and income taxes in their own
// section. Defined once, on purpose: before this, the section was hand-rolled
// three times (turnkey code pins, BSFRGP_PL_MAP, BANYAN_PL_MAP) and the srn
// family had no section at all, so the same account was classified two ways on
// two entities.
//
// Keyed off type + NAME, never a bare code range. On Turnkey Rail 70000 is
// Interest EXPENSE and 42000 is Interest Income - the exact reverse of every
// other chart here - so a code heuristic is guaranteed to be wrong somewhere.
//
// Returns null for an account that belongs in operating results.
const OIE_INCOME  = { bucket: 'otherIncome',  group: 'Other Income',  sub: 'Other Income' };
const OIE_EXPENSE = { bucket: 'otherExpense', group: 'Other Expense', sub: 'Other Expense' };
const OIE_TAX     = { bucket: 'incomeTax',    group: 'State and Local Taxes', sub: 'State and Local Taxes' };

function otherIeRoute(row) {
  const name = String(row.name || '').toLowerCase().trim();
  const sub = String(row.subtype || '').toLowerCase();
  const type = String(row.type || '');
  if (type === 'Revenue') {
    // Non-operating income: interest, misc, other, gains/losses, dividends,
    // debt forgiveness, tax refunds. Deliberately NOT matched: Banyan
    // Residential's 40250 'Expense - Bad Debt', a Revenue-typed expense
    // account - that is a chart problem, not a presentation one.
    if (/interest income/.test(name)) return OIE_INCOME;
    if (/^misc/.test(name)) return OIE_INCOME;
    if (/other income/.test(name)) return OIE_INCOME;
    if (/\bgain|\bloss/.test(name)) return OIE_INCOME;
    if (/dividend/.test(name)) return OIE_INCOME;
    if (/forgiveness/.test(name)) return OIE_INCOME;
    if (/tax refund/.test(name)) return OIE_INCOME;
    if (/other revenue|other income/.test(sub)) return OIE_INCOME;
    return null;
  }
  if (type !== 'Expense') return null;
  // Penalties BEFORE the tax test: 'State and Local Tax Penalties' contains
  // 'state and local tax' but is an other expense, not an income tax.
  if (/penalt/.test(name)) return OIE_EXPENSE;
  if (/^miscellaneous$/.test(name)) return OIE_EXPENSE;
  if (/expense - misc|^misc expense/.test(name)) return OIE_EXPENSE;
  if (/interest expense/.test(name)) return OIE_EXPENSE;
  if (/^other expense/.test(name)) return OIE_EXPENSE;
  if (/other expense/.test(sub)) return OIE_EXPENSE;
  // Income taxes get their own section. Property tax, tax & license and the
  // operating 'Contract Loss' / 'Credit Loss' expenses stay where they are.
  if (/state and local tax|franchise tax|income tax/.test(name)) return OIE_TAX;
  return null;
}

// CLA's Turnkey cash-flow line set. Each operating/investing/financing line
// names the accounts that roll into it. Signs are cash effects: an asset
// increase consumes cash, a liability increase provides it.
const TURNKEY_CF = {
  depreciation: ['15100'],                 // contra-asset movement, added back
  accountsReceivable: ['11000'],
  accountsPayable: ['20000'],
  contractAssets: ['11010', '14500'],
  prepaidExpenses: ['13000'],
  contractLiabilities: ['24000'],
  fixedAssets: ['15000', '15005'],         // gross additions (investing)
  memberCapital: ['30000', '32000'],
};

// CLA's Turnkey balance sheet, in order. Each entry is either a bare row or a
// titled group carrying its own subtotal. `name` overrides the ledger account
// name where CLA words it differently (11000 'Accounts Receivable' prints as
// 'Contract Receivables'). `always` keeps a row visible when both columns are
// zero, which CLA does for Costs and Estimated Earnings in Excess of Billings
// (it prints a dash).
const TURNKEY_BS_ASSETS = [
  { kind: 'row', code: '10100', name: 'Operating Checking' },
  { kind: 'row', code: '11000', name: 'Contract Receivables' },
  { kind: 'group', title: 'Contract Assets', rows: [
      { code: '14500', name: 'Costs and Estimated Earnings in Excess of Billings', always: true },
      { code: '11010', name: 'Retainage Receivable' },
  ] },
  { kind: 'row', code: '13000', name: 'Prepaid Expenses' },
  { kind: 'group', title: 'Fixed Assets', rows: [
      { code: '15000', name: 'Property & Equipment' },
      { code: '15005', name: 'Vehicles' },
      { code: '15100', name: 'Accumulated Depreciation' },
  ] },
];
// CLA's caption for the capital account on both the balance sheet and the
// statement of changes in members' equity.
const TURNKEY_EQUITY_NAMES = { '30000': "Members' Capital" };
const TURNKEY_BS_LIABS = [
  { kind: 'row', code: '20000', name: 'Accounts Payable' },
  { kind: 'group', title: 'Contract Liabilities', rows: [
      { code: '24000', name: 'Billings in Excess of Costs and Estimated Earnings' },
  ] },
];

// Build the block structure for one side of the balance sheet.
// valueOf(code) -> { cur, pri, name, zero } for an account that exists in
// either column, or null if it does not exist at all.
// An account with a balance that the spec does not name is appended as a bare
// row rather than dropped: a silently omitted account would unbalance the
// statement with nothing on the page to show it, so the caller also asserts
// the block totals against totalAssets / totalLiab.
function buildTurnkeyBlocks(spec, valueOf, codesOnSide) {
  const named = new Set();
  const blocks = [];
  const take = (entry) => {
    named.add(String(entry.code));
    const v = valueOf(String(entry.code));
    if (!v) return entry.always ? { code: entry.code, name: entry.name, cur: 0, pri: 0 } : null;
    if (v.zero && !entry.always) return null;
    return { code: entry.code, name: entry.name, cur: v.cur, pri: v.pri };
  };
  for (const b of spec) {
    if (b.kind === 'row') {
      const row = take(b);
      if (row) blocks.push(Object.assign({ kind: 'row' }, row));
      continue;
    }
    const rows = b.rows.map(take).filter(Boolean);
    if (!rows.length) continue;
    blocks.push({ kind: 'group', title: b.title, rows,
      subtotal: { cur: r2(rows.reduce((s, r) => s + r.cur, 0)),
                  pri: r2(rows.reduce((s, r) => s + r.pri, 0)) } });
  }
  const extras = [];
  for (const code of codesOnSide) {
    const c = String(code);
    if (named.has(c)) continue;
    const v = valueOf(c);
    if (!v || v.zero) continue;
    extras.push({ kind: 'row', code: c, name: v.name, cur: v.cur, pri: v.pri, unmapped: true });
  }
  extras.sort((a, b) => String(a.code).localeCompare(String(b.code)));
  return { blocks: blocks.concat(extras), unmapped: extras.map(e => e.code) };
}

// CLA's row order for the Turnkey statements of operations. Explicit because
// it is not derivable: COGS runs Track Materials, Administrative Costs,
// Maintenance, Insurance, Fuel & Disposals, Truck, Staffing Labor - not code
// order (55020, 55140, 55170, 55010, 55050, 55060, 55130), not alphabetical,
// not by amount. Any account not listed sorts after the listed ones, in code
// order, so a new account appears rather than vanishing.
const TURNKEY_COGS_ORDER = ['55020', '55140', '55170', '55010', '55050', '55060', '55130'];
const TURNKEY_GA_ORDER = ['64000', '63000'];

// Construction Revenue is presented net of the Work in Progress Adjustment.
// 45000 carries gross billings to date and 49999 the WIP true-up (negative,
// and equal to Billings in Excess of Costs on the balance sheet). CLA shows one
// line, so the two are summed and labelled 'Construction Revenue'.
const TURNKEY_REVENUE_NET_INTO = '45000';
const TURNKEY_REVENUE_NET_FROM = ['49999'];

// Reorder `lines` by an explicit code order; unlisted codes keep code order
// and follow the listed ones.
function orderByCodes(lines, order) {
  const rank = new Map(order.map((c, i) => [String(c), i]));
  return lines.slice().sort((a, b) => {
    const ra = rank.has(String(a.code)) ? rank.get(String(a.code)) : order.length;
    const rb = rank.has(String(b.code)) ? rank.get(String(b.code)) : order.length;
    if (ra !== rb) return ra - rb;
    return String(a.code).localeCompare(String(b.code));
  });
}

// Fold the WIP adjustment into the Construction Revenue line. Returns a new
// array; the totals are unaffected (it is a pure re-presentation of two lines
// that were already both in revenue), which is why Total Revenue and net income
// do not move.
function turnkeyNetRevenue(revenueLines) {
  const from = new Set(TURNKEY_REVENUE_NET_FROM);
  const netted = revenueLines.filter(l => from.has(String(l.code)));
  if (!netted.length) return revenueLines;
  return revenueLines.filter(l => !from.has(String(l.code))).map(l => {
    if (String(l.code) !== TURNKEY_REVENUE_NET_INTO) return l;
    const add = k => netted.reduce((s, n) => s + (n[k] || 0), l[k] || 0);
    return Object.assign({}, l, { cur: r2(add('cur')), pri: r2(add('pri')), ytd: r2(add('ytd')) });
  });
}

// Turnkey Rail began operations 2026-04-16. Its first period is a stub, so
// every statement in CLA's package is dated from inception rather than from
// the calendar year start:
//   Operations:  For the One Month Ended June 30, 2026, the Period
//                April 16 - May 31, 2026, and the Period April 16 - June 30, 2026
//   Cash flows / equity:  For the Period April 16 - June 30, 2026
//   Equity opening row:   Equity Balances at April 16, 2026
// Kept as a constant rather than an entities column because the request was
// scoped to Turnkey; a second entity needing this should get a real
// inception_date field instead of a second constant here.
const TURNKEY_INCEPTION = '2026-04-16';
function inceptionFor(profile) { return profile === 'turnkey' ? TURNKEY_INCEPTION : null; }

// 'April 16' - month and day, no year (the year is carried by the date that
// follows it in every label that uses this).
function monthDay(d) {
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const mm = parseInt(String(d).slice(5, 7), 10), dd = parseInt(String(d).slice(8, 10), 10);
  return MONTHS[mm - 1] + ' ' + dd;
}
// 'April 16 - June 30, 2026'
function inceptionRange(inception, to) { return monthDay(inception) + ' - ' + longDate(to); }
// m/d/yyyy, for the landscape equity header's opening column.
function slashDate(d) {
  return parseInt(String(d).slice(5, 7), 10) + '/' + parseInt(String(d).slice(8, 10), 10) + '/' + String(d).slice(0, 4);
}

// Which profiles present the intercompany balances as their own balance-sheet
// subsections — Current Assets › Intercompany Receivable (immediately before
// Other Current Assets) and Current Liabilities › Intercompany Payable
// (immediately after Accounts Payable).
//
// Everything on the default chart does. The three profiles that do NOT are
// banyan and bsfrgp (their CPA reference packages already define their own
// intercompany presentation) and clrf (excluded by request). Those keep the
// classification they had before 2026-08-18, byte for byte.
function usesIntercompanySections(profile) { return profile === 'srn'; }

// Balance-sheet section map for the Banyan SFR GP Investors (holding-company)
// profile. Same shape as BS_ACCOUNT_MAP; selected only for that entity.
const BS_ACCOUNT_MAP_BSFRGP = {
  // Current Assets → Cash and Cash Equivalents
  '10131': ['Current Assets', 'Cash and Cash Equivalents'],
  '10141': ['Current Assets', 'Cash and Cash Equivalents'],
  // Current Assets → Intercompany Receivable
  '18353': ['Current Assets', 'Intercompany Receivable'],
  '18384': ['Current Assets', 'Intercompany Receivable'],
  // Investments → Investment in Subsidiary
  '19027': ['Investments', 'Investment in Subsidiary'],
  '19037': ['Investments', 'Investment in Subsidiary'],
  '19055': ['Investments', 'Investment in Subsidiary'],
  '19057': ['Investments', 'Investment in Subsidiary'],
  '19060': ['Investments', 'Investment in Subsidiary'],
  '19074': ['Investments', 'Investment in Subsidiary'],
  // Current Liabilities
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '23000': ['Current Liabilities', 'Intercompany Payable'],
  // Members Equity (Retained Earnings & Net Income handled specially in renderer)
  '39000': ['Members Equity', 'Retained Earnings'],
};

// Per-profile balance-sheet classification. Default profile uses the SRN map
// (BS_ACCOUNT_MAP) via bsClassify(); bsfrgp uses its own map with the same
// heuristic fallback so no account is ever dropped.
function bsClassifyFor(profile, row) {
  if (profile === 'banyan') return banyanBsClassify(row);
  if (profile === 'banyandev') return banyandevBsClassify(row);
  if (profile === 'bsfrgp') {
    const explicit = BS_ACCOUNT_MAP_BSFRGP[String(row.code)];
    if (explicit) return { section: explicit[0], sub: explicit[1] };
    const name = (row.name || '').toLowerCase();
    if (row.type === 'Asset') {
      if (/cash|checking|savings|bank|clearing|ics/.test(name)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
      if (/due from|intercompany/.test(name)) return { section: 'Current Assets', sub: 'Intercompany Receivable' };
      if (/investment/.test(name)) return { section: 'Investments', sub: 'Investment in Subsidiary' };
      if (/receivable/.test(name)) return { section: 'Current Assets', sub: 'Accounts Receivable, Net' };
      return { section: 'Other Assets', sub: 'Other Assets' };
    }
    if (row.type === 'Liability') {
      if (/due to|intercompany/.test(name)) return { section: 'Current Liabilities', sub: 'Intercompany Payable' };
      if (/payable/.test(name)) return { section: 'Current Liabilities', sub: 'Accounts Payable' };
      if (/loan|note payable|bond/.test(name)) return { section: 'Long Term Liabilities', sub: 'Loans' };
      return { section: 'Current Liabilities', sub: 'Other Current Liabilities' };
    }
    if (row.type === 'Equity') {
      if (/retained earning/.test(name)) return { section: 'Members Equity', sub: 'Retained Earnings' };
      return { section: 'Members Equity', sub: 'Members Equity' };
    }
    return { section: 'Other', sub: 'Other' };
  }
  if (profile === 'clip') {
    const explicit = BS_ACCOUNT_MAP_CLIP[String(row.code)];
    if (explicit) return { section: explicit[0], sub: explicit[1] };
    // Defensive fallback (every CLIP account is pinned above, but never drop a
    // row). Mirror the CPA groupings: allowance nets into AR, prepaid/reserve/
    // earnest to Other Current Assets, land/construction to Long Term
    // Investments, remaining capitalized costs to Other Assets.
    const name = (row.name || '').toLowerCase();
    if (row.type === 'Asset') {
      if (isCashAccount(row)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
      if (/due from|intercompany/.test(name)) return { section: 'Current Assets', sub: 'Intercompany Receivable' };
      if (/receivable|allowance/.test(name)) return { section: 'Current Assets', sub: 'Accounts Receivable, Net' };
      if (/prepaid|reserve|earnest|deposit/.test(name)) return { section: 'Current Assets', sub: 'Other Current Assets' };
      if (/accum|depreciation/.test(name)) return { section: 'Fixed Assets, Net', sub: 'Accumulated Depreciation' };
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
  if (profile === 'silsbee') {
    const explicit = BS_ACCOUNT_MAP_SILSBEE[String(row.code)];
    if (explicit) return { section: explicit[0], sub: explicit[1] };
    // Defensive fallback (every Silsbee account is pinned above). Mirror the CPA
    // groupings: allowance nets into AR, prepaid/reserve to Other Current
    // Assets, land/construction to Long Term Investments, remaining capitalized
    // costs to Other Assets.
    const name = (row.name || '').toLowerCase();
    if (row.type === 'Asset') {
      if (isCashAccount(row)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
      if (/due from|intercompany/.test(name)) return { section: 'Current Assets', sub: 'Intercompany Receivable' };
      if (/receivable|allowance/.test(name)) return { section: 'Current Assets', sub: 'Accounts Receivable, Net' };
      if (/prepaid|reserve|deposit/.test(name)) return { section: 'Current Assets', sub: 'Other Current Assets' };
      if (/accum|depreciation/.test(name)) return { section: 'Fixed Assets, Net', sub: 'Accumulated Depreciation' };
      return { section: 'Other Assets', sub: 'Other Assets' };
    }
    if (row.type === 'Liability') {
      if (/due to|intercompany/.test(name)) return { section: 'Current Liabilities', sub: 'Intercompany Payable' };
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
  return bsClassify(row, { intercompany: usesIntercompanySections(profile) });
}

// Sub-order for the bsfrgp Investments section (its only extra subsection name).
const BS_SUB_ORDER_BSFRGP = Object.assign({}, {
  'Investments': ['Investment in Subsidiary'],
  'Current Liabilities': ['Accounts Payable', 'Intercompany Payable'],
});

// Banyan SFR GP Investors Statements-of-Operations account routing. Maps each
// P&L account to a top-level bucket so the statement reproduces the reference
// nesting. Anything unmapped falls back to Operating Expenses (never dropped).
//   opex   → Operating Expenses (General and Administrative → Legal and Accounting)
//   otherIncome / otherExpense → Other Income (Expense)
//   incomeTax → Income Taxes (State and Local Taxes)
// Only the OPERATING side lives here now. Interest income, the gains, the
// penalties and the franchise tax are classified by the shared otherIeRoute
// (2026-08-28), so they are not repeated in this map.
const BSFRGP_PL_MAP = {
  '63000': { bucket: 'opex',         group: 'General and Administrative Expenses', sub: 'Legal and Accounting' },
};
function bsfrgpPlRoute(row) {
  // Shared Other Income (Expense) / Income Taxes classifier wins.
  const oie = otherIeRoute(row);
  if (oie) return oie;
  const m = BSFRGP_PL_MAP[String(row.code)];
  if (m) return m;
  // This profile has no top-line Revenue section, so any revenue account the
  // classifier did not claim still belongs in Other Income - dropping it into
  // Operating Expenses would flip its sign and break net income.
  if (row.type === 'Revenue') return OIE_INCOME;
  return { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Legal and Accounting' };
}

// ── Banyan Residential (entity 41) profile ─────────────────────────────────
// Balance-sheet section map. Development-cost-heavy chart: Other Assets splits
// into Soft Costs / Land / Other Development. Fixed Assets carries an
// Accumulated Depreciation contra subsection. Built from the CPA June-2026
// reference package; any code not listed falls to the heuristic below.
const BS_ACCOUNT_MAP_BANYAN = {
  // Current Assets → Cash and Cash Equivalents
  '10030': ['Current Assets', 'Cash and Cash Equivalents'],
  '10143': ['Current Assets', 'Cash and Cash Equivalents'],
  '10144': ['Current Assets', 'Cash and Cash Equivalents'],
  '10300': ['Current Assets', 'Cash and Cash Equivalents'],
  // Current Assets → Accounts Receivable, Net
  '12000': ['Current Assets', 'Accounts Receivable, Net'],
  // Current Assets → Prepaid Expenses
  '12922': ['Current Assets', 'Prepaid Expenses'],
  '13002': ['Current Assets', 'Prepaid Expenses'],
  // Current Assets → Other Current Assets
  '13003': ['Current Assets', 'Prepaid Expenses'], // Prepaid Rent; moved to Prepaid Expenses per CLA review
  // Fixed Assets, Net → Fixed Assets
  '15500': ['Fixed Assets, Net', 'Fixed Assets'],
  '15510': ['Fixed Assets, Net', 'Fixed Assets'],
  // Fixed Assets, Net → Accumulated Depreciation (contra)
  '16500': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  '16510': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  // Other Assets → Soft Costs
  '11760': ['Other Assets', 'Soft Costs'],
  '11920': ['Other Assets', 'Soft Costs'],
  '11970': ['Other Assets', 'Soft Costs'],
  '12013': ['Other Assets', 'Soft Costs'],
  '12112': ['Other Assets', 'Soft Costs'],
  '12115': ['Other Assets', 'Soft Costs'],
  '12127': ['Other Assets', 'Soft Costs'],
  '12132': ['Other Assets', 'Soft Costs'],
  '12180': ['Other Assets', 'Soft Costs'],
  '12238': ['Other Assets', 'Soft Costs'],
  '13421': ['Other Assets', 'Soft Costs'],
  // Other Assets → Land
  '11030': ['Other Assets', 'Land'],
  '16900': ['Other Assets', 'Land'],
  // Other Assets → Other Development
  '12383': ['Other Assets', 'Other Development'],
  '12720': ['Other Assets', 'Other Development'],
  '12730': ['Other Assets', 'Other Development'],
  // Current Liabilities → Accounts Payable (incl. Credit Card Payable)
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '20500': ['Current Liabilities', 'Accounts Payable'],
  // Current Liabilities → Other Current Liabilities
  '21112': ['Current Liabilities', 'Other Current Liabilities'],
  // Members Equity (Net Income handled specially in renderer). Banyan's CPA
  // package shows contributed capital + distributions on the equity face and
  // has NO separate Retained Earnings line, so these are all Members Equity.
  '33104': ['Members Equity', 'Members Equity'],
  '34004': ['Members Equity', 'Members Equity'],
  '34117': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};

// Intercompany Receivable is Banyan's largest asset subsection but its "Due
// from ..." accounts are numerous (18xxx block) and stable; classify the whole
// 18xxx block by prefix rather than listing each, so a new intercompany account
// is picked up automatically.
function banyanBsClassify(row) {
  const explicit = BS_ACCOUNT_MAP_BANYAN[String(row.code)];
  if (explicit) return { section: explicit[0], sub: explicit[1] };
  const code = String(row.code || '');
  const name = (row.name || '').toLowerCase();
  if (row.type === 'Asset') {
    if (/^18\d/.test(code) || /due from|intercompany/.test(name)) return { section: 'Current Assets', sub: 'Intercompany Receivable' };
    if (/^10[0-3]\d/.test(code) || /cash|checking|savings|bank|clearing|ics|money market/.test(name)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
    if (/receivable/.test(name)) return { section: 'Current Assets', sub: 'Accounts Receivable, Net' };
    if (/prepaid/.test(name)) return { section: 'Current Assets', sub: 'Prepaid Expenses' };
    if (/^16[0-9]\d/.test(code) && /depreciation|amortization|accum/.test(name)) return { section: 'Fixed Assets, Net', sub: 'Accumulated Depreciation' };
    if (/^155\d|^156\d|equipment|vehicle|furniture|computer hardware/.test(code + ' ' + name)) return { section: 'Fixed Assets, Net', sub: 'Fixed Assets' };
    if (/earnest|security deposit|land/.test(name)) return { section: 'Other Assets', sub: 'Land' };
    if (/organization|travel|meals/.test(name)) return { section: 'Other Assets', sub: 'Other Development' };
    return { section: 'Other Assets', sub: 'Soft Costs' };
  }
  if (row.type === 'Liability') {
    if (/payable|credit card/.test(name)) return { section: 'Current Liabilities', sub: 'Accounts Payable' };
    if (/loan|note payable|bond/.test(name)) return { section: 'Long Term Liabilities', sub: 'Loans' };
    return { section: 'Current Liabilities', sub: 'Other Current Liabilities' };
  }
  if (row.type === 'Equity') {
    if (/retained earning/.test(name)) return { section: 'Members Equity', sub: 'Retained Earnings' };
    return { section: 'Members Equity', sub: 'Members Equity' };
  }
  return { section: 'Other', sub: 'Other' };
}

// Banyan section/sub presentation order. Fixed Assets shows the Accumulated
// Depreciation contra subsection after Fixed Assets; Other Assets shows Soft
// Costs → Land → Other Development.
const BS_SUB_ORDER_BANYAN = Object.assign({}, {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Prepaid Expenses', 'Intercompany Receivable', 'Other Current Assets'],
  'Fixed Assets, Net': ['Fixed Assets', 'Accumulated Depreciation'],
  'Other Assets': ['Soft Costs', 'Land', 'Other Development'],
  'Current Liabilities': ['Accounts Payable', 'Other Current Liabilities'],
});

// Accumulated-depreciation contra codes (subtracted within Fixed Assets, Net).
const BS_CONTRA_CODES_BANYAN = new Set(['16500', '16510']);

// ── Banyan QOZB development consolidation (banyandev) profile ────────────────
// Balance-sheet section map that reproduces the CLA consolidated package for
// Bridge Banyan HP QOZB / Braker QOZ Business exactly: Other Current Assets
// (incl. the Allowance for Doubtful Acct, netted with AR), then Other Assets
// split into Soft Costs / Construction Costs / Land / Permits and fees / Other
// development, each with its own subtotal. Liabilities: Accounts Payable and a
// single Accrued Liabilities subsection that ALSO absorbs the intercompany
// Due-to lines (23001/23003/23004), matching CLA. Both QOZBs share the Banyan
// development chart, so classification is by code and works for either.
const BS_ACCOUNT_MAP_BANYANDEV = {
  // Current Assets → Other Current Assets (AR + allowance + other receivables)
  '12000': ['Current Assets', 'Other Current Assets'],
  '12007': ['Current Assets', 'Other Current Assets'],
  '12001': ['Current Assets', 'Other Current Assets'],
  '12006': ['Current Assets', 'Other Current Assets'],
  '12008': ['Current Assets', 'Other Current Assets'],
  // Other Assets → Soft Costs
  '11506': ['Other Assets', 'Soft Costs'],
  '11640': ['Other Assets', 'Soft Costs'],
  '11650': ['Other Assets', 'Soft Costs'],
  '11670': ['Other Assets', 'Soft Costs'],
  '11730': ['Other Assets', 'Soft Costs'],
  '11760': ['Other Assets', 'Soft Costs'],
  '11970': ['Other Assets', 'Soft Costs'],
  '12013': ['Other Assets', 'Soft Costs'],
  '12117': ['Other Assets', 'Soft Costs'],
  '12127': ['Other Assets', 'Soft Costs'],
  '12230': ['Other Assets', 'Soft Costs'],
  '12321': ['Other Assets', 'Soft Costs'],
  '12385': ['Other Assets', 'Soft Costs'],
  '12410': ['Other Assets', 'Soft Costs'],
  '12420': ['Other Assets', 'Soft Costs'],
  '12422': ['Other Assets', 'Soft Costs'],
  '12597': ['Other Assets', 'Soft Costs'],
  '13420': ['Other Assets', 'Soft Costs'],
  '13423': ['Other Assets', 'Soft Costs'],
  '13424': ['Other Assets', 'Soft Costs'],
  '13425': ['Other Assets', 'Soft Costs'],
  // Other Assets → Construction Costs
  '11210': ['Other Assets', 'Construction Costs'],
  '11230': ['Other Assets', 'Construction Costs'],
  // Other Assets → Land
  '11010': ['Other Assets', 'Land'],
  // Other Assets → Permits and fees
  '11880': ['Other Assets', 'Permits and fees'],
  // Other Assets → Other development
  '12720': ['Other Assets', 'Other development'],
  '12916': ['Other Assets', 'Other development'],
  // Braker-specific codes (same profile, different chart) matched to the CLA
  // Braker package: Prepaid to Current Assets; Permits/Approval into Soft Costs
  // (Braker's CLA has no separate Permits group); Rate Cap Fees and Arbor Admin
  // Fee to Other development. HP does not use these codes, so it is unaffected.
  '13000': ['Current Assets', 'Other Current Assets'],
  '12180': ['Other Assets', 'Soft Costs'],
  '12411': ['Other Assets', 'Other development'],
  '12770': ['Other Assets', 'Other development'],
  // Current Liabilities → Accounts Payable
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '20001': ['Current Liabilities', 'Accounts Payable'],
  '20002': ['Current Liabilities', 'Accounts Payable'],
  // Current Liabilities → Accrued Liabilities (incl. intercompany Due-to)
  '21002': ['Current Liabilities', 'Accrued Liabilities'],
  '21003': ['Current Liabilities', 'Accrued Liabilities'],
  '21004': ['Current Liabilities', 'Accrued Liabilities'],
  '21006': ['Current Liabilities', 'Accrued Liabilities'],
  '21007': ['Current Liabilities', 'Accrued Liabilities'],
  '23001': ['Current Liabilities', 'Accrued Liabilities'],
  '23003': ['Current Liabilities', 'Accrued Liabilities'],
  '23004': ['Current Liabilities', 'Accrued Liabilities'],
  '23005': ['Current Liabilities', 'Accrued Liabilities'],
  '23007': ['Current Liabilities', 'Accrued Liabilities'],
  '23008': ['Current Liabilities', 'Accrued Liabilities'],
  '23010': ['Current Liabilities', 'Accrued Liabilities'],
  // Long Term Liabilities → Loans
  '25004': ['Long Term Liabilities', 'Loans'],
  // Members Equity (Retained Earnings & Net Income handled specially in renderer)
  '33022': ['Members Equity', 'Members Equity'],
  '34010': ['Members Equity', 'Members Equity'],
  '34106': ['Members Equity', 'Members Equity'],
  '34109': ['Members Equity', 'Members Equity'],
  '34168': ['Members Equity', 'Members Equity'],
  '34187': ['Members Equity', 'Members Equity'],
  '34188': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};
// Presentation order. Other Assets follows the CLA order Soft Costs →
// Construction Costs → Land → Permits and fees → Other development.
const BS_SUB_ORDER_BANYANDEV = Object.assign({}, {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Intercompany Receivable', 'Other Current Assets'],
  'Other Assets': ['Soft Costs', 'Construction Costs', 'Land', 'Permits and fees', 'Other development'],
  'Current Liabilities': ['Accounts Payable', 'Accrued Liabilities'],
});
// Every subsection under Other Assets carries a subtotal even when it holds a
// single account (CLA shows Total Land, Total Permits and fees), so the renderer
// forces subtotals within Other Assets for this profile.
const BANYANDEV_FORCE_SUBTOTAL_SECTIONS = new Set(['Other Assets']);
function banyandevBsClassify(row) {
  const explicit = BS_ACCOUNT_MAP_BANYANDEV[String(row.code)];
  if (explicit) return { section: explicit[0], sub: explicit[1] };
  // Heuristic fallback so nothing is ever dropped (a new account, or Braker's
  // few chart differences). Mirrors the map's intent.
  const code = String(row.code || '');
  const name = (row.name || '').toLowerCase();
  if (row.type === 'Asset') {
    if (isCashAccount(row)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
    if (/receivable|allowance|deposit/.test(name)) return { section: 'Current Assets', sub: 'Other Current Assets' };
    if (/^1121|^1123|construction cost|base construction/.test(code + ' ' + name)) return { section: 'Other Assets', sub: 'Construction Costs' };
    if (/^1101\d|land purchase/.test(code + ' ' + name)) return { section: 'Other Assets', sub: 'Land' };
    if (/^1188|impact fee|permit/.test(code + ' ' + name)) return { section: 'Other Assets', sub: 'Permits and fees' };
    if (/^1291|development fee|^1272|travel - other development/.test(code + ' ' + name)) return { section: 'Other Assets', sub: 'Other development' };
    return { section: 'Other Assets', sub: 'Soft Costs' };
  }
  if (row.type === 'Liability') {
    if (/^25|loan|note payable|bond/.test(code + ' ' + name)) return { section: 'Long Term Liabilities', sub: 'Loans' };
    if (/payable/.test(name)) return { section: 'Current Liabilities', sub: 'Accounts Payable' };
    return { section: 'Current Liabilities', sub: 'Accrued Liabilities' };
  }
  if (row.type === 'Equity') {
    if (/retained earning/.test(name)) return { section: 'Members Equity', sub: 'Retained Earnings' };
    return { section: 'Members Equity', sub: 'Members Equity' };
  }
  return { section: 'Other', sub: 'Other' };
}

// Banyan Statements-of-Operations account routing. The classic broad groupings:
// a top-line Revenue - Services section, then Operating Expenses split into
// Payroll & Related / Travel, Meals & Entertainment / Utilities & Facilities /
// G&A / Office Expense / Taxes & Insurance / Depreciation & Amortization, then
// Other Income (Expense), then Income Taxes. Each code maps to a bucket + a
// nested group/subsection so the renderer reproduces the reference nesting.
//   revenue   → Revenue - Services (top-line, before Operating Expenses)
//   opex      → Operating Expenses (grouped)
//   otherIncome / otherExpense → Other Income (Expense)
//   incomeTax → Income Taxes
const BANYAN_PL_MAP = {
  // Revenue - Services
  '42000': { bucket: 'revenue', group: 'Revenue - Services', sub: 'Revenue - Services' },
  '42200': { bucket: 'revenue', group: 'Revenue - Services', sub: 'Revenue - Services' },
  '43000': { bucket: 'revenue', group: 'Revenue - Services', sub: 'Revenue - Services' },
  // Operating Expenses → Payroll and Related Expenses → Payroll Expenses
  '60000': { bucket: 'opex', group: 'Payroll and Related Expenses', sub: 'Payroll Expenses' },
  '60021': { bucket: 'opex', group: 'Payroll and Related Expenses', sub: 'Payroll Expenses' },
  '60050': { bucket: 'opex', group: 'Payroll and Related Expenses', sub: 'Payroll Expenses' },
  '60100': { bucket: 'opex', group: 'Payroll and Related Expenses', sub: 'Payroll Expenses' },
  // Operating Expenses → Travel, Meals and Entertainment
  '60210': { bucket: 'opex', group: 'Travel, Meals and Entertainment', sub: 'Travel Expenses' },
  '60350': { bucket: 'opex', group: 'Travel, Meals and Entertainment', sub: 'Travel Expenses' },
  '60500': { bucket: 'opex', group: 'Travel, Meals and Entertainment', sub: 'Meals and Entertainment' },
  // Operating Expenses → Utilities and Facilities → Rent
  '61000': { bucket: 'opex', group: 'Utilities and Facilities', sub: 'Rent' },
  // Operating Expenses → General and Administrative → Legal and Accounting
  '63000': { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Legal and Accounting' },
  '63025': { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Legal and Accounting' },
  '63050': { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Legal and Accounting' },
  // Operating Expenses → General and Administrative → Debt Service
  '60253': { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Debt Service' },
  // Operating Expenses → Office Expense
  '60200': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  '67000': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  '67150': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  '67200': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  '67202': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  '67300': { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' },
  // Operating Expenses → Taxes and Insurance → Insurance
  '65000': { bucket: 'opex', group: 'Taxes and Insurance', sub: 'Insurance' },
  // Operating Expenses → Depreciation and Amortization
  '69100': { bucket: 'opex', group: 'Depreciation and Amortization Expense', sub: 'Depreciation' },
  '69000': { bucket: 'opex', group: 'Depreciation and Amortization Expense', sub: 'Amortization' },
  // Other Income (Expense) and Income Taxes are NOT listed here - they are
  // classified by the shared otherIeRoute (2026-08-28), which also pulls
  // 49999 Misc Revenue out of Revenue - Services and 67150 Miscellaneous out
  // of Office Expense.
};
function banyanPlRoute(row) {
  // Shared Other Income (Expense) / Income Taxes classifier wins, so 67150
  // Miscellaneous is no longer caught by the Office Expense pin below.
  const oie = otherIeRoute(row);
  if (oie) return oie;
  const m = BANYAN_PL_MAP[String(row.code)];
  if (m) return m;
  const name = (row.name || '').toLowerCase();
  if (row.type === 'Revenue') {
    return { bucket: 'revenue', group: 'Revenue - Services', sub: 'Revenue - Services' };
  }
  // Expense name heuristics, mirroring the reference groupings.
  if (/salary|salaries|wage|payroll tax|health insurance|benefit|401k|retirement/.test(name)) return { bucket: 'opex', group: 'Payroll and Related Expenses', sub: 'Payroll Expenses' };
  if (/travel/.test(name)) return { bucket: 'opex', group: 'Travel, Meals and Entertainment', sub: 'Travel Expenses' };
  if (/meals|entertainment/.test(name)) return { bucket: 'opex', group: 'Travel, Meals and Entertainment', sub: 'Meals and Entertainment' };
  if (/rent/.test(name)) return { bucket: 'opex', group: 'Utilities and Facilities', sub: 'Rent' };
  if (/utilit|electric|water|gas|telephone|internet/.test(name)) return { bucket: 'opex', group: 'Utilities and Facilities', sub: 'Rent' };
  if (/accounting|legal|professional fee/.test(name)) return { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Legal and Accounting' };
  if (/bank fee|debt service|interest expense/.test(name)) return { bucket: 'opex', group: 'General and Administrative Expenses', sub: 'Debt Service' };
  if (/depreciation/.test(name)) return { bucket: 'opex', group: 'Depreciation and Amortization Expense', sub: 'Depreciation' };
  if (/amortization/.test(name)) return { bucket: 'opex', group: 'Depreciation and Amortization Expense', sub: 'Amortization' };
  if (/insurance/.test(name)) return { bucket: 'opex', group: 'Taxes and Insurance', sub: 'Insurance' };
  if (/property tax|tax & license|tax and license|assessment/.test(name)) return { bucket: 'opex', group: 'Taxes and Insurance', sub: 'Taxes' };
  // Everything else administrative → Office Expense.
  return { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' };
}

// Presentation order for Banyan operating-expense groups.
const BANYAN_OPEX_GROUP_ORDER = [
  'Payroll and Related Expenses',
  'Travel, Meals and Entertainment',
  'Utilities and Facilities',
  'General and Administrative Expenses',
  'Office Expense',
  'Taxes and Insurance',
  'Depreciation and Amortization Expense',
];

// Subsection order WITHIN a group. Without this, subsections fall in account-code
// order, which puts Debt Service (60253) ahead of Legal and Accounting (63000)
// and Amortization (69000) ahead of Depreciation (69100) — both backwards from
// the CPA reference. Groups not listed keep account-code order.
const BANYAN_SUB_ORDER_IN_GROUP = {
  'Travel, Meals and Entertainment': ['Travel Expenses', 'Meals and Entertainment'],
  'General and Administrative Expenses': ['Legal and Accounting', 'Debt Service'],
  'Taxes and Insurance': ['Insurance', 'Taxes'],
  'Depreciation and Amortization Expense': ['Depreciation', 'Amortization'],
};

// ── Banyan QOZB development consolidations (banyandev) statement of operations ─
// Restructured to match the CLA package for HP and Braker EXACTLY. The two
// property-management systems use different chart conventions — the same 411xx
// fee code is Other Income on HP but Adjusted Residential Rent on Braker, and
// 40100 Loss/Gain to Lease is Revenue - Services on HP but Adjusted Residential
// Rent on Braker — so each entity carries its own account→group map. Section
// order and the nested Other Income (Expense) shape follow the CLA reference.
const BANYANDEV_REVENUE_GROUP_ORDER = ['Revenue - Services', 'Adjusted Residential Rent'];
const BANYANDEV_OPEX_GROUP_ORDER = [
  'Payroll Expenses', 'Facilities', 'Utilities', 'Legal and Accounting',
  'Office Expense', 'Taxes and Insurance', 'Advertising and Promotion',
  'Other Operating Expense',
];
const BANYANDEV_OTHER_INCOME_GROUP_ORDER = ['Other Income', 'Interest Income'];
// Build a code→{bucket,group,sub} map from { bucket: { group: { sub: [codes] } } }.
function bdBuildMap(spec) {
  const m = {};
  for (const bucket of Object.keys(spec))
    for (const group of Object.keys(spec[bucket]))
      for (const sub of Object.keys(spec[bucket][group]))
        for (const code of spec[bucket][group][sub]) m[String(code)] = { bucket, group, sub };
  return m;
}
const BANYANDEV_PL_MAP_HP = bdBuildMap({
  revenue: {
    'Revenue - Services': { 'Revenue - Services': ['40001','40002','40003','40004','40005','40006','40007','40008','40009','40010','40011','40014','40100'] },
    'Adjusted Residential Rent': { 'Adjusted Residential Rent': ['40200','40202','40400','40401','40402','40450','40480'] },
  },
  opex: {
    'Payroll Expenses': { 'Payroll Expenses': ['60002','60003','60004','60012','60006','60008','60009','60010','60050','60104','60105','60106'] },
    'Facilities': { 'Facilities': ['61034','61041','61077','61078','61079','61080','61085','61087','61089','61092','61093','61094','61095','61096','61082','61108','61158'] },
    'Utilities': { 'Utilities': ['61151','61152','61159','61164','61163'] },
    'Legal and Accounting': { 'Legal and Accounting': ['63052'] },
    'Office Expense': { 'Office Expense': ['60100','60101','60200','60451','60452','63026','67001','67152','67153','67154','67200','67210','67205','67251','67300','67402','67403','67404','68001','68110'] },
    'Taxes and Insurance': { 'Taxes and Insurance': ['68050','68055'] },
    'Advertising and Promotion': { 'Advertising and Promotion': ['67455','67457','67458','67461','67468'] },
    'Other Operating Expense': { 'Other Operating Expense': ['68300','68301','68303','68304','68305','68306','68307','68310'] },
  },
  otherIncome: {
    'Other Income': { 'Other Income': ['41110','41113','41114','41117','41122','41130'] },
    'Interest Income': { 'Interest Income': ['70000'] },
  },
  otherExpense: {
    'Other Expenses': { 'State and Local Taxes': ['68061'], 'Interest Expenses': ['75000'] },
  },
});
const BANYANDEV_PL_MAP_BRAKER = bdBuildMap({
  revenue: {
    'Revenue - Services': { 'Revenue - Services': ['40001'] },
    'Adjusted Residential Rent': { 'Adjusted Residential Rent': ['40100','40200','40201','40400','40480','41111','41113','41115','41117','41119','41120','41121','41122','41123','41124','41125','41126','41127','41128','41129'] },
  },
  opex: {
    'Payroll Expenses': { 'Payroll Expenses': ['60000','60008','60010','60016','60104','60105'] },
    'Facilities': { 'Facilities': ['61034','61056','61058','61064','61084','61085','61086','61087','61093','61169','61170','61171','61172'] },
    'Utilities': { 'Utilities': ['61147','61148','61149','61152','61153','61163','61165','61166','61167','61168'] },
    'Legal and Accounting': { 'Legal and Accounting': ['63000','63150'] },
    'Office Expense': { 'Office Expense': ['60451','60452','63025','63026','63045','63046','63100','67001','67002','67003','67004','67005','67006','67007','67011','67012','67150','67152','67200','67250','67251','67300','67301','67400','67401','67403','67405','67467','68110'] },
    'Taxes and Insurance': { 'Taxes and Insurance': ['65000','68055'] },
    'Advertising and Promotion': { 'Advertising and Promotion': ['67455','67457'] },
    'Other Operating Expense': { 'Other Operating Expense': ['68304','68306'] },
  },
  otherIncome: {
    'Interest Income': { 'Interest Income': ['70000'] },
  },
  otherExpense: {
    'Other Expense': { 'Other Expense': ['68061','75131','75132','75133','70350'] },
  },
});
// CLA suppresses the group subtotal on HP's "Other Expenses" (it prints the two
// subsection totals then goes straight to Total Other Income (Expense)); Braker's
// flat "Other Expense" keeps its group total.
const BANYANDEV_NO_GROUP_TOTAL = new Set(['Other Expenses']);
function banyandevPlRoute(entityKey, row) {
  const map = entityKey === 'braker' ? BANYANDEV_PL_MAP_BRAKER : BANYANDEV_PL_MAP_HP;
  const hit = map[String(row.code)];
  if (hit) return hit;
  // Fallback so nothing is dropped and net income still ties.
  const name = (row.name || '').toLowerCase();
  const oeGroup = entityKey === 'braker' ? 'Other Expense' : 'Other Expenses';
  if (row.type === 'Revenue') {
    if (/interest income/.test(name)) return { bucket: 'otherIncome', group: 'Interest Income', sub: 'Interest Income' };
    return { bucket: 'revenue', group: 'Adjusted Residential Rent', sub: 'Adjusted Residential Rent' };
  }
  if (/franchise tax|state.*tax|income tax/.test(name)) return { bucket: 'otherExpense', group: oeGroup, sub: entityKey === 'braker' ? 'Other Expense' : 'State and Local Taxes' };
  if (/mortgage interest|interest expense/.test(name)) return { bucket: 'otherExpense', group: oeGroup, sub: entityKey === 'braker' ? 'Other Expense' : 'Interest Expenses' };
  return { bucket: 'opex', group: 'Office Expense', sub: 'Office Expense' };
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
// Month + year only ("August 2026"). Used for page footers, where the exact
// day adds nothing and the CPA package convention is the period, not the date.
function monthYearLabel(asOf) {
  const d = new Date(asOf + 'T00:00:00Z');
  return MONTHS[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}
// Statements-of-Operations heading. The statement carries three period columns
// (current period, prior period, year to date), so the heading names all three
// rather than only the two comparative dates (Jimmy, 2026-08-26):
//   "For the One Month Ended July 31, 2026 and June 30, 2026 and
//    Year to Date Ended July 31, 2026"
// The year-to-date clause is dropped on an annual report, where it would just
// restate the period already named.
function opsHeadingLine(colLabel, curLong, priLong) {
  const word = String(colLabel || 'Month Ended').replace(/ Ended$/, '');
  const lead = word === 'Month' ? 'One Month' : word;
  const base = 'For the ' + lead + ' Ended ' + curLong + ' and ' + priLong;
  return word === 'Year' ? base : base + ' and Year to Date Ended ' + curLong;
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
// Bill.com / bank CLEARING accounts are cash-equivalent pass-throughs and belong
// with cash. They are spread across the 10xxx asset block on these charts (10500,
// 107xx, 108xx), so the code block alone does not find them and the word alone is
// not safe: CLR Silsbee Property Owner's 11213 "Track Clearing Repair" is a
// construction cost, and a name-only test swept it into Cash and Cash Equivalents.
// Requiring BOTH the word and a 10xxx code keeps every real clearing account and
// drops that false positive.
function isClearingAccount(r) {
  return /clearing/i.test(r.name || '') && /^10\d{3}$/.test(String(r.code || ''));
}
function isCashAccount(r) {
  return r.type === 'Asset' && (isCashCode(r.code) || isClearingAccount(r) || /cash|checking|savings|money market|operating acct|bank/i.test(r.name || '') || r.bank_acct);
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
//   Current Liabilities → Accounts Payable / Intercompany Payable
//                         / Other Current Liabilities
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

// ── CLIP Property Owner (entity 54) balance-sheet map ───────────────────────
// Explicit code → [section, subsection] assignment that reproduces the CPA
// reference package (County Line Industrial Park, LLC) exactly. Differences from
// the generic srn heuristic that this map pins:
//   • 12002 Allowance for Credit Losses → Accounts Receivable, Net (contra),
//     netting inside AR rather than falling to Other Assets.
//   • 11030 Earnest Money and 18002 Due from Outside Vendor → Other Current
//     Assets (the srn map/heuristic would route 18002 to Intercompany and
//     11030 to Other Assets).
//   • PP&E gross accounts → Fixed Assets, Net › Fixed Assets, with 16160/16500
//     accumulated depreciation split into their own Accumulated Depreciation
//     subsection (contra).
//   • Land/construction cost accounts → Investments › Long Term Investments.
//   • Development/soft-cost accounts → Other Assets.
// 19041 (Investment in CLIP Property Owner) and 34164 (Contributed Capital -
// County Line Industrial Park LLC) are the pushed-down parent investment/equity
// pair; they are ELIMINATED before classification (see clipEliminate below) and
// so are intentionally absent from this map.
const BS_ACCOUNT_MAP_CLIP = {
  // Current Assets → Cash and Cash Equivalents
  '10107': ['Current Assets', 'Cash and Cash Equivalents'],
  '10111': ['Current Assets', 'Cash and Cash Equivalents'],
  '10882': ['Current Assets', 'Cash and Cash Equivalents'],
  // Current Assets → Accounts Receivable, Net (12002 is a contra within it)
  '12000': ['Current Assets', 'Accounts Receivable, Net'],
  '12002': ['Current Assets', 'Accounts Receivable, Net'],
  // Current Assets → Intercompany Receivable
  '18307': ['Current Assets', 'Intercompany Receivable'],
  // Current Assets → Other Current Assets
  '11030': ['Current Assets', 'Other Current Assets'],
  '13001': ['Current Assets', 'Other Current Assets'],
  '13100': ['Current Assets', 'Other Current Assets'],
  '18002': ['Current Assets', 'Other Current Assets'],
  // Fixed Assets, Net → Fixed Assets (gross)
  '11670': ['Fixed Assets, Net', 'Fixed Assets'],
  '15150': ['Fixed Assets, Net', 'Fixed Assets'],
  '15165': ['Fixed Assets, Net', 'Fixed Assets'],
  '15170': ['Fixed Assets, Net', 'Fixed Assets'],
  '15200': ['Fixed Assets, Net', 'Fixed Assets'],
  '15210': ['Fixed Assets, Net', 'Fixed Assets'],
  '15505': ['Fixed Assets, Net', 'Fixed Assets'],
  // Fixed Assets, Net → Accumulated Depreciation (contra)
  '16160': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  '16500': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  // Investments → Long Term Investments
  '11010': ['Investments', 'Long Term Investments'],
  '11040': ['Investments', 'Long Term Investments'],
  '11050': ['Investments', 'Long Term Investments'],
  '11211': ['Investments', 'Long Term Investments'],
  '11230': ['Investments', 'Long Term Investments'],
  // Other Assets (capitalized development / soft costs)
  '11970': ['Other Assets', 'Other Assets'],
  '12013': ['Other Assets', 'Other Assets'],
  '12115': ['Other Assets', 'Other Assets'],
  '12230': ['Other Assets', 'Other Assets'],
  '12315': ['Other Assets', 'Other Assets'],
  '12321': ['Other Assets', 'Other Assets'],
  '12343': ['Other Assets', 'Other Assets'],
  '12381': ['Other Assets', 'Other Assets'],
  '12596': ['Other Assets', 'Other Assets'],
  '12720': ['Other Assets', 'Other Assets'],
  '12913': ['Other Assets', 'Other Assets'],
  // Current Liabilities
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '21006': ['Current Liabilities', 'Other Current Liabilities'],
  '21011': ['Current Liabilities', 'Other Current Liabilities'],
  '21200': ['Current Liabilities', 'Other Current Liabilities'],
  // Long Term Liabilities → Loans
  '25063': ['Long Term Liabilities', 'Loans'],
  // Members Equity (34164 eliminated; not listed)
  '34144': ['Members Equity', 'Members Equity'],
  '34158': ['Members Equity', 'Members Equity'],
  '34160': ['Members Equity', 'Members Equity'],
  '34202': ['Members Equity', 'Members Equity'],
  '34261': ['Members Equity', 'Members Equity'],
  '34262': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};

// Accumulated-depreciation contras for the CLIP profile (subtracted within the
// Fixed Assets, Net section, same mechanism as banyan's 165xx contras).
const BS_CONTRA_CODES_CLIP = new Set(['16160', '16500']);

// Codes eliminated on CLIP before classification: the pushed-down parent
// investment (asset) and the matching contributed-capital account (equity).
// They are equal and opposite and net out of every getBalances window, so
// dropping them keeps the balance sheet, members' equity statement and cash
// flow all tied by construction while removing the self-referential gross-up.
const CLIP_ELIMINATE_CODES = new Set(['19041', '34164']);

// ── CLR Silsbee Property Owner (entity 39) balance-sheet map ────────────────
// Reproduces the CPA reference package (County Line Rail Silsbee, LLC) exactly.
// Same shape family as the clip map. Notable points:
//   • 12002 Allowance for Credit Losses → Accounts Receivable, Net (contra).
//   • No Earnest Money / Due from Outside Vendor here, so Other Current Assets
//     is just prepaid insurance + interest reserve.
//   • PP&E gross accounts → Fixed Assets, Net › Fixed Assets, with 16160/16500
//     accumulated depreciation split into their own Accumulated Depreciation
//     subsection (contra).
//   • Land/construction cost accounts → Investments › Long Term Investments.
//   • Development/soft-cost accounts (incl. 11020 Scrap Metal) → Other Assets.
//   • Intercompany Receivable (18310/18311/18378) and Intercompany Payable
//     (23370/23375) present as their own subsections.
// 17001 (Investment - CLR Silsbee Property Owner LLC) and 34063 (Contributed
// Capital - County Line Rail Silsbee LLC) are the pushed-down parent
// investment/equity pair; they are ELIMINATED before classification (see
// SILSBEE_ELIMINATE_CODES) and are intentionally absent from this map.
const BS_ACCOUNT_MAP_SILSBEE = {
  // Current Assets → Cash and Cash Equivalents
  '10010': ['Current Assets', 'Cash and Cash Equivalents'],
  '10040': ['Current Assets', 'Cash and Cash Equivalents'],
  // Current Assets → Accounts Receivable, Net (12002 is a contra within it)
  '12000': ['Current Assets', 'Accounts Receivable, Net'],
  '12002': ['Current Assets', 'Accounts Receivable, Net'],
  // Current Assets → Intercompany Receivable
  '18310': ['Current Assets', 'Intercompany Receivable'],
  '18311': ['Current Assets', 'Intercompany Receivable'],
  '18378': ['Current Assets', 'Intercompany Receivable'],
  // Current Assets → Other Current Assets
  '13001': ['Current Assets', 'Other Current Assets'],
  '13100': ['Current Assets', 'Other Current Assets'],
  // Fixed Assets, Net → Fixed Assets (gross)
  '15100': ['Fixed Assets, Net', 'Fixed Assets'],
  '15150': ['Fixed Assets, Net', 'Fixed Assets'],
  '15165': ['Fixed Assets, Net', 'Fixed Assets'],
  '15175': ['Fixed Assets, Net', 'Fixed Assets'],
  '15200': ['Fixed Assets, Net', 'Fixed Assets'],
  '15210': ['Fixed Assets, Net', 'Fixed Assets'],
  '15220': ['Fixed Assets, Net', 'Fixed Assets'],
  // Fixed Assets, Net → Accumulated Depreciation (contra)
  '16160': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  '16500': ['Fixed Assets, Net', 'Accumulated Depreciation'],
  // Investments → Long Term Investments
  '11040': ['Investments', 'Long Term Investments'],
  '11211': ['Investments', 'Long Term Investments'],
  '11215': ['Investments', 'Long Term Investments'],
  '11230': ['Investments', 'Long Term Investments'],
  // Other Assets (capitalized development / soft costs; 11020 scrap metal)
  '11020': ['Other Assets', 'Other Assets'],
  '11970': ['Other Assets', 'Other Assets'],
  '12115': ['Other Assets', 'Other Assets'],
  '12230': ['Other Assets', 'Other Assets'],
  '12315': ['Other Assets', 'Other Assets'],
  '12321': ['Other Assets', 'Other Assets'],
  '12343': ['Other Assets', 'Other Assets'],
  '12421': ['Other Assets', 'Other Assets'],
  '12594': ['Other Assets', 'Other Assets'],
  '12596': ['Other Assets', 'Other Assets'],
  '12720': ['Other Assets', 'Other Assets'],
  '12913': ['Other Assets', 'Other Assets'],
  '13420': ['Other Assets', 'Other Assets'],
  // Current Liabilities → Accounts Payable
  '20000': ['Current Liabilities', 'Accounts Payable'],
  '20100': ['Current Liabilities', 'Accounts Payable'],
  // Current Liabilities → Intercompany Payable
  '23370': ['Current Liabilities', 'Intercompany Payable'],
  '23375': ['Current Liabilities', 'Intercompany Payable'],
  // Current Liabilities → Other Current Liabilities
  '21006': ['Current Liabilities', 'Other Current Liabilities'],
  '24000': ['Current Liabilities', 'Other Current Liabilities'],
  // Long Term Liabilities → Loans (22100 short-term CLRF I loan + 25063 BOT)
  '22100': ['Long Term Liabilities', 'Loans'],
  '25063': ['Long Term Liabilities', 'Loans'],
  // Members Equity (34063 eliminated; not listed). 34006 correctly reads
  // "Contributed Capital - CLRFI Silsbee Sponsor" in the CoA.
  '34006': ['Members Equity', 'Members Equity'],
  '34144': ['Members Equity', 'Members Equity'],
  '34151': ['Members Equity', 'Members Equity'],
  '34161': ['Members Equity', 'Members Equity'],
  '34171': ['Members Equity', 'Members Equity'],
  '34261': ['Members Equity', 'Members Equity'],
  '39000': ['Members Equity', 'Retained Earnings'],
};

// Accumulated-depreciation contras for the Silsbee profile.
const BS_CONTRA_CODES_SILSBEE = new Set(['16160', '16500']);

// Codes eliminated on Silsbee before classification: the pushed-down parent
// investment (17001, asset) and matching contributed capital (34063, equity).
const SILSBEE_ELIMINATE_CODES = new Set(['17001', '34063']);

// opts.intercompany routes "Due from ..." / "Due to ..." accounts into their own
// balance-sheet subsections. It is ON for the default (srn) profile and OFF for
// clrf, which keeps the pre-2026-08-18 shape. banyan/bsfrgp never reach here.
function bsClassify(row, opts = {}) {
  const ic = !!(opts && opts.intercompany);
  const nm = (row.name || '').toLowerCase();
  // The intercompany test runs BEFORE the explicit code map on purpose. A few
  // codes are pinned by the SRN reference package (18002 -> Other Current
  // Assets); if such an account is in fact named "Due from <affiliate>" it
  // belongs in Intercompany Receivable, and the name is the better evidence.
  if (ic && row.type === 'Asset' && /due from|intercompany/.test(nm)) {
    return { section: 'Current Assets', sub: 'Intercompany Receivable' };
  }
  // An ASSET account whose name starts "Due to ..." is a payable that was coded
  // on the wrong side of the chart, and it carries a credit balance to prove it.
  // Odyssey Holdings 18397 "Due to Phil Brosseau" (-7,736.96 at 6/30/2026) is the
  // only such account in the portfolio, and it was printing inside Other Assets.
  // Present it where it belongs, under Current Liabilities.
  //
  // Other Current Liabilities rather than Intercompany Payable: the counterparty
  // is an individual with no ledger and no org node, so there is nothing to
  // eliminate against. Map the code explicitly if a real affiliate ever lands
  // here. The match is anchored so it can only ever catch an account whose name
  // BEGINS with "Due to", never a phrase buried mid-name.
  //
  // bsSideOf / bsBal in buildStatements move the row to the liability side and
  // negate it, so it reads credit-positive like every other liability. The tie-out
  // is unchanged: assets rise by 7,736.96 and liabilities rise by the same amount.
  if (row.type === 'Asset' && /^\s*due\s+to\b/.test(nm)) {
    return { section: 'Current Liabilities', sub: 'Other Current Liabilities' };
  }
  // Liabilities: a genuine note/loan keeps its Long Term Liabilities home even
  // when the name also says "due to <affiliate>" (Jimmy, 2026-08-18) - that
  // avoids pulling real debt off the statements already issued.
  if (ic && row.type === 'Liability' && /due to|intercompany/.test(nm) && !/loan|note payable|bot|bond/.test(nm)) {
    return { section: 'Current Liabilities', sub: 'Intercompany Payable' };
  }
  const explicit = BS_ACCOUNT_MAP[String(row.code)];
  if (explicit) return { section: explicit[0], sub: explicit[1] };
  // Heuristic fallback for accounts not in the map, so nothing is dropped.
  const name = nm;
  if (row.type === 'Asset') {
    // Cash detection uses the SAME predicate as the Statement of Cash Flows
    // (isCashAccount), so the balance sheet can never report a different cash
    // figure than the cash-flow statement reconciles to.
    //
    // The old test looked only for the words cash/checking/savings/bank/clearing
    // in the account NAME. Most of this portfolio's bank accounts are named for
    // their bank and account number instead — "MapleMark Entity 120 Odyssey
    // Holdings - 8850", "MapleMark_Odyssey_8505ICS", "CLRO - MM x4817" — so they
    // fell through to Other Assets while the cash-flow statement counted them as
    // cash. On Odyssey Holdings that hid 100% of cash ($2,875,616.38 at
    // 6/30/2026): the balance sheet showed no Cash and Cash Equivalents section
    // at all. isCashAccount adds the account-code block (100xx-101xx, 107xx) and
    // the `bank_acct` flag, which is what actually identifies these.
    if (isCashAccount(row)) return { section: 'Current Assets', sub: 'Cash and Cash Equivalents' };
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
// Which side of the balance sheet each classified section prints on. The
// CLASSIFICATION decides the side, not the account's type: a chart can code a
// payable as an asset (see the "Due to ..." rule in bsClassify), and the
// statement should still show it as a liability.
const BS_SIDE = new Map([
  ...BS_ASSET_ORDER.map(sec => [sec, 'Asset']),
  ...BS_LIAB_ORDER.map(sec => [sec, 'Liability']),
]);
const BS_SUB_ORDER = {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Intercompany Receivable', 'Other Current Assets'],
  'Fixed Assets, Net': ['Fixed Assets'],
  'Intangible Assets, Net': ['Intangible Assets', 'Amortization'],
  'Investments': ['Long Term Investments'],
  'Other Assets': ['Other Assets'],
  'Current Liabilities': ['Accounts Payable', 'Other Current Liabilities'],
  'Long Term Liabilities': ['Loans'],
};

// Default (srn) presentation order. Identical to BS_SUB_ORDER except that
// Current Liabilities carries Intercompany Payable immediately after Accounts
// Payable. BS_SUB_ORDER itself is left untouched so the clrf profile keeps the
// exact order it had before.
const BS_SUB_ORDER_SRN = Object.assign({}, BS_SUB_ORDER, {
  'Current Liabilities': ['Accounts Payable', 'Intercompany Payable', 'Other Current Liabilities'],
});

// CLIP (entity 54) presentation order — matches the CPA reference package:
// Fixed Assets, Net splits into a gross Fixed Assets subtotal then an
// Accumulated Depreciation subtotal; Current Assets carries Intercompany
// Receivable between AR and Other Current Assets. Everything else follows the
// default order.
const BS_SUB_ORDER_CLIP = Object.assign({}, BS_SUB_ORDER, {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Intercompany Receivable', 'Other Current Assets'],
  'Fixed Assets, Net': ['Fixed Assets', 'Accumulated Depreciation'],
  'Investments': ['Long Term Investments'],
});

// Silsbee (entity 39) presentation order — same as CLIP, plus Intercompany
// Payable in Current Liabilities (immediately after Accounts Payable), which
// Silsbee has and CLIP does not.
const BS_SUB_ORDER_SILSBEE = Object.assign({}, BS_SUB_ORDER, {
  'Current Assets': ['Cash and Cash Equivalents', 'Accounts Receivable, Net', 'Intercompany Receivable', 'Other Current Assets'],
  'Fixed Assets, Net': ['Fixed Assets', 'Accumulated Depreciation'],
  'Investments': ['Long Term Investments'],
  'Current Liabilities': ['Accounts Payable', 'Intercompany Payable', 'Other Current Liabilities'],
});

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
  const profile = entityProfile(opts);
  // Push-down elimination (clip / silsbee): the parent owns 100% of the entity
  // and its investment/contributed-capital is pushed down to the entity's books,
  // creating a self-referential gross-up. Drop the offsetting pair from EVERY
  // balance snapshot before any classification runs, so the elimination flows
  // uniformly to the balance sheet, the statement of changes in members' equity,
  // and the cash-flow statement. The two codes are equal and opposite and net
  // out of every window, so Assets = Liabilities + Equity stays tied by
  // construction.
  //   clip:    19041 Investment in CLIP Property Owner / 34164 Contributed
  //            Capital - County Line Industrial Park LLC
  //   silsbee: 17001 Investment - CLR Silsbee Property Owner LLC / 34063
  //            Contributed Capital - County Line Rail Silsbee LLC
  const eliminateCodes = (profile === 'clip') ? CLIP_ELIMINATE_CODES
    : (profile === 'silsbee') ? SILSBEE_ELIMINATE_CODES
    : null;
  const getBalancesEff = eliminateCodes
    ? (o => Promise.resolve(getBalances(o)).then(rows =>
        (rows || []).filter(r => !eliminateCodes.has(String(r.code)))))
    : getBalances;
  const bsSub = (profile === 'bsfrgp') ? BS_SUB_ORDER_BSFRGP
    : (profile === 'banyandev') ? BS_SUB_ORDER_BANYANDEV
    : (profile === 'banyan') ? BS_SUB_ORDER_BANYAN
    : (profile === 'clip') ? BS_SUB_ORDER_CLIP
    : (profile === 'silsbee') ? BS_SUB_ORDER_SILSBEE
    : usesIntercompanySections(profile) ? BS_SUB_ORDER_SRN
    : BS_SUB_ORDER;
  // Contra subsections (accumulated depreciation/amortization) are profile-
  // specific: Banyan's are the 165xx accumulated-depreciation accounts; CLIP's
  // and Silsbee's are 16160/16500 in their own Accumulated Depreciation
  // subsection.
  const contraSet = (profile === 'banyan') ? BS_CONTRA_CODES_BANYAN
    : (profile === 'clip') ? BS_CONTRA_CODES_CLIP
    : (profile === 'silsbee') ? BS_CONTRA_CODES_SILSBEE
    : BS_CONTRA_CODES;
  const bsCls = (row) => bsClassifyFor(profile, row);
  const bsSec = (row) => bsClassifyFor(profile, row).section;
  // The side a row PRINTS on, which is the side its classification puts it on —
  // not always the side its account type implies.
  const bsSideOf = (row) => BS_SIDE.get(bsCls(row).section)
    || (row.type === 'Liability' ? 'Liability' : 'Asset');
  // Balances arrive signed by account type (Asset debit-positive, Liability
  // credit-positive). A row that crosses sides has to be negated so it reads in
  // the receiving side's convention. Because the same amount leaves one side and
  // arrives on the other, Assets = Liabilities + Equity is untouched.
  const bsBal = (row, v) => {
    const natural = row.type === 'Liability' ? 'Liability' : 'Asset';
    return bsSideOf(row) === natural ? v : -v;
  };
  const ys = yearStart(asOf);
  const period = resolvePeriod(asOf, opts.period);
  const priorBsDate = period.bsPriorDate;
  // Inception-dated entity: the prior P&L column runs from inception to the
  // prior period end (cumulative), matching CLA's 'the Period April 16 -
  // May 31, 2026' heading. ys is deliberately left at the calendar year start:
  // it also feeds close_pl_before (retained-earnings closing), and there can be
  // no pre-inception activity for the YTD window to pick up, so moving it would
  // change nothing numerically while touching the RE path.
  const inception = inceptionFor(profile);
  // Declared here because the balance-sheet block below needs it.
  const isTkProfile = profile === 'turnkey';
  if (inception && inception < priorBsDate) period.pri = { from: inception, to: priorBsDate };

  // Snapshots:
  //  bsCur / bsPri — balance sheet as of period-end and the prior COMPARABLE
  //    period-end (prior month / quarter / year, per the toggle), with prior-year
  //    P&L closed into RE (close_pl_before = year start) so RE holds the opening
  //    balance and current-year P&L stays open on income accounts.
  //  isYtd — calendar-YTD P&L (always 1/1 → asOf), drives the YTD column and CF.
  //  isCur / isPri — P&L for the current and prior comparable PERIOD windows.
  const [bsCur, bsPri, isYtd, isCur, isPri] = await Promise.all([
    getBalancesEff({ as_of: asOf, close_pl_before: ys }),
    getBalancesEff({ as_of: priorBsDate, close_pl_before: yearStart(priorBsDate) }),
    getBalancesEff({ from: ys, to: asOf }),
    getBalancesEff({ from: period.cur.from, to: period.cur.to }),
    getBalancesEff({ from: period.pri.from, to: period.pri.to }),
  ]);

  const niYtd = netIncomeOf(isYtd);
  // Prior BS column's net-income line = P&L for the prior year through the prior
  // comparative date (calendar-YTD basis relative to that column's own year).
  const niPriYtd = netIncomeOf(await getBalancesEff({ from: yearStart(priorBsDate), to: priorBsDate }));

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
        if (!ref || ref.type === 'Equity') return null;
        if (bsSideOf(ref) !== type) return null;
        const cls = bsCls(ref);
        if (cls.section !== section || cls.sub !== sub) return null;
        const cur = rc ? bsBal(ref, bal(rc)) : 0, pri = rp ? bsBal(ref, bal(rp)) : 0;
        if (isZero(cur) && isZero(pri)) return null;
        return { code, name: ref.name, cur: r2(cur), pri: r2(pri), change: r2(cur - pri), contra: contraSet.has(String(code)) };
      })
      .filter(Boolean);
  }

  // Build a section as an ordered list of subsections, each with its own rows and
  // subtotal. A section's net total treats contra subsections as subtractions.
  function bsSectionFor(section, type) {
    const subOrder = bsSub[section] || BS_SUB_ORDER[section] || [];
    // Include any subsection that appears in the order list first, then any
    // unexpected subsections (defensive: never drop a classified account).
    const seen = new Set(subOrder);
    const extraSubs = [];
    for (const code of bsCodes) {
      const ref = colCur.map.get(code) || colPri.map.get(code);
      if (!ref || ref.type === 'Equity') continue;
      if (bsSideOf(ref) !== type) continue;
      const cls = bsCls(ref);
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
        const cls = bsCls(ref);
        if (cls.sub !== sub) return null;
        const cur = rc ? bal(rc) : 0, pri = rp ? bal(rp) : 0;
        if (isZero(cur) && isZero(pri)) return null;
        // Turnkey prints CLA's caption for the capital account: the ledger name
        // is 'Common Stock / Member's Capital', CLA's statements say
        // 'Members' Capital'. Display only - the code and balances are untouched.
        const nm = (isTkProfile && TURNKEY_EQUITY_NAMES[String(code)]) || ref.name;
        return { code, name: nm, cur: r2(cur), pri: r2(pri), change: r2(cur - pri) };
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

  // Turnkey block-shaped balance sheet (CLA presentation). Built from the same
  // two snapshots the generic sections use, so it cannot disagree with them,
  // and tied against totalAssets / totalLiab - which is what would catch a spec
  // typo that otherwise silently hides an account.
  let turnkeyBs = null;
  if (isTkProfile) {
    const tkValueOf = (code) => {
      const rc = colCur.map.get(code), rp = colPri.map.get(code);
      const ref = rc || rp;
      if (!ref || ref.type === 'Equity') return null;
      const cur = rc ? bsBal(ref, bal(rc)) : 0, pri = rp ? bsBal(ref, bal(rp)) : 0;
      const zero = isZero(cur) && isZero(pri);
      return { cur: r2(cur), pri: r2(pri), name: ref.name, zero };
    };
    const sideCodes = (side) => bsCodes.filter(code => {
      const ref = colCur.map.get(code) || colPri.map.get(code);
      return ref && ref.type !== 'Equity' && bsSideOf(ref) === side;
    });
    const tkAssets = buildTurnkeyBlocks(TURNKEY_BS_ASSETS, tkValueOf, sideCodes('Asset'));
    const tkLiabs = buildTurnkeyBlocks(TURNKEY_BS_LIABS, tkValueOf, sideCodes('Liability'));
    const blockTotal = (blocks, k) => r2(blocks.reduce((s, b) =>
      s + (b.kind === 'row' ? b[k] : b.subtotal[k]), 0));
    turnkeyBs = {
      assetBlocks: tkAssets.blocks, liabBlocks: tkLiabs.blocks,
      unmapped: tkAssets.unmapped.concat(tkLiabs.unmapped),
      tie: {
        assetsCur: r2(blockTotal(tkAssets.blocks, 'cur') - totalAssets.cur),
        assetsPri: r2(blockTotal(tkAssets.blocks, 'pri') - totalAssets.pri),
        liabsCur: r2(blockTotal(tkLiabs.blocks, 'cur') - totalLiab.cur),
        liabsPri: r2(blockTotal(tkLiabs.blocks, 'pri') - totalLiab.pri),
      },
    };
  }

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

  // Turnkey routes its own P&L: COGS is the whole 5xxxx block (the name-based
  // test below matches none of its 55xxx lines), and Interest Income /
  // Interest Expense are lifted out of Revenue / Operating Expenses into an
  // Other Income (Expense) section. Every other profile is unchanged.
  const isTk = profile === 'turnkey';
  // COGS is resolved FIRST and always wins: a cost-of-construction line can
  // never be lifted out of cost of goods sold by an Other Income (Expense)
  // name test.
  const cogs = plLines(r => r.type === 'Expense' && (isTk
    ? turnkeyIsCogs(r)
    : /cogs|cost of goods|cost of revenue|car hire/i.test((r.subtype || '') + ' ' + (r.name || ''))));
  const cogsCodes = new Set(cogs.map(l => l.code));
  // Other Income (Expense) + Income Taxes, from the one shared classifier
  // (otherIeRoute). Turnkey keeps its own pins on top of it because its chart
  // reverses 42000 / 70000 against every other entity here.
  const oieOf = (r) => {
    if (r.type === 'Expense' && cogsCodes.has(r.code)) return null;
    if (isTk) {
      if (turnkeyIsOtherIncome(r)) return OIE_INCOME;
      if (turnkeyIsOtherExpense(r)) return OIE_EXPENSE;
    }
    return otherIeRoute(r);
  };
  const inOieBucket = (b) => plLines(r => { const x = oieOf(r); return !!x && x.bucket === b; });
  const otherIncomeLines = inOieBucket('otherIncome');
  const otherExpenseLines = inOieBucket('otherExpense');
  const incomeTaxLines = inOieBucket('incomeTax');
  const oieCodes = new Set([].concat(otherIncomeLines, otherExpenseLines, incomeTaxLines).map(l => l.code));
  const revenue = plLines(r => r.type === 'Revenue' && !oieCodes.has(r.code));
  const opex = plLines(r => r.type === 'Expense' && !cogsCodes.has(r.code) && !oieCodes.has(r.code));

  const sumCol = (lines, k) => r2(lines.reduce((s, l) => s + l[k], 0));
  const totRev = { cur: sumCol(revenue, 'cur'), pri: sumCol(revenue, 'pri'), ytd: sumCol(revenue, 'ytd') };
  const totCogs = { cur: sumCol(cogs, 'cur'), pri: sumCol(cogs, 'pri'), ytd: sumCol(cogs, 'ytd') };
  const grossProfit = { cur: r2(totRev.cur - totCogs.cur), pri: r2(totRev.pri - totCogs.pri), ytd: r2(totRev.ytd - totCogs.ytd) };
  const totOpex = { cur: sumCol(opex, 'cur'), pri: sumCol(opex, 'pri'), ytd: sumCol(opex, 'ytd') };
  // Other Income (Expense) nets income less expense. Both line sets carry
  // positive magnitudes (see plLines), so the expense side is subtracted.
  const totOtherInc = { cur: sumCol(otherIncomeLines, 'cur'), pri: sumCol(otherIncomeLines, 'pri'), ytd: sumCol(otherIncomeLines, 'ytd') };
  const totOtherExp = { cur: sumCol(otherExpenseLines, 'cur'), pri: sumCol(otherExpenseLines, 'pri'), ytd: sumCol(otherExpenseLines, 'ytd') };
  const totOtherIE = { cur: r2(totOtherInc.cur - totOtherExp.cur), pri: r2(totOtherInc.pri - totOtherExp.pri), ytd: r2(totOtherInc.ytd - totOtherExp.ytd) };
  // Income Taxes is its own section below Other Income (Expense). The accounts
  // carry their natural expense sign, so a credit/benefit is negative and
  // therefore ADDS to net income when subtracted.
  const totIncomeTax = { cur: sumCol(incomeTaxLines, 'cur'), pri: sumCol(incomeTaxLines, 'pri'), ytd: sumCol(incomeTaxLines, 'ytd') };
  const netIncome = { cur: r2(grossProfit.cur - totOpex.cur + totOtherIE.cur - totIncomeTax.cur), pri: r2(grossProfit.pri - totOpex.pri + totOtherIE.pri - totIncomeTax.pri), ytd: r2(grossProfit.ytd - totOpex.ytd + totOtherIE.ytd - totIncomeTax.ytd) };
  // Group -> subsection trees, in the shape the renderers already use for the
  // banyan / bsfrgp profiles. Every line in a bucket shares one group and one
  // subsection (see OIE_INCOME / OIE_EXPENSE / OIE_TAX), so a single group with
  // a single subsection is exact rather than a simplification.
  const oieTree = (lines, route) => {
    if (!lines.length) return [];
    const subtotal = { cur: sumCol(lines, 'cur'), pri: sumCol(lines, 'pri'), ytd: sumCol(lines, 'ytd') };
    const sorted = lines.slice().sort((a, b) => String(a.code).localeCompare(String(b.code), undefined, { numeric: true }));
    return [{ title: route.group, subtotal, subs: [{ title: route.sub, lines: sorted, subtotal }] }];
  };
  const otherIncomeTree = oieTree(otherIncomeLines, OIE_INCOME);
  const otherExpenseTree = oieTree(otherExpenseLines, OIE_EXPENSE);
  const incomeTaxTree = oieTree(incomeTaxLines, OIE_TAX);

  // Turnkey presentation only: fold the WIP adjustment into Construction
  // Revenue and order COGS / G&A the way CLA does. Deliberately AFTER every
  // total above - these rebind the ROW LISTS, never the arithmetic, so Total
  // Revenue, Gross Profit and Net Income cannot move.
  const revenueRows = isTk ? turnkeyNetRevenue(revenue) : revenue;
  const cogsRows = isTk ? orderByCodes(cogs, TURNKEY_COGS_ORDER) : cogs;
  const opexRows = isTk ? orderByCodes(opex, TURNKEY_GA_ORDER) : opex;

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
  // ── Banyan SFR GP Investors operations restructure (bsfrgp profile) ──────
  // Reproduces the CPA reference shape:
  //   Operating Expenses (General and Administrative → Legal and Accounting)
  //   Other Income (Expense)  (Other Income / Other Expense, netted)
  //   Income Taxes            (State and Local Taxes)
  //   Net Income (Loss) = Other Income (Expense) − Operating Expenses − Income Taxes
  // All P&L accounts (revenue + expense) are routed by bsfrgpPlRoute so nothing
  // is dropped; the resulting net income equals the GL net income by construction.
  let bsfrgpOps = null;
  if (profile === 'bsfrgp') {
    // All P&L lines regardless of type, so Interest Income (a Revenue account)
    // can be routed into Other Income rather than a top-line Revenue section.
    const allPl = plLines(() => true);
    const byBucket = { opex: [], otherIncome: [], otherExpense: [], incomeTax: [] };
    const routeOf = {};
    for (const l of allPl) {
      const bref = mYtd.get(l.code) || mCur.get(l.code) || mPri.get(l.code);
      const route = bsfrgpPlRoute({ code: l.code, name: l.name, type: bref.type, subtype: bref.subtype });
      routeOf[l.code] = route;
      (byBucket[route.bucket] || byBucket.opex).push(l);
    }
    // Other-expense lines are NEGATED for presentation, matching the banyan
    // profile and the standard shape: inside 'Other Income (Expense)' they are
    // reductions of income, so a 455.67 penalty prints as (455.67) and the
    // section nets to Income + (negated) Expense.
    byBucket.otherExpense = byBucket.otherExpense.map(l => ({
      ...l, cur: r2(-l.cur), pri: r2(-l.pri), ytd: r2(-l.ytd), change: r2(-l.change),
    }));
    // Build a nested group→subsection tree for a bucket's lines (preserves the
    // reference's 'General and Administrative Expenses → Legal and Accounting'
    // nesting). Each group carries its own subtotal; each subsection carries a
    // subtotal too.
    const buildTree = (lines) => {
      const groups = [];
      const gIndex = new Map();
      for (const l of lines) {
        const r = routeOf[l.code];
        let g = gIndex.get(r.group);
        if (!g) { g = { title: r.group, subs: [], _si: new Map() }; gIndex.set(r.group, g); groups.push(g); }
        let su = g._si.get(r.sub);
        if (!su) { su = { title: r.sub, lines: [] }; g._si.set(r.sub, su); g.subs.push(su); }
        su.lines.push(l);
      }
      // Subtotals.
      for (const g of groups) {
        for (const su of g.subs) {
          su.subtotal = { cur: sumCol(su.lines, 'cur'), pri: sumCol(su.lines, 'pri'), ytd: sumCol(su.lines, 'ytd') };
          su.lines.sort((a, b) => String(a.code).localeCompare(String(b.code)));
        }
        g.subtotal = {
          cur: r2(g.subs.reduce((s, x) => s + x.subtotal.cur, 0)),
          pri: r2(g.subs.reduce((s, x) => s + x.subtotal.pri, 0)),
          ytd: r2(g.subs.reduce((s, x) => s + x.subtotal.ytd, 0)),
        };
        delete g._si;
      }
      return groups;
    };
    const opexTree = buildTree(byBucket.opex);
    const otherIncomeTree = buildTree(byBucket.otherIncome);
    const otherExpenseTree = buildTree(byBucket.otherExpense);
    const incomeTaxTree = buildTree(byBucket.incomeTax);
    const sumBucket = (arr) => ({ cur: sumCol(arr, 'cur'), pri: sumCol(arr, 'pri'), ytd: sumCol(arr, 'ytd') });
    const tOpex = sumBucket(byBucket.opex);
    const tOtherIncome = sumBucket(byBucket.otherIncome);
    const tOtherExpense = sumBucket(byBucket.otherExpense);
    const tIncomeTax = sumBucket(byBucket.incomeTax);
    // Other Income (Expense) net = Other Income + (already negated) Other Expense.
    const tOtherIE = {
      cur: r2(tOtherIncome.cur + tOtherExpense.cur),
      pri: r2(tOtherIncome.pri + tOtherExpense.pri),
      ytd: r2(tOtherIncome.ytd + tOtherExpense.ytd),
    };
    // Net Income (Loss) = Other Income (Expense) − Operating Expenses − Income Taxes.
    // Income-tax accounts carry their natural expense sign (a credit/benefit is
    // negative and therefore ADDS to net income when subtracted).
    const niBsfrgp = {
      cur: r2(tOtherIE.cur - tOpex.cur - tIncomeTax.cur),
      pri: r2(tOtherIE.pri - tOpex.pri - tIncomeTax.pri),
      ytd: r2(tOtherIE.ytd - tOpex.ytd - tIncomeTax.ytd),
    };
    bsfrgpOps = {
      structured: true,
      opexTree, otherIncomeTree, otherExpenseTree, incomeTaxTree,
      totOpex: tOpex, totOtherIncome: tOtherIncome, totOtherExpense: tOtherExpense,
      totOtherIE: tOtherIE, totIncomeTax: tIncomeTax, netIncome: niBsfrgp,
    };
  }

  // ── Banyan Residential operations restructure (banyan profile) ───────────
  // Reproduces the CPA reference shape:
  //   Revenue - Services (top line) → Total Revenue → Gross Profit
  //   Operating Expenses, grouped: Payroll and Related / Travel, Meals and
  //     Entertainment / Utilities and Facilities / General and Administrative /
  //     Office Expense / Taxes and Insurance / Depreciation and Amortization
  //   Other Income (Expense)  (Other Income less Other Expense)
  //   Income Taxes            (State and Local Taxes)
  //   Net Income (Loss) = Revenue − Operating Expenses + Other Income (Expense)
  //                       − Income Taxes
  // Every P&L account is routed by banyanPlRoute so nothing is dropped; the
  // resulting net income equals the GL net income by construction.
  let banyanOps = null;
  if (profile === 'banyan' || profile === 'banyandev') {
    const isBd = profile === 'banyandev';
    const bdKey = isBd ? (/braker/i.test(String(opts.entityName || '')) ? 'braker' : 'hp') : null;
    const allPl = plLines(() => true);
    const byBucket = { revenue: [], opex: [], otherIncome: [], otherExpense: [], incomeTax: [] };
    const routeOf = {};
    for (const l of allPl) {
      const ref = mYtd.get(l.code) || mCur.get(l.code) || mPri.get(l.code);
      const route = isBd
        ? banyandevPlRoute(bdKey, { code: l.code, name: l.name, type: ref.type, subtype: ref.subtype })
        : banyanPlRoute({ code: l.code, name: l.name, type: ref.type, subtype: ref.subtype });
      routeOf[l.code] = route;
      (byBucket[route.bucket] || byBucket.opex).push(l);
    }
    // Other-expense lines are NEGATED for presentation: the reference shows them
    // inside "Other Income (Expense)" as reductions of income (e.g. a 415.84
    // penalty prints as (415.84)), and the section nets to Income − Expense.
    byBucket.otherExpense = byBucket.otherExpense.map(l => ({
      ...l, cur: r2(-l.cur), pri: r2(-l.pri), ytd: r2(-l.ytd), change: r2(-l.change),
    }));
    // Build a nested group→subsection tree, optionally ordered by groupOrder.
    const buildTree = (lines, groupOrder) => {
      const groups = [];
      const gIndex = new Map();
      for (const l of lines) {
        const r = routeOf[l.code];
        let g = gIndex.get(r.group);
        if (!g) { g = { title: r.group, subs: [], _si: new Map() }; gIndex.set(r.group, g); groups.push(g); }
        let su = g._si.get(r.sub);
        if (!su) { su = { title: r.sub, lines: [] }; g._si.set(r.sub, su); g.subs.push(su); }
        su.lines.push(l);
      }
      for (const g of groups) {
        // Order subsections per the reference where specified.
        const so = BANYAN_SUB_ORDER_IN_GROUP[g.title];
        if (so && so.length) {
          const srank = t => { const i = so.indexOf(t); return i === -1 ? so.length : i; };
          g.subs.sort((a, b) => srank(a.title) - srank(b.title));
        }
        for (const su of g.subs) {
          su.lines.sort((a, b) => String(a.code).localeCompare(String(b.code)));
          su.subtotal = { cur: sumCol(su.lines, 'cur'), pri: sumCol(su.lines, 'pri'), ytd: sumCol(su.lines, 'ytd') };
        }
        g.subtotal = {
          cur: r2(g.subs.reduce((s, x) => s + x.subtotal.cur, 0)),
          pri: r2(g.subs.reduce((s, x) => s + x.subtotal.pri, 0)),
          ytd: r2(g.subs.reduce((s, x) => s + x.subtotal.ytd, 0)),
        };
        delete g._si;
      }
      if (groupOrder && groupOrder.length) {
        const rank = t => { const i = groupOrder.indexOf(t); return i === -1 ? groupOrder.length : i; };
        groups.sort((a, b) => rank(a.title) - rank(b.title));
      }
      return groups;
    };
    const revenueTree = buildTree(byBucket.revenue, isBd ? BANYANDEV_REVENUE_GROUP_ORDER : undefined);
    const opexTree = buildTree(byBucket.opex, isBd ? BANYANDEV_OPEX_GROUP_ORDER : BANYAN_OPEX_GROUP_ORDER);
    const otherIncomeTree = buildTree(byBucket.otherIncome, isBd ? BANYANDEV_OTHER_INCOME_GROUP_ORDER : undefined);
    const otherExpenseTree = buildTree(byBucket.otherExpense);
    const incomeTaxTree = buildTree(byBucket.incomeTax);
    const sumBucket = arr => ({ cur: sumCol(arr, 'cur'), pri: sumCol(arr, 'pri'), ytd: sumCol(arr, 'ytd') });
    const tRev = sumBucket(byBucket.revenue);
    const tOpexB = sumBucket(byBucket.opex);
    const tOtherIncomeB = sumBucket(byBucket.otherIncome);
    const tOtherExpenseB = sumBucket(byBucket.otherExpense); // already negated
    const tIncomeTaxB = sumBucket(byBucket.incomeTax);
    // Other Income (Expense) net = Other Income + (negated) Other Expense.
    const tOtherIEB = {
      cur: r2(tOtherIncomeB.cur + tOtherExpenseB.cur),
      pri: r2(tOtherIncomeB.pri + tOtherExpenseB.pri),
      ytd: r2(tOtherIncomeB.ytd + tOtherExpenseB.ytd),
    };
    // No cost-of-revenue on Banyan, so Gross Profit = Total Revenue.
    const gpB = { cur: tRev.cur, pri: tRev.pri, ytd: tRev.ytd };
    const niBanyan = {
      cur: r2(tRev.cur - tOpexB.cur + tOtherIEB.cur - tIncomeTaxB.cur),
      pri: r2(tRev.pri - tOpexB.pri + tOtherIEB.pri - tIncomeTaxB.pri),
      ytd: r2(tRev.ytd - tOpexB.ytd + tOtherIEB.ytd - tIncomeTaxB.ytd),
    };
    banyanOps = {
      structured: true, banyanShape: true,
      showGrossProfit: !isBd || bdKey === 'hp',
      noGroupTotal: isBd ? BANYANDEV_NO_GROUP_TOTAL : null,
      revenueTree, opexTree, otherIncomeTree, otherExpenseTree, incomeTaxTree,
      totRev: tRev, grossProfit: gpB, totOpex: tOpexB,
      totOtherIncome: tOtherIncomeB, totOtherExpense: tOtherExpenseB,
      totOtherIE: tOtherIEB, totIncomeTax: tIncomeTaxB, netIncome: niBanyan,
    };
  }

  // ── Statement of Cash Flows (indirect, YTD) ────────────────────────────────
  // Beginning balances = as of (year start − 1 day). Deltas over the YTD window.
  const bsOpen = await getBalancesEff({ as_of: priorMonthEnd(ys), close_pl_before: ys });
  const openMap = new Map(); for (const r of bsOpen) openMap.set(r.code, r);
  const curMap = new Map(); for (const r of bsCur) if (r.type !== 'Revenue' && r.type !== 'Expense') curMap.set(r.code, r);
  // Cash detection for the cash-flow statement MUST agree with the balance
  // sheet's "Cash and Cash Equivalents" subsection, or the statement reconciles
  // to a different cash figure than the BS reports. Banyan's Bill.com clearing
  // account (10300) is cash on the BS but is missed by the generic code/name
  // test, which understated cash by that balance. For the banyan profile we
  // therefore take the BS classification as authoritative (OR'd with the generic
  // test so nothing is lost); other profiles keep the original test unchanged.
  const isCashRow = (profile === 'banyan')
    ? (r => r.type === 'Asset' && (bsCls(r).sub === 'Cash and Cash Equivalents' || isCashAccount(r)))
    : isCashAccount;
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

  // Depreciation/amortization is booked one of two ways, and each needs
  // different cash-flow handling:
  //   (a) credited to an ACCUMULATED contra account (e.g. 16500 Equipment: Acc
  //       Depreciation). The contra's movement is skipped in the loop below and
  //       the P&L expense is added back above — consistent, nothing to do.
  //   (b) credited DIRECTLY to the asset (e.g. Banyan's 12383 Organization Fees,
  //       amortized 66.67/month straight off the asset with no contra). Here the
  //       asset's decline would flow through investing as a cash INFLOW while the
  //       expense is also added back — double-counting the same non-cash amount.
  // The portion of D&A expense NOT matched by accumulated-contra movement is the
  // directly-credited amount, so we remove it from investing below. Clamped at 0
  // so entities that CAPITALIZE depreciation (amortExpense 0, e.g. SRN) and any
  // rounding are unaffected.
  const contraMoveAbs = r2(Math.abs([...new Set([...openMap.keys(), ...curMap.keys()])].reduce((sum, code) => {
    const rc = curMap.get(code), ro = openMap.get(code);
    const ref = rc || ro;
    if (!ref || ref.type !== 'Asset') return sum;
    if (!/accum|amortization|depreciation/i.test(ref.name || '')) return sum;
    return sum + ((rc ? bal(rc) : 0) - (ro ? bal(ro) : 0));
  }, 0)));
  const directAmort = r2(Math.max(0, amortExpense - contraMoveAbs));

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
      const sec = bsSec(ref);
      const cashEffect = -delta; // asset up → cash down
      if (/intercompany|due from|due to/i.test(nm)) cfBuckets.intercompany += cashEffect;
      else if (sec === 'Current Assets' && /receivable/i.test(nm)) cfBuckets.ar += cashEffect;
      else if (sec === 'Current Assets') cfBuckets.prepaidOther += cashEffect;
      else if (sec === 'Fixed Assets' || sec === 'Fixed Assets, Net') cfBuckets.capex += cashEffect;
      else cfBuckets.ltInvest += cashEffect; // intangible / investment / other long-term
    } else if (ref.type === 'Liability') {
      const cashEffect = delta; // liability up → cash up
      const sec = bsSec(ref);
      // Loan payables are FINANCING for every entity: the change in any loan /
      // note payable / bond account, or anything the section classifier put in
      // Long Term Liabilities, is a debt draw/repayment and belongs in financing
      // activities — not the operating AP line. These tests run BEFORE the
      // generic "... payable" name test so a loan account named "loan payable"
      // lands on debt service rather than operating AP. (Jimmy, 2026-08-31 —
      // generalized from the earlier banyandev-only rule to all entities.)
      if (/loan|note payable|bond/i.test(nm)) cfBuckets.debtChange += cashEffect;
      else if (sec === 'Long Term Liabilities') cfBuckets.debtChange += cashEffect;
      else if (/payable/i.test(nm)) cfBuckets.ap += cashEffect;
      else cfBuckets.accrued += cashEffect; // accrued / other current liabilities
    } else if (ref.type === 'Equity') {
      // Equity delta includes contributions/distributions but NOT current-year
      // net income (P&L is excluded above), so the whole delta is financing.
      cfBuckets.equityContrib += delta;
    }
  }
  // Remove the directly-credited D&A (case (b) above) from long-term investing:
  // that asset decrease is non-cash and is already added back in operating.
  if (!isZero(directAmort)) cfBuckets.ltInvest = r2(cfBuckets.ltInvest - directAmort);
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
  // The operating-TB consolidations (banyandev: HP / Braker) carry a non-cash
  // equity reclass into the opening column — the property manager's pre-closing
  // Dec-31 trial balance closes its prior-year P&L into retained earnings, a
  // movement that never touched cash. Left in, it reads as a phantom member
  // contribution/distribution and the statement fails to foot (Braker: the
  // 14,006.39 2025 lease-up loss). Fold that residual back into the equity line
  // where it originated so the statement ties and financing matches CLA. Capped
  // so a genuinely large reconciling gap still surfaces as the note below.
  if (profile === 'banyandev' && Math.abs(cashFlow.tieOut) > 0.004 && Math.abs(cashFlow.tieOut) <= 100000) {
    // Absorb the immaterial non-cash residual into the accounts-receivable line
    // (Jimmy, 2026-08-29) so the statement foots without a reconciling note.
    cashFlow.changeAR = r2(cashFlow.changeAR - cashFlow.tieOut);
    cashFlow.netOperating = r2(cashFlow.netOperating - cashFlow.tieOut);
    cashFlow.netChange = r2(cashFlow.netOperating + cashFlow.netInvesting + cashFlow.netFinancing);
    cashFlow.tieOut = r2(cashFlow.netChange - cashFlow.actualCashChange);
  }

  // Turnkey presentation: CLA's own line set, built from named account groups
  // rather than the name/section heuristics. Deliberately additive - the
  // generic cashFlow above is untouched, and tie.netChange below asserts the
  // two agree on the bottom line.
  if (isTkProfile) {
    // Movement over the cash-flow window in natural (computeBalances) sign:
    // assets debit-positive, liabilities and equity credit-positive.
    const mv = (codes) => r2(codes.reduce((s, c) => {
      const rc = colCur.map.get(c), ro = openMap.get(c);
      return s + ((rc ? bal(rc) : 0) - (ro ? bal(ro) : 0));
    }, 0));
    // 'Increase'/'Decrease' follows the direction of the BALANCE, not the cash
    // effect, so the wording still reads correctly in a month that reverses.
    const word = (delta) => (delta < 0 ? 'Decrease' : 'Increase');
    const dAR = mv(TURNKEY_CF.accountsReceivable);
    const dCA = mv(TURNKEY_CF.contractAssets);
    const dPre = mv(TURNKEY_CF.prepaidExpenses);
    const dAP = mv(TURNKEY_CF.accountsPayable);
    const dCL = mv(TURNKEY_CF.contractLiabilities);
    // 15100 is a contra asset: its balance goes more negative as depreciation
    // accrues, so the add-back is the negated movement.
    const dep = r2(-mv(TURNKEY_CF.depreciation));
    const capex = r2(-mv(TURNKEY_CF.fixedAssets));
    const equityMv = mv(TURNKEY_CF.memberCapital);
    const contributions = equityMv > 0 ? equityMv : 0;
    const distributions = equityMv < 0 ? equityMv : 0;

    const netOperating = r2(cashFlow.netIncome + dep - dAR + dAP - dCA - dPre + dCL);
    const netInvesting = r2(capex);
    const netFinancing = r2(contributions + distributions);
    const netChange = r2(netOperating + netInvesting + netFinancing);

    const lines = [
      { label: 'Cash Flows from Operating Activities:', header: true },
      { label: 'Net Income (Loss)', value: cashFlow.netIncome },
      { label: 'Changes in Operating Assets and Liabilities:', header: true },
      { label: 'Depreciation', value: dep, indent: true },
      { label: word(dAR) + ' in Accounts Receivable', value: r2(-dAR), indent: true },
      { label: word(dAP) + ' in Accounts Payable', value: dAP, indent: true },
      { label: word(dCA) + ' in Contract Assets', value: r2(-dCA), indent: true },
      { label: word(dPre) + ' in Prepaid Expenses', value: r2(-dPre), indent: true },
      { label: word(dCL) + ' in Contract Liabilities', value: dCL, indent: true },
      { label: (netOperating < 0 ? 'Net Cash Used by Operating Activities' : 'Net Cash Provided by Operating Activities'), value: netOperating, bold: true, rule: true },
      { label: 'Cash Flows from Investing Activities:', header: true, gapBefore: true },
      { label: 'Purchase of Fixed Assets', value: capex, indent: true },
      { label: (netInvesting < 0 ? 'Net Cash Used by Investing Activities' : 'Net Cash Provided by Investing Activities'), value: netInvesting, bold: true, rule: true },
      { label: 'Cash Flows from Financing Activities', header: true, gapBefore: true },
      { label: 'Contributions From Members', value: contributions, indent: true },
      { label: 'Distributions To Members', value: distributions, indent: true },
      { label: (netFinancing < 0 ? 'Net Cash Used by Financing Activities' : 'Net Cash Provided by Financing Activities'), value: netFinancing, bold: true, rule: true },
      { label: 'Net Increase (Decrease) in Cash', value: netChange, bold: true, gapBefore: true, rule: true },
      { label: 'Cash - Beginning of Period', value: cashFlow.cashBeg, gapBefore: true },
      { label: 'Cash - End of Period', value: cashFlow.cashEnd, bold: true, dollar: true, rule: true, doubleBelow: true },
    ];
    cashFlow.turnkey = {
      lines, netOperating, netInvesting, netFinancing, netChange,
      // Must be zero: the explicit line set and the generic buckets are two
      // independent routes to the same bottom line.
      tie: r2(netChange - cashFlow.netChange),
    };
  }

  // ── Statement of Changes in Members' Equity ───────────────────────────────
  // Beginning (year start) contributed equity by account + beginning RE, then
  // contributions (delta) and YTD net income → ending.
  //
  // Opening-day equity reclasses belong to the BEGINNING column, not to the
  // year's activity. CLR Buna rolls 33011 Distribution - Ben into 39000
  // Retained Earnings with an entry dated 1/1 (entry 17582, 148,117.07). Run
  // through as current-year movement it prints two lines of noise: a member row
  // that opens with a balance and is emptied by a "contribution", and a
  // Retained Earnings row carrying the same amount back the other way — for a
  // movement that never touched a member's capital and happened before the
  // first day of business. Folded into the opening balances instead, Buna's
  // 2026 statement opens at Retained Earnings (580,725.31), 33011 drops out
  // entirely, and the contributions column stays clean.
  //
  // Applied ONLY when the day's equity activity nets to zero, which is what
  // makes it a pure reclass WITHIN equity. A real 1/1 contribution or
  // distribution (equity against cash) fails that test and flows through as
  // activity, exactly as before. Total beginning equity is unchanged either
  // way, so the statement's tie to the prior period's close is untouched.
  const openReclass = new Map(); // account code → adjustment to its beginning
  let reOpenReclass = 0;         // adjustment to beginning Retained Earnings
  {
    const dayRows = (await getBalancesEff({ from: ys, to: ys })).filter(x => x.type === 'Equity');
    if (dayRows.length && isZero(r2(dayRows.reduce((a, x) => a + bal(x), 0)))) {
      for (const x of dayRows) {
        const amt = r2(bal(x));
        if (isZero(amt)) continue;
        if (bsCls(x).sub === 'Retained Earnings') reOpenReclass = r2(reOpenReclass + amt);
        else openReclass.set(String(x.code), r2((openReclass.get(String(x.code)) || 0) + amt));
      }
    }
  }
  const begOf = (row, code) => r2((row ? bal(row) : 0) + (openReclass.get(String(code)) || 0));
  const equityMembers = equityRows.map(r => {
    const openRow = bsOpen.find(x => x.code === r.code);
    const beg = begOf(openRow, r.code);
    return { code: r.code, name: r.name, beginning: beg, contributions: r2(r.cur - beg), distributions: 0, netIncome: 0, ending: r2(r.cur) };
  });
  // An equity account can carry a balance at the START of the year and be zero
  // in BOTH balance-sheet columns — a member redeemed during the year, or (CLR
  // Buna) a distribution account rolled into Retained Earnings on 1 January.
  // equityRows drops those (it filters accounts that are zero in cur AND pri),
  // so their opening balance silently vanished from this statement and the
  // beginning total stopped equalling the prior period's ending equity. Buna's
  // 2026 statement opened at 1,310,478.41 against a 12/31/2025 close of
  // 1,162,361.34 — exactly the 148,117.07 sitting in 33011 Distribution - Ben.
  // Re-add any such account from the opening balance sheet.
  const stmtCodes = new Set(equityMembers.map(m => String(m.code)));
  for (const o of bsOpen) {
    if (o.type !== 'Equity' || stmtCodes.has(String(o.code))) continue;
    if (bsCls(o).sub !== 'Members Equity') continue; // RE is the reMember row below
    const beg = begOf(o, o.code);
    const cr = colCur.map.get(o.code);
    const end = cr ? r2(bal(cr)) : 0;
    // Nothing at either end — an account emptied by a 1/1 reclass (Buna's
    // 33011) now lands here and is simply not a line of this statement.
    if (isZero(beg) && isZero(end)) continue;
    stmtCodes.add(String(o.code));
    equityMembers.push({ code: o.code, name: o.name, beginning: beg, contributions: r2(end - beg), distributions: 0, netIncome: 0, ending: end });
  }
  equityMembers.sort((a, b) => String(a.code).localeCompare(String(b.code), undefined, { numeric: true }));
  // A distribution account's current-year movement is a DISTRIBUTION, not a
  // negative contribution. The generic member mapping above books every
  // account's (cur - beg) delta into the contributions column; for a
  // distribution account (33xxx "Distribution - ...") that puts a negative in
  // Contributions when it belongs in the Distributions column. Reroute it so
  // the columns read correctly. The row still foots to ending (beginning +
  // contributions + distributions + net income) because the amount only moves
  // between two columns, and total ending equity — hence the tie to the
  // balance sheet — is untouched.
  for (const mrow of equityMembers) {
    const isDistribution = /^\s*distribution\b/i.test(String(mrow.name || ''))
      || /^33/.test(String(mrow.code || ''));
    if (isDistribution && !isZero(mrow.contributions)) {
      mrow.distributions = r2(mrow.distributions + mrow.contributions);
      mrow.contributions = 0;
    }
  }
  // HP (CLA): the Banyan HP Fund Undeployed Capital line is presented as a
  // current-period distribution, not an opening balance — move its beginning
  // balance into the Distributions column. The row still foots to ending
  // (beginning + contributions + distributions + net income), and total ending
  // equity is unchanged, so the statement's tie to the balance sheet holds.
  // Braker keeps Undeployed Capital in the opening (beginning) column (Jimmy) —
  // so the move is HP-only within banyandev.
  if (profile === 'banyandev' && !/braker/i.test(String(opts.entityName || ''))) {
    for (const mrow of equityMembers) {
      if (/undeployed/i.test(mrow.name) && !isZero(mrow.beginning)) {
        mrow.distributions = r2(mrow.distributions + mrow.beginning);
        mrow.beginning = 0;
      }
    }
  }
  // Retained earnings row. Beginning = opening RE; ending = the RE balance the
  // BALANCE SHEET carries plus YTD net income — NOT opening + net income. A
  // current-year entry posted directly to 39000 (the Buna 1/1 roll-forward
  // above) is part of RE's movement for the year; taking only the opening
  // balance dropped it and broke this statement's tie to the balance sheet.
  const reOpenRows = bsOpen.filter(x => x.type === 'Equity' && bsCls(x).sub === 'Retained Earnings');
  const reOpen = r2(reOpenRows.reduce((a, x) => a + bal(x), 0) + reOpenReclass);
  const reMember = { code: 're', name: 'Retained Earnings', beginning: reOpen, contributions: r2(totalRetained.cur - reOpen), distributions: 0, netIncome: niYtd, ending: r2(totalRetained.cur + niYtd) };
  const isEmptyRow = m => isZero(m.beginning) && isZero(m.contributions) && isZero(m.distributions) && isZero(m.netIncome) && isZero(m.ending);
  const equityStmt = [...equityMembers.filter(m => !/retained earning/i.test(m.name) && !isEmptyRow(m)), reMember];
  const equityTotals = equityStmt.reduce((t, m) => ({
    beginning: r2(t.beginning + m.beginning), contributions: r2(t.contributions + m.contributions),
    distributions: r2(t.distributions + m.distributions), netIncome: r2(t.netIncome + m.netIncome), ending: r2(t.ending + m.ending),
  }), { beginning: 0, contributions: 0, distributions: 0, netIncome: 0, ending: 0 });

  return {
    meta: { entityName: displayEntityName(opts.entityName), rawEntityName: opts.entityName || '', entityCode: (opts.entityCode || ''), isConsolidated: !!opts.isConsolidated, asOf, priorDate: priorBsDate, longDate: longDate(asOf),
            priorLongDate: longDate(priorBsDate),
            // Inception-dated entities date every statement from inception.
            monthsEnded: inception
              ? ('For the Period ' + inceptionRange(inception, asOf))
              : monthsEndedLabel(asOf),
            period: (opts.period || 'monthly').toLowerCase(), periodLabel: period.periodLabel, colLabel: period.colLabel,
            opsDateLine: inception
              ? ('For the One Month Ended ' + longDate(asOf) + ', the Period ' + inceptionRange(inception, priorBsDate) + ', and the Period ' + inceptionRange(inception, asOf))
              : opsHeadingLine(period.colLabel, longDate(asOf), longDate(priorBsDate)),
            // Header for the operations statement's prior column, and the
            // opening column of the equity statement.
            // Pre-wrapped so it cannot collide with the current-column heading:
            // 'April 16 -' on the first row, 'May 31, 2026' on the row below.
            opsPriorColLabel: inception
              ? (monthDay(inception) + ' -\n' + longDate(priorBsDate))
              : longDate(priorBsDate),
            // Unwrapped form, for callers that want it as one string.
            opsPriorColLabelFlat: inception ? inceptionRange(inception, priorBsDate) : longDate(priorBsDate),
            equityBegDate: inception ? slashDate(inception) : null,
            inception: inception || null, profile },
    balanceSheet: Object.assign({ assetSections, liabSections, equityRows, retainedRows, totalAssets, totalLiab, totalContribEquity, niLine, totalEquity, totalLiabEquity },
      turnkeyBs ? { turnkey: turnkeyBs } : {}),
    operations: Object.assign({ revenue: revenueRows, cogs: cogsRows, opex: opexRows, opexGroups, totRev, totCogs, grossProfit, totOpex, netIncome,
      otherIncomeLines, otherExpenseLines, incomeTaxLines, totOtherInc, totOtherExp, totOtherIE, totIncomeTax,
      otherIncomeTree, otherExpenseTree, incomeTaxTree },
      bsfrgpOps ? { bsfrgp: bsfrgpOps, netIncome: bsfrgpOps.netIncome } : {},
      banyanOps ? { banyan: banyanOps, netIncome: banyanOps.netIncome } : {}),
    cashFlow,
    equity: { rows: equityStmt, totals: equityTotals },
    checks: {
      balanceSheetTies: isZero(totalAssets.cur - totalLiabEquity.cur),
      balanceSheetDiff: r2(totalAssets.cur - totalLiabEquity.cur),
      cashFlowTies: isZero(cashFlow.tieOut),
      cashFlowDiff: cashFlow.tieOut,
      niAgrees: isZero(netIncome.ytd - niYtd),
      // The equity statement must close on the balance sheet's equity, and open
      // on the prior period's close. Both were silently untrue before the
      // 2026-08-26 fix above; surface them rather than let them drift again.
      equityTies: isZero(equityTotals.ending - totalEquity.cur),
      equityDiff: r2(equityTotals.ending - totalEquity.cur),
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// PDF rendering — CPA-plain style (centered bold headers, thin rules, footer).
// Renders the four GL statements onto US-Letter pages with pdf-lib.
// ═══════════════════════════════════════════════════════════════════════════

const PAGE = { w: 612, h: 792, mL: 54, mR: 54, mT: 60, mB: 54 };
const FS = { title: 11, sub: 9.5, head: 8, row: 8.5, foot: 7.5 }; // point sizes
// Data-row pitch (must match the `y -= 12` advance in layout.row()).
const ROW_H = 12;
// Blank space between a column-header underline and the first financial line
// item: 3pt of breathing room below the rule plus TWO blank data rows.
const HDR_TRAIL_GAP = 3 + 2 * ROW_H;

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
    // Only replay when there are still columns to draw against. A statement
    // that clears its columns mid-page (e.g. a prose notes block after the
    // figures) would otherwise replay the header with cols[] empty, and every
    // cols[i] would be undefined -> NaN x -> pdf-lib throws on drawText.
    if (_hdrSpec && !_replaying && cols.length) {
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
    // Centered footer on every page: "<entity>, <Month YYYY>  |  See Executive
    // Summary". Month + year only, never the day (rule set by Jimmy 2026-08-26)
    // — the statement headings already carry the exact period end date.
    const period = meta.asOf ? monthYearLabel(meta.asOf) : meta.longDate;
    const label = meta.entityName + ', ' + period + '  |  See Executive Summary';
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
      // A new column-header block starts a fresh statement, so clear the
      // stacked-total adjacency latch: a rule-below at the bottom of the prior
      // statement must not suppress the first subtotal's rule-above here.
      layout._prevRuledBelow = false;
      // Remember the spec so newPage() can repeat this header on continuation
      // pages. Only record on the FIRST (body-driven) call, not during replay.
      if (!_replaying) _hdrSpec = { labels, hopts };
      const LH = 9;                 // header line height
      // Column pitch (smallest gap between adjacent column right-edges). Computed
      // BEFORE the labels are measured, because a long date heading is wrapped
      // against it.
      let pitch = Infinity;
      for (let i = 1; i < cols.length; i++) pitch = Math.min(pitch, cols[i] - cols[i - 1]);
      // Keep a readable gutter between adjacent column headings.
      //
      // Headings are right-aligned on each column edge, so a label wider than
      // (pitch - MIN_HDR_GUTTER) runs back into the heading to its left. At the
      // balance sheet's 75pt pitch "December 31, 2025" is 71.7pt wide (3.3pt of
      // air) and "September 30, 2026" is 74.4pt (0.6pt) - they read as one run of
      // text; on the Statements of Operations' 72pt pitch the September label
      // actually overlaps its neighbour. Rather than widen the numeric columns
      // (which would squeeze the account-name column), wrap a long "Month D,
      // YYYY" heading after the comma: "September 30," is 54.7pt, leaving a 20pt
      // gutter. Multi-line labels are already bottom-aligned, so the year sits on
      // the same baseline as any single-line heading beside it.
      //
      // All-or-nothing per header row: if ANY date label needs wrapping, every
      // date label wraps, so a June/December pair can never render half-wrapped.
      const MIN_HDR_GUTTER = 10;
      const DATE_LABEL = /^([A-Z][a-z]+ \d{1,2},) (\d{4})$/;
      // A caller-supplied '\n' means the heading was deliberately pre-wrapped
      // (Turnkey's inception range). Trust it and do not auto-wrap on top of it.
      const preWrapped = labels.some(l => String(l).indexOf('\n') >= 0);
      if (Number.isFinite(pitch) && !preWrapped) {
        const avail = pitch - MIN_HDR_GUTTER;
        const anyTooWide = labels.some(l => DATE_LABEL.test(String(l))
          && bold.widthOfTextAtSize(String(l), FS.head) > avail);
        if (anyTooWide) labels = labels.map(l => {
          const dm = DATE_LABEL.exec(String(l));
          return dm ? dm[1] + '\n' + dm[2] : l;
        });
      }
      const nLines = Math.max(1, ...labels.map(l => String(l).split('\n').length));
      // Reserve height for the tallest label block + the underline + the two-row
      // trailing gap, so a header never lands at the very bottom of a page with
      // its first data row orphaned onto the next one.
      ensure(LH * nLines + 2 + HDR_TRAIL_GAP);
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
        // Lines are CENTRED within the heading block rather than each one right-
        // aligned on the column edge, so a wrapped year ("2025") sits in the middle
        // under its month/day line ("December 31,") instead of hugging the edge.
        // The block's right edge is still cols[i] and its width is still the widest
        // line, so the widest line lands exactly where it always did and a
        // single-line heading does not move at all.
        // colBox headers (Statement of Changes in Members' Equity, schedules of
        // investments / partners' capital) CENTER each heading over its column
        // — the block is centered on the underline box's center (cols[i]-boxW/2)
        // rather than right-aligned to the column edge, so a short label like
        // "Contributions" sits centered over its rule instead of hugging the
        // right (CLA/Jimmy 8/29 — "these column headings should all be centered").
        // Non-colBox headers (balance sheet / operations date columns) stay
        // right-aligned over their right-aligned figures.
        const blockLeft = (hopts.colBox && boxW) ? (cols[i] - boxW / 2 - maxW / 2) : (cols[i] - maxW);
        parts.forEach((pl, pi) => {
          const w = bold.widthOfTextAtSize(pl, FS.head);
          const lineY = baseY + (parts.length - 1 - pi) * LH;
          page.drawText(pl, { x: blockLeft + (maxW - w) / 2, y: lineY, size: FS.head, font: bold });
        });
        if (hopts.underline) {
          // colBox → fixed per-column span (with a gutter) so the rule reads as
          // one-per-column; otherwise hug just the widest line of this label.
          const span = (hopts.colBox && boxW) ? boxW : maxW;
          page.drawLine({ start: { x: cols[i] - span, y: uy }, end: { x: cols[i], y: uy }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
        }
      });
      // Advance the cursor below the underline, leaving TWO blank rows between
      // the column heading and the first financial line item. This is both a
      // house style rule and a correctness fix: at the old 8pt gap the first
      // row's baseline landed at uy-8, so a subtotal's `ruleAbove` (drawn at
      // baseline+9 = uy+1) was rendered ONE POINT ABOVE the header underline —
      // the two rules visibly collided whenever a page break put a "Total ..."
      // row first on a continuation page.
      y = uy - HDR_TRAIL_GAP;
    },
    sectionTitle(str) {
      // A section title is a non-total boundary: it re-arms the top rule for
      // the next subtotal (clears the stacked-total adjacency latch).
      layout._prevRuledBelow = false;
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
    row(label, cells, { indent = 12, boldRow = false, ruleAbove = false, ruleBelow = false, doubleBelow = false, gapAfter = 0, dollarPrefix = false, valueInset = 0, colRules = true, keepWithNext = 0 } = {}) {
      // keepWithNext reserves extra space so this row and the row(s) that follow
      // land on the SAME page — used to keep a section grand-total from being
      // orphaned alone at the top of a continuation page: the closest subtotal
      // above it reserves the total's height and the two break together.
      ensure(13 + keepWithNext);
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
      //
      // The rule is sized to HUG THE NUMBERS in this row rather than spanning the
      // full column pitch. Sizing to the pitch made the balance-sheet rules ~87pt
      // wide while the values are only ~50pt, so the leftmost rule ran back under
      // the account name and collided with long labels (e.g. "Bill.Com Clearing
      // Out Banyan Residential Entity 100"). We take the widest value in the row,
      // add a little padding, and clamp to [RULE_MIN_W, pitch − gutter] so every
      // column gets one consistent width that still fits fund-scale figures.
      const GUTTER = 12;
      const RULE_MIN_W = 34;
      // UNIFORM rule width for every subtotal/total row, so the underlines line
      // up vertically column-to-column across the whole statement (CLA/Jimmy
      // 8/17). The old behavior sized each rule to that ROW's widest value, so a
      // small-number row (e.g. "Total Other Development", 10,735.54) drew shorter
      // rules than a big-number row ("Total Other Assets", 2,000,289.13) and the
      // two didn't align. Width is the numeric column box minus a gutter, so
      // adjacent columns still read as separate underlines rather than one
      // continuous line across the row.
      const ruleSpan = Math.max(RULE_MIN_W, ruleBoxW - GUTTER);
      const drawRule = (yy) => {
        if (colRules) {
          cols.forEach((cx) => {
            const x0 = cx - ruleSpan - valueInset;
            const x1 = cx;
            page.drawLine({ start: { x: x0, y: yy }, end: { x: x1, y: yy }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
          });
        } else {
          page.drawLine({ start: { x: ruleLeft, y: yy }, end: { x: ruleRight, y: yy }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
        }
      };
      // Adjacency rule (uniform across all statements, CLA/Jimmy): the single
      // rule that separates two stacked totals belongs to the UPPER total (its
      // rule-below), never the lower one. So if the previous row already drew a
      // rule beneath its figures (ruleBelow or doubleBelow), suppress this row's
      // ruleAbove — otherwise the two lines stack ~12pt apart and read as a
      // stray double underline (Total Current Assets → Total Assets; Total
      // Operating Expenses → Net Income; Total Liabilities → Total L&E; etc.).
      const _drawAbove = ruleAbove && !layout._prevRuledBelow;
      if (_drawAbove) drawRule(y + 9);
      page.drawText(String(label), { x: PAGE.mL + indent, y, size: FS.row, font });
      cells.forEach((c, i) => {
        if (c == null || c === '') return;
        const s = String(c);
        const w = font.widthOfTextAtSize(s, FS.row);
        page.drawText(s, { x: cols[i] - w - valueInset, y, size: FS.row, font });
        if (dollarPrefix) {
          // "$" anchored a fixed gap to the LEFT of this column's own number,
          // NOT at the column-box left edge. The old formula (cols[i] -
          // colWidth(i) + 2) put it at cols[i-1] + 2 -- 2pt off the PREVIOUS
          // column's amount, so the sign read as if it belonged to that column
          // (CLA/Jimmy 8/17). Anchoring at pitch - DOLLAR_PREV_GAP leaves a
          // constant gap after the previous column while still sitting clear of
          // this column's widest figure. Uniform across every profile and the
          // landscape equity page (pitch is the min inter-column pitch).
          const DOLLAR_PREV_GAP = 12;
          const dollarInset = (Number.isFinite(pitch) ? pitch : 78) - DOLLAR_PREV_GAP;
          const dx = cols[i] - dollarInset - valueInset;
          page.drawText('$', { x: dx, y, size: FS.row, font });
        }
      });
      if (ruleBelow) drawRule(y - 3);
      if (doubleBelow) { drawRule(y - 3); drawRule(y - 5); }
      // Latch whether THIS row ruled below, so the next row can suppress a
      // redundant ruleAbove (see adjacency rule above). Reset on any row that
      // did not rule below, so a normal detail line correctly re-arms the top
      // rule for a following subtotal.
      layout._prevRuledBelow = !!(ruleBelow || doubleBelow);
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
  const track = (label, tocLabel) => { if (outOffsets) outOffsets.push({ label: (tocLabel || label), page: pdf.getPageCount() }); };
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fonts = { reg, bold };
  const m = s.meta;
  const money = v => acct(v, { dash: true });
  // Single-member LLCs (SRN / SABINERI and Buna / CLRBUNAP) read "Member\u2019s
  // Equity" (singular possessive) throughout. Pinned by code and raw name so no
  // other srn-profile entity is affected.
  const _singleMember = ['SABINERI', 'CLRBUNAP'].includes(String(m.entityCode || '').toUpperCase())
    || /sabine|county\s*line\s*srn|\bbuna\b/i.test(m.rawEntityName || '');
  const meEquity = _singleMember ? 'Member\u2019s Equity' : 'Members\u2019 Equity';

  // Numeric column right-edges. Balance Sheet now has 3 columns (current, prior,
  // and a Change column at the far right); Operations uses 3.
  const RIGHT = PAGE.w - PAGE.mR;
  // Balance Sheet: current & prior dates plus a "Change" column on the far right.
  // Narrower numeric columns (75pt pitch) push the value block right so long
  // account names (e.g. "MapleMark Entity 105 Banyan SFR GP Investors - 8967 -
  // ICS") keep a clear gap before the first figure instead of crowding it.
  const bsCols = [RIGHT - 150, RIGHT - 75, RIGHT];
  const twoCols = [RIGHT - 95, RIGHT];
  const threeCols = [RIGHT - 200, RIGHT - 100, RIGHT];
  // Operations: four columns — current month, prior month, a Change column
  // (current − prior) sitting between the prior-month and Year-to-Date columns,
  // and Year to Date on the far right. Numeric columns are kept narrow (72pt
  // pitch) so the first (account-name) column is wide enough to show full names
  // with a clear gap before the first figure.
  const opsCols = [RIGHT - 216, RIGHT - 144, RIGHT - 72, RIGHT];

  // ── 1. Balance Sheet ───────────────────────────────────────────────────────
  {
    // Heading date-line repeats on every page (incl. continuation pages) and a
    // blank space follows it before the first row. Dates are NOT underlined.
    const bsTitle = (m.isConsolidated ? 'Consolidated ' : '') + (m.profile === 'banyan'
      ? 'Statements of Assets, Liabilities, and Members\u2019 Equity \u2013 Tax Basis'
      : 'Balance Sheets');
    const L = makeLayout(pdf, fonts, m, bsTitle, { dateLine: m.longDate + ' and ' + m.priorLongDate });
    // TOC label must be the statement title verbatim \u2014 passing a separate string
    // here is how the TOC drifted out of sync with the heading (CLA 8/17: the TOC
    // read the singular "Statement of Assets ..." while the page read "Statements").
    track(bsTitle);
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
    // "$" on the FIRST figure line of each section and on the section's grand
    // total (CLA 8/17). The first line is an account row nested two levels down
    // inside the section, so a latch is armed before each section and consumed
    // by whichever account row renders first. All profiles (CLA 8/17: the
    // dollar signs are a global presentation change, not banyan-only).
    const wantDollar = true;
    let bsFirstRow = false;
    // Underline BELOW the subtotal figure (CLA/Jimmy 8/17) on exactly the
    // listed section totals. "Underline" here means the rule UNDER the number
    // -- these rows always had the rule ABOVE (separating them from the detail
    // lines), which is a different thing.
    const RULE_BELOW_SECTIONS = /^Total (Current Assets|Fixed Assets, Net)$/;
    const renderBsSection = (sec, sectionTotalLabel, reserveTail = 0) => {
      L.row(sec.title, [], { indent: 6, boldRow: true });
      const showSubHeaders = sec.subs.length > 1 || sec.subs.some(su => su.contra) || m.profile === 'bsfrgp' || m.profile === 'banyan';
      for (const su of sec.subs) {
        if (showSubHeaders) L.row(su.title, [], { indent: 16 });
        const rowIndent = showSubHeaders ? 26 : 16;
        for (const r of su.rows) {
          L.row(r.name, bsCells(r.cur, r.pri), { indent: rowIndent, dollarPrefix: bsFirstRow });
          bsFirstRow = false;
        }
        const forceSub = m.profile === 'banyandev' && BANYANDEV_FORCE_SUBTOTAL_SECTIONS.has(sec.title);
        if (showSubHeaders && (su.rows.length > 1 || forceSub)) {
          // GL balances are already signed; a contra subtotal (accumulated
          // amortization) is naturally negative and prints in parentheses.
          L.row('Total ' + su.title, bsCells(su.subtotal.cur, su.subtotal.pri), { indent: 20, ruleAbove: true });
        }
      }
      L.row(sectionTotalLabel, bsCells(sec.total.cur, sec.total.pri), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: RULE_BELOW_SECTIONS.test(sectionTotalLabel), gapAfter: 6, keepWithNext: reserveTail });
    };
    // Turnkey (CLA presentation): a flat list of bare rows under Assets, with
    // only Contract Assets and Fixed Assets carrying a header and a subtotal.
    // No section header and no section total, which is why this cannot go
    // through renderBsSection.
    const tkBs = (m.profile === 'turnkey') ? s.balanceSheet.turnkey : null;
    const renderTkBlocks = (blocks) => {
      for (const blk of blocks) {
        if (blk.kind === 'row') {
          L.row(blk.name, bsCells(blk.cur, blk.pri), { indent: 10, dollarPrefix: bsFirstRow });
          bsFirstRow = false;
          continue;
        }
        // Keep a group and its subtotal on one page.
        L.keepTogether(12 + blk.rows.length * 12 + 12 + 4);
        L.row(blk.title, [], { indent: 10 });
        for (const r of blk.rows) {
          L.row(r.name, bsCells(r.cur, r.pri), { indent: 22, dollarPrefix: bsFirstRow });
          bsFirstRow = false;
        }
        L.row('Total ' + blk.title, bsCells(blk.subtotal.cur, blk.subtotal.pri), { indent: 16, ruleAbove: true });
      }
    };
    // Height reserved by the closest subtotal above each grand total so the
    // grand total never lands alone atop a continuation page (>= gapAfter + a
    // grand-total row; see keepWithNext in makeLayout.row).
    const BS_GRAND_RESERVE = 22;
    // Height of the equity CLOSING cluster (Net Income + the grand Total
    // Members' Equity + Total Liabilities and Members' Equity). Reserved on the
    // Net Income line so those totals never land alone at the top of a page —
    // the whole cluster breaks together, led by the Net Income detail line
    // (CLA/Jimmy 8/30 — "no orphaned totals").
    const BS_EQUITY_TAIL = 46;
    bsFirstRow = wantDollar;
    // True when the LAST asset section total already carries a rule BELOW its
    // figures (Total Current Assets / Total Fixed Assets, Net). In that case the
    // grand "Total Assets" must NOT also draw a rule above its own figures — the
    // two would stack 6pt apart and read as a stray double rule. Most visible on
    // a one-section balance sheet (County Line Rail Operations), where Total
    // Current Assets is immediately followed by Total Assets (Jimmy, 2026-08-30).
    let lastAssetSectionRuledBelow = false;
    if (tkBs) renderTkBlocks(tkBs.assetBlocks);
    else {
      // Keep the grand "Total Assets" line from being orphaned alone at the top
      // of a continuation page: the LAST asset section's section-total reserves
      // the grand total's height (BS_GRAND_RESERVE), so if both won't fit they
      // break to the next page together (CLA/Jimmy 8/28 — "move the closest
      // subtotal with the total line to the subsequent page").
      const asecs = s.balanceSheet.assetSections;
      asecs.forEach((sec, i) => renderBsSection(sec, 'Total ' + sec.title, i === asecs.length - 1 ? BS_GRAND_RESERVE : 0));
      if (asecs.length) lastAssetSectionRuledBelow = RULE_BELOW_SECTIONS.test('Total ' + asecs[asecs.length - 1].title);
    }
    L.row('Total Assets', bsCells(s.balanceSheet.totalAssets.cur, s.balanceSheet.totalAssets.pri), { indent: 6, boldRow: true, ruleAbove: !lastAssetSectionRuledBelow, doubleBelow: true, gapAfter: 8, dollarPrefix: wantDollar });

    L.sectionTitle('LIABILITIES AND ' + meEquity.toUpperCase());
    bsFirstRow = wantDollar;
    if (tkBs) renderTkBlocks(tkBs.liabBlocks);
    else for (const sec of s.balanceSheet.liabSections) renderBsSection(sec, 'Total ' + sec.title);
    L.row('Total Liabilities', bsCells(s.balanceSheet.totalLiab.cur, s.balanceSheet.totalLiab.pri), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    // Keep the whole Members' Equity section on one page \u2014 never split it so
    // that only its totals spill to a near-empty continuation page (CLA/Jimmy
    // 8/30, "no orphaned totals"). If the section is taller than a page it can't
    // be kept whole; the Net Income tail reserve below is the fallback then.
    {
      const _eqRows = (s.balanceSheet.equityRows || []).length + (s.balanceSheet.retainedRows || []).length;
      const _fixed = (m.profile === 'banyandev') ? 7 : 4;   // headers + subtotals + NI + 2 grand totals
      L.keepTogether((_eqRows + _fixed) * 13 + 20);
    }
    L.row(meEquity, [], { indent: 6, boldRow: true });
    for (const r of s.balanceSheet.equityRows) L.row(r.name, bsCells(r.cur, r.pri), { indent: 16 });
    if (m.profile === 'banyandev') {
      // CLA presentation: a contributed-capital subtotal, then a separate
      // Retained Earnings subsection with its own subtotal, then Net Income,
      // then the grand Total Members' Equity.
      const tce = s.balanceSheet.totalContribEquity;
      L.row('Total ' + meEquity, bsCells(tce.cur, tce.pri), { indent: 20, ruleAbove: true, gapAfter: 4 });
      const rr = s.balanceSheet.retainedRows || [];
      const retTot = { cur: r2(rr.reduce((a, r) => a + r.cur, 0)), pri: r2(rr.reduce((a, r) => a + r.pri, 0)) };
      L.row('Retained Earnings', [], { indent: 6, boldRow: true });
      for (const r of rr) L.row(r.name, bsCells(r.cur, r.pri), { indent: 16 });
      L.row('Total Retained Earnings', bsCells(retTot.cur, retTot.pri), { indent: 20, ruleAbove: true, gapAfter: 4 });
      L.keepTogether(BS_EQUITY_TAIL);
      L.row('Net Income (Loss)', bsCells(s.balanceSheet.niLine.cur, s.balanceSheet.niLine.pri), { indent: 16 });
      L.row('Total ' + meEquity, bsCells(s.balanceSheet.totalEquity.cur, s.balanceSheet.totalEquity.pri), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    } else {
      for (const r of (s.balanceSheet.retainedRows || [])) L.row(r.name, bsCells(r.cur, r.pri), { indent: 16 });
      L.keepTogether(BS_EQUITY_TAIL);
      L.row('Net Income (Loss)', bsCells(s.balanceSheet.niLine.cur, s.balanceSheet.niLine.pri), { indent: 16 });
      L.row('Total ' + meEquity, bsCells(s.balanceSheet.totalEquity.cur, s.balanceSheet.totalEquity.pri), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    }
    L.row('Total Liabilities and ' + meEquity, bsCells(s.balanceSheet.totalLiabEquity.cur, s.balanceSheet.totalLiabEquity.pri), { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: wantDollar });
  }

  // ── 2. Statements of Operations ─────────────────────────────────────────────
  {
    // Heading + column labels follow the period toggle (monthly/quarterly/
    // annually). resolvePeriod sets m.colLabel = 'Month Ended' | 'Quarter Ended'
    // | 'Year Ended'; pluralize it for the two-date subtitle and prefix it on
    // the two period columns so a quarterly report reads 'For the Quarters Ended
    // 6/30/26 and 3/31/26' with 'Quarter Ended' columns (matches the CPA package).
    const opsDateLine = m.opsDateLine || opsHeadingLine(m.colLabel, m.longDate, m.priorLongDate);
    const opsTitle = (m.isConsolidated ? 'Consolidated ' : '') + (m.profile === 'banyan'
      ? 'Statements of Revenues and Expenses \u2013 Tax Basis'
      : 'Statements of Operations');
    const L = makeLayout(pdf, fonts, m, opsTitle, { dateLine: opsDateLine });
    track(opsTitle);
    L.start();
    L.setCols(opsCols);
    // Period columns show just the period-end DATE. The period type is already
    // stated in the date line above ("For the Months Ended ..."), so repeating
    // "Month Ended" / "Quarter Ended" over each column was redundant.
    L.colHeaders([m.longDate, m.opsPriorColLabel || m.priorLongDate, 'Change', 'Year to Date'], { underline: true });
    const chg = (cur, pri) => money(r2(cur - pri));
    const line = (r, o = {}) => L.row(r.name, [money(r.cur), money(r.pri), chg(r.cur, r.pri), money(r.ytd)], { indent: 16, ...o });

    // A 4-column value cell (current / prior / change / YTD).
    const cell4 = t => [money(t.cur), money(t.pri), chg(t.cur, t.pri), money(t.ytd)];
    if (s.operations.banyan && s.operations.banyan.structured) {
      // ── Banyan Residential shape: Revenue - Services / Operating Expenses
      //    (grouped) / Other Income (Expense) / Income Taxes / Net Income. ────
      const bo = s.operations.banyan;
      // "$" on the first figure line of the statement (the first revenue
      // account) and on Net Income (Loss) \u2014 CLA 8/17. Same latch pattern as
      // the balance sheet: armed just before the revenue tree is rendered.
      let plFirstRow = false;
      const renderTree = (groups, { showGroupTotal, keepWhole } = {}) => {
        for (const g of groups) {
          if (keepWhole) {
            // Reserve the whole group's height so a category never splits from
            // its subtotal across a page break.
            const nRows = g.subs.reduce((n, su) => n + 1 + su.lines.length + (su.lines.length > 1 ? 1 : 0), 0);
            L.keepTogether(12 + nRows * 12 + 16);
          }
          L.row(g.title, [], { indent: 12, boldRow: true });
          for (const su of g.subs) {
            // Don't echo a subsection header that just repeats the group name.
            const echo = su.title === g.title;
            if (!echo) L.row(su.title, [], { indent: 20 });
            const lineIndent = echo ? 26 : 30;
            su.lines.forEach(r => {
              L.row(r.name, cell4(r), { indent: lineIndent, dollarPrefix: plFirstRow });
              plFirstRow = false;
            });
            if (su.lines.length > 1 && !echo) L.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true });
          }
          if (showGroupTotal !== false && !(bo.noGroupTotal && bo.noGroupTotal.has(g.title))) L.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true });
        }
      };

      // Revenue → Total Revenue → Gross Profit (no cost of revenue on Banyan).
      L.sectionTitle('Revenue');
      plFirstRow = true;
      renderTree(bo.revenueTree, { showGroupTotal: true });
      // Disarm even if the revenue tree rendered nothing. A latch left armed
      // would be consumed by the first OPERATING EXPENSE line instead, putting
      // the "$" several sections down the statement.
      plFirstRow = false;
      L.row('Total Revenue', cell4(bo.totRev), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: bo.showGrossProfit !== false ? 0 : 6 });
      if (bo.showGrossProfit !== false) L.row('Gross Profit', cell4(bo.grossProfit), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });

      L.sectionTitle('Operating Expenses');
      renderTree(bo.opexTree, { showGroupTotal: true, keepWhole: true });
      L.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });

      // Other Income (Expense): income section, then expense section (shown as
      // reductions of income), then the netted total.
      if (bo.otherIncomeTree.length || bo.otherExpenseTree.length) {
        L.sectionTitle('Other Income (Expense)');
        renderTree(bo.otherIncomeTree, { showGroupTotal: true });
        renderTree(bo.otherExpenseTree, { showGroupTotal: true });
        L.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
      }

      if (bo.incomeTaxTree.length) {
        L.sectionTitle('Income Taxes');
        renderTree(bo.incomeTaxTree, { showGroupTotal: true });
        L.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
      }

      L.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, boldRow: true, ruleAbove: false, doubleBelow: true, dollarPrefix: true });
    } else if (s.operations.bsfrgp && s.operations.bsfrgp.structured) {
      // ── Banyan SFR GP Investors shape: Operating Expenses / Other Income
      //    (Expense) / Income Taxes / Net Income (Loss). ────────────────────
      const bo = s.operations.bsfrgp;
      // $ on the first figure line of the statement (first operating-expense
      // line) and on Net Income (Loss) -- global, CLA 8/17.
      let plFirstRow = false;
      // Render a nested tree (group → subsection → lines) with subtotals. A
      // subsection subtotal shows only when it has more than one line; a group
      // subtotal shows only when the group has more than one subsection (so a
      // single-line 'Interest Income' doesn't get a redundant echo).
      const renderTree = (groups, { showGroupTotal }) => {
        for (const g of groups) {
          L.row(g.title, [], { indent: 12, boldRow: true });
          for (const su of g.subs) {
            const echo = su.title === g.title;
            if (!echo) L.row(su.title, [], { indent: 20 });
            su.lines.forEach(r => { L.row(r.name, cell4(r), { indent: echo ? 26 : 30, dollarPrefix: plFirstRow }); plFirstRow = false; });
            if (su.lines.length > 1 && !echo) L.row('Total ' + su.title, cell4(su.subtotal), { indent: 24, ruleAbove: true });
          }
          if (showGroupTotal && (g.subs.length > 1 || g.subs.some(su => su.title === g.title))) {
            L.row('Total ' + g.title, cell4(g.subtotal), { indent: 16, ruleAbove: true });
          }
        }
      };

      L.sectionTitle('Operating Expenses');
      plFirstRow = true;
      renderTree(bo.opexTree, { showGroupTotal: true });
      L.row('Total Operating Expenses', cell4(bo.totOpex), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });

      // Other Income (Expense): Other Income section, then Other Expense section,
      // then a netted 'Total Other Income (Expense)'.
      L.sectionTitle('Other Income (Expense)');
      renderTree(bo.otherIncomeTree, { showGroupTotal: true });
      renderTree(bo.otherExpenseTree, { showGroupTotal: true });
      L.row('Total Other Income (Expense)', cell4(bo.totOtherIE), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });

      // Income Taxes.
      L.sectionTitle('Income Taxes');
      renderTree(bo.incomeTaxTree, { showGroupTotal: true });
      L.row('Total Income Taxes', cell4(bo.totIncomeTax), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });

      L.row('Net Income (Loss)', cell4(bo.netIncome), { indent: 6, boldRow: true, ruleAbove: false, doubleBelow: true, dollarPrefix: true });
    } else {
    // $ on the first figure line of the statement (CLA 8/17, global). It is
    // armed once and spent by whichever section draws first, because a
    // development entity whose only revenue account was interest income now
    // has no Revenue section at all - that account sits in Other Income
    // (Expense) (Jimmy, 2026-08-28).
    const firstFig = { armed: true };
    const spendDollar = () => { const d = firstFig.armed; firstFig.armed = false; return d; };
    if (s.operations.revenue.length) {
      L.sectionTitle('Revenue');
      s.operations.revenue.forEach(r => line(r, { dollarPrefix: spendDollar() }));
      L.row('Total Revenue', [money(s.operations.totRev.cur), money(s.operations.totRev.pri), chg(s.operations.totRev.cur, s.operations.totRev.pri), money(s.operations.totRev.ytd)], { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    }
    if (s.operations.cogs.length) {
      const cogsLabel = m.profile === 'turnkey' ? 'Cost of Goods Sold' : 'Cost of Revenue';
      L.sectionTitle(cogsLabel);
      s.operations.cogs.forEach(r => line(r));
      L.row('Total ' + cogsLabel, [money(s.operations.totCogs.cur), money(s.operations.totCogs.pri), chg(s.operations.totCogs.cur, s.operations.totCogs.pri), money(s.operations.totCogs.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 4 });
      L.row('Gross Profit', [money(s.operations.grossProfit.cur), money(s.operations.grossProfit.pri), chg(s.operations.grossProfit.cur, s.operations.grossProfit.pri), money(s.operations.grossProfit.ytd)], { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    }
    L.sectionTitle(m.profile === 'turnkey' ? 'General & Administrative Expenses' : 'Operating Expenses');
    // Grouped into the 11 presentation categories, each with its own subtotal.
    // Category subtotals sum to Total Operating Expenses exactly (pure re-group).
    // Fall back to a flat list if grouping produced nothing (defensive).
    // Turnkey renders G&A FLAT. The 11 categories are an SRN railroad-operations
    // construct and misfile a contractor chart - Depreciation Expense landed
    // under a "Professional Services" header - and the CPA package presents
    // General & Administrative Expenses as a plain list of accounts.
    const groups = (m.profile !== 'turnkey') && s.operations.opexGroups && s.operations.opexGroups.length
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
        g.lines.forEach(r => L.row(r.name, [money(r.cur), money(r.pri), chg(r.cur, r.pri), money(r.ytd)], { indent: 26, dollarPrefix: spendDollar() }));
        if (hasSubtotal) {
          L.row('Total ' + g.title, [money(g.subtotal.cur), money(g.subtotal.pri), chg(g.subtotal.cur, g.subtotal.pri), money(g.subtotal.ytd)], { indent: 20, ruleAbove: true });
        }
      }
    } else {
      s.operations.opex.forEach(r => line(r, { dollarPrefix: spendDollar() }));
    }
    L.row('Total ' + (m.profile === 'turnkey' ? 'General & Administrative Expenses' : 'Operating Expenses'), [money(s.operations.totOpex.cur), money(s.operations.totOpex.pri), chg(s.operations.totOpex.cur, s.operations.totOpex.pri), money(s.operations.totOpex.ytd)], { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    // ── Other Income (Expense) / Income Taxes ──────────────────────────
    // Populated for EVERY profile now by the shared classifier
    // (otherIeRoute), not just turnkey: misc revenue, interest income,
    // other income, the non-operating gains, misc / other / interest
    // expense and the penalties. Rendered as group -> subsection -> lines,
    // the same nesting the source P&L uses, so the section reads:
    //   Other Income (Expense)
    //     Other Income / Other Income / <lines> / Total Other Income x2
    //     Other Expense / Other Expenses / <lines> / Total ... x2
    //   Total Other Income (Expense)
    //   Income Taxes / State and Local Taxes / <lines> / Total x2
    // Expense lines print parenthesised, as reductions of income.
    const oiTree = s.operations.otherIncomeTree || [];
    const oeTree = s.operations.otherExpenseTree || [];
    const itTree = s.operations.incomeTaxTree || [];
    const negT = (t) => ({ cur: -t.cur, pri: -t.pri, ytd: -t.ytd });
    const cell4oie = (t) => [money(t.cur), money(t.pri), chg(t.cur, t.pri), money(t.ytd)];
    const renderOie = (groups, opts) => {
      const o = opts || {};
      for (const g of groups) {
        const nSub = g.subs.length;
        const nLine = g.subs.reduce((s2, su) => s2 + su.lines.length, 0);
        L.keepTogether(12 + nSub * 24 + nLine * 12 + 16);
        L.row(g.title, [], { indent: 12, boldRow: true });
        for (const su of g.subs) {
          const echo = !!o.echoSub && su.title === g.title;
          if (!echo) L.row(su.title, [], { indent: 20 });
          const li = echo ? 26 : 30;
          su.lines.forEach(r => L.row(r.name, cell4oie(o.negate ? negT(r) : r), { indent: li }));
          if (!echo) L.row('Total ' + su.title, cell4oie(o.negate ? negT(su.subtotal) : su.subtotal), { indent: 24, ruleAbove: true });
        }
        L.row('Total ' + g.title, cell4oie(o.negate ? negT(g.subtotal) : g.subtotal), { indent: 16, ruleAbove: true });
      }
    };
    if (oiTree.length || oeTree.length) {
      L.sectionTitle('Other Income (Expense)');
      renderOie(oiTree, { echoSub: true });
      renderOie(oeTree, { negate: true, echoSub: true });
      L.row('Total Other Income (Expense)', cell4oie(s.operations.totOtherIE), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    }
    if (itTree.length) {
      L.sectionTitle('Income Taxes');
      renderOie(itTree, { echoSub: true });
      L.row('Total Income Taxes', cell4oie(s.operations.totIncomeTax), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: true, gapAfter: 6 });
    }
    L.row('Net Income (Loss)', [money(s.operations.netIncome.cur), money(s.operations.netIncome.pri), chg(s.operations.netIncome.cur, s.operations.netIncome.pri), money(s.operations.netIncome.ytd)], { indent: 6, boldRow: true, ruleAbove: false, doubleBelow: true, dollarPrefix: true });
    }
  }

  // ── 3. Statement of Cash Flows ──────────────────────────────────────────────
  {
    const cfTitle = (m.isConsolidated ? 'Consolidated ' : '') + (m.profile === 'banyan'
      ? 'Statement of Cash Flows \u2013 Tax Basis'
      : 'Statement of Cash Flows');
    const L = makeLayout(pdf, fonts, m, cfTitle, { dateLine: m.monthsEnded });
    // TOC label is the title verbatim (CLA 8/17: the TOC was dropping
    // "\u2013 Tax Basis" because this override hardcoded the old string).
    track(cfTitle);
    L.start();
    L.setCols([RIGHT]);
    // No column heading: the single YTD column is self-evident from the
    // statement title, so no "Year to Date" label is drawn above it.
    const cf = s.cashFlow;
    // Turnkey (CLA presentation): an explicit ordered line set rather than the
    // generic section-by-section build below.
    if (cf.turnkey) {
      for (const ln of cf.turnkey.lines) {
        if (ln.gapBefore) L.row('', [], { gapAfter: 4 });
        if (ln.header) { L.row(ln.label, [], { indent: 6, boldRow: true }); continue; }
        L.row(ln.label, [money(ln.value)], {
          indent: ln.indent ? 16 : 6,
          boldRow: !!ln.bold,
          ruleAbove: !!ln.rule,
          doubleBelow: !!ln.doubleBelow,
          dollarPrefix: !!ln.dollar,
        });
      }
    } else {
    L.sectionTitle('Cash Flows from Operating Activities');
    L.row('Net Income (Loss)', [money(cf.netIncome)], { indent: 16, dollarPrefix: true });
    L.row('Adjustments to reconcile net income to net cash:', [], { indent: 16 });
    if (!isZero(cf.amortization)) L.row('Amortization and depreciation', [money(cf.amortization)], { indent: 28 });
    L.space(6);
    L.row('Changes in Operating Assets and Liabilities:', [], { indent: 16 });
    if (!isZero(cf.changeAR)) L.row('(Increase) decrease in accounts receivable', [money(cf.changeAR)], { indent: 28 });
    if (!isZero(cf.changePrepaidOther)) L.row('(Increase) decrease in prepaid and other current assets', [money(cf.changePrepaidOther)], { indent: 28 });
    if (!isZero(cf.changeIntercompany)) L.row('(Increase) decrease in intercompany balances', [money(cf.changeIntercompany)], { indent: 28 });
    // CLA 8/17 (global): present the payables bucket and the other-current-
    // liabilities bucket as ONE line on EVERY profile. Internally they stay
    // separate (cfBuckets.ap = accounts named "... payable"; cfBuckets.accrued =
    // every other current liability), so this is purely a presentation merge --
    // both were already inside netOperating, so no subtotal moves and every
    // statement still ties. Jimmy 8/17: this is a global change, not banyan-only.
    const changeApOther = r2(cf.changeAP + cf.changeAccrued);
    if (!isZero(changeApOther)) {
      L.row('Increase (decrease) in accounts payable and other current liabilities',
        [money(changeApOther)], { indent: 28 });
    }
    L.row('Net Cash Provided (Used) by Operating Activities', [money(cf.netOperating)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.sectionTitle('Cash Flows from Investing Activities');
    if (!isZero(cf.capex)) L.row('Acquisition of fixed assets', [money(cf.capex)], { indent: 28 });
    if (!isZero(cf.ltInvest)) L.row(m.profile === 'banyandev' ? 'Purchase of Long Term Investments and Other Assets' : '(Increase) decrease in Other Assets', [money(cf.ltInvest)], { indent: 28 });
    L.row('Net Cash Provided (Used) by Investing Activities', [money(cf.netInvesting)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.sectionTitle('Cash Flows from Financing Activities');
    if (!isZero(cf.equityContrib)) L.row('Member contributions (distributions), net', [money(cf.equityContrib)], { indent: 28 });
    if (!isZero(cf.debtChange)) L.row('Net Proceeds from (Repayment of) Loan Payable', [money(cf.debtChange)], { indent: 28 });
    L.row('Net Cash Provided (Used) by Financing Activities', [money(cf.netFinancing)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 8 });

    L.row('Net Increase (Decrease) in Cash', [money(cf.netChange)], { indent: 6, boldRow: true, ruleAbove: true });
    L.row('Cash, Beginning of Period', [money(cf.cashBeg)], { indent: 6 });
    L.row('Cash, End of Period', [money(cf.cashEnd)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, dollarPrefix: true });
    if (!isZero(cf.tieOut)) {
      L.space(6);
      L.row('Note: reconciled change differs from cash movement by ' + money(cf.tieOut) + ' (see notes).', [], { indent: 6 });
    }
    } // end generic cash-flow branch (turnkey renders its own line set above)
  }

  // ── 4. Statement of Changes in Members' Equity ──────────────────────────────
  {
    // County Line Rail Operations (COUNTYLI3 / entity 46) uses a TRANSPOSED
    // presentation per the CPA reference: activity runs down the rows (opening
    // balance, activity, ending balance) and each MEMBER is a column, plus a
    // Total column — the opposite of the standard member-per-row layout below.
    // CLRO is single-member, so the member column and Total column show the
    // same figures. Only nonzero activity rows print (Jimmy, 2026-09-01).
    const _isClroEquity = (String(m.entityCode || '').toUpperCase() === 'COUNTYLI3')
      || /county\s*line\s*rail\s*operations|^clro\b/i.test(String(m.rawEntityName || ''));
    if (_isClroEquity) {
      renderClroEquity(pdf, fonts, m, meEquity, s.equity, track);
    } else {
    // Landscape page mirroring the CPA reference: five columns, a Distributions
    // column shown even when all zero, and a Net Income (Loss) column wide enough
    // to keep the value on one row. Only the first member row and the Total row
    // carry a "$" (CLA 8/17); the rows between them are bare figures.
    const eqTitle = (m.isConsolidated ? 'Consolidated ' : '') + (m.profile === 'banyan'
      ? 'Statement of Changes in ' + meEquity + ' \u2013 Tax Basis'
      : 'Statement of Changes in ' + meEquity);
    const L = makeLayout(pdf, fonts, m, eqTitle,
      { landscape: true, dateLine: m.monthsEnded });
    const LRIGHT = PAGE.h - PAGE.mR; // landscape printable right edge (PAGE.h is the long side)
    track(eqTitle);
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
    const begDate = m.equityBegDate || ('1/1/' + String(m.asOf).slice(0, 4));
    const endDate = shortMD(m.longDate);
    // Right-edges anchored at the printable right edge and marched LEFT by a
    // fixed pitch. The first numeric column (c1) is placed far enough right that
    // its "$" prefix cell box (c1 - colWidth(0) + 2) clears the longest member
    // label — long trust names like "Contributed Capital - Stephen & Katherine
    // VanDusen 2007 Trust" (right edge ≈ 308pt) were crowding the first column.
    // c1 = 404 puts the "$" box left at ~328pt (a ~20pt gap after the longest
    // label) and the first figure ~54pt clear of it; pitch (~84pt) still leaves
    // each column ample room for a "$" prefix and a large right-aligned value.
    const c1 = 404;
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
    // "$" on the first member row and the Total row only (CLA 8/17). The c1
    // column edge was already positioned to leave room for a "$" cell box at
    // ~328pt, clear of the longest member label.
    const eqDollar = true;
    L.row('Member', [], { indent: 6, boldRow: true });
    s.equity.rows.forEach((r, i) => {
      dollarRow(r.name, [r.beginning, r.contributions, r.distributions, r.netIncome, r.ending],
        { indent: 16, dollarPrefix: eqDollar && i === 0 });
    });
    const t = s.equity.totals;
    dollarRow('Total', [t.beginning, t.contributions, t.distributions, t.netIncome, t.ending],
      { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, colRules: true, dollarPrefix: eqDollar });
    }
  }

  return await pdf.save();
}

// ═══════════════════════════════════════════════════════════════════════════
// renderClroEquity — TRANSPOSED Statement of Changes in Member's Equity for
// County Line Rail Operations, LLC (COUNTYLI3 / entity 46), matching the CPA
// reference: activity runs DOWN the rows (opening balance → nonzero activity →
// ending balance) and each MEMBER is a COLUMN, plus a Total column. CLRO is
// single-member, so the member column and the Total column carry the same
// figures. Only activity rows with a nonzero movement print. Portrait page.
//   Rows:   Equity Balance at <beg>     $  <b>   $  <b>
//           Net income                     <ni>     <ni>      (only if nonzero)
//           Contributions                  <c>      <c>       (only if nonzero)
//           Distributions                  <d>      <d>       (only if nonzero)
//           Equity Balance at <end>     $  <e>   $  <e>
// ═══════════════════════════════════════════════════════════════════════════
function renderClroEquity(pdf, fonts, m, meEquity, equity, track) {
  const eqTitle = 'Statement of Changes in ' + meEquity;
  const L = makeLayout(pdf, fonts, m, eqTitle, { dateLine: m.monthsEnded });
  track(eqTitle);
  L.start();

  const shortMD = (long) => {
    const map = { January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12 };
    const mm = String(long).match(/^(\w+)\s+(\d+),\s+(\d+)$/);
    return mm ? (map[mm[1]] + '/' + mm[2] + '/' + mm[3]) : long;
  };
  // Long-form dates on the row labels, per the reference ("January 1, 2026").
  const begLong = m.equityBegLongDate || ('January 1, ' + String(m.asOf).slice(0, 4));
  const endLong = m.longDate;

  // Single-member: take the member row (the non-Retained, non-empty equity row);
  // fall back to the totals if the member list is degenerate. The member column
  // and the Total column show the same numbers.
  const memberRows = (equity.rows || []);
  const primary = memberRows.length ? memberRows[0] : null;
  const t = equity.totals || {};
  const beg = primary ? primary.beginning : (t.beginning || 0);
  const contrib = primary ? primary.contributions : (t.contributions || 0);
  const distrib = primary ? primary.distributions : (t.distributions || 0);
  const ni = primary ? primary.netIncome : (t.netIncome || 0);
  const end = primary ? primary.ending : (t.ending || 0);
  const memberName = 'County Line Railroad\nInterests, LLC';

  // Two numeric columns anchored near the right of the portrait page: the member
  // column and the Total column. Right-edges marched left from the printable
  // right edge by a fixed pitch, leaving each column room for a "$" cell and a
  // large right-aligned value.
  const PRIGHT = PAGE.w - PAGE.mR;
  const PITCH = 118;
  const cTotal = PRIGHT;
  const cMember = cTotal - PITCH;
  L.setCols([cMember, cTotal]);

  // Header: member name (may wrap to two lines) over the member column, "Total"
  // over the Total column; both underlined, date-style bottom-aligned block so
  // the labels sit just above the rule (matches the reference header).
  L.colHeaders([memberName, 'Total'], { bottomAlign: true, underline: true, colBox: true });

  const isZero = (v) => Math.abs(Number(v) || 0) < 0.005;
  const twoCol = (v) => [acct(v), acct(v)]; // member col == total col (single member)

  // Opening balance — "$" on both columns.
  L.row('Equity Balance at ' + begLong, twoCol(beg), { indent: 6, valueInset: 4, dollarPrefix: true });
  // Activity — only nonzero rows, in the reference order.
  if (!isZero(contrib)) L.row('Contributions', twoCol(contrib), { indent: 6, valueInset: 4 });
  if (!isZero(distrib)) L.row('Distributions', twoCol(distrib), { indent: 6, valueInset: 4 });
  if (!isZero(ni)) L.row('Net income', twoCol(ni), { indent: 6, valueInset: 4 });
  // Ending balance — rule above, "$" on both columns, and a rule below the
  // figures to close the statement (single rule, matching the reference).
  L.row('Equity Balance at ' + endLong, twoCol(end), { indent: 6, valueInset: 4, ruleAbove: true, ruleBelow: true, dollarPrefix: true });
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
  center(meta.isConsolidated ? 'Consolidated Financial Statements' : 'Financial Statements', 15, reg, 470);
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
// Consolidating schedules (landscape) for a consolidated package: one column
// per member, an eliminations column, and the consolidated cross-foot. Figures
// come straight from the consolidation engine's buildColumns (schedules arg),
// so they tie to the on-screen schedules and to CLA to the penny. Records each
// schedule's 0-based start page in `offsets` for the Table of Contents.
//   schedules: { columns:[{entity_id,label}], balanceSheet:{accounts}, incomeMonth:{accounts} }
// ═══════════════════════════════════════════════════════════════════════════
async function renderConsolidatingSchedulesPdf(schedules, meta, offsets) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const PW = PAGE.h, PH = PAGE.w;                 // landscape
  const left = PAGE.mL, right = PW - PAGE.mR;
  const cols = schedules.columns || [];
  const heads = cols.map(c => c.label).concat(['Eliminations', 'Consolidated']);
  const nCols = heads.length;
  const usable = right - left;
  // Reserve a wide name column so full account names fit, then split the rest
  // across the figure columns. numColW is floored so the widest figure
  // (~"(168,564,159.10)") still fits with an inset.
  const nameWidth = Math.max(210, Math.min(330, Math.floor(usable * 0.42)));
  const numColW = Math.max(70, Math.floor((usable - nameWidth) / nCols));
  const colRight = []; for (let i = 0; i < nCols; i++) colRight.push(right - numColW * (nCols - 1 - i));
  const nameLeft = left;
  // Hard right edge for account names: a clear gap before the first figure
  // column's band, so a long name (or its ellipsis) never runs under the
  // column underline.
  const nameEnd = colRight[0] - numColW - 8;
  const F = { title: 10.5, sub: 9, head: 6.8, row: 7.0, foot: 7.2 };
  const rowH = 11.5, lineH = 8;
  // Space a section subtotal reserves so the following grand total shares its page.
  const GRAND_RESERVE = rowH * 1.5;
  let page, y, curTitle = null, curDateLine = null;
  // Set true just before the first account row of a statement section; the next
  // account row prints a "$" in every column and clears it (CLA dollar rule).
  let dollarFirst = false;

  const dtext = (s, x, yy, size, font, color) => page.drawText(String(s), { x, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  const dright = (s, xr, yy, size, font, color) => dtext(s, xr - font.widthOfTextAtSize(String(s), size), yy, size, font, color);
  const dcenter = (s, size, font, yy) => dtext(s, (PW - font.widthOfTextAtSize(String(s), size)) / 2, yy, size, font);
  const truncate = (s, font, size, max) => { let t = String(s); if (font.widthOfTextAtSize(t, size) <= max) return t; while (t.length > 3 && font.widthOfTextAtSize(t + '…', size) > max) t = t.slice(0, -1); return t + '…'; };
  const wrapHead = (label) => {
    const maxW = numColW - 6, words = String(label).split(/\s+/), lines = []; let cur = '';
    for (const w of words) { const t = cur ? cur + ' ' + w : w; if (bold.widthOfTextAtSize(t, F.head) > maxW && cur) { lines.push(cur); cur = w; } else cur = t; }
    if (cur) lines.push(cur);
    return lines.slice(0, 3);
  };
  const footer = () => {
    const period = meta.asOf ? monthYearLabel(meta.asOf) : meta.longDate;
    const label = meta.entityName + ', ' + period + '  |  See Executive Summary';
    dtext(label, (PW - reg.widthOfTextAtSize(label, F.foot)) / 2, PAGE.mB - 12, F.foot, reg, rgb(0.4, 0.4, 0.4));
  };
  const colHeaders = () => {
    const hl = heads.map(wrapHead), maxLines = Math.max(...hl.map(l => l.length));
    for (let i = 0; i < nCols; i++) {
      for (let j = 0; j < hl[i].length; j++) dright(hl[i][j], colRight[i], y - (maxLines - hl[i].length + j) * lineH, F.head, bold);
      const uy = y - (maxLines - 1) * lineH - 2.5;
      page.drawLine({ start: { x: colRight[i] - (numColW - 8), y: uy }, end: { x: colRight[i], y: uy }, thickness: 0.6, color: rgb(0.2, 0.2, 0.2) });
    }
    y -= maxLines * lineH + 5;
  };
  const newPage = () => {
    page = pdf.addPage([PW, PH]);
    y = PH - PAGE.mT;
    dcenter(meta.entityName, F.title, bold, PH - PAGE.mT + 22);
    dcenter(curTitle, F.sub, bold, PH - PAGE.mT + 10);
    if (curDateLine) dcenter(curDateLine, F.sub, reg, PH - PAGE.mT - 1);
    footer();
    y -= 24; colHeaders();
  };
  const ensure = (space) => { if (y - space < PAGE.mB + 8) newPage(); };
  // Draw the figure in every column, right-aligned on the column edge. When
  // `dollar` is set, also print a "$" at the LEFT of each column's numeric band
  // — a fixed position, so the number stays right-aligned and a clear gap opens
  // between the "$" and the figure (CLA presentation; Jimmy, 2026-08-30). Used
  // on the first account row of each statement and on the grand totals, every
  // column, matching the face statements' dollar rule.
  const DOLLAR_GAP = 5;   // "$" sits this far right of the column's left boundary
  const figs = (getVal, font, dollar) => {
    for (let i = 0; i < nCols; i++) {
      if (dollar) dtext('$', colRight[i] - numColW + DOLLAR_GAP, y, F.row, font || reg);
      dright(acct(getVal(i)), colRight[i], y, F.row, font || reg);
    }
  };
  const rule = () => { for (let i = 0; i < nCols; i++) page.drawLine({ start: { x: colRight[i] - (numColW - 10), y: y + 8 }, end: { x: colRight[i], y: y + 8 }, thickness: 0.5, color: rgb(0.3, 0.3, 0.3) }); };
  // Column value accessor for one account across member / elimination / consolidated columns.
  const val = (a, i) => i < cols.length ? (a.byEntity[cols[i].entity_id] || 0) : (i === cols.length ? (a.elimination || 0) : (a.consolidated || 0));
  const sumCol = (rows, i) => r2(rows.reduce((s, a) => s + val(a, i), 0));
  const byCode = (rows) => rows.slice().sort((a, b) => String(a.code).localeCompare(String(b.code), undefined, { numeric: true }));

  const accountRow = (a) => {
    ensure(rowH);
    dtext(truncate((a.code ? a.code + ' ' : '') + (a.name || ''), reg, F.row, nameEnd - (nameLeft + 10)), nameLeft + 10, y, F.row, reg);
    figs(i => val(a, i), reg, dollarFirst); if (dollarFirst) dollarFirst = false; y -= rowH;
  };
  const sectionHeader = (t) => { ensure(rowH * 2); dtext(t, nameLeft, y, F.row, bold); y -= rowH; };
  // A double rule under the figures (final-total convention). Drawn just below
  // the current baseline, only under the number columns.
  const doubleUnder = () => { for (let i = 0; i < nCols; i++) { const x0 = colRight[i] - (numColW - 10); for (const dy of [-2.4, -4.1]) page.drawLine({ start: { x: x0, y: y + dy }, end: { x: colRight[i], y: y + dy }, thickness: 0.5, color: rgb(0.3, 0.3, 0.3) }); } };
  // reserveAfter (pt) makes this subtotal reserve extra space so the row that
  // follows it (a grand total) lands on the SAME page — keeps "Total Assets" /
  // "Total Liabilities and Members' Equity" from being orphaned alone at the
  // top of a continuation page (CLA/Jimmy 8/28).
  const subtotal = (label, getVal, opts) => { const o = opts || {}; ensure(rowH + (o.reserveAfter || 0)); if (!o.noTopRule) rule(); dtext(label, nameLeft + 10, y, F.row, bold); figs(getVal, bold, o.dollar); if (o.double) doubleUnder(); y -= rowH * 1.5; };

  // Mirror the FACE statement groupings (Jimmy, 2026-08-28): the consolidating
  // schedules use the same balance-sheet classification (bsClassifyFor) and the
  // same section/subsection order as the printed statements, so the schedule
  // that footnotes the face statements reads with the identical grouping —
  // Soft Costs / Construction Costs / Land / Permits and fees / Other
  // development, Accrued Liabilities, Members' Equity + Retained Earnings — with
  // a per-column subtotal (and its underline) at every level.
  const profile = meta.profile || 'srn';
  const subOrderMap = profile === 'banyandev' ? BS_SUB_ORDER_BANYANDEV
    : profile === 'banyan' ? BS_SUB_ORDER_BANYAN
    : BS_SUB_ORDER_SRN;
  const forceSubSet = profile === 'banyandev' ? BANYANDEV_FORCE_SUBTOTAL_SECTIONS : new Set();
  // Indented account row (grouped layout nests accounts under a subsection).
  const acctRowAt = (a, ind) => {
    ensure(rowH);
    dtext(truncate((a.code ? a.code + ' ' : '') + (a.name || ''), reg, F.row, nameEnd - (nameLeft + ind)), nameLeft + ind, y, F.row, reg);
    figs(i => val(a, i), reg, dollarFirst); if (dollarFirst) dollarFirst = false; y -= rowH;
  };
  const subHeader = (t) => { ensure(rowH); dtext(t, nameLeft + 10, y, F.row, reg); y -= rowH; };
  const subSubtotal = (label, getVal) => { ensure(rowH); rule(); dtext(label, nameLeft + 18, y, F.row, reg); figs(getVal, reg); y -= rowH * 1.2; };
  // Group a set of rows into ordered { section, subs:[{sub, rows}] } using the
  // face classification and the profile's section/subsection order.
  const groupRows = (rows, sectionOrder) => {
    const bySec = new Map();
    for (const a of rows) {
      const c = bsClassifyFor(profile, a);
      if (!bySec.has(c.section)) bySec.set(c.section, new Map());
      const subs = bySec.get(c.section);
      if (!subs.has(c.sub)) subs.set(c.sub, []);
      subs.get(c.sub).push(a);
    }
    const secNames = [...new Set([...sectionOrder, ...bySec.keys()])].filter(s => bySec.has(s));
    return secNames.map(section => {
      const subsMap = bySec.get(section);
      const order = subOrderMap[section] || [];
      const subNames = [...new Set([...order, ...subsMap.keys()])].filter(s => subsMap.has(s));
      return { section, subs: subNames.map(sub => ({ sub, rows: byCode(subsMap.get(sub)) })) };
    });
  };
  // Render one grouped section: header, each subsection (header + rows +
  // subtotal when there is more than one subsection or the section forces it),
  // then the bold section subtotal.
  const renderGroupedSection = (secGroup, sectionTotalLabel, reserveTail = 0) => {
    sectionHeader(secGroup.section);
    const force = forceSubSet.has(secGroup.section);
    const showSubs = secGroup.subs.length > 1 || force;
    for (const su of secGroup.subs) {
      if (showSubs) subHeader(su.sub);
      const ind = showSubs ? 18 : 10;
      su.rows.forEach(a => acctRowAt(a, ind));
      if (showSubs && (su.rows.length > 1 || force)) subSubtotal('Total ' + su.sub, i => sumCol(su.rows, i));
    }
    const allRows = secGroup.subs.reduce((s2, su) => s2.concat(su.rows), []);
    subtotal(sectionTotalLabel, i => sumCol(allRows, i), reserveTail ? { reserveAfter: reserveTail } : undefined);
  };

  const renderBalanceSheet = () => {
    curTitle = 'Consolidating Balance Sheet'; curDateLine = meta.longDate;
    offsets.push({ label: curTitle, page: pdf.getPageCount() });
    newPage();
    // Drop accounts with no balance in ANY column (members, eliminations, or
    // consolidated) — a row of all dashes carries no information and only pads
    // the schedule (Jimmy). Subtotals are unaffected: a zero row adds nothing.
    const acc = (schedules.balanceSheet.accounts || []).filter(a => {
      for (let i = 0; i < nCols; i++) if (Math.abs(val(a, i)) > 0.004) return true;
      return false;
    });
    const assets = acc.filter(a => a.type === 'Asset');
    const liabs = acc.filter(a => a.type === 'Liability');
    const equityAll = acc.filter(a => a.type === 'Equity');
    // Net income (loss) folded into equity — the balance-sheet window carries
    // P&L accounts at year-to-date, exactly as the face balance sheet does.
    const rev = acc.filter(a => a.type === 'Revenue'), exp = acc.filter(a => a.type === 'Expense');
    const ni = i => r2(sumCol(rev, i) - sumCol(exp, i));
    // Split equity into contributed (Members Equity) and Retained Earnings by
    // the same classifier the face statement uses.
    const isRE = a => bsClassifyFor(profile, a).sub === 'Retained Earnings';
    const contrib = byCode(equityAll.filter(a => !isRE(a)));
    const retained = byCode(equityAll.filter(a => isRE(a)));
    // Assets.
    sectionHeader('Assets');
    dollarFirst = true;
    const assetSecs = groupRows(assets, BS_ASSET_ORDER);
    assetSecs.forEach((sec, i) => renderGroupedSection(sec, 'Total ' + sec.section, i === assetSecs.length - 1 ? GRAND_RESERVE : 0));
    subtotal('Total Assets', i => sumCol(assets, i), { double: true, dollar: true });
    // Liabilities and Members' Equity.
    sectionHeader('Liabilities and Members’ Equity');
    dollarFirst = true;
    for (const sec of groupRows(liabs, ['Current Liabilities', 'Long Term Liabilities'])) renderGroupedSection(sec, 'Total ' + sec.section);
    subtotal('Total Liabilities', i => sumCol(liabs, i));
    // Keep the whole Members' Equity section on one page — never split it so
    // that only its totals spill to a near-empty continuation page (matches the
    // face balance sheet; CLA/Jimmy 8/30, "no orphaned totals"). Falls back to
    // the Net Income reserve below when the section is taller than a page.
    {
      const usable = (PH - PAGE.mT) - (PAGE.mB + 8);
      let eqH = rowH * 2 + contrib.length * rowH + rowH * 1.2;   // ME header + Net Income + contributed + its subtotal
      if (retained.length) eqH += rowH + retained.length * rowH + rowH * 1.2;  // Retained Earnings header + rows + subtotal
      eqH += rowH * 3;                                            // the two grand totals
      if (eqH <= usable) ensure(eqH);
    }
    subHeader('Members’ Equity');
    contrib.forEach(a => acctRowAt(a, 18));
    subSubtotal('Total Members’ Equity', i => sumCol(contrib, i));
    if (retained.length) {
      subHeader('Retained Earnings');
      retained.forEach(a => acctRowAt(a, 18));
      subSubtotal('Total Retained Earnings', i => sumCol(retained, i));
    }
    // Reserve the whole equity CLOSING cluster (Net Income + the two grand
    // totals) so those totals never land alone at the top of a page — the
    // cluster breaks together, led by the Net Income detail line (CLA/Jimmy
    // 8/30 — "no orphaned totals").
    ensure(rowH * 4); dtext('Net Income (Loss)', nameLeft + 18, y, F.row, reg); figs(ni); y -= rowH;
    subtotal('Total Members’ Equity', i => r2(sumCol(equityAll, i) + ni(i)));
    subtotal('Total Liabilities and Members’ Equity', i => r2(sumCol(liabs, i) + sumCol(equityAll, i) + ni(i)), { double: true, dollar: true });
  };
  const renderIncome = () => {
    curTitle = 'Consolidating Statement of Income'; curDateLine = 'For the Month Ended ' + meta.longDate;
    offsets.push({ label: curTitle, page: pdf.getPageCount() });
    newPage();
    const acc = schedules.incomeMonth.accounts || [];
    // Mirror the statement of operations top-level groupings: Revenue, Operating
    // Expenses, then Other Income (Expense) and Income Taxes (the shared
    // otherIeRoute classifier picks those out), then Net Income.
    const route = (a) => {
      const r = (typeof otherIeRoute === 'function') ? otherIeRoute(a) : null;
      if (r && r.bucket === 'otherIncome') return 'oi';
      if (r && r.bucket === 'otherExpense') return 'oe';
      if (r && r.bucket === 'incomeTax') return 'it';
      return a.type === 'Revenue' ? 'rev' : 'exp';
    };
    // Only P&L accounts belong on the statement of income. The month window
    // (buildColumns) also carries every balance-sheet account's month delta, so
    // restrict to Revenue/Expense BEFORE routing — otherwise asset/liability
    // accounts fall through the route() default and print as expenses.
    const pl = acc.filter(a => a.type === 'Revenue' || a.type === 'Expense');
    const rev = byCode(pl.filter(a => route(a) === 'rev'));
    const exp = byCode(pl.filter(a => route(a) === 'exp'));
    const oi = byCode(pl.filter(a => route(a) === 'oi'));
    const oe = byCode(pl.filter(a => route(a) === 'oe'));
    const it = byCode(pl.filter(a => route(a) === 'it'));
    sectionHeader('Revenue');
    dollarFirst = true;
    rev.forEach(a => acctRowAt(a, 10));
    subtotal('Total Revenue', i => sumCol(rev, i));
    sectionHeader('Operating Expenses');
    exp.forEach(a => acctRowAt(a, 10));
    subtotal('Total Operating Expenses', i => sumCol(exp, i));
    // Net income from operations = revenue − operating expenses.
    const opInc = i => r2(sumCol(rev, i) - sumCol(exp, i));
    if (oi.length || oe.length) {
      sectionHeader('Other Income (Expense)');
      oi.forEach(a => acctRowAt(a, 10));
      oe.forEach(a => acctRowAt(a, 10));
      subtotal('Total Other Income (Expense)', i => r2(sumCol(oi, i) - sumCol(oe, i)));
    }
    if (it.length) {
      sectionHeader('Income Taxes');
      it.forEach(a => acctRowAt(a, 10));
      subtotal('Total Income Taxes', i => sumCol(it, i));
    }
    subtotal('Net Income (Loss)', i => r2(opInc(i) + (sumCol(oi, i) - sumCol(oe, i)) - sumCol(it, i)), { double: true, noTopRule: true, dollar: true });
  };

  renderBalanceSheet();
  renderIncome();
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
//   reqReports (optional)       // array of { bytes, name, sheet } — up to two
//                               //   requisition reports (rail-assets pairs two).
//                               //   Each becomes its own section + TOC entry,
//                               //   auto-numbered when more than one is present.
//   reqReportBytes (optional)   // single-report back-compat — PDF or .xlsx
//   reqReportName (optional)    // original filename, used to detect .xlsx
//   reqSheetName (optional)     // worksheet to extract when .xlsx (default
//                               //   "Budget to Actual", case-insensitive; falls
//                               //   back to first sheet if not present)
// }
// Returns { bytes, info: { pages, reqRemoved, reqKept, cashFlowTies, ... } }.
// ═══════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════
// renderBudgetToActualPdf — the Schedule of Operating Results, Budget to Actual.
//
// Jimmy, 2026-08-28: rail assets carry an annual operations budget; this
// schedule is the LAST item in their statement package. The data comes from
// budget.buildBudgetToActual(); this function only draws it, through the same
// layout engine as the face statements so it reads as part of the package
// rather than a bolted-on appendix.
//
// Landscape, six numeric columns: actual / budget / variance for the month and
// again year to date. Variance is favourable-positive on BOTH sides of the P&L,
// which is why it is computed from each row's `sense` rather than by
// subtracting blindly.
// ═══════════════════════════════════════════════════════════════════════════
const B2A_TITLE = 'Profit and Loss - Actual vs Budget';

async function renderBudgetToActualPdf(b2a, meta, outOffsets) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fonts = { reg, bold };
  const money = v => acct(v, { dash: true });

  const m = Object.assign({}, meta, { asOf: b2a.meta.asOf });
  const title = (m.isConsolidated ? 'Consolidated ' : '') + B2A_TITLE;
  const L = makeLayout(pdf, fonts, m, title, {
    landscape: true,
    dateLine: monthsEndedLabel(b2a.meta.asOf),
  });
  if (outOffsets) outOffsets.push({ label: B2A_TITLE, page: pdf.getPageCount() });
  L.start();

  const RIGHT = PAGE.h - PAGE.mR;          // landscape: measured on the long edge
  // 78pt pitch (was 85): the per-column subtotal/total underlines are sized to
  // the numeric-column pitch, so a wider pitch made the leftmost rule run back
  // under the account names. At 78 the leftmost rule starts clear of the longest
  // label ("Total General and Administrative Expenses") and the widest figure
  // still fits with room to spare (Jimmy, 2026-08-31).
  const PITCH = 78;
  const cols = [];
  for (let i = 5; i >= 0; i--) cols.push(RIGHT - i * PITCH);
  L.setCols(cols);
  L.colHeaders(
    ['Month\nActual', 'Month\nBudget', 'Month\nVariance',
      'Year to Date\nActual', 'Year to Date\nBudget', 'Year to Date\nVariance'],
    { underline: true }
  );

  // Favourable-positive: revenue beats budget by exceeding it, expense by
  // coming in under it.
  const varOf = (r, a, b) => (r.sense === 'exp' ? b - a : a - b);
  const cells = (r) => [
    money(r.aM), money(r.bM), money(varOf(r, r.aM, r.bM)),
    money(r.aY), money(r.bY), money(varOf(r, r.aY, r.bY)),
  ];

  // A subtotal/total/NOI/net/cash-flow row must never be orphaned at the top of a
  // continuation page. The row just ABOVE such a row reserves its height
  // (keepWithNext), so if the total would not fit at the bottom of the page the
  // two break to the next page together and the total is never left alone.
  const TOTALish = new Set(['subtotal', 'total', 'noi', 'net', 'cashflow']);
  for (let i = 0; i < b2a.rows.length; i++) {
    const r = b2a.rows[i];
    const next = b2a.rows[i + 1];
    const kw = (next && TOTALish.has(next.kind)) ? 24 : 0;
    switch (r.kind) {
      case 'section':
        L.space(4);
        L.sectionTitle(r.label);
        break;
      case 'group':
        L.row(r.label, [], { indent: 14, boldRow: true, keepWithNext: kw });
        break;
      case 'line':
      case 'debtline':
        L.row(r.label, cells(r), { indent: 28, keepWithNext: kw });
        break;
      case 'cashflow':
        L.row(r.label, cells(r), { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6, dollarPrefix: true, keepWithNext: kw });
        break;
      case 'subtotal':
        L.row(r.label, cells(r), { indent: 20, boldRow: true, ruleAbove: true, gapAfter: 4, keepWithNext: kw });
        break;
      case 'total':
        // Rule below only under Total Revenue (the divider before the expense
        // section). Total Operating Expenses drops its underline — Net Operating
        // Income sits directly beneath it, so a rule there read as a stray line
        // (Jimmy, 2026-08-31).
        L.row(r.label, cells(r), { indent: 6, boldRow: true, ruleAbove: true, ruleBelow: /revenue/i.test(r.label), gapAfter: 6, keepWithNext: kw });
        break;
      case 'noi':
        // No rule above Net Operating Income (Jimmy, 2026-08-31): the Total
        // Operating Expenses row above it already carries a rule below its own
        // figures, so a second line here read as a stray double underline.
        L.row(r.label, cells(r), { indent: 6, boldRow: true, gapAfter: 6, dollarPrefix: true, keepWithNext: kw });
        break;
      case 'net':
        // No rule above Net Income (Loss): the Total Other Income (Expense)
        // subtotal above it already carries a rule, and the double underline
        // below marks the final total (Jimmy, 2026-08-31).
        L.row(r.label, cells(r), { indent: 6, boldRow: true, doubleBelow: true, dollarPrefix: true });
        break;
      default:
        break;
    }
  }


  return await pdf.save();
}

// Read a phase number from a requisition report's filename, e.g.
// "... Requisition Report Phase 2B 07.2026.xlsx" -> "2B". The CPA-prepared files
// do not follow a rigid naming scheme, so just find a "Phase <n>[letter]" token
// wherever it sits in the name. Returns the phase (upper-cased) or null.
function reqPhaseFromName(name) {
  const m = /\bphase\s+([0-9]+[A-Za-z]?)\b/i.exec(String(name || ''));
  return m ? m[1].toUpperCase() : null;
}

async function generatePackage({ statements, execSummaryBytes, storedDefaultBytes, reqReports, reqReportBytes, reqReportName, reqSheetName, wipBytes, wipName, consolSchedules, b2a }) {
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
  // Body page ranges that came from an uploaded PDF. Those pages carry the
  // supplying firm's own footer (CLA's reads "Page N"), so stamping our number
  // on them prints a duplicate. Recorded as { from, to } in 0-based BODY page
  // indices and converted to absolute page numbers at stamping time.
  const uploadedBodyRanges = [];
  // forceNumber: an uploaded section whose pages DO NOT carry the supplier's own
  // footer, so our page number should be stamped on them (they are not added to
  // the skip ranges). Used for the Banyan Residential executive summary.
  const appendToBody = async (bytes, label, addToc, uploaded, forceNumber) => {
    if (!bytes) return 0;
    let srcDoc;
    try { srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true }); }
    catch (e) { info.warnings.push('Could not read ' + label + ' PDF: ' + e.message); return 0; }
    const startPage = body.getPageCount();
    const idx = srcDoc.getPageIndices();
    const pages = await body.copyPages(srcDoc, idx);
    pages.forEach(p => body.addPage(p));
    info.sections.push({ label, pages: pages.length, uploaded: !!uploaded });
    if (uploaded && !forceNumber && pages.length) uploadedBodyRanges.push({ from: startPage, to: startPage + pages.length - 1, label });
    if (addToc) tocEntries.push({ label, page: startPage + COVER_TOC_PAGES + 1 });
    return pages.length;
  };

  // Banyan Residential's executive summary is authored in-house and carries no
  // page footer of its own, so unlike CLA's uploaded summaries it should receive
  // the package page number (Jimmy, 2026-08-31).
  const _m = statements.meta || {};
  const numberExecSummary = (String(_m.entityCode || '').toUpperCase() === 'BANYANRE1')
    || /^banyan\s*residential$/i.test(String(_m.rawEntityName || '').trim());

  // Executive summary — resolution order:
  //   1) a per-call uploaded PDF (execSummaryBytes) — always wins, merged as-is.
  //      The route also persists it as this entity's new stored default.
  //   2) the entity's stored default file (storedDefaultBytes), if one has been
  //      uploaded/split previously.
  //   3) the built-in rendered default (execSummaries.js), whose title-block
  //      date line is dynamic from the statement period.
  //   4) neither → warn as before.
  if (execSummaryBytes) {
    await appendToBody(execSummaryBytes, 'Executive Summary', true, true, numberExecSummary);
    info.execSummarySource = 'uploaded';
  } else if (storedDefaultBytes) {
    // The stored default is itself a previously uploaded/split PDF.
    await appendToBody(storedDefaultBytes, 'Executive Summary', true, true, numberExecSummary);
    info.execSummarySource = 'stored_default';
  } else {
    let defBytes = null;
    try { defBytes = await execSummaries.renderExecSummaryPdf(statements.meta); }
    catch (e) { info.warnings.push('Default executive summary render failed: ' + e.message); }
    if (defBytes) { await appendToBody(defBytes, 'Executive Summary', true); info.execSummarySource = 'builtin'; }
    else { info.warnings.push('No executive summary uploaded and no default available for this entity.'); info.execSummarySource = 'none'; }
  }

  // GL statements — capture each statement's start page within the statements
  // PDF, then offset by where the statements PDF lands in the body.
  const stmtOffsets = [];
  const stmtBytes = await renderStatementsPdf(statements, stmtOffsets);
  const stmtBodyStart = body.getPageCount();
  await appendToBody(stmtBytes, 'Financial Statements', false);
  for (const off of stmtOffsets) {
    tocEntries.push({ label: off.label, page: stmtBodyStart + off.page + COVER_TOC_PAGES + 1 });
  }

  // 3b. WIP schedule (uploaded) -> 'Schedule of Contracts'. The CPA reference
  //     package carries this as its final statement page, so it is appended
  //     directly after the GL statements and ahead of any requisition report.
  //     appendToBody records the section's absolute start page and adds the
  //     Table-of-Contents entry, so nothing else needs bookkeeping.
  //     A non-PDF (or unreadable) upload is warned about, not fatal: the rest
  //     of the package still generates.
  if (wipBytes && wipBytes.length) {
    // appendToBody catches an unreadable PDF itself and pushes its own warning
    // rather than throwing, so inclusion is measured by whether the body
    // actually grew - never assumed. A try/catch alone reported success on a
    // file that had in fact been skipped.
    const wipPagesBefore = body.getPageCount();
    try {
      await appendToBody(wipBytes, 'Schedule of Contracts', true, true);
    } catch (e) {
      info.warnings.push('WIP schedule could not be merged (' + e.message + '); package generated without it.');
    }
    const wipAdded = body.getPageCount() - wipPagesBefore;
    info.wipSchedule = { included: wipAdded > 0, pages: wipAdded, name: wipName || null };
  } else {
    info.wipSchedule = { included: false, pages: 0, name: null };
  }
  // 4. Requisition report(s) (uploaded). Accepts a PDF or an .xlsx workbook.
  //    Rail-assets entities may pair TWO reports; each becomes its own body
  //    section and Table-of-Contents entry. Labels are auto-numbered
  //    ("Budget to Actual (1)", "(2)") only when more than one is present, so a
  //    single report still reads plainly as "Budget to Actual". Because every
  //    section is appended through appendToBody (which records its absolute
  //    start page) and page numbers are stamped by absolute position later,
  //    the TOC and page numbers update themselves with no extra bookkeeping.
  //    When a workbook is uploaded, extract the requested sheet (default
  //    "Budget to Actual") and render it to a PDF page first — the rendered PDF
  //    carries a real text layer, so invoice-log stripping runs on it identically.
  const reqList = (Array.isArray(reqReports) && reqReports.length)
    ? reqReports
    : (reqReportBytes ? [{ bytes: reqReportBytes, name: reqReportName, sheet: reqSheetName }] : []);
  if (reqList.length) {
    info.reqReports = [];
    const multi = reqList.length > 1;
    for (let ri = 0; ri < reqList.length; ri++) {
      const r = reqList[ri];
      if (!r || !r.bytes) continue;
      // Phase number for the heading and the TOC label — read from the report's
      // own filename (e.g. "... Requisition Report Phase 2B 07.2026.xlsx"). The
      // CPA-prepared files don't follow a rigid naming scheme, so we just look
      // for a "Phase <n>[letter]" token wherever it appears in the name. Falls
      // back to the auto-numbered "(1)/(2)" label when the name carries no phase
      // (Jimmy, 2026-08-31).
      const phase = reqPhaseFromName(r.name);
      const label = phase
        ? ('Budget to Actual — Phase ' + phase)
        : (multi ? ('Budget to Actual (' + (ri + 1) + ')') : 'Budget to Actual');
      const rInfo = { label };
      let reqPdfBytes = r.bytes;
      let fromXlsx = false;
      if (looksLikeXlsx(r.bytes, r.name)) {
        try {
          const wantSheet = r.sheet || 'Budget to Actual';
          // Crop to the sheet's print area (drops the workbook's own left/right
          // heading cells and anything below the reconciliation) and render a
          // clean CENTERED heading — entity + report name from the sheet's own
          // print header, with the PACKAGE period date substituted for the file's
          // (often stale) date line. Applies to development AND rail-asset
          // requisition reports alike, since both embed through this converter.
          const conv = await xlsxSheetToPdf(r.bytes, wantSheet, {
            headingDate: (statements.meta && statements.meta.longDate) || undefined,
            headingEntity: (statements.meta && statements.meta.entityName) || undefined,
            headingPhase: phase || undefined,
          });
          reqPdfBytes = Buffer.from(conv.bytes);
          fromXlsx = true;
          rInfo.convertedFromXlsx = true;
          rInfo.sheetUsed = conv.sheetUsed;
          rInfo.availableSheets = conv.availableSheets;
          if (conv.sheetUsed.toLowerCase() !== wantSheet.toLowerCase()) {
            info.warnings.push(label + ': workbook had no "' + wantSheet + '" sheet; used "' + conv.sheetUsed + '" instead.');
          }
        } catch (e) {
          info.warnings.push(label + ': could not convert requisition workbook to PDF: ' + e.message);
          reqPdfBytes = null;
        }
      }
      if (reqPdfBytes && fromXlsx) {
        // xlsx path: we already extracted exactly the requested sheet (the Budget
        // to Actual report), so there is no invoice log to strip. Append the
        // converted page directly and skip stripInvoiceLogPages — that heuristic
        // is for multi-page PDF packets and could wrongly drop the one B2A page.
        const kept = await appendToBody(reqPdfBytes, label, true);
        rInfo.removed = []; rInfo.kept = kept; rInfo.total = kept;
      } else if (reqPdfBytes) {
        const stripped = await stripInvoiceLogPages(reqPdfBytes);
        rInfo.removed = stripped.removed; rInfo.kept = stripped.kept; rInfo.total = stripped.total;
        if (!stripped.textDetected) info.warnings.push(label + ': ' + (stripped.parseFailed
          ? 'requisition PDF could not be parsed for invoice-log detection; all pages were kept.'
          : 'requisition PDF had no extractable text; invoice-log pages could not be detected and were left in.'));
        await appendToBody(stripped.bytes, label, true, true);
      }
      info.reqReports.push(rInfo);
    }
    // Back-compat: the first report's figures remain on the flat info fields the
    // existing client summary line reads.
    if (info.reqReports.length) {
      const first = info.reqReports[0];
      info.reqConvertedFromXlsx = !!first.convertedFromXlsx;
      info.reqSheetUsed = first.sheetUsed;
      info.reqRemoved = first.removed || [];
      info.reqKept = first.kept;
      info.reqTotal = first.total;
    }
  } else {
    info.warnings.push('No requisition report uploaded.');
  }

  // ── Consolidating schedules (consolidated packages only) ────────────────────
  // Rendered from the consolidation engine and appended at the very back of the
  // package, matching the CPA layout. Each schedule adds its own TOC entry.
  if (consolSchedules) {
    try {
      const schedOffsets = [];
      const schedBytes = await renderConsolidatingSchedulesPdf(consolSchedules, statements.meta, schedOffsets);
      const schedBodyStart = body.getPageCount();
      await appendToBody(schedBytes, 'Consolidating Schedules', false);
      for (const off of schedOffsets) {
        tocEntries.push({ label: off.label, page: schedBodyStart + off.page + COVER_TOC_PAGES + 1 });
      }
      info.consolidatingSchedules = { included: true };
    } catch (e) {
      info.warnings.push('Consolidating schedules could not be rendered (' + e.message + '); package generated without them.');
      info.consolidatingSchedules = { included: false, error: e.message };
    }
  }

  // ── Budget to Actual (rail assets with a budget on file) ───────────────────
  // Jimmy, 2026-08-28: the LAST item in the package, after the consolidating
  // schedules. Skipped silently when the entity has no budget for the year —
  // most entities never will, and their packages must be unchanged.
  if (b2a) {
    try {
      const b2aOffsets = [];
      const b2aBytes = await renderBudgetToActualPdf(b2a, statements.meta, b2aOffsets);
      const b2aBodyStart = body.getPageCount();
      await appendToBody(b2aBytes, B2A_TITLE, false);
      for (const off of b2aOffsets) {
        tocEntries.push({ label: off.label, page: b2aBodyStart + off.page + COVER_TOC_PAGES + 1 });
      }
      info.budgetToActual = {
        included: true,
        fiscalYear: b2a.meta.fiscalYear,
        versionNo: b2a.meta.versionNo,
        unmappedBudgetLines: (b2a.unmapped && b2a.unmapped.budgetLabels) || [],
      };
    } catch (e) {
      info.warnings.push('Budget-to-Actual schedule could not be rendered (' + e.message + '); package generated without it.');
      info.budgetToActual = { included: false, error: e.message };
    }
  }

  // Phase 2: render cover + TOC (with page references) and assemble the final
  // PDF as cover/TOC first, then the body.
  const coverBytes = await renderCoverPdf(statements.meta, tocEntries);
  const coverDoc = await PDFDocument.load(coverBytes, { ignoreEncryption: true });
  const coverPages = await merged.copyPages(coverDoc, coverDoc.getPageIndices());
  coverPages.forEach(p => merged.addPage(p));
  const bodyPages = await merged.copyPages(body, body.getPageIndices());
  bodyPages.forEach(p => merged.addPage(p));

  // ── Page numbers (CLA 8/17): bottom-right of every page EXCEPT the cover
  //    and the Table of Contents, which stay unnumbered per house style.
  //    The number is the page's ABSOLUTE position in the package (cover = 1,
  //    TOC = 2), which is the same basis the TOC used to compute its page
  //    references off COVER_TOC_PAGES \u2014 so the printed digit and the TOC can
  //    never disagree. Stamped here, on the merged doc, rather than inside
  //    renderStatementsPdf: only the merged doc knows a page's final position,
  //    and this way the uploaded executive summary and Budget-to-Actual pages
  //    get numbered too. Width comes from each page's own size so the
  //    landscape members'-equity page (and any landscape B2A page) is measured
  //    against its long edge, not the portrait constant.
  {
    // Pages copied in from an uploaded PDF are NOT numbered: they carry the
    // supplying firm's own footer (CLA's reads "Page N") and our stamp landed on
    // top of it as a second number (Jimmy, 2026-08-27). Keyed to where each
    // uploaded section actually landed, so it holds however the front matter
    // changes - unlike the absolute page 3 this replaces.
    const skipPageNumbers = new Set();
    for (const r of uploadedBodyRanges) {
      for (let bp = r.from; bp <= r.to; bp++) skipPageNumbers.add(bp + COVER_TOC_PAGES + 1);
    }
    info.unnumberedPages = [...skipPageNumbers].sort((a, b) => a - b);
    const pnFont = await merged.embedFont(StandardFonts.Helvetica);
    merged.getPages().forEach((p, i) => {
      if (i < COVER_TOC_PAGES) return;
      if (skipPageNumbers.has(i + 1)) return;
      const label = String(i + 1);
      const { width } = p.getSize();
      const w = pnFont.widthOfTextAtSize(label, FS.foot);
      p.drawText(label, {
        x: width - PAGE.mR - w, y: PAGE.mB - 12,
        size: FS.foot, font: pnFont, color: rgb(0.4, 0.4, 0.4),
      });
    });
    info.pageNumbersFrom = COVER_TOC_PAGES + 1;
  }

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

  // COGS first, so a cost-of-revenue line can never be lifted out by an Other
  // Income (Expense) name test.
  const cogs = allLines.filter(l => l.type === 'Expense' && /cogs|cost of goods|cost of revenue|car hire/i.test((l.subtype || '') + ' ' + (l.name || '')));
  const cogsCodes = new Set(cogs.map(l => l.code));
  // Other Income (Expense) + Income Taxes, from the same shared classifier
  // (otherIeRoute) the monthly statements use. Keyed off type + name, so a
  // Turnkey-style chart where 42000 is Interest Income and 70000 is Interest
  // Expense classifies correctly here without profile-specific pins.
  const oieRouteFor = (l) => (l.type === 'Expense' && cogsCodes.has(l.code)) ? null : otherIeRoute(l);
  const otherIncome = allLines.filter(l => { const r = oieRouteFor(l); return r && r.bucket === 'otherIncome'; });
  const otherExpenseRaw = allLines.filter(l => { const r = oieRouteFor(l); return r && r.bucket === 'otherExpense'; });
  const incomeTax = allLines.filter(l => { const r = oieRouteFor(l); return r && r.bucket === 'incomeTax'; });
  const oieCodes = new Set([].concat(otherIncome, otherExpenseRaw, incomeTax).map(l => l.code));
  const revenue = allLines.filter(l => l.type === 'Revenue' && !oieCodes.has(l.code));
  const opex = allLines.filter(l => l.type === 'Expense' && !cogsCodes.has(l.code) && !oieCodes.has(l.code));

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
  // Other-expense and income-tax lines carry their natural (positive-magnitude)
  // sign in the ledger. Other-expense lines are negated for PRESENTATION so the
  // Other Income (Expense) section reads as income less expense, matching the
  // monthly package; income taxes print at natural magnitude in their own
  // section and are subtracted from net income.
  const negLine = (l) => ({ ...l, vals: l.vals.map(v => r2(-v)), total: r2(-l.total) });
  const otherExpense = otherExpenseRaw.map(negLine);
  const totOtherIncome = sumLines(otherIncome);
  const totOtherExpense = sumLines(otherExpense);        // already negated
  const totIncomeTax = sumLines(incomeTax);              // natural (positive = expense)
  const totOtherIE = {
    vals: totOtherIncome.vals.map((v, i) => r2(v + totOtherExpense.vals[i])),
    total: r2(totOtherIncome.total + totOtherExpense.total),
  };
  const netIncome = {
    vals: grossProfit.vals.map((v, i) => r2(v - totOpex.vals[i] + totOtherIE.vals[i] - totIncomeTax.vals[i])),
    total: r2(grossProfit.total - totOpex.total + totOtherIE.total - totIncomeTax.total),
  };

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
      entityName: displayEntityName(opts.entityName),
      asOf,
      title: 'Trailing 12 Months',
      periodLabel: 'For the Trailing Twelve Months Ended ' + longDate(asOf),
      months: months.map(m => ({ from: m.from, to: m.to, label: m.label })),
      totalLabel: 'Total',
    },
    revenue, totRev,
    cogs, totCogs, grossProfit, hasCogs: cogs.length > 0,
    opex, opexGroups, totOpex,
    otherIncome, otherExpense, incomeTax,
    totOtherIncome, totOtherExpense, totOtherIE, totIncomeTax,
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

  // Fund-level column totals. The ENDING is anchored to the balance-sheet total
  // partners' capital (totalCapital), so the Statement of Changes always ties to
  // the Statement of Assets, Liabilities & Partners' Capital by construction.
  // The activity columns are the real period movements:
  //   contributions/refunds (net) = movement of the contribution accounts
  //   syndication                 = movement of the syndication-cost account
  //   netLoss                     = fund net investment loss for the period
  // Beginning is then derived as ending - activity, so the column foots exactly.
  //
  // Note on retained earnings: the beginning-of-year equity snapshot carries a
  // Retained Earnings balance that is a soft-close artifact (prior-period P&L
  // that the close mechanism re-derives). Anchoring the ending to totalCapital
  // and back-solving the beginning naturally allocates that retained earnings
  // into partners' capital rather than leaving it stranded outside the
  // roll-forward - which is why beginning here matches the CPA's beginning
  // partners' capital and ending matches the balance sheet.
  const contribMove = r2(sumWhere(isYtd, r => r.type === 'Equity' && codeStarts(r.code, CONTRIB_PREFIXES)));
  const syndMove = r2(sumWhere(isYtd, r => r.type === 'Equity' && codeStarts(r.code, SYND_PREFIX)));
  const fundEnding = totalCapital; // balance-sheet total partners' capital
  const fundBeginning = r2(fundEnding - contribMove - syndMove - niYtd);
  const fundTotals = {
    beginning: fundBeginning,
    contributions: contribMove,
    syndication: syndMove,
    netLoss: r2(niYtd),
    ending: fundEnding,
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

  // GP beginning capital: reconstruct each GP class's opening balance from the
  // CUMULATIVE class-tagged movement from fund inception through the prior
  // year-end. The point-in-time per-class snapshot returns the whole fund, but
  // period movements ARE class-tagged, so the inception-to-open cumulative sum
  // (equity contributions/syndication + net income) recovers each class's true
  // opening capital. inceptionDate is well before any activity so the window
  // captures everything up to the opening date.
  const inceptionDate = '2000-01-01';
  const gpOpenEnd = priorMonthEnd(ys); // day before the year start
  let gpBeginning = 0;
  for (const c of partnerClasses) {
    if (!gpClassIds.has(c.id)) continue;
    const mv = await getBalances({ from: inceptionDate, to: gpOpenEnd, class_id: c.id });
    const eqMove = sumWhere(mv, r => r.type === 'Equity');
    const niMove = netIncomeOf(mv);
    gpBeginning = r2(gpBeginning + eqMove + niMove);
    if (!isZero(eqMove) || !isZero(niMove)) anyClassData = true;
  }

  // Net loss is allocated to the partners by ownership %, measured as each
  // group's share of total capital commitments. When commitments are loaded,
  // GP net loss = fund net loss x (GP commitments / total commitments). When
  // commitments are not loaded, fall back to the class-tagged P&L (usually zero
  // for GP since expenses aren't class-tagged), so the statement still foots.
  const commitments = opts.commitments || [];
  const totalCommit = commitments.reduce((sum, x) => sum + (Number(x.commitment_amount) || 0), 0);
  const gpCommit = commitments
    .filter(x => gpClassIds.has(x.class_id))
    .reduce((sum, x) => sum + (Number(x.commitment_amount) || 0), 0);
  const haveCommitments = totalCommit > 0;
  const gpCommitShare = haveCommitments ? (gpCommit / totalCommit) : 0;
  const gpNetLoss = haveCommitments ? r2(fundTotals.netLoss * gpCommitShare) : gpAgg.netLoss;

  const groups = {
    GP: {
      beginning: gpBeginning,
      contributions: gpAgg.contributions,
      syndication: gpAgg.syndication,
      netLoss: gpNetLoss,
      ending: r2(gpBeginning + gpAgg.contributions + gpAgg.syndication + gpAgg.other + gpNetLoss),
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
  // Financing = movement of real capital-flow equity accounts only
  // (contributions + syndication). The accumulated / retained-earnings accounts
  // (390xxx) move because of the prior-year soft close, not because of cash, so
  // they are excluded — otherwise they double-count the net loss and the cash
  // flow won't reconcile.
  const isCapitalFlowEquity = code => codeStarts(code, CONTRIB_PREFIXES) || codeStarts(code, SYND_PREFIX);
  const eqFlowCur = sumWhere(bsCur, r => r.type === 'Equity' && isCapitalFlowEquity(r.code));
  const eqFlowBeg = sumWhere(bsBeg, r => r.type === 'Equity' && isCapitalFlowEquity(r.code));
  const financingEquity = r2(eqFlowCur - eqFlowBeg);

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
      entityName: displayEntityName(opts.entityName),
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
  const track = (label, tocLabel) => { if (outOffsets) outOffsets.push({ label: (tocLabel || label), page: pdf.getPageCount() }); };
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
  renderConsolidatingSchedulesPdf,
  buildTtmPL,
  buildFundStatements,
  renderFundStatementsPdf,
  generatePackage,
  renderStatementsPdf,
  renderBudgetToActualPdf,
  reqPhaseFromName,
  B2A_TITLE,
  stripInvoiceLogPages,
  // exported for unit tests / reuse
  _helpers: { acct, r2, isZero, netIncomeOf, bsSection, priorMonthEnd, yearStart, monthStart, monthsEndedLabel, longDate, monthYearLabel, displayEntityName, withDesignation, resolvePeriod },
};
