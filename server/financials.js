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

function bsSection(row) {
  const sub = (row.subtype || '').toLowerCase();
  const name = (row.name || '').toLowerCase();
  if (row.type === 'Asset') {
    if (/current/.test(sub)) return 'Current Assets';
    if (/fixed/.test(sub)) return 'Fixed Assets';
    if (/intangible/.test(sub)) return 'Intangible Assets';
    if (/investment/.test(sub)) return 'Investments';
    if (/other/.test(sub)) return 'Other Assets';
    // Name/heuristic fallbacks for un-subtyped charts.
    if (/cash|receivable|prepaid|reserve|due from|deposit/.test(name)) return 'Current Assets';
    return 'Other Assets';
  }
  if (row.type === 'Liability') {
    if (/current/.test(sub)) return 'Current Liabilities';
    if (/long|note|loan/.test(sub)) return 'Long Term Liabilities';
    if (/loan|note payable|bot/.test(name)) return 'Long Term Liabilities';
    return 'Current Liabilities';
  }
  if (row.type === 'Equity') return 'Equity';
  return 'Other';
}

const BS_ASSET_ORDER = ['Current Assets', 'Fixed Assets', 'Intangible Assets', 'Investments', 'Other Assets'];
const BS_LIAB_ORDER = ['Current Liabilities', 'Long Term Liabilities'];

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

  function bsRowsFor(section, type) {
    return bsCodes
      .map(code => {
        const rc = colCur.map.get(code), rp = colPri.map.get(code);
        const ref = rc || rp;
        if (!ref || ref.type !== type) return null;
        if (bsSection(ref) !== section) return null;
        const cur = rc ? bal(rc) : 0, pri = rp ? bal(rp) : 0;
        if (isZero(cur) && isZero(pri)) return null;
        return { code, name: ref.name, cur: r2(cur), pri: r2(pri), change: r2(cur - pri) };
      })
      .filter(Boolean);
  }

  const assetSections = BS_ASSET_ORDER
    .map(s => ({ title: s, rows: bsRowsFor(s, 'Asset') }))
    .filter(s => s.rows.length);
  const liabSections = BS_LIAB_ORDER
    .map(s => ({ title: s, rows: bsRowsFor(s, 'Liability') }))
    .filter(s => s.rows.length);
  const equityRows = bsRowsFor('Equity', 'Equity');

  const totalAssets = { cur: r2(assetSections.reduce((s, x) => s + x.rows.reduce((a, r) => a + r.cur, 0), 0)),
                        pri: r2(assetSections.reduce((s, x) => s + x.rows.reduce((a, r) => a + r.pri, 0), 0)) };
  const totalLiab = { cur: r2(liabSections.reduce((s, x) => s + x.rows.reduce((a, r) => a + r.cur, 0), 0)),
                      pri: r2(liabSections.reduce((s, x) => s + x.rows.reduce((a, r) => a + r.pri, 0), 0)) };
  const totalContribEquity = { cur: r2(equityRows.reduce((a, r) => a + r.cur, 0)),
                               pri: r2(equityRows.reduce((a, r) => a + r.pri, 0)) };
  // Net income line: current-year YTD (cur column) and prior-year-through-prior-
  // month YTD (pri column) — matches the hand-prepared "Net Income (Loss)" row.
  const niLine = { cur: niYtd, pri: niPriYtd };
  const totalEquity = { cur: r2(totalContribEquity.cur + niLine.cur), pri: r2(totalContribEquity.pri + niLine.pri) };
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
    meta: { entityName: opts.entityName || 'Entity', asOf, priorDate: priorBsDate, longDate: longDate(asOf),
            priorLongDate: longDate(priorBsDate), monthsEnded: monthsEndedLabel(asOf),
            period: (opts.period || 'monthly').toLowerCase(), periodLabel: period.periodLabel, colLabel: period.colLabel },
    balanceSheet: { assetSections, liabSections, equityRows, totalAssets, totalLiab, totalContribEquity, niLine, totalEquity, totalLiabEquity },
    operations: { revenue, cogs, opex, totRev, totCogs, grossProfit, totOpex, netIncome },
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
function makeLayout(pdf, fonts, meta, statementTitle) {
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
    const label = meta.entityName + ', ' + meta.longDate + '  |  See Executive Summary and accompanying notes';
    page.drawText(label, { x: PAGE.mL, y: PAGE.mB - 12, size: FS.foot, font: reg, color: rgb(0.4, 0.4, 0.4) });
  }
  function ensure(space) { if (y - space < PAGE.mB + 8) newPage(); }

  const layout = {
    _subline: null,
    setSubline(s) { this._subline = s; },
    start() { newPage(); },
    get y() { return y; },
    set y(v) { y = v; },
    get page() { return page; },
    space(n) { y -= n; },
    ensure,
    // Column layout: array of right-edge x positions for numeric columns.
    setCols(rightEdges) { cols.length = 0; rightEdges.forEach(e => cols.push(e)); },
    // Column headers (right-aligned above each numeric column).
    colHeaders(labels) {
      ensure(18);
      labels.forEach((lab, i) => {
        const parts = String(lab).split('\n');
        parts.forEach((pl, pi) => {
          const w = bold.widthOfTextAtSize(pl, FS.head);
          page.drawText(pl, { x: cols[i] - w, y: y - pi * 9, size: FS.head, font: bold });
        });
      });
      y -= 9 * Math.max(1, ...labels.map(l => String(l).split('\n').length));
      y -= 3;
      // thin rule under headers
      page.drawLine({ start: { x: PAGE.mL, y }, end: { x: PAGE.mR ? PAGE.w - PAGE.mR : PAGE.w, y }, thickness: 0.7, color: rgb(0.2, 0.2, 0.2) });
      y -= 10;
    },
    sectionTitle(str) {
      ensure(16);
      page.drawText(str, { x: PAGE.mL, y, size: FS.row, font: bold });
      y -= 13;
    },
    // A data row: label (optionally indented) + numeric cells (strings already formatted).
    row(label, cells, { indent = 12, boldRow = false, ruleAbove = false, ruleBelow = false, doubleBelow = false, gapAfter = 0 } = {}) {
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
    },
  };
  return layout;
}

// Render the four statements into a fresh PDFDocument and return its bytes.
async function renderStatementsPdf(s) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fonts = { reg, bold };
  const m = s.meta;
  const money = v => acct(v, { dash: true });

  // Numeric column right-edges. Two-column statements (BS) use 2; Operations uses 3.
  const RIGHT = PAGE.w - PAGE.mR;
  const twoCols = [RIGHT - 95, RIGHT];
  const threeCols = [RIGHT - 200, RIGHT - 100, RIGHT];

  // ── 1. Balance Sheet ───────────────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Balance Sheet');
    L.setSubline(m.longDate + '   (with comparative totals as of ' + m.priorLongDate + ')');
    L.start();
    L.setCols(twoCols);
    L.colHeaders([m.longDate, m.priorLongDate]);
    L.sectionTitle('ASSETS');
    for (const sec of s.balanceSheet.assetSections) {
      L.row(sec.title, [], { indent: 6, boldRow: true });
      for (const r of sec.rows) L.row(r.name, [money(r.cur), money(r.pri)], { indent: 16 });
    }
    L.row('Total Assets', [money(s.balanceSheet.totalAssets.cur), money(s.balanceSheet.totalAssets.pri)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true, gapAfter: 8 });

    L.sectionTitle('LIABILITIES AND MEMBERS\u2019 EQUITY');
    for (const sec of s.balanceSheet.liabSections) {
      L.row(sec.title, [], { indent: 6, boldRow: true });
      for (const r of sec.rows) L.row(r.name, [money(r.cur), money(r.pri)], { indent: 16 });
    }
    L.row('Total Liabilities', [money(s.balanceSheet.totalLiab.cur), money(s.balanceSheet.totalLiab.pri)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Members\u2019 Equity', [], { indent: 6, boldRow: true });
    for (const r of s.balanceSheet.equityRows) L.row(r.name, [money(r.cur), money(r.pri)], { indent: 16 });
    L.row('Net Income (Loss)', [money(s.balanceSheet.niLine.cur), money(s.balanceSheet.niLine.pri)], { indent: 16 });
    L.row('Total Members\u2019 Equity', [money(s.balanceSheet.totalEquity.cur), money(s.balanceSheet.totalEquity.pri)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Total Liabilities and Members\u2019 Equity', [money(s.balanceSheet.totalLiabEquity.cur), money(s.balanceSheet.totalLiabEquity.pri)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }

  // ── 2. Statements of Operations ─────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statements of Operations');
    L.setSubline(m.periodLabel + '   (with year-to-date)');
    L.start();
    L.setCols(threeCols);
    const curHdr = (m.period === 'monthly' ? 'Current Month' : (m.period === 'quarterly' ? 'Current Quarter' : 'Current Year'));
    const priHdr = (m.period === 'monthly' ? 'Prior Month' : (m.period === 'quarterly' ? 'Prior Quarter' : 'Prior Year'));
    L.colHeaders([curHdr, priHdr, 'Year to Date']);
    const line = (r, o = {}) => L.row(r.name, [money(r.cur), money(r.pri), money(r.ytd)], { indent: 16, ...o });

    L.sectionTitle('Revenue');
    s.operations.revenue.forEach(r => line(r));
    L.row('Total Revenue', [money(s.operations.totRev.cur), money(s.operations.totRev.pri), money(s.operations.totRev.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    if (s.operations.cogs.length) {
      L.sectionTitle('Cost of Revenue');
      s.operations.cogs.forEach(r => line(r));
      L.row('Total Cost of Revenue', [money(s.operations.totCogs.cur), money(s.operations.totCogs.pri), money(s.operations.totCogs.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 4 });
      L.row('Gross Profit', [money(s.operations.grossProfit.cur), money(s.operations.grossProfit.pri), money(s.operations.grossProfit.ytd)], { indent: 6, boldRow: true, gapAfter: 6 });
    }
    L.sectionTitle('Operating Expenses');
    s.operations.opex.forEach(r => line(r));
    L.row('Total Operating Expenses', [money(s.operations.totOpex.cur), money(s.operations.totOpex.pri), money(s.operations.totOpex.ytd)], { indent: 6, boldRow: true, ruleAbove: true, gapAfter: 6 });
    L.row('Net Income (Loss)', [money(s.operations.netIncome.cur), money(s.operations.netIncome.pri), money(s.operations.netIncome.ytd)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
  }

  // ── 3. Statement of Cash Flows ──────────────────────────────────────────────
  {
    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows');
    L.setSubline(m.monthsEnded);
    L.start();
    L.setCols([RIGHT]);
    L.colHeaders(['Year to Date']);
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
    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\u2019 Equity');
    L.setSubline(m.monthsEnded);
    L.start();
    const eCols = [RIGHT - 285, RIGHT - 190, RIGHT - 95, RIGHT];
    L.setCols(eCols);
    L.colHeaders(['Beginning', 'Contributions', 'Net Income\n(Loss)', 'Ending']);
    for (const r of s.equity.rows) {
      L.row(r.name, [acct(r.beginning), acct(r.contributions), acct(r.netIncome), acct(r.ending)], { indent: 10 });
    }
    const t = s.equity.totals;
    L.row('Total Members\u2019 Equity', [acct(t.beginning), acct(t.contributions), acct(t.netIncome), acct(t.ending)], { indent: 6, boldRow: true, ruleAbove: true, doubleBelow: true });
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
async function renderCoverPdf(meta) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const page = pdf.addPage([PAGE.w, PAGE.h]);
  const center = (str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  center(meta.entityName, 20, bold, 520);
  center('Financial Statements', 15, reg, 486);
  center(meta.periodLabel, 12, reg, 462);
  page.drawLine({ start: { x: 180, y: 448 }, end: { x: PAGE.w - 180, y: 448 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Table of Contents', 11, bold, 380);
  const toc = ['Executive Summary', 'Balance Sheet', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\u2019 Equity', 'Requisition Report'];
  let ty = 358;
  toc.forEach(t => { center(t, 10, reg, ty); ty -= 20; });
  center('Prepared by CloudLedger \u2014 ' + meta.longDate, 9, reg, 90, rgb(0.45, 0.45, 0.45));
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

  // 1. Cover
  await appendPdf(await renderCoverPdf(statements.meta), 'Cover');
  // 2. Executive summary (uploaded, merged as-is)
  if (execSummaryBytes) await appendPdf(execSummaryBytes, 'Executive Summary');
  else info.warnings.push('No executive summary uploaded.');
  // 3. GL statements
  await appendPdf(await renderStatementsPdf(statements), 'Financial Statements');
  // 4. Requisition report (uploaded). Accepts a PDF or an .xlsx workbook. When a
  //    workbook is uploaded, extract the requested sheet (default "Budget to
  //    Actual") and render it to a PDF page first — the rendered PDF carries a
  //    real text layer, so invoice-log stripping runs on it identically.
  if (reqReportBytes) {
    let reqPdfBytes = reqReportBytes;
    if (looksLikeXlsx(reqReportBytes, reqReportName)) {
      try {
        const wantSheet = reqSheetName || 'Budget to Actual';
        const conv = await xlsxSheetToPdf(reqReportBytes, wantSheet, {
          title: (statements.meta.entityName || 'Requisition') + ' \u2014 ' + wantSheet,
        });
        reqPdfBytes = Buffer.from(conv.bytes);
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
    if (reqPdfBytes) {
      const stripped = await stripInvoiceLogPages(reqPdfBytes);
      info.reqRemoved = stripped.removed;
      info.reqKept = stripped.kept;
      info.reqTotal = stripped.total;
      if (!stripped.textDetected) info.warnings.push(stripped.parseFailed
        ? 'Requisition PDF could not be parsed for invoice-log detection; all pages were kept.'
        : 'Requisition PDF had no extractable text; invoice-log pages could not be detected and were left in.');
      await appendPdf(stripped.bytes, 'Requisition Report');
    }
  } else {
    info.warnings.push('No requisition report uploaded.');
  }

  info.cashFlowTies = statements.checks.cashFlowTies;
  info.cashFlowDiff = statements.checks.cashFlowDiff;
  info.balanceSheetTies = statements.checks.balanceSheetTies;
  info.pages = merged.getPageCount();
  const bytes = await merged.save();
  return { bytes, info };
}

module.exports = {
  buildStatements,
  generatePackage,
  renderStatementsPdf,
  stripInvoiceLogPages,
  // exported for unit tests / reuse
  _helpers: { acct, r2, isZero, netIncomeOf, bsSection, priorMonthEnd, yearStart, monthStart, monthsEndedLabel, longDate, resolvePeriod },
};
