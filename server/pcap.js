// ─── CLRF workpaper: Partners' Capital Account Statements (PCAP) ──────────────
//
// The per-investor "Statement of Changes in Capital" for County Line Rail Fund I,
// LP (CL entity 40) — one statement per investor, combined into a single
// deliverable, mirroring the fund administrator's PCAP package. Prepared like the
// other CLRF workpapers (gpfees.js / carryclawback.js): buildData -> render ->
// saveToWorkpapers, registered by index.js, located by findWorkpaper().
//
// ── Data model (fully reproducible from CL's GL; no external inputs) ──────────
// Every PCAP line is class-tagged in the fund's EQUITY accounts, because the
// fund administrator's per-investor allocations of net investment income/(loss),
// management fee, and unrealized change are posted to the GL as class-tagged
// entries in the accumulation accounts. So each investor's statement is that
// investor class's slice of the equity roll-forward, and it ties to the fund
// Statement of Changes in Partners' Capital by construction. Verified to the
// penny against the fund administrator's Q1-2026 PCAP (e.g. the 1999 Britta
// Biesecker Trust: beginning 489,642, ending 364,844).
//
// Line -> GL mapping (per investor class, per period):
//   Beginning capital        = all equity accounts, balance as of period start
//   Contributions            = credits to contribution accounts in CASH calls
//   Return of Capital        = debits to contribution accounts in CASH refunds (neg.)
//   Transfers of interest    = contribution-account moves in non-cash inter-class
//                              transfer entries (net zero fund-wide)
//   Waived development fees   = contribution-account moves in other non-cash entries
//   Syndication/offering     = movement of 370200
//   Net investment income/(loss) = movement of 390100 + 390120
//   Management fee           = movement of 390110
//   Change in unrealized     = movement of 390130 (frozen intra-year)
//   Ending (before carry)    = all equity accounts, balance as of period end
//   GP carried interest reallocation = 0 in a shortfall (opts.carryByClass hook)
// Contributions/refunds/transfers/waived are split by JE composition: a real
// capital call or refund touches cash; an inter-class transfer touches two
// investor classes' contribution accounts and no cash; a waived development fee
// is the remaining non-cash contribution movement. YTD = fiscal-year start ->
// as-of; ITD = inception -> as-of.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const FUND_EID = 40;

const CONTRIB_ACCTS = ['30100', '301100', '301200', '301300', '301800'];
const CONTRIB_SQL = CONTRIB_ACCTS.map((c) => "'" + c + "'").join(',');
const SYND_ACCTS = ['370200'];
const NETINV_ACCTS = ['390100', '390120']; // accum inc/exp + accum investment gain/loss
const MGMT_ACCTS = ['390110'];             // accum management fees
const UNREAL_ACCTS = ['390130'];           // accum unrealized gain/loss (frozen intra-year)
const EQUITY_ACCTS = [...CONTRIB_ACCTS, ...SYND_ACCTS, ...NETINV_ACCTS, ...MGMT_ACCTS, ...UNREAL_ACCTS, '390500'];

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

function resolveQuarter(quarterEnd) {
  if (!isDate(quarterEnd)) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const [y, m, d] = quarterEnd.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) {
    throw new Error('quarter_end must be a quarter end date: 03-31, 06-30, 09-30 or 12-31. Received ' + quarterEnd);
  }
  return {
    label: y + '-Q' + (m / 3), year: String(y), quarter: 'Q' + (m / 3), end: quarterEnd,
    year_start: y + '-01-01', prior_year_end: (y - 1) + '-12-31',
    quarter_start: y + ({ 3: '-01-01', 6: '-04-01', 9: '-07-01', 12: '-10-01' }[m]),
    quarter_begin: ({ 3: (y - 1) + '-12-31', 6: y + '-03-31', 9: y + '-06-30', 12: y + '-09-30' }[m]),
  };
}

// Sum helpers over a computeBalances() row set (rows carry balance/total_debit/
// total_credit; for equity accounts balance = credit - debit).
const inSet = (code, set) => set.includes(String(code));
const sumBal = (rows, set) => r2(rows.filter((r) => inSet(r.code, set)).reduce((s, r) => s + (Number(r.balance) || 0), 0));

// Split each investor class's contribution-account movement over a period into
// contributions / return of capital / transfers / waived development fees, by the
// composition of each journal entry. Returns { [class_id]: {...} }. `from` null
// means inception-to-`to`.
function classifyContributions(db, eid, from, to, gpSet) {
  const dateWhere = from ? 'je.date >= ? AND je.date <= ?' : 'je.date <= ?';
  const dateArgs = from ? [from, to] : [to];
  // Entry-level flags for every entry (in the window) that touches a contribution
  // account: does it move cash, and how many distinct investor classes does it
  // touch on the contribution accounts.
  const flags = db.prepare(
    "SELECT je.id AS id, "
    + "MAX(CASE WHEN jl.account_code LIKE '1002%' THEN 1 ELSE 0 END) AS has_cash, "
    + "COUNT(DISTINCT CASE WHEN jl.account_code IN (" + CONTRIB_SQL + ") THEN jl.class_id END) AS n_classes "
    + "FROM journal_entries je JOIN journal_lines jl ON jl.entry_id = je.id "
    + "WHERE je.entity_id = ? AND " + dateWhere + " "
    + "GROUP BY je.id "
    + "HAVING SUM(CASE WHEN jl.account_code IN (" + CONTRIB_SQL + ") THEN 1 ELSE 0 END) > 0"
  ).all(eid, ...dateArgs);
  const flagBy = new Map(flags.map((f) => [f.id, f]));
  // Per entry + class contribution credit/debit.
  const rows = db.prepare(
    "SELECT jl.entry_id AS entry_id, jl.class_id AS class_id, "
    + "SUM(jl.credit) AS cr, SUM(jl.debit) AS dr "
    + "FROM journal_entries je JOIN journal_lines jl ON jl.entry_id = je.id "
    + "WHERE je.entity_id = ? AND " + dateWhere + " AND jl.account_code IN (" + CONTRIB_SQL + ") "
    + "GROUP BY jl.entry_id, jl.class_id"
  ).all(eid, ...dateArgs);
  // Net contribution-account movement per entry (across all classes), to tell an
  // inter-investor transfer (nets to zero) from a broad opening allocation.
  const entryNet = new Map();
  for (const r of rows) entryNet.set(r.entry_id, (entryNet.get(r.entry_id) || 0) + ((Number(r.cr) || 0) - (Number(r.dr) || 0)));
  // A real transfer of interest is a non-cash, net-zero move between a small
  // number of investor classes. The one-time opening-balance migration is also
  // non-cash and net-zero but distributes across many classes, so it is treated
  // as contributions/refunds (which is what it represents). A single-class
  // non-cash contribution move is a waived development fee.
  const TRANSFER_MAX_CLASSES = 8;
  const out = {};
  for (const r of rows) {
    const f = flagBy.get(r.entry_id) || { has_cash: 0, n_classes: 0 };
    const cr = Number(r.cr) || 0, dr = Number(r.dr) || 0;
    const net0 = Math.abs(entryNet.get(r.entry_id) || 0) < 1;
    const o = out[r.class_id] || (out[r.class_id] = { contributions: 0, returnOfCapital: 0, transfers: 0, waivedDevFees: 0 });
    // Net each investor's contribution-account movement WITHIN a journal entry:
    // an offsetting call+refund booked together nets out (Weaver presents these
    // net), while a pure call or pure refund is unaffected.
    const addNet = (v) => { if (v >= 0) o.contributions += v; else o.returnOfCapital += v; };
    if (f.has_cash) { addNet(cr - dr); }
    else if (net0 && f.n_classes >= 2 && f.n_classes <= TRANSFER_MAX_CLASSES) { o.transfers += (cr - dr); }
    else if (f.n_classes <= 1) {
      // A non-cash single-class contribution move is a WAIVED DEVELOPMENT FEE only
      // for a GP/promote class (ties to SRN entity 37 acct 34014). The same move on
      // an LP class is a non-cash capital contribution (contribution-in-kind).
      if (gpSet && gpSet.has(Number(r.class_id))) { o.waivedDevFees += (cr - dr); }
      else { addNet(cr - dr); }
    }
    else { addNet(cr - dr); }
  }
  for (const k of Object.keys(out)) {
    const o = out[k];
    o.contributions = r2(o.contributions); o.returnOfCapital = r2(o.returnOfCapital);
    o.transfers = r2(o.transfers); o.waivedDevFees = r2(o.waivedDevFees);
  }
  return out;
}

// One period column (YTD or ITD) for one investor class. `contribBreak` is the
// classified contribution split for this class/period (may be undefined -> zero).
function periodColumn(ctx, eid, classId, from, to, beginAsOf, carry, contribBreak) {
  const { computeBalances } = ctx;
  const move = computeBalances(eid, from ? { from, to, class_id: classId } : { to, class_id: classId });
  const beginning = beginAsOf ? sumBal(computeBalances(eid, { as_of: beginAsOf, class_id: classId }), EQUITY_ACCTS) : 0;
  const cb = contribBreak || { contributions: 0, returnOfCapital: 0, transfers: 0, waivedDevFees: 0 };
  const syndication = sumBal(move, SYND_ACCTS);
  const netInvestment = sumBal(move, NETINV_ACCTS);
  const managementFee = sumBal(move, MGMT_ACCTS);
  const unrealized = sumBal(move, UNREAL_ACCTS);
  const endingBeforeCarry = r2(beginning + cb.contributions + cb.returnOfCapital + cb.transfers + cb.waivedDevFees
    + syndication + netInvestment + managementFee + unrealized);
  const carryRealloc = r2(carry || 0);
  return {
    beginning, contributions: cb.contributions, returnOfCapital: cb.returnOfCapital,
    transfers: cb.transfers, waivedDevFees: cb.waivedDevFees,
    syndication, netInvestment, managementFee, unrealized,
    endingBeforeCarry, carryRealloc, ending: r2(endingBeforeCarry + carryRealloc),
  };
}

function buildData(ctx, quarter, opts = {}) {
  const { db, computeBalances } = ctx;
  const eid = opts.entity_id || FUND_EID;
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  const classes = db.prepare('SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?').all(eid);
  const commitRows = db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
  const commitBy = {}; commitRows.forEach((c) => { commitBy[c.class_id] = r2(c.commitment_amount); });
  const carryBy = opts.carryByClass || {};
  const gpSet = new Set(classes.filter((c) => String(c.partner_type || '').toUpperCase() === 'GP').map((c) => Number(c.id)));

  // Classify contribution activity once per period (all classes).
  const ytdContrib = classifyContributions(db, eid, quarter.year_start, quarter.end, gpSet);
  const itdContrib = classifyContributions(db, eid, null, quarter.end, gpSet);
  // Current-quarter split (for the quarter-only PCAP schedule and the two-period
  // Statement of Changes). For Q1 the quarter == the year, so reuse ytdContrib.
  const isQ1 = quarter.quarter === 'Q1';
  const qContrib = isQ1 ? ytdContrib : classifyContributions(db, eid, quarter.quarter_start, quarter.end, gpSet);

  const investors = [];
  for (const c of classes) {
    const carry = (c.id in carryBy) ? carryBy[c.id] : 0;
    const ytd = periodColumn(ctx, eid, c.id, quarter.year_start, quarter.end, quarter.prior_year_end, carry, ytdContrib[c.id]);
    const itd = periodColumn(ctx, eid, c.id, null, quarter.end, null, carry, itdContrib[c.id]);
    const q = isQ1 ? ytd : periodColumn(ctx, eid, c.id, quarter.quarter_start, quarter.end, quarter.quarter_begin, carry, qContrib[c.id]);
    const commitment = commitBy[c.id] || 0;
    // Skip classes with no economic presence (a stray tag, an emptied transfer).
    if (Math.abs(itd.ending) < 0.005 && Math.abs(itd.contributions) < 0.005 && commitment === 0) continue;
    // Contributed capital = ITD net contribution-account balance (calls net of refunds/transfers).
    const contributed = sumBal(computeBalances(eid, { to: quarter.end, class_id: c.id }), CONTRIB_ACCTS);
    const unfunded = r2(commitment - contributed);
    investors.push({
      class_id: c.id, name: c.name,
      partner_type: String(c.partner_type || '').toUpperCase() === 'GP' ? 'GP' : 'LP',
      commitment, contributed, unfunded,
      pct_contributed: commitment ? contributed / commitment : 0,
      pct_unfunded: commitment ? unfunded / commitment : 0,
      ytd, itd, q,
    });
  }
  investors.sort((a, b) => (a.partner_type === b.partner_type ? 0 : a.partner_type === 'LP' ? -1 : 1)
    || String(a.name).localeCompare(String(b.name)));

  const sumCol = (key, sub) => r2(investors.reduce((s, i) => s + (i[key][sub] || 0), 0));
  const totCols = (key) => ({
    beginning: sumCol(key, 'beginning'), contributions: sumCol(key, 'contributions'),
    returnOfCapital: sumCol(key, 'returnOfCapital'), transfers: sumCol(key, 'transfers'),
    waivedDevFees: sumCol(key, 'waivedDevFees'), syndication: sumCol(key, 'syndication'),
    netInvestment: sumCol(key, 'netInvestment'), managementFee: sumCol(key, 'managementFee'),
    unrealized: sumCol(key, 'unrealized'), ending: sumCol(key, 'ending'),
  });
  const totals = {
    count: investors.length, ytd: totCols('ytd'), itd: totCols('itd'), q: totCols('q'),
    commitment: r2(investors.reduce((s, i) => s + i.commitment, 0)),
    contributed: r2(investors.reduce((s, i) => s + i.contributed, 0)),
  };

  return { quarter, entity_id: eid, entity_name: ent ? ent.name : ('entity ' + eid), investors, totals };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const PCT = '0.00%';
const NAVY = 'FF1F3864';
const HDR_FONT = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const SF = (o = {}) => Object.assign({ name: 'Times New Roman', size: 10 }, o);
const SUM_MONEY = '_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)';
const THIN = { style: 'thin' };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const spellQuarterEnd = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };

function headerRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR_FONT;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true, vertical: 'bottom' };
  });
  if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

// Statement lines in display order. Zero-only optional lines (transfers, waived
// development fees, carried-interest reallocation) are hidden per investor when
// zero, matching the administrator's per-investor cards.
const STMT_LINES = [
  { key: 'beginning', label: 'Beginning capital balance', always: true },
  { key: 'contributions', label: 'Contributions', always: true },
  { key: 'returnOfCapital', label: 'Return of Capital', always: true },
  { key: 'transfers', label: 'Transfers of interest', always: false },
  { key: 'waivedDevFees', label: 'Waived development fees', always: false },
  { key: 'syndication', label: 'Syndication/offering costs', always: true },
  { key: 'netInvestment', label: 'Net investment income/(loss)', always: true },
  { key: 'managementFee', label: 'Management fee', always: true },
  { key: 'unrealized', label: 'Change in unrealized appreciation/(depreciation) on investments', always: true },
  { key: 'endingBeforeCarry', label: 'Ending capital balance (Before General Partner carried interest reallocation)', rule: true },
  { key: 'carryRealloc', label: 'General Partner carried interest reallocation', always: true },
  { key: 'ending', label: 'Ending capital balance', bold: true, rule: true },
];

function buildWorkbook(data) {
  const q = data.quarter;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  // ── 1. PCAP Summary — matrix of all investors (YTD), for review/tie-out ──────
  const sm = wb.addWorksheet('PCAP Summary', { views: [{ state: 'frozen', xSplit: 1, ySplit: 4, showGridLines: false }] });
  sm.getCell('A1').value = data.entity_name + ' — Partners’ Capital Accounts (Statement of Changes in Capital)';
  sm.getCell('A1').font = F({ size: 12, bold: true });
  sm.getCell('A2').value = 'For the Quarter Ended ' + spellQuarterEnd(q.end) + ' (year-to-date). These amounts are not to be used for income tax purposes.';
  sm.getCell('A2').font = F({ size: 9, italic: true });
  const cols = ['Investor', 'Type', 'Commitment', 'Contributed', 'Unfunded',
    'Beginning', 'Contributions', 'Return of Capital', 'Transfers', 'Waived Dev Fees',
    'Syndication', 'Net Investment Inc/(Loss)', 'Management Fee', 'Chg Unrealized',
    'GP Carry Realloc', 'Ending'];
  headerRow(sm, 4, cols, [34, 6, 15, 15, 14, 15, 14, 15, 13, 13, 12, 16, 13, 13, 13, 15]);
  const numCols = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'];
  let r = 5;
  const put = (row, col, v, fmt) => { const c = sm.getCell(col + row); c.value = v; if (fmt) c.numFmt = fmt; c.font = F(); };
  for (const inv of data.investors) {
    sm.getCell('A' + r).value = inv.name; sm.getCell('A' + r).font = F();
    sm.getCell('B' + r).value = inv.partner_type; sm.getCell('B' + r).font = F();
    const y = inv.ytd;
    const vals = [inv.commitment, inv.contributed, inv.unfunded, y.beginning, y.contributions,
      y.returnOfCapital, y.transfers, y.waivedDevFees, y.syndication, y.netInvestment,
      y.managementFee, y.unrealized, y.carryRealloc, y.ending];
    vals.forEach((v, i) => put(r, numCols[i], v, MONEY));
    r++;
  }
  // Totals row
  const t = data.totals; const tr = sm.getRow(r);
  tr.getCell(1).value = 'Total — ' + t.count + ' investors'; tr.getCell(1).font = F({ bold: true });
  const tvals = { C: t.commitment, D: t.contributed, E: r2(t.commitment - t.contributed),
    F: t.ytd.beginning, G: t.ytd.contributions, H: t.ytd.returnOfCapital, I: t.ytd.transfers,
    J: t.ytd.waivedDevFees, K: t.ytd.syndication, L: t.ytd.netInvestment, M: t.ytd.managementFee,
    N: t.ytd.unrealized, O: r2(data.investors.reduce((s, i) => s + i.ytd.carryRealloc, 0)), P: t.ytd.ending };
  Object.entries(tvals).forEach(([col, v]) => { const c = tr.getCell(col); c.value = v; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN, bottom: { style: 'double' } }; });

  // ── 2. Investor Statements — one card per investor (YTD + ITD), page-broken ──
  const st = wb.addWorksheet('Investor Statements', { views: [{ showGridLines: false }] });
  st.getColumn(1).width = 2; st.getColumn(2).width = 58; st.getColumn(3).width = 20; st.getColumn(4).width = 20;
  let R = 1;
  const sfCell = (row, col, v, o = {}) => { const c = st.getCell(col + row); c.value = v; c.font = SF(o); return c; };
  const sfMoney = (row, col, v, o = {}) => { const c = st.getCell(col + row); c.value = v; c.numFmt = SUM_MONEY; c.font = SF(o); return c; };
  const sfPct = (row, col, v) => { const c = st.getCell(col + row); c.value = v; c.numFmt = PCT; c.font = SF(); return c; };
  data.investors.forEach((inv, idx) => {
    const top = R;
    const title = (text, o = {}) => { const c = st.getCell('B' + R); c.value = text; c.font = SF(Object.assign({ bold: true }, o)); c.alignment = { horizontal: 'center' }; st.mergeCells('B' + R + ':D' + R); R++; };
    title(data.entity_name, { size: 11 });
    title('Statement of Changes in Capital');
    sfCell(R, 'B', 'For the Quarter Ended ' + spellQuarterEnd(q.end), { italic: true }).alignment = { horizontal: 'center' }; st.mergeCells('B' + R + ':D' + R); R++;
    sfCell(R, 'B', 'These amounts are not to be used for income tax purposes', { italic: true, size: 9 }).alignment = { horizontal: 'center' }; st.mergeCells('B' + R + ':D' + R); R += 2;
    sfCell(R, 'B', 'Investor: ' + inv.name, { bold: true }); R++;
    sfCell(R, 'B', inv.partner_type === 'GP' ? 'General Partner' : 'Limited Partner', { size: 9, italic: true }); R += 2;

    // Commitment summary
    sfCell(R, 'B', 'Capital Commitment Summary', { bold: true }); sfCell(R, 'D', 'Amount', { bold: true }).alignment = { horizontal: 'right' }; R++;
    sfCell(R, 'B', 'Capital Commitment'); sfPct(R, 'C', 1); sfMoney(R, 'D', inv.commitment); R++;
    sfCell(R, 'B', 'Contributed capital'); sfPct(R, 'C', -inv.pct_contributed); sfMoney(R, 'D', -inv.contributed); R++;
    sfCell(R, 'B', 'Unfunded commitment'); sfPct(R, 'C', inv.pct_unfunded); sfMoney(R, 'D', inv.unfunded);
    ['B', 'C', 'D'].forEach((c) => { st.getCell(c + R).border = { top: THIN }; }); R += 2;

    // Capital summary (YTD + ITD)
    sfCell(R, 'B', 'Capital Summary', { bold: true });
    sfCell(R, 'C', 'Year-to-Date', { bold: true }).alignment = { horizontal: 'right' };
    sfCell(R, 'D', 'Inception-to-Date', { bold: true }).alignment = { horizontal: 'right' }; R++;
    for (const line of STMT_LINES) {
      const vy = inv.ytd[line.key], vi = inv.itd[line.key];
      if (!line.always && !line.rule && Math.abs(vy) < 0.005 && Math.abs(vi) < 0.005) continue;
      sfCell(R, 'B', line.label, line.bold ? { bold: true } : {});
      sfMoney(R, 'C', vy, line.bold ? { bold: true } : {});
      sfMoney(R, 'D', vi, line.bold ? { bold: true } : {});
      if (line.rule) ['C', 'D'].forEach((c) => { st.getCell(c + R).border = { top: THIN }; });
      if (line.key === 'ending') ['C', 'D'].forEach((c) => { st.getCell(c + R).border = { top: THIN, bottom: { style: 'double' } }; });
      R++;
    }
    R += 1;
    sfCell(R, 'B', 'No Assurance Provided.', { italic: true, size: 9 }); R += 1;
    sfCell(R, 'B', 'Contact: countylinerail via CloudLedger', { size: 8, italic: true, color: { argb: 'FF808080' } }); R += 1;
    st.getRow(top).addPageBreak && st.getRow(top).addPageBreak();
    // Page break after each investor except the last.
    if (idx < data.investors.length - 1) { st.getRow(R).addPageBreak(); R += 2; }
  });

  // ── 3. Notes & Sources ───────────────────────────────────────────────────────
  const nt = wb.addWorksheet('Notes & Sources'); nt.getColumn(1).width = 118;
  const notes = [
    ['CLRF Partners’ Capital Account Statements (PCAP) — ' + q.label, F({ size: 12, bold: true })],
    ['No Assurance Provided. Weaver remains responsible for the official financial statements.', F({ size: 9, italic: true })],
    ['', F()],
    ['Each investor statement is that investor class’s slice of the CLRF equity roll-forward, sourced entirely from the general ledger (entity ' + data.entity_id + ') by class tag. It ties to the fund Statement of Changes in Partners’ Capital by construction.', F()],
    ['', F()],
    ['Line sources (per investor class, per period): Beginning/Ending = all equity accounts (' + EQUITY_ACCTS.join(', ') + '); Contributions = credits and Return of Capital = debits to the contribution accounts (' + CONTRIB_ACCTS.join(', ') + ') in cash capital calls/refunds; Transfers of interest = non-cash inter-class contribution moves; Waived development fees = other non-cash contribution moves; Syndication = 370200; Net investment income/(loss) = 390100 + 390120; Management fee = 390110; Change in unrealized = 390130 (held constant intra-year per the annual valuation).', F()],
    ['', F()],
    ['Contributions / Return of Capital / Transfers / Waived development fees are separated by journal-entry composition: a real capital call or refund settles in cash; an inter-class transfer moves capital between two investor classes with no cash; a waived development fee is a non-cash contribution.', F()],
    ['', F()],
    ['Year-to-Date = ' + q.year_start + ' through ' + q.end + '. Inception-to-Date = fund inception through ' + q.end + '.', F()],
    ['', F()],
    ['Tie-out (' + q.label + ', YTD): ' + data.totals.count + ' investors; contributions $' + data.totals.ytd.contributions.toLocaleString('en-US', { minimumFractionDigits: 2 }) + '; ending partners’ capital $' + data.totals.ytd.ending.toLocaleString('en-US', { minimumFractionDigits: 2 }) + ' (ties to the fund Statement of Changes in Partners’ Capital).', F()],
  ];
  let nr = 1;
  for (const [text, font] of notes) { const c = nt.getCell('A' + nr); c.value = text; c.font = font; c.alignment = { wrapText: true }; nr++; }

  return wb;
}

// ── Persistence + routes (mirror gpfees.js / carryclawback.js). ───────────────
const folderFor = (quarter) => 'Workpapers/Partners’ Capital Accounts/' + quarter.year + '/' + quarter.quarter;
const fileNameFor = (quarter) => 'CLRF_PCAP_' + quarter.label + '.xlsx';

function saveToWorkpapers(ctx, eid, quarter, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(quarter);
  const original = fileNameFor(quarter);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, '
    + "uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function findWorkpaper(ctx, eid, quarterEnd) {
  const quarter = resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ? ORDER BY id DESC LIMIT 1').get(eid, folderFor(quarter), fileNameFor(quarter));
  if (!row) return null;
  return Object.assign({}, row, { quarter, abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerPcapRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/pcap/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const quarter = resolveQuarter(body.quarter_end || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, quarter, { entity_id: eid, carryByClass: body.carry_by_class || null });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-PCAP-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          investors: data.totals.count, ytd_contributions: data.totals.ytd.contributions,
          ytd_ending: data.totals.ytd.ending, commitment: data.totals.commitment,
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { resolveQuarter, buildData, buildWorkbook, periodColumn, classifyContributions,
  saveToWorkpapers, findWorkpaper, registerPcapRoutes, FUND_EID, EQUITY_ACCTS };
