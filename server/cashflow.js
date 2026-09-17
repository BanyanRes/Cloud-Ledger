// ─── CLRF workpaper: Statement of Cash Flows Worksheet ───────────────────────
//
// A fund Statement of Cash Flows WORKSHEET (the matrix / spreadsheet tie-out an
// accountant builds behind the face SOCF), reproducing the layout, formulas and
// notes of the vendor "SOCF Worksheet" tab EXACTLY. The worksheet itself holds
// NO hard-coded numbers: every figure is a live formula that links to supporting
// tabs which are populated straight from the general ledger:
//
//   • "BS Data"     — a comparative balance sheet (beginning = prior quarter end,
//                     ending = quarter end) built from computeBalances. Each GL
//                     asset/liability account is classified to one of the SOCF
//                     presentation columns; the 20 summary lines are SUMIF rollups
//                     of that detail (so nothing is typed in).
//   • "PL Data"     — the period income statement from the GL; Net Income is a
//                     SUMIF of the revenue/expense detail and feeds the operating
//                     section (row 9).
//   • "Cap Activity"— fund capital contributions / refunds / syndication / waived
//                     development fees for the quarter from the PCAP roll-forward
//                     (partnerscap/pcap), feeding the financing section (rows 42-45).
//
// The matrix method: row 4 = ending balances (links), row 5 = beginning balances
// (links), row 6 = change; each column's change is allocated across the operating /
// investing / financing rows and the "TOTALS (s/b zero)" row 55 proves every column
// (and the fund) ties. Partners' Capital (col Y) is the balancing plug =-SUM(assets,
// liabilities); investment purchases (E12) are the residual of the investment change
// after the non-cash pieces; waived development fees are non-cash (E45 = -Y45).
// Filed under Workpapers > Cash Flow > <year> > <quarter>.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const pcap = require('./pcap');

const NUMFMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
const FONT = { name: 'Arial', size: 10 };
const FILL = {
  asset: 'FFE2EFDA', // green  (accent6 lighter 80%)  — asset columns / inputs
  liab: 'FFFFF2CC',  // gold   (accent4 lighter 80%)  — liability columns / inputs
  cap: 'FFDDEBF7',   // blue   (accent1 lighter 80%)  — partners' capital column
  inv: 'FFEDEDED',   // gray   (accent3 lighter 80%)  — investment-adjustment inputs
};
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

// Matrix columns D..X (assets D-M, liabilities N-X) plus the Y plug column. N is
// the (currently zero) line-of-credit column and carries no balance-sheet line.
const COLS = [
  { col: 'D', hdr: 'Cash and cash equivalents', side: 'asset' },
  { col: 'E', hdr: 'Investments', side: 'asset' },
  { col: 'F', hdr: 'Prepaid insurance', side: 'asset' },
  { col: 'G', hdr: 'Interest receivable', side: 'asset' },
  { col: 'H', hdr: 'Due from portfolio investments', side: 'asset' },
  { col: 'I', hdr: 'Prepaid advisory fees', side: 'asset' },
  { col: 'J', hdr: 'Due from affiliates', side: 'asset' },
  { col: 'K', hdr: 'Deferred Development Fees', side: 'asset' },
  { col: 'L', hdr: 'Other assets', side: 'asset' },
  { col: 'M', hdr: 'Capital contributions receivable', side: 'asset' },
  { col: 'O', hdr: 'Notes payable', side: 'liab' },
  { col: 'P', hdr: 'Accounts payable and accrued expenses', side: 'liab' },
  { col: 'Q', hdr: 'Due to affiliates', side: 'liab' },
  { col: 'R', hdr: 'Interest payable', side: 'liab' },
  { col: 'S', hdr: 'Management fees payable', side: 'liab' },
  { col: 'T', hdr: 'Due to Manager', side: 'liab' },
  { col: 'U', hdr: 'Other liabilities', side: 'liab' },
  { col: 'V', hdr: 'Due to portfolio investments', side: 'liab' },
  { col: 'W', hdr: 'Due to members', side: 'liab' },
  { col: 'X', hdr: 'Capital contributions received in advance', side: 'liab' },
];
// All matrix data columns in physical order (D..X incl. N spacer) + Y plug + Z check.
const DATA_COLS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
  'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'];
const ASSET_LETTERS = new Set(['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']);

// GL account (by type + name) -> SOCF presentation line. Returns the exact header
// string used in COLS, or the netting tags for the portfolio-company accounts, or
// null for accounts that do not belong on the balance sheet.
function socfLineFor(row) {
  const n = String(row.name || '').toLowerCase();
  if (row.type === 'Asset') {
    if (/cash|sweep|checking|clearing|money (in|out)/.test(n)) return 'Cash and cash equivalents';
    if (/interest receivable/.test(n)) return 'Interest receivable';
    if (/contribution receivable/.test(n)) return 'Capital contributions receivable';
    if (/due from (port|portfolio)/.test(n)) return '_dueFromPort';
    if (/prepaid insurance/.test(n)) return 'Prepaid insurance';
    if (/prepaid advisory|advisory fee/.test(n)) return 'Prepaid advisory fees';
    if (/deferred develop/.test(n)) return 'Deferred Development Fees';
    if (/due from affiliate/.test(n)) return 'Due from affiliates';
    if (/investment|unrealized|capitalized expense|appr\/depr/.test(n)) return 'Investments';
    return 'Other assets';
  }
  if (row.type === 'Liability') {
    if (/management fee/.test(n)) return 'Management fees payable';
    if (/due to management company|due to manager/.test(n)) return 'Due to Manager';
    if (/due to affiliate/.test(n)) return 'Due to affiliates';
    if (/due to (port|portfolio)/.test(n)) return '_dueToPort';
    if (/distributions? payable|due to member/.test(n)) return 'Due to members';
    if (/interest payable/.test(n)) return 'Interest payable';
    if (/note payable|line of credit|loan|bond/.test(n)) return 'Notes payable';
    if (/received in advance|contribution.*advance/.test(n)) return 'Capital contributions received in advance';
    if (/payable|accrued|a\/p/.test(n)) return 'Accounts payable and accrued expenses';
    return 'Other liabilities';
  }
  return null;
}

function folderFor(q) { return 'Workpapers/Cash Flow/' + q.year + '/' + q.quarter; }
function fileNameFor(q) { return 'CLRF_Cash_Flow_' + q.label + '.xlsx'; }

// ─── Data (from the general ledger) ──────────────────────────────────────────
function buildData(ctx, eid, asOf) {
  const { db } = ctx;
  const ent = db.prepare('SELECT name, code, entity_type FROM entities WHERE id=?').get(eid);
  const entityName = ent ? ent.name : ('Entity ' + eid);
  const quarter = pcap.resolveQuarter(asOf);
  const cb = (o) => ctx.computeBalances(eid, o);
  const ys = (d) => d.slice(0, 4) + '-01-01';

  // Balance-sheet snapshots: ending = quarter end, beginning = prior quarter end.
  const curRows = cb({ as_of: quarter.end, close_pl_before: ys(quarter.end) });
  const priRows = cb({ as_of: quarter.quarter_begin, close_pl_before: ys(quarter.quarter_begin) });
  const curBal = {}; curRows.forEach((r) => { curBal[r.code] = r; });
  const priBal = {}; priRows.forEach((r) => { priBal[r.code] = r; });
  const codes = Array.from(new Set([...curRows.map((r) => r.code), ...priRows.map((r) => r.code)]))
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  const bsDetail = [];
  for (const code of codes) {
    const ref = curBal[code] || priBal[code];
    if (ref.type !== 'Asset' && ref.type !== 'Liability') continue;
    const line = socfLineFor(ref);
    if (!line) continue;
    bsDetail.push({
      code, name: ref.name, type: ref.type, line,
      beg: r2(priBal[code] ? priBal[code].balance : 0),
      end: r2(curBal[code] ? curBal[code].balance : 0),
    });
  }

  // Income statement: quarter (Q window) and YTD, from the GL.
  const plQ = cb({ from: quarter.quarter_start, to: quarter.end });
  const plY = cb({ from: quarter.year_start, to: quarter.end });
  const qMap = {}; plQ.forEach((r) => { qMap[r.code] = r; });
  const yMap = {}; plY.forEach((r) => { yMap[r.code] = r; });
  const plCodes = Array.from(new Set([...plQ, ...plY].filter((r) => r.type === 'Revenue' || r.type === 'Expense').map((r) => r.code)))
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const plDetail = [];
  for (const code of plCodes) {
    const ref = qMap[code] || yMap[code];
    plDetail.push({
      code, name: ref.name, type: ref.type,
      q: r2(qMap[code] ? qMap[code].balance : 0),
      ytd: r2(yMap[code] ? yMap[code].balance : 0),
    });
  }

  // Fund capital activity for the quarter (PCAP roll-forward, GL-derived).
  let fin = { contributions: 0, refunds: 0, syndication: 0, waived: 0 };
  try {
    const pc = pcap.buildData({ db, computeBalances: (e, o) => ctx.computeBalances(e, o) }, quarter, { entity_id: eid });
    const q = (pc && pc.totals && pc.totals.q) || {};
    fin = {
      contributions: r2((q.contributions || 0) + (q.transfers || 0)),
      refunds: r2(q.returnOfCapital || 0),
      syndication: r2(q.syndication || 0),
      waived: r2(q.waivedDevFees || 0),
    };
  } catch (e) { /* leave financing at zero if PCAP is unavailable */ }

  // Summary subtotals for the UI card (the worksheet folds investing into the
  // operating section, so investing is reported as 0). Cash change = ending - beg;
  // financing = contributions + refunds + syndication (waived dev fees are non-cash
  // and net to zero in the cash column); operating = the balancing remainder.
  const cashBeg = r2(bsDetail.filter((d) => d.line === 'Cash and cash equivalents').reduce((s, d) => s + d.beg, 0));
  const cashEnd = r2(bsDetail.filter((d) => d.line === 'Cash and cash equivalents').reduce((s, d) => s + d.end, 0));
  const netFinancing = r2(fin.contributions + fin.refunds + fin.syndication);
  const netChange = r2(cashEnd - cashBeg);
  const summary = { netOperating: r2(netChange - netFinancing), netInvesting: 0, netFinancing, netChange, cashEnd };

  return { eid, entityName, asOf, quarter, bsDetail, plDetail, fin, summary };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
function buildWorkbook(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  const money = (cell) => { cell.numFmt = NUMFMT; cell.font = FONT; };
  const fillOf = (rgb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: rgb } });
  const thin = { style: 'thin' };
  const dbl = { style: 'double' };

  // Create the SOCF Worksheet FIRST so it is the leftmost/opening tab; it is
  // populated further below (after the supporting tabs' row layout is known).
  const ws = wb.addWorksheet('SOCF Worksheet', { views: [{ showGridLines: false }] });

  // ── Supporting tab: BS Data ────────────────────────────────────────────────
  const bs = wb.addWorksheet('BS Data');
  bs.getColumn(1).width = 40; bs.getColumn(2).width = 15; bs.getColumn(3).width = 15; bs.getColumn(4).width = 13;
  bs.getColumn(6).width = 2; bs.getColumn(7).width = 10; bs.getColumn(8).width = 44;
  bs.getColumn(9).width = 15; bs.getColumn(10).width = 15; bs.getColumn(11).width = 22;
  const setBS = (addr, v, opts = {}) => {
    const c = bs.getCell(addr); c.value = v; c.font = opts.bold ? { ...FONT, bold: true } : FONT;
    if (opts.num) c.numFmt = NUMFMT; if (opts.align) c.alignment = { horizontal: opts.align };
    return c;
  };
  setBS('A1', 'BALANCE SHEET DATA (from the general ledger)', { bold: true });
  setBS('A2', data.entityName);
  setBS('A3', 'Beginning = ' + data.quarter.quarter_begin + '   ·   Ending = ' + data.quarter.end);
  // Summary block header
  setBS('A5', 'Presentation line', { bold: true });
  setBS('B5', 'Beginning', { bold: true, align: 'right' });
  setBS('C5', 'Ending', { bold: true, align: 'right' });
  setBS('D5', 'Change', { bold: true, align: 'right' });
  // Detail block header (to the right)
  setBS('G5', 'Code', { bold: true });
  setBS('H5', 'Account', { bold: true });
  setBS('I5', 'Beginning', { bold: true, align: 'right' });
  setBS('J5', 'Ending', { bold: true, align: 'right' });
  setBS('K5', 'Line', { bold: true });
  // Detail rows
  let dr = 6;
  const detailFirst = dr;
  for (const d of data.bsDetail) {
    setBS('G' + dr, d.code); setBS('H' + dr, d.name);
    setBS('I' + dr, d.beg, { num: true }); setBS('J' + dr, d.end, { num: true });
    setBS('K' + dr, d.line);
    dr += 1;
  }
  const detailLast = dr - 1;
  const IRANGE = `$I$${detailFirst}:$I$${detailLast}`;
  const JRANGE = `$J$${detailFirst}:$J$${detailLast}`;
  const KRANGE = `$K$${detailFirst}:$K$${detailLast}`;
  // Summary rows 6..(5+COLS.length): one per SOCF line, in COLS order.
  const lineRow = {}; // col letter -> BS Data summary row
  COLS.forEach((cinfo, i) => {
    const row = 6 + i;
    lineRow[cinfo.col] = row;
    setBS('A' + row, cinfo.hdr);
    if (cinfo.hdr === 'Due from portfolio investments') {
      setBS('B' + row, { formula: `MAX(SUMIF(${KRANGE},"_dueFromPort",${IRANGE})-SUMIF(${KRANGE},"_dueToPort",${IRANGE}),0)` }, { num: true });
      setBS('C' + row, { formula: `MAX(SUMIF(${KRANGE},"_dueFromPort",${JRANGE})-SUMIF(${KRANGE},"_dueToPort",${JRANGE}),0)` }, { num: true });
    } else if (cinfo.hdr === 'Due to portfolio investments') {
      setBS('B' + row, { formula: `MAX(SUMIF(${KRANGE},"_dueToPort",${IRANGE})-SUMIF(${KRANGE},"_dueFromPort",${IRANGE}),0)` }, { num: true });
      setBS('C' + row, { formula: `MAX(SUMIF(${KRANGE},"_dueToPort",${JRANGE})-SUMIF(${KRANGE},"_dueFromPort",${JRANGE}),0)` }, { num: true });
    } else {
      setBS('B' + row, { formula: `SUMIF(${KRANGE},$A${row},${IRANGE})` }, { num: true });
      setBS('C' + row, { formula: `SUMIF(${KRANGE},$A${row},${JRANGE})` }, { num: true });
    }
    setBS('D' + row, { formula: `B${row}-C${row}` }, { num: true });
  });

  // ── Supporting tab: PL Data ────────────────────────────────────────────────
  const pl = wb.addWorksheet('PL Data');
  pl.getColumn(1).width = 40; pl.getColumn(2).width = 15; pl.getColumn(3).width = 15; pl.getColumn(4).width = 12;
  const setPL = (addr, v, opts = {}) => {
    const c = pl.getCell(addr); c.value = v; c.font = opts.bold ? { ...FONT, bold: true } : FONT;
    if (opts.num) c.numFmt = NUMFMT; if (opts.align) c.alignment = { horizontal: opts.align };
    return c;
  };
  setPL('A1', 'INCOME STATEMENT DATA (from the general ledger)', { bold: true });
  setPL('A2', data.entityName);
  setPL('A3', 'Quarter = ' + data.quarter.quarter_start + ' to ' + data.quarter.end + '   ·   YTD = ' + data.quarter.year_start + ' to ' + data.quarter.end);
  setPL('A5', 'Account', { bold: true });
  setPL('B5', 'Quarter', { bold: true, align: 'right' });
  setPL('C5', 'YTD', { bold: true, align: 'right' });
  setPL('D5', 'Type', { bold: true });
  let pr = 6; const plFirst = pr;
  for (const d of data.plDetail) {
    setPL('A' + pr, d.name); setPL('B' + pr, d.q, { num: true });
    setPL('C' + pr, d.ytd, { num: true }); setPL('D' + pr, d.type);
    pr += 1;
  }
  const plLast = pr - 1;
  const niRow = pr + 1;
  const PB = `$B$${plFirst}:$B$${plLast}`;
  const PC = `$C$${plFirst}:$C$${plLast}`;
  const PD = `$D$${plFirst}:$D$${plLast}`;
  setPL('A' + niRow, 'Net Income', { bold: true });
  setPL('B' + niRow, { formula: `SUMIF(${PD},"Revenue",${PB})-SUMIF(${PD},"Expense",${PB})` }, { num: true }).font = { ...FONT, bold: true };
  setPL('C' + niRow, { formula: `SUMIF(${PD},"Revenue",${PC})-SUMIF(${PD},"Expense",${PC})` }, { num: true }).font = { ...FONT, bold: true };

  // ── Supporting tab: Cap Activity ───────────────────────────────────────────
  const ca = wb.addWorksheet('Cap Activity');
  ca.getColumn(1).width = 34; ca.getColumn(2).width = 16;
  const setCA = (addr, v, opts = {}) => {
    const c = ca.getCell(addr); c.value = v; c.font = opts.bold ? { ...FONT, bold: true } : FONT;
    if (opts.num) c.numFmt = NUMFMT; if (opts.align) c.alignment = { horizontal: opts.align };
    return c;
  };
  setCA('A1', 'CAPITAL ACTIVITY (PCAP roll-forward, from the general ledger)', { bold: true });
  setCA('A2', data.entityName);
  setCA('A3', 'Quarter = ' + data.quarter.quarter_start + ' to ' + data.quarter.end);
  setCA('A5', 'Financing activity', { bold: true });
  setCA('B5', 'Quarter', { bold: true, align: 'right' });
  setCA('A6', 'Capital contributions'); setCA('B6', data.fin.contributions, { num: true });
  setCA('A7', 'Capital call refunds'); setCA('B7', data.fin.refunds, { num: true });
  setCA('A8', 'Syndication costs'); setCA('B8', data.fin.syndication, { num: true });
  setCA('A9', 'Waived development fees'); setCA('B9', data.fin.waived, { num: true });

  // ── Main tab: SOCF Worksheet (created above; populated here) ────────────────
  const WIDTHS = { A: 65.7, B: 1.0, C: 1.5, D: 12.8, E: 14.5, F: 12.7, G: 10.5, H: 12.8, I: 9.2, J: 10.2, K: 12.7, L: 11.5, M: 12.8, N: 11.5, O: 10.2, P: 12.3, Q: 12.8, R: 9.3, S: 13.0, T: 9.5, U: 9.0, V: 12.3, W: 12.5, X: 13.5, Y: 19.5, Z: 15.7 };
  Object.entries(WIDTHS).forEach(([c, w]) => { ws.getColumn(c).width = w; });
  const cell = (addr) => ws.getCell(addr);
  const put = (addr, v, o = {}) => {
    const c = cell(addr);
    c.value = v;
    c.font = o.bold ? { ...FONT, bold: true } : FONT;
    if (o.num !== false) c.numFmt = NUMFMT;
    if (o.gen) c.numFmt = 'General';
    if (o.align) c.alignment = { horizontal: o.align, ...(o.wrap ? { wrapText: true, vertical: 'center' } : {}) };
    if (o.fill) c.fill = fillOf(o.fill);
    if (o.border) c.border = o.border;
    return c;
  };
  const fillForCol = (col) => (col === 'Y' ? FILL.cap : ASSET_LETTERS.has(col) ? FILL.asset : FILL.liab);

  // Row 1 title
  put('A1', 'STATEMENT OF CASH FLOWS WORKSHEET', { bold: true, gen: true });
  // Row 2: as-of date + Assets/Liabilities banner
  const d = new Date(data.asOf + 'T00:00:00Z');
  const a2 = put('A2', d, { gen: true, align: 'left' }); a2.numFmt = 'mm-dd-yy';
  for (const c of DATA_COLS) {
    if (c === 'Y' || c === 'Z') continue;
    const isA = ASSET_LETTERS.has(c);
    put(c + '2', c === 'D' ? 'Assets' : c === 'N' ? 'Liabilities' : null,
      { gen: true, bold: c === 'D' || c === 'N', fill: isA ? FILL.asset : FILL.liab, border: { top: thin, bottom: thin } });
  }
  // Row 3: column headers
  ws.getRow(3).height = 52;
  for (const cinfo of COLS) {
    put(cinfo.col + '3', cinfo.hdr, { gen: true, bold: true, align: 'center', wrap: true,
      fill: cinfo.side === 'asset' ? FILL.asset : FILL.liab, border: { top: thin, bottom: thin } });
  }
  put('Y3', "Total Partners' Capital", { gen: true, bold: true, align: 'center', wrap: true, fill: FILL.cap, border: { top: thin, bottom: thin } });
  put('N3', null, { gen: true, fill: FILL.liab, border: { top: thin, bottom: thin } });

  // Row 4 (ending) and Row 5 (beginning): links into BS Data.
  put('A4', { formula: '"Ending Balance as of "&TEXT(A2,"m/d/yyyy")' }, { bold: true, gen: true, align: 'right' });
  put('A5', { formula: '"Beginning Balance as of 4/1/"&YEAR(A2)' }, { bold: true, gen: true, align: 'right' });
  for (const cinfo of COLS) {
    const row4Ref = `'BS Data'!C${lineRow[cinfo.col]}`;
    const row5Ref = `'BS Data'!B${lineRow[cinfo.col]}`;
    const sgn = cinfo.side === 'asset' ? '+' : '-';
    put(cinfo.col + '4', { formula: `${sgn}${row4Ref}` });
    put(cinfo.col + '5', { formula: `${sgn}${row5Ref}` });
  }
  put('Y4', { formula: '-SUM(C4:X4)' }, { fill: FILL.cap });
  put('Y5', { formula: '-SUM(C5:X5)' }, { fill: FILL.cap });
  put('Z4', { formula: 'SUM(C4:Y4)' });
  put('Z5', { formula: 'SUM(C5:Y5)' });

  // Row 6: CHANGE = beginning - ending, per column.
  put('A6', 'CHANGE', { bold: true, gen: true, align: 'right' });
  for (const c of DATA_COLS) {
    const f = (c === 'Y') ? 'Y5-Y4' : `+${c}5-${c}4`;
    put(c + '6', { formula: f }, { fill: c === 'Y' ? FILL.cap : undefined, border: { top: thin, bottom: thin } });
  }
  put('Z6', { formula: 'SUM(C6:Y6)' }, { border: { top: thin, bottom: thin } });

  // Operating section
  put('A8', 'Cash flows from operating activities:', { bold: true, gen: true });
  put('A9', "Net increase (decrease) in Partners' Capital from operations", { gen: true, align: 'left' });
  put('D9', { formula: 'SUM(E9:Y9)' });
  put('Y9', { formula: `+'PL Data'!B${niRow}` }, { fill: FILL.cap });
  put('A10', { formula: '"Adjustments to reconcile "&LOWER(A9)' }, { gen: true, align: 'left' });

  const opLabels = {
    12: 'Purchases of investments', 13: 'Proceeds from sale of investments',
    14: 'Payment-in-kind interest', 15: 'Return of Capital from investment',
    16: 'Accretion of discount on debt securities', 17: 'Amortization of Deferred Financing Costs',
    18: 'Net realized (gain) loss on portfolio investments',
    19: 'Net unrealized (appreciation) depreciation on portfolio investment',
  };
  for (const [r, label] of Object.entries(opLabels)) {
    put('A' + r, label, { gen: true, align: 'left' });
    put('D' + r, { formula: `SUM(E${r}:Y${r})` });
  }
  // Investment adjustment inputs (gray). E12 purchases = residual of the investment
  // change after the explicit (non-cash) pieces and the waived-dev-fee financing leg.
  put('E12', { formula: 'E6-E13-E15-E16-E18-E19-E45' }, { fill: FILL.inv, border: { top: thin } });
  put('E13', 0, { fill: FILL.inv });
  put('E15', 0, { fill: FILL.inv, border: { bottom: thin } });
  put('E16', 0, { fill: FILL.inv });
  put('E18', 0, { fill: FILL.inv });
  put('E19', 0, { fill: FILL.inv });
  put('K17', 0, { fill: FILL.inv });

  // Changes in assets and liabilities
  put('A20', 'Changes in assets and liabilities:', { bold: true, gen: true, align: 'left' });
  put('D20', { formula: 'SUM(E20:Y20)' });
  // A-column dynamic labels (verbatim), D-column row totals, and the single input
  // cell per row that pulls the column's change (row 6).
  const chg = [
    [21, 'K', 'K$3', '=-K17+K6', 'gtA'], [22, 'F', 'F$3', '=F$6', 'ltA'], [23, 'G', 'G$3', '=G6', 'ltA'],
    [24, 'H', 'H$3', '=H$6', 'ltA'], [25, 'I', 'I$3', '=I$6', 'ltA'], [26, 'J', 'J$3', '=J$6', 'ltA'],
    [27, 'L', 'L$3', '=L$6', 'ltA'], [28, 'M', 'M$3', '=M$6', 'ltA'], [29, 'P', 'P$3', '=P$6', 'ltL'],
    [30, 'Q', 'Q$3', '=Q$6', 'ltL'], [31, 'R', 'R$3', '=R$6', 'ltL'], [32, 'S', 'S$3', '=S$6', 'ltL'],
    [33, 'T', 'T$3', '=T$6', 'ltL'], [34, 'U', 'U$3', '=U$6', 'ltL'], [35, 'V', 'V$3', '=V$6', 'ltL'],
    [36, 'W', 'W$3', '=W$6', 'gtL'], [37, 'X', 'X$3', '=X$6', 'ltL'],
  ];
  for (const [r, col, hdrRef, inputF, kind] of chg) {
    // Wording: assets use D{r}>0 ; liabilities use D{r}<0. The first/last rows keep
    // the vendor's PROPER-case variant.
    const cmp = (kind === 'gtA' || kind === 'ltA') ? '>' : '<';
    const decrCase = (kind === 'gtA') ? `LOWER(LEFT(${hdrRef},1))` : `LOWER(LEFT(${hdrRef},1))`;
    const incrCase = (kind === 'gtA' || kind === 'gtL') ? `PROPER(LEFT(${hdrRef},1))` : `LOWER(LEFT(${hdrRef},1))`;
    const f = `IF(D${r}${cmp}0,"Decrease in "&${decrCase}&MID(${hdrRef}, 2, LEN(${hdrRef})),"Increase in "&${incrCase}&MID(${hdrRef}, 2, LEN(${hdrRef})))`;
    put('A' + r, { formula: f }, { gen: true, align: 'left' });
    put('D' + r, { formula: `SUM(E${r}:Y${r})` });
    put(col + r, { formula: inputF.slice(1) }, { fill: fillForCol(col) });
  }

  // Net cash used in operating activities
  put('A39', 'Net cash used in operating activities', { bold: true, gen: true });
  for (const c of DATA_COLS) {
    put(c + '39', { formula: `SUM(${c}9:${c}38)` }, { bold: true, align: 'right',
      fill: c === 'Y' ? FILL.cap : undefined, border: { top: thin, bottom: thin } });
  }

  // Financing section
  put('A41', 'Cash flows from financing activities:', { bold: true, gen: true });
  const finRows = [
    [42, 'Capital contributions', 'Y', "+'Cap Activity'!B6"],
    [43, 'Capital call refunds', 'Y', "+'Cap Activity'!B7"],
    [44, 'Syndication costs', 'Y', "+'Cap Activity'!B8"],
    [45, 'Waived Development Fees', 'Y', "+'Cap Activity'!B9"],
    [46, 'Proceeds from line of credit', 'N', '=N$6-N47'],
    [47, 'Repayment of line of credit', 'N', 0],
    [48, 'Proceeds from notes payable', 'O', '=O$6-O49'],
    [49, 'Repayment of notes payable', 'O', 0],
  ];
  for (const [r, label, col, val] of finRows) {
    put('A' + r, label, { gen: true, align: 'left' });
    put('D' + r, { formula: `SUM(E${r}:Y${r})` });
    if (typeof val === 'number') put(col + r, val, { fill: fillForCol(col) });
    else put(col + r, { formula: val.slice(1) }, { fill: fillForCol(col) });
  }
  // Waived development fees are non-cash: the investment leg (E45) offsets the
  // financing contribution (Y45) so investment purchases exclude the dev fee.
  put('E45', { formula: '-Y45' }, { fill: FILL.inv });

  // Net cash provided by financing activities
  put('A51', 'Net cash provided by financing activities', { bold: true, gen: true });
  for (const c of DATA_COLS) {
    put(c + '51', { formula: `SUM(${c}42:${c}50)` }, { bold: true, align: 'right',
      fill: c === 'Y' ? FILL.cap : undefined, border: { top: thin, bottom: thin } });
  }

  // Net increase (decrease) in cash
  put('A53', 'Net increase (decrease) in cash', { bold: true, gen: true });
  for (const c of DATA_COLS) {
    put(c + '53', { formula: `${c}51+${c}39` }, { fill: c === 'Y' ? FILL.cap : undefined });
  }

  // TOTALS (s/b zero) — the tie-out row.
  put('A55', 'TOTALS (s/b zero)', { bold: true, gen: true });
  for (const c of DATA_COLS) {
    const f = (c === 'D') ? '+D53+D6' : `${c}53-${c}6`;
    put(c + '55', { formula: f }, { bold: true, fill: c === 'Y' ? FILL.cap : undefined, border: { top: thin, bottom: dbl } });
  }

  // Supplemental disclosure of noncash financing activities
  put('A58', 'Supplemental disclosure of noncash financing activities', { bold: true, gen: true });
  const notes = [
    [59, 'Capitiized Expense (non-cash: deferred development fee allocation)', '=-E45', 'non cash', 'reduces investment cash purchase'],
    [60, 'Non-Cash purchases (rollover investors)', null, 'non cash', 'reduces investment cash purchase & cash contribuitons'],
    [61, 'Contributions in Kind (deferred development costs', null, 'non cash', 'reduces cash contributions'],
    [62, 'GP reduction of commitment', null, 'total including cash portion', null],
    [63, 'GP return of capital (cash out)', null, 'cash out (included in contributions as a reduction)', null],
    [64, 'GP reduction of commitment - non cash', null, 'non cash', null],
  ];
  for (const [r, a, e, f, g] of notes) {
    put('A' + r, a, { gen: true, align: 'left' });
    if (e) put('E' + r, { formula: e.slice(1) });
    if (f) put('F' + r, f, { gen: true });
    if (g) put('G' + r, g, { gen: true });
  }

  return wb.xlsx.writeBuffer();
}

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
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
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

function registerCashFlowRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/cash-flow/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const asOf = body.quarter_end || '';
        if (!/^\d{4}-\d{2}-\d{2}$/.test(asOf)) throw new Error('quarter_end (YYYY-MM-DD) is required');
        const data = buildData(ctx, eid, asOf);
        const quarter = data.quarter;
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const buf = Buffer.from(await buildWorkbook(data));
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        const s = data.summary || {};
        res.setHeader('X-CashFlow-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          net_operating: s.netOperating, net_investing: s.netInvesting, net_financing: s.netFinancing,
          net_change: s.netChange, cash_end: s.cashEnd,
          contributions: data.fin.contributions, refunds: data.fin.refunds,
          syndication: data.fin.syndication, waived: data.fin.waived,
          bs_accounts: data.bsDetail.length, pl_accounts: data.plDetail.length,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { buildData, buildWorkbook, saveToWorkpapers, registerCashFlowRoutes, socfLineFor, COLS };
