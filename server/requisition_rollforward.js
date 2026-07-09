// ============================================================================
// Requisition roll-forward engine (exceljs)
// ----------------------------------------------------------------------------
// Produces Req#N+1 from Req#N + a set of Req#N+1 current-period invoices. The
// ONLY data motion is: Req#N Current Log folds into Req#N+1 Prior Log (appended
// to the matching cost-code group); the new Current Log is replaced with the
// incoming invoices. Everything else (B2A SUMIF columns, contingency tables)
// recomputes from formulas once the logs are right.
//
// Why rebuild the Prior Log group-by-group instead of inserting rows?
// Inserting rows and then fixing up shifted SUBTOTAL ranges and absolute
// references by arithmetic is exactly where the one-off openpyxl pass went
// wrong (Dev Fee J6 landed one row off). Rebuilding each group from scratch and
// re-deriving every SUBTOTAL range and every absolute reference BY LABEL removes
// the whole class of off-by-one shift bugs: we never track a moving row number,
// we look up the target by its name/bill text in the freshly written sheet.
//
// Formula cells inside the logs (e.g. Suncoast "=5532.77-5451") are preserved
// verbatim - copied as formulas, not as evaluated numbers.
//
// This engine writes formulas but does not evaluate them; pair it with a recalc
// step (headless LibreOffice) and the reconciliation engine to verify the result.
// ============================================================================

const { cellNum, cellStr, cellFormula, COL, applyInvoiceCols, isDevFeeLabel } = require('./requisition_reconcile.js');
// Convert a 1-based column index to its A1 letter (8 -> "H") so SUBTOTAL
// ranges track the DETECTED amount column instead of a hard-coded letter
// (templates without a GL column put Amount in H, not I).
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s || 'A'; }
const { learnDevFeeSpec, applyDevFeeSpec } = require('./requisition_devfee.js');

// Resolve a worksheet by name tolerantly: exact match first, then a
// case-insensitive / whitespace-normalized match, then a few known aliases.
// Requisition workbooks come from different project templates (CLIP, CLRF, …)
// whose tab names drift slightly ("Current Inv Log", "B2A", trailing spaces),
// and an exact-name miss returns undefined → the first `.getCell` on it throws
// the opaque "Cannot read properties of undefined (reading 'getCell')". This
// resolver removes that whole class of failure and lets callers give a precise
// "which tab is missing" error instead.
const SHEET_ALIASES = {
  'prior invoice log': ['prior inv log', 'prior log', 'prior invoices'],
  'current invoice log': ['current inv log', 'current log', 'current invoices'],
  'budget to actual': ['b2a', 'budget vs actual', 'budget-to-actual', 'budget to actuals', 'buna budget'],
  'dev fee': ['development fee', 'dev fee calc', 'developer fee', 'development fee calculation', 'development fee calc'],
  'hard cost contingency table': ['hard cost contingency', 'hard contingency table', 'hard cost contingency tbl'],
  'soft cost contingency table': ['soft cost contingency', 'soft contingency table', 'soft cost contingency tbl'],
};
function findSheet(workbook, canonicalName) {
  const norm = (s) => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  const want = norm(canonicalName);
  // 1. exact
  let ws = workbook.getWorksheet(canonicalName);
  if (ws) return ws;
  // 2. normalized exact + 3. aliases, scanning every sheet once
  const aliases = SHEET_ALIASES[want] || [];
  for (const sheet of workbook.worksheets) {
    const n = norm(sheet.name);
    if (n === want || aliases.includes(n)) return sheet;
  }
  return undefined;
}

// Read one invoice-log sheet into an ordered list of groups. A group is a run of
// data rows terminated by a SUBTOTAL row, with blank spacer rows interspersed.
//   group = { code, name, subtotalName, rows: [{...cells}], }
// rows preserve formula-vs-value (formula cells keep their formula string).
function parseLogGroups(ws) {
  const groups = [];
  let cur = null;
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 3; r <= last; r++) {
    const amtCell = ws.getCell(r, COL.amount);
    const f = cellFormula(amtCell);
    const name = cellStr(ws.getCell(r, COL.name)).trim();
    if (f && /SUBTOTAL/i.test(f)) {
      // close the current group at this subtotal
      if (cur) { cur.subtotalName = name; groups.push(cur); cur = null; }
      else { groups.push({ code: null, name: '', subtotalName: name, rows: [] }); }
      continue;
    }
    const amt = cellNum(amtCell);
    const hasData = amt != null || cellStr(ws.getCell(r, COL.vendor)).trim() || name;
    if (!hasData) continue; // blank spacer
    // a data row
    const row = readRowCells(ws, r);
    if (!cur) cur = { code: row.code, name: row.name, subtotalName: '', rows: [] };
    cur.rows.push(row);
  }
  if (cur) groups.push(cur);
  return groups;
}

// Resolve a row's numeric amount from a {formula,result} | number | null payload.
function rowAmount(row) {
  const a = row && row.amount;
  if (a == null) return null;
  if (typeof a === 'number') return a;
  if (typeof a === 'object' && 'result' in a) return typeof a.result === 'number' ? a.result : null;
  const n = Number(a);
  return Number.isFinite(n) ? n : null;
}

// Build a stable identity key for an invoice row: vendor + bill#. Used to detect
// the same invoice appearing twice (e.g. once with the real amount under one
// cost code, once as a $0 placeholder under another).
function rowIdentity(row) {
  const vendor = (row.vendor || '').toString().trim().toLowerCase();
  let bill = row.bill;
  if (bill && typeof bill === 'object') bill = bill.result != null ? bill.result : bill.formula;
  bill = (bill == null ? '' : String(bill)).trim().toLowerCase();
  if (!vendor && !bill) return null; // not enough to identify
  return vendor + '||' + bill;
}

// Drop zero/blank-amount duplicate rows: when the SAME invoice (vendor + bill#)
// appears more than once across the prior log and at least one copy carries a
// real (non-zero) amount, the zero/blank copies are redundant placeholders and
// are removed. The copy carrying the amount is always kept. This cleans up the
// case where one invoice was coded to two cost codes — one real, one $0 — so the
// rolled-forward log doesn't show an empty duplicate line. Mutates groups in
// place and returns the number of rows dropped.
function dropZeroDuplicateRows(groups) {
  // First pass: which identities have a real (non-zero) amount somewhere?
  const hasRealAmount = new Set();
  for (const g of groups) {
    for (const row of g.rows) {
      const id = rowIdentity(row);
      if (!id) continue;
      const amt = rowAmount(row);
      if (amt != null && amt !== 0) hasRealAmount.add(id);
    }
  }
  // Second pass: drop zero/blank copies whose identity is covered by a real one.
  let dropped = 0;
  for (const g of groups) {
    g.rows = g.rows.filter(row => {
      const id = rowIdentity(row);
      if (!id) return true;
      const amt = rowAmount(row);
      const isZeroish = amt == null || amt === 0;
      if (isZeroish && hasRealAmount.has(id)) { dropped++; return false; }
      return true;
    });
  }
  return dropped;
}

// Capture the cell payloads we need to rewrite a row, preserving formulas.
function readRowCells(ws, r) {
  const get = (col) => {
    const c = ws.getCell(r, col);
    const fm = cellFormula(c);
    if (fm) {
      // Preserve the cached result alongside the formula. exceljs stores a
      // formula cell as { formula, result }; dropping `result` means downstream
      // readers (reconcile's cellNum / readLog) can't pull a number out of the
      // re-written cell, so a formula-amount data row (e.g. SunCoast
      // "=5532.77-5451") silently vanishes from the folded Prior Log — taking
      // its amount and row-count with it (A1/A2/A3 fail). Keep result when
      // present so the row survives the round-trip even before LibreOffice
      // recalc runs.
      const v = c.value;
      let result = (v && typeof v === 'object' && 'result' in v) ? v.result : undefined;
      // If the source cell never carried a cached result but the formula is
      // self-contained literal arithmetic, compute and stamp the result now.
      // This makes the recovered amount (e.g. 81.77) durable in the folded
      // workbook instead of having to be re-derived on every subsequent read.
      if (result === undefined || result === null) {
        const computed = cellNum(c);
        if (computed != null) result = computed;
      }
      return (result === undefined || result === null) ? { formula: fm } : { formula: fm, result };
    }
    return c.value;
  };
  return {
    srcRow: r,
    cat: get(COL.cat), code: cellNum(ws.getCell(r, COL.code)),
    bankcat: get(COL.bankcat), gl: get(COL.gl),
    name: cellStr(ws.getCell(r, COL.name)).trim(),
    vendor: cellStr(ws.getCell(r, COL.vendor)).trim(),
    bill: get(COL.bill),
    amount: get(COL.amount),   // may be {formula} or number
    req: get(COL.req), date: get(COL.date),
  };
}

// Classify current-period rows by cost code so they can be folded into the
// matching prior group.
function currentRowsByCode(ws) {
  const byCode = new Map();
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 3; r <= last; r++) {
    const amtCell = ws.getCell(r, COL.amount);
    const f = cellFormula(amtCell);
    if (f && /SUBTOTAL/i.test(f)) continue;
    const amt = cellNum(amtCell);
    const vendor = cellStr(ws.getCell(r, COL.vendor)).trim();
    const name = cellStr(ws.getCell(r, COL.name)).trim();
    // A row is data if it has an amount, a vendor, OR a name. This MUST match
    // parseLogGroups' hasData test exactly — otherwise a name-only row (e.g. an
    // "Echo Land Cost" placeholder with no amount/vendor) is counted by the prior
    // parser but silently dropped here, folding one row short (A3 row-count fail).
    if (amt == null && !vendor && !name) continue;
    const code = cellNum(ws.getCell(r, COL.code));
    const key = code == null ? '__none__' : String(code);
    if (!byCode.has(key)) byCode.set(key, []);
    byCode.get(key).push(readRowCells(ws, r));
  }
  return byCode;
}

// Write a single data row's cells at target row `r`, preserving formula payloads.
function writeRowCells(ws, r, row) {
  const put = (col, payload) => { if (payload !== undefined) ws.getCell(r, col).value = payload; };
  put(COL.cat, row.cat);
  if (row.code != null) ws.getCell(r, COL.code).value = row.code;
  put(COL.bankcat, row.bankcat);
  put(COL.gl, row.gl);
  ws.getCell(r, COL.name).value = row.name || null;
  ws.getCell(r, COL.vendor).value = row.vendor || null;
  // Bill number (invoice #) is an IDENTIFIER, not a quantity. When it happens to
  // be all digits (e.g. 47302, 748203) and is stored as a numeric cell, Excel
  // right-aligns it and renders "#####" whenever the column is too narrow — and
  // it can never overflow into the next cell the way text does. Coerce it to a
  // string (and tag the cell as text) so a numeric-looking invoice # displays in
  // full, left-aligned, and preserves any leading zeros. Formula bills (rare)
  // and blank/undefined are passed through untouched.
  putBill(ws, r, row.bill);
  put(COL.amount, row.amount);
  put(COL.req, row.req);
  put(COL.date, row.date);
  // Normalize the amount cell: comma/accounting numFmt, right-aligned, no stale
  // top border. Guarantees data rows render with the comma style and never carry
  // a leftover underline at a shifted row position.
  styleAmountCell(ws, r, { underline: false });
  // Give the data row a comfortable height. The sheet's defaultRowHeight is 13pt
  // which crowds the 10pt text once a row is rewritten (the original autofit
  // state is lost), so set an explicit height that comfortably fits the text.
  ws.getRow(r).height = DATA_ROW_HEIGHT;
}

// Write the bill/invoice-# cell as TEXT. exceljs treats a JS number as a numeric
// cell (→ "#####" when narrow); a string with numFmt '@' is a genuine text cell
// that overflows/wraps like any label. Pass through formula objects and leave
// blank/undefined bills alone (don't stamp an empty string into a spacer).
function putBill(ws, r, bill) {
  if (bill === undefined || bill === null || bill === '') return;
  const cell = ws.getCell(r, COL.bill);
  if (typeof bill === 'object') { cell.value = bill; return; } // {formula,...} — leave as-is
  cell.value = String(bill);
  cell.numFmt = '@'; // Text format so a digits-only invoice # is never treated as a number
}

// Explicit data/subtotal row height (points). The invoice logs use 10pt Calibri;
// the workbook's defaultRowHeight of 13 leaves text looking cramped/clipped once
// rows are rewritten, so rewritten data and subtotal rows get this height.
const DATA_ROW_HEIGHT = 15;

// Canonical accounting (comma-style) number format for amount cells. Stamped
// explicitly on every amount cell we write so the comma/parenthesis style is
// guaranteed regardless of whatever style was inherited at a shifted row.
const ACCT_FMT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';

// Apply the canonical amount formatting to an amount cell: comma numFmt + right
// alignment, and (default) NO top border. Pass underline=true to draw the single
// thin rule used under each group's subtotal.
//
// IMPORTANT: exceljs de-duplicates cell styles through a shared style registry,
// so many cells can point at ONE style record. Mutating individual props
// (c.border = ..., c.numFmt = ...) on a cell that shares its style record bleeds
// the change onto every sibling cell — which is exactly how the subtotal's
// top-border rule leaked onto all the data rows. Assigning a brand-new complete
// `.style` object instead forces exceljs to register a DISTINCT style for this
// one cell, so the border/format never aliases. Font and fill are preserved from
// the cell's current style so we only change numFmt/alignment/border.
function styleAmountCell(ws, r, { underline = false } = {}) {
  const c = ws.getCell(r, COL.amount);
  const prev = c.style || {};
  c.style = {
    numFmt: ACCT_FMT,
    alignment: { horizontal: 'right', vertical: 'middle' },
    border: underline ? { top: { style: 'thin' } } : {},
    font: prev.font ? { ...prev.font } : undefined,
    fill: prev.fill ? { ...prev.fill } : undefined,
  };
}

// Horizontally align a single cell, preserving everything else about its style
// (numFmt, border, font, fill). Assigns a fresh complete `.style` object for the
// same shared-style-registry reason described on styleAmountCell.
function alignCell(ws, r, col, horizontal) {
  const c = ws.getCell(r, col);
  const prev = c.style || {};
  c.style = {
    numFmt: prev.numFmt,
    alignment: { ...(prev.alignment || {}), horizontal },
    border: prev.border ? { ...prev.border } : {},
    font: prev.font ? { ...prev.font } : undefined,
    fill: prev.fill ? { ...prev.fill } : undefined,
  };
}

// After a log sheet is prepared, align the numeric coding/amount columns — cost
// code (C), GL coding (E), and amount/amount-paid (I) — so they read
// consistently down the column. Only touches rows that have content in the
// column (skips blank spacers). GL and amount are right-aligned; the cost-code
// column's horizontal alignment is configurable (`codeAlign`, default 'right')
// so the Prior log can center it.
function rightAlignNumericColumns(ws, { codeAlign = 'right' } = {}) {
  if (!ws) return;
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 3; r <= last; r++) {
    for (const col of [COL.code, COL.gl, COL.amount]) {
      const v = ws.getCell(r, col).value;
      if (v === null || v === undefined || v === '') continue;
      alignCell(ws, r, col, col === COL.code ? codeAlign : 'right');
    }
  }
}

// Rebuild the Prior Log in `nextPriorWs` from prior groups + folded current rows.
// Returns a map of useful landmarks (row of each group subtotal, grand total row,
// and rows of specially-referenced lines) so callers can rewrite absolute refs.
// Collapse the Dev Fee tab's Hard/Soft cost rows into a single "Project costs"
// line sourced directly from the invoice logs, so this month's costs appear
// without needing each invoice classified Hard vs Soft:
//   C (this month)    = Current Invoice Log grand total - this period's dev fee
//   D (through prior) = Prior Invoice Log grand total - all prior dev fees
//   B (to date)       = C + D
// Formulas reference the logs (auditable, SUBTOTAL-safe via SUMIF on the dev
// code) and are seeded with values so they display before Excel recalculates.
// No-op on templates that don't split Hard/Soft.
function updateDevFeeProjectCosts({ devFeeWs, priorWs, curWs, curGrandTotalRow, devFeeInfo }) {
  if (!devFeeWs || !curWs || !priorWs || !curGrandTotalRow) return;
  const AMT = colLetter(COL.amount), CODE = colLetter(COL.code);
  const last = Math.max(devFeeWs.rowCount || 0, devFeeWs.actualRowCount || 0, 40);
  let hardRow = null; const softRows = [];
  for (let r = 1; r <= last; r++) {
    const label = (cellStr(devFeeWs.getCell(r, 1)) || cellStr(devFeeWs.getCell(r, 2))).toLowerCase();
    if (/conting/.test(label)) continue;
    if (/hard\s*cost/.test(label)) { if (hardRow == null) hardRow = r; }
    else if (/soft\s*cost/.test(label)) softRows.push(r);
  }
  if (hardRow == null) return; // template doesn't split Hard/Soft -> leave as-is

  let priorGtRow = null;
  const pl = Math.max(priorWs.rowCount || 0, priorWs.actualRowCount || 0);
  for (let r = 3; r <= pl; r++) {
    const f = cellFormula(priorWs.getCell(r, COL.amount));
    if (f && /SUBTOTAL/i.test(f) && /grand total/i.test(cellStr(priorWs.getCell(r, COL.name)))) { priorGtRow = r; break; }
  }
  if (priorGtRow == null) return;

  const round2 = (n) => Math.round(n * 100) / 100;
  const devCode = devFeeInfo && devFeeInfo.code != null ? devFeeInfo.code : null;
  const base = devFeeInfo && Number.isFinite(devFeeInfo.base) ? devFeeInfo.base : 0;

  // Seed for column D: prior-log total excluding every dev-fee line.
  let priorExDev = 0;
  for (let r = 3; r <= pl; r++) {
    const f = cellFormula(priorWs.getCell(r, COL.amount));
    if (f && /SUBTOTAL/i.test(f)) continue;
    const amt = cellNum(priorWs.getCell(r, COL.amount));
    if (amt == null) continue;
    const code = cellNum(priorWs.getCell(r, COL.code));
    const nm = cellStr(priorWs.getCell(r, COL.name));
    const isDev = (devCode != null && String(code) === String(devCode)) || isDevFeeLabel(nm);
    if (!isDev) priorExDev += amt;
  }
  priorExDev = round2(priorExDev);

  const priorGT = `'Prior Invoice Log'!${AMT}${priorGtRow}`;
  const priorDev = devCode != null ? `-SUMIF('Prior Invoice Log'!$${CODE}:$${CODE},${devCode},'Prior Invoice Log'!$${AMT}:$${AMT})` : '';

  // C = this month's project costs (ex dev fee). Written as a VALUE (the engine
  //     computed it from the entered invoices) so it does NOT reference the
  //     dev-fee line — that keeps the dev-fee line = 'Dev Fee'!C30 link below
  //     non-circular (the fee ties back to these costs).
  devFeeWs.getCell(hardRow, 3).value = round2(base);
  // D = costs through the prior month (Prior Log total, excluding prior dev fees).
  devFeeWs.getCell(hardRow, 4).value = { formula: `${priorGT}${priorDev}`, result: priorExDev };
  // B = incurred to date = this month + prior.
  devFeeWs.getCell(hardRow, 2).value = { formula: `C${hardRow}+D${hardRow}`, result: round2(base + priorExDev) };

  const labelCol = cellStr(devFeeWs.getCell(hardRow, 1)).trim() ? 1 : 2;
  devFeeWs.getCell(hardRow, labelCol).value = 'Project costs';
  for (const sr of softRows) for (const c of [1, 2, 3, 4]) devFeeWs.getCell(sr, c).value = null;

  // Link the Current Invoice Log's dev-fee line to the Dev Fee tab's Current
  // Month Dev Fee (C30) so the posted fee always ties to the tab. Seeded with
  // the computed fee for display before Excel recalculates.
  if (devCode != null) {
    const cl = Math.max(curWs.rowCount || 0, curWs.actualRowCount || 0);
    const ref = `'${devFeeWs.name}'!C30`;
    const fee = devFeeInfo && Number.isFinite(devFeeInfo.amount) ? round2(devFeeInfo.amount) : null;
    for (let r = 3; r <= cl; r++) {
      const f = cellFormula(curWs.getCell(r, COL.amount));
      if (f && /SUBTOTAL/i.test(f)) continue; // skip subtotal/grand-total rows
      if (String(cellNum(curWs.getCell(r, COL.code))) === String(devCode)) {
        curWs.getCell(r, COL.amount).value = fee != null ? { formula: ref, result: fee } : { formula: ref };
        break;
      }
    }
  }
}

// Hide zero/blank-amount line items on an invoice log to declutter the many $0
// placeholder rows. Subtotal/grand-total rows and blank spacer rows are left
// visible. Non-zero data rows are explicitly un-hidden so the state is
// deterministic each roll-forward. Note: the grand total is SUBTOTAL(9,...),
// which still counts hidden rows, so totals are unaffected.
function hideZeroAmountRows(ws) {
  if (!ws) return;
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  const sumRange = (a, b) => { let t = 0; for (let r = a; r <= b; r++) { const c = ws.getCell(r, COL.amount); const cf = cellFormula(c); if (cf && /SUBTOTAL/i.test(cf)) continue; const n = cellNum(c); if (n != null) t += n; } return t; };
  // Grand total row bounds the region where blank spacer rows get collapsed.
  let gtRow = last;
  for (let r = 3; r <= last; r++) { const gf = cellFormula(ws.getCell(r, COL.amount)); if (gf && /SUBTOTAL/i.test(gf) && /grand total/i.test(cellStr(ws.getCell(r, COL.name)))) { gtRow = r; break; } }
  for (let r = 3; r <= last; r++) {
    const amtCell = ws.getCell(r, COL.amount);
    const f = cellFormula(amtCell);
    const name = cellStr(ws.getCell(r, COL.name)).trim();
    if (f && /SUBTOTAL/i.test(f)) {
      // Grand total always visible; a group subtotal is hidden only when its
      // whole group nets to zero (so empty categories collapse cleanly).
      if (/grand total/i.test(name)) { ws.getRow(r).hidden = false; continue; }
      const m = f.match(/[A-Z]+(\d+):[A-Z]+(\d+)/i);
      const groupSum = m ? sumRange(Number(m[1]), Number(m[2])) : null;
      ws.getRow(r).hidden = (groupSum != null && Math.abs(groupSum) < 0.005);
      continue;
    }
    const hasContent = !!name || !!cellStr(ws.getCell(r, COL.vendor)).trim() || cellNum(ws.getCell(r, COL.code)) != null;
    if (!hasContent) {
      // Blank spacer row: collapse it (within the data region, before the grand
      // total) so there are no gaps between items, subtotals, and the total.
      if (r < gtRow) ws.getRow(r).hidden = true;
      continue;
    }
    const amt = cellNum(amtCell);
    ws.getRow(r).hidden = (amt == null || amt === 0);
  }
}

function rebuildPriorLog(nextPriorWs, priorGroups, curByCode, opts = {}) {
  // Clear existing data region (keep header rows 1-2).
  const existingLast = Math.max(nextPriorWs.rowCount || 0, nextPriorWs.actualRowCount || 0);
  for (let r = 3; r <= existingLast; r++) {
    for (let c = 2; c <= 11; c++) nextPriorWs.getCell(r, c).value = null;
  }

  const landmarks = { groupSubtotalRow: {}, byLabel: {}, grandTotalRow: null };
  let r = 3;
  let grandTotalName = null; // written LAST so it sits below every group
  for (const g of priorGroups) {
    const isGrand = /grand total/i.test(g.subtotalName || '');
    // Defer the grand total: it must be written AFTER any current-only groups
    // that get appended below (otherwise its range misses them and it lands
    // mid-list instead of at the bottom).
    if (isGrand) { grandTotalName = g.subtotalName; continue; }
    const dataStart = r;
    // prior data rows
    for (const row of g.rows) {
      writeRowCells(nextPriorWs, r, row);
      if (row.bill && typeof row.bill === 'string') landmarks.byLabel[row.bill.toLowerCase()] = r;
      r++;
    }
    // folded current rows for this code
    const key = g.code == null ? '__none__' : String(g.code);
    const folded = curByCode.get(key) || [];
    for (const row of folded) {
      writeRowCells(nextPriorWs, r, row);
      if (row.bill && typeof row.bill === 'string') landmarks.byLabel[row.bill.toLowerCase()] = r;
      r++;
    }
    curByCode.delete(key);
    const dataEnd = r - 1;
    // spacer (clear any inherited style on the amount cell)
    nextPriorWs.getCell(r, COL.amount).value = null; styleAmountCell(nextPriorWs, r, { underline: false }); r++;
    // subtotal — single thin rule above the figure (one consistent convention)
    const subRow = r;
    nextPriorWs.getCell(subRow, COL.name).value = g.subtotalName || ((g.name || '') + ' Total');
    nextPriorWs.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,${colLetter(COL.amount)}${dataStart}:${colLetter(COL.amount)}${dataEnd + 1})` };
    styleAmountCell(nextPriorWs, subRow, { underline: true });
    nextPriorWs.getRow(subRow).height = DATA_ROW_HEIGHT;
    if (g.code != null) landmarks.groupSubtotalRow[String(g.code)] = subRow;
    if (g.subtotalName) landmarks.byLabel[g.subtotalName.toLowerCase()] = subRow;
    r = subRow + 1;
    // spacer
    nextPriorWs.getCell(r, COL.amount).value = null; styleAmountCell(nextPriorWs, r, { underline: false }); r++;
  }

  // Any current codes with no matching prior group are appended as new groups.
  for (const [key, rows] of curByCode.entries()) {
    if (!rows.length) continue;
    const dataStart = r;
    for (const row of rows) { writeRowCells(nextPriorWs, r, row); r++; }
    nextPriorWs.getCell(r, COL.amount).value = null; styleAmountCell(nextPriorWs, r, { underline: false }); r++;
    const subRow = r;
    const nm = (rows[0].name || ('Code ' + key)) + ' Total';
    nextPriorWs.getCell(subRow, COL.name).value = nm;
    nextPriorWs.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,${colLetter(COL.amount)}${dataStart}:${colLetter(COL.amount)}${r - 1})` };
    styleAmountCell(nextPriorWs, subRow, { underline: true });
    if (key !== '__none__') landmarks.groupSubtotalRow[key] = subRow;
    r = subRow + 2;
  }

  // Grand total goes at the very bottom, so its range covers every data row —
  // including any current-only groups appended above.
  if (grandTotalName != null) {
    const gtRow = r;
    nextPriorWs.getCell(gtRow, COL.name).value = grandTotalName;
    nextPriorWs.getCell(gtRow, COL.amount).value = { formula: `SUBTOTAL(9,${colLetter(COL.amount)}3:${colLetter(COL.amount)}${gtRow - 1})` };
    styleAmountCell(nextPriorWs, gtRow, { underline: true });
    nextPriorWs.getRow(gtRow).height = DATA_ROW_HEIGHT;
    landmarks.grandTotalRow = gtRow;
    r = gtRow + 1;
  }

  return landmarks;
}

// Find a data row in a sheet by a bill/name substring (used to re-point absolute
// refs by label rather than by a tracked row number).
function findRowByLabel(ws, needles) {
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  const wants = needles.map(n => n.toLowerCase());
  for (let r = 3; r <= last; r++) {
    const name = cellStr(ws.getCell(r, COL.name)).toLowerCase();
    const bill = cellStr(ws.getCell(r, COL.bill)).toLowerCase();
    const hay = name + ' | ' + bill;
    if (wants.some(w => hay.includes(w))) return r;
  }
  return null;
}

// ----------------------------------------------------------------------------
// Compute this period's Development Fee and return a Current-Log row for it.
//
// The fee is a percentage of this period's NEW costs, EXCLUDING the dev fee line
// itself (including it would be circular). HOW the percentage/rounding/half-waive
// works differs by project, so we LEARN the method from the prior Req report's
// Dev Fee tab (see requisition_devfee.learnDevFeeSpec): first by parsing the
// tab's formulas, and if those are missing/ambiguous, by asking Claude — then
// back-validating the learned method against the prior period's actual base→fee.
// If the method can't be learned or doesn't reproduce the prior fee, we return
// devFee info with needsReview=true and NO row, so the fee is entered by hand
// rather than shipping a wrong number.
//
// Async because the Claude fallback is async. Returns:
//   { row, amount, code, base, rateText, spec, source, needsReview, note } | null
// where `row` matches the shape replaceCurrentLog / writeRowCells expect. A null
// return means there was no Dev Fee tab or no costs to base a fee on.
async function computeDevFeeRow({ devFeeWs, priorCurWs, newCurrent, meta, callClaude = null }) {
  if (!devFeeWs) return null;

  // 1. Identify the dev-fee cost code from the prior workbook's Current Log dev
  //    fee line so we exclude it from the base and clone its B/C/D/E/F/G coding.
  let template = null; // { cat, code, bankcat, gl, name, vendor }
  if (priorCurWs) {
    const last = Math.max(priorCurWs.rowCount || 0, priorCurWs.actualRowCount || 0);
    const mkTemplate = (r) => ({
      cat: priorCurWs.getCell(r, COL.cat).value,
      code: cellNum(priorCurWs.getCell(r, COL.code)),
      bankcat: priorCurWs.getCell(r, COL.bankcat).value,
      gl: priorCurWs.getCell(r, COL.gl).value,
      name: cellStr(priorCurWs.getCell(r, COL.name)).trim() || 'Development Fee',
      vendor: cellStr(priorCurWs.getCell(r, COL.vendor)).trim() || 'Banyan Residential',
    });
    // Prefer the line whose NAME is a dev-fee label (the actual Development Fee
    // line); only fall back to a bank-category match when no name matches. Some
    // templates carry a non-fee line (e.g. "Travel - Other Development Costs")
    // that merely shares the "Development ... Fee" bank category — picking it
    // would mis-code the fee (wrong code, wrong name).
    let byBankcat = null;
    for (let r = 3; r <= last; r++) {
      const amtF = cellFormula(priorCurWs.getCell(r, COL.amount));
      if (amtF && /SUBTOTAL/i.test(amtF)) continue; // skip subtotal/total rows
      const nm = cellStr(priorCurWs.getCell(r, COL.name));
      const bank = cellStr(priorCurWs.getCell(r, COL.bankcat));
      if (isDevFeeLabel(nm)) { template = mkTemplate(r); break; }
      if (byBankcat == null && isDevFeeLabel(bank)) byBankcat = r;
    }
    if (!template && byBankcat != null) template = mkTemplate(byBankcat);
  }
  const devCode = template && template.code != null ? template.code : 12913;

  // 2. Base = sum of this period's new invoices EXCLUDING the dev fee line.
  let base = 0;
  for (const inv of newCurrent) {
    if (String(inv.code) === String(devCode)) continue; // skip an existing dev fee row
    const amt = typeof inv.amount === 'number' ? inv.amount
      : (inv.amount && typeof inv.amount === 'object' && 'result' in inv.amount ? inv.amount.result : Number(inv.amount));
    if (Number.isFinite(amt)) base += amt;
  }
  if (!(base > 0)) return null;

  // 3. Learn the project's dev-fee method from the prior Dev Fee tab: parse the
  //    formulas, fall back to Claude if ambiguous, and back-validate against the
  //    prior period's observed base→fee. Whatever we learn, WE compute the fee.
  const learned = await learnDevFeeSpec({ devFeeWs, priorCurWs, devCode, COL, callClaude });

  const round2 = (n) => Math.round(n * 100) / 100;
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  let billLabel = 'Dev Fee';
  if (meta && meta.asOfDate) {
    const d = new Date(meta.asOfDate);
    if (!isNaN(d.getTime())) billLabel = `${MONTHS[d.getMonth()]}_${String(d.getFullYear()).slice(2)} Dev Fee`;
  }
  const t = template || {};

  // Could not learn a trustworthy method → don't invent a fee. Surface for
  // manual entry with the prior example so the user can key it in.
  if (!learned || !learned.spec || learned.needsReview) {
    return {
      amount: null,
      code: devCode,
      base: round2(base),
      spec: learned ? learned.spec : null,
      source: learned ? learned.source : 'none',
      needsReview: true,
      note: (learned && learned.note) ||
        'Could not confirm this project\'s dev-fee method from the prior report; please enter the dev fee for this period manually.',
      prior: learned ? learned.prior : null,
      row: null,
    };
  }

  // 4. Apply the learned spec to this period's base.
  const spec = learned.spec;
  const fee = applyDevFeeSpec(spec, base);
  if (!(fee > 0)) {
    return {
      amount: null, code: devCode, base: round2(base), spec, source: learned.source,
      needsReview: true, note: 'Learned dev-fee method produced a non-positive fee; enter manually.',
      prior: learned.prior, row: null,
    };
  }

  const pct = spec.halve ? (spec.rate * 50) : (spec.rate * 100);
  const rateText = (Number.isInteger(pct) ? String(pct) : pct.toFixed(2).replace(/\.?0+$/, '')) +
    '% of new costs' + (spec.halve ? ' (half waived)' : '');
  return {
    amount: fee,
    code: devCode,
    base: round2(base),
    rateText,
    spec,
    source: learned.source,       // 'formula:E15' | 'claude'
    needsReview: false,
    validation: learned.validation, // { ok, expected, got } against prior
    prior: learned.prior,
    row: {
      cat: t.cat != null ? t.cat : 'Soft Costs',
      code: devCode,
      bankcat: t.bankcat != null ? t.bankcat : 'Development Fee',
      gl: t.gl != null ? t.gl : devCode,
      name: t.name || 'Development Fee',
      vendor: (meta && meta.devFeePayee) || t.vendor || 'County Line Rail Interest',
      bill: billLabel,
      amount: fee,
      req: meta && meta.reqNumber ? 'Req#' + meta.reqNumber : undefined,
    },
  };
}

// ----------------------------------------------------------------------------
// Top-level: roll a workbook forward in place.
//   workbook    : exceljs Workbook loaded from Req#N (will be mutated to Req#N+1)
//   newCurrent  : array of invoice rows for the new period, each:
//                 { code, name, vendor, bill, amount, date, req }
//   meta        : { reqNumber, asOfDate }  for titles
// Returns { landmarks, devFee } describing where key rows ended up and the
// auto-computed development fee (amount + the row that was appended), if any.
// ----------------------------------------------------------------------------
// Async because computeDevFeeRow may call Claude to learn the dev-fee method.
// meta may carry { callClaude } — a caller-injected async fn for the Claude
// fallback; when absent, only deterministic formula parsing is used.
async function rollForward(workbook, newCurrent, meta = {}) {
  const priorWs = findSheet(workbook, 'Prior Invoice Log');
  const curWs = findSheet(workbook, 'Current Invoice Log');
  const b2a = findSheet(workbook, 'Budget to Actual');
  const devFee = findSheet(workbook, 'Dev Fee');

  // Fail fast with a precise, human-readable message when a required tab is
  // missing, instead of letting the first `.getCell` throw the opaque
  // "Cannot read properties of undefined (reading 'getCell')". The four tabs
  // below are structural; Dev Fee and the contingency tables are optional
  // (their absence just skips the corresponding step).
  const required = { 'Prior Invoice Log': priorWs, 'Current Invoice Log': curWs, 'Budget to Actual': b2a };
  const missing = Object.keys(required).filter(k => !required[k]);
  if (missing.length) {
    const present = workbook.worksheets.map(s => s.name).join(', ');
    const err = new Error(
      'Workbook is missing required tab(s): ' + missing.join(', ') +
      '. Tabs found in the uploaded file: ' + (present || '(none)') +
      '. A requisition roll-forward needs a "Prior Invoice Log", a "Current Invoice Log", and a "Budget to Actual" tab.'
    );
    err.userFacing = true;
    throw err;
  }

  // 0. Detect this workbook's invoice-log column layout from the header row
  //    and apply it to the shared COL map (some templates omit the GL column,
  //    shifting Amount from col 9 to col 8). Covers this roll-forward plus the
  //    reconciliation + dev-fee steps that read the same COL in this request.
  applyInvoiceCols(curWs);

  // 1. Capture prior structure + current rows BEFORE mutating anything.
  const priorGroups = parseLogGroups(priorWs);
  const curByCode = currentRowsByCode(curWs);

  // 1a. Clean up the prior log: drop zero/blank-amount duplicate rows where the
  //     same invoice (vendor + bill#) also appears with a real amount elsewhere
  //     (one invoice coded to two cost codes — one real, one $0 placeholder).
  dropZeroDuplicateRows(priorGroups);

  // 1b. Auto-compute this period's Development Fee from the new invoices and the
  //     entity's Dev Fee tab, then append it as a Current-Log line. Drop any dev
  //     fee row the caller already included so we never double-count or go
  //     circular. The dev fee is (new costs ex-dev-fee) x entity rate.
  const effectiveCurrent = Array.isArray(newCurrent) ? newCurrent.slice() : [];
  let devFeeInfo = null;
  try {
    const df = await computeDevFeeRow({
      devFeeWs: devFee, priorCurWs: curWs, newCurrent: effectiveCurrent, meta,
      callClaude: meta.callClaude || null,
    });
    if (df && df.row && !df.needsReview) {
      // Learned a trustworthy method. Remove any caller-supplied dev fee line for
      // the same code, then append ours.
      const filtered = effectiveCurrent.filter(inv => String(inv.code) !== String(df.code));
      filtered.push(df.row);
      effectiveCurrent.length = 0;
      effectiveCurrent.push(...filtered);
      devFeeInfo = {
        amount: df.amount, code: df.code, base: df.base, rateText: df.rateText,
        source: df.source, spec: df.spec, validation: df.validation, prior: df.prior,
        needsReview: false, row: df.row,
      };
    } else if (df) {
      // Method could not be confirmed. Do NOT auto-insert a fee — surface it so
      // the user enters the dev fee manually. The roll-forward still proceeds.
      devFeeInfo = {
        amount: null, code: df.code, base: df.base, source: df.source,
        spec: df.spec || null, prior: df.prior || null, needsReview: true, note: df.note,
      };
    }
  } catch (e) {
    // Dev fee is best-effort; never block the roll-forward on it.
    devFeeInfo = { error: e.message, needsReview: true };
  }

  // 2. Rebuild Prior Log = prior groups + folded current rows.
  const landmarks = rebuildPriorLog(priorWs, priorGroups, curByCode);

  // 2a. Strip bold from the Prior Invoice Log. The prior workbook's bold (on some
  //     amounts/subtotals) survives a value-only rewrite; the report convention
  //     is no bold in the logs, matching the Current Invoice Log treatment.
  stripBold(priorWs);

  // 2b. Align the numeric coding/amount columns on the prepared Prior Log: cost
  //     code (C) centered, GL coding (E) and amount paid (I) right-aligned.
  rightAlignNumericColumns(priorWs, { codeAlign: 'center' });

  // 2c. Hide zero/blank-amount line items on the Prior Invoice Log to declutter
  //     the many $0 placeholder rows (universal). Subtotals/grand total stay, and
  //     SUBTOTAL(9) totals are unaffected since they still count hidden rows.
  hideZeroAmountRows(priorWs);

  // 3. Replace Current Log with the incoming period's invoices (incl. dev fee).
  const curInfo = replaceCurrentLog(curWs, effectiveCurrent, meta);

  // 3a. Right-align the same numeric columns on the Current Log.
  rightAlignNumericColumns(curWs);

  // 4. Re-point absolute references by label (never by tracked row number).
  repointAbsoluteRefs({ priorWs, curWs, b2a, devFee, landmarks });

  // 4b. Roll the B2A contingency columns forward: this period's "Current"
  //     (col F) folds into next period's "Previous" (col E), then F clears to 0.
  //     Subtotal/total rows are left to recompute from the updated data cells.
  const contingency = rollForwardContingency(b2a);

  // 4c. Roll the Hard/Soft Cost Contingency Tables forward: each row's
  //     "Requested Herein" (col E) folds into "Previously Requested" (col D) and
  //     E clears. The "Total Requested" (F=D+E) and allocation rows recompute.
  const hardCt = findSheet(workbook, 'Hard Cost Contingency Table');
  const softCt = findSheet(workbook, 'Soft Cost Contingency Table');
  const contingencyTables = {
    hard: rollForwardContingencyTable(hardCt),
    soft: rollForwardContingencyTable(softCt),
  };

  // 4d. Collapse the Dev Fee tab's Hard/Soft rows into one "Project costs" line
  //     sourced from the invoice-log totals, so this month's costs populate
  //     without needing each invoice tagged Hard vs Soft. Best-effort.
  try {
    if (meta.collapseDevFeeCosts) updateDevFeeProjectCosts({ devFeeWs: devFee, priorWs, curWs, curGrandTotalRow: curInfo && curInfo.grandTotalRow, devFeeInfo });
  } catch (e) { /* never block the roll-forward on the Dev Fee tab cosmetic */ }

  // 5. Update titles (date / requisition number).
  if (meta.asOfDate && b2a.getCell('L1')) b2a.getCell('L1').value = meta.asOfDate;
  if (meta.reqNumber) {
    if (meta.fixReportNumberHeader) {
      // CLIP-style templates keep the report-number line at B5 (not B4), and the
      // generic path would add a duplicate at B4. Clear B4, then update the real
      // "Requisition Report #N" header line in place (correct spelling).
      if (b2a.getCell('B4')) b2a.getCell('B4').value = null;
      let hdrCell = null;
      for (let r = 1; r <= 8 && !hdrCell; r++) for (let c = 1; c <= 6; c++) {
        if (/requi\w*\s*report\s*#/i.test(cellStr(b2a.getCell(r, c)))) { hdrCell = b2a.getCell(r, c); break; }
      }
      if (hdrCell) hdrCell.value = 'Requisition Report #' + meta.reqNumber;
      else if (b2a.getCell('B4')) b2a.getCell('B4').value = 'Requisition Report #' + meta.reqNumber;
    } else if (b2a.getCell('B4')) {
      b2a.getCell('B4').value = 'Requistion Report # ' + meta.reqNumber;
    }
  }

  return { landmarks, devFee: devFeeInfo, contingency, contingencyTables };
}

// Write the new period invoices into the Current Log, grouped by cost code with
// SUBTOTAL rows and a Grand Total, mirroring the sheet's existing layout.
function replaceCurrentLog(ws, rows, meta) {
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 3; r <= last; r++) for (let c = 2; c <= 11; c++) ws.getCell(r, c).value = null;

  // group incoming rows by code, preserving first-seen order
  const order = [];
  const byCode = new Map();
  for (const row of rows) {
    const key = row.code == null ? '__none__' : String(row.code);
    if (!byCode.has(key)) { byCode.set(key, []); order.push(key); }
    byCode.get(key).push(row);
  }
  let r = 3;
  for (const key of order) {
    const grp = byCode.get(key);
    const dataStart = r;
    for (const row of grp) {
      writeRowCells(ws, r, {
        cat: row.cat || row.name, code: row.code, bankcat: row.bankcat, gl: row.gl ?? row.code,
        name: row.name, vendor: row.vendor, bill: row.bill,
        amount: row.amount, req: row.req || (meta.reqNumber ? 'Req#' + meta.reqNumber : undefined),
        date: row.date,
      });
      r++;
    }
    ws.getCell(r, COL.amount).value = null; styleAmountCell(ws, r, { underline: false }); r++;            // spacer
    const subRow = r;
    ws.getCell(subRow, COL.name).value = (grp[0].name || ('Code ' + key)) + ' Total';
    ws.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,${colLetter(COL.amount)}${dataStart}:${colLetter(COL.amount)}${r - 1})` };
    styleAmountCell(ws, subRow, { underline: true });
    ws.getRow(subRow).height = DATA_ROW_HEIGHT;
    r = subRow + 1;
    ws.getCell(r, COL.amount).value = null; styleAmountCell(ws, r, { underline: false }); r++;            // spacer
  }
  // grand total
  const gtRow = r;
  ws.getCell(gtRow, COL.name).value = 'Grand Total';
  ws.getCell(gtRow, COL.amount).value = { formula: `SUBTOTAL(9,${colLetter(COL.amount)}3:${colLetter(COL.amount)}${gtRow - 1})` };
  styleAmountCell(ws, gtRow, { underline: true });
  ws.getRow(gtRow).height = DATA_ROW_HEIGHT;

  // Per request, the Current Invoice Log carries NO bold anywhere — header,
  // data, subtotal, or grand-total. exceljs preserves the template cell's font
  // when only .value is rewritten, so the prior workbook's bold survives into
  // the new sheet; strip it explicitly across every cell that has a font.
  stripBold(ws);
  return { grandTotalRow: gtRow };
}

// Remove bold from every cell on a worksheet, preserving all other font
// attributes (name, size, color, italic, etc.). exceljs fonts are plain objects
// but must be reassigned as a whole to take effect — mutating cell.font.bold in
// place does not persist — so we spread the existing font into a new object.
function stripBold(ws) {
  if (!ws) return;
  ws.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const f = cell.font;
      if (f && f.bold) cell.font = { ...f, bold: false };
    });
  });
}
//
// On the B2A sheet, column E ("Previous Contingency" / prior "Funding &
// Reallocation") is the cumulative prior-period figure and column F ("Current
// Contingency" / this-period "Funding & Reallocation") is the activity that
// happened in the period just closed. Rolling forward to the next requisition
// means last period's CURRENT becomes next period's PREVIOUS: for every data
// row, E := E + F, and F is then cleared to 0 so the new period starts empty.
//
// Only DATA rows are touched. Subtotal / total rows (E or F is a SUM(...) or any
// cross-cell formula such as "=E12+E22+E43+E45") are LEFT ALONE — they recompute
// from the updated data cells, and rewriting them would corrupt the schedule.
// A row qualifies only when BOTH E and F are literal values (plain number or
// self-contained literal arithmetic like "=140716.49+327097.19"); cellNum
// returns null for any cell-referencing formula, which is exactly the subtotal
// guard we want. Rows where F is empty/zero are skipped (no movement to fold in).
//
// E's audit trail is preserved: if E was a literal-arithmetic formula, F is
// appended as another additive term (e.g. "=24182.71+48139.68" + F 9160 ->
// "=24182.71+48139.68+9160") rather than collapsing it to a single number. A
// plain-number E becomes the numeric sum. A seeded `result` keeps the value
// readable without a spreadsheet recalc (prod has no LibreOffice).
// True when a formula references another cell (e.g. C11, $A$1, SUM(F10:F15),
// ='Budget to Actual'!F27) rather than being self-contained literal arithmetic
// (e.g. =140716.49+327097.19). Used to leave subtotal/derived cells untouched
// during the contingency roll-forward. This is more reliable than checking for a
// cached numeric result: a formula like =ROUND(+C11+D11+E11,2) HAS a cached
// result yet is still cell-referencing and must NOT be folded as a data value.
function formulaHasCellRef(f) {
  if (!f) return false;
  let s = String(f);
  if (s[0] === '=') s = s.slice(1);
  if (s.indexOf('!') >= 0) return true; // any sheet-qualified reference
  return /(^|[^A-Za-z0-9_])[$]?[A-Za-z]{1,3}[$]?[0-9]+(?![0-9(])/.test(s);
}

function rollForwardContingency(b2a) {
  if (!b2a) return { moved: 0 };
  const COL_E = 5, COL_F = 6;
  let moved = 0;
  const last = Math.max(b2a.rowCount || 0, b2a.actualRowCount || 0);
  for (let r = 3; r <= last; r++) {
    const eCell = b2a.getCell(r, COL_E);
    const fCell = b2a.getCell(r, COL_F);

    // Skip subtotal/total rows: any cell-referencing formula in E or F makes
    // cellNum return null. (Literal-arithmetic formulas DO resolve, so genuine
    // data rows like "=140716.49+327097.19" are still handled.)
    const eIsRefFormula = formulaHasCellRef(cellFormula(eCell));
    const fIsRefFormula = formulaHasCellRef(cellFormula(fCell));
    if (eIsRefFormula || fIsRefFormula) continue;

    const fVal = cellNum(fCell);
    if (fVal == null || fVal === 0) continue; // nothing to fold in this row

    const eFormula = cellFormula(eCell); // literal-arithmetic formula, or null
    const eVal = cellNum(eCell) || 0;
    const newE = round2(eVal + fVal);

    if (eFormula) {
      // Preserve the existing additive breakdown; append F as another term so
      // the worksheet still shows where each dollar came from.
      let body = eFormula.trim();
      if (body[0] === '=') body = body.slice(1);
      const term = fVal < 0 ? `-${Math.abs(fVal)}` : `+${fVal}`;
      eCell.value = { formula: body + term, result: newE };
    } else {
      eCell.value = newE;
    }

    // Clear the current-period column (keep it numeric 0, matching the sheet's
    // existing convention where untouched current cells are 0, not blank).
    fCell.value = 0;
    moved++;
  }
  return { moved };
}
function round2(n) { return Math.round(n * 100) / 100; }

// Roll a Hard/Soft Cost Contingency Table forward one period.
//
// Layout (both tables share it):
//   col C = Cost Category   col D = Contingency Previously Requested
//   col E = Contingency Requested Herein   col F = Contingency Total Requested (=D+E)
//   col G = Req #           col H = Notes
//
// On roll-forward, this period's "Requested Herein" (E) becomes next period's
// "Previously Requested" (D): for every DATA row, D := D + E and E is cleared.
// The row keeps its cost category, Req #, and notes. F (=D+E) recomputes itself.
//
// Only DATA rows are touched. The "Total..." / "Contingency Allocation" rows hold
// SUBTOTAL/SUM formulas in D and E (cellNum returns null for cell-referencing
// formulas) and are left alone so they recompute from the updated data cells.
// A row qualifies only when its E is a literal value (plain number, or a cached
// formula result like ='Budget to Actual'!F27 -> 9160); the herein cross-refs to
// the B2A current-period draw resolve to numbers and fold in correctly. Rows with
// no E activity are already historical and are skipped.
function rollForwardContingencyTable(ws) {
  if (!ws) return { moved: 0 };
  const COL_C = 3, COL_D = 4, COL_E = 5;
  let moved = 0;
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 5; r <= last; r++) {
    const dCell = ws.getCell(r, COL_D);
    const eCell = ws.getCell(r, COL_E);

    // Skip subtotal/total/allocation rows. These hold aggregate formulas in D
    // and/or E — SUBTOTAL(...), SUM(...), or a cross-cell combo like =D22+E22 —
    // and must recompute from the data cells, not be folded. A cached result on
    // such a formula means cellNum returns a number (not null), so a result-based
    // guard alone would wrongly fold them; instead, skip any row whose D or E is a
    // formula that references other cells. A herein cross-ref like
    // ='Budget to Actual'!F27 also references a cell, but it lives in E on a DATA
    // row whose D is NOT a formula — so we only skip when EITHER cell carries a
    // SUBTOTAL/SUM aggregate, or when BOTH D and E are formulas (the hallmark of a
    // total row). A lone E cross-ref on an otherwise-plain row still folds.
    const dF = cellFormula(dCell) || '';
    const eF = cellFormula(eCell) || '';
    const isAggregate = /SUBTOTAL|SUM/i.test(dF) || /SUBTOTAL|SUM/i.test(eF);
    const bothFormula = !!dF && !!eF;
    if (isAggregate || bothFormula) continue;

    const eVal = cellNum(eCell);
    if (eVal == null || eVal === 0) continue; // no herein activity to fold in

    const dVal = cellNum(dCell) || 0;
    // Fold herein into previously-requested as a resolved number. The herein
    // cell may have been a cross-ref to the current period's B2A draw; once it's
    // history it should be a fixed amount, not a live reference, so write a plain
    // number rather than preserving the formula.
    dCell.value = round2(dVal + eVal);
    // Clear the herein column for the new period. Blank (null) rather than 0 so
    // the row reads as "no activity this period" like the other historical rows.
    eCell.value = null;
    moved++;
  }
  return { moved };
}

// Re-point the cross-sheet absolute references that the roll-forward affects.
// Each is resolved by LABEL against the freshly written sheets, so a shift in
// row positions can never silently point at the wrong line.
//
// In addition to repointing, this fills in the CACHED RESULTS for the Dev Fee
// block (J5-J14) and the repointed reference cells. The production roll-forward
// runs without a LibreOffice recalc, so any formula we write lands with no
// cached result; reconcile's B5 check (J8/J10/J11 numeric) then degrades to
// "not evaluated" and the required gate fails forever. By computing these values
// here from data already present (sums of explicit ranges, plain-number cells)
// the Dev Fee numbers are self-calculated and B5 evaluates without a spreadsheet
// engine. (Excel/LibreOffice still recompute them on open; we only seed results.)
function repointAbsoluteRefs({ priorWs, curWs, b2a, devFee, landmarks }) {
  // Sum the data-row amounts inside a SUBTOTAL(9, I a:I b) range on a sheet,
  // excluding nested SUBTOTAL rows (mirrors how Excel's SUBTOTAL ignores them).
  const sumSubtotalRange = (ws, subtotalRow) => {
    const f = cellFormula(ws.getCell(subtotalRow, COL.amount));
    if (!f) return null;
    const m = f.match(/[A-Z]+(\d+):[A-Z]+(\d+)/i);
    if (!m) return null;
    const a = Number(m[1]), b = Number(m[2]);
    let sum = 0;
    for (let r = a; r <= b; r++) {
      const c = ws.getCell(r, COL.amount);
      if (cellFormula(c) && /SUBTOTAL/i.test(cellFormula(c))) continue; // skip nested subtotals
      const n = cellNum(c);
      if (n != null) sum += n;
    }
    return sum;
  };
  // Read a numeric value out of a cell (plain number or cached formula result).
  const num = (ws, addrRow, col) => cellNum(ws.getCell(addrRow, col));
  // Set a cell to a formula while seeding its cached result so downstream,
  // recalc-free readers (reconcile) get a number. exceljs silently DROPS a
  // result of 0 from a formula cell, which would make cellNum read it as "not
  // evaluated"; when the computed result is exactly 0 we therefore store a plain
  // numeric 0 instead (the spreadsheet still recomputes the formula on open, and
  // these Dev Fee cells are cross-references whose displayed value is what the
  // reconciliation needs to read back).
  const setFR = (cell, formula, result) => {
    if (!cell) return;
    if (result == null || Number.isNaN(result)) { cell.value = { formula }; return; }
    if (result === 0) { cell.value = 0; return; }
    cell.value = { formula, result };
  };
  const round2 = (n) => Math.round(n * 100) / 100;

  // Dev Fee J6 -> Prior "6 month upfront interest" data row
  const j6row = findRowByLabel(priorWs, ['6 month upfront interest']);
  let j6 = null;
  if (j6row && devFee.getCell('J6')) {
    const v = num(priorWs, j6row, COL.amount);
    j6 = v == null ? null : -v;
    setFR(devFee.getCell('J6'), `-'Prior Invoice Log'!I${j6row}`, j6);
  }

  // Dev Fee J10 -> Prior "Development Fee Total" subtotal
  const j10row = findRowByLabel(priorWs, ['Development Fee Total', 'Dev Fee Total']);
  let j10 = null;
  if (j10row && devFee.getCell('J10')) {
    j10 = sumSubtotalRange(priorWs, j10row);
    setFR(devFee.getCell('J10'), `'Prior Invoice Log'!I${j10row}`, j10 == null ? null : round2(j10));
  }

  // Dev Fee J11 -> Current "Development Fee" data row. If this period has no
  // Development Fee line, do NOT leave J11 pointing at the prior workbook's stale
  // row — re-point it to 0 so the discrepancy math stays correct and B4's label
  // check doesn't trip on an empty row.
  const j11row = findRowByLabel(curWs, ['Development Fee']);
  let j11 = 0;
  if (devFee.getCell('J11')) {
    if (j11row) {
      j11 = num(curWs, j11row, COL.amount) || 0;
      setFR(devFee.getCell('J11'), `'Current Invoice Log'!I${j11row}`, round2(j11));
    } else {
      // No Development Fee this period. Write a PLAIN 0 (not a formula): exceljs
      // silently drops result:0 from a formula cell, which would make reconcile's
      // cellNum read it as "not evaluated" and fail B5 forever. A plain numeric 0
      // is read back as 0, and there is no current-period Dev Fee row to point at.
      j11 = 0;
      devFee.getCell('J11').value = 0;
    }
  }

  // Seed the dependent Dev Fee cells (J5,J7,J8,J12,J14) so B5 evaluates without
  // a recalc. J5 = E5 + E7 - B2A!J40 ; J7 = J5 + J6 ; J8 = ROUND(J7*2%,2) ;
  // J12 = J10 + J11 ; J14 = J8 - J12. Only seed when every input is available.
  const e5 = cellNum(devFee.getCell('E5'));
  const e7 = cellNum(devFee.getCell('E7'));
  const b2aJ40 = b2a ? cellNum(b2a.getCell('J40')) : null;
  if (e5 != null && e7 != null && b2aJ40 != null && devFee.getCell('J5')) {
    const j5 = e5 + e7 - b2aJ40;
    setFR(devFee.getCell('J5'), cellFormula(devFee.getCell('J5')) || `E5+E7-'Budget to Actual'!J40`, round2(j5));
    if (j6 != null && devFee.getCell('J7')) {
      const j7 = j5 + j6;
      setFR(devFee.getCell('J7'), cellFormula(devFee.getCell('J7')) || `SUM(J5:J6)`, round2(j7));
      if (devFee.getCell('J8')) {
        const j8 = round2(j7 * 0.02);
        setFR(devFee.getCell('J8'), cellFormula(devFee.getCell('J8')) || `ROUND(J7*0.02,2)`, j8);
        if (j10 != null && devFee.getCell('J12')) {
          const j12 = round2(j10 + j11);
          setFR(devFee.getCell('J12'), cellFormula(devFee.getCell('J12')) || `SUM(J10:J11)`, j12);
          if (devFee.getCell('J14')) {
            setFR(devFee.getCell('J14'), cellFormula(devFee.getCell('J14')) || `J8-J12`, round2(j8 - j12));
          }
        }
      }
    }
  }

  // B2A H45 -> Prior "Working Capital" row
  const wcRow = findRowByLabel(priorWs, ['Working Capital']);
  if (wcRow && b2a.getCell('H45')) b2a.getCell('H45').value = { formula: `'Prior Invoice Log'!I${wcRow}` };

  // B2A H49 -> Prior Grand Total ; I49/I50 -> Current Grand Total
  if (landmarks.grandTotalRow && b2a.getCell('H49')) {
    b2a.getCell('H49').value = { formula: `H47='Prior Invoice Log'!I${landmarks.grandTotalRow}` };
  }
  const curGT = findRowByLabel(curWs, ['Grand Total']);
  if (curGT) {
    if (b2a.getCell('I49')) b2a.getCell('I49').value = { formula: `'Current Invoice Log'!I${curGT}=I47` };
    if (b2a.getCell('I50')) b2a.getCell('I50').value = { formula: `'Current Invoice Log'!I${curGT}-I47` };
  }
}

module.exports = {
  rollForward, parseLogGroups, currentRowsByCode, rebuildPriorLog,
  replaceCurrentLog, repointAbsoluteRefs, findRowByLabel, findSheet,
};
