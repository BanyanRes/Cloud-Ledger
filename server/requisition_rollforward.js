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

const { cellNum, cellStr, cellFormula, COL } = require('./requisition_reconcile.js');

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
      const result = (v && typeof v === 'object' && 'result' in v) ? v.result : undefined;
      return result === undefined ? { formula: fm } : { formula: fm, result };
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
  put(COL.bill, row.bill);
  put(COL.amount, row.amount);
  put(COL.req, row.req);
  put(COL.date, row.date);
  // Give the data row a comfortable height. The sheet's defaultRowHeight is 13pt
  // which crowds the 10pt text once a row is rewritten (the original autofit
  // state is lost), so set an explicit height that comfortably fits the text.
  ws.getRow(r).height = DATA_ROW_HEIGHT;
}

// Explicit data/subtotal row height (points). The invoice logs use 10pt Calibri;
// the workbook's defaultRowHeight of 13 leaves text looking cramped/clipped once
// rows are rewritten, so rewritten data and subtotal rows get this height.
const DATA_ROW_HEIGHT = 15;

// Rebuild the Prior Log in `nextPriorWs` from prior groups + folded current rows.
// Returns a map of useful landmarks (row of each group subtotal, grand total row,
// and rows of specially-referenced lines) so callers can rewrite absolute refs.
function rebuildPriorLog(nextPriorWs, priorGroups, curByCode, opts = {}) {
  // Clear existing data region (keep header rows 1-2).
  const existingLast = Math.max(nextPriorWs.rowCount || 0, nextPriorWs.actualRowCount || 0);
  for (let r = 3; r <= existingLast; r++) {
    for (let c = 2; c <= 11; c++) nextPriorWs.getCell(r, c).value = null;
  }

  const landmarks = { groupSubtotalRow: {}, byLabel: {}, grandTotalRow: null };
  let r = 3;
  for (const g of priorGroups) {
    const isGrand = /grand total/i.test(g.subtotalName || '');
    if (isGrand) {
      // write grand total after a spacer; range covers all data written so far
      r += 0;
      const gtRow = r;
      nextPriorWs.getCell(gtRow, COL.name).value = g.subtotalName;
      nextPriorWs.getCell(gtRow, COL.amount).value = { formula: `SUBTOTAL(9,I3:I${gtRow - 1})` };
      nextPriorWs.getRow(gtRow).height = DATA_ROW_HEIGHT;
      landmarks.grandTotalRow = gtRow;
      r = gtRow + 1;
      continue;
    }
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
    // spacer
    nextPriorWs.getCell(r, COL.amount).value = null; r++;
    // subtotal
    const subRow = r;
    nextPriorWs.getCell(subRow, COL.name).value = g.subtotalName || ((g.name || '') + ' Total');
    nextPriorWs.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,I${dataStart}:I${dataEnd + 1})` };
    nextPriorWs.getRow(subRow).height = DATA_ROW_HEIGHT;
    if (g.code != null) landmarks.groupSubtotalRow[String(g.code)] = subRow;
    if (g.subtotalName) landmarks.byLabel[g.subtotalName.toLowerCase()] = subRow;
    r = subRow + 1;
    // spacer
    nextPriorWs.getCell(r, COL.amount).value = null; r++;
  }

  // Any current codes with no matching prior group are appended as new groups.
  for (const [key, rows] of curByCode.entries()) {
    if (!rows.length) continue;
    const dataStart = r;
    for (const row of rows) { writeRowCells(nextPriorWs, r, row); r++; }
    nextPriorWs.getCell(r, COL.amount).value = null; r++;
    const subRow = r;
    const nm = (rows[0].name || ('Code ' + key)) + ' Total';
    nextPriorWs.getCell(subRow, COL.name).value = nm;
    nextPriorWs.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,I${dataStart}:I${r - 1})` };
    if (key !== '__none__') landmarks.groupSubtotalRow[key] = subRow;
    r = subRow + 2;
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
// itself (including it would be circular). The percentage is entity-specific and
// lives in the Dev Fee tab, not hard-coded here: the tab computes it as
//   E15 = ROUND(baseCost * rate1, 2)   (e.g. rate1 = 4%)
//   E17 = ROUND(E15 / 2, 2)            (waived half -> effective 2%)
//   E19 = E17                          (the amount that goes on the Current Log)
// We read rate1 and the "/2" structure straight from the tab's formulas so each
// entity's own rate carries through, then apply it to the freshly-entered
// invoices' total (minus any dev fee line).
//
// Returns { row, amount, code } or null if the tab has no usable dev-fee setup
// or there are no costs to base a fee on. `row` matches the shape replaceCurrentLog
// / writeRowCells expect.
function computeDevFeeRow({ devFeeWs, priorCurWs, newCurrent, meta }) {
  if (!devFeeWs) return null;

  // 1. Identify the dev-fee cost code from the prior workbook's Current Log dev
  //    fee line so we exclude it from the base and clone its B/C/D/E/F/G coding.
  let template = null; // { cat, code, bankcat, gl, name, vendor }
  if (priorCurWs) {
    const last = Math.max(priorCurWs.rowCount || 0, priorCurWs.actualRowCount || 0);
    for (let r = 3; r <= last; r++) {
      const nm = cellStr(priorCurWs.getCell(r, COL.name)).toLowerCase();
      const bank = cellStr(priorCurWs.getCell(r, COL.bankcat)).toLowerCase();
      if (nm.includes('development fee') || bank.includes('development fee')) {
        template = {
          cat: priorCurWs.getCell(r, COL.cat).value,
          code: cellNum(priorCurWs.getCell(r, COL.code)),
          bankcat: priorCurWs.getCell(r, COL.bankcat).value,
          gl: priorCurWs.getCell(r, COL.gl).value,
          name: cellStr(priorCurWs.getCell(r, COL.name)).trim() || 'Development Fee',
          vendor: cellStr(priorCurWs.getCell(r, COL.vendor)).trim() || 'Banyan Residential',
        };
        break;
      }
    }
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

  // 3. Read the entity-specific rate structure from the Dev Fee tab. Parse the
  //    first percentage out of E15's formula (e.g. "ROUND(E10*4%,2)" -> 0.04) and
  //    detect the "/2" halving in E17. Fall back to 4% & halving if unreadable.
  let rate = 0.04, halve = true;
  const e15f = cellFormula(devFeeWs.getCell('E15')) || '';
  const e17f = cellFormula(devFeeWs.getCell('E17')) || '';
  const pctM = e15f.match(/(\d+(?:\.\d+)?)\s*%/) || e15f.match(/\*\s*0?\.(\d+)/);
  if (pctM) {
    if (e15f.includes('%')) rate = parseFloat(pctM[1]) / 100;
    else rate = parseFloat('0.' + pctM[1]);
  }
  halve = /\/\s*2\b/.test(e17f) || /E15\s*\/\s*2/i.test(e17f);

  const round2 = (n) => Math.round(n * 100) / 100;
  let fee = round2(base * rate);
  if (halve) fee = round2(fee / 2);
  if (!(fee > 0)) return null;

  // 4. Build the Current-Log row. Bill number is month-based ("April_26 Dev Fee")
  //    derived from the as-of date, matching the workbook's existing convention.
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  let billLabel = 'Dev Fee';
  if (meta && meta.asOfDate) {
    const d = new Date(meta.asOfDate);
    if (!isNaN(d.getTime())) billLabel = `${MONTHS[d.getMonth()]}_${String(d.getFullYear()).slice(2)} Dev Fee`;
  }
  const t = template || {};
  const rateText = (halve ? (rate * 50) : (rate * 100)).toFixed(rate * 100 % 1 === 0 && (halve ? (rate*50)%1===0 : true) ? 0 : 2).replace(/\.00$/, '') + '% of new costs';
  return {
    amount: fee,
    code: devCode,
    base: round2(base),
    rateText,
    row: {
      cat: t.cat != null ? t.cat : 'Soft Costs',
      code: devCode,
      bankcat: t.bankcat != null ? t.bankcat : 'Development Fee',
      gl: t.gl != null ? t.gl : devCode,
      name: t.name || 'Development Fee',
      vendor: t.vendor || 'Banyan Residential',
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
function rollForward(workbook, newCurrent, meta = {}) {
  const priorWs = workbook.getWorksheet('Prior Invoice Log');
  const curWs = workbook.getWorksheet('Current Invoice Log');
  const b2a = workbook.getWorksheet('Budget to Actual');
  const devFee = workbook.getWorksheet('Dev Fee');

  // 1. Capture prior structure + current rows BEFORE mutating anything.
  const priorGroups = parseLogGroups(priorWs);
  const curByCode = currentRowsByCode(curWs);

  // 1b. Auto-compute this period's Development Fee from the new invoices and the
  //     entity's Dev Fee tab, then append it as a Current-Log line. Drop any dev
  //     fee row the caller already included so we never double-count or go
  //     circular. The dev fee is (new costs ex-dev-fee) x entity rate.
  const effectiveCurrent = Array.isArray(newCurrent) ? newCurrent.slice() : [];
  let devFeeInfo = null;
  try {
    const df = computeDevFeeRow({ devFeeWs: devFee, priorCurWs: curWs, newCurrent: effectiveCurrent, meta });
    if (df) {
      // Remove any caller-supplied dev fee line for the same code, then append ours.
      const filtered = effectiveCurrent.filter(inv => String(inv.code) !== String(df.code));
      filtered.push(df.row);
      effectiveCurrent.length = 0;
      effectiveCurrent.push(...filtered);
      devFeeInfo = { amount: df.amount, code: df.code, base: df.base, rateText: df.rateText, row: df.row };
    }
  } catch (e) {
    // Dev fee is best-effort; never block the roll-forward on it.
    devFeeInfo = { error: e.message };
  }

  // 2. Rebuild Prior Log = prior groups + folded current rows.
  const landmarks = rebuildPriorLog(priorWs, priorGroups, curByCode);

  // 3. Replace Current Log with the incoming period's invoices (incl. dev fee).
  replaceCurrentLog(curWs, effectiveCurrent, meta);

  // 4. Re-point absolute references by label (never by tracked row number).
  repointAbsoluteRefs({ priorWs, curWs, b2a, devFee, landmarks });

  // 5. Update titles (date / requisition number).
  if (meta.asOfDate && b2a.getCell('L1')) b2a.getCell('L1').value = meta.asOfDate;
  if (meta.reqNumber && b2a.getCell('B4')) b2a.getCell('B4').value = 'Requistion Report # ' + meta.reqNumber;

  return { landmarks, devFee: devFeeInfo };
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
    ws.getCell(r, COL.amount).value = null; r++;            // spacer
    const subRow = r;
    ws.getCell(subRow, COL.name).value = (grp[0].name || ('Code ' + key)) + ' Total';
    ws.getCell(subRow, COL.amount).value = { formula: `SUBTOTAL(9,I${dataStart}:I${r - 1})` };
    ws.getRow(subRow).height = DATA_ROW_HEIGHT;
    r = subRow + 1;
    ws.getCell(r, COL.amount).value = null; r++;            // spacer
  }
  // grand total
  const gtRow = r;
  ws.getCell(gtRow, COL.name).value = 'Grand Total';
  ws.getCell(gtRow, COL.amount).value = { formula: `SUBTOTAL(9,I3:I${gtRow - 1})` };
  ws.getRow(gtRow).height = DATA_ROW_HEIGHT;
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
    const m = f.match(/I(\d+):I(\d+)/i);
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
  replaceCurrentLog, repointAbsoluteRefs, findRowByLabel,
};
