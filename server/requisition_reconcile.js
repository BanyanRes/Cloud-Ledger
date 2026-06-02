// ============================================================================
// Requisition roll-forward reconciliation engine
// ----------------------------------------------------------------------------
// A roll-forward only MOVES data: Req#N+1 Prior Log = Req#N Prior + Req#N Current.
// No new amounts are introduced, so any discrepancy is a MECHANICAL error
// (a dropped cell, a shifted absolute reference, a stale SUBTOTAL range) - never
// a data/judgement error. These checks prove the move was exact; a FAIL is a bug
// to diagnose + fix, then re-verify until PASS.
//
// Checks (all deterministic arithmetic, no API):
//   A1  Prior total      : new.Prior.total == old.Prior.total + old.Current.total
//   A2  Per-cost-code     : for every code, new.Prior[code] == old.Prior[code] + old.Current[code]
//   A3  Row count         : new.Prior.rows == old.Prior.rows + old.Current.rows
//   A4  Grand Total cell  : Prior Log grand-total SUBTOTAL == A1 sum
//   B1  Group subtotals   : each group SUBTOTAL == sum of its data rows
//   B4  Absolute refs     : Dev Fee J6/J10/J11 hit the intended row (by label/value)
//   B5  Dev Fee discrep.  : Dev Fee J14 == 0  (J8 == J10 + J11)
//   C1  This-period tie   : B2A this-period column total == Current Log grand total
//
// Each check returns { id, level, pass, expected, actual, delta, detail }.
// level: 'required' | 'recommended'. The caller decides whether 'recommended'
// failures block. A FAIL bundle is what gets handed to the Claude API for
// diagnosis when (and only when) something does not reconcile.
// ============================================================================

const TOL = 0.01; // dollar tolerance for float/rounding noise

// ----- cell helpers ---------------------------------------------------------
// exceljs: plain cell -> .value is the value; formula cell -> .value is
// { formula, result }. Pull a number out of either shape.
function cellNum(cell) {
  if (!cell) return null;
  let v = cell.value;
  if (v && typeof v === 'object' && 'result' in v) v = v.result;
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
  return null;
}
function cellStr(cell) {
  if (!cell) return '';
  let v = cell.value;
  if (v && typeof v === 'object' && 'result' in v) v = v.result;
  if (v && typeof v === 'object' && 'richText' in v) {
    return v.richText.map(t => t.text).join('');
  }
  return v == null ? '' : String(v);
}
function cellFormula(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v && typeof v === 'object' && 'formula' in v) return v.formula;
  return null;
}

// Column map for the invoice-log sheets (1-based, matching the SRN workbook):
//   B(2)=Cost Category C(3)=Cost Code# D(4)=Bank Cost Cat E(5)=GL Coding
//   F(6)=Cost Code Name G(7)=Vendor H(8)=Bill# I(9)=Amount J(10)=Req# K(11)=Inv Date
const COL = { cat: 2, code: 3, bankcat: 4, gl: 5, name: 6, vendor: 7, bill: 8, amount: 9, req: 10, date: 11 };

// ----- log reader -----------------------------------------------------------
// Walk an invoice-log worksheet and return:
//   rows: [{ row, code, name, vendor, bill, amount }]  (vendor data rows only)
//   subtotals: [{ row, name, formula, result }]        (SUBTOTAL rows)
//   total: sum of data-row amounts
//   byCode: { code: sum }
function readLog(ws) {
  const rows = [];
  const subtotals = [];
  const byCode = {};
  let total = 0;
  // NOTE: exceljs actualRowCount = COUNT of non-empty rows, not the max row
  // number. Spacer rows between groups make it smaller than the real extent,
  // which truncates the scan. Use rowCount (max row index) so we reach the
  // Grand Total and the late groups (Dev Fee, etc.).
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = 3; r <= last; r++) {
    const amtCell = ws.getCell(r, COL.amount);
    const f = cellFormula(amtCell);
    if (f && /SUBTOTAL/i.test(f)) {
      subtotals.push({ row: r, name: cellStr(ws.getCell(r, COL.name)), formula: f, result: cellNum(amtCell) });
      continue;
    }
    const vendor = cellStr(ws.getCell(r, COL.vendor)).trim();
    const amt = cellNum(amtCell);
    // A data row is any row carrying a numeric amount that is not a SUBTOTAL.
    // Most have a vendor, but capital lines (Land, Acquisitions, Working Capital)
    // are vendorless yet still real amounts that belong to their group totals.
    if (amt != null) {
      const code = cellNum(ws.getCell(r, COL.code));
      const name = cellStr(ws.getCell(r, COL.name)).trim();
      rows.push({
        row: r, code, name,
        vendor, bill: cellStr(ws.getCell(r, COL.bill)).trim(),
        amount: amt,
        req: cellStr(ws.getCell(r, COL.req)).trim(),
      });
      total += amt;
      const k = code == null ? '__none__' : String(code);
      byCode[k] = (byCode[k] || 0) + amt;
    } else if (f && (vendor || cellStr(ws.getCell(r, COL.name)).trim())) {
      // Defensive: a non-SUBTOTAL formula-amount row (e.g. "=5532.77-5451")
      // whose cached result is missing. cellNum returns null for it, which would
      // drop the row from the count and undercount A3. If it carries a vendor or
      // name it is real data — count the row so A3 stays exact. We can't trust an
      // un-recalculated amount, so contribute 0 to total/byCode; the row's value
      // is verified once the workbook is recalculated (A1/A2 with results present).
      const code = cellNum(ws.getCell(r, COL.code));
      const name = cellStr(ws.getCell(r, COL.name)).trim();
      rows.push({
        row: r, code, name,
        vendor, bill: cellStr(ws.getCell(r, COL.bill)).trim(),
        amount: 0,
        req: cellStr(ws.getCell(r, COL.req)).trim(),
        unevaluatedFormula: f,
      });
    }
  }
  return { rows, subtotals, total, byCode };
}

function approxEq(a, b, tol = TOL) {
  return Math.abs((a || 0) - (b || 0)) <= tol;
}
function chk(id, level, pass, expected, actual, detail) {
  return { id, level, pass, expected, actual, delta: (actual || 0) - (expected || 0), detail: detail || '' };
}
function round2(n) { return n == null ? n : Math.round(n * 100) / 100; }

// ----------------------------------------------------------------------------
// Core: reconcile a freshly rolled-forward workbook against the prior period.
//   prev = { prior: <ws>, current: <ws> }   (Req#N sheets)
//   next = { prior: <ws>, current: <ws>, b2a: <ws>, devFee: <ws> }  (Req#N+1)
// Sheets are exceljs worksheet objects. b2a/devFee optional (enable B/C checks).
// ----------------------------------------------------------------------------
function reconcile(prev, next, opts = {}) {
  const tol = opts.tol != null ? opts.tol : TOL;
  const checks = [];

  const oPrior = readLog(prev.prior);
  const oCurr = readLog(prev.current);
  const nPrior = readLog(next.prior);

  // A1 - Prior total accumulation
  {
    const expected = oPrior.total + oCurr.total;
    checks.push(chk('A1', 'required', approxEq(nPrior.total, expected, tol),
      round2(expected), round2(nPrior.total),
      'Req#N+1 Prior total == Req#N Prior + Req#N Current'));
  }

  // A2 - per cost code
  {
    const codes = new Set([
      ...Object.keys(oPrior.byCode), ...Object.keys(oCurr.byCode), ...Object.keys(nPrior.byCode),
    ]);
    const fails = [];
    for (const code of codes) {
      const expected = (oPrior.byCode[code] || 0) + (oCurr.byCode[code] || 0);
      const actual = nPrior.byCode[code] || 0;
      if (!approxEq(actual, expected, tol)) {
        fails.push({ code, expected: round2(expected), actual: round2(actual), delta: round2(actual - expected) });
      }
    }
    checks.push(chk('A2', 'required', fails.length === 0, 0, fails.length,
      fails.length ? 'Per-code mismatch: ' + JSON.stringify(fails) : 'all cost codes tie'));
  }

  // A3 - row count
  {
    const expected = oPrior.rows.length + oCurr.rows.length;
    checks.push(chk('A3', 'required', nPrior.rows.length === expected,
      expected, nPrior.rows.length, 'Prior row count == old Prior + old Current rows'));
  }

  // A4 - Grand Total cell on the new Prior Log.
  // The Grand Total is SUBTOTAL(9, I3:I<last>) spanning the whole column, which
  // INCLUDES the per-group subtotal rows. Excel's SUBTOTAL ignores other
  // SUBTOTAL cells in its range, so in Excel this equals the data total. But
  // some engines (e.g. a LibreOffice headless recalc) flatten it and double-count
  // the subtotals. So compare against BOTH the data total and the
  // data-plus-subtotals figure; pass if it matches the data total, and flag the
  // double-count case as a known recalc artifact rather than a real break.
  {
    const gt = nPrior.subtotals.find(s => /grand total/i.test(s.name));
    if (gt && gt.result != null) {
      const subSum = nPrior.subtotals
        .filter(s => !/grand total/i.test(s.name) && s.result != null)
        .reduce((a, s) => a + s.result, 0);
      const pass = approxEq(gt.result, nPrior.total, tol);
      const doubleCounted = approxEq(gt.result, nPrior.total + subSum, tol);
      checks.push(chk('A4', 'recommended', pass,
        round2(nPrior.total), round2(gt.result),
        'Prior Log Grand Total cell (row ' + gt.row + ')' +
        (doubleCounted ? ' - matches data+subtotals; nested-SUBTOTAL recalc artifact, not a data break' : '')));
    } else {
      checks.push(chk('A4', 'recommended', false, round2(nPrior.total), null,
        'Grand Total SUBTOTAL row not found or not evaluated (recalc may be needed)'));
    }
  }

  // B1 - each group SUBTOTAL ties to the data rows inside ITS OWN formula range.
  // Parse the explicit range from SUBTOTAL(9,I<a>:I<b>) and sum only the data
  // rows in [a,b] (nested subtotal rows are naturally excluded - they're not in
  // `rows`). Using the formula range, not a guessed boundary, is what makes this
  // correct when groups are adjacent (e.g. two "Closing Costs Total" blocks).
  {
    const log = nPrior;
    const fails = [];
    const byRow = new Map(log.rows.map(rw => [rw.row, rw.amount]));
    for (const st of log.subtotals) {
      if (/grand total/i.test(st.name)) continue;
      const m = st.formula.match(/I(\d+):I(\d+)/i);
      if (!m) continue;
      const a = Number(m[1]), b = Number(m[2]);
      let sum = 0;
      for (let r = a; r <= b; r++) if (byRow.has(r)) sum += byRow.get(r);
      if (st.result != null && !approxEq(st.result, sum, tol)) {
        fails.push({ subtotalRow: st.row, name: st.name, range: `${a}:${b}`, expected: round2(sum), actual: round2(st.result) });
      }
    }
    checks.push(chk('B1', 'required', fails.length === 0, 0, fails.length,
      fails.length ? 'next.prior subtotal mismatch: ' + JSON.stringify(fails) : 'group subtotals tie'));
  }

  // B4 / B5 - absolute references + dev fee discrepancy (needs devFee sheet)
  if (next.devFee) {
    const dv = next.devFee;
    const refChecks = [
      { cell: 'J6', wantLabelIn: ['6 month upfront interest'], sheet: 'prior' },
      { cell: 'J10', wantLabelIn: ['Development Fee Total', 'Dev Fee Total'], sheet: 'prior' },
      { cell: 'J11', wantLabelIn: ['Development Fee'], sheet: 'current' },
    ];
    const refFails = [];
    for (const rc of refChecks) {
      const f = cellFormula(dv.getCell(rc.cell));
      if (!f) { refFails.push({ cell: rc.cell, issue: 'no formula' }); continue; }
      const m = f.match(/I(\d+)/);
      if (!m) { refFails.push({ cell: rc.cell, issue: 'no row ref', formula: f }); continue; }
      const targetRow = Number(m[1]);
      const sheet = rc.sheet === 'prior' ? next.prior : next.current;
      // The intended label may live in the Cost Code Name (F) or the Bill# (H)
      // column - e.g. "6 month upfront interest" is a bill memo, not a name.
      const label = cellStr(sheet.getCell(targetRow, COL.name)).trim();
      const bill = cellStr(sheet.getCell(targetRow, COL.bill)).trim();
      const hay = (label + ' | ' + bill).toLowerCase();
      const ok = rc.wantLabelIn.some(w => hay.includes(w.toLowerCase()));
      if (!ok) refFails.push({ cell: rc.cell, formula: f, targetRow, foundLabel: label, foundBill: bill, want: rc.wantLabelIn });
    }
    checks.push(chk('B4', 'required', refFails.length === 0, 0, refFails.length,
      refFails.length ? 'Absolute-ref mismatch: ' + JSON.stringify(refFails) : 'dev-fee absolute refs resolve correctly'));

    const j8 = cellNum(dv.getCell('J8'));
    const j10 = cellNum(dv.getCell('J10'));
    const j11 = cellNum(dv.getCell('J11'));
    if (j8 != null && j10 != null && j11 != null) {
      const disc = j8 - (j10 + j11);
      checks.push(chk('B5', 'required', approxEq(disc, 0, tol), 0, round2(disc),
        'Dev Fee discrepancy J8 - (J10+J11); J8=' + round2(j8) + ' J10=' + round2(j10) + ' J11=' + round2(j11)));
    } else {
      checks.push(chk('B5', 'required', false, 0, null,
        'Dev Fee J8/J10/J11 not evaluated (recalc the workbook before checking)'));
    }
  }

  // C1 - this-period total ties to Current Log grand total
  if (next.current && next.b2a) {
    const nCurr = readLog(next.current);
    const gt = nCurr.subtotals.find(s => /grand total/i.test(s.name));
    const curTotal = gt && gt.result != null ? gt.result : nCurr.total;
    let projAllI = null;
    const b2a = next.b2a;
    const bLast = b2a.actualRowCount || b2a.rowCount;
    for (let r = 1; r <= bLast; r++) {
      if (/project costs\s*-?\s*all/i.test(cellStr(b2a.getCell(r, 2)))) {
        projAllI = cellNum(b2a.getCell(r, COL.amount)); break;
      }
    }
    if (projAllI != null) {
      checks.push(chk('C1', 'recommended', approxEq(projAllI, curTotal, tol),
        round2(curTotal), round2(projAllI), 'B2A this-period Project-All == Current Log grand total'));
    }
  }

  const failed = checks.filter(c => !c.pass);
  const requiredFailed = failed.filter(c => c.level === 'required');
  return {
    pass: requiredFailed.length === 0,
    requiredPass: requiredFailed.length === 0,
    allPass: failed.length === 0,
    checks,
    failed,
    summary: {
      total: checks.length,
      passed: checks.length - failed.length,
      requiredFailed: requiredFailed.length,
      recommendedFailed: failed.length - requiredFailed.length,
    },
  };
}

module.exports = { reconcile, readLog, cellNum, cellStr, cellFormula, COL, TOL };
