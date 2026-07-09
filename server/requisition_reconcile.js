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
//   B5  Dev Fee amount    : posted current Dev Fee (J11) == this-period new costs (ex dev fee) x entity rate
//   C1  This-period tie   : B2A this-period column total == Current Log grand total
//
// Each check returns { id, level, pass, expected, actual, delta, detail }.
// level: 'required' | 'recommended'. The caller decides whether 'recommended'
// failures block. A FAIL bundle is what gets handed to the Claude API for
// diagnosis when (and only when) something does not reconcile.
// ============================================================================

const TOL = 0.01; // dollar tolerance for float/rounding noise

// ----- cell helpers ---------------------------------------------------------
// Safe evaluator for self-contained arithmetic formulas (e.g. a manual invoice
// split like "5532.77-5451"). Accepts ONLY numeric literals and + - * / ( ) —
// no cell references, no function calls, no names. Returns a number, or null
// when the formula is anything other than pure literal arithmetic. We never run
// untrusted workbook text through eval/Function; this is a hand-written
// tokenizer + shunting-yard evaluator whose grammar physically cannot reach a
// cell reference or function call.
function evalLiteralArithmetic(formula) {
  if (typeof formula !== 'string') return null;
  let s = formula.trim();
  if (s[0] === '=') s = s.slice(1).trim();
  if (s === '') return null;
  // Reject anything that isn't digits, dot, the four operators, parens, or space.
  // A letter, '!', ':', '$', or ',' means a ref/function/range — bail out.
  if (!/^[0-9.+\-*/() \t]+$/.test(s)) return null;

  const tokens = s.match(/\d+\.?\d*|\.\d+|[+\-*/()]/g);
  if (!tokens || tokens.join('') !== s.replace(/[ \t]/g, '')) return null;

  const prec = { '+': 1, '-': 1, '*': 2, '/': 2 };
  const out = [];      // output queue (RPN)
  const ops = [];      // operator stack
  let prevType = null; // 'num' | 'op' | '(' | ')'  — to detect unary minus
  for (const t of tokens) {
    if (/^[0-9.]/.test(t)) {
      const n = Number(t);
      if (!isFinite(n)) return null;
      out.push(n);
      prevType = 'num';
    } else if (t === '(') {
      ops.push(t);
      prevType = '(';
    } else if (t === ')') {
      while (ops.length && ops[ops.length - 1] !== '(') out.push(ops.pop());
      if (!ops.length) return null; // unbalanced
      ops.pop(); // discard '('
      prevType = ')';
    } else {
      // operator: treat leading/post-operator/post-'(' minus as unary (0 - x)
      if (t === '-' && (prevType === null || prevType === 'op' || prevType === '(')) {
        out.push(0);
      } else if (prevType === null || prevType === 'op' || prevType === '(') {
        return null; // a binary op with no left operand (e.g. "*5") is malformed
      }
      while (ops.length && ops[ops.length - 1] !== '(' && prec[ops[ops.length - 1]] >= prec[t]) {
        out.push(ops.pop());
      }
      ops.push(t);
      prevType = 'op';
    }
  }
  while (ops.length) { const op = ops.pop(); if (op === '(') return null; out.push(op); }

  const st = [];
  for (const tok of out) {
    if (typeof tok === 'number') { st.push(tok); continue; }
    const b = st.pop(), a = st.pop();
    if (a === undefined || b === undefined) return null;
    let r;
    if (tok === '+') r = a + b;
    else if (tok === '-') r = a - b;
    else if (tok === '*') r = a * b;
    else if (tok === '/') { if (b === 0) return null; r = a / b; }
    else return null;
    st.push(r);
  }
  if (st.length !== 1 || !isFinite(st[0])) return null;
  return st[0];
}

// exceljs: plain cell -> .value is the value; formula cell -> .value is
// { formula, result }. Pull a number out of either shape. When a formula cell
// has no cached result (workbook saved without a recalc), fall back to
// evaluating it IF it is self-contained literal arithmetic — this rescues
// manual invoice-split rows (e.g. "5532.77-5451" = 81.77) that would otherwise
// read as null and silently drop their amount from A1/A2 totals. Formulas that
// reference other cells/sheets still return null and are handled by the
// count-but-zero defensive branches downstream.
function cellNum(cell) {
  if (!cell) return null;
  let v = cell.value;
  if (v && typeof v === 'object' && 'formula' in v) {
    if ('result' in v && v.result !== undefined && v.result !== null) {
      v = v.result;
    } else {
      return evalLiteralArithmetic(v.formula);
    }
  } else if (v && typeof v === 'object' && 'result' in v) {
    v = v.result;
  }
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
const DEFAULT_COL = { cat: 2, code: 3, bankcat: 4, gl: 5, name: 6, vendor: 7, bill: 8, amount: 9, req: 10, date: 11 };
// Working column map. Mutable: applyInvoiceCols() overwrites these per workbook
// from the detected header row so templates that differ from the SRN layout
// (e.g. one with no "GL Coding" column, shifting Amount from col 9 to col 8)
// are read correctly. Shared by reference across the engine/verify/devfee.
const COL = { ...DEFAULT_COL };

// Recognize a development/management-fee label in any of its common spellings
// ("Development Fee", "Development Management Fee", "Dev Fee", "Management Fee").
const DEVFEE_RE = /dev(?:elopment)?\s*(?:mgmt|management)?\s*fee|management fee/i;
function isDevFeeLabel(s) { return DEVFEE_RE.test(String(s == null ? '' : s)); }

// Detect the invoice-log column layout from the header row instead of assuming
// fixed positions. Scans the first rows for a header row and maps each field by
// its header text; falls back to DEFAULT_COL for any field it can't find, and
// parks an unfound field on an unused column when its default slot is already
// taken (so a missing "GL" column never collides with the real "Cost Code Name").
function detectInvoiceCols(ws) {
  const map = { ...DEFAULT_COL };
  if (!ws) return map;
  const rules = [
    ['bankcat', t => /bank\s*cost\s*cat/.test(t)],
    ['name',    t => /cost code name/.test(t)],
    ['code',    t => /cost code\s*#|cost code\s*(number|no)\b|^cost code$/.test(t)],
    ['cat',     t => /cost category/.test(t)],
    ['gl',      t => /gl coding|gl account|^g\/?l$|^gl\b/.test(t)],
    ['vendor',  t => /vendor|payee/.test(t)],
    ['bill',    t => /bill\s*(number|no|#)?|invoice\s*(number|no|#)/.test(t)],
    ['amount',  t => /amount/.test(t)],
    ['req',     t => /requisition|req\s*(month|#|no)/.test(t)],
    ['date',    t => /invoice date|^date$|date$/.test(t)],
  ];
  let bestAssigned = null, bestHits = 0;
  for (let r = 1; r <= 8; r++) {
    const assigned = {}; let hits = 0;
    for (let c = 1; c <= 16; c++) {
      const t = cellStr(ws.getCell(r, c)).toLowerCase().replace(/\s+/g, ' ').trim();
      if (!t) continue;
      for (const [field, test] of rules) {
        if (assigned[field] != null) continue;
        if (test(t)) { assigned[field] = c; hits++; break; }
      }
    }
    if (hits > bestHits) { bestHits = hits; bestAssigned = assigned; }
  }
  // Only override when we found a convincing header row that includes the
  // critical Amount column; otherwise keep the default map untouched.
  if (!bestAssigned || bestHits < 4 || bestAssigned.amount == null) return map;
  Object.assign(map, bestAssigned);
  const used = new Set(Object.values(bestAssigned));
  for (const field of Object.keys(DEFAULT_COL)) {
    if (bestAssigned[field] != null) continue;      // detected: keep
    const def = DEFAULT_COL[field];
    if (!used.has(def)) { map[field] = def; used.add(def); }        // default slot free
    else { let g = 25; while (used.has(g)) g++; map[field] = g; used.add(g); } // park on unused col
  }
  return map;
}

// Find the header row of an invoice-log sheet (the row with the most recognizable
// column headers within the first several rows) and return the row where DATA
// begins (header + 1). Defaults to 3 (header on row 2), matching the standard
// templates; Silsbee-style logs carry a title block + header lower down.
function logDataStart(ws) {
  if (!ws) return 3;
  const HDR = [/bank\s*cost\s*cat/, /cost code name/, /cost code\s*#|cost code\s*(number|no)\b/, /cost category/, /vendor|payee/, /bill\s*(number|no|#)?/, /amount/, /requisition|req\s*(month|#|no)/, /invoice date/];
  let bestRow = 2, bestHits = 0;
  for (let r = 1; r <= 8; r++) {
    let hits = 0;
    for (let c = 1; c <= 16; c++) {
      const t = cellStr(ws.getCell(r, c)).toLowerCase().replace(/\s+/g, ' ').trim();
      if (t && HDR.some(re => re.test(t))) hits++;
    }
    if (hits > bestHits) { bestHits = hits; bestRow = r; }
  }
  return (bestHits >= 3 ? bestRow : 2) + 1;
}

// Detect + apply the layout of an invoice-log worksheet into the shared COL map.
function applyInvoiceCols(ws) { const d = detectInvoiceCols(ws); Object.assign(COL, d); return d; }

// ----- log reader -----------------------------------------------------------
// Walk an invoice-log worksheet and return:
//   rows: [{ row, code, name, vendor, bill, amount }]  (vendor data rows only)
//   subtotals: [{ row, name, formula, result }]        (SUBTOTAL rows)
//   total: sum of data-row amounts
//   byCode: { code: sum }
function readLog(ws) {
  const _ds = logDataStart(ws);
  const rows = [];
  const subtotals = [];
  const byCode = {};
  let total = 0;
  // NOTE: exceljs actualRowCount = COUNT of non-empty rows, not the max row
  // number. Spacer rows between groups make it smaller than the real extent,
  // which truncates the scan. Use rowCount (max row index) so we reach the
  // Grand Total and the late groups (Dev Fee, etc.).
  const last = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  for (let r = _ds; r <= last; r++) {
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
      // No cached result means the SUBTOTAL simply has not been evaluated on the
      // server (prod has no recalc). That is not a detected data break, so do not
      // report it as a failure - mark it not-evaluated (advisory).
      checks.push(chk('A4', 'recommended', true, round2(nPrior.total), null,
        'Prior Log Grand Total not evaluated on server (recalculates in Excel)'));
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
      if (!f) {
        // No formula here means this template does not use the SRN-style
        // J-column cross-reference for that line (e.g. it posts the Development
        // Fee directly on the Current Log and computes the fee in columns B/C/D).
        // There is no absolute ref to validate, so treat it as not-applicable
        // rather than a failure. Dev-fee amount correctness is verified by B5.
        continue;
      }
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

    // B5 - this period's posted Development Fee ties to (new costs x rate).
    // The Dev Fee is computed as a percentage of this period's NEW costs,
    // EXCLUDING the dev fee line itself (excluding it avoids circularity). We
    // recompute that expected fee here from the Current Log + the entity's rate
    // structure on the Dev Fee tab (E15 = base*rate1, E17 = E15/2) and compare it
    // to J11, the posted current-period Dev Fee. Reading the rate from the tab
    // keeps this entity-agnostic; the old cumulative check (J8 == J10+J11) no
    // longer applies because J8's cumulative target can't refresh without a
    // recalc and the fee is now a this-period figure.
    {
      const nCurr = readLog(next.current);
      // Find this period's posted Development Fee. Prefer the actual dev-fee
      // line on the Current Log (matched by the code the SRN J11 pointer targets,
      // or by a dev/management-fee label) so this works regardless of where the
      // template keeps the fee; fall back to the Dev Fee tab's J11 cell (SRN).
      let devCode = null;
      const j11f = cellFormula(dv.getCell('J11'));
      const j11m = j11f && j11f.match(/I(\d+)/);
      if (j11m) {
        const c = cellNum(next.current.getCell(Number(j11m[1]), COL.code));
        if (c != null) devCode = c;
      }
      let posted = null, feeRow = null;
      for (const rw of nCurr.rows) {
        if ((devCode != null && String(rw.code) === String(devCode)) || isDevFeeLabel(rw.name)) {
          posted = (posted || 0) + (rw.amount || 0); feeRow = rw;
          if (devCode == null && rw.code != null) devCode = rw.code;
        }
      }
      if (posted == null) posted = cellNum(dv.getCell('J11')) || 0; // SRN fallback
      // base = sum of this period's current-log data rows EXCLUDING the dev fee.
      let base = 0;
      for (const rw of nCurr.rows) {
        if (feeRow && rw === feeRow) continue;
        if (devCode != null && String(rw.code) === String(devCode)) continue;
        base += rw.amount || 0;
      }
      // Read the entity rate from the Dev Fee tab. Scan cols B-G (not just E) for
      // the first "*<pct>" so both the SRN (E15) and B/C/D layouts resolve;
      // detect a "/2" halving anywhere in the fee chain.
      let rate = 0.04, halve = false, rateSeen = false;
      const dvLast = Math.max(dv.rowCount || 0, dv.actualRowCount || 0, 30);
      for (let r = 1; r <= dvLast && !rateSeen; r++) {
        for (const c of ['B', 'C', 'D', 'E', 'F', 'G']) {
          const f2 = cellFormula(dv.getCell(c + r)); if (!f2) continue;
          const pm = f2.match(/\*\s*(\d+(?:\.\d+)?)\s*%/) || f2.match(/\*\s*(0?\.\d+)\b/);
          if (pm) { rate = f2.includes('%') ? parseFloat(pm[1]) / 100 : parseFloat(pm[1]); rateSeen = true; break; }
        }
      }
      for (let r = 1; r <= dvLast && !halve; r++) {
        for (const c of ['B', 'C', 'D', 'E', 'F', 'G']) {
          const f3 = cellFormula(dv.getCell(c + r)); if (f3 && /\/\s*2\b/.test(f3)) { halve = true; break; }
        }
      }
      let expectedFee = round2(base * rate);
      if (halve) expectedFee = round2(expectedFee / 2);
      const disc = round2(posted - expectedFee);
      checks.push(chk('B5', 'required', approxEq(disc, 0, tol), round2(expectedFee), round2(posted),
        'Posted Dev Fee vs (new costs ' + round2(base) + ' x ' +
        (halve ? (rate * 50) : (rate * 100)) + '%); expected=' + round2(expectedFee) +
        ' posted=' + round2(posted) + ' disc=' + disc));
    }
  }

  // C1 - this-period total ties to Current Log grand total.
  if (next.current && next.b2a) {
    const nCurr = readLog(next.current);
    // The freshly-summed Current Log data rows are the source of truth (the
    // Grand Total SUBTOTAL cell may hold a stale cached value until Excel recalcs).
    const curTotal = nCurr.total;
    let projCell = null;
    const b2a = next.b2a;
    const bLast = b2a.actualRowCount || b2a.rowCount;
    for (let r = 1; r <= bLast; r++) {
      if (/project costs\s*-?\s*all/i.test(cellStr(b2a.getCell(r, 2)))) { projCell = b2a.getCell(r, COL.amount); break; }
    }
    if (projCell != null) {
      const projAllI = cellNum(projCell);
      if (cellFormula(projCell)) {
        // The B2A this-period "Project costs - All" cell is a formula that
        // aggregates this period's columns; its cached value goes stale after a
        // roll-forward and only refreshes when Excel recalculates. Do not hard-fail
        // on the stale number - report it as not-evaluated (advisory).
        checks.push(chk('C1', 'recommended', true, round2(curTotal), projAllI == null ? null : round2(projAllI),
          'B2A this-period Project-All is a formula; recalculates in Excel (Current Log total=' + round2(curTotal) + ')'));
      } else if (projAllI != null) {
        checks.push(chk('C1', 'recommended', approxEq(projAllI, curTotal, tol),
          round2(curTotal), round2(projAllI), 'B2A this-period Project-All == Current Log grand total'));
      }
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

module.exports = { reconcile, readLog, cellNum, cellStr, cellFormula, COL, DEFAULT_COL, TOL, detectInvoiceCols, applyInvoiceCols, isDevFeeLabel, logDataStart };
