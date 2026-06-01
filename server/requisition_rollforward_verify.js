// ============================================================================
// Requisition roll-forward verification + auto-repair orchestrator
// ----------------------------------------------------------------------------
// Wraps a roll-forward with the deterministic reconciliation engine and, ONLY
// when something fails to reconcile, asks the Claude API to diagnose the
// mechanical cause and propose a structured fix. The fix is applied by THIS
// code (never by the model directly), then reconciliation is re-run. Re-passing
// is the safety guarantee: a wrong fix cannot pass, because the A/B/C identities
// are pure arithmetic. After `maxRetries` unsuccessful rounds it stops and
// returns the outstanding failures for a human to look at.
//
// A roll-forward only moves data, so a failure is always a mechanical bug
// (dropped cell, shifted absolute ref, stale SUBTOTAL range) - there is no
// "high-risk" data/judgement failure mode here, so no human gate is required on
// the happy path. The model is given a tightly-scoped patch vocabulary and any
// patch outside it is rejected before it can touch the workbook.
//
// Design notes:
//   - `recalc(workbook)` must recompute formula results (e.g. via a headless
//     LibreOffice convert) and return a workbook whose formula cells carry
//     cached results, because B5/A4 read evaluated values. If no recalc is
//     available, B5/A4 degrade to "not evaluated" and only the structural
//     checks (A1-A3, B1, B4) gate.
//   - `callClaude(promptPayload)` is injected so this module has no hard SDK
//     dependency and is unit-testable with a mock. In production it posts to the
//     Anthropic Messages API (see makeClaudeCaller).
// ============================================================================

const { reconcile, cellNum, cellStr, cellFormula, COL } = require('./requisition_reconcile.js');

// Allowed patch operations the model may return. Anything else is rejected.
//   setFormula : { op:'setFormula', sheet, cell, formula }    e.g. fix Dev Fee J6 ref
//   setValue   : { op:'setValue',   sheet, cell, value }      e.g. restore a dropped amount
// `sheet` is one of: 'Prior Invoice Log' | 'Current Invoice Log' |
//                    'Budget to Actual' | 'Dev Fee'
const ALLOWED_OPS = new Set(['setFormula', 'setValue']);
const ALLOWED_SHEETS = new Set([
  'Prior Invoice Log', 'Current Invoice Log', 'Budget to Actual', 'Dev Fee',
  'Hard Cost Contingency Table', 'Soft Cost Contingency Table',
]);

function getSheets(workbook) {
  return {
    prior: workbook.getWorksheet('Prior Invoice Log'),
    current: workbook.getWorksheet('Current Invoice Log'),
    b2a: workbook.getWorksheet('Budget to Actual'),
    devFee: workbook.getWorksheet('Dev Fee'),
  };
}

// Validate + apply one structured patch to an exceljs workbook in place.
// Returns { ok, error } so a bad patch is rejected, not silently applied.
function applyPatch(workbook, patch) {
  if (!patch || typeof patch !== 'object') return { ok: false, error: 'patch not an object' };
  if (!ALLOWED_OPS.has(patch.op)) return { ok: false, error: `op not allowed: ${patch.op}` };
  if (!ALLOWED_SHEETS.has(patch.sheet)) return { ok: false, error: `sheet not allowed: ${patch.sheet}` };
  if (!/^[A-Z]+[0-9]+$/.test(String(patch.cell || ''))) return { ok: false, error: `bad cell ref: ${patch.cell}` };
  const ws = workbook.getWorksheet(patch.sheet);
  if (!ws) return { ok: false, error: `sheet missing: ${patch.sheet}` };

  if (patch.op === 'setFormula') {
    if (typeof patch.formula !== 'string' || !patch.formula.length) {
      return { ok: false, error: 'setFormula requires a formula string' };
    }
    // Strip a leading '=' if present; exceljs wants the bare formula.
    const formula = patch.formula.replace(/^=/, '');
    ws.getCell(patch.cell).value = { formula };
    return { ok: true };
  }
  if (patch.op === 'setValue') {
    if (typeof patch.value !== 'number') {
      return { ok: false, error: 'setValue requires a numeric value' };
    }
    ws.getCell(patch.cell).value = patch.value;
    return { ok: true };
  }
  return { ok: false, error: 'unreachable' };
}

// Build the diagnosis payload handed to the model on a FAIL. It contains only
// the failing checks plus a little surrounding context - never the whole book.
function buildDiagnosisPayload(reconResult, nextSheets) {
  const failing = reconResult.failed.map(c => ({
    id: c.id, level: c.level, expected: c.expected, actual: c.actual, delta: c.delta, detail: c.detail,
  }));
  // Light context: dev-fee ref formulas + the prior-log grand-total formula,
  // since those are the usual mechanical culprits.
  const dv = nextSheets.devFee;
  const context = {};
  if (dv) {
    context.devFee = {
      J6: cellFormula(dv.getCell('J6')),
      J10: cellFormula(dv.getCell('J10')),
      J11: cellFormula(dv.getCell('J11')),
      J8_result: cellNum(dv.getCell('J8')),
      J10_result: cellNum(dv.getCell('J10')),
      J11_result: cellNum(dv.getCell('J11')),
    };
  }
  return { failing, context };
}

// ----------------------------------------------------------------------------
// Main orchestrator.
//   opts:
//     prevSheets   : { prior, current }  exceljs worksheets of Req#N
//     nextWorkbook : exceljs Workbook of the freshly rolled-forward Req#N+1
//     recalc       : async (workbook) => workbook  (recompute formula results)
//     callClaude   : async (payload) => { patches: [...], explanation }   (optional)
//     maxRetries   : number of diagnose->patch->reverify rounds (default 3)
//     tol          : dollar tolerance (default from reconcile)
//   returns:
//     { ok, attempts, finalResult, history, unresolved }
// ----------------------------------------------------------------------------
async function verifyRollforward(opts) {
  const {
    prevSheets, nextWorkbook, recalc = null, callClaude = null,
    maxRetries = 3, tol,
  } = opts;

  const history = [];
  let workbook = nextWorkbook;

  // Always recalc once up front so formula-dependent checks (A4/B5) have results.
  if (recalc) workbook = await recalc(workbook);

  let nextSheets = getSheets(workbook);
  let result = reconcile(prevSheets, nextSheets, tol != null ? { tol } : {});
  history.push({ stage: 'initial', summary: result.summary, failed: result.failed.map(f => f.id) });

  if (result.pass) {
    return { ok: true, attempts: 0, finalResult: result, history, unresolved: [] };
  }

  // FAIL -> diagnose + repair loop (only reached when something is wrong).
  if (!callClaude) {
    return { ok: false, attempts: 0, finalResult: result, history,
      unresolved: result.failed.filter(f => f.level === 'required'),
      note: 'reconciliation failed and no callClaude was provided to attempt repair' };
  }

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const payload = buildDiagnosisPayload(result, nextSheets);
    let diagnosis;
    try {
      diagnosis = await callClaude(payload);
    } catch (e) {
      history.push({ stage: `attempt${attempt}`, error: 'callClaude threw: ' + e.message });
      break;
    }
    const patches = (diagnosis && Array.isArray(diagnosis.patches)) ? diagnosis.patches : [];
    if (!patches.length) {
      history.push({ stage: `attempt${attempt}`, note: 'model returned no patches', explanation: diagnosis && diagnosis.explanation });
      break;
    }

    // Apply each patch, rejecting any that fall outside the allowed vocabulary.
    const applied = [];
    const rejected = [];
    for (const p of patches) {
      const r = applyPatch(workbook, p);
      (r.ok ? applied : rejected).push({ patch: p, error: r.error });
    }

    // Re-evaluate formulas, then re-run reconciliation - the safety gate.
    if (recalc) workbook = await recalc(workbook);
    nextSheets = getSheets(workbook);
    result = reconcile(prevSheets, nextSheets, tol != null ? { tol } : {});
    history.push({
      stage: `attempt${attempt}`,
      explanation: diagnosis.explanation,
      applied: applied.map(a => a.patch),
      rejected,
      summary: result.summary,
      failed: result.failed.map(f => f.id),
    });

    if (result.pass) {
      return { ok: true, attempts: attempt, finalResult: result, history, unresolved: [] };
    }
  }

  return {
    ok: false,
    attempts: maxRetries,
    finalResult: result,
    history,
    unresolved: result.failed.filter(f => f.level === 'required'),
    note: 'auto-repair exhausted retries; needs human review',
  };
}

// ----------------------------------------------------------------------------
// Production Claude caller. Posts the diagnosis payload to the Anthropic
// Messages API and parses a strict-JSON patch list back. No SDK dependency -
// uses global fetch (Node 18+). The model is instructed to return ONLY JSON.
// ----------------------------------------------------------------------------
function makeClaudeCaller({ apiKey = process.env.ANTHROPIC_API_KEY, model = 'claude-haiku-4-5-20251001', fetchImpl } = {}) {
  const doFetch = fetchImpl || globalThis.fetch;
  const SYSTEM = [
    'You diagnose MECHANICAL errors in a real-estate requisition roll-forward.',
    'A roll-forward only moves Req#N Current Log into Req#N+1 Prior Log, so every',
    'failure is mechanical: a dropped/zeroed amount cell, a shifted absolute',
    'reference, or a stale SUBTOTAL range. You never invent new amounts.',
    'Return ONLY a JSON object: {"explanation": string, "patches": [ ... ]}.',
    'Each patch is one of:',
    '  {"op":"setFormula","sheet":SHEET,"cell":"J6","formula":"-\'Prior Invoice Log\'!I217"}',
    '  {"op":"setValue","sheet":SHEET,"cell":"I415","value":81.77}',
    'SHEET is one of: "Prior Invoice Log","Current Invoice Log","Budget to Actual","Dev Fee".',
    'Only use setValue to restore an amount you can justify from the provided context;',
    'never fabricate. Prefer setFormula fixes for shifted references. No prose outside JSON.',
  ].join(' ');

  return async function callClaude(payload) {
    const body = {
      model,
      max_tokens: 1024,
      system: SYSTEM,
      messages: [{ role: 'user', content: 'Reconciliation failed. Diagnose and return JSON patches.\n\n' + JSON.stringify(payload, null, 2) }],
    };
    const res = await doFetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error('Anthropic API ' + res.status + ': ' + (await res.text()).slice(0, 300));
    const data = await res.json();
    const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
    const clean = text.replace(/```json|```/g, '').trim();
    let parsed;
    try { parsed = JSON.parse(clean); }
    catch (e) { throw new Error('model did not return valid JSON: ' + clean.slice(0, 200)); }
    return parsed;
  };
}

module.exports = { verifyRollforward, applyPatch, buildDiagnosisPayload, makeClaudeCaller, ALLOWED_OPS, ALLOWED_SHEETS };
