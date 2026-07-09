// ============================================================================
// Development-fee understanding + computation
// ----------------------------------------------------------------------------
// Every project computes its development fee differently: a different rate, a
// different base (some exclude interest/land, some include everything), some
// waive half, some round to the dollar, some don't round at all. Rather than
// hard-code one convention, this module LEARNS the method from the prior Req
// report's Dev Fee tab and applies that same method to the new period.
//
// Two-stage strategy (per Jimmy):
//   1. Deterministic formula parse — read the Dev Fee tab's E-column calc chain
//      (E15/E17/E19 …) and extract a structured spec {rate, halve, roundDecimals}.
//      This is exact and free when the tab uses the standard structure.
//   2. Claude fallback — when the formulas are missing/ambiguous, hand the whole
//      Dev Fee tab (formulas + cached values + labels) plus the prior period's
//      observed base→fee to Claude, which returns the SAME structured spec. The
//      model only proposes the spec; THIS code computes the number.
//
// Whichever stage produces the spec, we BACK-VALIDATE it against the prior
// period: recompute the prior fee from the prior base using the spec and check
// it matches the fee actually on the prior Current Log (within a cent). If it
// doesn't reconcile, we do NOT invent a fee — we return needsReview so the user
// enters it by hand instead of shipping a wrong number.
//
// A spec is:
//   { rate:Number, halve:Bool, roundDecimals:Number|null, source:String,
//     baseKind:'new_costs_ex_devfee', notes:String }
// applyDevFeeSpec(spec, base) => rounded fee number.
// ============================================================================

const { cellNum, cellStr, cellFormula, isDevFeeLabel } = require('./requisition_reconcile.js');

// Apply a learned spec to a base amount. Central so parse + Claude + validation
// all compute identically.
function applyDevFeeSpec(spec, base) {
  if (!spec || !(base > 0)) return null;
  let fee = base * spec.rate;
  if (spec.halve) fee = fee / 2;
  if (spec.roundDecimals != null) {
    const f = Math.pow(10, spec.roundDecimals);
    fee = Math.round(fee * f) / f;
  }
  // Always land on cents for a currency amount even when the spec rounds to a
  // coarser or no decimal (a raw float would look wrong on the log).
  return Math.round(fee * 100) / 100;
}

// ── Stage 1: deterministic parse of the Dev Fee tab's E-column calc chain ──
// The standard tab shape is:
//   E10 (or similar) = base cost
//   E15 = ROUND(E10 * <rate>, 2)      rate as 4% or 0.04
//   E17 = ROUND(E15 / 2, 2)           optional halving (waive half)
//   E19 = E17 (or =E15 when not halved) -> the number placed on the Current Log
// We read the rate out of the first "* <pct>" we find and detect a "/2" halving.
// roundDecimals is inferred from the ROUND(...,n) the tab actually uses.
function parseDevFeeSpec(devFeeWs) {
  if (!devFeeWs) return null;
  // Gather the E-column (and a couple neighbors) formulas so we can trace the
  // fee's calc chain regardless of the exact rows a given template uses.
  const scan = [];
  const last = Math.max(devFeeWs.rowCount || 0, devFeeWs.actualRowCount || 0, 30);
  for (let r = 1; r <= last; r++) {
    for (const col of ['B', 'C', 'D', 'E', 'F', 'G']) {
      const f = cellFormula(devFeeWs.getCell(col + r));
      if (f) scan.push({ addr: col + r, f });
    }
  }
  if (!scan.length) return null;

  // Find a rate: first formula containing "* N%" or "* 0.NN" or "*NN%".
  let rate = null, rateAddr = null;
  for (const { addr, f } of scan) {
    const pct = f.match(/\*\s*(\d+(?:\.\d+)?)\s*%/);          // * 4%  or *4%
    const dec = f.match(/\*\s*(0?\.\d+)\b/);                   // * 0.04
    if (pct) { rate = parseFloat(pct[1]) / 100; rateAddr = addr; break; }
    if (dec) { rate = parseFloat(dec[1]); rateAddr = addr; break; }
  }
  if (rate == null || !(rate > 0) || rate >= 1) return null; // not a usable rate

  // Detect halving: a "/2" or "/ 2" anywhere in the chain (waive-half convention).
  const halve = scan.some(({ f }) => /\/\s*2\b/.test(f));

  // Round decimals: read the ROUND(...,n) the tab uses on the fee cell; default 2.
  let roundDecimals = 2;
  const roundM = scan.map(s => s.f).join(' ').match(/ROUND\s*\([^,]+,\s*(\d+)\s*\)/i);
  if (roundM) roundDecimals = parseInt(roundM[1], 10);

  return {
    rate, halve, roundDecimals,
    baseKind: 'new_costs_ex_devfee',
    source: 'formula:' + (rateAddr || '?'),
    notes: `Parsed rate ${(rate * 100).toFixed(2)}%${halve ? ', halved' : ''}, round ${roundDecimals}dp from Dev Fee tab formulas.`,
  };
}

// Observe the prior period's actual base and fee so any spec can be validated,
// and so Claude has a concrete example to reason from. base = sum of the prior
// Current Log data rows EXCLUDING the dev-fee line; fee = the dev-fee line's
// amount. Uses the same COL map / cell readers as the engine.
function observePriorBaseFee(priorCurWs, devCode, COL) {
  if (!priorCurWs) return null;
  const last = Math.max(priorCurWs.rowCount || 0, priorCurWs.actualRowCount || 0);
  let base = 0, fee = null;
  for (let r = 3; r <= last; r++) {
    const f = cellFormula(priorCurWs.getCell(r, COL.amount));
    if (f && /SUBTOTAL/i.test(f)) continue; // skip subtotal/grand-total rows
    const amt = cellNum(priorCurWs.getCell(r, COL.amount));
    if (amt == null) continue;
    const nm = cellStr(priorCurWs.getCell(r, COL.name)).toLowerCase();
    const bank = cellStr(priorCurWs.getCell(r, COL.bankcat)).toLowerCase();
    const code = cellNum(priorCurWs.getCell(r, COL.code));
    const isDevFee = isDevFeeLabel(nm) || isDevFeeLabel(bank) ||
      (devCode != null && String(code) === String(devCode));
    if (isDevFee) { fee = (fee || 0) + amt; }
    else { base += amt; }
  }
  return { base: Math.round(base * 100) / 100, fee: fee == null ? null : Math.round(fee * 100) / 100 };
}

// Back-validate a spec against the prior period: does applying it to the prior
// base reproduce the prior fee (within a cent)? Returns { ok, expected, got }.
function validateSpecAgainstPrior(spec, prior) {
  if (!spec || !prior || prior.fee == null || !(prior.base > 0)) {
    return { ok: false, reason: 'no prior base/fee to validate against' };
  }
  const got = applyDevFeeSpec(spec, prior.base);
  const ok = got != null && Math.abs(got - prior.fee) <= 0.01;
  return { ok, expected: prior.fee, got, reason: ok ? null : 'spec did not reproduce prior fee' };
}

// ── Stage 2: Claude fallback ──
// Dump the Dev Fee tab (formulas + cached values + text labels) plus the prior
// observed base→fee, and ask for the structured spec. The model NEVER returns a
// dollar amount — only {rate, halve, roundDecimals} — so the number is always
// computed by us and back-validated.
function dumpDevFeeTab(devFeeWs) {
  if (!devFeeWs) return [];
  const cells = [];
  const last = Math.max(devFeeWs.rowCount || 0, devFeeWs.actualRowCount || 0, 30);
  for (let r = 1; r <= last; r++) {
    for (let c = 1; c <= 12; c++) {
      const cell = devFeeWs.getCell(r, c);
      const f = cellFormula(cell);
      const v = cell.value;
      let val;
      if (v && typeof v === 'object') val = ('result' in v) ? v.result : undefined;
      else val = v;
      if (f == null && (val == null || val === '')) continue;
      cells.push({ addr: cell.address, formula: f || undefined, value: (val === undefined ? undefined : val) });
    }
  }
  return cells;
}

function makeDevFeeSpecPrompt(cells, prior) {
  const SYSTEM = [
    'You infer HOW a real-estate development fee is calculated from a project\'s',
    '"Dev Fee" worksheet. Different projects use different rates, bases, rounding,',
    'and some waive half the fee. Return ONLY a JSON object describing the method:',
    '{"rate": <decimal, e.g. 0.04>, "halve": <true|false>, "roundDecimals": <int|null>,',
    ' "notes": "<one sentence on how you read it>"}.',
    'rate is the fraction applied to the base (4% -> 0.04). halve=true when the',
    'sheet divides the fee by 2 (waived half). roundDecimals is the number of',
    'decimals the sheet rounds the fee to (2 for cents; null if it never rounds).',
    'The base is this-period new costs excluding the dev-fee line — you do NOT',
    'decide the base, only rate/halve/rounding. Never output a dollar amount.',
    'No prose outside the JSON.',
  ].join(' ');
  const user = {
    instruction: 'Infer the dev-fee calculation method from these Dev Fee tab cells. ' +
      'Use the prior period example to confirm your rate/halve/rounding reproduce its fee.',
    dev_fee_tab_cells: cells,
    prior_period_example: prior && prior.fee != null
      ? { base_ex_devfee: prior.base, actual_dev_fee: prior.fee }
      : 'none available',
  };
  return { SYSTEM, user };
}

// callClaude(payload) -> parsed JSON. Injected so this module has no hard SDK
// dependency (mirrors requisition_rollforward_verify's makeClaudeCaller).
async function analyzeDevFeeWithClaude(devFeeWs, prior, callClaude) {
  if (!callClaude) return null;
  const cells = dumpDevFeeTab(devFeeWs);
  if (!cells.length) return null;
  const { SYSTEM, user } = makeDevFeeSpecPrompt(cells, prior);
  let parsed;
  try {
    parsed = await callClaude({ system: SYSTEM, user });
  } catch (e) {
    return { error: 'claude call failed: ' + e.message };
  }
  if (!parsed || typeof parsed.rate !== 'number' || !(parsed.rate > 0) || parsed.rate >= 1) {
    return { error: 'model did not return a usable rate', raw: parsed };
  }
  return {
    rate: parsed.rate,
    halve: !!parsed.halve,
    roundDecimals: (parsed.roundDecimals === null || Number.isInteger(parsed.roundDecimals)) ? parsed.roundDecimals : 2,
    baseKind: 'new_costs_ex_devfee',
    source: 'claude',
    notes: (parsed.notes || 'Inferred by Claude from the Dev Fee tab.').toString().slice(0, 300),
  };
}

// A Claude caller specialized for the dev-fee spec prompt. Same transport as the
// roll-forward verifier's caller. Returns parsed JSON {rate,halve,...}.
function makeDevFeeClaudeCaller({ apiKey = process.env.ANTHROPIC_API_KEY, model = 'claude-haiku-4-5-20251001', fetchImpl } = {}) {
  const doFetch = fetchImpl || globalThis.fetch;
  return async function callClaude({ system, user }) {
    const res = await doFetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model, max_tokens: 512, system,
        messages: [{ role: 'user', content: JSON.stringify(user, null, 2) }],
      }),
    });
    if (!res.ok) throw new Error('Anthropic API ' + res.status + ': ' + (await res.text()).slice(0, 200));
    const data = await res.json();
    const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
    const clean = text.replace(/```json|```/g, '').trim();
    return JSON.parse(clean);
  };
}

// ── Orchestrator: learn the spec (parse → validate → Claude → validate) ──
// Returns { spec, prior, validation, needsReview, source } where spec may be
// null (couldn't be learned or didn't validate). The caller decides whether to
// write a fee (spec present, !needsReview) or leave it for manual entry.
async function learnDevFeeSpec({ devFeeWs, priorCurWs, devCode, COL, callClaude = null, trustParsed = false }) {
  const prior = observePriorBaseFee(priorCurWs, devCode, COL);

  // Stage 1: deterministic parse.
  const parsed = parseDevFeeSpec(devFeeWs);
  if (parsed) {
    const v = validateSpecAgainstPrior(parsed, prior);
    // If we have a prior fee to check against and it reconciles, trust it. If we
    // have no prior fee to check (first-ever dev fee), accept the parsed spec but
    // flag it lightly via source. If it FAILS validation, fall through to Claude.
    if (v.ok || prior == null || prior.fee == null || trustParsed) {
      // trustParsed: for collapse entities the base is redefined as this period's
      // new-invoice total x the parsed rate, so the prior-fee reproduction gate
      // (which assumes the prior fee used the same base) does not apply.
      return { spec: parsed, prior, validation: v, needsReview: false, source: parsed.source };
    }
  }

  // Stage 2: Claude fallback (only when parse missing or failed validation).
  if (callClaude) {
    const inferred = await analyzeDevFeeWithClaude(devFeeWs, prior, callClaude);
    if (inferred && !inferred.error) {
      const v = validateSpecAgainstPrior(inferred, prior);
      if (v.ok || prior == null || prior.fee == null) {
        return { spec: inferred, prior, validation: v, needsReview: !(v.ok), source: 'claude' };
      }
      // Claude produced a spec but it didn't reproduce the prior fee → don't ship.
      return { spec: null, prior, validation: v, needsReview: true, source: 'claude-unvalidated',
        note: 'Claude proposed a method but it did not reproduce the prior fee; enter dev fee manually.' };
    }
  }

  // Nothing usable. If parse existed but failed validation, surface that.
  if (parsed) {
    const v = validateSpecAgainstPrior(parsed, prior);
    return { spec: null, prior, validation: v, needsReview: true, source: 'formula-unvalidated',
      note: 'Parsed a method from the Dev Fee tab but it did not reproduce the prior fee; enter dev fee manually.' };
  }
  return { spec: null, prior, validation: null, needsReview: true, source: 'none',
    note: 'Could not determine the dev-fee method from the Dev Fee tab; enter dev fee manually.' };
}

module.exports = {
  applyDevFeeSpec, parseDevFeeSpec, observePriorBaseFee, validateSpecAgainstPrior,
  analyzeDevFeeWithClaude, makeDevFeeClaudeCaller, learnDevFeeSpec, dumpDevFeeTab,
};
