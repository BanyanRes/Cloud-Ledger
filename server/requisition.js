// ============================================================================
// Requisition coding engine
// ----------------------------------------------------------------------------
// Predicts the COST CODE (the value that flows into the Budget-to-Actual report)
// for each invoice on a development project, learned from how the same vendor /
// invoice was coded in prior requisitions. GL coding is intentionally ignored:
// GL sub-codes collapse 1:1 into a single cost code, so cost code is the target.
//
// Validated on County Line SRN Req#10-14 (walk-forward backtest):
//   - HIGH-confidence auto-coding: 100% correct (190/190)
//   - auto-coverage: ~73% of lines
//   - 85% land on the correct cost code after a quick human review of the rest
//
// Confidence tiers returned by predict():
//   high   -> unique historical cost code for this vendor (+bill signature). Auto-apply.
//   review -> vendor historically split across multiple cost codes. Propose top
//             candidate + full candidate list; a human confirms.
//   new    -> vendor not seen in history. Leave blank for one-time manual coding;
//             it becomes history for next month.
// ============================================================================

// Normalize a vendor name for matching: lowercase, strip punctuation and common
// entity suffixes (llc/inc/co/...), collapse whitespace.
function normVendor(v) {
  if (!v) return '';
  let s = String(v).toLowerCase();
  s = s.replace(/[.,]/g, '');
  s = s.replace(/\b(pllc|pllp|llp|llc|pc|pa|inc|incorporated|company|co|corp|corporation|ltd|limited|lp|the)\b/g, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// Bill-number signature: the disambiguator for vendors that split across codes
// (e.g. payroll "... - RRB" vs "... - Payroll Taxes"). Strip dates and digits,
// keep alphabetic tokens (>=2 chars) as a sorted, de-duplicated key.
function billSig(bill) {
  let b = String(bill == null ? '' : bill).toLowerCase();
  b = b.replace(/\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g, ''); // remove dates
  b = b.replace(/\d/g, '');                                   // remove remaining digits
  const toks = (b.match(/[a-z]{2,}/g) || []);
  const uniq = Array.from(new Set(toks)).sort();
  return uniq.join(' ');
}

// Build the in-memory history index for one entity from requisition_coding_history.
// For split vendors we pick the MOST RECENT coding (highest req_number), so each
// cost-code entry tracks last_req; weight is kept as a tiebreaker.
// Returns:
//   keyMap:    Map "<vendorNorm>|<billSig>" -> Map(cost_code -> {weight, last_req, sample})
//   vendorMap: Map "<vendorNorm>"           -> Map(cost_code -> {weight, last_req, sample})
//   vendorTokens: Array of { norm, tokens:Set } for token-based (similar) matching.
function buildHistoryIndex(db, entityId) {
  const rows = db.prepare(
    'SELECT vendor_norm, bill_signature, cost_category, cost_code, bank_cost_category, gl_coding, cost_code_name, weight, req_number ' +
    'FROM requisition_coding_history WHERE entity_id = ?'
  ).all(entityId);

  const keyMap = new Map();
  const vendorMap = new Map();

  const bump = (map, mapKey, row) => {
    let codes = map.get(mapKey);
    if (!codes) { codes = new Map(); map.set(mapKey, codes); }
    const cc = row.cost_code == null ? '' : String(row.cost_code);
    let entry = codes.get(cc);
    if (!entry) {
      entry = {
        weight: 0,
        last_req: -Infinity,
        sample: {
          cost_category: row.cost_category,
          cost_code: row.cost_code,
          bank_cost_category: row.bank_cost_category,
          gl_coding: row.gl_coding,
          cost_code_name: row.cost_code_name,
        },
      };
      codes.set(cc, entry);
    }
    entry.weight += (row.weight || 1);
    const rq = row.req_number != null ? Number(row.req_number) : null;
    if (rq != null && Number.isFinite(rq) && rq > entry.last_req) {
      entry.last_req = rq;
      // Keep the sample from the most recent occurrence (coding fields can drift).
      entry.sample = {
        cost_category: row.cost_category,
        cost_code: row.cost_code,
        bank_cost_category: row.bank_cost_category,
        gl_coding: row.gl_coding,
        cost_code_name: row.cost_code_name,
      };
    }
  };

  const vendorTokensMap = new Map(); // norm -> Set(tokens)
  for (const r of rows) {
    const v = r.vendor_norm || '';
    const sig = r.bill_signature || '';
    bump(keyMap, v + '|' + sig, r);
    bump(vendorMap, v, r);
    if (v && !vendorTokensMap.has(v)) vendorTokensMap.set(v, tokenSet(v));
  }
  const vendorTokens = Array.from(vendorTokensMap.entries()).map(([norm, tokens]) => ({ norm, tokens }));
  return { keyMap, vendorMap, vendorTokens };
}

// Tokenize a normalized vendor string into a Set of word tokens (>=2 chars).
function tokenSet(normalized) {
  return new Set((String(normalized || '').match(/[a-z0-9]{2,}/g) || []));
}

// Token-overlap similarity between two token Sets: |intersection| / |smaller set|.
// 1.0 means one set is a subset of the other (e.g. "mullins law group" vs
// "mullins law group pllc" after suffix-strip, or any subset relationship).
function tokenSimilarity(a, b) {
  if (!a.size || !b.size) return 0;
  let inter = 0;
  for (const t of a) if (b.has(t)) inter++;
  return inter / Math.min(a.size, b.size);
}

// Rank a Map(cost_code -> {weight, last_req, sample}) into a list.
// For split vendors the MOST RECENT coding wins (highest last_req); weight is
// the tiebreaker when two codes share the same most-recent requisition.
function rankCodes(codes) {
  return Array.from(codes.values())
    .map((e) => ({ cost_code: e.sample.cost_code, weight: e.weight, last_req: e.last_req, sample: e.sample }))
    .sort((a, b) => (b.last_req - a.last_req) || (b.weight - a.weight));
}

// Predict the cost code for one invoice line.
// `line` needs at least { vendor, bill_number }.
// Returns { confidence, cost_code, coding, candidates } where:
//   coding     = full coding fields to copy onto the line (null when unmatched)
//   candidates = ranked list of historical codings (most-recent first)
// Matching is by normalized vendor, then by TOKEN SIMILARITY (so e.g. an extra
// "PLLC"/"PC" token, a dropped word, or reordering still matches the same
// vendor). When a vendor's history is split across cost codes, the most recent
// coding is proposed. Every prediction is editable in the UI; nothing is final
// until the user clicks Roll Forward & Download.
const SIMILARITY_THRESHOLD = 0.6;
function predict(line, index) {
  const v = normVendor(line.vendor);
  const sig = billSig(line.bill_number);
  const empty = { confidence: 'new', cost_code: null, coding: null, candidates: [] };
  if (!v) return empty;

  // 1) Exact vendor + bill-signature match (strongest signal).
  const keyCodes = index.keyMap.get(v + '|' + sig);
  if (keyCodes && keyCodes.size > 0) {
    const ranked = rankCodes(keyCodes);
    return {
      confidence: keyCodes.size === 1 ? 'high' : 'review',
      cost_code: ranked[0].cost_code,
      coding: ranked[0].sample,
      candidates: ranked.map((r) => r.sample),
    };
  }

  // 2) Exact vendor-only match.
  const vCodes = index.vendorMap.get(v);
  if (vCodes && vCodes.size > 0) {
    const ranked = rankCodes(vCodes);
    return {
      confidence: vCodes.size === 1 ? 'high' : 'review',
      cost_code: ranked[0].cost_code,
      coding: ranked[0].sample,
      candidates: ranked.map((r) => r.sample),
    };
  }

  // 3) Similar (token-based) vendor match. Pick the historical vendor whose
  // token set overlaps ours the most, above the threshold; ties broken by the
  // larger overlap then by recency of its coding.
  const myTokens = tokenSet(v);
  if (myTokens.size && Array.isArray(index.vendorTokens)) {
    let best = null, bestSim = 0;
    for (const cand of index.vendorTokens) {
      const sim = tokenSimilarity(myTokens, cand.tokens);
      if (sim > bestSim) { bestSim = sim; best = cand; }
    }
    if (best && bestSim >= SIMILARITY_THRESHOLD) {
      const codes = index.vendorMap.get(best.norm);
      if (codes && codes.size > 0) {
        const ranked = rankCodes(codes);
        return {
          confidence: 'review',
          cost_code: ranked[0].cost_code,
          coding: ranked[0].sample,
          candidates: ranked.map((r) => r.sample),
        };
      }
    }
  }

  // 4) Unseen vendor.
  return empty;
}

// Record one coded invoice line into history so future months learn from it.
// `weight` lets callers bias recent periods higher (default 1).
function recordHistory(db, entityId, line, reqNumber, weight) {
  db.prepare(
    'INSERT INTO requisition_coding_history ' +
    '(entity_id, vendor_norm, bill_signature, cost_category, cost_code, bank_cost_category, gl_coding, cost_code_name, req_number, weight, created_at) ' +
    'VALUES (?,?,?,?,?,?,?,?,?,?,?)'
  ).run(
    entityId,
    normVendor(line.vendor),
    billSig(line.bill_number),
    line.cost_category != null ? String(line.cost_category) : null,
    line.cost_code != null ? String(line.cost_code) : null,
    line.bank_cost_category != null ? String(line.bank_cost_category) : null,
    line.gl_coding != null ? String(line.gl_coding) : null,
    line.cost_code_name != null ? String(line.cost_code_name) : null,
    reqNumber != null ? Number(reqNumber) : null,
    weight != null ? Number(weight) : 1,
    new Date().toISOString()
  );
}

module.exports = {
  normVendor,
  billSig,
  buildHistoryIndex,
  predict,
  recordHistory,
};
