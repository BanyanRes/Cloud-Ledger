// ─── CLIP Development Costs workpaper ────────────────────────────────────────
//
// Returns the total capitalized development cost for a development entity as of a
// date, matching the "Total Long Term Investments + Total Other Assets" figure
// that feeds the CLIP line of the CLRF valuation schedule.
//
// This is the pre-transfer land + hard-cost + soft-cost population ONLY. It
// deliberately EXCLUDES:
//   - placed-in-service fixed assets (15xxx railroad track, building, equipment)
//     and their accumulated depreciation (16xxx),
//   - cash, AR / allowance, prepaids, interest reserve held as a current asset,
//     and intercompany balances (18xxx / 19xxx).
//
// The account set below was validated against the CLIP Property Owner (entity 54)
// balance sheet as of 2026-03-31: Total Long Term Investments 45,755,262.80 +
// Total Other Assets 5,280,033.34 = 51,035,296.14.
//
// The two groups mirror the entity's balance-sheet presentation ("Long Term
// Investments" and "Other Assets"). Membership is declared here rather than
// inferred from account type/subtype, so a new development account shows up as an
// omission on the face of the report instead of silently changing the total.

const LONG_TERM_INVESTMENTS = [
  '11010', // Acquisitions Costs (Land Purchase)
  '11040', // Land Contract Payments
  '11050', // Other Land Costs
  '11211', // Future Expansion Project Costs
  '11230', // Other Construction Costs
];

const OTHER_ASSETS = [
  '11970', // Other Legal - Legal
  '12013', // Civil Engineering Plans
  '12115', // A&E
  '12230', // Professional Services - Accounting
  '12315', // Appraisal
  '12321', // Construction Period Interest
  '12343', // Loan Fees
  '12381', // Acquisition Fee
  '12596', // Closing Costs
  '12720', // Travel - Other Development Costs
  '12913', // Development Fee
];

const r2 = n => Math.round((Number(n) || 0) * 100) / 100;

// Build the dev-cost figure from a balance snapshot.
//   computeBalances(eid, { as_of }) -> [{ code, name, type, balance, ... }]
function buildDevCosts(computeBalances, eid, asOf) {
  const rows = computeBalances(eid, { as_of: asOf });
  const byCode = new Map(rows.map(r => [String(r.code), r]));

  const pick = codes => codes.map(code => {
    const row = byCode.get(code);
    return {
      code,
      name: row ? row.name : null,
      balance: row ? r2(row.balance) : 0,
      present: !!row,
    };
  });

  const lti = pick(LONG_TERM_INVESTMENTS);
  const oa = pick(OTHER_ASSETS);
  const ltiTotal = r2(lti.reduce((s, a) => s + a.balance, 0));
  const oaTotal = r2(oa.reduce((s, a) => s + a.balance, 0));
  const total = r2(ltiTotal + oaTotal);

  // Accounts declared in the report set but absent from the GL snapshot -
  // surfaced so a chart-of-accounts change is visible rather than silent.
  const missing = [...lti, ...oa].filter(a => !a.present).map(a => a.code);

  return {
    as_of: asOf,
    longTermInvestments: { accounts: lti, total: ltiTotal },
    otherAssets: { accounts: oa, total: oaTotal },
    totalDevelopmentCost: total,
    missingAccounts: missing,
  };
}

function registerDevCostsRoutes(app, ctx) {
  const { auth, requireEntityAccess, computeBalances } = ctx;

  // GET /api/entities/:eid/dev-costs?as_of=YYYY-MM-DD
  // Returns the two subtotals and the total capitalized development cost.
  app.get('/api/entities/:eid/dev-costs', auth, requireEntityAccess(), (req, res) => {
    try {
      const asOf = req.query && req.query.as_of;
      if (!asOf || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) {
        return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
      }
      const data = buildDevCosts(computeBalances, req.params.eid, asOf);
      res.json(data);
    } catch (e) {
      console.error('dev-costs failed:', e);
      res.status(500).json({ error: e.message });
    }
  });
}

module.exports = {
  registerDevCostsRoutes,
  buildDevCosts,
  LONG_TERM_INVESTMENTS,
  OTHER_ASSETS,
};
