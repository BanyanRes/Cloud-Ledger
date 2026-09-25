// ─── Monthly Closing Workpapers — CLA seed for Banyan Residential ─────────────
//
// The prepaid-expense items and the fixed-asset register exactly as they stand in
// CLA's July 2026 close workpapers (05 Prepaid Expenses Leadsheet, 07 Fixed Asset
// Leadsheet). Loaded once into cla_prepaid_items / cla_fixed_assets for the
// Banyan Residential entity the first time the workpaper is generated; after
// that the registers live in CloudLedger and are rolled forward each month.
//
// Conventions carried over from CLA's schedules:
//   • Prepaid opening_balance = the "Balance P0" column (balance at 12/31/2025).
//   • Prepaid amortization begins the month AFTER the policy start month (CLA's
//     Lloyd's policy, start 3/30/2026, first amortized in April).
//   • The IMA / Wichita renewal is seeded at the $343.52/month CLA actually
//     expensed (their sheet's formula shows 2922.25/12 but the booked expense
//     was 343.52 each month, with a $30.57 true-up), so the balance ties.
//   • IMA / Wichita start date is corrected to 6/16/2025 (CLA's cell held a
//     corrupt 1909 date; the description reads "6/16/2025 - 6/16/2026").
//   • Fixed-asset accum_dep_beg = accumulated depreciation at 12/31/2025 (the
//     schedule's "Accumulated Depr/Amort Beg Balance" column).
module.exports = {
  // Entity is matched by name (Banyan Residential), not by id, so this applies
  // in any environment where that entity exists.
  entityMatch: /^banyan\s*residential/i,

  prepaid: [
    // 12922 — Prepaid Insurance
    { account_code: '12922', date_paid: '2024-08-06', vendor: 'Armstrong', expense_account: '65000', description: 'GL (Inv#28344), Commercial (Inv#283445)', start_date: '2024-09-01', end_date: '2025-09-01', monthly: 354.73, opening_balance: 0, premium: null },
    { account_code: '12922', date_paid: '2024-08-07', vendor: 'Armstrong', expense_account: '65000', description: 'Armstrong was originally 26k -15k when Rc', start_date: '2024-09-30', end_date: '2025-03-30', monthly: 2194.88, opening_balance: 0, premium: null },
    { account_code: '12922', date_paid: '2025-04-22', vendor: 'Armstrong', expense_account: '65000', description: '', start_date: '2025-03-30', end_date: '2026-03-30', monthly: 2303.39, opening_balance: 3307.75, premium: null },
    { account_code: '12922', date_paid: '2025-07-17', vendor: 'IMA, Inc - Wichita', expense_account: '65000', description: 'Renewal Property - Buna (inc fees); Policy#: IMA432919A; 6/16/2025 - 6/16/2026', start_date: '2025-06-16', end_date: '2026-06-16', monthly: 343.52, opening_balance: 1061.13, premium: null },
    { account_code: '12922', date_paid: '2025-08-14', vendor: 'Armstrong', expense_account: '65000', description: 'GL (Inv#28344), Commercial (Inv#283445)', start_date: '2025-09-01', end_date: '2026-09-01', monthly: 701.15, opening_balance: 5608.59, premium: null },
    { account_code: '12922', date_paid: '2025-09-02', vendor: 'Armstrong', expense_account: '65000', description: 'Inv#30126, Cyber Liability - Coalition', start_date: '2025-09-21', end_date: '2026-09-21', monthly: 267.00, opening_balance: 2403.00, premium: null },
    { account_code: '12922', date_paid: '2025-09-08', vendor: 'Armstrong', expense_account: '65000', description: 'Inv#30141, Cyber Liability - Ategrity', start_date: '2025-09-21', end_date: '2026-09-21', monthly: 216.58, opening_balance: 1949.20, premium: null },
    { account_code: '12922', date_paid: '2026-03-01', vendor: 'Armstrong', expense_account: '65000', description: "Certain Underwriters at Lloyd's London", start_date: '2026-03-30', end_date: '2027-03-30', monthly: 2307.69, opening_balance: 0, premium: 27692.31 },
    // 13003 — Prepaid Rent (reclassed from 61000 by JE dated 7/31/2026; no
    // amortization schedule set — monthly 0 until start/end/monthly are filled in)
    { account_code: '13003', date_paid: '2026-07-31', vendor: 'Continental', expense_account: '61000', description: 'Rent - paid to Continental (reclass from 61000, JE 7/31/2026)', start_date: null, end_date: null, monthly: 0, opening_balance: 0, premium: 7408.05 },
  ],

  fixedAssets: [
    // 15500 Equipment / 16500 Accumulated Depreciation
    { asset_account: '15500', dep_account: '16500', description: '02/18/2019 - HP *HP.COM STORE - BENJAMIN BROSSEAU', life_years: 5, in_service: '2019-02-26', cost: 1306.00, accum_dep_beg: 1306.00 },
    { asset_account: '15500', dep_account: '16500', description: '3/4/2026 - PC Express Online/Ecom - MELBEN LAPTOP', life_years: 0.8333, in_service: '2026-03-04', cost: 1459.34, accum_dep_beg: 0 },
    { asset_account: '15500', dep_account: '16500', description: "4/6/2026 - SHOPEE PH - DIANNE'S LAPTOP", life_years: 0.75, in_service: '2026-04-06', cost: 1198.01, accum_dep_beg: 0 },
    // 12383 Organization Fee / 12384 Accumulated Amortization
    { asset_account: '12383', dep_account: '12384', description: 'Organization Cost', life_years: 15, in_service: '2019-01-01', cost: 12000.00, accum_dep_beg: 5600.00 },
    // 15510 Vehicles / 16510 Accumulated Depreciation
    { asset_account: '15510', dep_account: '16510', description: '2022 Chevy Tahoe', life_years: 0.4167, in_service: '2025-08-11', cost: 42586.49, accum_dep_beg: 38327.49 },
  ],
};
