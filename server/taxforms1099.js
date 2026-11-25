// Accounts Payable: 1099 Reporting
//
// Per-entity, per-calendar-year 1099 workfile. Identifies vendors paid during
// the year and the amount, flags reportable payees, and (Phase 2) enriches each
// vendor with the TIN / 1099 flag / address captured in Bill.com.
//
// Amounts come from the A/P control account (billcom_config.default_ap_account,
// else 202000), grouped by journal_entries.vendor for entries dated in the year:
//   paid   = SUM(debits)  on A/P lines  -> cash applied to the vendor (1099 basis)
//   billed = SUM(credits) on A/P lines  -> bills booked to the vendor (accrual)
//   net    = billed - paid              -> change in payable
// Default basis is PAID (cash), the IRS 1099 basis. ?basis=billed thresholds on
// bills instead. Vendors paid straight from cash with no A/P entry do not appear.
//
// Reportability uses the Bill.com track1099 flag once vendor tax data is synced
// (POST /api/ap/form1099/:entity_id/sync-vendor-tax in index.js). REPORTABLE =
// 1099-flagged in Bill.com AND at/above threshold; over-threshold with no TIN =
// NEEDS W-9. Threshold defaults to $2,000 (OBBBA 2026+); override via ?threshold=.
//
// Route (Admin + Accountant, entity-scoped):
//   GET /api/ap/form1099/:entity_id?year=YYYY&threshold=2000&basis=paid

const DEFAULT_THRESHOLD = 2000;
const normKey = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');

function buildForm1099(db, entityId, year, threshold, basis) {
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(entityId);
  if (!ent) return { error: 'entity_not_found' };

  let apAccount = '202000';
  try {
    const cfg = db.prepare('SELECT default_ap_account FROM billcom_config WHERE entity_id = ?').get(entityId);
    if (cfg && cfg.default_ap_account) apAccount = String(cfg.default_ap_account);
  } catch (e) { /* not every entity has billcom_config */ }

  const y = Number(year);
  const from = y + '-01-01';
  const to = y + '-12-31';
  const useBasis = basis === 'billed' ? 'billed' : 'paid';

  const rows = db.prepare(
    `SELECT je.vendor AS vendor,
            SUM(jl.credit) AS billed,
            SUM(jl.debit)  AS paid,
            SUM(jl.credit) - SUM(jl.debit) AS net,
            COUNT(DISTINCT je.id) AS txns
       FROM journal_entries je
       JOIN journal_lines jl ON jl.entry_id = je.id
      WHERE je.entity_id = ?
        AND je.vendor IS NOT NULL AND TRIM(je.vendor) <> ''
        AND je.date >= ? AND je.date <= ?
        AND jl.account_code = ?
      GROUP BY je.vendor
      HAVING SUM(jl.credit) <> 0 OR SUM(jl.debit) <> 0
      ORDER BY paid DESC`
  ).all(entityId, from, to, apAccount);

  const taxByKey = new Map();
  let taxSyncedAt = null;
  try {
    for (const t of db.prepare('SELECT * FROM vendor_tax WHERE entity_id = ?').all(entityId)) {
      taxByKey.set(t.name_key, t);
      if (!taxSyncedAt || (t.synced_at && t.synced_at > taxSyncedAt)) taxSyncedAt = t.synced_at;
    }
  } catch (e) { /* vendor_tax may not exist yet */ }

  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const thr = Number(threshold) > 0 ? Number(threshold) : DEFAULT_THRESHOLD;

  const out = rows.map((r) => {
    const paid = r2(r.paid), billed = r2(r.billed);
    const amt = useBasis === 'billed' ? billed : paid;
    const t = taxByKey.get(normKey(r.vendor)) || null;
    const track1099 = t && t.track1099 != null ? (t.track1099 ? 1 : 0) : null;
    const w9 = !!(t && t.tax_id_last4);
    const over = amt >= thr;
    return {
      vendor: r.vendor, paid, billed, net: r2(r.net), txns: r.txns, over,
      tin_last4: t ? (t.tax_id_last4 || null) : null,
      tin_type: t ? (t.tax_id_type || null) : null,
      track1099,
      w9_on_file: w9,
      reportable: track1099 === 1 && over,
      needs_w9: over && track1099 !== 0 && !w9,
      address: t ? (t.address || null) : null,
    };
  });

  const reportable = out.filter((r) => r.reportable);
  const amtOf = (r) => (useBasis === 'billed' ? r.billed : r.paid);
  return {
    entity_id: ent.id, entity_name: ent.name, year: y, threshold: thr,
    ap_account: apAccount, basis: useBasis,
    tax_synced: taxByKey.size > 0, tax_synced_at: taxSyncedAt, tax_vendor_count: taxByKey.size,
    rows: out,
    summary: {
      vendor_count: out.length,
      over_count: out.filter((r) => r.over).length,
      reportable_count: reportable.length,
      reportable_total: r2(reportable.reduce((s, r) => s + amtOf(r), 0)),
      needs_w9_count: out.filter((r) => r.needs_w9).length,
      unmatched_count: out.filter((r) => r.track1099 === null).length,
    },
  };
}

function registerForm1099Routes(app, { db, auth, requireRole, requireEntityAccess }) {
  app.get('/api/ap/form1099/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
    const entityId = Number(req.params.entity_id);
    const year = Number(req.query.year) || (new Date().getUTCFullYear());
    const threshold = req.query.threshold != null ? Number(req.query.threshold) : DEFAULT_THRESHOLD;
    try {
      const result = buildForm1099(db, entityId, year, threshold, req.query.basis);
      if (result.error) return res.status(404).json(result);
      res.json(result);
    } catch (e) {
      res.status(500).json({ error: String((e && e.message) || e) });
    }
  });
}

module.exports = { registerForm1099Routes, buildForm1099 };
