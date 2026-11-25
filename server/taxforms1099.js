// ─── Accounts Payable: 1099 Reporting ────────────────────────────────────────
//
// Per-entity, per-calendar-year 1099 workfile. Identifies vendors to whom the
// entity booked payables during the year and the amount, so reportable payees
// (>= the IRS threshold) can be furnished and filed in-house — the data that
// previously justified CLA's per-entity 1099 prep fee now lives here.
//
// Amount basis (MVP): GROSS bills booked to the A/P control account per vendor
//   amount_billed(vendor) = SUM(journal_lines.credit) on A/P-control lines of
//   journal entries dated in the year whose journal_entries.vendor = vendor.
//   This is complete for AP-routed spend and is reversal / credit-memo aware
//   (a reversing or credit-memo entry debits A/P, lowering the net column).
//   Vendors paid straight from cash with no A/P bill do not appear — a known
//   MVP limitation, surfaced in the UI.
//
// Threshold: defaults to $2,000 (OBBBA, payments on/after 2026-01-01; prior
//   years used $600). Caller may override via ?threshold=.
//
// Phase 2 (not here): enrich each vendor with taxId / track1099 / W-9 address
//   pulled from the Bill.com vendor record, exclude card/third-party-network
//   payments, and export in IRS e-file (IRIS) format.
//
// Route (Admin + Accountant, entity-scoped):
//   GET /api/ap/form1099/:entity_id?year=YYYY&threshold=2000  -> JSON workfile

const DEFAULT_THRESHOLD = 2000;

function buildForm1099(db, entityId, year, threshold) {
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(entityId);
  if (!ent) return { error: 'entity_not_found' };

  // A/P control account: per-entity Bill.com config override, else 202000.
  let apAccount = '202000';
  try {
    const cfg = db.prepare('SELECT default_ap_account FROM billcom_config WHERE entity_id = ?').get(entityId);
    if (cfg && cfg.default_ap_account) apAccount = String(cfg.default_ap_account);
  } catch (e) { /* billcom_config may not exist for every entity */ }

  const y = Number(year);
  const from = y + '-01-01';
  const to = y + '-12-31';

  const rows = db.prepare(
    `SELECT je.vendor AS vendor,
            SUM(jl.credit) AS billed,
            SUM(jl.credit) - SUM(jl.debit) AS net,
            COUNT(DISTINCT je.id) AS txns
       FROM journal_entries je
       JOIN journal_lines jl ON jl.entry_id = je.id
      WHERE je.entity_id = ?
        AND je.vendor IS NOT NULL AND TRIM(je.vendor) <> ''
        AND je.date >= ? AND je.date <= ?
        AND jl.account_code = ?
      GROUP BY je.vendor
      HAVING SUM(jl.credit) <> 0
      ORDER BY billed DESC`
  ).all(entityId, from, to, apAccount);

  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const thr = Number(threshold) > 0 ? Number(threshold) : DEFAULT_THRESHOLD;

  const out = rows.map((r) => {
    const billed = r2(r.billed);
    return { vendor: r.vendor, billed, net: r2(r.net), txns: r.txns, over: billed >= thr };
  });

  const reportable = out.filter((r) => r.over);
  return {
    entity_id: ent.id,
    entity_name: ent.name,
    year: y,
    threshold: thr,
    ap_account: apAccount,
    basis: 'Gross bills booked to A/P control account ' + apAccount + ' (accrual).',
    rows: out,
    summary: {
      vendor_count: out.length,
      reportable_count: reportable.length,
      reportable_total: r2(reportable.reduce((s, r) => s + r.billed, 0)),
    },
  };
}

function registerForm1099Routes(app, { db, auth, requireRole, requireEntityAccess }) {
  app.get('/api/ap/form1099/:entity_id', auth, requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), (req, res) => {
    const entityId = Number(req.params.entity_id);
    const year = Number(req.query.year) || (new Date().getUTCFullYear());
    const threshold = req.query.threshold != null ? Number(req.query.threshold) : DEFAULT_THRESHOLD;
    try {
      const result = buildForm1099(db, entityId, year, threshold);
      if (result.error) return res.status(404).json(result);
      res.json(result);
    } catch (e) {
      res.status(500).json({ error: String((e && e.message) || e) });
    }
  });
}

module.exports = { registerForm1099Routes, buildForm1099 };
