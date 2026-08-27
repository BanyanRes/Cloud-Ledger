import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/index.js'
src = open(p, encoding='utf-8').read()

# Insert new routes right after the existing turnkey health endpoint
anchor = "// Health check (no auth — useful for Turnkey to verify connectivity)\napp.get('/api/turnkey/health', (req, res) => {\n  res.json({ status: 'ok', integration: 'turnkey-rail', timestamp: new Date().toISOString() });\n});"

new_routes = """// Health check (no auth — useful for Turnkey to verify connectivity)
app.get('/api/turnkey/health', (req, res) => {
  res.json({ status: 'ok', integration: 'turnkey-rail', timestamp: new Date().toISOString() });
});

// === Turnkey integration config (admin via JWT) ===
//
// Sets the company entity (the single "Turnkey Rail" entity that holds all
// project activity). Must be set before any project linking or sync events.
app.get('/api/admin/turnkey/config', auth, requireRole('Admin'), (req, res) => {
  const row = db.prepare('SELECT * FROM turnkey_config WHERE id = 1').get();
  res.json(row || { id: 1, enabled: 0, default_entity_id: null });
});

app.put('/api/admin/turnkey/config', auth, requireRole('Admin'), (req, res) => {
  const enabled = req.body.enabled ? 1 : 0;
  const entityId = req.body.default_entity_id != null ? Number(req.body.default_entity_id) : null;
  // Validate the entity exists
  if (entityId != null) {
    const ent = db.prepare('SELECT id FROM entities WHERE id = ?').get(entityId);
    if (!ent) return res.status(400).json({ error: 'default_entity_id refers to a non-existent entity' });
  }
  const now = new Date().toISOString();
  db.prepare(
    'INSERT INTO turnkey_config (id, enabled, default_entity_id, updated_by, updated_at) ' +
    'VALUES (1, ?, ?, ?, ?) ' +
    'ON CONFLICT(id) DO UPDATE SET enabled = excluded.enabled, default_entity_id = excluded.default_entity_id, ' +
    '  updated_by = excluded.updated_by, updated_at = excluded.updated_at'
  ).run(enabled, entityId, req.user.email, now);
  // Seed POC chart of accounts on the configured entity (idempotent)
  if (entityId != null) {
    const added = turnkey.seedPOCAccountsIfMissing(db, entityId);
    return res.json({ ok: true, enabled, default_entity_id: entityId, poc_accounts_added: added });
  }
  res.json({ ok: true, enabled, default_entity_id: entityId });
});

// === WIP Schedule (Job Schedule) endpoint ===
// Returns JSON. as_of query param defaults to today.
app.get('/api/turnkey/wip-schedule', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
    ? req.query.as_of
    : new Date().toISOString().slice(0, 10);
  try {
    const schedule = turnkey.computeWipSchedule(db, asOf);
    res.json(schedule);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Export the WIP schedule as Excel (.xlsx). Uses the existing 'xlsx' lib.
app.get('/api/turnkey/wip-schedule.xlsx', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
    ? req.query.as_of
    : new Date().toISOString().slice(0, 10);
  try {
    const schedule = turnkey.computeWipSchedule(db, asOf);
    const header = [
      'Job #', 'Job Name', 'Contract Amount', 'Revised Contract',
      'Costs to Date', 'Est Cost to Complete', 'Est Total Cost',
      'Est Gross Profit', '% Complete', 'Earned Revenue',
      'Billed to Date', 'Over/(Under) Billing'
    ];
    const dataRows = schedule.rows.map(r => [
      r.project_code || r.turnkey_project_id, r.project_name || '',
      r.contract_amount, r.revised_contract,
      r.costs_to_date, r.estimated_cost_to_complete, r.estimated_total_cost,
      r.estimated_gross_profit, r.percent_complete / 100, // store as fraction; format will render %
      r.earned_revenue, r.billed_to_date, r.over_under_billing,
    ]);
    const t = schedule.total;
    const totalRow = [
      'TOTAL', '',
      t.contract_amount, t.revised_contract,
      t.costs_to_date, t.estimated_cost_to_complete, t.estimated_total_cost,
      t.estimated_gross_profit, '',
      t.earned_revenue, t.billed_to_date, t.over_under_billing,
    ];
    const aoa = [
      ['Turnkey Rail \u2014 WIP Schedule'],
      ['As of:', asOf],
      [],
      header,
      ...dataRows,
      totalRow,
    ];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Column widths
    ws['!cols'] = [
      { wch: 12 }, { wch: 28 }, { wch: 16 }, { wch: 16 },
      { wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 16 },
      { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 20 },
    ];
    // Number formats for numeric cols
    const moneyFmt = '#,##0.00;(#,##0.00)';
    const pctFmt = '0.0%';
    const numericCols = [2,3,4,5,6,7,9,10,11];
    const headerRowIdx = 3; // 0-based, '$Job #' is row 4 in spreadsheet (after title+as_of+blank)
    const firstDataRow = headerRowIdx + 1;
    for (let i = 0; i < dataRows.length + 1; i++) { // +1 for total row
      for (const c of numericCols) {
        const addr = XLSX.utils.encode_cell({ r: firstDataRow + i, c });
        if (ws[addr]) ws[addr].z = moneyFmt;
      }
      // % column (index 8) — only data rows, not total
      if (i < dataRows.length) {
        const addr = XLSX.utils.encode_cell({ r: firstDataRow + i, c: 8 });
        if (ws[addr]) ws[addr].z = pctFmt;
      }
    }
    // Title cell — bold-ish via merge
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }];
    XLSX.utils.book_append_sheet(wb, ws, 'WIP Schedule');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=\"WIP_Schedule_' + asOf + '.xlsx\"');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});"""

if anchor not in src:
    print('MISSING anchor'); sys.exit(1)
src = src.replace(anchor, new_routes, 1)
open(p, 'w', encoding='utf-8').write(src)
print('OK')
