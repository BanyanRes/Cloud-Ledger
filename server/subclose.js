// ─── CLRF workpaper: Subsequent-Closing (Sub Close) workpaper ────────────────
//
// The Legacy Knight subsequent-closing / equalization workpaper for County Line
// Rail Fund I, LP (CL entity 40), reproduced from the general ledger. Mirrors the
// fund administrator's "subclose workpaper": a per-investor rebalance (Sub Close
// Calcu) plus the subscriber calls and posted JEs (Sub Close Summary). Prepared
// like the other CLRF workpapers (pcap.js / gpfees.js): buildData -> buildWorkbook
// -> saveToWorkpapers, registered by index.js.
//
// Every rebalance column is GL-derived, keyed to the journal-entry line
// descriptions the subclose postings carry:
//   GCM Equalization      = '%equalization distribution per GCM%'
//   Odyssey Correcting     = '%true-up due to Odyssey commitment overstatement%'
//   Legacy Knight Equaliz. = '%changes in contribution due to Legacy Knight%'
//   Capital Call (May-26)  = '%capital call for 2 millions%' + '%true up for roundings%'
//   Ending Capital         = net balance of the contribution accounts (= PCAP
//                            "Contributed capital"; ties the workpaper to the PCAP
//                            statements and to the fund Statement of Changes).
// Commitment / Unfunded come from investor_commitments and the contribution-account
// balance, so Odyssey ties to its corrected PCAP figures once its commitment is set.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const FUND_EID = 40;
const CONTRIB_ACCTS = ['30100', '301100', '301200', '301300', '301800'];
const CONTRIB_SQL = CONTRIB_ACCTS.map((c) => "'" + c + "'").join(',');
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

// Description buckets that make up the rebalance columns.
const BUCKETS = {
  gcmEqualization: '%equalization distribution per GCM%',
  odysseyCorrecting: '%true-up due to Odyssey commitment overstatement%',
  lkEqualization: '%changes in contribution due to Legacy Knight%',
  jnbEqualization: '%James and Nat%',
  capitalCall2m: '%capital call for 2 millions%',
  roundingTrueup: '%true up for roundings%',
};

function bucketByClass(db, eid, asOf, pattern) {
  const rows = db.prepare(
    "SELECT jl.class_id AS cid, SUM(jl.credit) - SUM(jl.debit) AS net "
    + "FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id "
    + "WHERE je.entity_id = ? AND je.date <= ? AND jl.account_code IN (" + CONTRIB_SQL + ") "
    + "AND jl.description LIKE ? GROUP BY jl.class_id"
  ).all(eid, asOf, pattern);
  const m = {}; rows.forEach((r) => { m[r.cid] = r.net; }); return m;
}

// A subscriber call entry (Legacy Knight / James Bloomingdale): contribution-account
// credits by category + the receivable, read from one journal entry.
function callEntry(db, eid, memoLike) {
  const je = db.prepare(
    "SELECT id, date, memo FROM journal_entries WHERE entity_id = ? AND memo LIKE ? ORDER BY date, id LIMIT 1"
  ).get(eid, memoLike);
  if (!je) return null;
  const rows = db.prepare(
    "SELECT jl.account_code AS code, a.name AS name, SUM(jl.debit) AS d, SUM(jl.credit) AS c "
    + "FROM journal_lines jl JOIN accounts a ON a.entity_id = ? AND a.code = jl.account_code "
    + "WHERE jl.entry_id = ? GROUP BY jl.account_code ORDER BY jl.account_code"
  ).all(eid, je.id);
  const lines = rows.map((r) => ({ code: r.code, name: r.name, debit: r2(r.d), credit: r2(r.c) }));
  const total = r2(lines.reduce((s, l) => s + l.debit, 0));
  return { id: je.id, date: je.date, memo: je.memo, lines, total };
}

// Account-level net summary of one entry (for the equalization JE), plus the
// per-investor cash distribution (100200 credits tagged to the entry).
function equalizationEntry(db, eid, memoLike) {
  const je = db.prepare(
    "SELECT id, date, memo FROM journal_entries WHERE entity_id = ? AND memo LIKE ? ORDER BY date, id LIMIT 1"
  ).get(eid, memoLike);
  if (!je) return null;
  const acct = db.prepare(
    "SELECT jl.account_code AS code, a.name AS name, SUM(jl.debit) AS d, SUM(jl.credit) AS c "
    + "FROM journal_lines jl JOIN accounts a ON a.entity_id = ? AND a.code = jl.account_code "
    + "WHERE jl.entry_id = ? GROUP BY jl.account_code ORDER BY jl.account_code"
  ).all(eid, je.id).map((r) => ({ code: r.code, name: r.name, debit: r2(r.d), credit: r2(r.c) }));
  const dist = db.prepare(
    "SELECT jl.class_id AS cid, SUM(jl.credit) - SUM(jl.debit) AS cash "
    + "FROM journal_lines jl WHERE jl.entry_id = ? AND jl.account_code LIKE '1002%' AND jl.class_id IS NOT NULL "
    + "GROUP BY jl.class_id"
  ).all(je.id);
  const distBy = {}; dist.forEach((r) => { distBy[r.cid] = r2(r.cash); });
  const totalD = r2(acct.reduce((s, l) => s + l.debit, 0));
  const totalC = r2(acct.reduce((s, l) => s + l.credit, 0));
  return { id: je.id, date: je.date, memo: je.memo, acct, distBy, totalD, totalC };
}

function buildData(ctx, asOf, opts = {}) {
  const { db, computeBalances } = ctx;
  if (!isDate(asOf)) throw new Error('as_of must be YYYY-MM-DD');
  const eid = opts.entity_id || FUND_EID;
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  const classes = db.prepare('SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?').all(eid);
  const commitRows = db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
  const commitBy = {}; commitRows.forEach((c) => { commitBy[c.class_id] = r2(c.commitment_amount); });

  const b = {};
  for (const k of Object.keys(BUCKETS)) b[k] = bucketByClass(db, eid, asOf, BUCKETS[k]);

  const investors = [];
  for (const c of classes) {
    const commitment = commitBy[c.id] || 0;
    const gcmEqualization = r2(b.gcmEqualization[c.id] || 0);
    const odysseyCorrecting = r2(b.odysseyCorrecting[c.id] || 0);
    const lkEqualization = r2(b.lkEqualization[c.id] || 0);
    const jnbEqualization = r2(b.jnbEqualization[c.id] || 0);
    const capitalCall = r2((b.capitalCall2m[c.id] || 0) + (b.roundingTrueup[c.id] || 0));
    // Ending capital = net contribution-account balance (= PCAP "Contributed capital").
    const contributed = r2((computeBalances(eid, { to: asOf, class_id: c.id }) || [])
      .filter((r) => CONTRIB_ACCTS.includes(String(r.code)))
      .reduce((s, r) => s + (Number(r.balance) || 0), 0));
    const anyActivity = commitment || contributed || gcmEqualization || lkEqualization || capitalCall;
    if (!anyActivity) continue;
    const unfunded = r2(commitment - contributed);
    // Recalc check: commitment rolled forward by the rebalance columns.
    const recalc = r2(commitment + gcmEqualization + odysseyCorrecting + lkEqualization + jnbEqualization + capitalCall);
    investors.push({
      class_id: c.id, name: c.name,
      partner_type: String(c.partner_type || '').toUpperCase() === 'GP' ? 'GP' : 'LP',
      commitment, gcmEqualization, odysseyCorrecting, lkEqualization, jnbEqualization,
      capitalCall, endingCapital: contributed, unfunded,
      recalc, recalcDiff: r2(recalc - contributed),
    });
  }
  investors.sort((a, z) => String(a.name).localeCompare(String(z.name)));

  const legacyKnightCall = callEntry(db, eid, '%LC05282601%');
  const jamesCall = callEntry(db, eid, '%contribution rec%James Bloomingdale%')
    || callEntry(db, eid, '%James Bloomingdale%');
  const equalization = equalizationEntry(db, eid, '%LC05282603%');
  // Attach per-investor equalization cash to each investor row.
  if (equalization) for (const inv of investors) inv.equalizationCash = r2(equalization.distBy[inv.class_id] || 0);

  const sum = (k) => r2(investors.reduce((s, i) => s + (i[k] || 0), 0));
  const totals = {
    count: investors.length,
    commitment: sum('commitment'), gcmEqualization: sum('gcmEqualization'),
    odysseyCorrecting: sum('odysseyCorrecting'), lkEqualization: sum('lkEqualization'),
    jnbEqualization: sum('jnbEqualization'), capitalCall: sum('capitalCall'),
    endingCapital: sum('endingCapital'), unfunded: sum('unfunded'),
    equalizationCash: sum('equalizationCash'),
  };
  return {
    entity_id: eid, entity_name: ent ? ent.name : ('entity ' + eid), as_of: asOf,
    investors, totals, legacyKnightCall, jamesCall, equalization,
  };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const NAVY = 'FF1F3864';
const HDR = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const MONEY = '#,##0.00;(#,##0.00);-';
const THIN = { style: 'thin' };

function hdrRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true, vertical: 'bottom' };
  });
  if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

function buildWorkbook(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  const dateStr = data.as_of;

  // ── Tab 1: Sub Close Calcu — per-investor rebalance ──────────────────────────
  const ws = wb.addWorksheet('Sub Close Calcu', { views: [{ state: 'frozen', xSplit: 1, ySplit: 4, showGridLines: false }] });
  ws.getCell('A1').value = data.entity_name + ' — Subsequent-Closing Rebalance (Legacy Knight)';
  ws.getCell('A1').font = F({ size: 12, bold: true });
  ws.getCell('A2').value = 'As of ' + dateStr + '. GL-derived; ties to the PCAP statements. Not for income tax purposes.';
  ws.getCell('A2').font = F({ size: 9, italic: true });
  const cols = ['Investor Name', 'Commitment', 'Distribution — GCM Equalization',
    'Correcting Distribution (Odyssey overstatement)', 'Distribution — Legacy Knight Equalization',
    'Distribution — James & Natalie Bloomingdale', 'May 2026 — Capital Call', 'Ending Capital (Contributed)',
    'Unfunded', 'Equalization Cash to/(from) Investor', 'Recalc Check', 'Diff'];
  hdrRow(ws, 4, cols, [40, 15, 18, 18, 18, 16, 15, 16, 14, 16, 15, 10]);
  const keys = ['commitment', 'gcmEqualization', 'odysseyCorrecting', 'lkEqualization', 'jnbEqualization',
    'capitalCall', 'endingCapital', 'unfunded', 'equalizationCash', 'recalc', 'recalcDiff'];
  let r = 5;
  for (const inv of data.investors) {
    ws.getCell('A' + r).value = inv.name; ws.getCell('A' + r).font = F();
    keys.forEach((k, i) => { const c = ws.getCell(r, i + 2); c.value = inv[k]; c.numFmt = MONEY; c.font = F(); });
    r++;
  }
  const t = data.totals; const tr = ws.getRow(r);
  tr.getCell(1).value = 'Total — ' + t.count + ' investors'; tr.getCell(1).font = F({ bold: true });
  const tvals = [t.commitment, t.gcmEqualization, t.odysseyCorrecting, t.lkEqualization, t.jnbEqualization,
    t.capitalCall, t.endingCapital, t.unfunded, t.equalizationCash, null, null];
  tvals.forEach((v, i) => { if (v === null) return; const c = tr.getCell(i + 2); c.value = v; c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN, bottom: { style: 'double' } }; });

  // ── Tab 2: Sub Close Summary — subscriber calls, equalization JE, distribution ─
  const sm = wb.addWorksheet('Sub Close Summary', { views: [{ showGridLines: false }] });
  sm.getColumn(1).width = 44; sm.getColumn(2).width = 18; sm.getColumn(3).width = 18; sm.getColumn(4).width = 40;
  let R = 1;
  const put = (col, v, o = {}, fmt) => { const c = sm.getCell(col + R); c.value = v; c.font = F(o); if (fmt) c.numFmt = fmt; return c; };
  const callBlock = (title, call) => {
    put('A', title, { bold: true, size: 11 }); R++;
    if (!call) { put('A', '(entry not found in GL)', { italic: true }); R += 2; return; }
    put('A', 'JE' + call.id + '  ' + String(call.date).slice(0, 10)); R++;
    put('A', 'Account', { bold: true }); put('B', 'Debit', { bold: true }); put('C', 'Credit', { bold: true }); R++;
    for (const l of call.lines) { put('A', l.code + ' ' + (l.name || '')); put('B', l.debit || null, {}, MONEY); put('C', l.credit || null, {}, MONEY); R++; }
    put('A', 'Total call', { bold: true }); put('B', call.total, { bold: true }, MONEY); R += 2;
  };
  put('A', data.entity_name + ' — Subsequent-Closing Summary (Legacy Knight)', { bold: true, size: 12 }); R += 2;
  callBlock('Legacy Knight — Capital Call (LC05282601)', data.legacyKnightCall);
  callBlock('James Bloomingdale — Capital Call (LC05282602)', data.jamesCall);

  if (data.equalization) {
    put('A', 'Equalization / true-up entry — JE' + data.equalization.id + '  ' + String(data.equalization.date).slice(0, 10), { bold: true, size: 11 }); R++;
    put('A', 'Account', { bold: true }); put('B', 'Debit', { bold: true }); put('C', 'Credit', { bold: true }); R++;
    for (const l of data.equalization.acct) { put('A', l.code + ' ' + (l.name || '')); put('B', l.debit || null, {}, MONEY); put('C', l.credit || null, {}, MONEY); R++; }
    put('A', 'Total', { bold: true }); put('B', data.equalization.totalD, { bold: true }, MONEY); put('C', data.equalization.totalC, { bold: true }, MONEY); R += 2;

    put('A', 'Equalization distribution — cash to/(from) each investor', { bold: true, size: 11 }); R++;
    put('A', 'Investor', { bold: true }); put('B', 'Cash', { bold: true }); R++;
    let dt = 0;
    for (const inv of data.investors) {
      if (!inv.equalizationCash) continue;
      put('A', inv.name); put('B', inv.equalizationCash, {}, MONEY); R++; dt = r2(dt + inv.equalizationCash);
    }
    put('A', 'Total distribution', { bold: true }); put('B', dt, { bold: true }, MONEY); R += 1;
  }
  return wb;
}

// ── Persistence + route (mirror pcap.js). ─────────────────────────────────────
const folderFor = (asOf) => 'Workpapers/Subsequent Closings/' + String(asOf).slice(0, 4);
const fileNameFor = (asOf) => 'CLRF_SubClose_' + String(asOf) + '.xlsx';

function saveToWorkpapers(ctx, eid, asOf, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(asOf);
  const original = fileNameFor(asOf);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, '
    + "uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function registerSubcloseRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/subclose/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const asOf = (req.body && (req.body.as_of || req.body.quarter_end)) || '';
        if (!isDate(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, asOf, { entity_id: eid });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, asOf, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-SubClose-Summary', JSON.stringify({
          as_of: asOf, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          investors: data.totals.count, ending_capital: data.totals.endingCapital,
          equalization_cash: data.totals.equalizationCash,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });

  // Download-only (regenerate on the fly) for the Reports view.
  app.get('/api/entities/:eid/subclose.xlsx', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const eid = Number(req.params.eid);
      const asOf = req.query && req.query.as_of;
      if (!isDate(asOf)) return res.status(400).json({ error: 'as_of (YYYY-MM-DD) is required' });
      const data = buildData(ctx, asOf, { entity_id: eid });
      const wb = buildWorkbook(data);
      const buf = Buffer.from(await wb.xlsx.writeBuffer());
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fileNameFor(asOf) + '"');
      res.send(buf);
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
}

module.exports = { buildData, buildWorkbook, saveToWorkpapers, registerSubcloseRoutes, FUND_EID };
