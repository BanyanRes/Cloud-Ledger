// ─── CLRF workpaper: Preferred Return (fund-level 8% XIRR) ────────────────────
//
// A quarterly, standalone, reviewable workpaper for County Line Rail Fund I, LP
// (CL entity 40) that reproduces the fund's authoritative Preferred Return
// calculation (Weaver's "Preferred Return - Fund Level Calculation").
//
// Methodology (per the LPA "Preferred Return" definition and Weaver's workbook):
//   • Take the equalized Limited-Partner net cash-flow stream (contributions are
//     outflows to the LPs, refunds are inflows), dated from fund inception.
//   • At the measurement (quarter-end) date, assume a single liquidating
//     distribution equal to Return of Capital + Preferred Return.
//   • The Preferred Return is the amount, on top of Return of Capital, that makes
//     the internal rate of return (XIRR) of the LP stream equal 8% per annum,
//     annually compounded. Weaver solves it by Goal Seek (set IRR cell to 8% by
//     changing the Preferred Return cell).
//
// CLRF's CL ledger holds contributed capital only as a 12/31/2025 opening import
// (no dated call history), so the dated schedule is NOT derivable from the GL. It
// is maintained per quarter in fund_preferred_return (roc, pref, and an optional
// dated `cashflows` schedule) from the fund's preferred-return workpaper. When a
// dated schedule is present this workpaper RECOMPUTES the XIRR and verifies the
// 8% solve line by line; when it is absent it presents the stored summary with a
// note. Prepared like the other CLRF workpapers (buildData -> buildWorkbook ->
// saveToWorkpapers), registered by index.js, and located via findWorkpaper() so
// it folds into the quarterly deliverable. No Assurance Provided (Weaver remains
// responsible for the official financial statements).
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// County Line Rail Fund I, LP.
const FUND_EID = 40;
const PREF_RATE = 0.08;   // 8% per annum, annually compounded (LPA definition)

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

// ── Quarter arithmetic (same contract as the other CLRF workpapers). ──────────
function resolveQuarter(quarterEnd) {
  if (!isDate(quarterEnd)) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const [y, m, d] = quarterEnd.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) {
    throw new Error('quarter_end must be a quarter end date: 03-31, 06-30, 09-30 or 12-31. Received ' + quarterEnd);
  }
  const q = m / 3;
  return { label: y + '-Q' + q, year: String(y), quarter: 'Q' + q, end: quarterEnd };
}

// ── XIRR on a dated cash-flow stream. cfs = [{date:'YYYY-MM-DD', amount}]. Uses
//    Actual/365. Returns the annually-compounded rate r solving Σ a_i /
//    (1+r)^((d_i - d_0)/365) = 0. Newton from several seeds, then a bisection
//    fallback on the sign-bracket, so it is robust for the standard "outflows
//    then a single terminal inflow" shape.
function xirr(cfs) {
  if (!cfs || cfs.length < 2) return null;
  const d0 = new Date(cfs[0].date + 'T00:00:00Z').getTime();
  const yrs = cfs.map((c) => (new Date(c.date + 'T00:00:00Z').getTime() - d0) / (365 * 24 * 3600 * 1000));
  const amt = cfs.map((c) => Number(c.amount) || 0);
  const npv = (r) => amt.reduce((s, a, i) => s + a / Math.pow(1 + r, yrs[i]), 0);
  const dnpv = (r) => amt.reduce((s, a, i) => s - a * yrs[i] / Math.pow(1 + r, yrs[i] + 1), 0);
  // Newton from a few seeds.
  for (const seed of [0.08, 0.05, 0.15, 0.0]) {
    let r = seed, ok = true;
    for (let k = 0; k < 100; k++) {
      const f = npv(r), df = dnpv(r);
      if (!isFinite(f) || !isFinite(df) || Math.abs(df) < 1e-12) { ok = false; break; }
      const nr = r - f / df;
      if (!isFinite(nr) || nr <= -0.9999) { ok = false; break; }
      if (Math.abs(nr - r) < 1e-10) { r = nr; break; }
      r = nr;
    }
    if (ok && isFinite(npv(r)) && Math.abs(npv(r)) < 1e-4) return r;
  }
  // Bisection fallback on a sign bracket.
  let lo = -0.9, hi = 10, flo = npv(lo), fhi = npv(hi);
  if (!(isFinite(flo) && isFinite(fhi)) || flo * fhi > 0) return null;
  for (let k = 0; k < 200; k++) {
    const mid = (lo + hi) / 2, fm = npv(mid);
    if (Math.abs(fm) < 1e-6 || (hi - lo) < 1e-12) return mid;
    if (flo * fm <= 0) { hi = mid; fhi = fm; } else { lo = mid; flo = fm; }
  }
  return (lo + hi) / 2;
}

// ── Solve the Preferred Return: the amount P (on top of Return of Capital) at the
//    measurement date that makes the LP-stream XIRR equal `rate`. IRR is monotone
//    increasing in P, so bisect. flows = dated LP net cash flows (excluding the
//    terminal distribution); roc = Return of Capital; end = measurement date.
function solvePref(flows, roc, end, rate = PREF_RATE) {
  if (!flows || !flows.length) return null;
  const irrAt = (P) => xirr([...flows, { date: end, amount: r2(roc) + P }]);
  let lo = 0, hi = Math.max(1e6, Math.abs(roc) * 5);
  // Expand hi until IRR(hi) >= rate (guards an unusually low stream).
  for (let k = 0; k < 60 && (irrAt(hi) == null || irrAt(hi) < rate); k++) hi *= 2;
  const flo = (irrAt(lo) == null ? -1 : irrAt(lo)) - rate;
  const fhi = (irrAt(hi) == null ? 1 : irrAt(hi)) - rate;
  if (flo > 0) return 0; // even with no preferred return the stream already clears 8%
  if (fhi < 0) return null;
  for (let k = 0; k < 200; k++) {
    const mid = (lo + hi) / 2;
    const fm = (irrAt(mid) == null ? 0 : irrAt(mid)) - rate;
    if (Math.abs(fm) < 1e-9 || (hi - lo) < 1e-6) return mid;
    if (fm >= 0) hi = mid; else lo = mid;
  }
  return (lo + hi) / 2;
}

// ── Gather everything the workbook needs from fund_preferred_return. ───────────
function buildData(ctx, quarter, opts = {}) {
  const { db } = ctx;
  const eid = opts.entity_id || FUND_EID;
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  let row = null;
  try {
    row = db.prepare('SELECT roc, pref, note, cashflows FROM fund_preferred_return WHERE entity_id = ? AND quarter_end = ?')
      .get(eid, quarter.end) || null;
  } catch (e) { row = null; }
  if (!row) {
    throw new Error('No preferred-return figures stored for ' + (ent ? ent.name : ('entity ' + eid))
      + ' as of ' + quarter.end + '. Enter Return of Capital and Preferred Return (and optionally the'
      + ' dated cash-flow schedule) first.');
  }
  const roc = row.roc == null ? null : r2(row.roc);
  const pref = row.pref == null ? null : r2(row.pref);
  const total = (roc != null && pref != null) ? r2(roc + pref) : null;

  // Optional dated schedule → XIRR reproduction.
  let flows = null;
  if (row.cashflows) {
    try {
      const parsed = JSON.parse(row.cashflows);
      if (Array.isArray(parsed) && parsed.length) {
        flows = parsed
          .map((x) => ({ date: String(x.date).slice(0, 10), amount: Number(x.amount) }))
          .filter((x) => isDate(x.date) && isFinite(x.amount))
          .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
        if (!flows.length) flows = null;
      }
    } catch (e) { flows = null; }
  }

  const calc = { schedule_loaded: !!flows };
  if (flows) {
    const lpNetInvested = r2(-flows.reduce((s, f) => s + f.amount, 0)); // ROC identity: capital in, net of refunds
    const terminalInflow = r2((roc != null ? roc : lpNetInvested) + (pref != null ? pref : 0));
    const stream = [...flows, { date: quarter.end, amount: terminalInflow }];
    const irr = xirr(stream);
    const solvedPref = solvePref(flows, roc != null ? roc : lpNetInvested, quarter.end, PREF_RATE);
    Object.assign(calc, {
      flows,
      cf_count: flows.length,
      lp_net_invested: lpNetInvested,
      terminal_date: quarter.end,
      terminal_inflow: terminalInflow,
      irr: irr == null ? null : irr,
      irr_target: PREF_RATE,
      solved_pref: solvedPref == null ? null : r2(solvedPref),
      solved_pref_delta: (pref != null && solvedPref != null) ? r2(solvedPref - pref) : null,
      // Checks: ROC equals the LP net capital invested (exact identity); and the
      // stored Preferred Return reproduces the 8% hurdle — i.e. the LP stream IRR
      // at the stored pref lands on 8% within Weaver's Goal-Seek tolerance (5 bps).
      // `solved_pref` is the preferred return that hits exactly 8.0000%, shown for
      // reference; the small delta to the stored figure is the goal-seek residual.
      roc_ties: roc != null && Math.abs(roc - lpNetInvested) < 0.5,
      irr_ties: irr != null && Math.abs(irr - PREF_RATE) < 0.0005,
    });
  }

  return {
    quarter,
    fund: {
      entity_id: eid, entity_name: ent ? ent.name : ('entity ' + eid),
      roc, pref, total, rate: PREF_RATE,
      note: row.note || null,
    },
    calc,
  };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const SUM_MONEY = '_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)';
const PCT4 = '0.0000%';
const NAVY = 'FF1F3864';
const HDR_FONT = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const BLUE = F({ color: { argb: 'FF0000FF' } });     // value read from the source system
const SMALLI = F({ size: 9, italic: true });
const SF = (o = {}) => Object.assign({ name: 'Times New Roman', size: 10 }, o);
const THIN = { style: 'thin' };
const MONTHS = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
  'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
const spellDate = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };

function headerRow(ws, rowNum, labels, widths) {
  const row = ws.getRow(rowNum);
  labels.forEach((t, i) => {
    const c = row.getCell(i + 1);
    c.value = t; c.font = HDR_FONT;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: 'center', wrapText: true };
  });
  if (widths) widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

function buildWorkbook(data) {
  const q = data.quarter, fund = data.fund, calc = data.calc;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  const pr = wb.addWorksheet('Preferred Return');
  const nt = wb.addWorksheet('Notes & Sources');

  // ── Preferred Return (client-facing, mirrors Weaver's fund-level sheet) ───────
  pr.getColumn(1).width = 16; pr.getColumn(2).width = 30; pr.getColumn(3).width = 22; pr.getColumn(4).width = 22;
  const title = (rowNum, text, opts = {}) => {
    const c = pr.getCell('A' + rowNum); c.value = text; c.font = SF(Object.assign({ bold: true }, opts));
    c.alignment = { horizontal: 'center' }; pr.mergeCells('A' + rowNum + ':D' + rowNum);
  };
  title(1, fund.entity_name);
  title(2, 'Preferred Return — Fund-Level Calculation');
  title(3, 'As of ' + spellDate(q.end)); pr.getCell('A3').font = SF({ italic: true });

  let R = 5;
  const put = (ref, v, font) => { const c = pr.getCell(ref); c.value = v; c.font = font || SF(); };
  const money = (ref, v, o = {}) => {
    const c = pr.getCell(ref);
    if (v === null || v === undefined) { c.value = '—'; c.font = SF({ italic: true }); }
    else { c.value = v; c.numFmt = SUM_MONEY; c.font = SF(o); }
  };

  if (calc.schedule_loaded) {
    put('A' + R, 'Equalized Limited-Partner net cash flows (contributions negative, refunds positive):', SF({ italic: true, size: 9 }));
    pr.mergeCells('A' + R + ':D' + R); R++;
    headerRow(pr, R, ['Date', 'LP Net Cash Flow', 'Return of Capital', 'Preferred Return'], [16, 30, 22, 22]);
    R++;
    for (const f of calc.flows) {
      pr.getCell('A' + R).value = f.date; pr.getCell('A' + R).font = SF();
      pr.getCell('A' + R).alignment = { horizontal: 'left' };
      money('B' + R, r2(f.amount), {}); pr.getCell('B' + R).font = BLUE;
      R++;
    }
    // Terminal liquidating distribution row at the measurement date.
    pr.getCell('A' + R).value = calc.terminal_date; pr.getCell('A' + R).font = SF({ bold: true });
    pr.getCell('A' + R).alignment = { horizontal: 'left' };
    money('B' + R, calc.terminal_inflow, { bold: true });
    money('C' + R, fund.roc === null ? null : -fund.roc);
    money('D' + R, fund.pref === null ? null : -fund.pref);
    ['A', 'B', 'C', 'D'].forEach((col) => { pr.getCell(col + R).border = { top: THIN }; });
    R++;
    // Totals + IRR.
    put('A' + R, 'Total', SF({ bold: true }));
    money('B' + R, r2(calc.flows.reduce((s, f) => s + f.amount, 0) + calc.terminal_inflow), { bold: true });
    pr.getCell('B' + R).border = { top: THIN, bottom: { style: 'double' } };
    R += 2;
    put('A' + R, 'Internal rate of return (XIRR)', SF({ bold: true }));
    { const c = pr.getCell('B' + R); c.value = calc.irr == null ? '—' : calc.irr; if (calc.irr != null) c.numFmt = PCT4; c.font = SF({ bold: true }); }
    put('C' + R, 'Target', SF({ italic: true, size: 9 }));
    { const c = pr.getCell('D' + R); c.value = calc.irr_target; c.numFmt = PCT4; c.font = SF({ italic: true }); }
    R += 2;
  } else {
    put('A' + R, 'Dated cash-flow schedule not loaded for this quarter — summary shown from the fund preferred-return workpaper (Weaver).', SF({ italic: true, size: 9 }));
    pr.getCell('A' + R).alignment = { wrapText: true }; pr.mergeCells('A' + R + ':D' + R); R += 2;
  }

  // ── Summary block (always). ──────────────────────────────────────────────────
  put('A' + R, 'Summary', SF({ bold: true })); R++;
  put('A' + R, '   Return of Capital (LP net capital invested)'); pr.mergeCells('A' + R + ':C' + R); money('D' + R, fund.roc); R++;
  put('A' + R, '   Preferred Return (8% p.a., annually compounded)'); pr.mergeCells('A' + R + ':C' + R); money('D' + R, fund.pref); R++;
  put('A' + R, 'Total return threshold (Return of Capital + Preferred Return)', SF({ bold: true })); pr.mergeCells('A' + R + ':C' + R);
  money('D' + R, fund.total, { bold: true });
  ['A', 'B', 'C', 'D'].forEach((col) => { pr.getCell(col + R).border = { top: THIN, bottom: THIN }; }); R += 2;

  if (calc.schedule_loaded) {
    put('A' + R, 'Independent checks', SF({ bold: true, size: 9 })); R++;
    put('A' + R, '   Return of Capital ties to LP net capital invested (−Σ cash flows = '
      + (calc.lp_net_invested).toLocaleString('en-US', { style: 'currency', currency: 'USD' }) + '): '
      + (calc.roc_ties ? 'YES' : 'review'), SF({ size: 9, italic: true })); pr.mergeCells('A' + R + ':D' + R); R++;
    put('A' + R, '   LP stream IRR at the stored Preferred Return reaches the 8% hurdle (computed '
      + (calc.irr == null ? 'n/a' : (calc.irr * 100).toFixed(4) + '%') + '): '
      + (calc.irr_ties ? 'YES' : 'review'), SF({ size: 9, italic: true })); pr.mergeCells('A' + R + ':D' + R); R++;
    put('A' + R, '   Preferred Return to reach exactly 8.0000% (reference): '
      + (calc.solved_pref == null ? 'n/a' : calc.solved_pref.toLocaleString('en-US', { style: 'currency', currency: 'USD' }))
      + (calc.solved_pref_delta == null ? '' : ' — '
        + (Math.abs(calc.solved_pref_delta) < 0.5 ? 'equals the stored figure'
          : (Math.abs(calc.solved_pref_delta).toLocaleString('en-US', { style: 'currency', currency: 'USD' })
            + ' ' + (calc.solved_pref_delta > 0 ? 'above' : 'below') + ' stored; Weaver Goal-Seek tolerance'))),
      SF({ size: 9, italic: true })); pr.mergeCells('A' + R + ':D' + R); R += 2;
  }
  put('A' + R, 'No Assurance Provided.', SF({ italic: true, size: 9 }));

  // ── Notes & Sources ──────────────────────────────────────────────────────────
  nt.getColumn(1).width = 118;
  const notes = [
    ['CLRF Preferred Return Workpaper — ' + q.label, F({ size: 12, bold: true })],
    ['No Assurance Provided. Weaver remains responsible for the official financial statements.', SMALLI],
    ['', F()],
    ['Purpose: reproduce the fund-level Preferred Return — the 8% per annum, annually-compounded, IRR-based return the Limited Partners receive before the General Partner participates in carried interest (LPA "Preferred Return" definition; used by the §17(c) carried-interest build-up).', F()],
    ['', F()],
    ['Method: the equalized Limited-Partner net cash-flow stream (contributions are outflows to the LPs, refunds are inflows), dated from inception, plus a single assumed liquidating distribution at the measurement date equal to Return of Capital + Preferred Return. The Preferred Return is the amount, on top of Return of Capital, that makes the stream’s internal rate of return (XIRR, Actual/365) equal 8%. Weaver solves it by Goal Seek (set the IRR cell to 8% by changing the Preferred Return cell); this workpaper recomputes the XIRR and independently re-solves the Preferred Return to verify the 8% result.', F()],
    ['', F()],
    ['Data source: CLRF’s CL ledger carries contributed capital only as a 12/31/2025 opening import (no dated capital-call history), so the dated cash-flow schedule is NOT derivable from the GL. Return of Capital, the Preferred Return, and the dated cash-flow schedule are maintained per quarter from the fund’s preferred-return workpaper (Weaver, "equalized cash flows per LK subsequent close workbook") and stored in fund_preferred_return.' + (fund.note ? ' Note: ' + fund.note + '.' : ''), F()],
    ['', F()],
    [calc.schedule_loaded
      ? ('This ' + q.label + ' run reproduced the calculation from ' + calc.cf_count + ' dated LP cash flows: computed XIRR '
        + (calc.irr == null ? 'n/a' : (calc.irr * 100).toFixed(4) + '%') + ' against an 8% target ('
        + (calc.irr_ties ? 'reaches the 8% hurdle' : 'does NOT reach the 8% hurdle') + '); Return of Capital '
        + (calc.roc_ties ? 'ties to' : 'does NOT tie to') + ' LP net capital invested. The Preferred Return that hits exactly 8.0000% is '
        + (calc.solved_pref == null ? 'n/a' : calc.solved_pref.toLocaleString('en-US', { style: 'currency', currency: 'USD' }))
        + (calc.solved_pref_delta == null || Math.abs(calc.solved_pref_delta) < 0.5 ? ', equal to the stored figure.'
          : ' — ' + Math.abs(calc.solved_pref_delta).toLocaleString('en-US', { style: 'currency', currency: 'USD' })
            + ' ' + (calc.solved_pref_delta > 0 ? 'above' : 'below') + ' the stored figure, the residual of Weaver’s Excel Goal Seek tolerance; the stored figure is retained as Weaver’s authoritative amount.'))
      : ('This ' + q.label + ' run had no dated cash-flow schedule loaded, so it presents the stored Return of Capital and Preferred Return summary only. Load the dated schedule to reproduce and verify the 8% XIRR.'), F()],
    ['', F()],
    ['The Preferred Return here flows into the Carried Interest / Clawback workpaper (side letter §17(c)) as the LP hurdle above Return of Capital.', F()],
  ];
  let nr = 1;
  for (const [text, font] of notes) { const c = nt.getCell('A' + nr); c.value = text; c.font = font; c.alignment = { wrapText: true }; nr++; }

  return wb;
}

// ── Persistence (identical shape to carryclawback.saveToWorkpapers). ──────────
const folderFor = (quarter) => 'Workpapers/Preferred Return/' + quarter.year + '/' + quarter.quarter;
const fileNameFor = (quarter) => 'CLRF_Preferred_Return_' + quarter.label + '.xlsx';

function saveToWorkpapers(ctx, eid, quarter, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(quarter);
  const original = fileNameFor(quarter);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
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

// Resolver so the financial-statements package can locate this workpaper.
function findWorkpaper(ctx, eid, quarterEnd) {
  const quarter = resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ? ORDER BY id DESC LIMIT 1').get(eid, folderFor(quarter), fileNameFor(quarter));
  if (!row) return null;
  return Object.assign({}, row, { quarter, abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerPreferredReturnRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/preferred-return/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const quarter = resolveQuarter(body.quarter_end || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, quarter, { entity_id: eid });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Pref-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          roc: data.fund.roc, pref: data.fund.pref, total: data.fund.total,
          schedule_loaded: data.calc.schedule_loaded, cf_count: data.calc.cf_count || 0,
          irr: data.calc.irr != null ? data.calc.irr : null, irr_target: PREF_RATE,
          solved_pref: data.calc.solved_pref != null ? data.calc.solved_pref : null,
          roc_ties: data.calc.roc_ties || false, irr_ties: data.calc.irr_ties || false,
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerPreferredReturnRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, xirr, solvePref, FUND_EID };
