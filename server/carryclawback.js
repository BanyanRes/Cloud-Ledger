// ─── CLRF workpaper: Carried Interest / Clawback (Side Letter §17(c)) ──────────
//
// A quarterly workpaper for County Line Rail Fund I, LP (CL entity 40) that
// reproduces the GCM Grosvenor side-letter §17(c) disclosures:
//   (i)   the "build-up" to carried interest,
//   (ii)  carried interest earned to date (cumulative), and
//   (iii) the clawback the GP would owe if the Partnership were dissolved and
//         liquidated today.
//
// It is prepared exactly like the GP Fees & Expenses workpaper (server/gpfees.js):
// buildData -> buildWorkbook -> saveToWorkpapers, registered by index.js, and
// located by the financial-statements package through findWorkpaper() so it folds
// into the quarterly deliverable.
//
// ── Economics (LPA §9.3 + "Preferred Return" definition) ──────────────────────
// Per-Limited-Partner (American) waterfall, applied to each LP independently:
//   (a) 100% to the LP until cumulative distributions equal aggregate Capital
//       Contributions (Return of Capital);
//   (b) 100% to the LP until it has received an 8% annually-compounded, IRR-based
//       Preferred Return;
//   (c) 80% to the Promoting Partner / 20% to the LP (catch-up) until the
//       Promoting Partner has received 20% of [Preferred Return + amounts under
//       this clause (c)];
//   (d) 80% to the LP / 20% to the Promoting Partner (residual).
// Carried Interest = distributions to the Promoting Partner under (c) and (d).
//
// ── Preferred Return input ────────────────────────────────────────────────────
// The 8% Preferred Return is IRR-based and depends on each LP's true capital-call
// dates from fund inception. CLRF's CL ledger holds the contributed-capital
// balance only as a 12/31/2025 opening import (no dated call history), so the
// Preferred Return is NOT computed from the GL. It is supplied from the fund's
// authoritative Preferred Return workpaper (Weaver), injected here as
// opts.prefByClass (class_id -> accrued unpaid preferred return) and/or
// opts.prefTotal (fund-level aggregate). When neither is supplied the workpaper
// prints every Preferred-Return-dependent figure as "PENDING" so the deliverable
// shows the gap on its face rather than understating carry.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// County Line Rail Fund I, LP.
const FUND_EID = 40;

// Equity account groups on the CLRF ledger (validated against the 3/31/2026
// package). Capital = every equity account; Contributed = the contribution
// accounts net of capital-call refunds (refunds post as debits to these same
// accounts and are, per Jimmy, the same movement as "Return of Capital").
const CONTRIB_ACCTS = ['30100', '301100', '301200', '301300', '301800'];
const SYND_ACCTS = ['370200'];
const ACCUM_ACCTS = ['390100', '390110', '390120', '390130', '390500'];
const EQUITY_ACCTS = [...CONTRIB_ACCTS, ...SYND_ACCTS, ...ACCUM_ACCTS];

const PREF_RATE = 0.08;      // 8% per annum, annually compounded (LPA definition)
const CATCHUP_GP = 0.80, CATCHUP_LP = 0.20;   // clause (c)
const RESIDUAL_LP = 0.80, RESIDUAL_GP = 0.20; // clause (d)

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));

// ── Quarter arithmetic (same contract as gpfees.resolveQuarter). ──────────────
function resolveQuarter(quarterEnd) {
  if (!isDate(quarterEnd)) throw new Error('quarter_end must be a date in YYYY-MM-DD form');
  const [y, m, d] = quarterEnd.split('-').map(Number);
  const ENDS = { 3: 31, 6: 30, 9: 30, 12: 31 };
  if (!ENDS[m] || d !== ENDS[m]) {
    throw new Error('quarter_end must be a quarter end date: 03-31, 06-30, 09-30 or 12-31. Received ' + quarterEnd);
  }
  const q = m / 3;
  const priorEnd = new Date(Date.UTC(y, m - 3, 1));
  priorEnd.setUTCDate(0);
  return {
    label: y + '-Q' + q, year: String(y), quarter: 'Q' + q,
    start: y + '-' + String(m - 2).padStart(2, '0') + '-01',
    end: quarterEnd, prior_end: priorEnd.toISOString().slice(0, 10),
  };
}

// ── Per-LP waterfall (LPA §9.3(a)-(d)). Pure function; distributable/roc/pref are
// dollars. Returns the split of a hypothetical full distribution of the LP's
// distributable assets as of the measurement date. In a shortfall (distributable
// < roc + pref) every carry tier is zero — which is CLRF's Q1 2026 case.
function waterfallLP(distributable, roc, pref) {
  let avail = r2(distributable);
  const rocReturn = Math.min(avail, r2(roc)); avail = r2(avail - rocReturn);
  const prefReturn = Math.min(avail, r2(pref)); avail = r2(avail - prefReturn);
  // Catch-up: split CATCHUP (80 GP / 20 LP). The Promoting Partner target is 20%
  // of (pref + total catch-up pool); solving 0.8T = 0.2(pref + T) gives T = pref/3.
  const catchTarget = r2(pref / 3);
  const catchPool = Math.min(avail, catchTarget); avail = r2(avail - catchPool);
  const catchupGP = r2(CATCHUP_GP * catchPool);
  const catchupLP = r2(CATCHUP_LP * catchPool);
  // Residual (80 LP / 20 GP).
  const residualLP = r2(RESIDUAL_LP * avail);
  const residualGP = r2(RESIDUAL_GP * avail);
  const carryGP = r2(catchupGP + residualGP);
  return {
    distributable: r2(distributable), roc: r2(roc), pref: r2(pref),
    excess: r2(distributable - roc - pref),
    rocReturn, prefReturn, catchupGP, catchupLP, residualLP, residualGP, carryGP,
  };
}

// ── Gather everything the workbook needs. `opts` may carry:
//     prefByClass    { [class_id]: accruedUnpaidPreferredReturn }  (Weaver)
//     prefTotal      fund-level accrued unpaid preferred return    (Weaver)
//     accumCarryByClass / accumCarryTotal  carry distributed to PP to date
function buildData(ctx, quarter, opts = {}) {
  const { db, computeBalances } = ctx;
  const eid = opts.entity_id || FUND_EID;
  const ent = db.prepare('SELECT id, name FROM entities WHERE id = ?').get(eid);
  const classes = db.prepare('SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?').all(eid);
  const commitRows = db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
  const commitBy = {}; commitRows.forEach((c) => { commitBy[c.class_id] = r2(c.commitment_amount); });

  const prefByClass = opts.prefByClass || null;
  const accumCarryByClass = opts.accumCarryByClass || null;
  const sumAcct = (rows, codes) => r2(rows
    .filter((b) => codes.includes(String(b.code)))
    .reduce((s, b) => s + (Number(b.balance) || 0), 0));

  const partners = [];
  for (const c of classes) {
    const rows = computeBalances(eid, { as_of: quarter.end, class_id: c.id });
    const capital = sumAcct(rows, EQUITY_ACCTS);        // partners' capital = distributable
    const contributed = sumAcct(rows, CONTRIB_ACCTS);   // unreturned capital = return of capital
    // Skip empty classes (no capital, no commitment) so a stray tag doesn't print.
    if (Math.abs(capital) < 0.005 && Math.abs(contributed) < 0.005 && !commitBy[c.id]) continue;
    const isGP = String(c.partner_type || '').toUpperCase() === 'GP';
    const pref = prefByClass && (c.id in prefByClass) ? r2(prefByClass[c.id]) : null;
    const p = {
      class_id: c.id, name: c.name, partner_type: isGP ? 'GP' : 'LP',
      commitment: commitBy[c.id] || 0,
      distributable: r2(capital), roc: r2(contributed), pref,
      syndication: sumAcct(rows, SYND_ACCTS), accumulated: sumAcct(rows, ACCUM_ACCTS),
      accum_carry: accumCarryByClass && (c.id in accumCarryByClass) ? r2(accumCarryByClass[c.id]) : 0,
    };
    if (!isGP && pref !== null) Object.assign(p, waterfallLP(p.distributable, p.roc, pref));
    partners.push(p);
  }
  partners.sort((a, b) => (a.partner_type === b.partner_type ? 0 : a.partner_type === 'LP' ? -1 : 1)
    || String(a.name).localeCompare(String(b.name)));

  const lps = partners.filter((p) => p.partner_type === 'LP');
  const sum = (arr, k) => r2(arr.reduce((s, x) => s + (Number(x[k]) || 0), 0));
  const prefKnown = prefByClass ? lps.every((p) => p.pref !== null) : (opts.prefTotal != null);
  const prefTotal = prefByClass ? sum(lps, 'pref') : (opts.prefTotal != null ? r2(opts.prefTotal) : null);
  const distributable = sum(lps, 'distributable');
  const roc = sum(lps, 'roc');

  const fund = {
    entity_id: eid, entity_name: ent ? ent.name : ('entity ' + eid),
    lp_count: lps.length, gp_count: partners.length - lps.length,
    distributable, roc,
    pref: prefKnown ? prefTotal : null,
    excess: prefKnown ? r2(distributable - roc - prefTotal) : null,
    // Carry build-up: sum of per-LP carry (0 in a shortfall). Only meaningful when
    // pref is per-class; with a fund-level pref override the tiers stay 0 unless
    // the aggregate excess is positive.
    catchupGP: prefByClass && prefKnown ? sum(lps, 'catchupGP') : 0,
    catchupLP: prefByClass && prefKnown ? sum(lps, 'catchupLP') : 0,
    residualLP: prefByClass && prefKnown ? sum(lps, 'residualLP') : 0,
    residualGP: prefByClass && prefKnown ? sum(lps, 'residualGP') : 0,
    carry_quarter: prefByClass && prefKnown ? sum(lps, 'carryGP') : (prefKnown && prefTotal != null && (distributable - roc - prefTotal) > 0 ? null : 0),
    accum_carry: accumCarryByClass ? sum(lps, 'accum_carry') : r2(opts.accumCarryTotal || 0),
    pref_known: prefKnown,
    pref_source: prefByClass ? 'per-LP (Weaver preferred-return workpaper)'
      : (opts.prefTotal != null ? 'fund-level override' : null),
  };
  // §17(c)(iii) clawback if liquidated & dissolved today = carry received to date
  // in excess of carry earned on a hypothetical liquidation. Zero while in
  // shortfall and no carry has been distributed.
  fund.carry_to_date = fund.accum_carry;
  fund.clawback = prefKnown ? Math.max(0, r2(fund.accum_carry - (fund.carry_quarter || 0))) : null;

  return { quarter, fund, partners, lps };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const MONEY = '$#,##0.00;($#,##0.00);-';
const NAVY = 'FF1F3864';
const HDR_FONT = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const F = (o = {}) => Object.assign({ name: 'Arial', size: 10 }, o);
const BLUE = F({ color: { argb: 'FF0000FF' } });     // value read from the source system
const SMALLI = F({ size: 9, italic: true });
const SUM_MONEY = '_($* #,##0.00_);_($* (#,##0.00);_($* -??_);_(@_)';
const SF = (o = {}) => Object.assign({ name: 'Times New Roman', size: 10 }, o);
const THIN = { style: 'thin' };
const MONTHS = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
  'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
const spellQuarterEnd = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };
const PENDING = 'PENDING — per Weaver preferred-return workpaper';

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
  const q = data.quarter, fund = data.fund;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();

  const ci = wb.addWorksheet('Carried Interest');
  const wf = wb.addWorksheet('Waterfall Detail', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
  const gl = wb.addWorksheet('GL Data', { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
  const nt = wb.addWorksheet('Notes & Sources');

  // ── Carried Interest (client-facing, mirrors the §17(c) FS schedule) ─────────
  ci.getColumn(1).width = 6; ci.getColumn(2).width = 62; ci.getColumn(3).width = 22;
  const title = (rowNum, text, opts = {}) => {
    const c = ci.getCell('B' + rowNum); c.value = text; c.font = SF(Object.assign({ bold: true }, opts));
    c.alignment = { horizontal: 'center' };
    ci.mergeCells('B' + rowNum + ':C' + rowNum);
  };
  title(1, fund.entity_name); title(2, 'Carried Interest Reporting per Section §17(c) of GCM Grosvenor Side Letter Agreement');
  title(3, 'For the Quarter Ended ' + spellQuarterEnd(q.end));
  ci.getCell('B3').font = SF({ italic: true });
  const money = (ref, v, o = {}) => {
    const c = ci.getCell(ref);
    if (v === null || v === undefined) { c.value = PENDING; c.font = SF({ italic: true, color: { argb: 'FFC00000' } }); }
    else { c.value = v; c.numFmt = SUM_MONEY; c.font = SF(o); }
  };
  const label = (ref, v, o = {}) => { const c = ci.getCell(ref); c.value = v; c.font = SF(o); };
  let R = 5;
  const secHdr = (t) => { ci.getCell('B' + R).value = t; ci.getCell('B' + R).font = SF({ bold: true }); R++; };
  secHdr('Carried interest build-up');
  label('B' + R, '   Distributable assets as of quarter-end', { indent: 1 }); money('C' + R, fund.distributable); R++;
  label('B' + R, '      Return of Capital'); money('C' + R, fund.roc === null ? null : -fund.roc); R++;
  label('B' + R, '      Preferred Return'); money('C' + R, fund.pref === null ? null : -fund.pref); R++;
  label('B' + R, 'Excess/(Shortfall)', { bold: true }); money('C' + R, fund.excess, { bold: true });
  ['B', 'C'].forEach((col) => { ci.getCell(col + R).border = { top: THIN, bottom: THIN }; }); R += 2;

  label('B' + R, 'Current distributable assets subject to carried interest', { bold: true });
  money('C' + R, fund.pref_known ? Math.max(0, fund.excess || 0) : null); R += 2;

  label('B' + R, '§17(c)(i)', { italic: true, bold: true }); R++;
  label('B' + R, '   Catch-Up to GP (80%)'); money('C' + R, fund.pref_known ? fund.catchupGP : null); R++;
  label('B' + R, '   Catch-Up to LP (20%)'); money('C' + R, fund.pref_known ? fund.catchupLP : null); R++;
  label('B' + R, '   Residual Split – LP (80%)'); money('C' + R, fund.pref_known ? fund.residualLP : null); R++;
  label('B' + R, '   Residual Split – GP (20%)'); money('C' + R, fund.pref_known ? fund.residualGP : null); R += 2;
  label('B' + R, '   Accumulated carried interest as of beginning of quarter'); money('C' + R, fund.accum_carry); R++;
  label('B' + R, '   Accumulated carried interest as of end of quarter'); money('C' + R, fund.pref_known ? r2(fund.accum_carry + (fund.carry_quarter || 0)) : null); R++;
  label('B' + R, '   Total carried interest in quarter'); money('C' + R, fund.carry_quarter); R += 2;

  label('B' + R, '§17(c)(ii)', { italic: true, bold: true });
  label('B' + (R + 1), '   Carried interest earned to date (cumulative)'); money('C' + (R + 1), fund.carry_to_date); R += 3;
  label('B' + R, '§17(c)(iii)', { italic: true, bold: true });
  label('B' + (R + 1), '   Clawback that the GP would have to pay if the Partnership were dissolved & liquidated today');
  ci.getCell('B' + (R + 1)).alignment = { wrapText: true };
  money('C' + (R + 1), fund.clawback); R += 3;
  label('B' + R, 'No Assurance Provided.', { italic: true, size: 9 });

  // ── Waterfall Detail (per-LP) ────────────────────────────────────────────────
  wf.getCell('A1').value = 'PER-LIMITED-PARTNER WATERFALL — ' + fund.entity_name;
  wf.getCell('A1').font = F({ size: 12, bold: true });
  wf.getCell('A2').value = 'Hypothetical full distribution of each LP’s capital as of ' + q.end
    + ' under LPA §9.3: Return of Capital → 8% Preferred Return → 80/20 catch-up → 80/20 residual.';
  wf.getCell('A2').font = SMALLI;
  wf.getCell('A3').value = 'Preferred Return is sourced from the fund’s preferred-return workpaper (Weaver); see Notes.';
  wf.getCell('A3').font = SMALLI;
  const wfCols = ['Partner', 'Type', 'Commitment', 'Distributable (capital)', 'Return of Capital', 'Preferred Return',
    'Excess/(Shortfall)', 'Catch-Up GP', 'Catch-Up LP', 'Residual LP', 'Residual GP', 'Carried Interest (GP)'];
  headerRow(wf, 5, wfCols, [34, 7, 16, 18, 17, 16, 16, 14, 14, 14, 14, 18]);
  let wr = 6;
  const wcell = (col, v, fmt) => { const c = wf.getCell(col + wr); if (v === null || v === undefined) { c.value = '—'; c.font = F({ italic: true }); } else { c.value = v; if (fmt) c.numFmt = fmt; c.font = F(); } };
  for (const p of data.lps) {
    wf.getCell('A' + wr).value = p.name; wf.getCell('A' + wr).font = F();
    wf.getCell('B' + wr).value = p.partner_type; wf.getCell('B' + wr).font = F();
    wcell('C', p.commitment, MONEY); wcell('D', p.distributable, MONEY); wcell('E', p.roc, MONEY);
    wcell('F', p.pref, MONEY); wcell('G', p.pref === null ? null : r2(p.distributable - p.roc - p.pref), MONEY);
    wcell('H', p.catchupGP, MONEY); wcell('I', p.catchupLP, MONEY);
    wcell('J', p.residualLP, MONEY); wcell('K', p.residualGP, MONEY); wcell('L', p.carryGP, MONEY);
    wr++;
  }
  // Total row
  const tot = wf.getRow(wr); tot.getCell(1).value = 'Total — Limited Partners'; tot.getCell(1).font = F({ bold: true });
  const sums = { C: 'commitment', D: 'distributable', E: 'roc', H: 'catchupGP', I: 'catchupLP', J: 'residualLP', K: 'residualGP', L: 'carryGP' };
  const s = (k) => r2(data.lps.reduce((a, x) => a + (Number(x[k]) || 0), 0));
  Object.entries(sums).forEach(([col, k]) => { const c = tot.getCell(col); c.value = s(k); c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; });
  { const c = tot.getCell('F'); c.value = fund.pref === null ? '—' : fund.pref; if (fund.pref !== null) c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }
  { const c = tot.getCell('G'); c.value = fund.excess === null ? '—' : fund.excess; if (fund.excess !== null) c.numFmt = MONEY; c.font = F({ bold: true }); c.border = { top: THIN }; }

  // ── GL Data (per-class equity balances the figures come from) ────────────────
  gl.getCell('A1').value = 'GL DATA — PARTNERS’ CAPITAL BY INVESTOR CLASS AS OF ' + q.end;
  gl.getCell('A1').font = F({ size: 12, bold: true });
  gl.getCell('A2').value = 'Distributable = all equity accounts (' + EQUITY_ACCTS.join(', ') + '). '
    + 'Return of Capital = contribution accounts net of refunds (' + CONTRIB_ACCTS.join(', ') + ').';
  gl.getCell('A2').font = SMALLI;
  headerRow(gl, 4, ['Investor class', 'Type', 'Distributable (capital)', 'Return of Capital', 'Syndication', 'Accumulated inc/exp/mgmt/unreal'], [40, 8, 20, 18, 14, 26]);
  let gr = 5;
  for (const p of data.partners) {
    const row = gl.getRow(gr);
    row.getCell(1).value = p.name; row.getCell(1).font = F();
    row.getCell(2).value = p.partner_type; row.getCell(2).font = F();
    [['C', p.distributable], ['D', p.roc], ['E', p.syndication], ['F', p.accumulated]].forEach(([col, v]) => {
      const c = row.getCell({ C: 3, D: 4, E: 5, F: 6 }[col]); c.value = v; c.numFmt = MONEY; c.font = BLUE;
    });
    gr++;
  }

  // ── Notes & Sources ──────────────────────────────────────────────────────────
  nt.getColumn(1).width = 118;
  const notes = [
    ['CLRF Carried Interest / Clawback Workpaper — ' + q.label, F({ size: 12, bold: true })],
    ['No Assurance Provided. Weaver remains responsible for the official financial statements.', SMALLI],
    ['', F()],
    ['Purpose: GCM Grosvenor side-letter §17(c) disclosures — (i) build-up to carried interest, (ii) carried interest earned to date, (iii) clawback if the Partnership were dissolved and liquidated today.', F()],
    ['', F()],
    ['Waterfall (LPA §9.3), applied per Limited Partner: (a) 100% Return of Capital; (b) 100% to an 8% annually-compounded, IRR-based Preferred Return; (c) 80% Promoting Partner / 20% LP catch-up until the Promoting Partner holds 20% of [Preferred Return + clause (c)]; (d) 80% LP / 20% Promoting Partner residual. Carried Interest = distributions to the Promoting Partner under (c) and (d).', F()],
    ['', F()],
    ['Preferred Return: 8% is IRR-based on each LP’s true capital-call dates. CLRF’s CL ledger carries contributed capital only as a 12/31/2025 opening import, so the Preferred Return is NOT computed from the GL — it is taken from the fund’s preferred-return workpaper (Weaver) and injected as prefByClass/prefTotal. Source used this run: ' + (fund.pref_source || 'NONE — Preferred Return figures pending'), F()],
    ['', F()],
    ['Distributable assets = Limited Partners’ capital = every equity account (' + EQUITY_ACCTS.join(', ') + ') summed by investor class as of ' + q.end + '. Return of Capital = unreturned contributions = contribution accounts (' + CONTRIB_ACCTS.join(', ') + ') net of capital-call refunds. GP classes (partner_type = GP) are the Promoting Partner side and are excluded from the LP build-up.', F()],
    ['', F()],
    ['Tie-out to the ' + q.label + ' fund statements: Distributable assets should equal Limited Partners’ capital on the Statement of Assets, Liabilities and Partners’ Capital. This run: $' + fund.distributable.toLocaleString('en-US', { minimumFractionDigits: 2 }) + ' across ' + fund.lp_count + ' LP classes.', F()],
    ['', F()],
    ['Q1 2026 is a shortfall (distributable < Return of Capital + Preferred Return), so every carry tier, carried interest to date, and clawback is $0.', F()],
  ];
  let nr = 1;
  for (const [text, font] of notes) { const c = nt.getCell('A' + nr); c.value = text; c.font = font; c.alignment = { wrapText: true }; nr++; }

  return wb;
}

// ── Persistence (identical shape to gpfees.saveToWorkpapers). ─────────────────
const folderFor = (quarter) => 'Workpapers/Carried Interest & Clawback/' + quarter.year + '/' + quarter.quarter;
const fileNameFor = (quarter) => 'CLRF_Carried_Interest_' + quarter.label + '.xlsx';

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

// Resolver so the financial-statements package can locate this workpaper for a
// period without hardcoding the path (mirrors gpfees.findWorkpaper).
function findWorkpaper(ctx, eid, quarterEnd) {
  const quarter = resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ? ORDER BY id DESC LIMIT 1').get(eid, folderFor(quarter), fileNameFor(quarter));
  if (!row) return null;
  return Object.assign({}, row, { quarter, abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerCarryClawbackRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/carry-clawback/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const quarter = resolveQuarter(body.quarter_end || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, quarter, {
          entity_id: eid,
          prefByClass: body.pref_by_class || null,
          prefTotal: body.pref_total != null ? body.pref_total : null,
          accumCarryByClass: body.accum_carry_by_class || null,
          accumCarryTotal: body.accum_carry_total != null ? body.accum_carry_total : 0,
        });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-Carry-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          distributable: data.fund.distributable, roc: data.fund.roc, pref: data.fund.pref,
          excess: data.fund.excess, carry_quarter: data.fund.carry_quarter, clawback: data.fund.clawback,
          pref_known: data.fund.pref_known, lp_count: data.fund.lp_count,
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { registerCarryClawbackRoutes, findWorkpaper, resolveQuarter, buildData, buildWorkbook, waterfallLP, FUND_EID };
