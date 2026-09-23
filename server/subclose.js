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
const JSZip = require('jszip');

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

// ─── Workbook (fund-administrator template population) ───────────────────────
//
// The output reproduces the fund administrator's (Weaver) subclose workbook
// exactly — same two tabs, fonts, number formats, borders, section notes and
// column layout — by loading their file as a template and populating CloudLedger's
// GL-derived figures into it. Only the per-investor rebalance INPUT columns are
// overwritten (B commitment, C GCM eq., D Odyssey correcting, E Legacy Knight eq.,
// F James & Natalie eq., G May-26 call, K ending capital per WB); every derived
// column is rewired as a live formula so the sheet foots to CloudLedger's numbers.
//
// The subsequent-closing-specific rows (Odyssey true-up + the two subscriber rows
// Legacy Knight and James Bloomingdale) carry bespoke distribution formulas in the
// admin file and are left exactly as delivered — the same "pin to the administrator"
// treatment used elsewhere (preferred return @ 6/30/26). The "Sub Close Summary"
// tab is preserved verbatim.
const TEMPLATE_PATH = path.join(__dirname, 'assets', 'subclose_template.xlsx');
const CALC_SHEET = 'Sub Close Calcu';
const FIRST_ROW = 4;
const LAST_ROW = 85;
// Interest factor the admin file applies to the investment distribution (col Q).
const INT_FACTOR = 0.06847168947158275;
// Rows whose distribution columns are bespoke in the admin file (Odyssey true-up,
// Legacy Knight and James Bloomingdale subscriber rows) — left untouched.
const SPECIAL_ROWS = new Set([57, 82, 85]);

const normName = (s) => String(s == null ? '' : s)
  .toLowerCase()
  .replace(/&/g, ' and ')
  .replace(/[^a-z0-9]+/g, '');

// Admin-file investor names that differ from the CloudLedger class name (short
// form vs. full trust name). Map: normalized template name -> normalized CL name.
// (The admin's "James and Natelie Bloomingdale" row is intentionally NOT mapped —
// it is a subclose-entangled row the admin zeroes, with the funding shown on the
// James Bloomingdale subscriber row, so it stays pinned to the admin file.)
const NAME_ALIASES = {
  stewarttate: 'stewartseviertaterevocabletrustdatedoctober42013',
};

// Effective numeric value of a cell (formula result or literal).
function numOf(cell) {
  const v = cell && cell.value;
  if (v == null || v === '') return 0;
  if (typeof v === 'object') return ('result' in v) ? (Number(v.result) || 0) : 0;
  return Number(v) || 0;
}

// ExcelJS emits <pageSetUpPr> before <outlinePr> inside <sheetPr>, which Excel
// rejects ("Repaired Records ... Load error"). Restore schema order post-write.
async function fixSheetPrOrder(buf) {
  try {
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    let changed = false;
    for (const name of names) {
      const xml = await zip.file(name).async('string');
      const fixed = xml.replace(/(<pageSetUpPr\b[^>]*\/>)\s*(<outlinePr\b[^>]*\/>)/g, '$2$1');
      if (fixed !== xml) { zip.file(name, fixed); changed = true; }
    }
    if (!changed) return buf;
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) { return buf; }
}

// Weaver's source workbooks carry thousands of legacy named ranges (FactSet/model
// junk like "\a", "_______EPS91", plus many pointing at #REF!). ExcelJS drops
// defined names on save, so app output is normally clean, but this is a defensive
// guard: remove every defined name that is broken (#REF!) or references a sheet not
// in this workbook. Such names make Excel show "We found a problem with some content
// ... recover?" on open. Legitimate names (_xlnm.* print areas on existing sheets,
// or names actually used by a formula) are kept.
async function stripBadDefinedNames(buf) {
  try {
    const zip = await JSZip.loadAsync(buf);
    let wb = await zip.file('xl/workbook.xml').async('string');
    const m = wb.match(/<definedNames>([\s\S]*?)<\/definedNames>/);
    if (!m) return buf;
    const sheets = new Set([...wb.matchAll(/<sheet [^>]*name="([^"]*)"/g)].map((x) => x[1]));
    // formula text across worksheets, to keep names a formula actually uses
    let blob = '';
    for (const n of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      blob += (await zip.file(n).async('string')).replace(/[\s\S]*?(<f\b)/g, '$1');
    }
    const keep = [];
    for (const dn of m[1].match(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g) || []) {
      const nm = (dn.match(/name="([^"]*)"/) || [])[1] || '';
      const val = dn.replace(/<[^>]+>/g, '');
      if (dn.includes('#REF!')) continue;
      const refs = [...val.matchAll(/(?:^|[=,+\-*/(! ])'?([A-Za-z0-9 _.&\-]+?)'?!/g)].map((x) => x[1]);
      if (refs.some((r) => !sheets.has(r))) continue; // orphaned sheet reference
      const used = new RegExp('(?<![A-Za-z0-9_.])' + nm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![A-Za-z0-9_.])').test(blob);
      if (nm.startsWith('_xlnm.') || used) keep.push(dn);
    }
    if (keep.length === (m[1].match(/<definedName\b/g) || []).length) return buf; // nothing removed
    const block = keep.length ? '<definedNames>' + keep.join('') + '</definedNames>' : '';
    wb = wb.slice(0, m.index) + block + wb.slice(m.index + m[0].length);
    zip.file('xl/workbook.xml', wb);
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) { return buf; }
}

// The Weaver template's "Sub Close Summary" sheet is built from formulas that
// reference an EXTERNAL workbook ('[1]...'!Ref — Weaver's own source file). ExcelJS
// drops the external-link parts on save (xl/externalLinks + the workbook's
// <externalReference>), which leaves those formulas pointing at a workbook that no
// longer exists, so Excel reports "Removed Records: Formula" and strips them on open.
// Those cells only ever displayed cached values anyway (the Weaver file is never
// shipped with this workpaper and can't recompute on anyone's machine), so freeze
// every externally-linked formula to its cached value: drop the <f> element, keep
// the <v>. Only sheets that actually carry an external '[1]' reference are touched,
// so the Calculation sheet's live internal formulas are left intact.
async function freezeExternalLinkFormulas(buf) {
  try {
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    let changed = false;
    for (const name of names) {
      const xml = await zip.file(name).async('string');
      if (!xml.includes('[1]')) continue; // sheet has no external-workbook links
      // Strip the formula (shared follower <f .../> or full <f>…</f>) from any cell
      // that has a cached <v>, keeping the value and the cell's attributes/style.
      const fixed = xml.replace(
        /<c\b([^>]*)>(?:<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>)(<v>[\s\S]*?<\/v>)<\/c>/g,
        '<c$1>$2</c>',
      );
      if (fixed !== xml) { zip.file(name, fixed); changed = true; }
    }
    if (!changed) return buf;
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) { return buf; }
}

async function buildWorkbook(data) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_PATH);
  const ws = wb.getWorksheet(CALC_SHEET);
  if (!ws) throw new Error('subclose template missing "' + CALC_SHEET + '" sheet');

  // Index CloudLedger investors by normalized name.
  const byName = new Map();
  for (const inv of data.investors) {
    const key = normName(inv.name);
    if (key && !byName.has(key)) byName.set(key, inv);
  }

  const setVal = (addr, v) => { ws.getCell(addr).value = v; };
  const setFormula = (addr, formula, result) => { ws.getCell(addr).value = { formula: formula, result: result }; };

  let matched = 0;
  for (let r = FIRST_ROW; r <= LAST_ROW; r++) {
    if (SPECIAL_ROWS.has(r)) continue;
    const nameCell = ws.getCell('A' + r).value;
    const key = normName(typeof nameCell === 'object' && nameCell ? (nameCell.result || nameCell.text) : nameCell);
    const inv = key && (byName.get(key) || (NAME_ALIASES[key] && byName.get(NAME_ALIASES[key])));
    if (!inv) continue; // leave the administrator's values for any unmatched row
    matched++;

    const b = r2(inv.commitment);
    const c = r2(inv.gcmEqualization);
    const d = r2(inv.odysseyCorrecting);
    const e = r2(inv.lkEqualization);
    const f = r2(inv.jnbEqualization);
    const g = r2(inv.capitalCall);
    const k = r2(-inv.endingCapital);

    // Input columns (overwrite; existing number formats are preserved).
    setVal('B' + r, b); setVal('C' + r, c); setVal('D' + r, d);
    setVal('E' + r, e); setVal('F' + r, f); setVal('G' + r, g); setVal('K' + r, k);

    // Derived columns — live formulas with cached results (S/T blank on normal rows).
    const S = numOf(ws.getCell('S' + r));
    const T = numOf(ws.getCell('T' + r));
    const Y = numOf(ws.getCell('Y' + r));
    const H = b + c + d + e + f + g;
    const I = b - H;
    const L = H + k - g;
    const O = d;
    const P = e;
    const Q = e * INT_FACTOR;
    const R = O + P + Q;
    const U = R + S + T;
    const X = g;
    const Z = X + Y;
    const AB = Z + U;

    setFormula('H' + r, '+B' + r + '+C' + r + '+D' + r + '+E' + r + '+F' + r + '+G' + r, H);
    setFormula('I' + r, '+B' + r + '-H' + r, I);
    setFormula('L' + r, '+H' + r + '+K' + r + '-G' + r, L);
    setFormula('O' + r, '+D' + r, O);
    setFormula('P' + r, '+E' + r, P);
    setFormula('Q' + r, '+P' + r + '*' + INT_FACTOR, Q);
    setFormula('R' + r, '+O' + r + '+P' + r + '+Q' + r, R);
    setFormula('U' + r, '+R' + r + '+S' + r + '+T' + r, U);
    setFormula('V' + r, '+P' + r, P);
    setFormula('X' + r, '+G' + r, X);
    setFormula('Z' + r, '+X' + r + '+Y' + r, Z);
    setFormula('AB' + r, '+Z' + r + '+U' + r, AB);
  }

  // Refoot the header/footer totals over the (now CloudLedger-populated) rows,
  // so the workbook foots regardless of which rows were matched.
  const refoot = (addr, col) => {
    const cell = ws.getCell(addr);
    if (cell.value == null || cell.value === '') return;
    let s = 0;
    for (let r = FIRST_ROW; r <= LAST_ROW; r++) s += numOf(ws.getCell(col + r));
    cell.value = { formula: 'SUM(' + col + FIRST_ROW + ':' + col + LAST_ROW + ')', result: r2(s) };
  };
  ['B', 'C', 'D', 'E', 'H', 'K'].forEach((col) => refoot(col + '1', col));
  ['N', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'X', 'Y', 'Z', 'AB'].forEach((col) => refoot(col + '86', col));

  wb._clMatched = matched; // for logging
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
        const wb = await buildWorkbook(data);
        const buf = await stripBadDefinedNames(await freezeExternalLinkFormulas(await fixSheetPrOrder(Buffer.from(await wb.xlsx.writeBuffer()))));
        const saved = saveToWorkpapers(ctx, eid, asOf, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-SubClose-Summary', JSON.stringify({
          as_of: asOf, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          investors: data.totals.count, ending_capital: data.totals.endingCapital,
          equalization_cash: data.totals.equalizationCash,
          template_rows_matched: wb._clMatched,
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
      const wb = await buildWorkbook(data);
      const buf = await stripBadDefinedNames(await freezeExternalLinkFormulas(await fixSheetPrOrder(Buffer.from(await wb.xlsx.writeBuffer()))));
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fileNameFor(asOf) + '"');
      res.send(buf);
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
}

module.exports = { buildData, buildWorkbook, saveToWorkpapers, registerSubcloseRoutes, FUND_EID };
