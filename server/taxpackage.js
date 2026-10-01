// ─── Tax Packages (TB & GL) ───────────────────────────────────────────────────
//
// One workbook per entity reproducing the year-end "TB & GL Tax Package" that
// goes to the tax preparer (EJ CPAs). Every tab is derived from CloudLedger's
// general ledger through the same computeBalances the balance/TB reports use, so
// the package ties to the books exactly:
//
//   • TB   — the trial balance as of the period end: Code | Account | Type |
//            Debit | Credit | Closing Balance, each closing a =D-E formula, a
//            footed Total, and a Net Income / Net Loss line (yellow) summing the
//            P&L rows. Balance-sheet accounts carry a "See <code>" note pointing
//            at their supporting tab.
//   • GL   — the full general ledger for the period, grouped by account: a
//            "<code> - <name>" header then Date | JE | Memo | Debit | Credit |
//            Balance, with a running natural balance. Balance-sheet accounts get
//            a synthesized Opening Balance line when the opening balance predates
//            the window; P&L accounts show period activity only.
//   • <code> supporting tabs — the GL block for each balance-sheet account on its
//            own tab (the "See <code>" target).
//   • Org Chart — a placeholder tab (the org-chart diagram is pasted in by hand).
//
// Two modes, same engine:
//   • Annual  — the tax year: TB as of 12/31, GL window prior-12/31 .. 12/31 (the
//               opening-balance JE at the prior year end shows in-window, exactly
//               as the hand-built 2025 packages do).
//   • Monthly — a single month: TB as of month end, GL window month-01 .. month
//               end, balance-sheet accounts opening at the prior month end.
//
// Display matches the delivered packages: Verdana 8, the accounting number
// format, a light-blue TB header, and the period/date headings. A copy is filed
// under Workpapers > Tax Packages > Annual|Monthly > <year>.
//
// ctx = { db, auth, requireEntityAccess, requireRole, workpapersDir, computeBalances }.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const multer = require('multer');

// ── Display constants (match the delivered TB & GL packages) ──────────────────
const FONT = 'Verdana';
const SIZE = 8;
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const HDRFILL = 'FFCEDEEF';
const YELLOW = 'FFFFFF00';
const F = (o = {}) => Object.assign({ name: FONT, size: SIZE }, o);
const fillOf = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const pad2 = (n) => String(n).padStart(2, '0');
const pad4 = (n) => String(n == null ? '' : n).padStart(4, '0');
const isBS = (t) => t === 'Asset' || t === 'Liability' || t === 'Equity';
const isDrNat = (t) => t === 'Asset' || t === 'Expense';
const numCode = (c) => { const n = Number(String(c).replace(/[^\d.-]/g, '')); return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY; };
const sortByCode = (a, b) => { const d = numCode(a) - numCode(b); return d !== 0 ? d : String(a).localeCompare(String(b)); };
const dayBefore = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return new Date(Date.UTC(y, mo - 1, dd - 1)).toISOString().slice(0, 10); };
const mdy = (d) => { const [y, mo, dd] = String(d).split('-').map(Number); return mo + '/' + dd + '/' + y; };

// Excel-safe sheet name (<=31 chars, unique, no illegal chars).
function sheetNameFor(base, used) {
  let name = String(base).replace(/[\[\]\:\*\?\/\\]/g, '_').slice(0, 31) || 'acct';
  let root = name, i = 1;
  while (used.has(name.toLowerCase())) { const suf = '_' + (++i); name = root.slice(0, 31 - suf.length) + suf; }
  used.add(name.toLowerCase());
  return name;
}

// Resolve the reporting period for either mode.
function resolvePeriod(dateInput, mode) {
  const m = String(dateInput || '').match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) throw new Error('period_end must be a date in YYYY-MM-DD form');
  const y = Number(m[1]), mo = Number(m[2]);
  if (mo < 1 || mo > 12) throw new Error('period_end must be a valid date');
  if (mode === 'annual') {
    const tbAsOf = y + '-12-31';
    const windowFrom = (y - 1) + '-12-31';
    return {
      mode, year: String(y), label: String(y),
      tbAsOf, periodTo: tbAsOf, windowFrom, openingAsOf: dayBefore(windowFrom),
      periodLine: 'Period: ' + windowFrom + ' to ' + tbAsOf,
      clsDate: mdy(tbAsOf), title: 'Annual',
    };
  }
  // monthly
  const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  const end = y + '-' + pad2(mo) + '-' + pad2(last);
  const windowFrom = y + '-' + pad2(mo) + '-01';
  return {
    mode: 'monthly', year: String(y), label: y + '-' + pad2(mo),
    tbAsOf: end, periodTo: end, windowFrom, openingAsOf: dayBefore(windowFrom),
    periodLine: 'Period: ' + windowFrom + ' to ' + end,
    clsDate: mdy(end), title: 'Monthly',
  };
}

// ── Data layer: trial balance, opening balances, and GL activity ──────────────
function buildData(ctx, eid, p) {
  const { db, computeBalances } = ctx;
  const ent = db.prepare('SELECT id, name, code, display_id FROM entities WHERE id = ?').get(eid);
  if (!ent) throw new Error('Entity not found');

  // Trial balance as of the period end. closing = total_debit - total_credit
  // (debit-positive; liabilities/equity/revenue therefore negative), mirroring
  // the delivered packages exactly.
  const balRows = computeBalances(eid, { as_of: p.tbAsOf }) || [];
  const meta = new Map(); // code -> { name, type }
  const tb = [];
  for (const r of balRows) {
    const code = String(r.code);
    meta.set(code, { name: r.name || '', type: r.type || '' });
    const closing = r2((Number(r.total_debit) || 0) - (Number(r.total_credit) || 0));
    if (Math.abs(closing) >= 0.005) tb.push({ code, name: r.name || '', type: r.type || '', closing });
  }
  tb.sort((a, b) => sortByCode(a.code, b.code));

  // Opening (natural) balance as of the day before the window, by account. Only
  // used to seed balance-sheet GL blocks whose opening predates the window.
  const openRows = computeBalances(eid, { as_of: p.openingAsOf }) || [];
  const opening = new Map();
  for (const r of openRows) { opening.set(String(r.code), r2(Number(r.balance) || 0)); if (!meta.has(String(r.code))) meta.set(String(r.code), { name: r.name || '', type: r.type || '' }); }

  // GL activity within the window, grouped by account.
  const rows = db.prepare(`
    SELECT je.date AS date, je.entry_num AS entry_num, je.doc_number AS doc_number,
           je.memo AS memo, jl.description AS description,
           jl.account_code AS code, a.type AS type, a.name AS name,
           jl.debit AS debit, jl.credit AS credit
    FROM journal_lines jl
    JOIN journal_entries je ON je.id = jl.entry_id
    LEFT JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code
    WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ?
    ORDER BY jl.account_code, je.date, je.entry_num, jl.id
  `).all(eid, p.windowFrom, p.periodTo);

  const act = new Map(); // code -> [lines]
  for (const r of rows) {
    const code = String(r.code);
    if (!meta.has(code)) meta.set(code, { name: r.name || '', type: r.type || '' });
    if (!act.has(code)) act.set(code, []);
    act.get(code).push({
      date: r.date,
      je: (r.entry_num != null && r.entry_num !== '') ? ('JE-' + pad4(r.entry_num)) : (r.doc_number ? String(r.doc_number) : ''),
      memo: r.memo || r.description || '',
      debit: r2(Number(r.debit) || 0), credit: r2(Number(r.credit) || 0),
    });
  }

  // Build GL blocks (account header + optional opening line + activity, running
  // natural balance). An account appears if it has window activity or a nonzero
  // balance-sheet opening.
  const codes = new Set([...act.keys()]);
  for (const [code, bal] of opening) { if (Math.abs(bal) >= 0.005 && isBS((meta.get(code) || {}).type)) codes.add(code); }
  const blocks = [];
  for (const code of [...codes].sort(sortByCode)) {
    const mt = meta.get(code) || { name: '', type: '' };
    const lines = [];
    let running = 0;
    const op = opening.get(code) || 0;
    if (isBS(mt.type) && Math.abs(op) >= 0.005) {
      running = op;
      lines.push({ date: p.openingAsOf, je: '', memo: 'Opening balance', debit: null, credit: null, balance: r2(running) });
    }
    for (const ln of (act.get(code) || [])) {
      running = r2(running + (isDrNat(mt.type) ? (ln.debit - ln.credit) : (ln.credit - ln.debit)));
      lines.push({ date: ln.date, je: ln.je, memo: ln.memo, debit: ln.debit, credit: ln.credit, balance: r2(running) });
    }
    if (!lines.length) continue;
    blocks.push({ code, name: mt.name, type: mt.type, lines });
  }

  // Which balance-sheet accounts get a "See <code>" supporting tab: every
  // non-zero asset/liability on the TB that has a GL block.
  const blockByCode = new Map(blocks.map((b) => [b.code, b]));
  const support = tb
    .filter((t) => (t.type === 'Asset' || t.type === 'Liability') && blockByCode.has(t.code))
    .map((t) => t.code);

  return { ent, p, tb, blocks, blockByCode, support };
}

// ── Workbook ──────────────────────────────────────────────────────────────────
function titleRows(ws, entName, line2, line3) {
  ws.getCell('A1').value = entName; ws.getCell('A1').font = F();
  ws.getCell('A2').value = line2; ws.getCell('A2').font = F();
  ws.getCell('A3').value = line3; ws.getCell('A3').font = F();
}

function buildTB(wb, d) {
  const { ent, p, tb, support } = d;
  const ws = wb.addWorksheet('TB');
  titleRows(ws, ent.name, 'Trial Balance', 'As of ' + p.tbAsOf);
  const supportSet = new Set(support);
  const HDR = ['Code', 'Account', 'Type', 'Debit', 'Credit', 'Closing Balance on ' + p.clsDate, ''];
  HDR.forEach((h, i) => {
    const c = ws.getCell(5, i + 1); c.value = h; c.font = F({ bold: true });
    c.numFmt = '@'; if (i <= 5) c.fill = fillOf(HDRFILL);
  });
  let row = 6; const firstRow = row; const plRows = [];
  for (const t of tb) {
    const codeVal = (numCode(t.code) !== Number.POSITIVE_INFINITY && /^\d+$/.test(t.code)) ? Number(t.code) : t.code;
    ws.getCell(row, 1).value = codeVal; ws.getCell(row, 1).font = F();
    ws.getCell(row, 2).value = t.name; ws.getCell(row, 2).font = F();
    ws.getCell(row, 3).value = t.type; ws.getCell(row, 3).font = F();
    const dr = t.closing > 0 ? t.closing : null;
    const cr = t.closing < 0 ? r2(-t.closing) : null;
    if (dr != null) { const c = ws.getCell(row, 4); c.value = dr; c.font = F(); c.numFmt = ACCT; }
    if (cr != null) { const c = ws.getCell(row, 5); c.value = cr; c.font = F(); c.numFmt = ACCT; }
    const fc = ws.getCell(row, 6); fc.value = { formula: 'D' + row + '-E' + row, result: t.closing }; fc.font = F(); fc.numFmt = ACCT;
    if (supportSet.has(t.code)) { const nc = ws.getCell(row, 7); nc.value = 'See ' + t.code; nc.font = F({ bold: true }); }
    if (t.type === 'Revenue' || t.type === 'Expense') plRows.push(row);
    row++;
  }
  const lastRow = row - 1;
  row++; // blank
  const totRow = row;
  ws.getCell(totRow, 3).value = 'Total'; ws.getCell(totRow, 3).font = F();
  const totDr = r2(tb.reduce((s, t) => s + (t.closing > 0 ? t.closing : 0), 0));
  const totCr = r2(tb.reduce((s, t) => s + (t.closing < 0 ? -t.closing : 0), 0));
  const cD = ws.getCell(totRow, 4); cD.value = lastRow >= firstRow ? { formula: 'SUM(D' + firstRow + ':D' + lastRow + ')', result: totDr } : totDr; cD.font = F(); cD.numFmt = ACCT;
  const cE = ws.getCell(totRow, 5); cE.value = lastRow >= firstRow ? { formula: 'SUM(E' + firstRow + ':E' + lastRow + ')', result: totCr } : totCr; cE.font = F(); cE.numFmt = ACCT;
  const cF = ws.getCell(totRow, 6); cF.value = lastRow >= firstRow ? { formula: 'SUM(F' + firstRow + ':F' + lastRow + ')', result: r2(totDr - totCr) } : 0; cF.font = F(); cF.numFmt = ACCT;
  // Net Income / Net Loss over the P&L rows.
  const niVal = r2(tb.filter((t) => t.type === 'Revenue' || t.type === 'Expense').reduce((s, t) => s + t.closing, 0));
  const niRow = totRow + 2;
  const lab = ws.getCell(niRow, 5); lab.value = niVal >= 0 ? 'Net Income' : 'Net Loss'; lab.font = F(); lab.fill = fillOf(YELLOW);
  const niF = ws.getCell(niRow, 6);
  let niFormula = '0';
  if (plRows.length) {
    const contiguous = plRows[plRows.length - 1] - plRows[0] === plRows.length - 1;
    niFormula = contiguous ? ('SUM(F' + plRows[0] + ':F' + plRows[plRows.length - 1] + ')') : plRows.map((r) => 'F' + r).join('+');
  }
  niF.value = plRows.length ? { formula: niFormula, result: niVal } : 0; niF.font = F(); niF.numFmt = ACCT; niF.fill = fillOf(YELLOW);
  // Column widths (match the delivered packages).
  const W = [18.2, 41.0, 7.9, 12.9, 12.9, 24.0, 20.0];
  W.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

// One GL block (used by the GL tab and by each supporting tab).
function writeBlock(ws, startRow, block) {
  let row = startRow;
  const h = ws.getCell(row, 1); h.value = block.code + ' - ' + block.name; h.font = F({ bold: true }); row++;
  ['Date', 'JE', 'Memo', 'Debit', 'Credit', 'Balance'].forEach((t, i) => { const c = ws.getCell(row, i + 1); c.value = t; c.font = F({ bold: true }); });
  row++;
  for (const ln of block.lines) {
    ws.getCell(row, 1).value = ln.date; ws.getCell(row, 1).font = F();
    ws.getCell(row, 2).value = ln.je; ws.getCell(row, 2).font = F();
    ws.getCell(row, 3).value = ln.memo; ws.getCell(row, 3).font = F();
    if (ln.debit != null && ln.debit !== 0) { const c = ws.getCell(row, 4); c.value = ln.debit; c.font = F(); c.numFmt = ACCT; }
    if (ln.credit != null && ln.credit !== 0) { const c = ws.getCell(row, 5); c.value = ln.credit; c.font = F(); c.numFmt = ACCT; }
    const b = ws.getCell(row, 6); b.value = ln.balance; b.font = F(); b.numFmt = ACCT;
    row++;
  }
  return row; // next free row
}

function glColWidths(ws) {
  const W = [47.8, 7.2, 76.6, 12.9, 12.9, 12.9];
  W.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
}

function buildGL(wb, d) {
  const { ent, p, blocks } = d;
  const ws = wb.addWorksheet('GL', { views: [{ state: 'frozen', ySplit: 3 }] });
  titleRows(ws, ent.name, 'General Ledger', p.periodLine);
  let row = 5;
  for (const b of blocks) { row = writeBlock(ws, row, b); row++; /* blank between blocks */ }
  glColWidths(ws);
}

function buildSupportTabs(wb, d, used) {
  const { ent, p, support, blockByCode } = d;
  for (const code of support) {
    const block = blockByCode.get(code); if (!block) continue;
    const ws = wb.addWorksheet(sheetNameFor(code, used));
    titleRows(ws, ent.name, 'General Ledger', p.periodLine);
    writeBlock(ws, 5, block);
    glColWidths(ws);
  }
}

function buildWorkbook(d) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  addGlDerivedSheets(wb, d, true);
  return wb;
}

// Add the GL-derived sheets (TB, GL, supporting tabs, optional Org Chart) into a
// workbook — either a fresh one (standalone) or a loaded reference (carry).
function addGlDerivedSheets(wb, d, includeOrgChart) {
  const used = new Set(wb.worksheets.map((w) => String(w.name).toLowerCase()));
  used.add('tb'); used.add('gl');
  buildTB(wb, d);
  buildGL(wb, d);
  buildSupportTabs(wb, d, used);
  if (includeOrgChart && !used.has('org chart')) wb.addWorksheet('Org Chart'); // placeholder — diagram pasted in by hand
}

// A sheet in the reference is "GL-derived" (owned by this generator, regenerated
// each run) when it is the current TB/GL or a numeric per-account tab. Everything
// else — Equity Rollforward, Contributed Capital, Org Chart, cap tables,
// joinders, a prior-year TB like "TB 2023" — is bespoke and carried verbatim.
function isGlDerivedSheet(name, year) {
  const n = String(name || '').trim();
  if (/^(tb|gl|trial balance|general ledger)$/i.test(n)) return true;
  if (/^\d+$/.test(n)) return true;
  if (new RegExp('^(tb|trial balance)\\s*' + year + '$', 'i').test(n)) return true;
  if (new RegExp('^' + year + '\\s*(tb|trial balance)$', 'i').test(n)) return true;
  return false;
}
function bespokeSheetNames(names, year) { return (names || []).filter((n) => !isGlDerivedSheet(n, year)); }

// Carry mode: load the reference as the base, drop its GL-derived sheets, and add
// freshly-built TB/GL/supporting tabs. The bespoke tabs (and their embedded
// images) pass through untouched.
async function buildCarriedWorkbook(d, refBuf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(refBuf);
  for (const name of wb.worksheets.map((w) => w.name)) {
    if (isGlDerivedSheet(name, d.p.year)) { const ws = wb.getWorksheet(name); if (ws) wb.removeWorksheet(ws.id); }
  }
  addGlDerivedSheets(wb, d, false); // keep the reference's Org Chart, don't add a placeholder
  return wb;
}

// Tab order after a carry: TB, GL, numeric supporting tabs (ascending), then the
// bespoke tabs in their original order. Reorders the <sheet> entries in
// workbook.xml only — parts, relationships and drawings are left untouched.
async function reorderSheets(buf) {
  let JSZip; try { JSZip = require('jszip'); } catch (e) { return buf; }
  try {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file('xl/workbook.xml'); if (!f) return buf;
    let xml = await f.async('string');
    const m = xml.match(/<sheets>([\s\S]*?)<\/sheets>/);
    if (!m) return buf;
    const els = m[1].match(/<sheet\b[^>]*\/>/g) || [];
    if (els.length < 2) return buf;
    const nameOf = (el) => { const mm = el.match(/name="([^"]*)"/); return mm ? mm[1].trim() : ''; };
    const rank = (el) => { const n = nameOf(el); if (/^tb$/i.test(n)) return [0, 0]; if (/^gl$/i.test(n)) return [1, 0]; if (/^\d+$/.test(n)) return [2, Number(n)]; return [3, 0]; };
    const ordered = els.map((el, i) => ({ el, i, r: rank(el) }))
      .sort((a, b) => (a.r[0] - b.r[0]) || (a.r[1] - b.r[1]) || (a.i - b.i))
      .map((x) => x.el).join('');
    xml = xml.replace(/<sheets>[\s\S]*?<\/sheets>/, '<sheets>' + ordered + '</sheets>');
    zip.file('xl/workbook.xml', xml);
    return Buffer.from(await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  } catch (e) { return buf; }
}

// ── Filing to the Workpapers folder ───────────────────────────────────────────
function safeName(s) { return String(s || '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, ''); }
function folderFor(p) { return 'Workpapers/Tax Packages/' + p.title + '/' + p.year; }
function fileNameFor(ent, p) {
  const nm = safeName(ent.name) || ('entity_' + ent.id);
  return nm + '_TB_GL_Tax_Package_' + p.label + '.xlsx';
}
function saveToWorkpapers(ctx, eid, ent, p, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(p), original = fileNameFor(ent, p);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const pr of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), pr.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(pr.id);
  }
  const dir = path.join(workpapersDir, String(eid)); fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function ensureRefSchema(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS taxpkg_reference (
    entity_id INTEGER PRIMARY KEY,
    original_name TEXT,
    data BLOB,
    sheet_names TEXT,
    uploaded_by TEXT,
    created_at TEXT
  )`);
}

function registerTaxPackageRoutes(app, ctx) {
  const { db, auth, requireEntityAccess, requireRole } = ctx;
  ensureRefSchema(db);
  const refUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

  app.post('/api/workpapers/tax-package/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const mode = (req.body && req.body.mode) === 'monthly' ? 'monthly' : 'annual';
        const p = resolvePeriod((req.body && req.body.period_end) || '', mode);
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const d = buildData(ctx, eid, p);
        const ref = db.prepare('SELECT data, sheet_names FROM taxpkg_reference WHERE entity_id = ?').get(eid);
        let buf; let carried = null;
        if (ref && ref.data && ref.data.length) {
          const wb = await buildCarriedWorkbook(d, ref.data);
          buf = await reorderSheets(Buffer.from(await wb.xlsx.writeBuffer()));
          let names = []; try { names = JSON.parse(ref.sheet_names || '[]'); } catch (e) { names = []; }
          carried = bespokeSheetNames(names, p.year);
        } else {
          const wb = buildWorkbook(d);
          buf = Buffer.from(await wb.xlsx.writeBuffer());
        }
        let filed = null;
        try { filed = saveToWorkpapers(ctx, eid, d.ent, p, buf, who); } catch (e) { console.error('[tax-package] filing failed:', e.message); }
        const fname = fileNameFor(d.ent, p);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
        res.setHeader('X-TaxPackage-Summary', JSON.stringify({
          mode: p.mode, title: p.title, label: p.label, as_of: p.tbAsOf,
          accounts: d.tb.length, gl_accounts: d.blocks.length, supporting: d.support.length,
          carried: carried || null, carried_count: carried ? carried.length : 0,
          folder: filed ? filed.folder_path : null, replaced: filed ? filed.replaced : 0,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });

  // Reference package: a prior tax package whose bespoke tabs (Equity Rollforward,
  // Contributed Capital, Org Chart, cap tables, …) are carried into every generated
  // package. TB, GL and the per-account tabs are always regenerated from the ledger.
  app.get('/api/workpapers/tax-package/:entity_id/reference', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const row = db.prepare('SELECT original_name, sheet_names, uploaded_by, created_at FROM taxpkg_reference WHERE entity_id = ?').get(eid);
        if (!row) return res.json({ has_reference: false });
        let names = []; try { names = JSON.parse(row.sheet_names || '[]'); } catch (e) { names = []; }
        res.json({ has_reference: true, original_name: row.original_name, uploaded_by: row.uploaded_by, uploaded_at: row.created_at, sheets: names, carried: bespokeSheetNames(names, '') });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });

  app.post('/api/workpapers/tax-package/:entity_id/reference', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), refUpload.single('file'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        if (!req.file || !req.file.buffer || !req.file.buffer.length) return res.status(400).json({ error: 'No file uploaded' });
        const buf = req.file.buffer;
        if (buf.slice(0, 2).toString('latin1') !== 'PK') return res.status(400).json({ error: 'The reference must be an .xlsx workbook' });
        let sheetNames = [];
        try { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf); sheetNames = wb.worksheets.map((w) => w.name); }
        catch (e) { return res.status(400).json({ error: 'Could not read the workbook: ' + e.message }); }
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        db.prepare(`INSERT INTO taxpkg_reference (entity_id, original_name, data, sheet_names, uploaded_by, created_at)
          VALUES (?, ?, ?, ?, ?, datetime('now'))
          ON CONFLICT(entity_id) DO UPDATE SET original_name = excluded.original_name, data = excluded.data,
            sheet_names = excluded.sheet_names, uploaded_by = excluded.uploaded_by, created_at = datetime('now')`)
          .run(eid, req.file.originalname || 'reference.xlsx', buf, JSON.stringify(sheetNames), who);
        res.json({ has_reference: true, original_name: req.file.originalname || 'reference.xlsx', sheets: sheetNames, carried: bespokeSheetNames(sheetNames, '') });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });

  app.delete('/api/workpapers/tax-package/:entity_id/reference', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        db.prepare('DELETE FROM taxpkg_reference WHERE entity_id = ?').run(eid);
        res.json({ ok: true, has_reference: false });
      } catch (e) { res.status(400).json({ error: e.message }); }
    });
}

module.exports = { registerTaxPackageRoutes, buildData, buildWorkbook, buildCarriedWorkbook, reorderSheets, resolvePeriod };
