// ═══════════════════════════════════════════════════════════════════════════
// Administration: Monthly Time Allocation
//
// Portfolio-company time allocation reports for Jimmy Yun (CAO) and Irvin
// Bermudez (Controller). Each month the only thing that changes is the date —
// the allocation itself (companies, %, hours, activity detail) carries forward
// unchanged. Given a month-end date this:
//
//   1. Finds each person's most recent prior .xlsx in the Time Allocation folder
//      and uses it as the template (so the latest content is preserved).
//   2. Rolls the dates forward:
//        - Title (A1)            -> "{Month} {Year} Portfolio Company Time
//                                    Allocation - {Name}" (set canonically, so a
//                                    stale copy-pasted title is corrected).
//        - "Month" field (B7)    -> the new month/year, keeping the template's
//                                    own style (abbreviated vs. full).
//        - Signature / Approval  -> the run date; each person's on-file
//          dates                    signature image is stamped onto the
//                                   employee line, and the approver's signature
//                                   (Ben for Jimmy's report, Jimmy for Irvin's)
//                                   onto the "Approved By CAO/CEO" line. Images
//                                   live locally in ~/.cloudledger/signatures
//                                   (never committed to git).
//   3. Writes  {Name}'s .xlsx and a print-faithful .pdf into
//        Accounting - Documents / General / 2. Monthly Time Allocation / {Year}
//
// Runs on Jimmy's PC via the shared reports runner (same as Compliance
// Certificates): the .xlsx is edited with ExcelJS and the .pdf is produced with
// Excel itself (COM, scripts/xlsx-to-pdf.ps1) so it matches a hand-made
// "Save as PDF" exactly — honoring the sheet's landscape / fit-to-one-page /
// A1:D## print area. OneDrive then syncs the files up to SharePoint.
// Nothing is emailed; the reports are stamped with each person's on-file
// signature so they come out signed in one pass.
// ═══════════════════════════════════════════════════════════════════════════

const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFileSync } = require('child_process');

const LIB = 'Accounting - Documents';
const BASE = 'General/2. Monthly Time Allocation';
const DEST_BASE = BASE;   // files land in BASE/{year}/

const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = MON.map((m) => m.slice(0, 3));
const pad2 = (n) => String(n).padStart(2, '0');

// People the report is produced for. `match` identifies both the template file
// and (defensively) the signer, from a filename or a label.
const PEOPLE = [
  { key: 'jimmy', name: 'Jimmy Yun', match: /jimmy\s*yun/i },
  { key: 'irvin', name: 'Irvin Bermudez', match: /irvin\s*bermudez/i },
];

const DEFAULTS = {};   // no editable configuration today; kept for route symmetry

// Who signs which line on each person's report:
//   employee  — the person whose report it is signs the Employee Signature line
//   approver  — the "Approved By CAO/CEO" line is signed by the approver, where
//               we hold that person's signature on file. The CAO (Jimmy)
//               approves the Controller's (Irvin's) report. The CEO's (Ben's)
//               approval of the CAO's (Jimmy's) own report is NOT stamped — that
//               line is left blank for Ben to sign off by hand (approver: null).
// A null/absent approver leaves that line blank. Signature images (transparent
// PNGs) live locally, never in git.
const SIGN = {
  jimmy: { employee: 'jimmy', approver: null },
  irvin: { employee: 'irvin', approver: 'jimmy' },
};
function signaturesDir() {
  return process.env.TA_SIGNATURES_DIR || path.join(os.homedir(), '.cloudledger', 'signatures');
}
// PNG intrinsic size from the IHDR chunk (width @ byte 16, height @ 20, BE).
function pngSize(buf) {
  if (!buf || buf.length < 24 || buf.readUInt32BE(0) !== 0x89504e47) return null;
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}

// ── month-end helpers ───────────────────────────────────────────────────────
function parseMonthEnd(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Month-end date must be YYYY-MM-DD');
  const y = +m[1], mo = +m[2], d = +m[3];
  if (mo < 1 || mo > 12) throw new Error('Not a valid month');
  const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  if (d !== last) throw new Error('Month-end must be the last day of the month (' + y + '-' + pad2(mo) + '-' + pad2(last) + ')');
  return { iso: s, y, m: mo, d, ym: y * 12 + mo, monthFull: MON[mo - 1], monthAbbr: MON3[mo - 1], year: y, label: MON[mo - 1] + ' ' + y };
}
function lastMonthEnd(today = new Date()) {   // last day of the previous month
  const dt = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), 0));
  return dt.toISOString().slice(0, 10);
}

// Match a "<month> <year>" token (full or 3-letter, optional trailing dot),
// longest alternative first so "September" wins over "Sep".
const MONTH_ALT = MON.concat(MON3).sort((a, b) => b.length - a.length).join('|');
const RE_MONTH_YEAR = new RegExp('\\b(' + MONTH_ALT + ')\\.?\\s+(\\d{4})\\b', 'i');
function isAbbrToken(tok) {
  const t = tok.replace(/\.$/, '').toLowerCase();
  return MON3.some((a) => a.toLowerCase() === t) && !MON.some((f) => f.toLowerCase() === t);
}
// Replace the first month/year token in a string with the target month/year,
// keeping the matched token's own style (abbreviated vs. full).
function replaceMonthYear(str, T) {
  return String(str).replace(RE_MONTH_YEAR, (_m, mon) => (isAbbrToken(mon) ? T.monthAbbr : T.monthFull) + ' ' + T.year);
}
// Parse the reporting month out of a time-allocation filename.
function monthFromName(name) {
  const m = RE_MONTH_YEAR.exec(name);
  if (!m) return null;
  const i = MON.findIndex((f) => f.toLowerCase() === m[1].toLowerCase());
  const j = i >= 0 ? i : MON3.findIndex((a) => a.toLowerCase() === m[1].replace(/\.$/, '').toLowerCase());
  if (j < 0) return null;
  return { y: +m[2], m: j + 1, ym: (+m[2]) * 12 + (j + 1) };
}

// ── local OneDrive-synced store ─────────────────────────────────────────────
// Time allocation only ever runs where the folder is synced (the runner on
// Jimmy's PC, or a dev machine with the root set) — it also needs Excel for the
// PDF — so there is no Graph path here.
function localRoot() { return process.env.TA_LOCAL_ROOT || process.env.CC_LOCAL_ROOT || ''; }
function makeStore() {
  const root = localRoot();
  const writeRoot = process.env.TA_LOCAL_WRITE_ROOT || process.env.CC_LOCAL_WRITE_ROOT || root;
  if (!root) throw new Error('Time Allocation local root is not set (TA_LOCAL_ROOT / CC_LOCAL_ROOT)');
  const absIn = (rel) => path.join(root, LIB, ...String(rel).split('/').filter(Boolean));
  const absOut = (rel) => path.join(writeRoot, LIB, ...String(rel).split('/').filter(Boolean));
  return {
    root, writeRoot,
    // recursive list of .xlsx under BASE, as { rel (relative to BASE), name, full }
    listXlsx() {
      const baseAbs = absIn(BASE);
      const out = [];
      const walk = (dir, rel) => {
        let ents; try { ents = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
        for (const e of ents) {
          const full = path.join(dir, e.name), r = rel ? rel + '/' + e.name : e.name;
          if (e.isDirectory()) walk(full, r);
          else if (/\.xlsx$/i.test(e.name) && !/^~\$/.test(e.name)) out.push({ rel: r, name: e.name, full });
        }
      };
      walk(baseAbs, '');
      return out;
    },
    readFull(full) { return fs.readFileSync(full); },
    // write buffer to BASE/rel, return the absolute path
    write(rel, buf) { const f = absOut(BASE + '/' + rel); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, buf); return f; },
    outDir(rel) { return absOut(BASE + '/' + rel); },
  };
}

// ── workbook date roll-forward (ExcelJS) ────────────────────────────────────
// The template keeps all of its own formatting, merges, column widths and page
// setup; only the dated cells are changed, then the signatures are stamped.
function rollForward(wb, ws, person, T, runDate, out, say) {
  // 1) Title (row 1, master A1) — set canonically so a stale title is fixed.
  ws.getCell('A1').value = T.monthFull + ' ' + T.year + ' Portfolio Company Time Allocation - ' + person.name;

  // Scan column A for the labelled rows and update the value cell in column B.
  const labelRow = {};
  for (let r = 1; r <= ws.rowCount; r++) {
    const a = ws.getCell(r, 1).value;
    const t = (a == null ? '' : String(typeof a === 'object' && a && a.richText ? a.richText.map((x) => x.text).join('') : a)).trim().toLowerCase();
    if (t) labelRow[t] = r;
  }
  // 2) "Month" field — keep the template's own abbreviated/full style.
  if (labelRow['month']) {
    const c = ws.getCell(labelRow['month'], 2);
    if (c.value instanceof Date) c.value = new Date(Date.UTC(T.y, T.m - 1, 1));
    else if (typeof c.value === 'string' && RE_MONTH_YEAR.test(c.value)) c.value = replaceMonthYear(c.value, T);
    else c.value = T.monthFull + ' ' + T.year;
  }
  // 3) Signature / approval dates — set to the run date. A formula cell (e.g.
  //    Approval Date = the signature cell) is left alone so it keeps following.
  for (const lbl of ['employee signature date', 'approval date']) {
    if (!labelRow[lbl]) continue;
    const c = ws.getCell(labelRow[lbl], 2);
    if (c.value && typeof c.value === 'object' && c.value.formula) continue;
    c.value = new Date(Date.UTC(runDate.getFullYear(), runDate.getMonth(), runDate.getDate()));
  }
  // 4) Stamp the signature images onto the employee and approval lines.
  applySignatures(wb, ws, person, labelRow, out, say);
}

// Anchor each person's on-file signature image over the value cell (column B,
// merged B:D) of its signature row. The image floats above the cell, sized to a
// natural signature width with its own aspect preserved; it rotates and scales
// with the sheet on PDF export, so Jimmy's rotated page is handled the same way.
function applySignatures(wb, ws, person, labelRow, out, say) {
  const map = SIGN[person.key];
  if (!map) return;
  const dir = signaturesDir();
  const empRow = labelRow['employee signature'];
  let apprRow = null;
  for (const [k, r] of Object.entries(labelRow)) { if (/^approved by\b/.test(k)) { apprRow = r; break; } }

  const place = (sigKey, row, slot) => {
    if (!sigKey || !row) return;
    const file = path.join(dir, sigKey + '.png');
    let buf; try { buf = fs.readFileSync(file); }
    catch (e) { if (out) out.warnings.push(person.name + ': ' + slot + ' line left blank — signature ' + sigKey + '.png not found'); return; }
    const sz = pngSize(buf);
    const targetW = 150;                                   // ~1.5in at 96dpi
    const h = sz ? Math.max(12, Math.round(targetW * sz.h / sz.w)) : 24;
    const imageId = wb.addImage({ buffer: buf, extension: 'png' });
    // tl is 0-based col/row with a fractional inset into column B of the row.
    ws.addImage(imageId, {
      tl: { col: 1.12, row: (row - 1) + 0.12 },
      ext: { width: targetW, height: h },
      editAs: 'oneCell',
    });
    if (say) say(person.name + ': signed ' + slot + ' (' + sigKey + ')');
  };
  place(map.employee, empRow, 'employee');
  place(map.approver, apprRow, 'approver');
}

// ── Excel COM (PDF) ─────────────────────────────────────────────────────────
// Convert each .xlsx to .pdf with Excel itself, honoring the workbook's own
// print area / landscape / fit-to-one-page — identical to a manual export.
function convertToPdf(pairs, log) {
  if (!pairs.length) return;
  const ps1 = path.join(__dirname, '..', 'scripts', 'xlsx-to-pdf.ps1');
  if (!fs.existsSync(ps1)) throw new Error('xlsx-to-pdf.ps1 not found at ' + ps1);
  const tmp = path.join(os.tmpdir(), 'ta_pairs_' + Date.now() + '.tsv');
  fs.writeFileSync(tmp, pairs.map((p) => p.xlsx + '\t' + p.pdf).join('\r\n'), 'utf8');
  try {
    const out = execFileSync('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-PairsFile', tmp],
      { encoding: 'utf8', timeout: 5 * 60 * 1000, windowsHide: true });
    String(out || '').split(/\r?\n/).forEach((l) => { if (l.trim() && log) log('  pdf: ' + l.trim()); });
  } catch (e) {
    const msg = (e.stdout || '') + (e.stderr || '') || e.message;
    throw new Error('Excel PDF conversion failed: ' + String(msg).replace(/\s+/g, ' ').trim().slice(0, 400));
  } finally { try { fs.unlinkSync(tmp); } catch (e) { /* ignore */ } }
}

// ════════════════════════════════════════════════════════════════════════════
// The report job
// ════════════════════════════════════════════════════════════════════════════
async function runReport(monthEnd, cfg, log) {
  const ExcelJS = require('exceljs');
  const T = parseMonthEnd(monthEnd);
  const store = makeStore();
  const runDate = new Date();
  const out = { month_end: T.iso, month: T.label, storage: 'local', dest: DEST_BASE + '/' + T.year, items: [], warnings: [] };
  const say = (s) => { if (log) log(s); };
  const item = (who, file, status, note) => { out.items.push({ pkg: who, file, status, note: note || '' }); say(who + ': ' + status + ' - ' + file + (note ? ' (' + note + ')' : '')); };

  const all = store.listXlsx();
  const pdfPairs = [];

  for (const person of PEOPLE) {
    try {
      // candidate templates for this person, newest reporting month first
      const mine = all
        .map((f) => ({ f, mo: monthFromName(f.name) }))
        .filter((x) => x.mo && person.match.test(x.f.name))
        .sort((a, b) => b.mo.ym - a.mo.ym);
      if (!mine.length) { item(person.name, '(no template)', 'error', 'no prior .xlsx found for ' + person.name); continue; }
      // prefer the latest month strictly before the target (clean overwrite /
      // roll-forward); else the overall latest.
      const tpl = (mine.find((x) => x.mo.ym < T.ym) || mine[0]).f;

      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(store.readFull(tpl.full));
      const ws = wb.worksheets.find((w) => (w.pageSetup && w.pageSetup.printArea)) || wb.worksheets[0];
      rollForward(wb, ws, person, T, runDate, out, say);

      const stem = T.monthFull + ' ' + T.year + ' Portfolio Company Time Allocation - ' + person.name;
      const xlsxRel = T.year + '/' + stem + '.xlsx';
      const pdfRel = T.year + '/' + stem + '.pdf';
      const xlsxFull = store.write(xlsxRel, Buffer.from(await wb.xlsx.writeBuffer()));
      item(person.name, stem + '.xlsx', 'created', 'from ' + tpl.name);
      pdfPairs.push({ xlsx: xlsxFull, pdf: store.outDir(pdfRel), who: person.name, file: stem + '.pdf' });
    } catch (e) {
      item(person.name, '(error)', 'error', e.message);
    }
  }

  // Render all PDFs in one Excel session.
  if (pdfPairs.length) {
    try {
      convertToPdf(pdfPairs.map((p) => ({ xlsx: p.xlsx, pdf: p.pdf })), say);
      for (const p of pdfPairs) {
        if (fs.existsSync(p.pdf)) item(p.who, p.file, 'created', 'Excel PDF export');
        else item(p.who, p.file, 'error', 'Excel did not produce the PDF');
      }
    } catch (e) {
      for (const p of pdfPairs) item(p.who, p.file, 'error', e.message);
      out.warnings.push('PDF step failed — the .xlsx files were saved; open each and "Save as PDF" if needed. ' + e.message);
    }
  }

  out.counts = out.items.reduce((a, it) => { a[it.status] = (a[it.status] || 0) + 1; return a; }, {});
  return out;
}

// ════════════════════════════════════════════════════════════════════════════
// Routes  (mirrors the Compliance Certificates run/runner pattern)
// ════════════════════════════════════════════════════════════════════════════
function ensureSchema(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS time_alloc_runs (id INTEGER PRIMARY KEY AUTOINCREMENT, month_end TEXT NOT NULL, status TEXT NOT NULL, log TEXT, result TEXT, error TEXT,
    by_who TEXT, started_at TEXT DEFAULT (datetime('now')), finished_at TEXT, claimed_at TEXT);`);
}

function registerTimeAllocationRoutes(app, ctx) {
  const { db, auth, requireRole } = ctx;
  ensureSchema(db);
  const crypto = require('crypto');
  const running = new Map();
  const who = (req) => (req.user && (req.user.email || req.user.username || req.user.name)) || 'unknown';
  const runnerKey = () => { const n = Object.keys(process.env).find((k) => k.toUpperCase() === 'CC_RUNNER_KEY'); return n ? String(process.env[n]).trim() : ''; };
  const mode = () => (localRoot() ? 'local' : runnerKey() ? 'runner' : 'none');
  let runnerSeen = 0, runnerHost = '';
  const runnerAuth = (req, res, next) => {
    const k = runnerKey(), got = String(req.get('x-runner-key') || '');
    if (!k || got.length !== k.length || !crypto.timingSafeEqual(Buffer.from(got), Buffer.from(k))) return res.status(401).json({ error: 'invalid runner key' });
    runnerSeen = Date.now(); runnerHost = String(req.get('x-runner-host') || '').slice(0, 80); next();
  };

  app.get('/api/admin/time-allocation', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const runs = db.prepare('SELECT id, month_end, status, by_who, started_at, finished_at, error FROM time_alloc_runs ORDER BY id DESC LIMIT 15').all();
    res.json({
      storage: mode(),
      runner_configured: !!runnerKey(), runner_online: Date.now() - runnerSeen < 45000, runner_host: runnerHost,
      runner_last_seen: runnerSeen ? new Date(runnerSeen).toISOString() : null,
      default_month_end: lastMonthEnd(), dest_base: DEST_BASE, people: PEOPLE.map((p) => p.name), runs,
    });
  });

  app.post('/api/admin/time-allocation/report', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    let T; try { T = parseMonthEnd((req.body || {}).month_end); } catch (e) { return res.status(400).json({ error: e.message }); }
    if ([...running.values()].some((r) => r.month_end === T.iso) ||
        db.prepare("SELECT id FROM time_alloc_runs WHERE month_end = ? AND status IN ('queued','running')").get(T.iso))
      return res.status(409).json({ error: 'A report for ' + T.label + ' is already queued or running' });
    if (mode() === 'none') return res.status(400).json({ error: 'Report runner is not configured (set CC_RUNNER_KEY)' });
    if (mode() === 'runner') {
      const qid = db.prepare("INSERT INTO time_alloc_runs (month_end, status, log, by_who) VALUES (?, 'queued', '', ?)").run(T.iso, who(req)).lastInsertRowid;
      return res.json({ id: qid, month: T.label, queued: true });
    }
    // local (dev) mode — run in-process
    const id = db.prepare("INSERT INTO time_alloc_runs (month_end, status, log, by_who) VALUES (?, 'running', '', ?)").run(T.iso, who(req)).lastInsertRowid;
    const state = { month_end: T.iso, log: [] }; running.set(id, state);
    const upd = db.prepare('UPDATE time_alloc_runs SET log = ? WHERE id = ?');
    runReport(T.iso, {}, (line) => { state.log.push(line); upd.run(state.log.join('\n'), id); })
      .then((result) => db.prepare("UPDATE time_alloc_runs SET status = 'done', result = ?, finished_at = datetime('now') WHERE id = ?").run(JSON.stringify(result), id))
      .catch((e) => db.prepare("UPDATE time_alloc_runs SET status = 'failed', error = ?, finished_at = datetime('now') WHERE id = ?").run(String(e && e.message || e), id))
      .finally(() => running.delete(id));
    res.json({ id, month: T.label });
  });

  // ── runner endpoints (authenticated by CC_RUNNER_KEY) ──
  app.post('/api/time-allocation-runner/heartbeat', runnerAuth, (req, res) => res.json({ ok: true }));
  app.post('/api/time-allocation-runner/claim', runnerAuth, (req, res) => {
    db.prepare("UPDATE time_alloc_runs SET status = 'failed', error = 'Runner stopped before finishing (timed out)', finished_at = datetime('now') WHERE status = 'running' AND claimed_at IS NOT NULL AND claimed_at < datetime('now', '-30 minutes')").run();
    const job = db.transaction(() => {
      const r = db.prepare("SELECT id, month_end FROM time_alloc_runs WHERE status = 'queued' ORDER BY id LIMIT 1").get();
      if (!r) return null;
      db.prepare("UPDATE time_alloc_runs SET status = 'running', claimed_at = datetime('now'), log = ? WHERE id = ?").run('Picked up by ' + (runnerHost || 'runner'), r.id);
      return r;
    })();
    res.json({ job: job ? { id: job.id, month_end: job.month_end } : null });
  });
  app.post('/api/time-allocation-runner/runs/:id/log', runnerAuth, (req, res) => {
    const r = db.prepare('SELECT log FROM time_alloc_runs WHERE id = ?').get(+req.params.id); if (!r) return res.status(404).json({ error: 'not found' });
    const lines = ((req.body || {}).lines || []).map(String).slice(0, 200);
    db.prepare('UPDATE time_alloc_runs SET log = ? WHERE id = ?').run([r.log || ''].concat(lines).filter(Boolean).join('\n'), +req.params.id);
    res.json({ ok: true });
  });
  app.post('/api/time-allocation-runner/runs/:id/finish', runnerAuth, (req, res) => {
    const b = req.body || {};
    if (b.ok) db.prepare("UPDATE time_alloc_runs SET status = 'done', result = ?, finished_at = datetime('now') WHERE id = ?").run(JSON.stringify(b.result || {}), +req.params.id);
    else db.prepare("UPDATE time_alloc_runs SET status = 'failed', error = ?, finished_at = datetime('now') WHERE id = ?").run(String(b.error || 'unknown error').slice(0, 2000), +req.params.id);
    res.json({ ok: true });
  });

  app.get('/api/admin/time-allocation/runs/:id', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    const r = db.prepare('SELECT * FROM time_alloc_runs WHERE id = ?').get(+req.params.id);
    if (!r) return res.status(404).json({ error: 'not found' });
    res.json(Object.assign({}, r, { result: r.result ? JSON.parse(r.result) : null, log: r.log ? r.log.split('\n') : [] }));
  });
}

module.exports = { registerTimeAllocationRoutes, runReport, parseMonthEnd, DEFAULTS };
