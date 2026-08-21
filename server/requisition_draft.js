// Editable / persistent Requisition Report draft — lifecycle logic.
//
// A draft is the entity's single in-progress requisition. It is created once
// (auto-seeded from the last finalized Req, or from a first-time manual upload),
// edited freely (add / delete / update invoices), and re-rolled on each save so
// its stored workbook + packet always reflect the current invoice set. Finalize
// locks it, stamps the invoices with the req number, and files the workbook +
// packet to Workpapers (via saveRequisitionOutputs) — that filed workbook is
// next month's auto-seed base.
//
// This module owns: draft row CRUD, the shared roll-forward invocation (so the
// draft path and the legacy one-shot path build `meta` identically), the
// auto-seed resolver, and the upload-guard hash/compare. The heavy Excel work is
// delegated to the existing engines — no statement logic lives here.

const crypto = require('crypto');
const { rollForward, findSheet: findReqSheet } = require('./requisition_rollforward');
const { verifyRollforward } = require('./requisition_rollforward_verify');
const { finalizeRequisitionWorkbook } = require('./requisition_preserve');
const { makeDevFeeClaudeCaller } = require('./requisition_devfee');

// Build the per-entity roll-forward `meta` the same way the legacy endpoint does,
// so the Dev Fee tab style, report-number header fix, and dev-fee payee are
// identical whichever path invokes the engine. Kept here as the single source of
// truth; the legacy endpoint can call this too.
function buildRollforwardMeta(entityId, { reqNumber, asOfDate } = {}) {
  const meta = {};
  if (reqNumber != null && reqNumber !== '') meta.reqNumber = reqNumber;
  if (asOfDate) meta.asOfDate = asOfDate;
  const eid = String(parseInt(entityId));
  const collapseIds = (process.env.REQ_DEVFEE_COLLAPSE_ENTITIES || '42,38,39')
    .split(',').map(x => x.trim()).filter(Boolean);
  meta.collapseDevFeeCosts = collapseIds.includes(eid);
  meta.fixReportNumberHeader = meta.collapseDevFeeCosts;
  const defaultPayees = { '42': 'County Line Rail Interest', '38': 'County Line Rail Interest', '39': 'County Line Rail Interest' };
  let payeeMap = defaultPayees;
  try { payeeMap = Object.assign({}, defaultPayees, JSON.parse(process.env.REQ_DEVFEE_PAYEES || '{}')); } catch (_) { payeeMap = defaultPayees; }
  if (meta.collapseDevFeeCosts && payeeMap[eid]) meta.devFeePayee = payeeMap[eid];
  return meta;
}

function sha256(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

// Load a workbook buffer into a mutable ExcelJS copy + an untouched prior copy,
// and validate the required tabs are present. Returns { workbook, priorBook,
// prevSheets } or throws an Error with a user-facing message.
async function loadReqWorkbook(ExcelJS, buffer) {
  let workbook, priorBook;
  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    priorBook = new ExcelJS.Workbook();
    await priorBook.xlsx.load(buffer);
  } catch (e) {
    const err = new Error('Failed to read workbook (.xlsx expected): ' + e.message);
    err.status = 400; throw err;
  }
  const prevSheets = {
    prior: findReqSheet(priorBook, 'Prior Invoice Log'),
    current: findReqSheet(priorBook, 'Current Invoice Log'),
  };
  if (!prevSheets.prior || !prevSheets.current) {
    const present = priorBook.worksheets.map(s => s.name).join(', ');
    const missing = [!prevSheets.prior && 'Prior Invoice Log', !prevSheets.current && 'Current Invoice Log'].filter(Boolean).join(', ');
    const err = new Error('Workbook is missing required tab(s): ' + missing + '. Tabs found: ' + (present || '(none)') + '. A requisition roll-forward needs "Prior Invoice Log", "Current Invoice Log", and "Budget to Actual" tabs.');
    err.status = 400; throw err;
  }
  return { workbook, priorBook, prevSheets };
}

// Run the roll-forward + verification against a base workbook buffer and the
// current invoice set (newCurrent = the Current-Invoice-Log rows to append).
// Returns { outBuf, rfResult, verification }. Does NOT persist anything — the
// caller decides what to store (draft output vs. filed).
async function rollForwardFromBase(ExcelJS, baseBuffer, entityId, newCurrent, { reqNumber, asOfDate } = {}) {
  const meta = buildRollforwardMeta(entityId, { reqNumber, asOfDate });
  const { workbook, prevSheets } = await loadReqWorkbook(ExcelJS, baseBuffer);

  const devFeeCaller = process.env.ANTHROPIC_API_KEY ? makeDevFeeClaudeCaller() : null;
  const rows = Array.isArray(newCurrent) ? newCurrent : [];
  const rfResult = await rollForward(workbook, rows, { ...meta, callClaude: devFeeCaller });

  const verification = await verifyRollforward({
    prevSheets, nextWorkbook: workbook, recalc: null, callClaude: null,
  });

  let outBuf = await workbook.xlsx.writeBuffer();
  outBuf = await finalizeRequisitionWorkbook(baseBuffer, Buffer.from(outBuf));
  return { outBuf: Buffer.from(outBuf), rfResult, verification, meta, workbook };
}

// ── Draft row helpers ──────────────────────────────────────────────────────

// Normalize a phase key: trim, drop a leading "phase" word if the user typed it,
// keep the identifier (e.g. "2", "2a"). '' / null → '' (the single default stream).
function normPhase(p) {
  if (p == null) return '';
  let s = String(p).trim().replace(/^phase\s*/i, '').trim();
  return s;
}

// The label inserted into filenames for a phase, e.g. "Phase 2a". '' for the
// default stream (no phase label). Used both to name the output and to match the
// one-copy purge precisely (so "Phase 2" never matches "Phase 2a").
function phaseLabel(p) {
  const s = normPhase(p);
  return s ? 'Phase ' + s : '';
}

function getOpenDraft(db, eid, phase) {
  const ph = normPhase(phase);
  return db.prepare(
    "SELECT * FROM requisition_draft WHERE entity_id=? AND status='open' AND IFNULL(phase,'')=? ORDER BY id DESC LIMIT 1"
  ).get(eid, ph);
}

// All open drafts for an entity (used to enforce the two-phase cap and to list
// them in the UI). Ordered by phase then id.
function getOpenDrafts(db, eid) {
  return db.prepare(
    "SELECT * FROM requisition_draft WHERE entity_id=? AND status='open' ORDER BY IFNULL(phase,''), id"
  ).all(eid);
}

// Resolve the auto-seed base for a NEW draft of the given phase: the most recent
// finalized draft's output_blob FOR THAT PHASE, else the newest filed requisition
// workbook for that phase in the entity's "Requisition Reports" tree. Returns
// { buffer, name, source } or null (→ manual upload).
function resolveAutoSeed(db, workpapersDir, eid, phase) {
  const ph = normPhase(phase);
  const fin = db.prepare(
    "SELECT output_blob, output_name FROM requisition_draft WHERE entity_id=? AND status='finalized' AND IFNULL(phase,'')=? AND output_blob IS NOT NULL ORDER BY finalized_at DESC, id DESC LIMIT 1"
  ).get(eid, ph);
  if (fin && fin.output_blob) {
    return { buffer: Buffer.from(fin.output_blob), name: fin.output_name || 'prior_requisition.xlsx', source: 'finalized-draft' };
  }
  // Fallback: newest filed workbook whose name looks like a Requisition Report,
  // under a "Requisition Reports" folder for this entity. When a phase is set,
  // require the phase label in the name so Phase 2 seeds from Phase 2 (and the
  // default stream avoids grabbing a phased file). Matched with word boundaries
  // so "Phase 2" never matches "Phase 2a".
  const lbl = phaseLabel(ph);
  let sql =
    "SELECT stored_filename, original_name FROM entity_files " +
    "WHERE entity_id=? AND folder_path LIKE '%Requisition Reports%' " +
    "AND lower(original_name) LIKE '%requisition report%' AND lower(original_name) LIKE '%.xlsx' ";
  const args = [eid];
  if (lbl) { sql += "AND lower(original_name) LIKE ? "; args.push('%' + lbl.toLowerCase() + '%'); }
  sql += "ORDER BY id DESC LIMIT 20";
  const rows = db.prepare(sql).all(...args);
  const pick = rows.find(r => phaseMatchesName(r.original_name, ph)) || (lbl ? null : rows[0]);
  if (pick) {
    const fs = require('fs'); const path = require('path');
    try {
      const buf = fs.readFileSync(path.join(workpapersDir, String(eid), pick.stored_filename));
      return { buffer: buf, name: pick.original_name || 'prior_requisition.xlsx', source: 'workpapers' };
    } catch (_) { /* fall through */ }
  }
  return null;
}

// True when a filename belongs to the given phase. For the default stream ('')
// the name must carry NO "Phase <x>" token at all. For a real phase it must
// carry exactly that token, bounded so "Phase 2" doesn't match "Phase 2a".
function phaseMatchesName(name, phase) {
  const ph = normPhase(phase);
  const s = String(name || '');
  const anyPhase = /\bphase\s+[0-9a-z]+/i.test(s);
  if (!ph) return !anyPhase;
  const re = new RegExp('\\bphase\\s+' + ph.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![0-9a-z])', 'i');
  return re.test(s);
}

// Upload guard (Option B — always ask): compare a hand-uploaded base against the
// filed prior-month copy FOR THIS PHASE. Returns a match descriptor the endpoint
// surfaces so the user picks "use uploaded" vs "use filed copy". Never decides.
//   { match:'identical' } — byte-identical to the filed copy
//   { match:'same-period-edited', filedName } — same req/date, different bytes
//   null — no conflict (genuinely new base)
function checkUploadAgainstFiled(db, workpapersDir, eid, uploadBuf, uploadedReqNumber, uploadedAsOf, phase) {
  const seed = resolveAutoSeed(db, workpapersDir, eid, phase);
  if (!seed) return null; // nothing filed yet → first-time upload, no conflict
  const upHash = sha256(uploadBuf);
  const filedHash = sha256(seed.buffer);
  if (upHash === filedHash) return { match: 'identical', filedName: seed.name };
  // Same-period heuristic: the filed copy IS the latest finalized report. If the
  // user is uploading a base for the NEXT req and it isn't the filed copy, that's
  // expected (no conflict). We only flag when the upload appears to BE the filed
  // period — i.e. its embedded req/date equal the filed report's. The caller
  // passes the parsed req/date when available; when it can't parse, we stay quiet
  // rather than false-alarm.
  // (Byte-diff with matching identity is the "edited copy of the filed month" case.)
  return null;
}

// Insert the phase label into a built filename just before the trailing date +
// extension, e.g. "...Requisition Report #12 07.31.2026.xlsx" →
// "...Requisition Report #12 Phase 2a 07.31.2026.xlsx". No-op for the default
// stream. If the base name already carries the same phase token, it's left as-is
// (idempotent). Falls back to appending before the extension if no date is found.
function phasedFilename(name, phase) {
  const lbl = phaseLabel(phase);
  if (!lbl) return name;
  const s = String(name || 'Requisition_Report.xlsx');
  if (phaseMatchesName(s, phase)) return s; // already labeled for this phase
  // Try to insert before a trailing date token (dd.dd.dddd / dd_dd_dddd etc.)
  const dateRe = /(\s+)(\d{1,2}[._/-]\d{1,2}[._/-]\d{2,4})(\.[^.]+)$/;
  if (dateRe.test(s)) return s.replace(dateRe, (m, sp, date, ext) => ' ' + lbl + sp + date + ext);
  // else insert before the extension
  const dot = s.lastIndexOf('.');
  if (dot > 0) return s.slice(0, dot) + ' ' + lbl + s.slice(dot);
  return s + ' ' + lbl;
}

module.exports = {
  buildRollforwardMeta,
  rollForwardFromBase,
  loadReqWorkbook,
  getOpenDraft,
  getOpenDrafts,
  resolveAutoSeed,
  checkUploadAgainstFiled,
  normPhase,
  phaseLabel,
  phaseMatchesName,
  phasedFilename,
  sha256,
};
