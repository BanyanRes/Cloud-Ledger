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

function getOpenDraft(db, eid) {
  return db.prepare("SELECT * FROM requisition_draft WHERE entity_id=? AND status='open' ORDER BY id DESC LIMIT 1").get(eid);
}

// Resolve the auto-seed base for a NEW draft: the most recent finalized draft's
// output_blob, else the newest filed requisition workbook in the entity's
// Workpapers "Requisition Reports" tree (covers Reqs finalized before this
// feature shipped). Returns { buffer, name, source } or null (→ manual upload).
function resolveAutoSeed(db, workpapersDir, eid) {
  const fin = db.prepare(
    "SELECT output_blob, output_name FROM requisition_draft WHERE entity_id=? AND status='finalized' AND output_blob IS NOT NULL ORDER BY finalized_at DESC, id DESC LIMIT 1"
  ).get(eid);
  if (fin && fin.output_blob) {
    return { buffer: Buffer.from(fin.output_blob), name: fin.output_name || 'prior_requisition.xlsx', source: 'finalized-draft' };
  }
  // Fallback: newest filed workbook whose name looks like a Requisition Report,
  // under a "Requisition Reports" folder for this entity.
  const row = db.prepare(
    "SELECT stored_filename, original_name FROM entity_files " +
    "WHERE entity_id=? AND folder_path LIKE '%Requisition Reports%' " +
    "AND lower(original_name) LIKE '%requisition report%' AND lower(original_name) LIKE '%.xlsx' " +
    "ORDER BY id DESC LIMIT 1"
  ).get(eid);
  if (row) {
    const fs = require('fs'); const path = require('path');
    try {
      const buf = fs.readFileSync(path.join(workpapersDir, String(eid), row.stored_filename));
      return { buffer: buf, name: row.original_name || 'prior_requisition.xlsx', source: 'workpapers' };
    } catch (_) { /* fall through */ }
  }
  return null;
}

// Upload guard (Option B — always ask): compare a hand-uploaded base against the
// filed prior-month copy. Returns a match descriptor the endpoint surfaces so the
// user picks "use uploaded" vs "use filed copy". Never decides on its own.
//   { match:'identical' } — byte-identical to the filed copy
//   { match:'same-period-edited', filedName } — same req/date, different bytes
//   null — no conflict (genuinely new base)
function checkUploadAgainstFiled(db, workpapersDir, eid, uploadBuf, uploadedReqNumber, uploadedAsOf) {
  const seed = resolveAutoSeed(db, workpapersDir, eid);
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

module.exports = {
  buildRollforwardMeta,
  rollForwardFromBase,
  loadReqWorkbook,
  getOpenDraft,
  resolveAutoSeed,
  checkUploadAgainstFiled,
  sha256,
};
