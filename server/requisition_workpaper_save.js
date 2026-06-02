// Auto-save requisition outputs (the rolled-forward workbook + a merged invoice
// packet) into the entity's Workpapers tree under:
//
//   <YYYY> / Requisition Reports / <Month YYYY> /
//
// e.g. an As-of Date in May 2026 lands in "2026/Requisition Reports/May 2026".
//
// Both files are written to disk under WORKPAPERS_DIR/<eid>/ and registered in
// the entity_files / entity_folders tables exactly like a manual upload, so they
// show up in the Workpapers UI with no further wiring. Name collisions get a
// timestamp suffix rather than overwriting (audit-friendly).

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { PDFDocument } = require('pdf-lib');

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// Parse a date-ish string (YYYY-MM-DD or anything Date accepts) into {year, monthName}.
// Falls back to today if the input is missing or unparseable.
function periodFolderParts(asOfDate) {
  let d = asOfDate ? new Date(asOfDate) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  return { year: String(d.getFullYear()), monthName: MONTHS[d.getMonth()] + ' ' + d.getFullYear() };
}

// Build the "2026/Requisition Reports/May 2026" virtual folder path.
function requisitionFolderPath(asOfDate) {
  const { year, monthName } = periodFolderParts(asOfDate);
  return `${year}/Requisition Reports/${monthName}`;
}

// Register a folder and all its ancestors in entity_folders (idempotent).
function ensureFolders(db, eid, folderPath, who) {
  const parts = folderPath.split('/').filter(Boolean);
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)');
  let acc = '';
  for (const p of parts) {
    acc = acc ? acc + '/' + p : p;
    ins.run(eid, acc, who || 'system');
  }
}

// Pick a stored_filename that doesn't collide with an existing entity_files row
// in the same folder. If the desired original_name is taken, append a timestamp.
function uniqueOriginalName(db, eid, folderPath, desiredName) {
  const exists = db.prepare(
    'SELECT 1 FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name=? LIMIT 1'
  ).get(eid, folderPath, desiredName);
  if (!exists) return desiredName;
  const dot = desiredName.lastIndexOf('.');
  const base = dot > 0 ? desiredName.slice(0, dot) : desiredName;
  const ext = dot > 0 ? desiredName.slice(dot) : '';
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  return `${base} (${stamp})${ext}`;
}

// Write a buffer to WORKPAPERS_DIR/<eid>/<random> and insert an entity_files row.
function saveBufferToWorkpapers(db, workpapersDir, eid, folderPath, originalName, mimeType, buffer, who) {
  const entityDir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(entityDir, { recursive: true });
  const storedName = crypto.randomBytes(16).toString('hex') + path.extname(originalName);
  fs.writeFileSync(path.join(entityDir, storedName), buffer);
  const finalName = uniqueOriginalName(db, eid, folderPath, originalName);
  const info = db.prepare(
    'INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)'
  ).run(eid, folderPath, storedName, finalName, buffer.length, mimeType, who || 'system');
  return { id: info.lastInsertRowid, original_name: finalName, size: buffer.length };
}

// Merge a list of stored invoices into one PDF. Each invoice carries
// { file_blob (Buffer), mime_type, original_name }. PDFs are appended page by
// page; images (png/jpg) are embedded one per page. Unsupported types are
// skipped. Returns a Buffer, or null if nothing usable was produced.
async function buildInvoicePacket(invoices) {
  if (!invoices || !invoices.length) return null;
  const out = await PDFDocument.create();
  let added = 0;
  for (const inv of invoices) {
    const buf = inv.file_blob;
    if (!buf || !buf.length) continue;
    const mime = (inv.mime_type || '').toLowerCase();
    const name = (inv.original_name || '').toLowerCase();
    try {
      if (mime.includes('pdf') || name.endsWith('.pdf')) {
        const src = await PDFDocument.load(buf, { ignoreEncryption: true });
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p => out.addPage(p));
        added++;
      } else if (mime.includes('png') || name.endsWith('.png')) {
        const img = await out.embedPng(buf);
        const page = out.addPage([img.width, img.height]);
        page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        added++;
      } else if (mime.includes('jpeg') || mime.includes('jpg') || name.endsWith('.jpg') || name.endsWith('.jpeg')) {
        const img = await out.embedJpg(buf);
        const page = out.addPage([img.width, img.height]);
        page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        added++;
      }
      // other types: skip silently
    } catch (e) {
      // a single bad file shouldn't sink the whole packet
      console.error('invoice packet: skipped a file:', name || inv.id, e.message);
    }
  }
  if (!added) return null;
  const bytes = await out.save();
  return Buffer.from(bytes);
}

// Main entry. Saves the workbook and (if any invoices) the merged packet into
// the period folder. Best-effort: never throws — returns a result summary so the
// caller can log it without risking the user's download.
async function saveRequisitionOutputs({ db, workpapersDir, eid, reqNumber, asOfDate, workbookBuffer, invoices, who }) {
  const result = { folder: null, workbook: null, packet: null, errors: [] };
  try {
    const folderPath = requisitionFolderPath(asOfDate);
    result.folder = folderPath;
    ensureFolders(db, eid, folderPath, who);

    const reqLabel = reqNumber != null && reqNumber !== '' ? `Req ${reqNumber}` : 'Requisition';

    if (workbookBuffer && workbookBuffer.length) {
      try {
        result.workbook = saveBufferToWorkpapers(
          db, workpapersDir, eid, folderPath,
          `${reqLabel} Report.xlsx`,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          workbookBuffer, who
        );
      } catch (e) { result.errors.push('workbook: ' + e.message); }
    }

    try {
      const packetBuf = await buildInvoicePacket(invoices);
      if (packetBuf) {
        result.packet = saveBufferToWorkpapers(
          db, workpapersDir, eid, folderPath,
          `${reqLabel} Invoice Packet.pdf`,
          'application/pdf',
          packetBuf, who
        );
      }
    } catch (e) { result.errors.push('packet: ' + e.message); }
  } catch (e) {
    result.errors.push('save: ' + e.message);
  }
  return result;
}

module.exports = { saveRequisitionOutputs, requisitionFolderPath, buildInvoicePacket };
