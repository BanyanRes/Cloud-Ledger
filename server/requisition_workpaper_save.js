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
const { PDFDocument, PDFName, PDFNumber, PDFString, PDFArray, PDFDict, PDFNull } = require('pdf-lib');

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
function saveBufferToWorkpapers(db, workpapersDir, eid, folderPath, originalName, mimeType, buffer, who, opts = {}) {
  const entityDir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(entityDir, { recursive: true });
  const storedName = crypto.randomBytes(16).toString('hex') + path.extname(originalName);
  fs.writeFileSync(path.join(entityDir, storedName), buffer);
  // overwrite: re-running the roll-forward for the same period should REPLACE the
  // prior output, not pile up timestamp-suffixed copies. Remove any existing rows
  // (and their stored blobs) of the same name in this folder, then keep the name
  // exactly as given. Without overwrite, fall back to a unique (timestamped) name.
  let finalName = originalName;
  if (opts.overwrite) {
    const dupes = db.prepare(
      'SELECT id, stored_filename FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name=?'
    ).all(eid, folderPath, originalName);
    for (const d of dupes) {
      try { fs.unlinkSync(path.join(entityDir, d.stored_filename)); } catch (_) {}
      db.prepare('DELETE FROM entity_files WHERE id=?').run(d.id);
    }
  } else {
    finalName = uniqueOriginalName(db, eid, folderPath, originalName);
  }
  const info = db.prepare(
    'INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)'
  ).run(eid, folderPath, storedName, finalName, buffer.length, mimeType, who || 'system');
  return { id: info.lastInsertRowid, original_name: finalName, size: buffer.length };
}

// Build a one-page Development Fee invoice PDF (the fee has no vendor invoice of
// its own, so we generate one for the packet). Shows the period, payee, the
// computed amount, and the basis. Returns a Buffer.
async function buildDevFeePdf({ amount, reqNumber, asOfDate, vendor, baseAmount, rateText }) {
  const { rgb, StandardFonts } = require('pdf-lib');
  const doc = await PDFDocument.create();
  const page = doc.addPage([612, 792]); // US Letter
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const money = (n) => {
    const v = Number(n); if (!Number.isFinite(v)) return '';
    const s = v < 0 ? '-' : '';
    return s + '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  let monthYear = '';
  if (asOfDate) { const d = new Date(asOfDate); if (!isNaN(d.getTime())) monthYear = MONTHS[d.getMonth()] + ' ' + d.getFullYear(); }
  const draw = (t, x, y, f = font, size = 12, color = rgb(0.1,0.1,0.1)) => page.drawText(String(t), { x, y, size, font: f, color });

  let y = 720;
  draw('Development Fee', 56, y, bold, 22, rgb(0.13,0.21,0.36)); y -= 14;
  page.drawLine({ start: { x: 56, y }, end: { x: 556, y }, thickness: 1.5, color: rgb(0.13,0.21,0.36) }); y -= 36;

  draw('Payee:', 56, y, bold, 12); draw(vendor || 'Banyan Residential', 160, y); y -= 22;
  if (reqNumber != null && reqNumber !== '') { draw('Requisition:', 56, y, bold, 12); draw('Req# ' + reqNumber, 160, y); y -= 22; }
  if (monthYear) { draw('Period:', 56, y, bold, 12); draw(monthYear, 160, y); y -= 22; }
  if (asOfDate) { draw('As-of date:', 56, y, bold, 12); draw(String(asOfDate), 160, y); y -= 22; }
  y -= 16;

  page.drawLine({ start: { x: 56, y }, end: { x: 556, y }, thickness: 0.75, color: rgb(0.7,0.7,0.7) }); y -= 28;
  if (baseAmount != null) { draw('New costs this period (basis):', 56, y, font, 12); draw(money(baseAmount), 400, y, font, 12); y -= 22; }
  if (rateText) { draw('Development fee rate:', 56, y, font, 12); draw(rateText, 400, y, font, 12); y -= 22; }
  y -= 6;
  page.drawLine({ start: { x: 56, y }, end: { x: 556, y }, thickness: 0.75, color: rgb(0.7,0.7,0.7) }); y -= 30;

  draw('Development Fee Due:', 56, y, bold, 14); draw(money(amount), 380, y, bold, 14, rgb(0.13,0.21,0.36)); y -= 40;

  draw('Auto-generated by CloudLedger from the requisition roll-forward.', 56, 64, font, 9, rgb(0.5,0.5,0.5));
  const bytes = await doc.save();
  return Buffer.from(bytes);
}

// Merge a list of stored invoices into one PDF, in the given order, with a
// top-level PDF bookmark (outline) at the first page of each invoice so the
// packet is navigable and its order mirrors the Current Invoice Log.
//
// Each invoice carries { file_blob (Buffer), mime_type, original_name } and
// optionally { vendor, bill_number, amount, cost_code, cost_code_name } used to
// label its bookmark. PDFs are appended page by page; images (png/jpg) are
// embedded one per page. Unsupported types are skipped. The caller's array
// order is preserved exactly. Returns a Buffer, or null if nothing usable.
async function buildInvoicePacket(invoices) {
  if (!invoices || !invoices.length) return null;
  const out = await PDFDocument.create();
  let added = 0;
  // Track where each invoice starts in the merged packet and its label, so we
  // can build the outline after all pages are placed.
  const marks = []; // { pageIndex, title }

  const money = (n) => {
    const v = Number(n);
    if (!Number.isFinite(v)) return '';
    const sign = v < 0 ? '-' : '';
    return ' ' + sign + '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  const labelFor = (inv) => {
    const gl = inv.cost_code != null && inv.cost_code !== '' ? String(inv.cost_code) + ' ' : '';
    const vendor = (inv.vendor || inv.cost_code_name || inv.original_name || 'Invoice').toString().trim();
    const bill = inv.bill_number ? ' Inv. ' + inv.bill_number : '';
    const amt = money(inv.amount);
    return (gl + vendor + bill + amt).trim() || 'Invoice';
  };

  for (const inv of invoices) {
    const buf = inv.file_blob;
    if (!buf || !buf.length) continue;
    const mime = (inv.mime_type || '').toLowerCase();
    const name = (inv.original_name || '').toLowerCase();
    const startIndex = out.getPageCount();
    let placed = false;
    try {
      if (mime.includes('pdf') || name.endsWith('.pdf')) {
        const src = await PDFDocument.load(buf, { ignoreEncryption: true });
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p => out.addPage(p));
        placed = pages.length > 0;
      } else if (mime.includes('png') || name.endsWith('.png')) {
        const img = await out.embedPng(buf);
        const page = out.addPage([img.width, img.height]);
        page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        placed = true;
      } else if (mime.includes('jpeg') || mime.includes('jpg') || name.endsWith('.jpg') || name.endsWith('.jpeg')) {
        const img = await out.embedJpg(buf);
        const page = out.addPage([img.width, img.height]);
        page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        placed = true;
      }
      // other types: skip silently
    } catch (e) {
      // a single bad file shouldn't sink the whole packet
      console.error('invoice packet: skipped a file:', name || inv.id, e.message);
    }
    if (placed) {
      marks.push({ pageIndex: startIndex, title: labelFor(inv) });
      added++;
    }
  }
  if (!added) return null;
  addOutline(out, marks);
  const bytes = await out.save();
  return Buffer.from(bytes);
}

// Attach a flat PDF outline (bookmarks) to `doc`. `marks` is an ordered list of
// { pageIndex, title }. pdf-lib has no high-level outline API, so we assemble
// the /Outlines dictionary and the linked list of outline items by hand and
// register them in the catalog. Each item is a top-level bookmark whose /Dest
// jumps to the top of its page (XYZ with null left/top = "keep position").
function addOutline(doc, marks) {
  if (!marks || !marks.length) return;
  const context = doc.context;
  const pages = doc.getPages();

  // Create the parent /Outlines dict first; children reference it as /Parent.
  const outlinesDict = context.obj({ Type: 'Outlines' });
  const outlinesRef = context.register(outlinesDict);

  // Pre-create a ref for each item so siblings can link /Next and /Prev.
  const itemRefs = marks.map(() => context.nextRef());

  marks.forEach((m, i) => {
    const page = pages[Math.min(m.pageIndex, pages.length - 1)];
    const dest = PDFArray.withContext(context);
    dest.push(page.ref);
    dest.push(PDFName.of('XYZ'));
    dest.push(PDFNull); // left  — null = keep current horizontal position
    dest.push(PDFNull); // top   — null = keep current vertical position
    dest.push(PDFNumber.of(0));    // zoom (0 = keep current zoom)

    const item = new Map();
    item.set(PDFName.of('Title'), PDFString.of(m.title));
    item.set(PDFName.of('Parent'), outlinesRef);
    item.set(PDFName.of('Dest'), dest);
    if (i > 0) item.set(PDFName.of('Prev'), itemRefs[i - 1]);
    if (i < marks.length - 1) item.set(PDFName.of('Next'), itemRefs[i + 1]);
    context.assign(itemRefs[i], PDFDict.fromMapWithContext(item, context));
  });

  outlinesDict.set(PDFName.of('First'), itemRefs[0]);
  outlinesDict.set(PDFName.of('Last'), itemRefs[itemRefs.length - 1]);
  outlinesDict.set(PDFName.of('Count'), PDFNumber.of(marks.length));

  doc.catalog.set(PDFName.of('Outlines'), outlinesRef);
}

// Main entry. Saves the workbook and (if any invoices) the merged packet into
// the period folder. Best-effort: never throws — returns a result summary so the
// caller can log it without risking the user's download.
async function saveRequisitionOutputs({ db, workpapersDir, eid, reqNumber, asOfDate, workbookBuffer, invoices, devFee, who, packetPrefix, workbookFilename }) {
  const result = { folder: null, workbook: null, packet: null, errors: [] };
  try {
    const folderPath = requisitionFolderPath(asOfDate);
    result.folder = folderPath;
    ensureFolders(db, eid, folderPath, who);

    const reqLabel = reqNumber != null && reqNumber !== '' ? `Req ${reqNumber}` : 'Requisition';
    // Prefer the caller-derived name (same as the download — e.g.
    // "0005 B1 County Line SRN Requisition Report #12 02.28.2026.xlsx") so the
    // saved workbook matches the prior report's naming convention. Only fall
    // back to the bare label if no name was supplied.
    const workbookName = (workbookFilename && String(workbookFilename).trim()) || `${reqLabel} Report.xlsx`;

    if (workbookBuffer && workbookBuffer.length) {
      try {
        result.workbook = saveBufferToWorkpapers(
          db, workpapersDir, eid, folderPath,
          workbookName,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          workbookBuffer, who, { overwrite: true }
        );
      } catch (e) { result.errors.push('workbook: ' + e.message); }
    }

    try {
      // The Development Fee has no vendor invoice of its own, so generate a
      // one-page Dev Fee invoice and append it to the packet as the last item —
      // matching its position at the end of the Current Invoice Log. It rides the
      // same ordering + bookmark path as the real invoices.
      const packetInvoices = Array.isArray(invoices) ? invoices.slice() : [];
      if (devFee && devFee.amount != null) {
        try {
          const df = devFee.row || {};
          const baseAmount = devFee.base != null ? devFee.base : null;
          const rateText = devFee.rateText || '2% of new costs';
          const pdfBuf = await buildDevFeePdf({
            amount: devFee.amount, reqNumber, asOfDate,
            vendor: df.vendor || 'Banyan Residential',
            baseAmount, rateText,
          });
          packetInvoices.push({
            file_blob: pdfBuf, mime_type: 'application/pdf',
            original_name: 'Development Fee.pdf',
            vendor: df.vendor || 'Banyan Residential',
            bill_number: df.bill || 'Dev Fee',
            amount: devFee.amount,
            cost_code: devFee.code || df.code,
            cost_code_name: df.name || 'Development Fee',
          });
        } catch (e) { result.errors.push('devfee pdf: ' + e.message); }
      }

      const packetBuf = await buildInvoicePacket(packetInvoices);
      if (packetBuf) {
        // Packet filename: "<entity id / name> Invoice Packet <Month Year>.pdf"
        // e.g. "0005 B1a County Line SRN Invoice Packet February 2026.pdf".
        // Falls back to the Req label prefix if no entity prefix was supplied.
        const { monthName } = periodFolderParts(asOfDate);
        const prefix = (packetPrefix && String(packetPrefix).trim()) || reqLabel;
        const packetName = `${prefix} Invoice Packet ${monthName}.pdf`;
        result.packet = saveBufferToWorkpapers(
          db, workpapersDir, eid, folderPath,
          packetName,
          'application/pdf',
          packetBuf, who, { overwrite: true }
        );
      }
    } catch (e) { result.errors.push('packet: ' + e.message); }
  } catch (e) {
    result.errors.push('save: ' + e.message);
  }
  return result;
}

module.exports = { saveRequisitionOutputs, requisitionFolderPath, buildInvoicePacket, buildDevFeePdf };
