// Patch 7: TOC page references (pg2). Two-phase package assembly:
//  1) renderStatementsPdf records each statement's starting page (offsets).
//  2) generatePackage builds the body first (exec summary, statements, req),
//     tracking each section's absolute start page, then renders cover+TOC with
//     real page numbers and prepends it.
//  renderCoverPdf(meta, tocEntries) draws a right-aligned page number + dotted
//  leader for each TOC line.
const fs = require('fs');
const P = 'C:/Users/JimmyYun/Cloud-Ledger/server/financials.js';
let src = fs.readFileSync(P, 'utf8');
const EOL = src.includes('\r\n') ? '\r\n' : '\n';
const E = s => s.replace(/\n/g, EOL);
let applied = 0;
function replace(label, oldStr, newStr) {
  const o = E(oldStr), n = E(newStr);
  const count = src.split(o).length - 1;
  if (count === 0) throw new Error('ANCHOR NOT FOUND: ' + label + '\n---\n' + JSON.stringify(o.slice(0, 260)));
  if (count > 1) throw new Error('ANCHOR NOT UNIQUE (' + count + '): ' + label);
  src = src.replace(o, () => n); applied++; console.log('ok:', label);
}

// ── 7.1 Instrument renderStatementsPdf to record per-statement start pages. ───
replace('renderStatementsPdf signature + offsets init',
`// Render the four statements into a fresh PDFDocument and return its bytes.
async function renderStatementsPdf(s) {
  const pdf = await PDFDocument.create();`,
`// Render the four statements into a fresh PDFDocument and return its bytes.
// If an \`outOffsets\` array is passed, it is filled with { label, page } entries
// giving each statement's 0-based starting page index within this PDF (used to
// compute Table-of-Contents page references).
async function renderStatementsPdf(s, outOffsets) {
  const track = (label) => { if (outOffsets) outOffsets.push({ label, page: pdf.getPageCount() }); };
  const pdf = await PDFDocument.create();`);

// Record the start page immediately before each statement's L.start().
replace('BS start track',
`    const L = makeLayout(pdf, fonts, m, 'Balance Sheets', { dateLine: m.longDate + ' and ' + m.priorLongDate });
    L.start();`,
`    const L = makeLayout(pdf, fonts, m, 'Balance Sheets', { dateLine: m.longDate + ' and ' + m.priorLongDate });
    track('Balance Sheets');
    L.start();`);

replace('Operations start track',
`    const L = makeLayout(pdf, fonts, m, 'Statements of Operations', { dateLine: 'For the Months Ended ' + m.longDate + ' and ' + m.priorLongDate });
    L.start();`,
`    const L = makeLayout(pdf, fonts, m, 'Statements of Operations', { dateLine: 'For the Months Ended ' + m.longDate + ' and ' + m.priorLongDate });
    track('Statements of Operations');
    L.start();`);

replace('Cash Flow start track',
`    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows', { dateLine: m.monthsEnded });
    L.start();`,
`    const L = makeLayout(pdf, fonts, m, 'Statement of Cash Flows', { dateLine: m.monthsEnded });
    track('Statement of Cash Flows');
    L.start();`);

replace('Members Equity start track',
`    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\\u2019 Equity',
      { landscape: true, dateLine: m.monthsEnded });
    const LRIGHT = PAGE.h - PAGE.mR; // landscape printable right edge (PAGE.h is the long side)
    L.start();`,
`    const L = makeLayout(pdf, fonts, m, 'Statement of Changes in Members\\u2019 Equity',
      { landscape: true, dateLine: m.monthsEnded });
    const LRIGHT = PAGE.h - PAGE.mR; // landscape printable right edge (PAGE.h is the long side)
    track('Statement of Changes in Members\\u2019 Equity');
    L.start();`);

// ── 7.2 Rewrite renderCoverPdf to accept tocEntries [{label, page}] and draw a
//        dotted leader + right-aligned page number for each. ──────────────────
replace('renderCoverPdf body',
`async function renderCoverPdf(meta) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const page = pdf.addPage([PAGE.w, PAGE.h]);
  const center = (str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  // ── Cover page ────────────────────────────────────────────────────────────
  // Clean, professional cover: entity name, "Financial Statements", the as-of
  // date, framed by two thin rules. No table of contents here (it lives on its
  // own page that follows).
  center(meta.entityName, 22, bold, 512);
  page.drawLine({ start: { x: 150, y: 494 }, end: { x: PAGE.w - 150, y: 494 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Financial Statements', 15, reg, 470);
  center(meta.longDate, 12, reg, 448);
  page.drawLine({ start: { x: 150, y: 430 }, end: { x: PAGE.w - 150, y: 430 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });

  // ── Table of Contents page (separate) ────────────────────────────────────
  const toc2 = pdf.addPage([PAGE.w, PAGE.h]);
  const centerOn = (pg, str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    pg.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  centerOn(toc2, meta.entityName, 13, bold, PAGE.h - PAGE.mT + 10);
  centerOn(toc2, 'Table of Contents', 15, bold, PAGE.h - 150);
  toc2.drawLine({ start: { x: 180, y: PAGE.h - 168 }, end: { x: PAGE.w - 180, y: PAGE.h - 168 }, thickness: 0.6, color: rgb(0.3, 0.3, 0.3) });
  const toc = ['Executive Summary', 'Balance Sheets', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\\u2019 Equity', 'Budget to Actual'];
  let ty = PAGE.h - 210;
  toc.forEach(t => { centerOn(toc2, t, 11, reg, ty); ty -= 26; });
  return await pdf.save();
}`,
`async function renderCoverPdf(meta, tocEntries) {
  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const page = pdf.addPage([PAGE.w, PAGE.h]);
  const center = (str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  // ── Cover page ────────────────────────────────────────────────────────────
  center(meta.entityName, 22, bold, 512);
  page.drawLine({ start: { x: 150, y: 494 }, end: { x: PAGE.w - 150, y: 494 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });
  center('Financial Statements', 15, reg, 470);
  center(meta.longDate, 12, reg, 448);
  page.drawLine({ start: { x: 150, y: 430 }, end: { x: PAGE.w - 150, y: 430 }, thickness: 0.8, color: rgb(0.3, 0.3, 0.3) });

  // ── Table of Contents page (separate) with page references ────────────────
  const toc2 = pdf.addPage([PAGE.w, PAGE.h]);
  const centerOn = (pg, str, size, font, yy, color) => {
    const w = font.widthOfTextAtSize(str, size);
    pg.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font, color: color || rgb(0.1, 0.1, 0.1) });
  };
  centerOn(toc2, meta.entityName, 13, bold, PAGE.h - PAGE.mT + 10);
  centerOn(toc2, 'Table of Contents', 15, bold, PAGE.h - 150);
  toc2.drawLine({ start: { x: 180, y: PAGE.h - 168 }, end: { x: PAGE.w - 180, y: PAGE.h - 168 }, thickness: 0.6, color: rgb(0.3, 0.3, 0.3) });
  // Fall back to a label-only list if no page references were supplied.
  const entries = (tocEntries && tocEntries.length)
    ? tocEntries
    : ['Executive Summary', 'Balance Sheets', 'Statements of Operations', 'Statement of Cash Flows', 'Statement of Changes in Members\\u2019 Equity', 'Budget to Actual'].map(label => ({ label, page: null }));
  const LX = 120, RX = PAGE.w - 120;
  let ty = PAGE.h - 210;
  const sz = 11;
  for (const e of entries) {
    const label = e.label;
    toc2.drawText(label, { x: LX, y: ty, size: sz, font: reg, color: rgb(0.1, 0.1, 0.1) });
    if (e.page != null) {
      const num = String(e.page);
      const numW = reg.widthOfTextAtSize(num, sz);
      toc2.drawText(num, { x: RX - numW, y: ty, size: sz, font: reg, color: rgb(0.1, 0.1, 0.1) });
      // Dotted leader between label and page number.
      const labW = reg.widthOfTextAtSize(label, sz);
      const dotStart = LX + labW + 6, dotEnd = RX - numW - 6;
      const dotY = ty + 2;
      for (let dx = dotStart; dx < dotEnd; dx += 4) {
        toc2.drawText('.', { x: dx, y: dotY - 2, size: sz, font: reg, color: rgb(0.5, 0.5, 0.5) });
      }
    }
    ty -= 26;
  }
  return await pdf.save();
}`);

// ── 7.3 Rewrite generatePackage assembly to two-phase (body first, then cover
//        + TOC with page refs prepended). ──────────────────────────────────────
replace('generatePackage assembly',
`  // 1. Cover
  await appendPdf(await renderCoverPdf(statements.meta), 'Cover');
  // 2. Executive summary (uploaded, merged as-is)
  if (execSummaryBytes) await appendPdf(execSummaryBytes, 'Executive Summary');
  else info.warnings.push('No executive summary uploaded.');
  // 3. GL statements
  await appendPdf(await renderStatementsPdf(statements), 'Financial Statements');`,
`  // ── Two-phase assembly so the Table of Contents can show real page numbers. ─
  // Phase 1: build the BODY (everything after cover+TOC) into a separate doc,
  // recording the absolute start page of each TOC section. The cover + TOC are
  // two pages, so body page N (0-based within the body) is printed page N+3.
  const COVER_TOC_PAGES = 2;
  const body = await PDFDocument.create();
  const tocEntries = [];
  const appendToBody = async (bytes, label, addToc) => {
    if (!bytes) return 0;
    let srcDoc;
    try { srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true }); }
    catch (e) { info.warnings.push('Could not read ' + label + ' PDF: ' + e.message); return 0; }
    const startPage = body.getPageCount();
    const idx = srcDoc.getPageIndices();
    const pages = await body.copyPages(srcDoc, idx);
    pages.forEach(p => body.addPage(p));
    info.sections.push({ label, pages: pages.length });
    if (addToc) tocEntries.push({ label, page: startPage + COVER_TOC_PAGES + 1 });
    return pages.length;
  };

  // Executive summary (uploaded, merged as-is).
  if (execSummaryBytes) await appendToBody(execSummaryBytes, 'Executive Summary', true);
  else { info.warnings.push('No executive summary uploaded.'); }

  // GL statements — capture each statement's start page within the statements
  // PDF, then offset by where the statements PDF lands in the body.
  const stmtOffsets = [];
  const stmtBytes = await renderStatementsPdf(statements, stmtOffsets);
  const stmtBodyStart = body.getPageCount();
  await appendToBody(stmtBytes, 'Financial Statements', false);
  for (const off of stmtOffsets) {
    tocEntries.push({ label: off.label, page: stmtBodyStart + off.page + COVER_TOC_PAGES + 1 });
  }`);

// The requisition append must go into the body (not merged) and add a TOC entry.
replace('req xlsx append to body',
`      const kept = await appendPdf(reqPdfBytes, 'Requisition Report');
      info.reqRemoved = [];
      info.reqKept = kept;
      info.reqTotal = kept;`,
`      const kept = await appendToBody(reqPdfBytes, 'Budget to Actual', true);
      info.reqRemoved = [];
      info.reqKept = kept;
      info.reqTotal = kept;`);

replace('req pdf append to body',
`      if (!stripped.textDetected) info.warnings.push(stripped.parseFailed
        ? 'Requisition PDF could not be parsed for invoice-log detection; all pages were kept.'
        : 'Requisition PDF had no extractable text; invoice-log pages could not be detected and were left in.');
      await appendPdf(stripped.bytes, 'Requisition Report');`,
`      if (!stripped.textDetected) info.warnings.push(stripped.parseFailed
        ? 'Requisition PDF could not be parsed for invoice-log detection; all pages were kept.'
        : 'Requisition PDF had no extractable text; invoice-log pages could not be detected and were left in.');
      await appendToBody(stripped.bytes, 'Budget to Actual', true);`);

// Phase 2: render cover+TOC with the collected refs, then merge cover + body.
replace('generatePackage finalize',
`  info.cashFlowTies = statements.checks.cashFlowTies;
  info.cashFlowDiff = statements.checks.cashFlowDiff;
  info.balanceSheetTies = statements.checks.balanceSheetTies;
  info.pages = merged.getPageCount();
  const bytes = await merged.save();
  return { bytes, info };`,
`  // Phase 2: render cover + TOC (with page references) and assemble the final
  // PDF as cover/TOC first, then the body.
  const coverBytes = await renderCoverPdf(statements.meta, tocEntries);
  const coverDoc = await PDFDocument.load(coverBytes, { ignoreEncryption: true });
  const coverPages = await merged.copyPages(coverDoc, coverDoc.getPageIndices());
  coverPages.forEach(p => merged.addPage(p));
  const bodyPages = await merged.copyPages(body, body.getPageIndices());
  bodyPages.forEach(p => merged.addPage(p));

  info.cashFlowTies = statements.checks.cashFlowTies;
  info.cashFlowDiff = statements.checks.cashFlowDiff;
  info.balanceSheetTies = statements.checks.balanceSheetTies;
  info.tocEntries = tocEntries;
  info.pages = merged.getPageCount();
  const bytes = await merged.save();
  return { bytes, info };`);

fs.writeFileSync(P, src, 'utf8');
console.log('\nPATCH7 APPLIED', applied, 'edits.');
