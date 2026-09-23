// ─────────────────────────────────────────────────────────────────────────────
// xlsx_sanitize.js — one place that makes every workbook CloudLedger writes open
// cleanly in Excel the first time (no "We found a problem with some content …
// recover?" prompt).
//
// That prompt is Excel rejecting small structural defects in the file. The three
// we have seen, all inherited from fund-administrator (Weaver) source workbooks or
// produced by the spreadsheet library:
//   1. <sheetPr> child order — <pageSetUpPr> written before <outlinePr>. OOXML
//      requires outlinePr first; ExcelJS 4.x emits the wrong order when a sheet has
//      both fit-to-page and outline/grouping settings.
//   2. Orphaned external links — xl/externalLinks/* parts and <externalReferences>
//      declared in the workbook but referenced by no formula (dead pointers to
//      other people's files). Excel flags these on open.
//   3. Broken / junk defined names — thousands of legacy named ranges (#REF!,
//      references to sheets that no longer exist, or unused model/FactSet junk).
//
// `sanitizeXlsxBuffer` fixes all three on a finished .xlsx buffer, losslessly for
// everything else. `install(ExcelJS)` hooks it onto XLSX.prototype.writeBuffer so
// EVERY generator (current and future) is covered automatically — no per-report
// patching. Fail-safe: any error returns the original buffer unchanged.
// ─────────────────────────────────────────────────────────────────────────────
const JSZip = require('jszip');

async function sanitizeXlsxBuffer(input) {
  try {
    const buf = Buffer.isBuffer(input) ? input : Buffer.from(input);
    const zip = await JSZip.loadAsync(buf);
    let changed = false;

    // 1. Fix <sheetPr> child order on every worksheet.
    for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      const xml = await zip.file(name).async('string');
      const fixed = xml.replace(/(<pageSetUpPr\b[^>]*\/>)\s*(<outlinePr\b[^>]*\/>)/g, '$2$1');
      if (fixed !== xml) { zip.file(name, fixed); changed = true; }
    }

    // Do any formulas still reference an external workbook ([1], [2], …)?
    let usesExternal = false;
    for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      if (/\[\d+\]/.test(await zip.file(name).async('string'))) { usesExternal = true; break; }
    }

    const wbPath = 'xl/workbook.xml';
    const relPath = 'xl/_rels/workbook.xml.rels';
    const ctPath = '[Content_Types].xml';
    let wb = zip.file(wbPath) ? await zip.file(wbPath).async('string') : null;

    // 2. Remove orphaned external links (only when nothing references them).
    if (wb && !usesExternal && Object.keys(zip.files).some((n) => n.startsWith('xl/externalLinks/'))) {
      for (const n of Object.keys(zip.files).filter((n) => n.startsWith('xl/externalLinks/'))) { zip.remove(n); changed = true; }
      wb = wb.replace(/<externalReferences>[\s\S]*?<\/externalReferences>/g, '');
      if (zip.file(relPath)) {
        const rels = (await zip.file(relPath).async('string'))
          .replace(/<Relationship\b[^>]*Target="[^"]*externalLinks\/[^"]*"[^>]*\/>/g, '');
        zip.file(relPath, rels);
      }
      if (zip.file(ctPath)) {
        const ct = (await zip.file(ctPath).async('string'))
          .replace(/<Override\b[^>]*PartName="\/xl\/externalLinks\/[^"]*"[^>]*\/>/g, '');
        zip.file(ctPath, ct);
      }
    }

    // 3. Strip broken / junk defined names (keep _xlnm.* built-ins on existing
    //    sheets and names a formula actually uses).
    if (wb) {
      const m = wb.match(/<definedNames>([\s\S]*?)<\/definedNames>/);
      if (m) {
        const sheets = new Set([...wb.matchAll(/<sheet [^>]*name="([^"]*)"/g)].map((x) => x[1]));
        let blob = '';
        for (const n of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
          blob += (await zip.file(n).async('string')).replace(/[\s\S]*?(<f\b)/g, '$1');
        }
        const all = m[1].match(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g) || [];
        const keep = [];
        for (const dn of all) {
          const nm = (dn.match(/name="([^"]*)"/) || [])[1] || '';
          if (dn.includes('#REF!')) continue;
          const val = dn.replace(/<[^>]+>/g, '');
          const refs = [...val.matchAll(/(?:^|[=,+\-*/(! ])'?([A-Za-z0-9 _.&\-]+?)'?!/g)].map((x) => x[1]);
          if (refs.some((r) => !sheets.has(r))) continue; // references a missing sheet
          const used = new RegExp('(?<![A-Za-z0-9_.])' + nm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![A-Za-z0-9_.])').test(blob);
          if (nm.startsWith('_xlnm.') || used) keep.push(dn);
        }
        if (keep.length !== all.length) {
          const block = keep.length ? '<definedNames>' + keep.join('') + '</definedNames>' : '';
          wb = wb.slice(0, m.index) + block + wb.slice(m.index + m[0].length);
          changed = true;
        }
      }
      zip.file(wbPath, wb);
    }

    if (!changed) return buf;
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) {
    return Buffer.isBuffer(input) ? input : Buffer.from(input);
  }
}

// Hook the sanitizer onto ExcelJS so every wb.xlsx.writeBuffer() is cleaned.
// Idempotent: installing more than once is a no-op.
function install(ExcelJS) {
  try {
    const XLSX = new ExcelJS.Workbook().xlsx.constructor;
    if (XLSX.prototype.__clSanitized) return;
    const orig = XLSX.prototype.writeBuffer;
    XLSX.prototype.writeBuffer = async function writeBuffer(...args) {
      const out = await orig.apply(this, args);
      return sanitizeXlsxBuffer(out);
    };
    XLSX.prototype.__clSanitized = true;
  } catch (e) { /* leave ExcelJS unpatched if the internal shape changed */ }
}

module.exports = { sanitizeXlsxBuffer, install };
