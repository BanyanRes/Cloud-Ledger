// ─────────────────────────────────────────────────────────────────────────
// Finalize a roll-forward workbook after ExcelJS writes it.
//
// ExcelJS's write has two side effects we correct here:
//
// 1. It drops the calc chain and leaves STALE cached formula results (and it
//    does not set fullCalcOnLoad). So Excel opens showing prior-period numbers
//    everywhere a formula feeds off the invoice logs — the Dev Fee tab's column
//    C (Budget-to-Actual SUMIF of the Current Invoice Log), the group subtotals,
//    and the grand total. Setting <calcPr fullCalcOnLoad="1"> makes Excel
//    recompute the whole workbook from the rolled-forward data on open.
//
// 2. It drops external-link parts (xl/externalLinks/…) but leaves the '[n]'
//    formula references, so Excel shows a "we found a problem… recover?" repair
//    prompt. We re-inject the external links from the source so the package is
//    consistent again.
//
// No formulas or values are altered. Best-effort: any failure returns the
// ExcelJS output unchanged so a download is never blocked.
// ─────────────────────────────────────────────────────────────────────────
const JSZip = require('jszip');

async function finalizeRequisitionWorkbook(originalBuf, outBuf) {
  try {
    const out = await JSZip.loadAsync(outBuf);
    let changed = false;
    let wb = await out.file('xl/workbook.xml').async('string');

    // (1) Force a full recalculation when the workbook opens.
    if (/<calcPr\b[^>]*\/>/.test(wb)) {
      if (!/fullCalcOnLoad=/.test(wb)) {
        wb = wb.replace(/<calcPr\b([^\/>]*)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
        changed = true;
      }
    } else if (!/<calcPr\b/.test(wb)) {
      wb = wb.replace('</workbook>', '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>');
      changed = true;
    }

    // (2) Re-inject external links the write dropped (if the source had any).
    const src = await JSZip.loadAsync(originalBuf);
    const extNames = Object.keys(src.files).filter(n => n.startsWith('xl/externalLinks/') && !src.files[n].dir);
    const alreadyHas = Object.keys(out.files).some(n => n.startsWith('xl/externalLinks/') && !out.files[n].dir);
    if (extNames.length && !alreadyHas) {
      const links = extNames.filter(n => /externalLink\d+\.xml$/.test(n));
      for (const n of extNames) out.file(n, await src.file(n).async('nodebuffer'));

      let ct = await out.file('[Content_Types].xml').async('string');
      let ctAdd = '';
      for (const n of links) {
        const pn = '/' + n;
        if (!ct.includes(pn)) ctAdd += `<Override PartName="${pn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`;
      }
      if (ctAdd) out.file('[Content_Types].xml', ct.replace('</Types>', ctAdd + '</Types>'));

      let rels = await out.file('xl/_rels/workbook.xml.rels').async('string');
      let maxId = 0;
      for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxId = Math.max(maxId, Number(m[1]));
      const ids = [];
      let relAdd = '';
      for (const n of links) {
        const id = 'rId' + (++maxId);
        ids.push(id);
        relAdd += `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="${n.slice('xl/'.length)}"/>`;
      }
      out.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', relAdd + '</Relationships>'));

      if (!/<externalReferences/.test(wb)) {
        const er = '<externalReferences>' + ids.map(id => `<externalReference r:id="${id}"/>`).join('') + '</externalReferences>';
        if (/<definedNames>/.test(wb)) wb = wb.replace('<definedNames>', er + '<definedNames>');
        else if (/<\/sheets>/.test(wb)) wb = wb.replace('</sheets>', '</sheets>' + er);
      }
      changed = true;
    }

    if (changed) out.file('xl/workbook.xml', wb);
    return changed ? await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) : outBuf;
  } catch (e) {
    return outBuf; // never block a download over finalization
  }
}

// Back-compat alias (older name referred to the external-link step only).
module.exports = { finalizeRequisitionWorkbook, preserveExternalLinks: finalizeRequisitionWorkbook };
