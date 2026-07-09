// ─────────────────────────────────────────────────────────────────────────
// Preserve parts that ExcelJS silently drops on a load→write round-trip.
//
// ExcelJS does not model external links. When it rewrites a workbook that has
// an external reference (xl/externalLinks/externalLink*.xml, declared via
// <externalReferences> in workbook.xml), it drops the part AND the declaration
// — but leaves the formulas that use external-reference syntax ('[1]Sheet'!A1)
// intact. Excel then finds references to external index [1] with nothing behind
// them and shows the "We found a problem with some content… recover?" repair
// prompt on open.
//
// This restores the workbook to the same, consistent state as the source: it
// copies the external-link parts back in and re-declares them (Content_Types,
// workbook rels, and the <externalReferences> node), so the '[n]' formulas
// resolve again exactly as they did in the uploaded file. No formulas or values
// are altered. Best-effort: any failure returns the ExcelJS output unchanged so
// a download is never blocked.
// ─────────────────────────────────────────────────────────────────────────
const JSZip = require('jszip');

async function preserveExternalLinks(originalBuf, outBuf) {
  try {
    const src = await JSZip.loadAsync(originalBuf);
    const extNames = Object.keys(src.files).filter(n => n.startsWith('xl/externalLinks/') && !src.files[n].dir);
    if (!extNames.length) return outBuf; // source had no external links → nothing to do

    const out = await JSZip.loadAsync(outBuf);
    // If ExcelJS (or a prior pass) already kept them, don't double-add.
    if (Object.keys(out.files).some(n => n.startsWith('xl/externalLinks/') && !out.files[n].dir)) return outBuf;

    const links = extNames.filter(n => /externalLink\d+\.xml$/.test(n)); // the link parts (not their .rels)

    // 1) Copy every external-link part (link xml + its _rels, which point at the
    //    linked file via TargetMode="External").
    for (const n of extNames) out.file(n, await src.file(n).async('nodebuffer'));

    // 2) [Content_Types].xml — declare each external-link part.
    let ct = await out.file('[Content_Types].xml').async('string');
    let ctAdd = '';
    for (const n of links) {
      const pn = '/' + n;
      if (!ct.includes(pn)) ctAdd += `<Override PartName="${pn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`;
    }
    if (ctAdd) { ct = ct.replace('</Types>', ctAdd + '</Types>'); out.file('[Content_Types].xml', ct); }

    // 3) workbook.xml.rels — add a relationship (fresh rIds) for each link.
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
    rels = rels.replace('</Relationships>', relAdd + '</Relationships>');
    out.file('xl/_rels/workbook.xml.rels', rels);

    // 4) workbook.xml — re-add <externalReferences> (schema order: after
    //    </sheets>/functionGroups, before <definedNames>).
    let wb = await out.file('xl/workbook.xml').async('string');
    if (!/<externalReferences/.test(wb)) {
      const er = '<externalReferences>' + ids.map(id => `<externalReference r:id="${id}"/>`).join('') + '</externalReferences>';
      if (/<definedNames>/.test(wb)) wb = wb.replace('<definedNames>', er + '<definedNames>');
      else if (/<\/sheets>/.test(wb)) wb = wb.replace('</sheets>', '</sheets>' + er);
      out.file('xl/workbook.xml', wb);
    }

    return await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) {
    return outBuf; // never block a download over link preservation
  }
}

module.exports = { preserveExternalLinks };
