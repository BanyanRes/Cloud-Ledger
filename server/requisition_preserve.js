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

    // ExcelJS mangles full-column print-area/title refs ($A:$G) into $ANaN:$GNaN
    // (NaN where a row number would go), which makes Excel show a repair prompt.
    // Strip the bogus NaN so the reference is valid again.
    if (/[A-Z]NaN/.test(wb)) { wb = wb.replace(/([A-Z])NaN/g, '$1'); changed = true; }

    // Strip invalid defined names carried over from source templates that were
    // built with data plugins (FactSet / S&P Capital IQ: EV__CVPARAMS__,
    // HTML_Control, IQR*). Their value is a bare A1 range with NO sheet qualifier
    // (e.g. "$C$17:$AAB$38"); a defined name must point to Sheet!range, so Excel
    // rejects them on open and shows the "repaired records: Named range from
    // /xl/workbook.xml" prompt. ExcelJS keeps a handful of them when it re-saves
    // (Phase-2 Silsbee books carry thousands). Remove any <definedName> whose
    // reference is a sheetless cell/range; constants and Sheet!-qualified names
    // (which contain '!') are left untouched.
    if (/<definedName\b/.test(wb)) {
      const before = wb;
      wb = wb.replace(/<definedName\b[^>]*>([\s\S]*?)<\/definedName>/g, (full, val) => {
        const v = val.trim();
        const sheetless = v.indexOf('!') === -1 &&
          /^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(v);
        return sheetless ? '' : full;
      });
      // Drop the container if stripping emptied it (empty <definedNames/> is itself
      // a schema violation Excel would flag).
      wb = wb.replace(/<definedNames>\s*<\/definedNames>/g, '');
      if (wb !== before) changed = true;
    }

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

    // (2b) Remove empty <conditionalFormatting> elements. ExcelJS sometimes drops
    //      the <cfRule> children it cannot represent but KEEPS the wrapper, leaving
    //      <conditionalFormatting sqref="..."/> with no rule. A conditionalFormatting
    //      must contain at least one cfRule, so this is a schema violation: the XML is
    //      well-formed (generic parsers accept it) but Excel rejects the worksheet part
    //      and shows "we found a problem... recover?". Strip the empty wrappers.
    for (const name of Object.keys(out.files)) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name) || out.files[name].dir) continue;
      const ws = await out.file(name).async('string');
      if (!/<conditionalFormatting\b/.test(ws)) continue;
      const cleaned = ws
        .replace(/<conditionalFormatting\b[^>]*\/>/g, '')
        .replace(/<conditionalFormatting\b[^>]*>\s*<\/conditionalFormatting>/g, '');
      if (cleaned !== ws) { out.file(name, cleaned); changed = true; }
    }

    // (3) Strip bare directory entries. A proper OOXML/OPC package (like the one
    //     Excel writes) contains only file parts, never folder entries. JSZip
    //     re-emits a folder entry for every directory when it generates the zip,
    //     and Excel treats those undeclared zero-length "parts" as corruption,
    //     showing the "we found a problem... recover?" repair prompt. Remove any
    //     folder objects so the regenerated package is clean. Done last, after all
    //     out.file() calls, so nothing re-creates them.
    let hadDirs = false;
    for (const n of Object.keys(out.files)) {
      if (out.files[n].dir) { delete out.files[n]; hadDirs = true; }
    }

    return (changed || hadDirs)
      ? await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
      : outBuf;
  } catch (e) {
    return outBuf; // never block a download over finalization
  }
}

// Back-compat alias (older name referred to the external-link step only).
module.exports = { finalizeRequisitionWorkbook, preserveExternalLinks: finalizeRequisitionWorkbook };
