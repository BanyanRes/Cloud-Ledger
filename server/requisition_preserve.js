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

// Map each worksheet's DISPLAY NAME to its part path (xl/worksheets/sheetN.xml)
// by resolving workbook.xml's <sheet name r:id> through workbook.xml.rels. The
// physical part-file numbering is NOT stable across writers: ExcelJS renumbers
// the sheetN.xml files when it re-saves, so the same logical sheet can be
// sheet22.xml in the source and a different sheetN.xml in the output. Copying
// per-sheet content (e.g. conditional formatting) by identical FILENAME therefore
// lands it on the wrong worksheet. Callers pair sheets by name via this map.
async function sheetPartMap(zip) {
  const map = new Map();
  try {
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rid2t = {};
    for (const m of relsXml.match(/<Relationship\b[^>]*>/g) || []) {
      const id = (m.match(/Id="([^"]+)"/) || [])[1];
      const tgt = (m.match(/Target="([^"]+)"/) || [])[1];
      if (id && tgt) rid2t[id] = tgt;
    }
    const unesc = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d));
    for (const m of wbXml.match(/<sheet\b[^>]*?\/?>/g) || []) {
      const nm = (m.match(/name="([^"]*)"/) || [])[1];
      const rid = (m.match(/r:id="([^"]+)"/) || [])[1];
      if (nm == null || !rid || !rid2t[rid]) continue;
      let t = rid2t[rid].replace(/^\//, '');
      if (!t.startsWith('xl/')) t = 'xl/' + t;
      map.set(unesc(nm), t);
    }
  } catch (_) { /* leave empty; caller falls back to no CF restore */ }
  return map;
}

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
        // A workbook-global _FilterDatabase (no localSheetId) is invalid — it must
        // be sheet-scoped. exceljs sometimes drops the localSheetId when it
        // round-trips a sheet's autofilter, and Excel then strips the name on open
        // ("Removed Records: Named range from /xl/workbook.xml"). Drop it here.
        if (/name="_xlnm\._FilterDatabase"/.test(full) && !/\blocalSheetId=/.test(full)) return '';
        const v = val.trim();
        // Malformed quoted sheet reference: exceljs does NOT escape an apostrophe
        // inside a sheet name, so "Members' Capital" is written 'Members' Capital'!…
        // instead of the correct 'Members'' Capital'!… . Excel can't resolve it and
        // strips the name ("Removed Records: Named range"). Detect a leading quoted
        // sheet name whose closing quote isn't immediately followed by '!' (after
        // un-doubling escaped quotes) and drop the whole defined name — these are
        // Print_Area/Print_Titles print cosmetics.
        {
          const s = v.replace(/&apos;/g, "'").replace(/^[+\-]/, '');
          if (s[0] === "'") {
            let i = 1;
            while (i < s.length) { if (s[i] === "'") { if (s[i + 1] === "'") { i += 2; continue; } break; } i++; }
            if (s[i] !== "'" || s[i + 1] !== '!') return '';
          }
        }
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

    // (CF) Restore each worksheet's original conditional formatting. ExcelJS does
    // not reliably round-trip <conditionalFormatting> on sheets it modifies (it can
    // emit a malformed, typeless <cfRule priority="1"/> with no rule body), which
    // makes Excel show "we found a problem" and strip the formatting on open. We
    // copy the exact original CF blocks from the source sheet back into the output
    // sheet, PAIRED BY WORKSHEET NAME (not part filename): ExcelJS renumbers the
    // sheetN.xml parts, so matching by filename dropped the Budget-to-Actual's
    // over-budget "Balance remaining < 0 -> red" rules and scattered other sheets'
    // rules onto the wrong tabs (Max, HP, 2026-09-10). Best-effort and non-fatal.
    try {
      const srcMap = await sheetPartMap(src);
      const outMap = await sheetPartMap(out);
      for (const [sheetName, srcPart] of srcMap) {
        const outPart = outMap.get(sheetName);
        if (!outPart || !src.files[srcPart] || !out.files[outPart]) continue;
        const srcXml = await src.file(srcPart).async('string');
        const origCF = (srcXml.match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g) || []).join('');
        let outXml = await out.file(outPart).async('string');
        const outHasCF = /<conditionalFormatting\b/.test(outXml);
        if (!origCF) {
          if (outHasCF) { outXml = outXml.replace(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, ''); out.file(outPart, outXml); changed = true; }
          continue;
        }
        if (outHasCF) {
          let replaced = false;
          outXml = outXml.replace(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, () => { if (replaced) return ''; replaced = true; return origCF; });
        } else if (/<pageMargins\b/.test(outXml)) {
          outXml = outXml.replace(/<pageMargins\b/, origCF + '<pageMargins');
        } else {
          outXml = outXml.replace('</worksheet>', origCF + '</worksheet>');
        }
        out.file(outPart, outXml); changed = true;
      }
      const srcStyles = await src.file('xl/styles.xml').async('string');
      const srcDxfs = (srcStyles.match(/<dxfs\b[\s\S]*?<\/dxfs>/) || srcStyles.match(/<dxfs\b[^>]*\/>/) || [null])[0];
      if (srcDxfs && out.files['xl/styles.xml']) {
        let outStyles = await out.file('xl/styles.xml').async('string');
        if (/<dxfs\b/.test(outStyles)) {
          const rep2 = outStyles.replace(/<dxfs\b[\s\S]*?<\/dxfs>|<dxfs\b[^>]*\/>/, srcDxfs);
          if (rep2 !== outStyles) { out.file('xl/styles.xml', rep2); changed = true; }
        } else if (/<\/cellStyles>/.test(outStyles)) {
          out.file('xl/styles.xml', outStyles.replace('</cellStyles>', '</cellStyles>' + srcDxfs)); changed = true;
        } else {
          out.file('xl/styles.xml', outStyles.replace('</styleSheet>', srcDxfs + '</styleSheet>')); changed = true;
        }
      }
    } catch (_) { /* non-fatal: leave ExcelJS CF as-is */ }
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
      let cleaned = ws
        .replace(/<conditionalFormatting\b[^>]*\/>/g, '')
        .replace(/<conditionalFormatting\b[^>]*>\s*<\/conditionalFormatting>/g, '');
      // Drop invalid cached formula results of NaN. exceljs emits <v>NaN</v> when
      // it re-writes a formula whose result it computed as NaN (e.g. references to
      // #REF! or date math over an empty cell — present on Braker's hidden
      // consolidation tabs). "NaN" is not a valid numeric cached value, so Excel
      // flags the worksheet part and shows a repair prompt. Removing the stale
      // cached value lets Excel recompute the formula on open (fullCalcOnLoad set).
      cleaned = cleaned.replace(/<v>NaN<\/v>/g, '');
      if (cleaned !== ws) { out.file(name, cleaned); changed = true; }
    }

    // (2c) Remove WMF pictures from drawings. exceljs re-emits any drawing that
    //      embeds a WMF image (a vector format it doesn't render) with a shape that
    //      Excel repairs on open ("Repaired Records: Drawing shape"). The WMF bytes
    //      are intact but the shape XML trips Excel's loader. On Bridge/Intacct books
    //      (Braker) these are decorative images on hidden consolidation tabs, so drop
    //      the WMF picture anchors + their relationships; the now-unreferenced .wmf
    //      media are left in place (Excel ignores unreferenced parts).
    for (const relName of Object.keys(out.files)) {
      const dm = relName.match(/^xl\/drawings\/_rels\/(drawing\d+\.xml)\.rels$/);
      if (!dm || out.files[relName].dir) continue;
      let relXml = await out.file(relName).async('string');
      const wmfRids = [...relXml.matchAll(/Id="([^"]+)"[^>]*Target="[^"]*\.wmf"/g)].map(m => m[1]);
      if (!wmfRids.length) continue;
      const drawName = 'xl/drawings/' + dm[1];
      const df = out.file(drawName); if (!df) continue;
      let drawXml = await df.async('string');
      const before = drawXml;
      drawXml = drawXml.replace(/<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/xdr:\1>/g,
        blk => wmfRids.some(r => blk.includes('r:embed="' + r + '"')) ? '' : blk);
      relXml = relXml.replace(/<Relationship\b[^>]*Target="[^"]*\.wmf"[^>]*\/>/g, '');
      if (drawXml !== before) { out.file(drawName, drawXml); out.file(relName, relXml); changed = true; }
    }

    // (2d) Strip broken EXTERNAL hyperlinks. Uploaded invoice/JE source workbooks
    //      sometimes carry a hyperlink on a description cell pointing at a file on
    //      the original author's machine (e.g. a "...\AppData\...\Loan Closing
    //      JE.xlsx" path baked into a cell by whoever built the source). That path
    //      is a relative external OPC relationship Excel cannot resolve, so on open
    //      it rewrites the worksheet part and shows "Replaced Part: /xl/worksheets/
    //      sheetN.xml part with XML error. Load error." It also leaks a third
    //      party's folder layout into our report. Remove every worksheet hyperlink
    //      relationship whose TargetMode is External, then drop the matching
    //      <hyperlink> cell elements (and an emptied <hyperlinks> wrapper). Internal
    //      in-workbook hyperlinks (location="Sheet!A1", no relationship) are left
    //      untouched. Best-effort and non-fatal.
    for (const relName of Object.keys(out.files)) {
      const sm = relName.match(/^xl\/worksheets\/_rels\/(sheet\d+\.xml)\.rels$/);
      if (!sm || out.files[relName].dir) continue;
      let relXml = await out.file(relName).async('string');
      const extIds = [];
      const newRelXml = relXml.replace(/<Relationship\b[^>]*\/>/g, (rel) => {
        if (/Type="[^"]*\/hyperlink"/.test(rel) && /TargetMode="External"/.test(rel)) {
          const idm = rel.match(/Id="([^"]+)"/);
          if (idm) { extIds.push(idm[1]); return ''; }
        }
        return rel;
      });
      if (!extIds.length) continue;
      const sheetName = 'xl/worksheets/' + sm[1];
      const sf = out.file(sheetName); if (!sf) continue;
      let sheetXml = await sf.async('string');
      const idSet = new Set(extIds);
      sheetXml = sheetXml.replace(/<hyperlink\b[^>]*\/>/g, (hl) => {
        const idm = hl.match(/r:id="([^"]+)"/);
        return (idm && idSet.has(idm[1])) ? '' : hl;
      });
      // Drop the wrapper if it is now empty (an empty <hyperlinks/> is itself a
      // schema violation Excel would flag).
      sheetXml = sheetXml
        .replace(/<hyperlinks>\s*<\/hyperlinks>/g, '')
        .replace(/<hyperlinks\s*\/>/g, '');
      out.file(relName, newRelXml);
      out.file(sheetName, sheetXml);
      changed = true;
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
