path = r"C:\Users\JimmyYun\Cloud-Ledger\server\requisition_preserve.js"
raw=open(path,"rb").read(); nl="\r\n" if b"\r\n" in raw else "\n"
s=raw.decode("utf-8"); orig=s

# Insert a CF-restore block right after the external-links section, before the final
# save. Anchor: the "(2) Re-inject external links" section ends and later there is a
# save. We add our block after the src load (src is already available). Simplest safe
# anchor: right after `const src = await JSZip.loadAsync(originalBuf);` we do nothing
# (src used later); instead insert our restore just before the function returns/saves.
# Find the return/generate near the end of the function.
# We anchor on the external-links block's `const src = await JSZip.loadAsync(originalBuf);`
# and add CF restore immediately after it (src is in scope from there on).

anchor = "    const src = await JSZip.loadAsync(originalBuf);\n"
cf_block = (
"    const src = await JSZip.loadAsync(originalBuf);\n"
"\n"
"    // (CF) Restore each worksheet's original conditional formatting. ExcelJS does\n"
"    // not reliably round-trip <conditionalFormatting> on sheets it modifies (it can\n"
"    // emit a malformed, typeless <cfRule priority=\"1\"/> with no rule body), which\n"
"    // makes Excel show \"we found a problem\" and strip the formatting on open. We\n"
"    // copy the exact original CF blocks from the source sheet back into the output\n"
"    // sheet, matched by worksheet part name (sheetN.xml aligns because sheet order\n"
"    // is preserved). Best-effort and non-fatal.\n"
"    try {\n"
"      const sheetRe = /^xl\\/worksheets\\/sheet\\d+\\.xml$/;\n"
"      const srcSheets = Object.keys(src.files).filter(n => sheetRe.test(n) && !src.files[n].dir);\n"
"      for (const name of srcSheets) {\n"
"        if (!out.files[name]) continue;\n"
"        const srcXml = await src.file(name).async('string');\n"
"        const origCF = (srcXml.match(/<conditionalFormatting\\b[\\s\\S]*?<\\/conditionalFormatting>/g) || []).join('');\n"
"        let outXml = await out.file(name).async('string');\n"
"        const outHasCF = /<conditionalFormatting\\b/.test(outXml);\n"
"        // Nothing to restore if the source had no CF.\n"
"        if (!origCF) {\n"
"          // If the source had none but ExcelJS invented a (broken) block, drop it.\n"
"          if (outHasCF) { outXml = outXml.replace(/<conditionalFormatting\\b[\\s\\S]*?<\\/conditionalFormatting>/g, ''); out.file(name, outXml); changed = true; }\n"
"          continue;\n"
"        }\n"
"        if (outHasCF) {\n"
"          // Replace the first CF block with the full original set, drop any extras.\n"
"          let replaced = false;\n"
"          outXml = outXml.replace(/<conditionalFormatting\\b[\\s\\S]*?<\\/conditionalFormatting>/g, () => {\n"
"            if (replaced) return '';\n"
"            replaced = true; return origCF;\n"
"          });\n"
"        } else {\n"
"          // ExcelJS dropped CF entirely; re-insert before pageMargins (schema order)\n"
"          // or before </worksheet> if there is no pageMargins.\n"
"          if (/<pageMargins\\b/.test(outXml)) outXml = outXml.replace(/<pageMargins\\b/, origCF + '<pageMargins');\n"
"          else outXml = outXml.replace('</worksheet>', origCF + '</worksheet>');\n"
"        }\n"
"        out.file(name, outXml);\n"
"        changed = true;\n"
"      }\n"
"      // The restored CF references dxfIds from the source styles. Make sure the\n"
"      // output's <dxfs> collection matches the source's so those indices resolve.\n"
"      const srcStyles = await src.file('xl/styles.xml').async('string');\n"
"      const srcDxfs = (srcStyles.match(/<dxfs\\b[\\s\\S]*?<\\/dxfs>/) || srcStyles.match(/<dxfs\\b[^>]*\\/>/) || [null])[0];\n"
"      if (srcDxfs && out.files['xl/styles.xml']) {\n"
"        let outStyles = await out.file('xl/styles.xml').async('string');\n"
"        const outHasDxfs = /<dxfs\\b/.test(outStyles);\n"
"        if (outHasDxfs) {\n"
"          const replacedStyles = outStyles.replace(/<dxfs\\b[\\s\\S]*?<\\/dxfs>|<dxfs\\b[^>]*\\/>/, srcDxfs);\n"
"          if (replacedStyles !== outStyles) { out.file('xl/styles.xml', replacedStyles); changed = true; }\n"
"        } else {\n"
"          // Insert <dxfs> in its schema position: after </cellStyles> (or before\n"
"          // </styleSheet> as a fallback).\n"
"          let ins;\n"
"          if (/<\\/cellStyles>/.test(outStyles)) ins = outStyles.replace('</cellStyles>', '</cellStyles>' + srcDxfs);\n"
"          else ins = outStyles.replace('</styleSheet>', srcDxfs + '</styleSheet>');\n"
"          if (ins !== outStyles) { out.file('xl/styles.xml', ins); changed = true; }\n"
"        }\n"
"      }\n"
"    } catch (_) { /* non-fatal: leave ExcelJS CF as-is */ }\n"
)
assert s.count(anchor)==1, "anchor count=%d"%s.count(anchor)
s=s.replace(anchor, cf_block, 1)
assert s!=orig
open(path,"w",encoding="utf-8",newline="").write(s)
print("CF + dxfs restoration added to finalizeRequisitionWorkbook")
