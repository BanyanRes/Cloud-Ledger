const fs = require('fs');
const JSZip = require('jszip');
const SRC = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 3-31-26_updated by JY.xlsx';
const OUT = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 6-30-26.xlsx';
(async () => {
  for (const [tag, f] of [['SRC', SRC], ['OUT', OUT]]) {
    const zip = await JSZip.loadAsync(fs.readFileSync(f));
    const x = await zip.file('xl/worksheets/sheet1.xml').async('string');
    console.log('=== ' + tag + ' sheet1 shared formulas ===');
    for (const m of x.matchAll(/<c r="([A-Z]+\d+)"[^>]*>\s*<f t="shared"([^>]*)>([^<]*)<\/f>|<c r="([A-Z]+\d+)"[^>]*>\s*<f t="shared"([^>]*)\/>/g)) {
      if (m[1]) console.log(m[1] + ' shared' + m[2] + ' formula="' + m[3] + '"');
      else console.log(m[4] + ' shared' + m[5] + ' (member, no formula)');
    }
    // also confirm which sheet sheet1 is
    const wb = await zip.file('xl/workbook.xml').async('string');
    const wr = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rid = (wr.match(/Id="(rId\d+)"[^>]*Target="worksheets\/sheet1\.xml"/) || wr.match(/Target="worksheets\/sheet1\.xml"[^>]*Id="(rId\d+)"/) || [])[1];
    const nm = (wb.match(new RegExp('<sheet name="([^"]*)"[^>]*r:id="' + rid + '"')) || [])[1];
    console.log('sheet1 = "' + nm + '"');
  }
})().catch(e => console.error(e.stack));
