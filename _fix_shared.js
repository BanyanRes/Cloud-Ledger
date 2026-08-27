// Find orphaned shared-formula groups (members whose host was replaced) in ALL
// sheets of the output, and fix by converting members to plain formulas with
// row/col-shifted expressions derived from the source host formula.
const fs = require('fs');
const JSZip = require('jszip');
const SRC = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 3-31-26_updated by JY.xlsx';
const OUT = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 6-30-26.xlsx';

const colToN = c => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const nToCol = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const parseRef = ref => { const m = ref.match(/^([A-Z]+)(\d+)$/); return { c: colToN(m[1]), r: Number(m[2]) }; };

// Shift relative A1 refs in a formula by (dr, dc); $-anchored parts stay fixed.
function shiftFormula(f, dr, dc) {
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d{1,7})/g, (all, dA, col, dB, row, off, s) => {
    // skip if part of a sheet name or function-like token: preceded by letter/quote
    const prev = s[off - 1];
    if (prev && /[A-Za-z0-9_.'"!]/.test(prev) && prev !== '!' && prev !== ',' && prev !== '(' && prev !== '+' && prev !== '-' && prev !== '*' && prev !== '/' && prev !== '=' && prev !== ':' && prev !== '<' && prev !== '>' && prev !== ' ') return all;
    const nc = dA ? colToN(col) : colToN(col) + dc;
    const nr = dB ? Number(row) : Number(row) + dr;
    if (nc < 1 || nr < 1) return all;
    return dA + nToCol(nc) + dB + nr;
  });
}

(async () => {
  const zSrc = await JSZip.loadAsync(fs.readFileSync(SRC));
  const zOut = await JSZip.loadAsync(fs.readFileSync(OUT));
  const wb = await zOut.file('xl/workbook.xml').async('string');
  const wr = await zOut.file('xl/_rels/workbook.xml.rels').async('string');
  const ridMap = {};
  for (const m of wr.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = (m[0].match(/Id="(rId\d+)"/) || [])[1];
    const tg = (m[0].match(/Target="([^"]+)"/) || [])[1];
    if (id && tg) ridMap[id] = tg;
  }
  let fixedTotal = 0;
  for (const m of wb.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)) {
    const name = m[1].trim();
    const part = 'xl/' + ridMap[m[2]].replace(/^\//, '');
    let x = await zOut.file(part).async('string');
    // hosts and members in OUTPUT
    const hosts = new Set([...x.matchAll(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"/g)].map(mm => mm[1]));
    const members = [...x.matchAll(/<c r="([A-Z]+\d+)"[^>]*>\s*<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"[^>]*\/>\s*(?:<v>([^<]*)<\/v>)?/g)]
      .map(mm => ({ ref: mm[1], si: mm[2], v: mm[3] }));
    const orphans = members.filter(mm => !hosts.has(mm.si));
    if (!orphans.length) continue;
    // source host formula for each orphaned si
    const sx = await zSrc.file(part).async('string');
    const srcHosts = {};
    for (const mm of sx.matchAll(/<c r="([A-Z]+\d+)"[^>]*>\s*<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"[^>]*>([^<]*)<\/f>/g)) {
      srcHosts[mm[2]] = { ref: mm[1], f: mm[3] };
    }
    for (const o of orphans) {
      const h = srcHosts[o.si];
      if (!h) { console.log('ERROR: no source host for si=' + o.si + ' on ' + name); process.exit(1); }
      const hp = parseRef(h.ref), op = parseRef(o.ref);
      const nf = shiftFormula(h.f, op.r - hp.r, op.c - hp.c);
      // pull the source member's cached value as fallback (recalc-on-load will refresh)
      const cache = (o.v !== undefined && o.v !== '' && o.v !== '#REF!') ? o.v : '0';
      const cellRe = new RegExp('<c r="' + o.ref + '"([^>]*)>\\s*<f t="shared"[^>]*/>\\s*(?:<v>[^<]*</v>)?\\s*</c>');
      const mm = x.match(cellRe);
      if (!mm) { console.log('ERROR: could not isolate member cell ' + o.ref + ' on ' + name); process.exit(1); }
      x = x.replace(cellRe, '<c r="' + o.ref + '"$1><f>' + nf + '</f><v>' + cache + '</v></c>');
      console.log(name + ' ' + o.ref + ': orphaned si=' + o.si + ' -> plain =' + nf);
      fixedTotal++;
    }
    zOut.file(part, x);
  }
  if (fixedTotal) {
    fs.writeFileSync('C:\Users\JimmyYun\Cloud-Ledger\_fixed_inv.xlsx', await zOut.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    console.log('FIXED ' + fixedTotal + ' orphaned shared-formula members; file rewritten');
  } else {
    console.log('no orphans found');
  }
  // final re-scan
  const z2 = await JSZip.loadAsync(fs.readFileSync('C:\Users\JimmyYun\Cloud-Ledger\_fixed_inv.xlsx'));
  let remaining = 0;
  for (const m of wb.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)) {
    const x = await z2.file('xl/' + ridMap[m[2]].replace(/^\//, '')).async('string');
    const hosts = new Set([...x.matchAll(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"/g)].map(mm => mm[1]));
    for (const mm of x.matchAll(/<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"/g)) if (!hosts.has(mm[1])) remaining++;
  }
  console.log('orphaned members remaining after fix: ' + remaining);
  process.exit(remaining ? 1 : 0);
})().catch(e => { console.error(e.stack); process.exit(1); });
