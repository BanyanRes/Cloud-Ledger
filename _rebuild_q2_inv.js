// CONSOLIDATED build: CLRF Investment Balance 6-30-26.xlsx from the 3-31-26 file.
// Incorporates: TB rebuilds, solved Q2 valuations (frozen-unrealized-gain method),
// shared-formula-safe cell replacement, cache refreshes, calcChain drop.
const fs = require('fs');
const JSZip = require('jszip');

const SRC = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 3-31-26_updated by JY.xlsx';
const OUT = 'C:\\Users\\JimmyYun\\OneDrive - banyanres.com\\Desktop\\CLRF Investment Balance 6-30-26.xlsx';
const TBJ = 'C:\\Users\\JimmyYun\\Downloads\\_tb_63026.json';

const r2 = n => Math.round((Number(n) || 0) * 100) / 100;
const xmlEsc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const dpOf = r => (r.type === 'Asset' || r.type === 'Expense') ? r.balance : -r.balance;

// ============ model (validated vs Q1 to the penny) ============
const S = ymd => { const [y, m, d] = ymd.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5); };
const FVn = (r, n, pv) => -pv * Math.pow(1 + r, n);
const LIQ = S('2026-06-30');
const accr = (contrib, months, anchor, rate) => {
  const fv = FVn(rate / 12, months, -contrib);
  return fv + fv * rate / 365 * (LIQ - S(anchor));
};
function wf(avail, capital, g14, i14, l14) {
  const D24 = -Math.min(avail, capital); let rem = avail + D24;
  const G24 = -Math.min(rem, g14 + D24); rem += G24;
  const I23 = -Math.min(rem, (i14 + D24 + G24) / 0.8); const I24 = I23 * 0.8, I25 = I23 * 0.2; rem += I24 + I25;
  const K23 = -Math.min(rem, (l14 + D24 + G24 + I24) / 0.7); const K24 = K23 * 0.7, K25 = K23 * 0.3; rem += K24 + K25;
  return { sponsor: I25 + K25 + (-rem) * 0.4 };
}
const USC_PCT = 0.11326577820164344;
function clipProceeds(V, loan, nwc) {
  const C6 = V + loan + nwc;
  const em = wf(C6 * (1 - USC_PCT), 45944023, accr(45944023, 35, '2025-05-16', .10), accr(45944023, 35, '2025-05-16', .15), accr(45944023, 35, '2025-05-16', .30));
  const usc = wf(C6 * USC_PCT, 8000000,
    accr(6200000, 15, '2025-05-28', .10) + accr(1800000, 5, '2025-05-31', .10),
    accr(6200000, 15, '2025-05-28', .15) + accr(1800000, 5, '2025-05-31', .15),
    accr(6200000, 15, '2025-05-28', .30) + accr(1800000, 5, '2025-05-31', .30));
  const promote = em.sponsor + usc.sponsor;
  return { proceeds: (C6 + promote) * 0.7651 - promote, promote, C6 };
}
function silsbeeProceeds(V, loan, nwc) {
  const C6 = V + loan + nwc;
  const em = wf(C6, 11712181, accr(11712181, 35, '2025-05-16', .10), accr(11712181, 35, '2025-05-16', .15), accr(11712181, 35, '2025-05-16', .30));
  return { proceeds: (C6 + em.sponsor) * 0.5453 - em.sponsor, promote: em.sponsor, C6 };
}

// ============ data ============
const tbData = JSON.parse(fs.readFileSync(TBJ, 'utf8'));
const balOf = (k, c) => { const r = tbData[k].closing.find(x => x.code === c); return r ? r2(dpOf(r)) : 0; };
const loan = { clip: balOf('clip', '25063'), silsbee: balOf('silsbee', '25063'), buna: balOf('buna', '25063'), srn: balOf('srn', '25063') };
const nwc = { clip: 7337275.42, silsbee: 243613.24, buna: -6157382.07, srn: 2139585.83 };
const book = { clip: 65319338, silsbee: 8787580, buna: 922600, srn: 60408356 };
const BUF = 350000;

// CLIP: exact solve for proceeds = book + 141,167
let lo = 100e6, hi = 250e6;
for (let i = 0; i < 300; i++) { const mid = (lo + hi) / 2; (clipProceeds(mid, loan.clip, nwc.clip).proceeds < book.clip + 141167) ? lo = mid : hi = mid; }
const V_CLIP = r2((lo + hi) / 2);
// Silsbee: current 27,080,000 gives proceeds < book -> solve for proceeds = book + BUF, ceil 10k
let loS = 15e6, hiS = 80e6;
for (let i = 0; i < 300; i++) { const mid = (loS + hiS) / 2; (silsbeeProceeds(mid, loan.silsbee, nwc.silsbee).proceeds < book.silsbee + BUF) ? loS = mid : hiS = mid; }
const V_SIL = Math.ceil(((loS + hiS) / 2) / 10000) * 10000;
// SRN/Buna 100%: V = book + BUF - loan - nwc, ceil 10k
const V_SRN = Math.ceil((book.srn + BUF - loan.srn - nwc.srn) / 10000) * 10000;
const V_BUNA = Math.ceil((book.buna + BUF - loan.buna - nwc.buna) / 10000) * 10000;

const DEV_H12 = 54241570.57; // CLIP GL Dev Costs C28 per Q2 valuation
const G12 = r2(V_CLIP - DEV_H12); // income-producing plug
const clipRes = clipProceeds(V_CLIP, loan.clip, nwc.clip);
const silRes = silsbeeProceeds(V_SIL, loan.silsbee, nwc.silsbee);
const srnProc = r2(V_SRN + loan.srn + nwc.srn);
const bunaProc = r2(V_BUNA + loan.buna + nwc.buna);
const I = { clip: Math.round(clipRes.proceeds), silsbee: Math.round(silRes.proceeds), buna: Math.round(bunaProc), srn: Math.round(srnProc) };
const K = { clip: I.clip, silsbee: Math.min(I.silsbee, book.silsbee), buna: Math.min(I.buna, book.buna), srn: Math.min(I.srn, book.srn) };
const L = { clip: K.clip - book.clip, silsbee: K.silsbee - book.silsbee, buna: K.buna - book.buna, srn: K.srn - book.srn };

// ============ XML helpers (shared-formula-safe) ============
function replaceCell(xml, ref, newCell) {
  const re = new RegExp('<c r="' + ref + '"(?:[^>]*)(?:/>|>[\\s\\S]*?</c>)');
  const m = xml.match(re);
  if (!m) throw new Error('cell ' + ref + ' not found');
  if ((m[0].match(/<c r=/g) || []).length !== 1) throw new Error('overran cell ' + ref);
  // if this cell HOSTS a shared formula group, unshare members first
  const host = m[0].match(/<f t="shared"[^>]*ref="([^"]*)"[^>]*si="(\d+)"[^>]*>([^<]*)<\/f>/);
  let out = xml;
  if (host) out = unshareMembers(out, host[2], ref, host[3], ref);
  const m2 = out.match(re);
  return out.slice(0, m2.index) + newCell + out.slice(m2.index + m2[0].length);
}
const colToN = c => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const nToCol = n => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const parseRef = ref => { const m = ref.match(/^([A-Z]+)(\d+)$/); return { c: colToN(m[1]), r: Number(m[2]) }; };
function shiftFormula(f, dr, dc) {
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d{1,7})/g, (all, dA, col, dB, row, off, s) => {
    const prev = s[off - 1];
    if (prev && /[A-Za-z0-9_.'"]/.test(prev)) return all;
    const nc = dA ? colToN(col) : colToN(col) + dc;
    const nr = dB ? Number(row) : Number(row) + dr;
    if (nc < 1 || nr < 1) return all;
    return dA + nToCol(nc) + dB + nr;
  });
}
// convert all members of shared group si (host hostRef, formula hostF) to plain formulas
function unshareMembers(xml, si, hostRef, hostF, skipRef) {
  const hp = parseRef(hostRef);
  return xml.replace(new RegExp('<c r="([A-Z]+\\d+)"([^>]*)>\\s*<f t="shared"(?![^>]*ref=)[^>]*si="' + si + '"[^>]*/>\\s*(<v>[^<]*</v>)?\\s*</c>', 'g'),
    (all, ref, attrs, v) => {
      if (ref === skipRef) return all;
      const op = parseRef(ref);
      const nf = shiftFormula(hostF, op.r - hp.r, op.c - hp.c);
      return '<c r="' + ref + '"' + attrs + '><f>' + nf + '</f>' + (v || '<v>0</v>') + '</c>';
    });
}
const styleOf = (xml, ref) => { const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*\\bs="(\\d+)"')); return m ? m[1] : null; };
const numCell = (ref, s, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><v>' + v + '</v></c>';
const strCell = (ref, s, t) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(t) + '</t></is></c>';
const fCell = (ref, s, f, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><f>' + xmlEsc(f) + '</f><v>' + v + '</v></c>';

// ============ TB tab builders (same classification as prior build) ============
const isCA = c => /^10\d{3}$/.test(c) || c === '11030' || c === '12000' || c === '12002' || c === '13001' || c === '13100' || /^18\d{3}$/.test(c);
const isCL = c => /^2[0-4]\d{3}$/.test(c);
const isLoan = c => /^25\d{3}$/.test(c);
function buildEntityRows(ds) {
  const acc = {};
  const touch = code => (acc[code] = acc[code] || { open: 0, close: 0, td: 0, tc: 0, name: '', type: '' });
  for (const r of ds.opening) { const a = touch(r.code); a.open = r2(dpOf(r)); a.name = r.name; a.type = r.type; }
  for (const r of ds.closing) { const a = touch(r.code); a.close = r2(dpOf(r)); a.name = r.name; a.type = r.type; }
  for (const r of ds.activity) { const a = touch(r.code); a.td = r2(r.total_debit); a.tc = r2(r.total_credit); if (!a.name) { a.name = r.name; a.type = r.type; } }
  const codes = Object.keys(acc).filter(c => { const a = acc[c]; return Math.abs(a.open) > 0.005 || Math.abs(a.close) > 0.005 || a.td > 0.005 || a.tc > 0.005; })
    .sort((x, y) => Number(x) - Number(y));
  const bad = codes.filter(c => { const a = acc[c]; return Math.abs(r2(a.open + a.td - a.tc) - a.close) > 0.02; });
  return { acc, codes, bad };
}
function build4Col(entName, ds) {
  const { acc, codes, bad } = buildEntityRows(ds);
  const rows = [];
  rows.push([1, [strCell('A1', null, 'Per CloudLedger (cloud-ledger.up.railway.app)')]]);
  rows.push([2, [strCell('A2', null, 'Company name:'), strCell('B2', null, entName)]]);
  rows.push([3, [strCell('A3', null, 'Report name:'), strCell('B3', null, 'Trial balance report')]]);
  rows.push([4, [strCell('A4', null, 'Reporting Book:'), strCell('B4', null, 'ACCRUAL')]]);
  rows.push([5, [strCell('A5', null, 'Start Date:'), strCell('B5', null, '01/01/2026')]]);
  rows.push([6, [strCell('A6', null, 'End Date:'), strCell('B6', null, '06/30/2026')]]);
  rows.push([7, [strCell('A7', null, 'Location:'), strCell('B7', null, 'All (entity total)')]]);
  rows.push([8, [strCell('A8', null, 'Account'), strCell('B8', null, 'Account'), strCell('C8', null, 'Opening balance'), strCell('F8', null, 'Closing balance')]]);
  rows.push([9, [strCell('A9', null, 'Number'), strCell('B9', null, 'Name'), strCell('C9', null, 'on 01/01/2026'), strCell('D9', null, 'Debit'), strCell('E9', null, 'Credit'), strCell('F9', null, 'on 06/30/2026')]]);
  let r = 10; const dataStart = r; const nwcRefs = []; let loanRef = null; const loanCodes = [];
  let tO = 0, tD = 0, tC = 0, tF = 0; let nwcSum = 0;
  for (const c of codes) {
    const a = acc[c];
    rows.push([r, [strCell('A' + r, null, c), strCell('B' + r, null, a.name),
      numCell('C' + r, null, a.open), numCell('D' + r, null, a.td), numCell('E' + r, null, a.tc), numCell('F' + r, null, a.close)]]);
    if (isCA(c) || isCL(c)) { nwcRefs.push('F' + r); nwcSum = r2(nwcSum + a.close); }
    if (isLoan(c)) { loanCodes.push(c); loanRef = 'F' + r; }
    tO = r2(tO + a.open); tD = r2(tD + a.td); tC = r2(tC + a.tc); tF = r2(tF + a.close); r++;
  }
  const dataEnd = r - 1;
  rows.push([r, [strCell('A' + r, null, 'Totals:'),
    fCell('C' + r, null, 'SUM(C' + dataStart + ':C' + dataEnd + ')', tO), fCell('D' + r, null, 'SUM(D' + dataStart + ':D' + dataEnd + ')', tD),
    fCell('E' + r, null, 'SUM(E' + dataStart + ':E' + dataEnd + ')', tC), fCell('F' + r, null, 'SUM(F' + dataStart + ':F' + dataEnd + ')', tF)]]);
  rows.find(p => p[0] === 10)[1].push(strCell('G10', null, 'Net assets at 6/30/26'), fCell('H10', null, nwcRefs.join('+'), nwcSum));
  rows.find(p => p[0] === 12)[1].push(strCell('G12', null, 'NWC = current assets (cash, AR net, prepaids, interest reserve, due-froms) less current liabilities (AP, accruals, due-tos, deposits); excludes 25xxx loan payable.'));
  return { rows, loanRef, loanCodes, nwc: nwcSum, bad, nAccts: codes.length };
}
function build1Col(entName, ds, style) {
  const { acc, codes, bad } = buildEntityRows(ds);
  const rows = [];
  rows.push([1, [strCell('A1', null, 'Company name:'), strCell('B1', null, entName)]]);
  rows.push([2, [strCell('A2', null, 'Report name:'), strCell('B2', null, 'Trial balance report')]]);
  rows.push([3, [strCell('A3', null, 'Reporting Book:'), strCell('B3', null, 'ACCRUAL')]]);
  rows.push([4, [strCell('A4', null, 'Start Date:'), strCell('B4', null, '01/01/2026')]]);
  rows.push([5, [strCell('A5', null, 'End Date:'), strCell('B5', null, '06/30/2026')]]);
  rows.push([6, [strCell('A6', null, 'Location:'), strCell('B6', null, 'All (entity total) — per CloudLedger')]]);
  rows.push([7, [strCell('A7', null, 'Account'), strCell('B7', null, 'Account'), strCell('C7', null, 'Closing balance')]]);
  rows.push([8, [strCell('A8', null, 'Number'), strCell('B8', null, 'Name'), strCell('C8', null, 'on 06/30/2026')]]);
  let r = 9; const dataStart = r; const nwcRefs = []; let loanRef = null; const loanCodes = []; let tot = 0; let nwcSum = 0;
  for (const c of codes) {
    const a = acc[c];
    rows.push([r, [strCell('A' + r, null, c), strCell('B' + r, null, a.name), numCell('C' + r, null, a.close)]]);
    if (isCA(c) || isCL(c)) { nwcRefs.push('C' + r); nwcSum = r2(nwcSum + a.close); }
    if (isLoan(c)) { loanCodes.push(c); loanRef = 'C' + r; }
    tot = r2(tot + a.close); r++;
  }
  const dataEnd = r - 1;
  rows.push([r, [strCell('A' + r, null, 'Totals:'), fCell('C' + r, null, 'SUM(C' + dataStart + ':C' + dataEnd + ')', tot)]]);
  if (style === 'srn') {
    rows.find(p => p[0] === 8)[1].push(strCell('D8', null, 'Net assets 6/30/26'), fCell('F8', null, '+' + nwcRefs.join('+'), nwcSum));
    rows.find(p => p[0] === 10)[1].push(strCell('E10', null, 'NWC = current assets less current liabilities; excludes 25xxx loan payable. Per CloudLedger.'));
  } else {
    rows.find(p => p[0] === 9)[1].push(strCell('D9', null, 'Net assets at 6/30/26'), fCell('E9', null, '+' + nwcRefs.join('+'), nwcSum));
    rows.find(p => p[0] === 11)[1].push(strCell('D11', null, 'NWC = current assets less current liabilities; excludes 25xxx loan payable. Per CloudLedger.'));
  }
  return { rows, loanRef, loanCodes, nwc: nwcSum, bad, nAccts: codes.length };
}
function replaceSheetData(xml, rows) {
  const rowXml = rows.sort((a, b) => a[0] - b[0]).map(p => '<row r="' + p[0] + '">' + p[1].join('') + '</row>').join('');
  let out = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rowXml + '</sheetData>');
  out = out.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, '');
  return out;
}

// ============ main ============
(async () => {
  console.log('SOLVED: CLIP=' + V_CLIP.toFixed(2) + ' Silsbee=' + V_SIL + ' Buna=' + V_BUNA + ' SRN=' + V_SRN);
  console.log('Proceeds: CLIP=' + I.clip + ' Silsbee=' + I.silsbee + ' Buna=' + I.buna + ' SRN=' + I.srn);
  console.log('Unrealized: CLIP=' + L.clip + ' Silsbee=' + L.silsbee + ' Buna=' + L.buna + ' SRN=' + L.srn);
  const zip = await JSZip.loadAsync(fs.readFileSync(SRC));
  const wb0 = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const ridMap = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = (m[0].match(/Id="(rId\d+)"/) || [])[1];
    const tg = (m[0].match(/Target="([^"]+)"/) || [])[1];
    if (id && tg) ridMap[id] = tg;
  }
  const P = {};
  for (const m of wb0.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)) P[m[1].replace(/&amp;/g, '&').trim()] = 'xl/' + ridMap[m[2]].replace(/^\//, '');

  // 1) TB tabs
  const clip = build4Col('County Line Industrial Park LLC (CLIP Property Owner)', tbData.clip);
  const sils = build4Col('CLR Silsbee Property Owner LLC', tbData.silsbee);
  const srn = build1Col('County Line SRN LLC (Sabine River & Northern Railroad)', tbData.srn, 'srn');
  const buna = build1Col('CLR Buna Property Owner LLC', tbData.buna, 'buna');
  for (const [name, b] of [['CLIP TB', clip], ['Silsbee TB', sils], ['SRN TB', srn], ['Buna TB', buna]]) {
    if (b.bad.length) throw new Error(name + ' roll mismatches: ' + b.bad.join(','));
    if (b.loanCodes.length !== 1) console.log('WARN ' + name + ' loans: ' + b.loanCodes.join(','));
    zip.file(P[name], replaceSheetData(await zip.file(P[name]).async('string'), b.rows));
  }

  // 2) entity tabs
  { let x = await zip.file(P['CLIP']).async('string');
    x = replaceCell(x, 'C3', fCell('C3', styleOf(x, 'C3'), "'Investment Balance'!C5", V_CLIP));
    x = replaceCell(x, 'C4', fCell('C4', styleOf(x, 'C4'), "'CLIP TB'!" + clip.loanRef, loan.clip));
    zip.file(P['CLIP'], x); }
  { let x = await zip.file(P['Silsbee']).async('string');
    x = replaceCell(x, 'C4', fCell('C4', styleOf(x, 'C4'), "'Silsbee TB'!" + sils.loanRef, loan.silsbee));
    zip.file(P['Silsbee'], x); }
  { let x = await zip.file(P['SRN']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), "'SRN TB'!" + srn.loanRef, loan.srn));
    zip.file(P['SRN'], x); }
  { let x = await zip.file(P['Buna']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), "'Buna TB'!" + buna.loanRef, loan.buna));
    zip.file(P['Buna'], x); }

  // 3) Valuations: solved amounts. CLIP row: G12=I12 plug, J12=K12 solved total.
  { let x = await zip.file(P['Valuations']).async('string');
    x = replaceCell(x, 'C8', numCell('C8', styleOf(x, 'C8'), S('2026-06-30')));
    x = replaceCell(x, 'D12', numCell('D12', styleOf(x, 'D12'), 65319338.46));
    x = replaceCell(x, 'D13', numCell('D13', styleOf(x, 'D13'), 8787579.67));
    x = replaceCell(x, 'D14', numCell('D14', styleOf(x, 'D14'), 922599.55));
    x = replaceCell(x, 'D15', numCell('D15', styleOf(x, 'D15'), 60408356.37));
    x = replaceCell(x, 'D16', fCell('D16', styleOf(x, 'D16'), 'SUM(D12:D15)', r2(65319338.46 + 8787579.67 + 922599.55 + 60408356.37)));
    x = replaceCell(x, 'F12', numCell('F12', styleOf(x, 'F12'), 144688277));
    x = replaceCell(x, 'G12', numCell('G12', styleOf(x, 'G12'), G12));
    x = replaceCell(x, 'H12', numCell('H12', styleOf(x, 'H12'), DEV_H12));
    x = replaceCell(x, 'I12', fCell('I12', styleOf(x, 'I12'), 'G12', G12));
    x = replaceCell(x, 'J12', fCell('J12', styleOf(x, 'J12'), 'SUM(H12:I12)', V_CLIP));
    x = replaceCell(x, 'K12', fCell('K12', styleOf(x, 'K12'), 'J12', V_CLIP));
    x = replaceCell(x, 'E13', numCell('E13', styleOf(x, 'E13'), V_SIL));
    x = replaceCell(x, 'G13', numCell('G13', styleOf(x, 'G13'), V_SIL));
    x = replaceCell(x, 'E14', numCell('E14', styleOf(x, 'E14'), V_BUNA));
    x = replaceCell(x, 'G14', numCell('G14', styleOf(x, 'G14'), V_BUNA));
    x = replaceCell(x, 'E15', numCell('E15', styleOf(x, 'E15'), V_SRN));
    x = replaceCell(x, 'G15', numCell('G15', styleOf(x, 'G15'), V_SRN));
    for (const [ref, v] of Object.entries({ E20: 16848, F20: 102712, E21: 418611, F21: 17243, E22: 3500, F22: 3969, E23: 15506, F23: 242141 })) {
      x = replaceCell(x, ref, numCell(ref, styleOf(x, ref), v));
    }
    x = x.replace('</sheetData>', '<row r="27"><c r="B27" t="inlineStr"><is><t xml:space="preserve">Interim convention: unrealized gain/(loss) held at 12/31/25 amounts (CLIP +141,167; others 0). Valuations solved from 6/30/26 net assets and loan balances per CloudLedger so estimated sale proceeds equal book carrying value plus the frozen gain (CLIP, exact) or exceed book carrying value by at least $350,000 (Silsbee/Buna/SRN, cost and sales approach figures raised to the required amount). Book carrying values per CLRF GL investment accounts at 6/30/26.</t></is></c></row></sheetData>');
    zip.file(P['Valuations'], x); }

  // 4) Carrying Value
  { let x = await zip.file(P['Carrying Value']).async('string');
    x = replaceCell(x, 'A4', strCell('A4', styleOf(x, 'A4'), 'JUNE 30, 2026'));
    x = replaceCell(x, 'J7', numCell('J7', styleOf(x, 'J7'), book.srn));
    x = replaceCell(x, 'J9', numCell('J9', styleOf(x, 'J9'), book.clip));
    x = replaceCell(x, 'J11', numCell('J11', styleOf(x, 'J11'), book.silsbee));
    x = replaceCell(x, 'J13', numCell('J13', styleOf(x, 'J13'), book.buna));
    zip.file(P['Carrying Value'], x); }

  // 5) waterfall dates
  for (const name of ['CLIP Summary', 'Silsbee Summary']) {
    let x = await zip.file(P[name]).async('string');
    x = replaceCell(x, 'B5', numCell('B5', styleOf(x, 'B5'), S('2026-06-30')));
    zip.file(P[name], x);
  }

  // 6) Investment Balance: repairs + fresh caches (shared-formula-safe replaceCell)
  { let x = await zip.file(P['Investment Balance']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), 'Valuations!K12', V_CLIP));
    x = replaceCell(x, 'C6', fCell('C6', styleOf(x, 'C6'), 'Valuations!K13', V_SIL));
    x = replaceCell(x, 'C7', fCell('C7', styleOf(x, 'C7'), 'Valuations!K14', V_BUNA));
    x = replaceCell(x, 'C8', fCell('C8', styleOf(x, 'C8'), 'Valuations!K15', V_SRN));
    x = replaceCell(x, 'D5', fCell('D5', styleOf(x, 'D5'), "'CLIP TB'!H10", nwc.clip));
    x = replaceCell(x, 'D6', fCell('D6', styleOf(x, 'D6'), "'Silsbee TB'!H10", nwc.silsbee));
    x = replaceCell(x, 'D7', fCell('D7', styleOf(x, 'D7'), "'Buna TB'!E9", nwc.buna));
    x = replaceCell(x, 'D8', fCell('D8', styleOf(x, 'D8'), "'SRN TB'!F8", nwc.srn));
    // E5 host of shared si=1? No: E6 is host (E6:E8). replaceCell(E6) unshares E7/E8 first.
    x = replaceCell(x, 'E5', fCell('E5', styleOf(x, 'E5'), 'SUM(C5:D5)', r2(V_CLIP + nwc.clip)));
    x = replaceCell(x, 'E6', fCell('E6', styleOf(x, 'E6'), 'SUM(C6:D6)', r2(V_SIL + nwc.silsbee)));
    x = replaceCell(x, 'E7', fCell('E7', styleOf(x, 'E7'), 'SUM(C7:D7)', r2(V_BUNA + nwc.buna)));
    x = replaceCell(x, 'E8', fCell('E8', styleOf(x, 'E8'), 'SUM(C8:D8)', r2(V_SRN + nwc.srn)));
    x = replaceCell(x, 'F5', fCell('F5', styleOf(x, 'F5'), "'CLIP TB'!H10", nwc.clip));
    x = replaceCell(x, 'F6', fCell('F6', styleOf(x, 'F6'), "'Silsbee TB'!H10", nwc.silsbee));
    x = replaceCell(x, 'F7', fCell('F7', styleOf(x, 'F7'), "'Buna TB'!E9", nwc.buna));
    x = replaceCell(x, 'F8', fCell('F8', styleOf(x, 'F8'), "'SRN TB'!F8", nwc.srn));
    x = replaceCell(x, 'I5', fCell('I5', styleOf(x, 'I5'), 'ROUND(CLIP!C13,0)', I.clip));
    x = replaceCell(x, 'I6', fCell('I6', styleOf(x, 'I6'), 'ROUND(IF(H6="transitional","N/A",Silsbee!C13),0)', I.silsbee));
    x = replaceCell(x, 'I7', fCell('I7', styleOf(x, 'I7'), 'ROUND(IF(H7="transitional","N/A",Buna!C10),0)', I.buna));
    x = replaceCell(x, 'I8', fCell('I8', styleOf(x, 'I8'), 'ROUND(IF(H8="transitional","N/A",SRN!C10),0)', I.srn));
    x = replaceCell(x, 'J5', fCell('J5', styleOf(x, 'J5'), "'Carrying Value'!J9", book.clip));
    x = replaceCell(x, 'J6', fCell('J6', styleOf(x, 'J6'), "'Carrying Value'!J11", book.silsbee));
    x = replaceCell(x, 'J7', fCell('J7', styleOf(x, 'J7'), "'Carrying Value'!J13", book.buna));
    x = replaceCell(x, 'J8', fCell('J8', styleOf(x, 'J8'), "'Carrying Value'!J7", book.srn));
    x = replaceCell(x, 'K5', fCell('K5', styleOf(x, 'K5'), 'I5', K.clip));
    x = replaceCell(x, 'K6', fCell('K6', styleOf(x, 'K6'), 'IF(I6>J6,J6,I6)', K.silsbee));
    x = replaceCell(x, 'K7', fCell('K7', styleOf(x, 'K7'), 'IF(I7>J7,J7,I7)', K.buna));
    x = replaceCell(x, 'K8', fCell('K8', styleOf(x, 'K8'), 'IF(I8>J8,J8,I8)', K.srn));
    x = replaceCell(x, 'L5', fCell('L5', styleOf(x, 'L5'), 'K5-J5', L.clip));
    x = replaceCell(x, 'L6', fCell('L6', styleOf(x, 'L6'), 'K6-J6', L.silsbee));
    x = replaceCell(x, 'L7', fCell('L7', styleOf(x, 'L7'), 'K7-J7', L.buna));
    x = replaceCell(x, 'L8', fCell('L8', styleOf(x, 'L8'), 'K8-J8', L.srn));
    x = replaceCell(x, 'C9', fCell('C9', styleOf(x, 'C9'), 'SUM(C5:C8)', r2(V_CLIP + V_SIL + V_BUNA + V_SRN)));
    x = replaceCell(x, 'D9', fCell('D9', styleOf(x, 'D9'), 'SUM(D5:D8)', r2(nwc.clip + nwc.silsbee + nwc.buna + nwc.srn)));
    x = replaceCell(x, 'E9', fCell('E9', styleOf(x, 'E9'), 'SUM(E5:E8)', r2(V_CLIP + V_SIL + V_BUNA + V_SRN + nwc.clip + nwc.silsbee + nwc.buna + nwc.srn)));
    x = replaceCell(x, 'J9', fCell('J9', styleOf(x, 'J9'), 'SUM(J5:J8)', book.clip + book.silsbee + book.buna + book.srn));
    x = replaceCell(x, 'K9', fCell('K9', styleOf(x, 'K9'), 'SUM(K5:K8)', K.clip + K.silsbee + K.buna + K.srn));
    x = replaceCell(x, 'L9', fCell('L9', styleOf(x, 'L9'), 'SUM(L5:L8)', L.clip + L.silsbee + L.buna + L.srn));
    zip.file(P['Investment Balance'], x); }

  // 7) Loan Balances_DO NOT USE loan refs
  { let x = await zip.file(P['Loan Balances_DO NOT USE']).async('string');
    x = replaceCell(x, 'W7', fCell('W7', styleOf(x, 'W7'), "-'CLIP TB'!" + clip.loanRef, -loan.clip));
    x = replaceCell(x, 'W8', fCell('W8', styleOf(x, 'W8'), "-'Silsbee TB'!" + sils.loanRef, -loan.silsbee));
    x = replaceCell(x, 'W9', fCell('W9', styleOf(x, 'W9'), "-'SRN TB'!" + srn.loanRef, -loan.srn));
    x = replaceCell(x, 'W10', fCell('W10', styleOf(x, 'W10'), "-'Buna TB'!" + buna.loanRef, -loan.buna));
    zip.file(P['Loan Balances_DO NOT USE'], x); }

  // 8) workbook: fullCalcOnLoad + drop calcChain
  { let wb = await zip.file('xl/workbook.xml').async('string');
    if (/<calcPr/.test(wb)) wb = wb.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
    else wb = wb.replace(/<\/workbook>/, '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/workbook.xml', wb);
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      const ct = await zip.file('[Content_Types].xml').async('string');
      zip.file('[Content_Types].xml', ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
      const wr2 = await zip.file('xl/_rels/workbook.xml.rels').async('string');
      zip.file('xl/_rels/workbook.xml.rels', wr2.replace(/<Relationship [^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
    } }

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(OUT, buf);
  console.log('WROTE ' + OUT + ' (' + buf.length + ' bytes)');

  // ============ verify ============
  const z2 = await JSZip.loadAsync(buf);
  let ok = true;
  const chk = (l, c) => { console.log((c ? 'PASS' : 'FAIL') + ': ' + l); if (!c) ok = false; };
  chk('no calcChain', !z2.file('xl/calcChain.xml'));
  // orphaned shared members scan
  let orphans = 0, refs = 0;
  for (const nm of Object.keys(P)) {
    const x = await z2.file(P[nm]).async('string');
    const hosts = new Set([...x.matchAll(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"/g)].map(mm => mm[1]));
    for (const mm of x.matchAll(/<f t="shared"(?![^>]*ref=)[^>]*si="(\d+)"/g)) if (!hosts.has(mm[1])) orphans++;
    refs += (x.match(/#REF!/g) || []).length;
  }
  chk('0 orphaned shared-formula members', orphans === 0);
  chk('#REF! only pre-existing USC K44 (2 tokens), got ' + refs, refs === 2);
  const inv = await z2.file(P['Investment Balance']).async('string');
  chk('L5 cache = 141167', new RegExp('<c r="L5"[^>]*><f>K5-J5</f><v>141167</v>').test(inv));
  chk('L9 cache = 141167', /<c r="L9"[^>]*><f>SUM\(L5:L8\)<\/f><v>141167<\/v>/.test(inv));
  console.log(ok ? 'ALL CHECKS PASS' : 'FAILURES ABOVE');
  process.exit(ok ? 0 : 1);
})().catch(e => { console.error('FAIL:', e.stack); process.exit(1); });
