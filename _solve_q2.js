// Solve the Q2 valuation amounts that hold unrealized G/L at the frozen amounts
// (CLIP +141,167; Silsbee/Buna/SRN 0). Validates the waterfall replication
// against Q1 known outputs first.
const fs = require('fs');
const r2 = n => Math.round(n * 100) / 100;
const dpOf = r => (r.type === 'Asset' || r.type === 'Expense') ? r.balance : -r.balance;
const tb = JSON.parse(fs.readFileSync('C:\\Users\\JimmyYun\\Downloads\\_tb_63026.json', 'utf8'));
const bal = (k, c) => { const r = tb[k].closing.find(x => x.code === c); return r ? dpOf(r) : 0; };

const S = ymd => { const [y, m, d] = ymd.split('-').map(Number); return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5); };
const FV = (r, n, pv) => -pv * Math.pow(1 + r, n);

function makeModel(liqDateStr) {
  const liq = S(liqDateStr);
  const accr = (contrib, months, anchor, rate) => {
    const fv = FV(rate / 12, months, -contrib);
    return fv + fv * rate / 365 * (liq - S(anchor));
  };
  function waterfall(avail, capital, g14, i14, l14) {
    const D24 = -Math.min(avail, capital); let rem = avail + D24;
    const G22 = g14 + D24; const G24 = -Math.min(rem, G22); rem += G24;
    const I22 = (i14 + D24 + G24) / 0.8; const I23 = -Math.min(rem, I22); const I24 = I23 * 0.8, I25 = I23 * 0.2; rem += I24 + I25;
    const K22 = (l14 + D24 + G24 + I24) / 0.7; const K23 = -Math.min(rem, K22); const K24 = K23 * 0.7, K25 = K23 * 0.3; rem += K24 + K25;
    const N23 = -rem; const N25 = N23 * 0.4;
    return { sponsor: I25 + K25 + N25 };
  }
  const uscPct = 0.11326577820164344;
  function clipProceeds(V, loan, nwc) {
    const C6 = V + loan + nwc;
    const em = waterfall(C6 * (1 - uscPct), 45944023,
      accr(45944023, 35, '2025-05-16', 0.10), accr(45944023, 35, '2025-05-16', 0.15), accr(45944023, 35, '2025-05-16', 0.30));
    const g32 = accr(6200000, 15, '2025-05-28', 0.10) + accr(1800000, 5, '2025-05-31', 0.10);
    const i32 = accr(6200000, 15, '2025-05-28', 0.15) + accr(1800000, 5, '2025-05-31', 0.15);
    const l32 = accr(6200000, 15, '2025-05-28', 0.30) + accr(1800000, 5, '2025-05-31', 0.30);
    const usc = waterfall(C6 * uscPct, 8000000, g32, i32, l32);
    const promote = em.sponsor + usc.sponsor;
    return { proceeds: (C6 + promote) * 0.7651 - promote, promote };
  }
  function silsbeeProceeds(V, loan, nwc) {
    const C6 = V + loan + nwc;
    const em = waterfall(C6, 11712181,
      accr(11712181, 35, '2025-05-16', 0.10), accr(11712181, 35, '2025-05-16', 0.15), accr(11712181, 35, '2025-05-16', 0.30));
    return { proceeds: (C6 + em.sponsor) * 0.5453 - em.sponsor, promote: em.sponsor };
  }
  return { clipProceeds, silsbeeProceeds };
}

// ---- VALIDATE vs Q1 (3/31/26) ----
const q1 = makeModel('2026-03-31');
const v1 = q1.clipProceeds(149431786.26, -72003929.13, 7602097.11);
console.log('Q1 validation CLIP: proceeds=' + v1.proceeds.toFixed(2) + ' (expect 65460505.00), promote=' + v1.promote.toFixed(2) + ' (expect -1720251.13)');
const v2 = q1.silsbeeProceeds(27080000, -10971611.24, 105291.53);
console.log('Q1 validation Silsbee: proceeds=' + v2.proceeds.toFixed(2) + ' (expect 8841319.86), promote=' + v2.promote.toFixed(2) + ' (expect 0)');
if (Math.abs(v1.proceeds - 65460505) > 1 || Math.abs(v2.proceeds - 8841319.86) > 1) { console.log('VALIDATION FAILED'); process.exit(1); }
console.log('MODEL VALIDATED\n');

// ---- SOLVE Q2 (6/30/26) ----
const q2 = makeModel('2026-06-30');
const loan = { clip: r2(bal('clip', '25063')), silsbee: r2(bal('silsbee', '25063')), buna: r2(bal('buna', '25063')), srn: r2(bal('srn', '25063')) };
const nwc = { clip: 7337275.42, silsbee: 243613.24, buna: -6157382.07, srn: 2139585.83 };
const book = { clip: 65319338, silsbee: 8787580, buna: 922600, srn: 60408356 };

// CLIP: bisect V so proceeds == book + 141,167 = 65,460,505 (exact, unrounded valuation)
const targetClip = book.clip + 141167;
let lo = 100e6, hi = 200e6;
for (let i = 0; i < 200; i++) {
  const mid = (lo + hi) / 2;
  (q2.clipProceeds(mid, loan.clip, nwc.clip).proceeds < targetClip) ? lo = mid : hi = mid;
}
const Vclip = r2((lo + hi) / 2);
const chk = q2.clipProceeds(Vclip, loan.clip, nwc.clip);
console.log('CLIP solved valuation = ' + Vclip.toFixed(2) + ' -> proceeds ' + chk.proceeds.toFixed(2) + ' (target ' + targetClip + '), promote ' + chk.promote.toFixed(2));

// Silsbee: min V so proceeds >= book, then round UP to nearest 10k; verify
let loS = 20e6, hiS = 60e6;
for (let i = 0; i < 200; i++) {
  const mid = (loS + hiS) / 2;
  (q2.silsbeeProceeds(mid, loan.silsbee, nwc.silsbee).proceeds < book.silsbee) ? loS = mid : hiS = mid;
}
const VsilMin = (loS + hiS) / 2;
const VsilRound = Math.ceil(VsilMin / 10000) * 10000;
const silChk = q2.silsbeeProceeds(VsilRound, loan.silsbee, nwc.silsbee);
console.log('Silsbee min valuation = ' + VsilMin.toFixed(2) + ' -> rounded up ' + VsilRound + ' -> proceeds ' + silChk.proceeds.toFixed(2) + ' vs book ' + book.silsbee + ', promote ' + silChk.promote.toFixed(2));

// SRN / Buna: 100% owned -> proceeds = V + loan + NWC; min V = book - loan - nwc
for (const k of ['srn', 'buna']) {
  const Vmin = book[k] - loan[k] - nwc[k];
  const Vround = Math.ceil(Vmin / 10000) * 10000;
  const proceeds = Vround + loan[k] + nwc[k];
  console.log(k.toUpperCase() + ' min valuation = ' + Vmin.toFixed(2) + ' -> rounded up ' + Vround + ' -> proceeds ' + proceeds.toFixed(2) + ' vs book ' + book[k]);
}
console.log('\nQ1 comparison valuations: CLIP 149,431,786.26 | Silsbee 27,080,000 | Buna 6,840,000 | SRN 77,950,000');
