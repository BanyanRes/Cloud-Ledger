// --- CLRF Investment & Valuation workpapers (quarterly) -----------------------
//
// One run produces TWO workbooks for County Line Rail Fund I, LP (entity 40),
// filed under Workpapers/Investment & Valuation/<Qn YYYY>/:
//
//   1. "CLRF Investment Balance M-D-YY.xlsx" -- rolled forward from the prior
//      quarter's investment workpaper. The four portfolio-company TB tabs are
//      rebuilt from the live CloudLedger trial balances (entities 54/39/38/37),
//      net working capital and loan balances are re-derived, and the Valuations
//      tab carries the SOLVED quarter valuations (below).
//
//   2. "CLRF Qn YYYY Valuation_dist.xlsx" -- the appraiser-format valuation
//      workbook (via server/valuation.js transform), generated AGAINST the
//      solved amounts so its Summary matches the investment workpaper's
//      Valuations tab exactly.
//
// Interim valuation convention (per JY): during the year the unrealized
// gain/(loss) per investment is FROZEN at the 12/31 amounts (CLIP +141,167 --
// the balance of CLRF account 121012; all others 0). Each quarter:
//   - CLIP: solve the total valuation so the waterfall's estimated sale
//     proceeds equal book carrying value (121011, rounded) + frozen gain,
//     exactly. The development component comes straight from the CLIP GL
//     (dev-cost account set), so the sales-comparison / income-producing
//     component is the plug: total - dev component.
//   - Silsbee / Buna / SRN: keep the prior quarter's valuation if it still
//     yields proceeds above book carrying value; otherwise raise the cost- and
//     sales-approach figures to the smallest 10k-rounded valuation whose
//     proceeds exceed book by at least $350,000.
//
// The hypothetical-liquidation waterfall (pref accrual = FV for the anchored
// full months + simple interest for the stub through the quarter end) is
// replicated here for the solver; every model parameter (member capital,
// ownership %, USC split, contribution dates/amounts, full-month anchors) is
// PARSED FROM THE TEMPLATE's cached cells, not hardcoded, so the solver tracks
// the workpaper if those inputs ever change.

const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');

const CLRF = 40;
const PORTFOLIO = { clip: 54, silsbee: 39, buna: 38, srn: 37 };
const BUFFER = 350000; // proceeds must exceed book by at least this (non-CLIP)
const IV_FOLDER = (qtr) => 'Workpapers/Investment & Valuation/' + qtr.quarter + ' ' + qtr.year;

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const dpOf = (row) => (row.type === 'Asset' || row.type === 'Expense') ? row.balance : -row.balance;

// Excel serial (1900 system, base 1899-12-30)
function excelSerial(ymd) {
  const [y, m, d] = ymd.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}
const MONTHS = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
  'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
const shortDate = (ymd) => { const [y, m, d] = ymd.split('-').map(Number); return m + '-' + d + '-' + String(y).slice(2); };
const slashDate = (ymd) => { const [y, m, d] = ymd.split('-').map(Number); return String(m).padStart(2, '0') + '/' + String(d).padStart(2, '0') + '/' + y; };
const mdyy = (ymd) => { const [y, m, d] = ymd.split('-').map(Number); return m + '/' + d + '/' + String(y).slice(2); };

// -- NWC classification (mirrors the Q1 2026 manual selections) ---------------
const isCA = (c) => /^10\d{3}$/.test(c) || c === '11030' || c === '12000' || c === '12002'
  || c === '13001' || c === '13100' || /^18\d{3}$/.test(c);
const isCL = (c) => /^2[0-4]\d{3}$/.test(c);
const isLoan = (c) => /^25\d{3}$/.test(c);

// -- Portfolio TB gather -------------------------------------------------------
function gatherPortfolio(ctx, qtr) {
  const { computeBalances } = ctx;
  const y = Number(qtr.year);
  const openAsOf = (y - 1) + '-12-31';
  const closeBefore = y + '-01-01';
  const out = {};
  for (const key of Object.keys(PORTFOLIO)) {
    const eid = PORTFOLIO[key];
    const opening = computeBalances(eid, { as_of: openAsOf, close_pl_before: closeBefore });
    const closing = computeBalances(eid, { as_of: qtr.end, close_pl_before: closeBefore });
    const activity = computeBalances(eid, { from: closeBefore, to: qtr.end });
    const acc = {};
    const touch = (c) => (acc[c] = acc[c] || { open: 0, close: 0, td: 0, tc: 0, name: '', type: '' });
    for (const r of opening) { const a = touch(r.code); a.open = r2(dpOf(r)); a.name = r.name; a.type = r.type; }
    for (const r of closing) { const a = touch(r.code); a.close = r2(dpOf(r)); a.name = r.name; a.type = r.type; }
    for (const r of activity) { const a = touch(r.code); a.td = r2(r.total_debit); a.tc = r2(r.total_credit); if (!a.name) { a.name = r.name; a.type = r.type; } }
    const codes = Object.keys(acc).filter((c) => {
      const a = acc[c];
      return Math.abs(a.open) > 0.005 || Math.abs(a.close) > 0.005 || a.td > 0.005 || a.tc > 0.005;
    }).sort((a, b) => Number(a) - Number(b));
    const bad = codes.filter((c) => { const a = acc[c]; return Math.abs(r2(a.open + a.td - a.tc) - a.close) > 0.02; });
    if (bad.length) throw new Error('TB roll check failed for ' + key + ' accounts: ' + bad.join(', '));
    let nwc = 0; let loanBal = 0; const loanCodes = [];
    for (const c of codes) {
      if (isCA(c) || isCL(c)) nwc = r2(nwc + acc[c].close);
      if (isLoan(c) && Math.abs(acc[c].close) > 0.005) { loanCodes.push(c); loanBal = r2(loanBal + acc[c].close); }
    }
    out[key] = { eid, acc, codes, nwc, loanBal, loanCodes };
  }
  return out;
}

// -- Template parameter parsing -------------------------------------------------
function cellCache(xml, ref) {
  const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*>(?:<f[^>]*>[\\s\\S]*?</f>|<f[^>]*/>)?<v>([^<]+)</v>'));
  if (!m) return null;
  const n = Number(m[1]);
  return isFinite(n) ? n : null;
}
async function sheetMap(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rid = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = (m[0].match(/Id="(rId\d+)"/) || [])[1];
    const tg = (m[0].match(/Target="([^"]+)"/) || [])[1];
    if (id && tg) rid[id] = tg;
  }
  const P = {};
  for (const m of wb.matchAll(/<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g)) {
    if (rid[m[2]]) P[m[1].replace(/&amp;/g, '&').trim()] = 'xl/' + rid[m[2]].replace(/^\//, '');
  }
  return P;
}
// Model params from the prior investment workpaper (all cached values).
async function parseModelParams(zip, P) {
  const need = async (n) => zip.file(P[n]).async('string');
  const clipTab = await need('CLIP');
  const silTab = await need('Silsbee');
  const clipSum = await need('CLIP Summary');
  const emw = await need('Existing Member Waterfall');
  const uscw = await need('USC Waterfall');
  const silw = await need('Existing Members Waterfall');
  const val = await need('Valuations');
  const req = (v, what) => { if (v === null || v === undefined || !isFinite(v)) throw new Error('template parse failed: ' + what); return v; };
  return {
    clipOwnPct: req(cellCache(clipTab, 'C9'), 'CLIP!C9 ownership %'),
    silOwnPct: req(cellCache(silTab, 'C9'), 'Silsbee!C9 ownership %'),
    uscPct: req(cellCache(clipSum, 'B8'), 'CLIP Summary!B8 USC %'),
    em: {
      capital: req(cellCache(emw, 'B28'), 'EMW B28 capital'),
      months: req(cellCache(emw, 'E12'), 'EMW E12 full months'),
      anchor: req(cellCache(emw, 'C12'), 'EMW C12 stub anchor'),
    },
    sil: {
      capital: req(cellCache(silw, 'B28'), 'Silsbee EMW B28 capital'),
      months: req(cellCache(silw, 'E12'), 'Silsbee EMW E12 months'),
      anchor: req(cellCache(silw, 'C12'), 'Silsbee EMW C12 anchor'),
    },
    usc: {
      capital: req(cellCache(uscw, 'B44'), 'USC WF B44 capital'),
      c1: -req(cellCache(uscw, 'C7'), 'USC WF C7 contribution 1'),
      m1: req(cellCache(uscw, 'E17'), 'USC WF E17 months 1'),
      a1: req(cellCache(uscw, 'C17'), 'USC WF C17 anchor 1'),
      c2: -req(cellCache(uscw, 'C11'), 'USC WF C11 contribution 2'),
      m2: req(cellCache(uscw, 'E25'), 'USC WF E25 months 2'),
      a2: req(cellCache(uscw, 'C25'), 'USC WF C25 anchor 2'),
    },
    priorVals: {
      clip: cellCache(val, 'K12'),
      silsbee: cellCache(val, 'K13'),
      buna: cellCache(val, 'K14'),
      srn: cellCache(val, 'K15'),
    },
    incomeF12: cellCache(val, 'F12'),
    propInfo: ['E20', 'F20', 'E21', 'F21', 'E22', 'F22', 'E23', 'F23']
      .reduce((o, ref) => { o[ref] = cellCache(val, ref); return o; }, {}),
  };
}

// -- Waterfall model ------------------------------------------------------------
function makeModel(params, liqSerial) {
  const FV = (r, n, pv) => -pv * Math.pow(1 + r, n);
  const accr = (contrib, months, anchorSerial, rate) => {
    const fv = FV(rate / 12, months, -contrib);
    return fv + fv * rate / 365 * (liqSerial - anchorSerial);
  };
  function wf(avail, capital, g14, i14, l14) {
    const D24 = -Math.min(avail, capital); let rem = avail + D24;
    const G24 = -Math.min(rem, g14 + D24); rem += G24;
    const I23 = -Math.min(rem, (i14 + D24 + G24) / 0.8); const I24 = I23 * 0.8, I25 = I23 * 0.2; rem += I24 + I25;
    const K23 = -Math.min(rem, (l14 + D24 + G24 + I24) / 0.7); const K24 = K23 * 0.7, K25 = K23 * 0.3; rem += K24 + K25;
    return { sponsor: I25 + K25 + (-rem) * 0.4 };
  }
  const emAcc = (rate) => accr(params.em.capital, params.em.months, params.em.anchor, rate);
  const silAcc = (rate) => accr(params.sil.capital, params.sil.months, params.sil.anchor, rate);
  const uscAcc = (rate) => accr(params.usc.c1, params.usc.m1, params.usc.a1, rate)
    + accr(params.usc.c2, params.usc.m2, params.usc.a2, rate);
  return {
    clipProceeds(V, loan, nwc) {
      const C6 = V + loan + nwc;
      const em = wf(C6 * (1 - params.uscPct), params.em.capital, emAcc(.10), emAcc(.15), emAcc(.30));
      const usc = wf(C6 * params.uscPct, params.usc.capital, uscAcc(.10), uscAcc(.15), uscAcc(.30));
      const promote = em.sponsor + usc.sponsor;
      return { proceeds: (C6 + promote) * params.clipOwnPct - promote, promote };
    },
    silsbeeProceeds(V, loan, nwc) {
      const C6 = V + loan + nwc;
      const em = wf(C6, params.sil.capital, silAcc(.10), silAcc(.15), silAcc(.30));
      return { proceeds: (C6 + em.sponsor) * params.silOwnPct - em.sponsor, promote: em.sponsor };
    },
  };
}

// -- Solver ----------------------------------------------------------------------
function solveValuations(model, port, books, fvAdj, priorVals) {
  const bisect = (fn, target, lo, hi) => {
    for (let i = 0; i < 300; i++) { const mid = (lo + hi) / 2; (fn(mid) < target) ? lo = mid : hi = mid; }
    return (lo + hi) / 2;
  };
  const out = {};
  // CLIP: exact solve for proceeds = book + frozen gain
  const clipTarget = books.clip + fvAdj.clip;
  out.clip = {
    valuation: r2(bisect((v) => model.clipProceeds(v, port.clip.loanBal, port.clip.nwc).proceeds, clipTarget, 50e6, 500e6)),
    target_proceeds: clipTarget, changed: true,
  };
  out.clip.promote = r2(model.clipProceeds(out.clip.valuation, port.clip.loanBal, port.clip.nwc).promote);
  // Silsbee (waterfall) and SRN/Buna (100%): keep prior if proceeds clear book, else raise
  const evalP = {
    silsbee: (v) => model.silsbeeProceeds(v, port.silsbee.loanBal, port.silsbee.nwc).proceeds,
    buna: (v) => v + port.buna.loanBal + port.buna.nwc,
    srn: (v) => v + port.srn.loanBal + port.srn.nwc,
  };
  for (const k of ['silsbee', 'buna', 'srn']) {
    const prior = priorVals[k];
    if (prior !== null && evalP[k](prior) > books[k]) {
      out[k] = { valuation: prior, changed: false };
    } else {
      const minV = (k === 'silsbee')
        ? bisect(evalP.silsbee, books.silsbee + BUFFER, 5e6, 200e6)
        : books[k] + BUFFER - port[k].loanBal - port[k].nwc;
      out[k] = { valuation: Math.ceil(minV / 10000) * 10000, changed: true };
    }
    out[k].proceeds = r2(evalP[k](out[k].valuation));
  }
  out.clip.proceeds = r2(model.clipProceeds(out.clip.valuation, port.clip.loanBal, port.clip.nwc).proceeds);
  return out;
}

// -- XML cell helpers (shared-formula-safe) --------------------------------------
const colToN = (c) => { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const nToCol = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const parseRef = (ref) => { const m = ref.match(/^([A-Z]+)(\d+)$/); return { c: colToN(m[1]), r: Number(m[2]) }; };
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
function unshareMembers(xml, si, hostCellRef, hostF, skipRef) {
  const hp = parseRef(hostCellRef);
  return xml.replace(new RegExp('<c r="([A-Z]+\\d+)"([^>]*)>\\s*<f t="shared"(?![^>]*ref=)[^>]*si="' + si + '"[^>]*/>\\s*(<v>[^<]*</v>)?\\s*</c>', 'g'),
    (all, ref, attrs, v) => {
      if (ref === skipRef) return all;
      const op = parseRef(ref);
      const nf = shiftFormula(hostF, op.r - hp.r, op.c - hp.c);
      return '<c r="' + ref + '"' + attrs + '><f>' + nf + '</f>' + (v || '<v>0</v>') + '</c>';
    });
}
// Replace a cell; if it hosts a shared-formula group, first convert the group's
// member cells to standalone formulas so Excel does not flag orphaned members.
function replaceCell(xml, ref, newCell) {
  const re = new RegExp('<c r="' + ref + '"(?:[^>]*)(?:/>|>[\\s\\S]*?</c>)');
  const m = xml.match(re);
  if (!m) throw new Error('cell ' + ref + ' not found in sheet XML');
  if ((m[0].match(/<c r=/g) || []).length !== 1) throw new Error('overran cell ' + ref);
  const host = m[0].match(/<f t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"[^>]*>([^<]*)<\/f>/);
  let out = xml;
  if (host) out = unshareMembers(out, host[1], ref, host[2], ref);
  const m2 = out.match(re);
  return out.slice(0, m2.index) + newCell + out.slice(m2.index + m2[0].length);
}
const styleOf = (xml, ref) => { const m = xml.match(new RegExp('<c r="' + ref + '"[^>]*\\bs="(\\d+)"')); return m ? m[1] : null; };
const numCell = (ref, s, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><v>' + v + '</v></c>';
const strCell = (ref, s, t) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(t) + '</t></is></c>';
const fCell = (ref, s, f, v) => '<c r="' + ref + '"' + (s ? ' s="' + s + '"' : '') + '><f>' + xmlEsc(f) + '</f><v>' + v + '</v></c>';

// -- TB tab builders --------------------------------------------------------------
function tbRows4Col(entName, ent, qtr) {
  const rows = [];
  rows.push([1, [strCell('A1', null, 'Per CloudLedger (cloud-ledger.up.railway.app)')]]);
  rows.push([2, [strCell('A2', null, 'Company name:'), strCell('B2', null, entName)]]);
  rows.push([3, [strCell('A3', null, 'Report name:'), strCell('B3', null, 'Trial balance report')]]);
  rows.push([4, [strCell('A4', null, 'Reporting Book:'), strCell('B4', null, 'ACCRUAL')]]);
  rows.push([5, [strCell('A5', null, 'Start Date:'), strCell('B5', null, '01/01/' + qtr.year)]]);
  rows.push([6, [strCell('A6', null, 'End Date:'), strCell('B6', null, slashDate(qtr.end))]]);
  rows.push([7, [strCell('A7', null, 'Location:'), strCell('B7', null, 'All (entity total)')]]);
  rows.push([8, [strCell('A8', null, 'Account'), strCell('B8', null, 'Account'), strCell('C8', null, 'Opening balance'), strCell('F8', null, 'Closing balance')]]);
  rows.push([9, [strCell('A9', null, 'Number'), strCell('B9', null, 'Name'), strCell('C9', null, 'on 01/01/' + qtr.year), strCell('D9', null, 'Debit'), strCell('E9', null, 'Credit'), strCell('F9', null, 'on ' + slashDate(qtr.end))]]);
  let r = 10; const dataStart = r; const nwcRefs = []; let loanRef = null;
  let tO = 0, tD = 0, tC = 0, tF = 0;
  for (const c of ent.codes) {
    const a = ent.acc[c];
    rows.push([r, [strCell('A' + r, null, c), strCell('B' + r, null, a.name),
      numCell('C' + r, null, a.open), numCell('D' + r, null, a.td), numCell('E' + r, null, a.tc), numCell('F' + r, null, a.close)]]);
    if (isCA(c) || isCL(c)) nwcRefs.push('F' + r);
    if (isLoan(c) && Math.abs(a.close) > 0.005) loanRef = 'F' + r;
    tO = r2(tO + a.open); tD = r2(tD + a.td); tC = r2(tC + a.tc); tF = r2(tF + a.close); r++;
  }
  const dataEnd = r - 1;
  rows.push([r, [strCell('A' + r, null, 'Totals:'),
    fCell('C' + r, null, 'SUM(C' + dataStart + ':C' + dataEnd + ')', tO), fCell('D' + r, null, 'SUM(D' + dataStart + ':D' + dataEnd + ')', tD),
    fCell('E' + r, null, 'SUM(E' + dataStart + ':E' + dataEnd + ')', tC), fCell('F' + r, null, 'SUM(F' + dataStart + ':F' + dataEnd + ')', tF)]]);
  rows.find((p) => p[0] === 10)[1].push(strCell('G10', null, 'Net assets at ' + mdyy(qtr.end)), fCell('H10', null, nwcRefs.join('+'), ent.nwc));
  rows.find((p) => p[0] === 12)[1].push(strCell('G12', null, 'NWC = current assets (cash, AR net, prepaids, interest reserve, due-froms) less current liabilities (AP, accruals, due-tos, deposits); excludes 25xxx loan payable.'));
  return { rows, loanRef };
}
function tbRows1Col(entName, ent, qtr, style) {
  const rows = [];
  rows.push([1, [strCell('A1', null, 'Company name:'), strCell('B1', null, entName)]]);
  rows.push([2, [strCell('A2', null, 'Report name:'), strCell('B2', null, 'Trial balance report')]]);
  rows.push([3, [strCell('A3', null, 'Reporting Book:'), strCell('B3', null, 'ACCRUAL')]]);
  rows.push([4, [strCell('A4', null, 'Start Date:'), strCell('B4', null, '01/01/' + qtr.year)]]);
  rows.push([5, [strCell('A5', null, 'End Date:'), strCell('B5', null, slashDate(qtr.end))]]);
  rows.push([6, [strCell('A6', null, 'Location:'), strCell('B6', null, 'All (entity total) — per CloudLedger')]]);
  rows.push([7, [strCell('A7', null, 'Account'), strCell('B7', null, 'Account'), strCell('C7', null, 'Closing balance')]]);
  rows.push([8, [strCell('A8', null, 'Number'), strCell('B8', null, 'Name'), strCell('C8', null, 'on ' + slashDate(qtr.end))]]);
  let r = 9; const dataStart = r; const nwcRefs = []; let loanRef = null; let tot = 0;
  for (const c of ent.codes) {
    const a = ent.acc[c];
    rows.push([r, [strCell('A' + r, null, c), strCell('B' + r, null, a.name), numCell('C' + r, null, a.close)]]);
    if (isCA(c) || isCL(c)) nwcRefs.push('C' + r);
    if (isLoan(c) && Math.abs(a.close) > 0.005) loanRef = 'C' + r;
    tot = r2(tot + a.close); r++;
  }
  const dataEnd = r - 1;
  rows.push([r, [strCell('A' + r, null, 'Totals:'), fCell('C' + r, null, 'SUM(C' + dataStart + ':C' + dataEnd + ')', tot)]]);
  if (style === 'srn') {
    rows.find((p) => p[0] === 8)[1].push(strCell('D8', null, 'Net assets ' + mdyy(qtr.end)), fCell('F8', null, '+' + nwcRefs.join('+'), ent.nwc));
    rows.find((p) => p[0] === 10)[1].push(strCell('E10', null, 'NWC = current assets less current liabilities; excludes 25xxx loan payable. Per CloudLedger.'));
  } else {
    rows.find((p) => p[0] === 9)[1].push(strCell('D9', null, 'Net assets at ' + mdyy(qtr.end)), fCell('E9', null, '+' + nwcRefs.join('+'), ent.nwc));
    rows.find((p) => p[0] === 11)[1].push(strCell('D11', null, 'NWC = current assets less current liabilities; excludes 25xxx loan payable. Per CloudLedger.'));
  }
  return { rows, loanRef };
}
function replaceSheetData(xml, rows) {
  const rowXml = rows.sort((a, b) => a[0] - b[0]).map((p) => '<row r="' + p[0] + '">' + p[1].join('') + '</row>').join('');
  let out = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rowXml + '</sheetData>');
  out = out.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, '');
  return out;
}

// -- Investment workbook builder ---------------------------------------------------
async function buildInvestmentWorkbook(templateBuf, data) {
  const { qtr, port, books, fvAdj, booksExact, solve, devTotal, params } = data;
  const zip = await JSZip.loadAsync(templateBuf);
  const P = await sheetMap(zip);
  const need = ['Investment Balance', 'Valuations', 'Carrying Value', 'CLIP TB', 'SRN TB', 'Buna TB', 'Silsbee TB',
    'CLIP', 'SRN', 'Buna', 'Silsbee', 'CLIP Summary', 'Silsbee Summary'];
  for (const n of need) if (!P[n]) throw new Error('investment template missing sheet: ' + n);
  const V = { clip: solve.clip.valuation, silsbee: solve.silsbee.valuation, buna: solve.buna.valuation, srn: solve.srn.valuation };
  const nwc = { clip: port.clip.nwc, silsbee: port.silsbee.nwc, buna: port.buna.nwc, srn: port.srn.nwc };
  const I = { clip: Math.round(solve.clip.proceeds), silsbee: Math.round(solve.silsbee.proceeds), buna: Math.round(solve.buna.proceeds), srn: Math.round(solve.srn.proceeds) };
  const K = { clip: I.clip, silsbee: Math.min(I.silsbee, books.silsbee), buna: Math.min(I.buna, books.buna), srn: Math.min(I.srn, books.srn) };
  const L = { clip: K.clip - books.clip, silsbee: K.silsbee - books.silsbee, buna: K.buna - books.buna, srn: K.srn - books.srn };

  // TB tabs
  const built = {
    clip: tbRows4Col('County Line Industrial Park LLC (CLIP Property Owner)', port.clip, qtr),
    silsbee: tbRows4Col('CLR Silsbee Property Owner LLC', port.silsbee, qtr),
    srn: tbRows1Col('County Line SRN LLC (Sabine River & Northern Railroad)', port.srn, qtr, 'srn'),
    buna: tbRows1Col('CLR Buna Property Owner LLC', port.buna, qtr, 'buna'),
  };
  const tbSheet = { clip: 'CLIP TB', silsbee: 'Silsbee TB', srn: 'SRN TB', buna: 'Buna TB' };
  for (const k of Object.keys(tbSheet)) {
    if (!built[k].loanRef) throw new Error('no 25xxx loan balance found for ' + k);
    zip.file(P[tbSheet[k]], replaceSheetData(await zip.file(P[tbSheet[k]]).async('string'), built[k].rows));
  }

  // entity tabs: valuation link + loan refs
  { let x = await zip.file(P['CLIP']).async('string');
    x = replaceCell(x, 'C3', fCell('C3', styleOf(x, 'C3'), "'Investment Balance'!C5", V.clip));
    x = replaceCell(x, 'C4', fCell('C4', styleOf(x, 'C4'), "'CLIP TB'!" + built.clip.loanRef, port.clip.loanBal));
    zip.file(P['CLIP'], x); }
  { let x = await zip.file(P['Silsbee']).async('string');
    x = replaceCell(x, 'C4', fCell('C4', styleOf(x, 'C4'), "'Silsbee TB'!" + built.silsbee.loanRef, port.silsbee.loanBal));
    zip.file(P['Silsbee'], x); }
  { let x = await zip.file(P['SRN']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), "'SRN TB'!" + built.srn.loanRef, port.srn.loanBal));
    zip.file(P['SRN'], x); }
  { let x = await zip.file(P['Buna']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), "'Buna TB'!" + built.buna.loanRef, port.buna.loanBal));
    zip.file(P['Buna'], x); }

  // Valuations tab
  { let x = await zip.file(P['Valuations']).async('string');
    const G12 = r2(V.clip - devTotal);
    x = replaceCell(x, 'C8', numCell('C8', styleOf(x, 'C8'), excelSerial(qtr.end)));
    x = replaceCell(x, 'D12', numCell('D12', styleOf(x, 'D12'), booksExact.clip));
    x = replaceCell(x, 'D13', numCell('D13', styleOf(x, 'D13'), booksExact.silsbee));
    x = replaceCell(x, 'D14', numCell('D14', styleOf(x, 'D14'), booksExact.buna));
    x = replaceCell(x, 'D15', numCell('D15', styleOf(x, 'D15'), booksExact.srn));
    x = replaceCell(x, 'D16', fCell('D16', styleOf(x, 'D16'), 'SUM(D12:D15)', r2(booksExact.clip + booksExact.silsbee + booksExact.buna + booksExact.srn)));
    if (params.incomeF12 !== null) x = replaceCell(x, 'F12', numCell('F12', styleOf(x, 'F12'), params.incomeF12));
    x = replaceCell(x, 'G12', numCell('G12', styleOf(x, 'G12'), G12));
    x = replaceCell(x, 'H12', numCell('H12', styleOf(x, 'H12'), devTotal));
    x = replaceCell(x, 'I12', fCell('I12', styleOf(x, 'I12'), 'G12', G12));
    x = replaceCell(x, 'J12', fCell('J12', styleOf(x, 'J12'), 'SUM(H12:I12)', V.clip));
    x = replaceCell(x, 'K12', fCell('K12', styleOf(x, 'K12'), 'J12', V.clip));
    for (const [k, row] of [['silsbee', 13], ['buna', 14], ['srn', 15]]) {
      x = replaceCell(x, 'E' + row, numCell('E' + row, styleOf(x, 'E' + row), V[k]));
      x = replaceCell(x, 'G' + row, numCell('G' + row, styleOf(x, 'G' + row), V[k]));
      x = replaceCell(x, 'K' + row, fCell('K' + row, styleOf(x, 'K' + row), 'ROUND(MIN(E' + row + ',G' + row + '),-4)', V[k]));
    }
    for (const [ref, v] of Object.entries(params.propInfo)) {
      if (v !== null) x = replaceCell(x, ref, numCell(ref, styleOf(x, ref), v));
    }
    const note = 'Interim convention: unrealized gain/(loss) held at prior year-end amounts (CLIP = CLRF acct 121012 balance; others 0). '
      + 'Valuations solved from ' + mdyy(qtr.end) + ' net assets and loan balances per CloudLedger: CLIP exactly (dev component per CLIP GL, sales-comparison component is the plug); '
      + 'Silsbee/Buna/SRN kept at prior valuation when proceeds clear book carrying value, otherwise cost and sales approach figures raised so proceeds exceed book by at least $' + BUFFER.toLocaleString() + '. '
      + 'Book carrying values per CLRF GL investment accounts at ' + mdyy(qtr.end) + '.';
    if (/<row r="27">/.test(x)) x = x.replace(/<row r="27">[\s\S]*?<\/row>/, '<row r="27">' + strCell('B27', null, note) + '</row>');
    else x = x.replace('</sheetData>', '<row r="27">' + strCell('B27', null, note) + '</row></sheetData>');
    zip.file(P['Valuations'], x); }

  // Carrying Value
  { let x = await zip.file(P['Carrying Value']).async('string');
    const [y, m, d] = qtr.end.split('-').map(Number);
    x = replaceCell(x, 'A4', strCell('A4', styleOf(x, 'A4'), MONTHS[m - 1] + ' ' + d + ', ' + y));
    x = replaceCell(x, 'J7', numCell('J7', styleOf(x, 'J7'), books.srn));
    x = replaceCell(x, 'J9', numCell('J9', styleOf(x, 'J9'), books.clip));
    x = replaceCell(x, 'J11', numCell('J11', styleOf(x, 'J11'), books.silsbee));
    x = replaceCell(x, 'J13', numCell('J13', styleOf(x, 'J13'), books.buna));
    zip.file(P['Carrying Value'], x); }

  // waterfall liquidation dates
  for (const name of ['CLIP Summary', 'Silsbee Summary']) {
    let x = await zip.file(P[name]).async('string');
    x = replaceCell(x, 'B5', numCell('B5', styleOf(x, 'B5'), excelSerial(qtr.end)));
    zip.file(P[name], x);
  }

  // Investment Balance grid: formulas + fresh caches
  { let x = await zip.file(P['Investment Balance']).async('string');
    x = replaceCell(x, 'C5', fCell('C5', styleOf(x, 'C5'), 'Valuations!K12', V.clip));
    x = replaceCell(x, 'C6', fCell('C6', styleOf(x, 'C6'), 'Valuations!K13', V.silsbee));
    x = replaceCell(x, 'C7', fCell('C7', styleOf(x, 'C7'), 'Valuations!K14', V.buna));
    x = replaceCell(x, 'C8', fCell('C8', styleOf(x, 'C8'), 'Valuations!K15', V.srn));
    const nref = { clip: "'CLIP TB'!H10", silsbee: "'Silsbee TB'!H10", buna: "'Buna TB'!E9", srn: "'SRN TB'!F8" };
    x = replaceCell(x, 'D5', fCell('D5', styleOf(x, 'D5'), nref.clip, nwc.clip));
    x = replaceCell(x, 'D6', fCell('D6', styleOf(x, 'D6'), nref.silsbee, nwc.silsbee));
    x = replaceCell(x, 'D7', fCell('D7', styleOf(x, 'D7'), nref.buna, nwc.buna));
    x = replaceCell(x, 'D8', fCell('D8', styleOf(x, 'D8'), nref.srn, nwc.srn));
    x = replaceCell(x, 'E5', fCell('E5', styleOf(x, 'E5'), 'SUM(C5:D5)', r2(V.clip + nwc.clip)));
    x = replaceCell(x, 'E6', fCell('E6', styleOf(x, 'E6'), 'SUM(C6:D6)', r2(V.silsbee + nwc.silsbee)));
    x = replaceCell(x, 'E7', fCell('E7', styleOf(x, 'E7'), 'SUM(C7:D7)', r2(V.buna + nwc.buna)));
    x = replaceCell(x, 'E8', fCell('E8', styleOf(x, 'E8'), 'SUM(C8:D8)', r2(V.srn + nwc.srn)));
    x = replaceCell(x, 'F5', fCell('F5', styleOf(x, 'F5'), nref.clip, nwc.clip));
    x = replaceCell(x, 'F6', fCell('F6', styleOf(x, 'F6'), nref.silsbee, nwc.silsbee));
    x = replaceCell(x, 'F7', fCell('F7', styleOf(x, 'F7'), nref.buna, nwc.buna));
    x = replaceCell(x, 'F8', fCell('F8', styleOf(x, 'F8'), nref.srn, nwc.srn));
    x = replaceCell(x, 'I5', fCell('I5', styleOf(x, 'I5'), 'ROUND(CLIP!C13,0)', I.clip));
    x = replaceCell(x, 'I6', fCell('I6', styleOf(x, 'I6'), 'ROUND(IF(H6="transitional","N/A",Silsbee!C13),0)', I.silsbee));
    x = replaceCell(x, 'I7', fCell('I7', styleOf(x, 'I7'), 'ROUND(IF(H7="transitional","N/A",Buna!C10),0)', I.buna));
    x = replaceCell(x, 'I8', fCell('I8', styleOf(x, 'I8'), 'ROUND(IF(H8="transitional","N/A",SRN!C10),0)', I.srn));
    x = replaceCell(x, 'J5', fCell('J5', styleOf(x, 'J5'), "'Carrying Value'!J9", books.clip));
    x = replaceCell(x, 'J6', fCell('J6', styleOf(x, 'J6'), "'Carrying Value'!J11", books.silsbee));
    x = replaceCell(x, 'J7', fCell('J7', styleOf(x, 'J7'), "'Carrying Value'!J13", books.buna));
    x = replaceCell(x, 'J8', fCell('J8', styleOf(x, 'J8'), "'Carrying Value'!J7", books.srn));
    x = replaceCell(x, 'K5', fCell('K5', styleOf(x, 'K5'), 'I5', K.clip));
    x = replaceCell(x, 'K6', fCell('K6', styleOf(x, 'K6'), 'IF(I6>J6,J6,I6)', K.silsbee));
    x = replaceCell(x, 'K7', fCell('K7', styleOf(x, 'K7'), 'IF(I7>J7,J7,I7)', K.buna));
    x = replaceCell(x, 'K8', fCell('K8', styleOf(x, 'K8'), 'IF(I8>J8,J8,I8)', K.srn));
    x = replaceCell(x, 'L5', fCell('L5', styleOf(x, 'L5'), 'K5-J5', L.clip));
    x = replaceCell(x, 'L6', fCell('L6', styleOf(x, 'L6'), 'K6-J6', L.silsbee));
    x = replaceCell(x, 'L7', fCell('L7', styleOf(x, 'L7'), 'K7-J7', L.buna));
    x = replaceCell(x, 'L8', fCell('L8', styleOf(x, 'L8'), 'K8-J8', L.srn));
    x = replaceCell(x, 'C9', fCell('C9', styleOf(x, 'C9'), 'SUM(C5:C8)', r2(V.clip + V.silsbee + V.buna + V.srn)));
    x = replaceCell(x, 'D9', fCell('D9', styleOf(x, 'D9'), 'SUM(D5:D8)', r2(nwc.clip + nwc.silsbee + nwc.buna + nwc.srn)));
    x = replaceCell(x, 'E9', fCell('E9', styleOf(x, 'E9'), 'SUM(E5:E8)', r2(V.clip + V.silsbee + V.buna + V.srn + nwc.clip + nwc.silsbee + nwc.buna + nwc.srn)));
    x = replaceCell(x, 'J9', fCell('J9', styleOf(x, 'J9'), 'SUM(J5:J8)', books.clip + books.silsbee + books.buna + books.srn));
    x = replaceCell(x, 'K9', fCell('K9', styleOf(x, 'K9'), 'SUM(K5:K8)', K.clip + K.silsbee + K.buna + K.srn));
    x = replaceCell(x, 'L9', fCell('L9', styleOf(x, 'L9'), 'SUM(L5:L8)', L.clip + L.silsbee + L.buna + L.srn));
    zip.file(P['Investment Balance'], x); }

  // Loan Balances_DO NOT USE: repoint loan cells if the tab exists
  if (P['Loan Balances_DO NOT USE']) {
    let x = await zip.file(P['Loan Balances_DO NOT USE']).async('string');
    const tryRep = (ref, f, v) => { try { x = replaceCell(x, ref, fCell(ref, styleOf(x, ref), f, v)); } catch (e) { /* cell absent */ } };
    tryRep('W7', "-'CLIP TB'!" + built.clip.loanRef, -port.clip.loanBal);
    tryRep('W8', "-'Silsbee TB'!" + built.silsbee.loanRef, -port.silsbee.loanBal);
    tryRep('W9', "-'SRN TB'!" + built.srn.loanRef, -port.srn.loanBal);
    tryRep('W10', "-'Buna TB'!" + built.buna.loanRef, -port.buna.loanBal);
    zip.file(P['Loan Balances_DO NOT USE'], x);
  }

  // workbook-level: full recalc on load; drop the (now stale) calc chain
  { let wb = await zip.file('xl/workbook.xml').async('string');
    if (/<calcPr[^>]*fullCalcOnLoad=/.test(wb)) wb = wb.replace(/(<calcPr[^>]*fullCalcOnLoad=")[^"]*(")/, '$11$2');
    else if (/<calcPr/.test(wb)) wb = wb.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
    else wb = wb.replace(/<\/workbook>/, '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    zip.file('xl/workbook.xml', wb);
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      const ct = await zip.file('[Content_Types].xml').async('string');
      zip.file('[Content_Types].xml', ct.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
      const wr = await zip.file('xl/_rels/workbook.xml.rels').async('string');
      zip.file('xl/_rels/workbook.xml.rels', wr.replace(/<Relationship [^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
    } }

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buf, invBalance: { I, J: books, K, L } };
}

// -- Template lookup / save --------------------------------------------------------
function findFileIn(ctx, eid, folder, likeName) {
  const { db, workpapersDir } = ctx;
  const row = db.prepare('SELECT * FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name LIKE ? ORDER BY id DESC LIMIT 1')
    .get(eid, folder, likeName);
  if (!row) return null;
  return Object.assign({}, row, { abs_path: path.join(workpapersDir, String(eid), row.stored_filename) });
}
function findLatestLike(ctx, eid, folderLike, likeName) {
  const { db, workpapersDir } = ctx;
  const row = db.prepare('SELECT * FROM entity_files WHERE entity_id=? AND folder_path LIKE ? AND original_name LIKE ? ORDER BY id DESC LIMIT 1')
    .get(eid, folderLike, likeName);
  if (!row) return null;
  return Object.assign({}, row, { abs_path: path.join(workpapersDir, String(eid), row.stored_filename) });
}
function saveFile(ctx, eid, folder, original, buf, who) {
  const { db, workpapersDir } = ctx;
  const parts = folder.split('/');
  const ins = db.prepare("INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id=? AND folder_path=? AND original_name=?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id=?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  const ins2 = db.prepare("INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { file_id: Number(ins2.lastInsertRowid), folder_path: folder, original_name: original, replaced: prior.length };
}

// -- Route ---------------------------------------------------------------------------
function registerInvValRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  const val = require('./valuation');

  app.post('/api/workpapers/investment-valuation/:entity_id/generate', auth,
    requireEntityAccess('entity_id'), requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        if (eid !== CLRF) return res.status(400).json({ error: 'Investment & Valuation is a CLRF (entity 40) workpaper.' });
        const qtr = val.resolveQuarter((req.body && req.body.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const prior = qtr.q === 1 ? { quarter: 'Q4', year: String(Number(qtr.year) - 1) }
          : { quarter: 'Q' + (qtr.q - 1), year: qtr.year };
        const priorIV = 'Workpapers/Investment & Valuation/' + prior.quarter + ' ' + prior.year;

        // 1) investment template: prior quarter's investment workpaper
        const invTpl = findFileIn(ctx, eid, priorIV, 'CLRF Investment Balance%.xlsx')
          || findLatestLike(ctx, eid, 'Workpapers/Investment & Valuation/%', 'CLRF Investment Balance%.xlsx');
        if (!invTpl || !fs.existsSync(invTpl.abs_path)) {
          return res.status(400).json({ error: "No prior investment workpaper found. Upload the prior quarter's "
            + '"CLRF Investment Balance ..." file under Workpapers/Investment & Valuation/' + prior.quarter + ' ' + prior.year + '.' });
        }
        const invTplBuf = fs.readFileSync(invTpl.abs_path);

        // 2) gather live data
        const port = gatherPortfolio(ctx, qtr);
        const clrfRows = ctx.computeBalances(CLRF, { as_of: qtr.end });
        const clrfBy = {}; for (const x of clrfRows) clrfBy[String(x.code)] = x;
        const costOf = (code) => { const rw = clrfBy[code]; return rw ? r2(dpOf(rw)) : 0; };
        const booksExact = { clip: costOf('121011'), silsbee: costOf('121041'), buna: costOf('121021'), srn: costOf('121031') };
        const books = { clip: Math.round(booksExact.clip), silsbee: Math.round(booksExact.silsbee), buna: Math.round(booksExact.buna), srn: Math.round(booksExact.srn) };
        const fvAdj = { clip: Math.round(costOf('121012')), silsbee: Math.round(costOf('121042')), buna: Math.round(costOf('121022')), srn: Math.round(costOf('121032')) };
        const glVal = val.gatherGl(ctx, qtr); // CLRF TB tab + CLIP dev costs for the valuation workbook
        const devTotal = r2(glVal.dev.ltiTotal + glVal.dev.oaTotal);

        // 3) parse model params from the template, solve valuations
        const tplZip = await JSZip.loadAsync(invTplBuf);
        const tplP = await sheetMap(tplZip);
        const params = await parseModelParams(tplZip, tplP);
        const model = makeModel(params, excelSerial(qtr.end));
        const solve = solveValuations(model, port, books, fvAdj, params.priorVals);

        // 4) build + save the investment workbook
        const inv = await buildInvestmentWorkbook(invTplBuf, { qtr, port, books, booksExact, fvAdj, solve, devTotal, params });
        const folder = IV_FOLDER(qtr);
        const invName = 'CLRF Investment Balance ' + shortDate(qtr.end) + '.xlsx';
        const savedInv = saveFile(ctx, eid, folder, invName, inv.buf, who);

        // 5) valuation workbook against the solved targets (Summary matches the
        //    investment workpaper's Valuations tab exactly)
        const valTpl = val.findTemplate(ctx, eid, qtr);
        if (!valTpl || !fs.existsSync(valTpl.abs_path)) {
          return res.status(400).json({ error: 'No prior valuation workbook found under Workpapers/Investment & Valuation or Workpapers/Valuation.' });
        }
        const targets = {
          clipJ12: solve.clip.valuation,
          approaches: {
            silsbee: solve.silsbee.changed ? solve.silsbee.valuation : null,
            buna: solve.buna.changed ? solve.buna.valuation : null,
            srn: solve.srn.changed ? solve.srn.valuation : null,
          },
        };
        const valResult = await val.transform(fs.readFileSync(valTpl.abs_path), glVal, qtr, targets);
        const valName = val.valFileName(qtr);
        const savedVal = saveFile(ctx, eid, folder, valName, valResult.buf, who);

        res.json({
          quarter: qtr.label,
          folder: folder,
          investment: Object.assign({ template_from: invTpl.folder_path + '/' + invTpl.original_name }, savedInv, inv.invBalance),
          valuation: Object.assign({ template_from: valTpl.source_folder + '/' + valTpl.original_name }, savedVal, valResult.summary),
          solve: {
            clip: { valuation: solve.clip.valuation, proceeds: solve.clip.proceeds, promote: solve.clip.promote, frozen_gain: fvAdj.clip },
            silsbee: { valuation: solve.silsbee.valuation, proceeds: solve.silsbee.proceeds, changed: solve.silsbee.changed },
            buna: { valuation: solve.buna.valuation, proceeds: solve.buna.proceeds, changed: solve.buna.changed },
            srn: { valuation: solve.srn.valuation, proceeds: solve.srn.proceeds, changed: solve.srn.changed },
          },
          nwc: { clip: port.clip.nwc, silsbee: port.silsbee.nwc, buna: port.buna.nwc, srn: port.srn.nwc },
          loans: { clip: port.clip.loanBal, silsbee: port.silsbee.loanBal, buna: port.buna.loanBal, srn: port.srn.loanBal },
          unrealized: inv.invBalance.L,
        });
      } catch (e) {
        console.error('investment-valuation failed:', e);
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = {
  registerInvValRoutes: registerInvValRoutes,
  gatherPortfolio: gatherPortfolio,
  solveValuations: solveValuations,
  makeModel: makeModel,
  parseModelParams: parseModelParams,
  buildInvestmentWorkbook: buildInvestmentWorkbook,
  IV_FOLDER: IV_FOLDER,
};
