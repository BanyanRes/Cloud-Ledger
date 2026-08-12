// Insurance Allocation engine (pure module).
// Given a carrier health-insurance billing invoice and a consolidated billing
// report (both .xlsx buffers), allocates the monthly premium across the
// commonly-owned entities and builds a fully FORMULA-LINKED workpaper workbook:
// source tabs (Invoice, Consolidated Billing) carry the raw uploaded numbers,
// Member Detail links to them by formula, a by-entity pivot summarizes the
// detail, and the Allocation tab references the pivot. Almost nothing is a
// hard-coded amount — every dollar traces back to a source tab.
//
// Rules (confirmed by Jimmy, 2026-08-12):
//  - Employee share (monthly) = SUM(Employee Cost Per Pay Period across ALL
//    benefit lines, HSA included) x Pay Periods / 12.
//  - Employer share = Premium - Employee share.
//  - Entity = the employee's Department in the consolidated billing, matched by
//    Last + First name, subject to an override table.
//  - Eligibility-change amounts are booked 100% to the employer, tagged to the
//    same entity as the subscriber.
//  - Override table (defaults): MARTINEZ, JONATHAN -> SRN;
//    employer-paid-in-full (EE = 0): BROSSEAU, BENJAMIN.

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const DEPT_TO_CODE = [
  [/county\s*line\s*rail\s*op/i, 'CLRO'],
  [/sabine\s*river/i,            'SRN'],
  [/turn\s*key\s*rail/i,         'TKR'],
  [/banyan\s*res/i,              'BR'],
];
const CODE_META = {
  CLRO: { name: 'County Line Rail Operations' },
  SRN:  { name: 'Sabine River and Northern Railroad' },
  BR:   { name: 'Banyan Residential' },
  TKR:  { name: 'TurnKey Rail' },
};
const CODE_ORDER = ['CLRO', 'SRN', 'BR', 'TKR'];

const DEFAULT_OVERRIDES = {
  entity: { 'MARTINEZ, JONATHAN': 'SRN' },
  employerPaid: ['BROSSEAU, BENJAMIN'],
};

function deptToCode(dept) {
  const s = String(dept || '');
  for (const [re, code] of DEPT_TO_CODE) if (re.test(s)) return code;
  return null;
}
function nameKey(last, first) {
  const l = String(last || '').trim().toUpperCase();
  const f = String(first || '').trim().split(/\s+/)[0].toUpperCase().replace(/[.,]$/, '');
  return l && f ? `${l}, ${f}` : (l || '');
}
function nameKeyFromFull(full) {
  const s = String(full || '').trim();
  const comma = s.indexOf(',');
  if (comma < 0) return s.toUpperCase();
  return nameKey(s.slice(0, comma), s.slice(comma + 1));
}
function money(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
  return Number.isFinite(n) ? n : 0;
}
function readSheets(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const out = {};
  for (const name of wb.SheetNames) out[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
  return out;
}
function findHeader(rows, required) {
  for (let r = 0; r < Math.min(rows.length, 40); r++) {
    const row = rows[r] || [];
    const map = {};
    row.forEach((cell, c) => { const key = String(cell == null ? '' : cell).replace(/\s+/g, ' ').trim().toLowerCase(); if (key) map[key] = c; });
    if (required.every(req => req.some(alias => map[alias] != null))) {
      const pick = req => { for (const a of req) if (map[a] != null) return map[a]; return -1; };
      return { rowIndex: r, pick, map };
    }
  }
  return null;
}
const round2 = n => Math.round((Number(n) + Number.EPSILON) * 100) / 100;

// ── Consolidated billing → { dept, ee, rows } (rows kept for the source tab). ──
function parseConsolidated(sheets) {
  let rows = null, hdr = null;
  const names = Object.keys(sheets).sort((a, b) => (/summary/i.test(b) ? 1 : 0) - (/summary/i.test(a) ? 1 : 0));
  for (const name of names) {
    const rws = sheets[name];
    const h = findHeader(rws, [['last name'], ['department'], ['plan type'], ['employee cost per pay period'], ['pay periods']]);
    if (h) { rows = rws; hdr = h; break; }
  }
  if (!rows) throw new Error('Consolidated billing: could not find the long-format summary sheet (needs Department, Plan Type, Employee Cost Per Pay Period, Pay Periods columns).');
  const cLast = hdr.pick(['last name']), cFirst = hdr.pick(['first name']), cDept = hdr.pick(['department']);
  const cPlan = hdr.pick(['plan type']), cEcpp = hdr.pick(['employee cost per pay period']), cPp = hdr.pick(['pay periods']);
  const dept = {}, ee = {}, outRows = [];
  for (let r = hdr.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const last = row[cLast], first = cFirst >= 0 ? row[cFirst] : '';
    if (last == null || String(last).trim() === '') continue;
    const key = nameKey(last, first);
    const code = deptToCode(row[cDept]);
    if (code && dept[key] == null) dept[key] = code;
    const pp = money(row[cPp]), ecpp = money(row[cEcpp]);
    ee[key] = (ee[key] || 0) + ecpp * (pp ? pp / 12 : 0);
    outRows.push({ last: String(last).trim(), first: String(first || '').trim(), dept: String(row[cDept] || '').trim(),
      code: code || '', plan: String(row[cPlan] || '').trim(), pp, ecpp, key });
  }
  return { dept, ee, rows: outRows };
}

// ── Carrier invoice → { members, elig, invMeta } (full columns kept). ──
function parseInvoice(sheets) {
  let memRows = null, memHdr = null, invMeta = {};
  for (const name of Object.keys(sheets)) {
    const rws = sheets[name];
    const h = findHeader(rws, [['subscriber name'], ['premium amount']]);
    if (h) {
      memRows = rws; memHdr = h;
      for (let r = 0; r < h.rowIndex; r++) {
        const row = rws[r] || [];
        for (let c = 0; c < row.length - 1; c++) {
          const lab = String(row[c] == null ? '' : row[c]).toLowerCase();
          if (/invoice\s*(#|no\.?|number)/.test(lab) && row[c + 1] != null) invMeta.invoice = String(row[c + 1]);
          if (lab.includes('billing period') && row[c + 1] != null) invMeta.period = String(row[c + 1]);
        }
      }
      break;
    }
  }
  if (!memRows) throw new Error('Invoice: could not find a sheet with Subscriber Name + Premium Amount columns.');
  const cId = memHdr.pick(['member id no.', 'member id']), cName = memHdr.pick(['subscriber name']);
  const cProd = memHdr.pick(['product']), cType = memHdr.pick(['contract type']), cCov = memHdr.pick(['number covered']);
  const cSub = memHdr.pick(['subscriber amount']), cDep = memHdr.pick(['dependent amount']), cPrem = memHdr.pick(['premium amount']);
  const members = [];
  for (let r = memHdr.rowIndex + 1; r < memRows.length; r++) {
    const row = memRows[r] || [];
    if ((row || []).some(cell => /^(sub)?total$|rate change legend/i.test(String(cell == null ? '' : cell).trim()))) break;
    const nm = row[cName];
    if (nm == null || String(nm).trim() === '') continue;
    const prem = money(row[cPrem]);
    if (!prem) continue;
    members.push({ memberId: cId >= 0 ? String(row[cId] || '') : '', name: String(nm).trim(), key: nameKeyFromFull(nm),
      product: cProd >= 0 ? String(row[cProd] || '') : '', contractType: cType >= 0 ? String(row[cType] || '') : '',
      numCovered: cCov >= 0 ? (row[cCov] == null ? '' : row[cCov]) : '',
      subscriberAmt: money(row[cSub]), dependentAmt: money(row[cDep]), premium: prem });
  }
  const elig = [];
  for (const name of Object.keys(sheets)) {
    const rws = sheets[name];
    const h = findHeader(rws, [['subscriber name'], ['premium amount'], ['change code']]);
    if (!h) continue;
    const eId = h.pick(['member id no.', 'member id']), en = h.pick(['subscriber name']), eProd = h.pick(['product']);
    const eEff = h.pick(['effective date']), eCode = h.pick(['change code']), ep = h.pick(['premium amount']);
    for (let r = h.rowIndex + 1; r < rws.length; r++) {
      const row = rws[r] || [];
      const nm = row[en];
      if (nm == null || String(nm).trim() === '') continue;
      if (/^subtotal|^total\b/i.test(String(nm).trim())) break;
      elig.push({ memberId: eId >= 0 ? String(row[eId] || '') : '', name: String(nm).trim(), key: nameKeyFromFull(nm),
        product: eProd >= 0 ? String(row[eProd] || '') : '', effDate: eEff >= 0 ? String(row[eEff] || '') : '',
        changeCode: eCode >= 0 ? String(row[eCode] || '') : '', amount: money(row[ep]) });
    }
    break;
  }
  return { members, elig, invMeta };
}

// ── Compute the allocation. Per-entity employer is BASE (excludes eligibility);
//    eligibility is tracked separately (eligByEntity) and shown on its own line. ──
function computeAllocation({ invoiceBuf, consolidatedBuf, overrides }) {
  const ov = {
    entity: { ...DEFAULT_OVERRIDES.entity, ...((overrides && overrides.entity) || {}) },
    employerPaid: new Set([...DEFAULT_OVERRIDES.employerPaid, ...((overrides && overrides.employerPaid) || [])].map(s => s.toUpperCase())),
  };
  const cons = parseConsolidated(readSheets(consolidatedBuf));
  const { members, elig, invMeta } = parseInvoice(readSheets(invoiceBuf));

  const ent = {}; const bucket = c => (ent[c] || (ent[c] = { premium: 0, employer: 0, employee: 0 }));
  const eligByEntity = {};
  const flags = [], memberRows = [], unmatched = [];

  for (const m of members) {
    const forcedEntity = ov.entity[m.key];
    const deptEntity = cons.dept[m.key];
    const code = forcedEntity || deptEntity || null;
    const employerPaid = ov.employerPaid.has(m.key);
    const ee = code && employerPaid ? 0 : (code ? (cons.ee[m.key] || 0) : 0);
    const er = m.premium - ee;
    const note = forcedEntity && deptEntity && forcedEntity !== deptEntity ? `reclassified from ${deptEntity}` : (employerPaid ? 'employer-paid in full' : '');
    memberRows.push({ memberId: m.memberId, name: m.name, key: m.key, entity: code || '', premium: m.premium, ee, er,
      employerPaid, product: m.product, contractType: m.contractType, numCovered: m.numCovered,
      subscriberAmt: m.subscriberAmt, dependentAmt: m.dependentAmt, note });
    if (!code) { unmatched.push(m.name); continue; }
    const b = bucket(code); b.premium += m.premium; b.employer += er; b.employee += ee;
    if (forcedEntity && deptEntity && forcedEntity !== deptEntity) flags.push({ type: 'reclass', name: m.name, from: deptEntity, to: forcedEntity, premium: m.premium });
    if (employerPaid) flags.push({ type: 'employerPaid', name: m.name, entity: code, premium: m.premium });
  }

  const eligRows = [];
  let eligibilityTotal = 0;
  for (const e of elig) {
    const code = ov.entity[e.key] || cons.dept[e.key] || null;
    eligibilityTotal += e.amount;
    if (code) eligByEntity[code] = (eligByEntity[code] || 0) + e.amount;
    eligRows.push({ ...e, entity: code || '' });
    flags.push({ type: 'eligibility', name: e.name, entity: code, amount: e.amount });
  }

  const codes = [...CODE_ORDER.filter(c => ent[c]), ...Object.keys(ent).filter(c => !CODE_ORDER.includes(c))];
  const entities = codes.map(c => ({ code: c, name: (CODE_META[c] && CODE_META[c].name) || c,
    premium: round2(ent[c].premium), employer: round2(ent[c].employer), employee: round2(ent[c].employee),
    eligibility: round2(eligByEntity[c] || 0) }));

  const premiumSubtotal = round2(members.reduce((s, m) => s + m.premium, 0));
  const employeeTotal = round2(memberRows.reduce((s, m) => s + (m.entity ? m.ee : 0), 0));
  const employerBase = round2(premiumSubtotal - employeeTotal);

  return {
    period: (invMeta.period || '').trim(), invoice: (invMeta.invoice || '').trim(),
    entities, codes,
    subtotal: { premium: premiumSubtotal, employer: employerBase, employee: employeeTotal },
    eligibility: eligRows, eligibilityTotal: round2(eligibilityTotal),
    employerTotal: round2(employerBase + eligibilityTotal), employeeTotal, totalBilled: round2(premiumSubtotal + eligibilityTotal),
    subscriberCount: members.length, flags, members: memberRows, consolidated: cons.rows, unmatched,
    reconciled: unmatched.length === 0,
  };
}

// ── Build the FORMULA-LINKED workpaper workbook. ──
async function buildAllocationWorkbook(result, opts = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  wb.calcProperties.fullCalcOnLoad = true; // force Excel/LibreOffice to recompute on open
  const NAVY = 'FF0D1B2E', HEAD = 'FFF2F4F7', GREY = 'FF8A90A0', MONEY = '#,##0.00;(#,##0.00)';
  const bold = { bold: true };
  const styleHead = row => row.eachCell(c => { c.font = { bold: true, size: 10, color: { argb: GREY } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD } }; });
  const money = cell => { cell.numFmt = MONEY; cell.alignment = { horizontal: 'right' }; };

  // Create tabs in display order: Allocation, Member Detail, Invoice, Eligibility,
  // Consolidated Billing. They are populated below in dependency order (sources →
  // detail → pivot → allocation); cross-sheet formulas are strings, so build order
  // and tab order are independent.
  const al = wb.addWorksheet('Allocation', { views: [{ showGridLines: false }] });
  const md = wb.addWorksheet('Member Detail', { views: [{ showGridLines: false }] });
  const inv = wb.addWorksheet('Invoice', { views: [{ showGridLines: false }] });
  const el = wb.addWorksheet('Eligibility', { views: [{ showGridLines: false }] });
  const con = wb.addWorksheet('Consolidated Billing', { views: [{ showGridLines: false }] });

  // ---- Source tab: Invoice (membership detail) ----
  inv.columns = [{ width: 14 }, { width: 26 }, { width: 12 }, { width: 13 }, { width: 11 }, { width: 15 }, { width: 15 }, { width: 15 }];
  inv.addRow(['Carrier Billing Invoice — Membership Detail']).font = { bold: true, size: 14, color: { argb: NAVY } };
  const im = []; if (result.invoice) im.push('Invoice #: ' + result.invoice); if (result.period) im.push('Billing Period: ' + result.period);
  inv.addRow([im.join('     ')]).font = { color: { argb: GREY }, size: 10 };
  inv.addRow([]);
  const invHdr = inv.addRow(['Member ID', 'Subscriber Name', 'Product', 'Contract Type', '# Covered', 'Subscriber Amount', 'Dependent Amount', 'Premium Amount']);
  styleHead(invHdr);
  const invStart = invHdr.number + 1;
  for (const m of result.members) {
    const r = inv.addRow([m.memberId, m.name, m.product, m.contractType, m.numCovered, round2(m.subscriberAmt), round2(m.dependentAmt), round2(m.premium)]);
    [6, 7, 8].forEach(i => money(r.getCell(i)));
  }
  const invEnd = invStart + result.members.length - 1;
  const invSub = inv.addRow(['', 'Subtotal', '', '', '', '', '', { formula: `SUM(H${invStart}:H${invEnd})` }]);
  invSub.font = bold; money(invSub.getCell(8)); invSub.getCell(8).border = { top: { style: 'thin' } };
  if (result.eligibility.length) {
    const invElig = inv.addRow(['', 'Eligibility changes', '', '', '', '', '', { formula: `SUM(Eligibility!$F:$F)` }]);
    money(invElig.getCell(8)); invElig.getCell(8).font = { color: { argb: GREY } }; invElig.getCell(2).font = { color: { argb: GREY } };
    const invTot = inv.addRow(['', 'Total invoice amount', '', '', '', '', '', { formula: `H${invSub.number}+H${invElig.number}` }]);
    invTot.font = bold; money(invTot.getCell(8)); invTot.getCell(8).border = { top: { style: 'thin' }, bottom: { style: 'double' } };
  }

  // ---- Source tab: Eligibility ----
  el.columns = [{ width: 14 }, { width: 26 }, { width: 12 }, { width: 13 }, { width: 13 }, { width: 15 }, { width: 8 }];
  el.addRow(['Eligibility Changes']).font = { bold: true, size: 14, color: { argb: NAVY } };
  el.addRow([]);
  const elHdr = el.addRow(['Member ID', 'Subscriber Name', 'Product', 'Effective Date', 'Change Code', 'Amount', 'Entity']);
  styleHead(elHdr);
  const elStart = elHdr.number + 1;
  for (const e of result.eligibility) {
    const r = el.addRow([e.memberId, e.name, e.product, e.effDate, e.changeCode, round2(e.amount), e.entity]);
    money(r.getCell(6));
  }
  const elEnd = elStart + result.eligibility.length - 1;
  const hasElig = result.eligibility.length > 0;

  // ---- Source tab: Consolidated Billing ----
  con.columns = [{ width: 16 }, { width: 14 }, { width: 30 }, { width: 8 }, { width: 20 }, { width: 12 }, { width: 16 }, { width: 14 }, { width: 24 }];
  con.addRow(['Consolidated Billing — Employee Deductions']).font = { bold: true, size: 14, color: { argb: NAVY } };
  con.addRow(['Monthly EE = Employee Cost Per Pay Period × Pay Periods ÷ 12']).font = { color: { argb: GREY }, size: 10 };
  con.addRow([]);
  const conHdr = con.addRow(['Last Name', 'First Name', 'Department', 'Entity', 'Plan Type', 'Pay Periods', 'Employee Cost / Pay Period', 'Monthly EE', 'Name Key']);
  styleHead(conHdr);
  const conStart = conHdr.number + 1;
  for (const c of result.consolidated) {
    const r = con.addRow([c.last, c.first, c.dept, c.code, c.plan, c.pp, round2(c.ecpp), null, c.key]);
    money(r.getCell(7));
    r.getCell(8).value = { formula: `G${r.number}*F${r.number}/12` }; money(r.getCell(8)); // Monthly EE links to source
  }
  const conEnd = conStart + result.consolidated.length - 1;

  // ---- Member Detail (formula-linked) + by-entity pivot ----
  md.columns = [{ width: 26 }, { width: 8 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 24 }, { width: 18 },
                { width: 3 }, { width: 10 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }];
  md.addRow(['Member Detail']).font = { bold: true, size: 14, color: { argb: NAVY } };
  md.addRow(['Premium linked to Invoice; Employee linked to Consolidated Billing; Employer = Premium − Employee.']).font = { color: { argb: GREY }, size: 10 };
  md.addRow([]);
  const mdHdr = md.addRow(['Subscriber', 'Entity', 'Premium', 'Employee', 'Employer', 'Note', 'Name Key']);
  styleHead(mdHdr);
  const mdStart = mdHdr.number + 1;
  for (const m of result.members) {
    const r = md.addRow([m.name, m.entity, null, null, null, m.note, m.key]);
    const rn = r.number;
    r.getCell(3).value = { formula: `SUMIFS(Invoice!$H:$H,Invoice!$B:$B,$A${rn})` };          // Premium
    r.getCell(4).value = m.employerPaid ? 0 : { formula: `SUMIFS('Consolidated Billing'!$H:$H,'Consolidated Billing'!$I:$I,$G${rn})` }; // Employee
    r.getCell(5).value = { formula: `C${rn}-D${rn}` };                                          // Employer
    [3, 4, 5].forEach(i => money(r.getCell(i)));
  }
  const mdEnd = mdStart + result.members.length - 1;

  // Pivot block (columns I..M), aligned near the top of the same sheet.
  const pvTitleRow = md.getRow(3); pvTitleRow.getCell(9).value = 'By-Entity Pivot'; pvTitleRow.getCell(9).font = { bold: true, color: { argb: NAVY } };
  const pvHdr = md.getRow(mdHdr.number);
  ['Entity', 'Premium', 'Employer', 'Employee', 'Eligibility'].forEach((h, i) => {
    const c = pvHdr.getCell(9 + i); c.value = h; c.font = { bold: true, size: 10, color: { argb: GREY } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD } };
  });
  const pvStart = mdStart;
  const pivotRowByCode = {};
  result.codes.forEach((code, i) => {
    const rn = pvStart + i;
    const row = md.getRow(rn);
    row.getCell(9).value = code;
    row.getCell(10).value = { formula: `SUMIF($B:$B,$I${rn},$C:$C)` };  // Premium (base)
    row.getCell(11).value = { formula: `SUMIF($B:$B,$I${rn},$E:$E)` };  // Employer (base)
    row.getCell(12).value = { formula: `SUMIF($B:$B,$I${rn},$D:$D)` };  // Employee
    row.getCell(13).value = hasElig ? { formula: `SUMIF(Eligibility!$G:$G,$I${rn},Eligibility!$F:$F)` } : 0; // Eligibility
    [10, 11, 12, 13].forEach(c => money(row.getCell(c)));
    pivotRowByCode[code] = rn;
  });
  const pvEnd = pvStart + result.codes.length - 1;
  const pvTotRn = pvEnd + 1;
  const pvTot = md.getRow(pvTotRn);
  pvTot.getCell(9).value = 'Total'; pvTot.getCell(9).font = bold;
  [10, 11, 12, 13].forEach(c => { const L = String.fromCharCode(64 + c); pvTot.getCell(c).value = { formula: `SUM(${L}${pvStart}:${L}${pvEnd})` }; money(pvTot.getCell(c)); pvTot.getCell(c).font = bold; });

  // ---- Allocation (references the pivot) ----
  al.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }];
  al.addRow([opts.title || 'Health Insurance Allocation']).font = { bold: true, size: 15, color: { argb: NAVY } };
  const am = []; if (opts.entityName) am.push('Entity: ' + opts.entityName); if (result.period) am.push('Billing Period: ' + result.period);
  if (result.invoice) am.push('Invoice #: ' + result.invoice); am.push('Subscribers: ' + result.subscriberCount);
  al.addRow([am.join('     ')]).font = { color: { argb: GREY }, size: 10 };
  al.addRow([]);
  const alHdr = al.addRow(['Entity', 'Premium', 'Employer', 'Employee']); styleHead(alHdr); alHdr.getCell(1).alignment = { horizontal: 'left' };
  const entRowStart = alHdr.number + 1;
  result.entities.forEach((e) => {
    const pr = pivotRowByCode[e.code];
    const r = al.addRow([`${e.name}  ·  ${e.code}`, null, null, null]);
    r.getCell(2).value = { formula: `'Member Detail'!J${pr}` };
    r.getCell(3).value = { formula: `'Member Detail'!K${pr}` };
    r.getCell(4).value = { formula: `'Member Detail'!L${pr}` };
    [2, 3, 4].forEach(i => money(r.getCell(i)));
  });
  const entRowEnd = entRowStart + result.entities.length - 1;
  const sub = al.addRow(['Subtotal', { formula: `SUM(B${entRowStart}:B${entRowEnd})` }, { formula: `SUM(C${entRowStart}:C${entRowEnd})` }, { formula: `SUM(D${entRowStart}:D${entRowEnd})` }]);
  sub.font = bold; [2, 3, 4].forEach(i => money(sub.getCell(i))); sub.eachCell(c => { c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } }; });
  const subRn = sub.number;
  let totRn;
  if (hasElig) {
    const eRow = al.addRow(['Eligibility change (100% employer)',
      { formula: `'Member Detail'!M${pvTotRn}` }, { formula: `'Member Detail'!M${pvTotRn}` }, null]);
    eRow.font = { color: { argb: GREY } }; [2, 3].forEach(i => money(eRow.getCell(i)));
    const eRn = eRow.number;
    const tot = al.addRow(['Total billed', { formula: `B${subRn}+B${eRn}` }, { formula: `C${subRn}+C${eRn}` }, { formula: `D${subRn}` }]);
    totRn = tot.number;
  } else {
    const tot = al.addRow(['Total billed', { formula: `B${subRn}` }, { formula: `C${subRn}` }, { formula: `D${subRn}` }]);
    totRn = tot.number;
  }
  const tot = al.getRow(totRn);
  [2, 3, 4].forEach(i => money(tot.getCell(i)));
  tot.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; });

  // Review flags + unmatched, below the table.
  if (result.flags.length) {
    al.addRow([]); const fh = al.addRow(['Review flags']); fh.font = { bold: true, color: { argb: NAVY } };
    for (const f of result.flags) {
      let txt = '';
      if (f.type === 'reclass') txt = `Entity reclassified — ${f.name}: billed ${f.from}, allocated to ${f.to}.`;
      else if (f.type === 'employerPaid') txt = `Employer-paid in full — ${f.name}: employee share $0.00 (${f.entity}).`;
      else if (f.type === 'eligibility') txt = `Eligibility change — ${f.name}: booked 100% employer → ${f.entity || 'unmatched'}.`;
      al.addRow([txt]).font = { size: 10, color: { argb: 'FF52596B' } };
    }
  }
  if (result.unmatched && result.unmatched.length) {
    al.addRow([]); al.addRow(['Unmatched subscribers (no entity — review): ' + result.unmatched.join(', ')]).font = { size: 10, color: { argb: 'FFB9791A' } };
  }

  return await wb.xlsx.writeBuffer();
}

module.exports = { computeAllocation, buildAllocationWorkbook, DEFAULT_OVERRIDES };
