// Insurance Allocation engine (pure module).
// Given a carrier health-insurance billing invoice and a consolidated billing
// report (both .xlsx buffers), allocates the monthly premium across the
// commonly-owned entities and builds a workpaper workbook.
//
// Rules (confirmed by Jimmy, 2026-08-12):
//  - Employee share (monthly) = SUM(Employee Cost Per Pay Period across ALL
//    benefit lines, HSA included) x Pay Periods / 12.
//  - Employer share = Premium - Employee share.
//  - Entity = the employee's Department in the consolidated billing, matched by
//    Last + First name, subject to an override table.
//  - Eligibility-change amounts are booked 100% to the employer, tagged to the
//    same entity as the subscriber.
//  - Override table (per entity settings, defaults here):
//      entity reassignment: MARTINEZ, JONATHAN -> SRN
//      employer-paid-in-full (EE share = 0): BROSSEAU, BENJAMIN
//
// The module never runs the DB or Express; index.js wires the route.

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

// Department string (consolidated billing) -> short entity code.
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
  // "LAST, FIRST" (upper) -> entity code, overriding the department mapping.
  entity: { 'MARTINEZ, JONATHAN': 'SRN' },
  // "LAST, FIRST" (upper) that the employer pays in full (employee share = 0).
  employerPaid: ['BROSSEAU, BENJAMIN'],
};

function deptToCode(dept) {
  const s = String(dept || '');
  for (const [re, code] of DEPT_TO_CODE) if (re.test(s)) return code;
  return null;
}

// Normalize a subscriber/employee name to a match key "LAST, FIRST" (upper),
// using only the first token of the first name so middle initials don't break
// the match, and so duplicate last names resolve on first name.
function nameKey(last, first) {
  const l = String(last || '').trim().toUpperCase();
  const f = String(first || '').trim().split(/\s+/)[0].toUpperCase().replace(/[.,]$/, '');
  return l && f ? `${l}, ${f}` : (l || '');
}
// From a carrier "LAST, FIRST M" string.
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

// Read every sheet of a workbook into arrays-of-rows (header:1), preserving cells.
function readSheets(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const out = {};
  for (const name of wb.SheetNames) {
    out[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
  }
  return out;
}

// Locate a header row by required column labels; returns {rowIndex, pick, map} or null.
function findHeader(rows, required) {
  for (let r = 0; r < Math.min(rows.length, 40); r++) {
    const row = rows[r] || [];
    const map = {};
    row.forEach((cell, c) => {
      const key = String(cell == null ? '' : cell).replace(/\s+/g, ' ').trim().toLowerCase();
      if (key) map[key] = c;
    });
    if (required.every(req => req.some(alias => map[alias] != null))) {
      const pick = req => { for (const a of req) if (map[a] != null) return map[a]; return -1; };
      return { rowIndex: r, pick, map };
    }
  }
  return null;
}

// ── Consolidated billing: name -> { code, ee } (monthly employee share). ──
function parseConsolidated(sheets) {
  // Require the long-format tab: it carries a single "Plan Type" column (one row
  // per employee-per-benefit-line). The wide "Details" tab has per-plan column
  // groups and no generic Plan Type — matching it would sum only one plan's EE.
  let rows = null, hdr = null;
  const names = Object.keys(sheets).sort((a, b) => (/summary/i.test(b) ? 1 : 0) - (/summary/i.test(a) ? 1 : 0));
  for (const name of names) {
    const rws = sheets[name];
    const h = findHeader(rws, [
      ['last name'], ['department'], ['plan type'], ['employee cost per pay period'], ['pay periods'],
    ]);
    if (h) { rows = rws; hdr = h; break; }
  }
  if (!rows) throw new Error('Consolidated billing: could not find the long-format summary sheet (needs Department, Plan Type, Employee Cost Per Pay Period, Pay Periods columns).');
  const cLast = hdr.pick(['last name']);
  const cFirst = hdr.pick(['first name']);
  const cDept = hdr.pick(['department']);
  const cEcpp = hdr.pick(['employee cost per pay period']);
  const cPp = hdr.pick(['pay periods']);

  const dept = {};   // key -> code
  const ee = {};     // key -> monthly employee share
  for (let r = hdr.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const last = row[cLast], first = cFirst >= 0 ? row[cFirst] : '';
    if (last == null || String(last).trim() === '') continue;
    const key = nameKey(last, first);
    const code = deptToCode(row[cDept]);
    if (code && dept[key] == null) dept[key] = code;
    const pp = money(row[cPp]);
    ee[key] = (ee[key] || 0) + money(row[cEcpp]) * (pp ? pp / 12 : 0);
  }
  return { dept, ee };
}

// ── Carrier invoice: membership detail + eligibility changes. ──
function parseInvoice(sheets) {
  // Membership detail: header with Subscriber Name + Premium Amount.
  let memRows = null, memHdr = null, invMeta = {};
  for (const name of Object.keys(sheets)) {
    const rws = sheets[name];
    const h = findHeader(rws, [['subscriber name'], ['premium amount']]);
    if (h) {
      memRows = rws; memHdr = h;
      // Scrape a few labeled meta cells above the table (Invoice #, Billing Period).
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
  const cName = memHdr.pick(['subscriber name']);
  const cPrem = memHdr.pick(['premium amount']);
  const members = [];
  for (let r = memHdr.rowIndex + 1; r < memRows.length; r++) {
    const row = memRows[r] || [];
    // End of the membership table: a Subtotal/Total row, or the Rate Change
    // Legend block that follows it (their labels sit in other columns, so scan
    // the whole row, not just the name cell).
    if ((row || []).some(cell => /^(sub)?total$|rate change legend/i.test(String(cell == null ? '' : cell).trim()))) break;
    const nm = row[cName];
    if (nm == null || String(nm).trim() === '') continue;
    const prem = money(row[cPrem]);
    if (!prem) continue; // legend/blank rows carry no premium
    members.push({ name: String(nm).trim(), key: nameKeyFromFull(nm), premium: prem });
  }

  // Eligibility changes: a sheet whose header carries Change Code (+ Premium Amount).
  const elig = [];
  for (const name of Object.keys(sheets)) {
    const rws = sheets[name];
    const h = findHeader(rws, [['subscriber name'], ['premium amount'], ['change code']]);
    if (!h) continue;
    const en = h.pick(['subscriber name']), ep = h.pick(['premium amount']);
    for (let r = h.rowIndex + 1; r < rws.length; r++) {
      const row = rws[r] || [];
      const nm = row[en];
      if (nm == null || String(nm).trim() === '') continue;
      if (/^subtotal|^total\b/i.test(String(nm).trim())) break;
      elig.push({ name: String(nm).trim(), key: nameKeyFromFull(nm), amount: money(row[ep]) });
    }
    break;
  }
  return { members, elig, invMeta };
}

// ── Compute the allocation. ──
function computeAllocation({ invoiceBuf, consolidatedBuf, overrides }) {
  const ov = {
    entity: { ...DEFAULT_OVERRIDES.entity, ...((overrides && overrides.entity) || {}) },
    employerPaid: new Set([...DEFAULT_OVERRIDES.employerPaid, ...((overrides && overrides.employerPaid) || [])].map(s => s.toUpperCase())),
  };
  const cons = parseConsolidated(readSheets(consolidatedBuf));
  const { members, elig, invMeta } = parseInvoice(readSheets(invoiceBuf));

  const ent = {}; // code -> {premium, employer, employee}
  const bucket = code => (ent[code] || (ent[code] = { premium: 0, employer: 0, employee: 0 }));
  const flags = [];
  const memberRows = [];
  const unmatched = [];

  for (const m of members) {
    const forcedEntity = ov.entity[m.key];
    const deptEntity = cons.dept[m.key];
    const code = forcedEntity || deptEntity || null;
    if (!code) { unmatched.push(m.name); continue; }
    const employerPaid = ov.employerPaid.has(m.key);
    const ee = employerPaid ? 0 : (cons.ee[m.key] || 0);
    const er = m.premium - ee;
    const b = bucket(code);
    b.premium += m.premium; b.employer += er; b.employee += ee;
    memberRows.push({ name: m.name, entity: code, premium: m.premium, ee, er,
      note: forcedEntity && deptEntity && forcedEntity !== deptEntity ? `reclassified from ${deptEntity}` : (employerPaid ? 'employer-paid in full' : '') });
    if (forcedEntity && deptEntity && forcedEntity !== deptEntity)
      flags.push({ type: 'reclass', name: m.name, from: deptEntity, to: forcedEntity, premium: m.premium });
    if (employerPaid)
      flags.push({ type: 'employerPaid', name: m.name, entity: code, premium: m.premium });
    if (!deptEntity && !forcedEntity) unmatched.push(m.name);
  }

  // Eligibility changes -> 100% employer, same entity as the subscriber.
  const eligRows = [];
  let eligibilityTotal = 0;
  for (const e of elig) {
    const code = ov.entity[e.key] || cons.dept[e.key] || null;
    eligibilityTotal += e.amount;
    if (code) { bucket(code).employer += e.amount; }
    eligRows.push({ name: e.name, entity: code, amount: e.amount });
    flags.push({ type: 'eligibility', name: e.name, entity: code, amount: e.amount });
  }

  const codes = [...CODE_ORDER.filter(c => ent[c]), ...Object.keys(ent).filter(c => !CODE_ORDER.includes(c))];
  const entities = codes.map(c => ({ code: c, name: (CODE_META[c] && CODE_META[c].name) || c,
    premium: round2(ent[c].premium), employer: round2(ent[c].employer), employee: round2(ent[c].employee) }));

  const premiumSubtotal = round2(members.reduce((s, m) => s + m.premium, 0));
  const employeeTotal = round2(memberRows.reduce((s, m) => s + m.ee, 0));
  const employerBase = round2(premiumSubtotal - employeeTotal);
  const employerTotal = round2(employerBase + eligibilityTotal);
  const totalBilled = round2(premiumSubtotal + eligibilityTotal);

  return {
    period: (invMeta.period || '').trim(),
    invoice: (invMeta.invoice || '').trim(),
    entities,
    subtotal: { premium: premiumSubtotal, employer: employerBase, employee: employeeTotal },
    eligibility: eligRows, eligibilityTotal: round2(eligibilityTotal),
    totalBilled, employerTotal, employeeTotal,
    subscriberCount: members.length,
    flags, members: memberRows, unmatched,
    reconciled: unmatched.length === 0,
  };
}

function round2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }

// ── Build the workpaper workbook (ExcelJS) from a computeAllocation result. ──
async function buildAllocationWorkbook(result, opts = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger';
  const NAVY = 'FF0D1B2E', HEADFILL = 'FFF2F4F7', GREY = 'FF8A90A0';
  const money = { numFmt: '#,##0.00;(#,##0.00)' };
  const ws = wb.addWorksheet('Allocation', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 26 }];

  const title = ws.addRow([opts.title || 'Health Insurance Allocation']);
  title.font = { bold: true, size: 15, color: { argb: NAVY } };
  const meta = [];
  if (opts.entityName) meta.push('Entity: ' + opts.entityName);
  if (result.period) meta.push('Billing Period: ' + result.period);
  if (result.invoice) meta.push('Invoice #: ' + result.invoice);
  meta.push('Subscribers: ' + result.subscriberCount);
  const mrow = ws.addRow([meta.join('     ')]); mrow.font = { color: { argb: GREY }, size: 10 };
  ws.addRow([]);

  const hdr = ws.addRow(['Entity', 'Premium', 'Employer', 'Employee']);
  hdr.eachCell(c => { c.font = { bold: true, size: 10, color: { argb: GREY } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADFILL } }; });
  hdr.getCell(1).alignment = { horizontal: 'left' };
  const numAlign = (r) => { [2, 3, 4].forEach(i => { r.getCell(i).numFmt = money.numFmt; r.getCell(i).alignment = { horizontal: 'right' }; }); };
  for (const e of result.entities) {
    const r = ws.addRow([`${e.name}  ·  ${e.code}`, e.premium, e.employer, e.employee]); numAlign(r);
  }
  const sub = ws.addRow(['Subtotal', result.subtotal.premium, result.subtotal.employer, result.subtotal.employee]);
  sub.font = { bold: true }; numAlign(sub);
  sub.eachCell(c => { c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } }; });
  if (result.eligibilityTotal) {
    const er = ws.addRow(['Eligibility change (100% employer)', result.eligibilityTotal, result.eligibilityTotal, null]);
    numAlign(er); er.font = { color: { argb: GREY } };
  }
  const tot = ws.addRow(['Total billed', result.totalBilled, result.employerTotal, result.employeeTotal]);
  tot.font = { bold: true, color: { argb: 'FFFFFFFF' } }; numAlign(tot);
  tot.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; });

  // Review flags
  if (result.flags.length) {
    ws.addRow([]); const fh = ws.addRow(['Review flags']); fh.font = { bold: true, color: { argb: NAVY } };
    for (const f of result.flags) {
      let txt = '';
      if (f.type === 'reclass') txt = `Entity reclassified — ${f.name}: billed ${f.from}, allocated to ${f.to} (premium ${f.premium.toFixed(2)}).`;
      else if (f.type === 'employerPaid') txt = `Employer-paid in full — ${f.name}: employee share $0.00; company covers ${f.premium.toFixed(2)} (${f.entity}).`;
      else if (f.type === 'eligibility') txt = `Eligibility change — ${f.name}: ${f.amount.toFixed(2)} booked 100% employer → ${f.entity || 'unmatched'}.`;
      const rr = ws.addRow([txt]); rr.font = { size: 10, color: { argb: 'FF52596B' } };
    }
  }
  if (result.unmatched && result.unmatched.length) {
    ws.addRow([]); const uh = ws.addRow(['Unmatched subscribers (no entity — review)']); uh.font = { bold: true, color: { argb: 'FFB9791A' } };
    for (const n of result.unmatched) ws.addRow([n]);
  }

  // Member detail
  const det = wb.addWorksheet('Member Detail', { views: [{ showGridLines: false }] });
  det.columns = [{ width: 30 }, { width: 10 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 26 }];
  const dh = det.addRow(['Subscriber', 'Entity', 'Premium', 'Employee', 'Employer', 'Note']);
  dh.eachCell(c => { c.font = { bold: true, size: 10, color: { argb: GREY } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADFILL } }; });
  for (const m of result.members) {
    const r = det.addRow([m.name, m.entity, round2(m.premium), round2(m.ee), round2(m.er), m.note || '']);
    [3, 4, 5].forEach(i => { r.getCell(i).numFmt = money.numFmt; r.getCell(i).alignment = { horizontal: 'right' }; });
  }
  return await wb.xlsx.writeBuffer();
}

module.exports = { computeAllocation, buildAllocationWorkbook, DEFAULT_OVERRIDES };
