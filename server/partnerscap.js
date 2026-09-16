// ─── CLRF workpaper: Partners' Capital Accounts (FS supplemental disclosure) ──
//
// The fund-level Partners' Capital Accounts schedule for County Line Rail Fund I,
// LP (CL entity 40) — one row per partner (the full un-merged roster) with the
// quarterly capital roll-forward, matching the fund's FS supplemental disclosure
// (the "Partners cap" workpaper). Prepared like the other CLRF workpapers:
// buildData (reuses pcap.buildData, noMerge, quarter period) -> buildWorkbook ->
// saveToWorkpapers, registered by index.js.
//
// Partner # ties to the FS supplemental disclosure via a fixed roster (below):
// original partners are numbered alphabetically (1-70); later subscribers are
// appended in join order (71-81). CL stores no partner number, so the roster is
// embedded and keyed by name (with a small alias map for name variants).
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const pcap = require('./pcap.js');

const FUND_EID = 40;
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const norm = (s) => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();

// Partner # per the FS supplemental disclosure (2026).
const FS_PARTNER_NO = {
  "1999 britta biesecker declaration of trust dtd 5/12/99": 1,
  "1999 frederick n biesecker ii declaration of trust dtd 5/12/99": 2,
  "1999 lissa biesecker longacre declaration of trust dtd 5/12/99": 3,
  "abigail levy": 4, "akwt llc": 5, "alexander s. bloomingdale": 6,
  "alexander sharpe moore": 7, "all dogs go to heaven, llc": 8, "alten investments llc": 9,
  "avery ryan klann": 10, "bbr real assets fund 2024 llc": 11, "belclaire asset management, llc": 12,
  "benaiah inc.": 13, "benjamin johnson": 14, "bloomingdale family trust share #3": 15,
  "brian carroll 2011 long-term trust": 16, "bryan c. white llc": 17, "cameron j. tringale": 18,
  "charing hatcherville llc": 19, "charing silsbee, llc": 20, "chase f. monroe": 21,
  "chase the lion, inc.": 22, "chris schaaf": 23, "christopher j. campbell": 24,
  "christopher johnson": 25, "daniel ross": 26, "endurance companies llc": 27,
  "forem capital, llc": 28, "francis brawley": 29, "geoffrey and elizabeth bloomingdale living trust": 30,
  "greene 2000 family trust": 31, "greene marital trust": 32, "gregory brent wood": 33,
  "hundley family trust": 34, "in god we trust 2017 living trust": 35, "james campbell and patricia campbell": 36,
  "james kane": 37, "james thorp": 38, "jeffrey a. lipsitz revocable trust": 39,
  "jeffrey d furber revocable trust": 40,
  "jeremy m. jacobs trust u/a/d april 1, 1973 fbo margaret l. reichenback": 41,
  "john mark comer": 42, "landscape i, lp": 43, "lisa bell trust, elisabeth b. bell, trustee": 44,
  "madten investments llc": 45, "matt mccoy": 46, "matterhorn holdings, llc": 47,
  "max o. reinbach iii": 48, "meghann c. robinson": 49, "nahas investments ii llc": 50,
  "pease river holdings llc": 51, "pt burros, llc": 52, "robert d. falese iii": 53,
  "robert parker yates": 54, "robert s. shafir & donna r. shafir ten/com": 55,
  "robert s. shafir 2011 children's trust": 56, "romar partners, lp": 57,
  "steven and chani laufer 2012 dynasty trust": 58, "steven m. laufer 2012 trust": 59,
  "stewart tate": 60, "tc wilson family trust": 61, "terminal refrigerating & warehousing corp.": 62,
  "the 2015 pipe reality delaware trust": 63, "the david l. schnadig 2017 trust": 64,
  "the diane rosen 2021 irrevocable family trust": 65,
  "the mark t. kirchdorfer 2015 family delaware dynasty trust": 66,
  "university of southern california": 67, "walker d. zimmerman": 68,
  "wilson family trust dtd 3-6-1987": 69, "wilson gst exempt trust": 70,
  "legacy knight strategic opportunities fund llc - cl rail series": 71,
  "argos holdings, llc": 72,
  "texas emerging managers private markets program, l.p. (2025-1 re investment series)": 73,
  "gcm grosvenor nj re emerging managers program, lp (2023-1 investment series)": 74,
  "gcm grosvenor - nystrs real estate investment partners, l.p. (2024-1 series)": 75,
  "crptf-gcm middle-market re partnership, l.p. (2025-01 investment series)": 76,
  "clip sponsor llc": 77, "clr silsbee sponsor llc": 78, "odyssey holdings, llc": 79,
  "palmatum hr illiquid, llc": 80, "james bloomingdale": 81,
};
// CL class name -> FS roster name, for name variants.
const NAME_ALIAS = {
  "stewart sevier tate revocable trust dated october 4, 2013": "stewart tate",
};
function partnerNo(name) {
  const n = norm(name);
  if (FS_PARTNER_NO[n] != null) return FS_PARTNER_NO[n];
  const a = NAME_ALIAS[n];
  return (a && FS_PARTNER_NO[a] != null) ? FS_PARTNER_NO[a] : null;
}

// Build the schedule rows from the quarterly PCAP data (un-merged roster).
function buildData(ctx, quarter, opts = {}) {
  const eid = opts.entity_id || FUND_EID;
  const data = pcap.buildData(ctx, quarter, { entity_id: eid, noMerge: true });
  const totalCommit = data.totals.commitment || 0;
  const fundIncomeQ = (data.investmentIncome && data.investmentIncome.q) || 0;
  // Operations = net investment income + management fee (excludes syndication and
  // unrealized, which are separate columns). Fund investment income is allocated
  // pro-rata to each partner's capital commitment (the FS disclosure basis); the
  // balance of operations is Total Expenses. Net Decrease = the two combined and
  // is GL-exact, so the ending capital foots regardless of the income/expense split.

  const rows = [];
  for (const inv of data.investors) {
    const no = partnerNo(inv.name);
    if (no == null) continue; // classes not in the FS roster (stray/merge-target)
    const y = inv.q;
    const opsNet = r2((y.netInvestment || 0) + (y.managementFee || 0));
    const investmentIncome = totalCommit ? r2(fundIncomeQ * (inv.commitment / totalCommit)) : 0;
    const totalExpenses = r2(opsNet - investmentIncome);
    rows.push({
      no, name: inv.name, partner_type: inv.partner_type,
      commitment: inv.commitment,
      pct: totalCommit ? inv.commitment / totalCommit : 0,
      beginning: y.beginning, contributions: y.contributions, refunds: y.returnOfCapital,
      syndication: y.syndication, waived: y.waivedDevFees,
      investmentIncome, managementFee: y.managementFee, totalExpenses, netDecrease: opsNet,
      unrealized: y.unrealized, transfers: y.transfers, ending: y.ending,
    });
  }
  rows.sort((a, b) => a.no - b.no);
  const lps = rows.filter((r) => r.partner_type === 'LP');
  const gps = rows.filter((r) => r.partner_type === 'GP');
  return { entity_id: eid, entity_name: data.entity_name, quarter, rows, lps, gps, totalCommit, fundIncomeQ };
}

// ─── Workbook ────────────────────────────────────────────────────────────────
const NAVY = 'FF1F3864';
const HDR = { name: 'Times New Roman', size: 8, bold: true, color: { argb: 'FFFFFFFF' } };
const TNR = (o = {}) => Object.assign({ name: 'Times New Roman', size: 9 }, o);
const MONEY = '#,##0;(#,##0);-';
const PCT = '0.0000%';
const THIN = { style: 'thin' };
const DOUBLE = { style: 'double' };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const spell = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };

const FIELDS = [
  { key: 'commitment', label: 'Capital Commitments', money: true },
  { key: 'pct', label: '% of Commitments', pct: true },
  { key: 'beginning', label: 'Partners’ Capital at {QSTART}', money: true },
  { key: 'contributions', label: 'Contributions', money: true },
  { key: 'refunds', label: 'Capital Call Refunds', money: true },
  { key: 'syndication', label: 'Syndication Costs', money: true },
  { key: 'waived', label: 'Waived Development Fees', money: true },
  { key: 'investmentIncome', label: 'Investment Income', money: true },
  { key: 'managementFee', label: 'Management Fees', money: true },
  { key: 'totalExpenses', label: 'Total Expenses', money: true },
  { key: 'netDecrease', label: 'Net Decrease in Partners’ Capital Resulting from Operations', money: true },
  { key: 'unrealized', label: 'Change in Unrealized Gain/Loss', money: true },
  { key: 'transfers', label: 'Transfers of Interest', money: true },
  { key: 'ending', label: 'Partners’ Capital at {QEND}', money: true },
];

function buildWorkbook(data) {
  const q = data.quarter;
  const [ay, am] = String(q.end).split('-').map(Number);
  const QSL = { 3: 'January 1, ', 6: 'April 1, ', 9: 'July 1, ', 12: 'October 1, ' };
  const qStart = (QSL[am] || 'January 1, ') + ay;
  const qEnd = spell(q.end);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  const ws = wb.addWorksheet('Partners Cap', {
    views: [{ state: 'frozen', xSplit: 2, ySplit: 6, showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  const title = (row, text, o = {}) => { const c = ws.getCell('A' + row); c.value = text; c.font = TNR(Object.assign({ bold: true, size: 11 }, o)); };
  title(1, data.entity_name); title(3, 'Partners’ Capital Accounts', { size: 10 });
  title(5, 'For the Quarter Ended ' + qEnd, { size: 9, italic: true, bold: false });

  const headers = ['Partner #', 'Partners'].concat(FIELDS.map((f) => f.label.replace('{QSTART}', qStart).replace('{QEND}', qEnd)));
  const widths = [8, 42].concat(FIELDS.map(() => 13));
  const hr = ws.getRow(6);
  headers.forEach((t, i) => {
    const c = hr.getCell(i + 1);
    c.value = t; c.font = HDR;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    c.alignment = { horizontal: i < 2 ? 'left' : 'center', wrapText: true, vertical: 'bottom' };
    ws.getColumn(i + 1).width = widths[i];
  });

  let r = 7;
  const writeRow = (row) => {
    const xl = ws.getRow(r);
    xl.getCell(1).value = row.no; xl.getCell(1).font = TNR(); xl.getCell(1).alignment = { horizontal: 'center' };
    xl.getCell(2).value = row.name; xl.getCell(2).font = TNR();
    FIELDS.forEach((f, i) => {
      const c = xl.getCell(i + 3);
      c.value = row[f.key];
      c.numFmt = f.pct ? PCT : MONEY;
      c.font = TNR();
    });
    r += 1;
  };
  const writeTotal = (label, rows, opts = {}) => {
    const xl = ws.getRow(r);
    xl.getCell(2).value = label; xl.getCell(2).font = TNR({ bold: true });
    FIELDS.forEach((f, i) => {
      const c = xl.getCell(i + 3);
      if (f.pct) c.value = data.totalCommit ? r2(rows.reduce((s, x) => s + x.commitment, 0)) / data.totalCommit : 0;
      else c.value = r2(rows.reduce((s, x) => s + (x[f.key] || 0), 0));
      c.numFmt = f.pct ? PCT : MONEY;
      c.font = TNR({ bold: true });
      c.border = { top: THIN, bottom: opts.grand ? DOUBLE : undefined };
    });
    r += 1;
  };

  ws.getCell('A' + r).value = 'Limited Partners'; ws.getCell('A' + r).font = TNR({ bold: true }); r += 1;
  data.lps.forEach(writeRow);
  writeTotal('Total Limited Partners', data.lps);
  r += 1;
  ws.getCell('A' + r).value = 'General Partners'; ws.getCell('A' + r).font = TNR({ bold: true }); r += 1;
  data.gps.forEach(writeRow);
  writeTotal('Total General Partners', data.gps);
  r += 1;
  writeTotal('Total Partners’ Capital', data.rows, { grand: true });
  r += 2;
  ws.getCell('A' + r).value = 'Partner # ties to the FS supplemental disclosure. Sourced from the CLRF general ledger (entity ' + data.entity_id + ') by class tag; ties to the PCAP statements. No Assurance Provided.';
  ws.getCell('A' + r).font = TNR({ italic: true, size: 8 });
  return wb;
}

// ── Persistence + route (mirror pcapschedule.js). ─────────────────────────────
const folderFor = (q) => 'Workpapers/Partners’ Capital Accounts Schedule/' + q.year + '/' + q.quarter;
const fileNameFor = (q) => 'CLRF_Partners_Cap_' + q.label + '.xlsx';

function saveToWorkpapers(ctx, eid, q, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(q);
  const original = fileNameFor(q);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* gone */ }
    db.prepare('DELETE FROM entity_files WHERE id = ?').run(p.id);
  }
  const dir = path.join(workpapersDir, String(eid));
  fs.mkdirSync(dir, { recursive: true });
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + original.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, '
    + "uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))")
    .run(eid, folder, stored, original, buf.length,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', who);
  return { folder_path: folder, original_name: original, replaced: prior.length };
}

function registerPartnersCapRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/partners-cap/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const q = pcap.resolveQuarter((req.body && req.body.quarter_end) || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, q, { entity_id: eid });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, q, buf, who);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-PartnersCap-Summary', JSON.stringify({
          quarter: q.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          partners: data.rows.length,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) { res.status(400).json({ error: e.message }); }
    });

  app.get('/api/entities/:eid/partners-cap.xlsx', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const eid = Number(req.params.eid);
      const q = pcap.resolveQuarter((req.query && req.query.as_of) || '');
      const data = buildData(ctx, q, { entity_id: eid });
      const wb = buildWorkbook(data);
      const buf = Buffer.from(await wb.xlsx.writeBuffer());
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="' + fileNameFor(q) + '"');
      res.send(buf);
    } catch (e) { res.status(400).json({ error: e.message }); }
  });
}

module.exports = { buildData, buildWorkbook, saveToWorkpapers, registerPartnersCapRoutes, partnerNo, FS_PARTNER_NO, FUND_EID };
