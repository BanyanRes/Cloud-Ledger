// ─── CLRF workpaper: Partners' Capital Accounts schedule (80-line) ────────────
//
// The supplementary Partners' Capital Accounts schedule for County Line Rail
// Fund I, LP (CL entity 40) — one row per investor, Limited Partners then General
// Partners, with subtotals and a grand total — matching the fund administrator's
// FS supplementary schedule column-for-column. Prepared like the other CLRF
// workpapers (gpfees.js / carryclawback.js / pcap.js): buildData -> render ->
// saveToWorkpapers, registered by index.js, located by findWorkpaper().
//
// It reuses the PCAP engine (server/pcap.js buildData): each row is that
// investor class's year-to-date equity roll-forward, so the schedule ties to the
// PCAP statements and to the fund Statement of Changes in Partners' Capital by
// construction. Columns (year-to-date):
//   Capital Commitments | % of Commitments | Partners' Capital at Jan 1 |
//   Contributions | Capital Call Refunds | Syndication Costs |
//   Waived Development Fees | Total Expenses | Net Decrease Resulting from
//   Operations | Transfers of Interest | Partners' Capital at <quarter end>
// Total Expenses = Net Decrease Resulting from Operations = the class's net
// investment loss for the period (net investment income/(loss) + management fee).
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const pcap = require('./pcap');

const FUND_EID = pcap.FUND_EID;
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

const SUM_MONEY = '_($* #,##0_);_($* (#,##0);_($* -_);_(@_)';
const PCT = '0.0000%';
const NAVY = 'FF1F3864';
const HDR_FONT = { name: 'Times New Roman', size: 8, bold: true, color: { argb: 'FFFFFFFF' } };
const TNR = (o = {}) => Object.assign({ name: 'Times New Roman', size: 9 }, o);
const THIN = { style: 'thin' };
const DOUBLE = { style: 'double' };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const spellQuarterEnd = (end) => { const [y, m, d] = String(end).split('-').map(Number); return MONTHS[m - 1] + ' ' + d + ', ' + y; };

// Column layout: label, width, and the per-investor value accessor (year-to-date).
const COLS = [
  { key: 'num', label: 'Partner #', w: 9 },
  { key: 'name', label: 'Investor', w: 40 },
  { key: 'commitment', label: 'Capital Commitments', w: 15, money: true, v: (i) => i.commitment },
  { key: 'pct', label: '% of Commitments', w: 13, pct: true, v: (i, tot) => (tot ? i.commitment / tot : 0) },
  { key: 'beginning', label: 'Partners’ Capital at January 1', w: 15, money: true, v: (i) => i.ytd.beginning },
  { key: 'contributions', label: 'Contributions', w: 14, money: true, v: (i) => i.ytd.contributions },
  { key: 'refunds', label: 'Capital Call Refunds', w: 15, money: true, v: (i) => i.ytd.returnOfCapital },
  { key: 'syndication', label: 'Syndication Costs', w: 13, money: true, v: (i) => i.ytd.syndication },
  { key: 'waived', label: 'Waived Development Fees', w: 13, money: true, v: (i) => i.ytd.waivedDevFees },
  { key: 'totalExp', label: 'Total Expenses', w: 13, money: true, v: (i) => r2(i.ytd.netInvestment + i.ytd.managementFee) },
  { key: 'netDecrease', label: 'Net Decrease Resulting from Operations', w: 15, money: true, v: (i) => r2(i.ytd.netInvestment + i.ytd.managementFee) },
  { key: 'transfers', label: 'Transfers of Interest', w: 13, money: true, v: (i) => i.ytd.transfers },
  { key: 'ending', label: 'Partners’ Capital at End', w: 15, money: true, v: (i) => i.ytd.ending },
];
const MONEY_KEYS = COLS.filter((c) => c.money).map((c) => c.key);

function sectionTotals(rows) {
  const t = {};
  for (const c of COLS) if (c.money || c.key === 'commitment') t[c.key] = 0;
  for (const inv of rows) for (const c of COLS) if (c.money) t[c.key] = r2((t[c.key] || 0) + c.v(inv));
  t.commitment = r2(rows.reduce((s, i) => s + i.commitment, 0));
  return t;
}

function buildWorkbook(data) {
  const q = data.quarter;
  const grandCommit = data.totals.commitment;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CloudLedger'; wb.created = new Date();
  const ws = wb.addWorksheet('Partners’ Capital Accounts', {
    views: [{ state: 'frozen', xSplit: 2, ySplit: 6, showGridLines: false }],
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  // Titles
  const wsTitle = (row, text, o = {}) => { const c = ws.getCell('A' + row); c.value = text; c.font = TNR(Object.assign({ bold: true, size: 11 }, o)); };
  wsTitle(1, data.entity_name); wsTitle(2, 'Partners’ Capital Accounts', { size: 10 });
  wsTitle(3, 'For the Quarter Ended ' + spellQuarterEnd(q.end), { size: 9, italic: true, bold: false });

  // Header row (row 5)
  COLS.forEach((c, i) => {
    const cell = ws.getRow(5).getCell(i + 1);
    cell.value = c.label; cell.font = HDR_FONT;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    cell.alignment = { horizontal: i < 2 ? 'left' : 'center', wrapText: true, vertical: 'bottom' };
    ws.getColumn(i + 1).width = c.w;
  });

  let r = 6;
  let num = 0;
  const writeRow = (inv) => {
    num += 1;
    const row = ws.getRow(r);
    COLS.forEach((c, i) => {
      const cell = row.getCell(i + 1);
      if (c.key === 'num') { cell.value = num; cell.font = TNR(); cell.alignment = { horizontal: 'center' }; }
      else if (c.key === 'name') { cell.value = inv.name; cell.font = TNR(); }
      else if (c.pct) { cell.value = c.v(inv, grandCommit); cell.numFmt = PCT; cell.font = TNR(); }
      else { cell.value = c.v(inv); cell.numFmt = SUM_MONEY; cell.font = TNR(); }
    });
    r += 1;
  };
  const writeTotal = (label, rows, opts = {}) => {
    const t = sectionTotals(rows);
    const row = ws.getRow(r);
    row.getCell(2).value = label; row.getCell(2).font = TNR({ bold: true });
    COLS.forEach((c, i) => {
      if (c.key === 'num' || c.key === 'name') return;
      const cell = row.getCell(i + 1);
      if (c.pct) { cell.value = grandCommit ? t.commitment / grandCommit : 0; cell.numFmt = PCT; }
      else { cell.value = t[c.key]; cell.numFmt = SUM_MONEY; }
      cell.font = TNR({ bold: true });
      cell.border = { top: THIN, bottom: opts.grand ? DOUBLE : undefined };
    });
    r += 1;
    return t;
  };

  const lps = data.investors.filter((i) => i.partner_type === 'LP');
  const gps = data.investors.filter((i) => i.partner_type === 'GP');

  ws.getCell('A' + r).value = 'Limited Partners'; ws.getCell('A' + r).font = TNR({ bold: true }); r += 1;
  lps.forEach(writeRow);
  writeTotal('Total Limited Partners', lps);
  r += 1;
  ws.getCell('A' + r).value = 'General Partners'; ws.getCell('A' + r).font = TNR({ bold: true }); r += 1;
  gps.forEach(writeRow);
  writeTotal('Total General Partners', gps);
  r += 1;
  writeTotal('Total Partners’ Capital', data.investors, { grand: true });
  r += 1;
  ws.getCell('A' + r).value = 'No Assurance Provided.'; ws.getCell('A' + r).font = TNR({ italic: true, size: 8 }); r += 2;
  ws.getCell('A' + r).value = 'Total Expenses = Net Decrease Resulting from Operations = net investment income/(loss) + management fee (year-to-date, per investor class).'
    + ' Sourced from the CLRF general ledger by class tag; ties to the PCAP statements and the fund Statement of Changes in Partners’ Capital.';
  ws.getCell('A' + r).font = TNR({ italic: true, size: 8 });

  return wb;
}

// ── Persistence + routes (mirror the other CLRF workpapers). ──────────────────
const folderFor = (quarter) => 'Workpapers/Partners’ Capital Accounts Schedule/' + quarter.year + '/' + quarter.quarter;
const fileNameFor = (quarter) => 'CLRF_Partners_Capital_Accounts_' + quarter.label + '.xlsx';

function saveToWorkpapers(ctx, eid, quarter, buf, who) {
  const { db, workpapersDir } = ctx;
  const folder = folderFor(quarter);
  const original = fileNameFor(quarter);
  const parts = folder.split('/');
  const ins = db.prepare('INSERT OR IGNORE INTO entity_folders (entity_id, folder_path, created_by, created_at) '
    + "VALUES (?, ?, ?, datetime('now'))");
  for (let i = 1; i <= parts.length; i++) ins.run(eid, parts.slice(0, i).join('/'), who);
  const prior = db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ?').all(eid, folder, original);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(workpapersDir, String(eid), p.stored_filename)); } catch (e) { /* already gone */ }
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

function findWorkpaper(ctx, eid, quarterEnd) {
  const quarter = pcap.resolveQuarter(quarterEnd);
  const row = ctx.db.prepare('SELECT * FROM entity_files WHERE entity_id = ? AND folder_path = ? '
    + 'AND original_name = ? ORDER BY id DESC LIMIT 1').get(eid, folderFor(quarter), fileNameFor(quarter));
  if (!row) return null;
  return Object.assign({}, row, { quarter, abs_path: path.join(ctx.workpapersDir, String(eid), row.stored_filename) });
}

function registerPcapScheduleRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/pcap-schedule/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const quarter = pcap.resolveQuarter(body.quarter_end || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = pcap.buildData(ctx, quarter, { entity_id: eid, carryByClass: body.carry_by_class || null });
        const wb = buildWorkbook(data);
        const buf = Buffer.from(await wb.xlsx.writeBuffer());
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        const lps = data.investors.filter((i) => i.partner_type === 'LP');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-PCAP-Schedule-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          rows: data.totals.count, lp_rows: lps.length, gp_rows: data.totals.count - lps.length,
          commitment: data.totals.commitment, ending: data.totals.ytd.ending,
        }).replace(/[\r\n]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { buildWorkbook, saveToWorkpapers, findWorkpaper, registerPcapScheduleRoutes, sectionTotals, COLS };
