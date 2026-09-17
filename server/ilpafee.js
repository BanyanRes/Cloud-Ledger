// ─── CLRF workpaper: ILPA Fee Reporting Template (GCM Grosvenor investors) ────
//
// County Line Rail Fund I, LP (CL entity 40) is the "Master Fund". Its four GCM
// Grosvenor-affiliated LP investors each receive an ILPA Fee Reporting Template
// (one tab per investor). This workpaper reproduces the vendor template EXACTLY
// (server/assets/ilpa_fee_template.xlsx keeps the fonts, notes, merges, blue
// input cells and formulas) and only fills the blue input cells from the general
// ledger for the requested quarter.
//
// Presentation: Contributions are shown GROSS and returns of capital are shown as
// a positive value on the Distributions line (the template's Total Cash Flows
// formula nets them). Every subtotal stays a template formula, so the four
// statements foot and tie to the PCAP / GL ending capital by construction.
//
// Blue input cells populated per tab (E = QTD col 5, F = YTD col 6; the Since-
// Inception column G is the template's own =+F link, and Beginning NAV SI = 0):
//   row16 Beginning NAV, row17 Contributions (gross), row18 Distributions (+),
//   row21 Management Fees, row23 Partnership Expenses, row27 Interest Income,
//   row32 Placement Fees ; plus C4 Investor Remaining Commitment and the period
//   date cells.
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const pcap = require('./pcap');

// ExcelJS serializes <sheetPr> children as <pageSetUpPr/><outlinePr/>, but the
// OOXML CT_SheetPr schema requires <outlinePr/> BEFORE <pageSetUpPr/>. Excel's
// strict loader rejects the reversed order and "repairs" the file on open
// (Repaired Records: worksheet — "Load error. Line 2, column 0."). Swap them back
// in every worksheet part so the workbook opens clean. LibreOffice/openpyxl are
// lenient about the order, which is why it renders fine everywhere except Excel.
async function fixSheetPrOrder(buf) {
  try {
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    let changed = false;
    for (const name of names) {
      const xml = await zip.file(name).async('string');
      const fixed = xml.replace(/(<pageSetUpPr\b[^>]*\/>)\s*(<outlinePr\b[^>]*\/>)/g, '$2$1');
      if (fixed !== xml) { zip.file(name, fixed); changed = true; }
    }
    if (!changed) return buf;
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) { return buf; }
}

const FUND_EID = 40;
const TEMPLATE = path.join(__dirname, 'assets', 'ilpa_fee_template.xlsx');
const INTEREST_ACCT = '401000';           // fund interest income, allocated pro-rata
const INCEPTION = new Date(Date.UTC(2024, 6, 12)); // 7/12/2024 master-fund inception

// Investor class_id -> template tab name (the four GCM Grosvenor sleeves).
const CLASS_TAB = {
  250: 'CRPTF-GCM',
  248: 'GCM Grosvenor NJ RE',
  247: 'Texas Emerging Managers',
  249: 'NYSTRS',
};

const r0 = (n) => Math.round(Number(n) || 0);   // whole-dollar (matches vendor template)

function folderFor(q) { return 'Workpapers/ILPA Fee/' + q.year + '/' + q.quarter; }
function fileNameFor(q) { return 'CLRF_ILPA_Fee_' + q.label + '.xlsx'; }

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

// Fund interest income (positive) over a period, from the GL.
function fundInterest(ctx, eid, from, to) {
  const rows = ctx.computeBalances(eid, { from, to });
  const r = (rows || []).find((x) => String(x.code) === INTEREST_ACCT);
  return r ? (Number(r.balance) || 0) : 0;
}

// Assemble the ILPA line values for the four GCM sleeves for one quarter.
function buildData(ctx, quarter) {
  const eid = FUND_EID;
  // Per-class roll-forward (un-merged: keep every class and its return-of-capital).
  const data = pcap.buildData(ctx, quarter, { entity_id: eid, noMerge: true });
  const byClass = {}; (data.investors || []).forEach((i) => { byClass[i.class_id] = i; });
  // Ownership pct (commitment basis) for the interest-income allocation.
  const coms = ctx.db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
  const totalCommit = coms.reduce((s, c) => s + (Number(c.commitment_amount) || 0), 0) || 1;
  const commitBy = {}; coms.forEach((c) => { commitBy[c.class_id] = Number(c.commitment_amount) || 0; });
  // Fund interest income for the two reported periods (SI uses YTD: these sleeves
  // were first funded in the current year, so since-inception == year-to-date).
  const intQTD = fundInterest(ctx, eid, quarter.quarter_start, quarter.end);
  const intYTD = fundInterest(ctx, eid, quarter.year_start, quarter.end);

  const out = {};
  for (const cid of Object.keys(CLASS_TAB)) {
    const inv = byClass[cid];
    if (!inv) continue;
    const commit = commitBy[cid] || inv.commitment || 0;
    const pct = commit / totalCommit;
    const line = (period, fundInt) => {
      const intInc = fundInt * pct;
      return {
        beg: period.beginning,
        contrib: period.contributions,                 // gross (positive nets)
        distrib: -period.returnOfCapital,              // shown positive
        mgmt: period.managementFee,
        partExp: period.netInvestment - intInc,        // 390120 remainder
        intInc,
        place: period.syndication,
      };
    };
    out[cid] = {
      tab: CLASS_TAB[cid], commit,
      remaining: commit - (inv.itd.contributions + inv.itd.returnOfCapital), // net contributed = contrib + (neg ROC)
      QTD: line(inv.q, intQTD),
      YTD: line(inv.ytd, intYTD),
      ending: inv.itd.ending,
    };
  }
  return { quarter, entity_id: eid, sleeves: out };
}

function excelDate(d) { return d instanceof Date ? d : new Date(d + 'T00:00:00Z'); }

async function buildWorkbook(data) {
  const q = data.quarter;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE);
  const yStart = excelDate(q.year_start), qStart = excelDate(q.quarter_start), pEnd = excelDate(q.end);
  for (const cid of Object.keys(data.sleeves)) {
    const s = data.sleeves[cid];
    const ws = wb.getWorksheet(s.tab);
    if (!ws) continue;
    // Header
    ws.getCell('C4').value = r0(s.remaining);
    // Period dates (Inception J9 + E11/G11 stay from template; set the rest)
    ws.getCell('J10').value = yStart;   // Current Year Start
    ws.getCell('J11').value = qStart;   // Current Period Start
    ws.getCell('J12').value = pEnd;     // Period End
    ws.getCell('E11').value = qStart;   // QTD start
    ws.getCell('F11').value = yStart;   // YTD start
    ws.getCell('E12').value = pEnd; ws.getCell('F12').value = pEnd; ws.getCell('G12').value = pEnd;
    // Numeric inputs: E = QTD (col 5), F = YTD (col 6). SI (G) is the template =+F link.
    const put = (row, per) => {
      ws.getCell('E' + row).value = r0(s.QTD[per]);
      ws.getCell('F' + row).value = r0(s.YTD[per]);
    };
    ws.getCell('E16').value = r0(s.QTD.beg); ws.getCell('F16').value = r0(s.YTD.beg); // Beginning NAV (G16 stays 0)
    put(17, 'contrib'); put(18, 'distrib'); put(21, 'mgmt'); put(23, 'partExp'); put(27, 'intInc'); put(32, 'place');
  }
  return wb;
}

function registerIlpaFeeRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/ilpa-fee/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const quarter = pcap.resolveQuarter(body.quarter_end || '');
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const data = buildData(ctx, quarter);
        const wb = await buildWorkbook(data);
        const buf = await fixSheetPrOrder(Buffer.from(await wb.xlsx.writeBuffer()));
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        const totContrib = Object.values(data.sleeves).reduce((s, x) => s + r0(x.YTD.contrib), 0);
        const totEnd = Object.values(data.sleeves).reduce((s, x) => s + r0(x.ending), 0);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-ILPA-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          investors: Object.keys(data.sleeves).length, ytd_contributions: totContrib, ending: totEnd,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { buildData, buildWorkbook, registerIlpaFeeRoutes, FUND_EID, CLASS_TAB };
