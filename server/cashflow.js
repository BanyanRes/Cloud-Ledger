// ─── CLRF workpaper: Statement of Cash Flows worksheet ───────────────────────
//
// A standalone Statement of Cash Flows (indirect method, year-to-date) rendered
// as an .xlsx workpaper. It builds the same statements model the Financial
// Statements Excel export uses (financials.buildStatements) and renders ONLY the
// Cash Flows sheet (financials_xlsx cashFlowOnly), so it is identical to the Cash
// Flows tab of the full financial-statements workbook. Every subtotal is a live
// SUM formula. Filed under Workpapers > Cash Flow > <year> > <quarter>.
const path = require('path');
const fs = require('fs');
const pcap = require('./pcap');
const financials = require('./financials');
const financials_xlsx = require('./financials_xlsx');

function folderFor(q) { return 'Workpapers/Cash Flow/' + q.year + '/' + q.quarter; }
function fileNameFor(q) { return 'CLRF_Cash_Flow_' + q.label + '.xlsx'; }

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

async function buildModel(ctx, eid, asOf, period) {
  const { db } = ctx;
  const ent = db.prepare('SELECT name, code, entity_type FROM entities WHERE id=?').get(eid);
  const entityName = ent ? ent.name : ('Entity ' + eid);
  const getBalances = (o) => Promise.resolve(ctx.computeBalances(eid, o));
  return financials.buildStatements(getBalances, {
    asOf, period: period || 'quarterly', entityName,
    entityCode: ent ? ent.code : '', entityType: ent ? ent.entity_type : '',
    isConsolidated: false, nci: null,
  });
}

function registerCashFlowRoutes(app, ctx) {
  const { auth, requireEntityAccess, requireRole } = ctx;
  app.post('/api/workpapers/cash-flow/:entity_id/generate', auth, requireEntityAccess('entity_id'),
    requireRole('Admin', 'Accountant'), async (req, res) => {
      try {
        const eid = Number(req.params.entity_id);
        const body = req.body || {};
        const asOf = body.quarter_end || '';
        if (!/^\d{4}-\d{2}-\d{2}$/.test(asOf)) throw new Error('quarter_end (YYYY-MM-DD) is required');
        const quarter = pcap.resolveQuarter(asOf);
        const who = (req.user && (req.user.email || req.user.name)) || 'system';
        const model = await buildModel(ctx, eid, asOf, body.period);
        const buf = Buffer.from(await financials_xlsx.buildStatementsWorkbook(model, { cashFlowOnly: true }));
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        const cf = (model && model.cashFlow) || {};
        const n = (v) => (typeof v === 'number' ? v : null);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-CashFlow-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          net_operating: n(cf.netOperating), net_investing: n(cf.netInvesting), net_financing: n(cf.netFinancing),
          net_change: n(cf.netChange), cash_end: n(cf.cashEnd),
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { buildModel, registerCashFlowRoutes };
