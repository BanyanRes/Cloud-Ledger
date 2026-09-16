// ─── CLRF workpaper: Statement of Cash Flows worksheet ───────────────────────
//
// A standalone Statement of Cash Flows (indirect method, year-to-date) for the
// fund, rendered as an .xlsx workpaper. It reuses the exact same model the fund
// financial-statements package builds (financials.buildFundStatements) and
// renders ONLY the Cash Flows sheet (financials_xlsx cashFlowOnly), so this
// worksheet ties to the fund statements by construction. Every subtotal is a
// live SUM formula. Filed under Workpapers > Cash Flow > <year> > <quarter>.
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

// Investment cash activity for the Statement of Cash Flows (mirrors the fund
// statements route): investment-account debits (purchases) / credits (returns)
// in entries that also touch a cash account.
function traceInvestCF(db, eid, asOf) {
  const [iy, im] = String(asOf).split('-').map(Number);
  const ysR = iy + '-01-01';
  const qStartR = iy + ({ 3: '-01-01', 6: '-04-01', 9: '-07-01', 12: '-10-01' }[im] || '-01-01');
  const traceInv = (from, to) => {
    const rows = db.prepare(
      "SELECT jl.debit AS dr, jl.credit AS cr FROM journal_entries je JOIN journal_lines jl ON jl.entry_id = je.id "
      + "WHERE je.entity_id = ? AND je.date >= ? AND je.date <= ? "
      + "AND (jl.account_code LIKE '1201%' OR jl.account_code LIKE '1202%' OR jl.account_code LIKE '1210%' OR jl.account_code LIKE '1218%') "
      + "AND EXISTS (SELECT 1 FROM journal_lines jc WHERE jc.entry_id = je.id AND (jc.account_code LIKE '1002%' OR jc.account_code LIKE '1003%' OR jc.account_code LIKE '1005%' OR jc.account_code LIKE '1072%'))"
    ).all(eid, from, to);
    let purch = 0, ret = 0;
    for (const r of rows) { purch += Number(r.dr) || 0; ret += Number(r.cr) || 0; }
    return { purchases: Math.round(purch * 100) / 100, returns: Math.round(ret * 100) / 100 };
  };
  return { q: traceInv(qStartR, asOf), ytd: traceInv(ysR, asOf) };
}

async function buildModel(ctx, eid, asOf) {
  const { db } = ctx;
  const ent = db.prepare('SELECT name FROM entities WHERE id=?').get(eid);
  const entityName = ent ? ent.name : ('Entity ' + eid);
  const investments = db.prepare('SELECT id, parent_name, name, acquisition_date, cost, fair_value, sort_order '
    + 'FROM fund_investments WHERE entity_id = ? ORDER BY sort_order, id').all(eid);
  const partnerClasses = db.prepare('SELECT id, name, partner_type FROM dim_classes WHERE entity_id = ?').all(eid);
  const commitments = db.prepare('SELECT class_id, commitment_amount FROM investor_commitments WHERE entity_id = ?').all(eid);
  const getBalances = (o) => Promise.resolve(ctx.computeBalances(eid, o));
  let pcapData = null;
  try {
    const quarter = pcap.resolveQuarter(asOf);
    pcapData = pcap.buildData({ db, computeBalances: (e, o) => ctx.computeBalances(e, o) }, quarter, { entity_id: Number(eid) });
  } catch (e) { pcapData = null; }
  let investCF = null;
  try { investCF = traceInvestCF(db, eid, asOf); } catch (e) { investCF = null; }
  return financials.buildFundStatements({ asOf, entityName, getBalances, investments, partnerClasses, commitments, pcap: pcapData, investCF });
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
        const model = await buildModel(ctx, eid, asOf);
        const buf = Buffer.from(await financials_xlsx.buildStatementsWorkbook(model, { cashFlowOnly: true }));
        const saved = saveToWorkpapers(ctx, eid, quarter, buf, who);
        const cf = (model && model.cashFlow) || {};
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="' + saved.original_name + '"');
        res.setHeader('X-CashFlow-Summary', JSON.stringify({
          quarter: quarter.label, saved_to: saved.folder_path + '/' + saved.original_name, replaced: saved.replaced,
          net_operating: cf.netOperating, net_investing: cf.netInvesting, net_financing: cf.netFinancing,
          net_change: cf.netChange, cash_end: cf.cashEnd,
        }).replace(/[^\x20-\x7E]/g, ' '));
        res.send(buf);
      } catch (e) {
        res.status(400).json({ error: e.message });
      }
    });
}

module.exports = { buildModel, registerCashFlowRoutes };
