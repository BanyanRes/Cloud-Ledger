// ─── Administration document folders ─────────────────────────────────────────
//
// Each document-generating Administration tool gets a folder of the documents it
// has produced, saved on CloudLedger's own data volume (no SharePoint). Saving is
// automatic: a response-capture middleware watches the generate/download routes of
// the nine tools and files whatever file they send. The sidebar renders a folder
// icon next to each of those items; clicking it opens the tool's folder (list /
// download / delete) served by the routes below.
//
// Storage: <dataDir>/admin-docs/<tool>/<stored-name> on disk, indexed by the
// admin_documents table. Re-generating the same-named document replaces the prior
// one (newest wins), so repeated generation doesn't pile up duplicates.
const fs = require('fs');
const path = require('path');

// The nine document tools, and the generate/download route(s) whose file output
// is filed into each tool's folder. originalUrl is matched (so query strings and
// path params are fine).
const TOOL_KEYS = ['assignment', 'usc_capital', 'muha_invoice', 'am_fee',
  'compliance_certs', 'time_allocation', 'letterhead', 'dissolution', 'bank_account_app'];
const DOC_ROUTES = [
  { re: /^\/api\/assignments\/generate(?:[/?]|$)/, tool: 'assignment' },
  { re: /^\/api\/admin\/usc-capital\/statement\.pdf(?:[/?]|$)/, tool: 'usc_capital' },
  { re: /^\/api\/admin\/usc-capital\/statements\.zip(?:[/?]|$)/, tool: 'usc_capital' },
  { re: /^\/api\/admin\/muha-invoice\/[^/]+\/pdf(?:[/?]|$)/, tool: 'muha_invoice' },
  { re: /^\/api\/admin\/am-fee\/[^/]+\/preview\.pdf(?:[/?]|$)/, tool: 'am_fee' },
  { re: /^\/api\/admin\/compliance-certs\/report(?:[/?]|$)/, tool: 'compliance_certs' },
  { re: /^\/api\/admin\/time-allocation\/report(?:[/?]|$)/, tool: 'time_allocation' },
  { re: /^\/api\/letterhead\/generate(?:[/?]|$)/, tool: 'letterhead' },
  { re: /^\/api\/admin\/dissolution\/generate(?:[/?]|$)/, tool: 'dissolution' },
  { re: /^\/api\/admin\/bankapp\/generate(?:[/?]|$)/, tool: 'bank_account_app' },
];

const safeTool = (t) => String(t || '').replace(/[^a-z0-9_]/gi, '').slice(0, 40);
const dirFor = (dataDir, tool) => path.join(dataDir, 'admin-docs', safeTool(tool));

function ensureTable(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS admin_documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tool TEXT NOT NULL,
    original_name TEXT NOT NULL,
    stored_name TEXT NOT NULL,
    mime TEXT,
    size INTEGER,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now'))
  )`);
  try { db.exec('CREATE INDEX IF NOT EXISTS idx_admindocs_tool ON admin_documents(tool)'); } catch { /* ignore */ }
}

// File a generated document into a tool's folder, replacing any prior document of
// the same name (newest wins).
function saveAdminDoc(db, dataDir, tool, buf, originalName, who, mime) {
  if (!TOOL_KEYS.includes(tool) || !buf || !buf.length) return;
  const dir = dirFor(dataDir, tool);
  try { fs.mkdirSync(dir, { recursive: true }); } catch { /* ignore */ }
  const name = String(originalName || (tool + '.pdf')).replace(/[/\\]/g, '_').trim().slice(0, 200) || (tool + '.pdf');
  const prior = db.prepare('SELECT id, stored_name FROM admin_documents WHERE tool = ? AND original_name = ?').all(tool, name);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(dir, p.stored_name)); } catch { /* already gone */ }
    db.prepare('DELETE FROM admin_documents WHERE id = ?').run(p.id);
  }
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + name.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO admin_documents (tool, original_name, stored_name, mime, size, created_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, datetime('now'))").run(tool, name, stored, mime || null, buf.length, who || null);
}

// Response-capture middleware: for the nine tools' generate/download routes, wrap
// res.send so any successful file response is also filed into the tool's folder.
// Registered once, early (before the tool routes), so it precedes them.
function registerCaptureMiddleware(app, ctx) {
  const { db, dataDir } = ctx;
  ensureTable(db);
  app.use((req, res, next) => {
    let hit = null;
    for (const d of DOC_ROUTES) { if (d.re.test(req.originalUrl || req.url)) { hit = d; break; } }
    if (!hit) return next();
    const orig = res.send.bind(res);
    res.send = function (body) {
      try {
        const code = res.statusCode || 200;
        const buf = Buffer.isBuffer(body) ? body : null;
        const ct = String(res.getHeader('Content-Type') || '');
        if (code >= 200 && code < 300 && buf && buf.length > 100 && !/application\/json|text\/html/i.test(ct)) {
          const cd = String(res.getHeader('Content-Disposition') || '');
          const m = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(cd);
          let name;
          try { name = m ? decodeURIComponent(m[1]) : null; } catch { name = m ? m[1] : null; }
          if (!name) { name = hit.tool + '_' + Date.now() + (/\bpdf\b/i.test(ct) ? '.pdf' : /zip/i.test(ct) ? '.zip' : ''); }
          const who = (req.user && (req.user.email || req.user.name)) || 'system';
          setImmediate(() => { try { saveAdminDoc(db, dataDir, hit.tool, buf, name, who, ct); } catch { /* best-effort */ } });
        }
      } catch { /* never block the response */ }
      return orig(body);
    };
    next();
  });
}

// Folder API: list / download / delete a tool's saved documents.
function registerAdminDocsRoutes(app, ctx) {
  const { db, auth, requireRole, dataDir } = ctx;
  ensureTable(db);
  const BOTH = ['Admin', 'Accountant'];

  app.get('/api/admin/docs/:tool', auth, requireRole(...BOTH), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const rows = db.prepare('SELECT id, original_name, mime, size, created_by, created_at FROM admin_documents '
        + 'WHERE tool = ? ORDER BY created_at DESC, id DESC').all(tool);
      res.json({ tool, rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/admin/docs/:tool/:id/download', auth, requireRole(...BOTH), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const row = db.prepare('SELECT * FROM admin_documents WHERE id = ? AND tool = ?').get(req.params.id, tool);
      if (!row) return res.status(404).json({ error: 'Not found' });
      const abs = path.join(dirFor(dataDir, tool), row.stored_name);
      if (!fs.existsSync(abs)) return res.status(404).json({ error: 'File missing on disk' });
      res.setHeader('Content-Type', row.mime || 'application/octet-stream');
      res.setHeader('Content-Disposition', 'attachment; filename="' + String(row.original_name).replace(/"/g, '') + '"');
      res.send(fs.readFileSync(abs));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.delete('/api/admin/docs/:tool/:id', auth, requireRole('Admin'), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const row = db.prepare('SELECT * FROM admin_documents WHERE id = ? AND tool = ?').get(req.params.id, tool);
      if (row) {
        try { fs.unlinkSync(path.join(dirFor(dataDir, tool), row.stored_name)); } catch { /* already gone */ }
        db.prepare('DELETE FROM admin_documents WHERE id = ?').run(row.id);
      }
      res.json({ deleted: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
}

module.exports = { registerCaptureMiddleware, registerAdminDocsRoutes, saveAdminDoc, TOOL_KEYS };
