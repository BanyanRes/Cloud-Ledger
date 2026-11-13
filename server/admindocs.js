// ─── Administration document folders ─────────────────────────────────────────
//
// Each document-generating Administration tool gets a folder of the documents it
// has produced, saved on CloudLedger's own data volume (no SharePoint). Saving is
// automatic: a response-capture middleware watches the generate/download routes of
// the nine tools and files whatever file they send. The sidebar renders a folder
// icon next to each item; clicking it opens the tool's folder browser.
//
// The browser works like the entity Workpapers browser: subfolders, upload,
// download, delete, download-all. The Bank Account Application tool files each
// entity's packet (and any supporting documents added later) under a subfolder
// named for that entity, so one entity's complete bank-opening package lives
// together.
//
// Storage: files live flat on disk at <dataDir>/admin-docs/<tool>/<stored-name>;
// the subfolder a file belongs to is metadata in admin_documents.folder, and
// empty folders are tracked in admin_doc_folders. Re-generating the same-named
// document in the same folder replaces the prior one (newest wins).
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const memUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 60 * 1024 * 1024 } });

const TOOL_KEYS = ['assignment', 'usc_capital', 'muha_invoice', 'am_fee',
  'compliance_certs', 'time_allocation', 'letterhead', 'dissolution', 'bank_account_app'];
// Generate/download route(s) whose file output is filed into each tool's folder.
// originalUrl is matched (query strings and path params are fine).
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

// A folder path: '/'-separated segments. Slashes within a typed name are turned
// into '-' per segment; leading/trailing slashes and empty segments are dropped.
function normFolder(s) {
  return String(s || '').split('/').map((x) => x.replace(/\\/g, '-').trim()).filter(Boolean).join('/').slice(0, 300);
}
// Every ancestor path of a folder, inclusive (a/b/c -> a, a/b, a/b/c).
function ancestors(folder) {
  const parts = normFolder(folder).split('/').filter(Boolean);
  const out = [];
  for (let i = 0; i < parts.length; i++) out.push(parts.slice(0, i + 1).join('/'));
  return out;
}

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
  // folder column added after the fact for installs created before subfolders.
  try {
    const cols = db.prepare('PRAGMA table_info(admin_documents)').all().map((c) => c.name);
    if (!cols.includes('folder')) db.exec("ALTER TABLE admin_documents ADD COLUMN folder TEXT DEFAULT ''");
  } catch { /* ignore */ }
  db.exec(`CREATE TABLE IF NOT EXISTS admin_doc_folders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tool TEXT NOT NULL,
    folder TEXT NOT NULL,
    created_by TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(tool, folder)
  )`);
  try { db.exec('CREATE INDEX IF NOT EXISTS idx_admindocs_tool ON admin_documents(tool)'); } catch { /* ignore */ }
}

// File a generated/uploaded document into a tool's folder, replacing any prior
// document of the same name in the same folder (newest wins).
function saveAdminDoc(db, dataDir, tool, buf, originalName, who, mime, folder) {
  if (!TOOL_KEYS.includes(tool) || !buf || !buf.length) return;
  const fold = normFolder(folder);
  const dir = dirFor(dataDir, tool);
  try { fs.mkdirSync(dir, { recursive: true }); } catch { /* ignore */ }
  const name = String(originalName || (tool + '.pdf')).replace(/[/\\]/g, '_').trim().slice(0, 200) || (tool + '.pdf');
  const prior = db.prepare('SELECT id, stored_name FROM admin_documents WHERE tool = ? AND IFNULL(folder, \'\') = ? AND original_name = ?').all(tool, fold, name);
  for (const p of prior) {
    try { fs.unlinkSync(path.join(dir, p.stored_name)); } catch { /* already gone */ }
    db.prepare('DELETE FROM admin_documents WHERE id = ?').run(p.id);
  }
  const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + name.replace(/[^A-Za-z0-9._-]/g, '_');
  fs.writeFileSync(path.join(dir, stored), buf);
  db.prepare('INSERT INTO admin_documents (tool, folder, original_name, stored_name, mime, size, created_by, created_at) '
    + "VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))").run(tool, fold, name, stored, mime || null, buf.length, who || null);
}

// The distinct folders of a tool: every folder (and ancestor) that holds a file,
// plus every explicitly-created folder.
function listFolders(db, tool) {
  const set = new Set();
  for (const r of db.prepare('SELECT DISTINCT IFNULL(folder, \'\') AS folder FROM admin_documents WHERE tool = ?').all(tool)) {
    for (const a of ancestors(r.folder)) set.add(a);
  }
  for (const r of db.prepare('SELECT folder FROM admin_doc_folders WHERE tool = ?').all(tool)) {
    for (const a of ancestors(r.folder)) set.add(a);
  }
  return Array.from(set).filter(Boolean).sort();
}

// Response-capture middleware: for the tools' generate/download routes, wrap
// res.send so any successful file response is also filed into the tool's folder.
// Bank Account Application files under a subfolder named for the entity.
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
          // Bank applications are grouped per entity (the name typed on generate).
          let folder = '';
          if (hit.tool === 'bank_account_app') { try { folder = normFolder((req.body && req.body.entity) || ''); } catch { folder = ''; } }
          const who = (req.user && (req.user.email || req.user.name)) || 'system';
          setImmediate(() => { try { saveAdminDoc(db, dataDir, hit.tool, buf, name, who, ct, folder); } catch { /* best-effort */ } });
        }
      } catch { /* never block the response */ }
      return orig(body);
    };
    next();
  });
}

// Folder API: browse / upload / download / delete a tool's saved documents.
function registerAdminDocsRoutes(app, ctx) {
  const { db, auth, requireRole, dataDir } = ctx;
  ensureTable(db);
  const BOTH = ['Admin', 'Accountant'];

  // Browse: files (each with its folder) + the folder tree.
  app.get('/api/admin/docs/:tool', auth, requireRole(...BOTH), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const files = db.prepare('SELECT id, IFNULL(folder, \'\') AS folder, original_name, mime, size, created_by, created_at '
        + 'FROM admin_documents WHERE tool = ? ORDER BY created_at DESC, id DESC').all(tool);
      res.json({ tool, files, folders: listFolders(db, tool) });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Upload one or more documents into a folder (defaults to root).
  app.post('/api/admin/docs/:tool/upload', auth, requireRole(...BOTH), memUpload.array('files', 20), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      if (!TOOL_KEYS.includes(tool)) return res.status(400).json({ error: 'Unknown tool' });
      const folder = normFolder((req.body && req.body.folder) || '');
      const who = (req.user && (req.user.email || req.user.name)) || 'system';
      const files = req.files || [];
      if (!files.length) return res.status(400).json({ error: 'No files uploaded.' });
      for (const f of files) saveAdminDoc(db, dataDir, tool, f.buffer, f.originalname, who, f.mimetype, folder);
      res.json({ uploaded: files.length, folder });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Create an (empty) folder.
  app.post('/api/admin/docs/:tool/folders', auth, requireRole(...BOTH), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const folder = normFolder((req.body && req.body.folder) || '');
      if (!folder) return res.status(400).json({ error: 'Folder name required' });
      db.prepare('INSERT OR IGNORE INTO admin_doc_folders (tool, folder, created_by) VALUES (?, ?, ?)')
        .run(tool, folder, (req.user && (req.user.email || req.user.name)) || 'system');
      res.json({ created: true, folder });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Delete a folder (must be empty of files, itself and descendants).
  app.delete('/api/admin/docs/:tool/folders', auth, requireRole(...BOTH), (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const folder = normFolder(req.query.folder || '');
      if (!folder) return res.status(400).json({ error: 'Folder required' });
      const n = db.prepare('SELECT COUNT(*) AS n FROM admin_documents WHERE tool = ? AND (IFNULL(folder, \'\') = ? OR IFNULL(folder, \'\') LIKE ?)')
        .get(tool, folder, folder + '/%').n;
      if (n > 0) return res.status(400).json({ error: 'Folder is not empty (' + n + ' file(s)).' });
      db.prepare('DELETE FROM admin_doc_folders WHERE tool = ? AND (folder = ? OR folder LIKE ?)').run(tool, folder, folder + '/%');
      res.json({ deleted: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Download one document.
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

  // Download every file at/under a folder as one ZIP (root '' = the whole tool).
  app.get('/api/admin/docs/:tool/download-all', auth, requireRole(...BOTH), async (req, res) => {
    try {
      const tool = safeTool(req.params.tool);
      const folder = normFolder(req.query.folder || '');
      const rows = db.prepare('SELECT * FROM admin_documents WHERE tool = ? ORDER BY folder, original_name').all(tool)
        .filter((f) => {
          const fp = String(f.folder || '');
          return folder === '' ? true : (fp === folder || fp.startsWith(folder + '/'));
        });
      if (!rows.length) return res.status(404).json({ error: 'This folder has no files to download.' });
      const JSZip = require('jszip');
      const zip = new JSZip();
      const dir = dirFor(dataDir, tool);
      for (const f of rows) {
        const abs = path.join(dir, f.stored_name);
        if (!fs.existsSync(abs)) continue;
        const rel = folder === '' ? String(f.folder || '') : String(f.folder || '').slice(folder.length).replace(/^\//, '');
        zip.file((rel ? rel + '/' : '') + f.original_name, fs.readFileSync(abs));
      }
      const out = await zip.generateAsync({ type: 'nodebuffer' });
      const base = (folder ? folder.split('/').pop() : tool).replace(/[^A-Za-z0-9._-]+/g, '_');
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename="' + base + '.zip"');
      res.send(out);
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
