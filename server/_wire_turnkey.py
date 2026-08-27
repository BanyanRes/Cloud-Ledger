import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/index.js'
src = open(p, encoding='utf-8').read()

# Add require near the top
old_req = "const fs = require('fs');"
new_req = "const fs = require('fs');\nconst turnkey = require('./turnkey');"
if old_req not in src:
    print('MISSING require anchor'); sys.exit(1)
src = src.replace(old_req, new_req, 1)

# Find the very end of the file (before module.exports if any, or before app.listen)
# We'll insert routes right before the listen call.
anchor = "// === Bill.com integration routes ==="
if anchor not in src:
    print('MISSING billcom routes anchor'); sys.exit(1)

routes = '''// === Turnkey Rail integration routes ===

// All routes use API key auth, NOT JWT.
const turnkeyAuth = turnkey.apiKeyAuth(db);

// Health check (no auth — useful for Turnkey to verify connectivity)
app.get('/api/turnkey/health', (req, res) => {
  res.json({ status: 'ok', integration: 'turnkey-rail', timestamp: new Date().toISOString() });
});

// === API key management (these use JWT/Admin, not API key) ===

// List API keys (no raw key visible, only metadata)
app.get('/api/admin/api-keys', auth, requireRole('Admin'), (req, res) => {
  const rows = db.prepare(
    'SELECT id, key_prefix, name, scopes, last_used_at, created_by, created_at, revoked_at FROM api_keys ORDER BY id DESC'
  ).all();
  res.json(rows);
});

// Create a new API key. Returns the raw key ONCE — admin must save it.
app.post('/api/admin/api-keys', auth, requireRole('Admin'), (req, res) => {
  const name = (req.body.name || '').trim();
  const scopes = (req.body.scopes || 'turnkey:sync').trim();
  if (!name) return res.status(400).json({ error: 'name required' });
  const rawKey = turnkey.generateApiKey();
  const hash = turnkey.hashApiKey(rawKey);
  const prefix = turnkey.apiKeyPrefix(rawKey);
  const now = new Date().toISOString();
  const result = db.prepare(
    'INSERT INTO api_keys (key_hash, key_prefix, name, scopes, created_by, created_at) VALUES (?, ?, ?, ?, ?, ?)'
  ).run(hash, prefix, name, scopes, req.user.email, now);
  res.json({ id: result.lastInsertRowid, raw_key: rawKey, key_prefix: prefix, name, scopes,
             warning: 'Save this key now. It will never be shown again.' });
});

// Revoke an API key
app.post('/api/admin/api-keys/:id/revoke', auth, requireRole('Admin'), (req, res) => {
  db.prepare('UPDATE api_keys SET revoked_at = ? WHERE id = ?').run(new Date().toISOString(), req.params.id);
  res.json({ success: true });
});

// === Project linking (Turnkey calls these with its API key) ===

// Link Turnkey project to CL entity + seed POC chart of accounts
app.post('/api/turnkey/projects/link', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  try {
    const { turnkey_project_id, project_code, project_name } = req.body;
    if (!turnkey_project_id || !project_code || !project_name) {
      return res.status(400).json({ error: 'turnkey_project_id, project_code, project_name required' });
    }
    const map = turnkey.linkProject(db, { turnkey_project_id, project_code, project_name });
    res.json({ ok: true, project_map: map });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Get project mapping (for Turnkey to verify the link exists)
app.get('/api/turnkey/projects/:id', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const map = db.prepare('SELECT * FROM turnkey_project_map WHERE turnkey_project_id = ?').get(req.params.id);
  if (!map) return res.status(404).json({ error: 'Project not linked' });
  res.json(map);
});

// === Sync event endpoints ===
// All accept JSON payload; all return { ok, cl_entry_id, idempotent } on success.

function syncRoute(syncFn) {
  return (req, res) => {
    try {
      const result = syncFn(db, req.body || {});
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  };
}

app.post('/api/turnkey/sync/sub-payapp-approved',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncSubPayAppApproved));

app.post('/api/turnkey/sync/sub-payapp-paid',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncSubPayAppPaid));

app.post('/api/turnkey/sync/owner-payapp-issued',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncOwnerPayAppIssued));

app.post('/api/turnkey/sync/owner-payment-received',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncOwnerPaymentReceived));

app.post('/api/turnkey/sync/month-end-poc',
  turnkeyAuth, turnkey.requireScope('turnkey:sync'),
  syncRoute(turnkey.syncMonthEndPOC));

// View sync log for a project (last 50 events)
app.get('/api/turnkey/sync-log/:turnkey_project_id', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  const map = db.prepare('SELECT cl_entity_id FROM turnkey_project_map WHERE turnkey_project_id = ?').get(req.params.turnkey_project_id);
  if (!map) return res.status(404).json({ error: 'Project not linked' });
  const rows = db.prepare(
    'SELECT id, sync_type, turnkey_id, cl_entry_id, status, message, created_at FROM turnkey_sync_log ' +
    'WHERE cl_entity_id = ? ORDER BY id DESC LIMIT 50'
  ).all(map.cl_entity_id);
  res.json(rows);
});

'''
src = src.replace(anchor, routes + anchor, 1)

open(p, 'w', encoding='utf-8').write(src)
print('OK')
