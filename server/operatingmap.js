// ═══════════════════════════════════════════════════════════════════════════
// Operating TB Account Mapping — a standalone home, under Accounting, for the
// one job the (retired) Consolidation page used to bury: mapping a property
// manager's operating trial balance to CloudLedger statement lines.
//
// SCOPE: Braker and HP only — the two projects that hand leasing to a property
// manager keeping its own general ledger. Every month the manager adds new
// sub-accounts, so a map built once drifts out of date and new accounts arrive
// unmapped. This module lets an accountant upload a TB, see what is mapped and
// what is not, and map/correct accounts in one discoverable place.
//
// The data is the SAME operating_tb / operating_tb_map tables the consolidation
// engine reads (server/consolidation.js owns their schema and the roll-up), so
// existing maps carry over untouched and anything saved here flows into any
// consolidating schedule still in use. Two columns are added for this page:
//   • confirmed — 0 for an auto-mapped row awaiting a human's OK, 1 otherwise.
//                 Auto-mapped rows still flow into the schedules (so a close is
//                 never blocked), they are just surfaced for review.
//   • source    — 'manual' | 'auto', so an auto-map is distinguishable from a
//                 hand-mapped one and from a confirmed correction.
//
// Auto-mapping never guesses. A new account is proposed a target only on a
// signal that cannot reasonably be wrong: an identical account name already
// mapped (a rename/renumber), a collection-loss / contra account whose base fee
// is mapped (it nets to the same line), or a property-system account family
// whose every mapped member already lands on one line. Anything else is left
// unmapped for a person to map — the same stance the consolidation engine takes
// everywhere else (report the difference, never plug it).
// ═══════════════════════════════════════════════════════════════════════════

const multer = require('multer');

const C = require('./consolidation');
const { parseOperatingTb } = require('./consolidation_routes');
const { r2, monthEnd } = C._helpers;

const tbUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } });
const TYPES = ['Asset', 'Liability', 'Equity', 'Revenue', 'Expense'];
const isType = (t) => TYPES.includes(t);

// ── Schema additions (the base tables live in consolidation.js) ──
function ensureSchema(db) {
  C.ensureSchema(db); // idempotent; guarantees operating_tb / operating_tb_map exist
  for (const [col, def] of [
    ['confirmed', 'INTEGER NOT NULL DEFAULT 1'],
    ['source', "TEXT NOT NULL DEFAULT 'manual'"],
  ]) {
    try { db.exec('ALTER TABLE operating_tb_map ADD COLUMN ' + col + ' ' + def); }
    catch (e) { if (!/duplicate column/i.test(e.message)) throw e; }
  }
}

// ── Books: the operating ledgers that have a mapping (HP, Braker) ──
// Resolved from the consolidation setup while it exists, with a fallback that
// scans the TB data itself so the page keeps working if that setup is removed.
function books(db) {
  let rows = [];
  try {
    rows = db.prepare(
      "SELECT g.parent_entity_id, g.scope_key, g.name parent_name, m.entity_id oper_eid, m.label "
      + "FROM consol_members m JOIN consol_groups g ON g.id = m.group_id "
      + "WHERE m.source = 'tb' AND g.scope_key IN ('hp','braker') ORDER BY g.scope_key"
    ).all();
  } catch (e) { rows = []; }
  if (rows.length) return rows;
  // Fallback: every entity_id that carries an operating TB, labelled from the
  // entities table where it is a real id, else a generic label.
  const out = [];
  for (const r of db.prepare('SELECT DISTINCT entity_id FROM operating_tb').all()) {
    const ent = db.prepare('SELECT name FROM entities WHERE id = ?').get(r.entity_id);
    out.push({ parent_entity_id: r.entity_id < 0 ? -r.entity_id : null, scope_key: null,
      parent_name: null, oper_eid: r.entity_id, label: ent ? ent.name : ('Operating ' + r.entity_id) });
  }
  return out;
}
function bookFor(db, operEid) { return books(db).find((b) => Number(b.oper_eid) === Number(operEid)) || null; }

// ── Name helpers ──
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
// A collection-loss / contra account nets against the fee it offsets; strip the
// contra prefix to recover that base fee's name.
const CONTRA_RX = /^(collection\s*loss|bad\s*debt|allowance(\s+for)?)\s*[-:]?\s*/i;
const isContra = (name) => CONTRA_RX.test(String(name || ''));
const baseName = (name) => String(name || '').replace(CONTRA_RX, '').trim();
// Property-system code families: AAAA and AAAA.BBB from AAAA.BBB.CCC (either
// separator). A family is a group of sub-accounts of one roll-up account.
const seg = (code) => String(code).split(/[.\-]/);
const fam4 = (code) => seg(code)[0] || String(code);
const famBBB = (code) => seg(code).slice(0, 2).join('.') || String(code);

// ── Auto-mapper ──
// Returns the proposals it wrote (confirmed = 0). Never throws on a single bad
// line; a failure to propose just leaves that account unmapped.
function autoMap(db, operEid, asOf, by) {
  const eom = monthEnd(asOf);
  const unmapped = C.unmappedFor(db, operEid, eom);
  if (!unmapped.length) return [];
  const map = C.mapFor(db, operEid); // source_code -> row (target_code/name/type)

  // Indexes over what is already mapped.
  const byName = new Map();          // normalized source name -> target row
  const byTargetName = new Map();    // normalized target line name -> its target
  const famBBBTargets = new Map();   // AAAA.BBB -> Set(target_code)
  const fam4Targets = new Map();     // AAAA -> Set(target_code)
  const targetByCode = new Map();    // target_code -> {target_code,name,type}
  for (const [code, m] of map) {
    const nm = norm(m.source_name || db.prepare('SELECT source_name FROM operating_tb WHERE entity_id=? AND source_code=? ORDER BY as_of DESC LIMIT 1').get(operEid, code)?.source_name);
    if (nm) byName.set(nm, m);
    const t = { target_code: m.target_code, target_name: m.target_name, target_type: m.target_type };
    targetByCode.set(String(m.target_code), t);
    const tn = norm(m.target_name);
    if (tn) byTargetName.set(tn, t);
    const bk = famBBB(code), fk = fam4(code);
    if (!famBBBTargets.has(bk)) famBBBTargets.set(bk, new Map());
    if (!fam4Targets.has(fk)) fam4Targets.set(fk, new Map());
    famBBBTargets.get(bk).set(String(m.target_code), (famBBBTargets.get(bk).get(String(m.target_code)) || 0) + 1);
    fam4Targets.get(fk).set(String(m.target_code), (fam4Targets.get(fk).get(String(m.target_code)) || 0) + 1);
  }
  const unanimous = (famMap) => {
    if (!famMap) return null;
    const codes = [...famMap.keys()];
    const members = [...famMap.values()].reduce((s, n) => s + n, 0);
    if (codes.length === 1 && members >= 2) return targetByCode.get(codes[0]) || null;
    return null;
  };

  const now = new Date().toISOString();
  const ins = db.prepare(
    'INSERT INTO operating_tb_map (entity_id, source_code, target_code, target_name, target_type, notes, confirmed, source, updated_at, updated_by) '
    + "VALUES (?,?,?,?,?,?,0,'auto',?,?)");
  const log = db.prepare('INSERT INTO operating_tb_map_log (entity_id, source_code, old_target, new_target, changed_at, changed_by) VALUES (?,?,?,?,?,?)');

  const proposals = [];
  const propose = (u, t, basis) => {
    if (!t || !t.target_code || !isType(t.target_type)) return false;
    try {
      ins.run(operEid, u.source_code, String(t.target_code), String(t.target_name || ''), t.target_type,
        'Auto-mapped — ' + basis + '. Confirm or change.', now, by || 'system');
      log.run(operEid, u.source_code, null, String(t.target_code), now, (by || 'system') + ' (auto)');
      proposals.push({ source_code: u.source_code, source_name: u.source_name, target_code: t.target_code,
        target_name: t.target_name, target_type: t.target_type, basis });
      return true;
    } catch (e) { return false; }
  };

  const tx = db.transaction(() => {
    for (const u of unmapped) {
      const nm = norm(u.source_name);
      // 1. Exact match to an existing statement line's name. CLA names its lines
      //    after the source accounts, so this is the primary intended mapping.
      if (nm && byTargetName.has(nm)) { if (propose(u, byTargetName.get(nm), 'matches the statement line "' + byTargetName.get(nm).target_name + '"')) continue; }
      // 2. Identical name already mapped on another source code (rename / renumber).
      if (nm && byName.has(nm)) { if (propose(u, byName.get(nm), 'matches the mapped account "' + u.source_name + '"')) continue; }
      // 3. Contra / collection-loss account → the line of the base fee it offsets.
      if (isContra(u.source_name)) {
        const bn = norm(baseName(u.source_name));
        const base = byTargetName.get(bn) || byName.get(bn);
        if (base && propose(u, base, 'offsets "' + baseName(u.source_name) + '", which is mapped')) continue;
      }
      // 4. Unanimous family: every mapped sibling lands on one line.
      const ut = unanimous(famBBBTargets.get(famBBB(u.source_code))) || unanimous(fam4Targets.get(fam4(u.source_code)));
      if (ut && propose(u, ut, 'every mapped account in its group maps to ' + ut.target_code + ' ' + ut.target_name)) continue;
    }
  });
  tx();
  return proposals;
}

// ── Read model for a book + month ──
function viewFor(db, operEid, asOf) {
  const months = C.tbMonths(db, operEid);
  const eom = asOf ? monthEnd(asOf) : (months[0] || {}).as_of;
  const map = C.mapFor(db, operEid);
  const mapMeta = new Map();
  for (const r of db.prepare('SELECT source_code, confirmed, source, notes FROM operating_tb_map WHERE entity_id = ?').all(operEid)) mapMeta.set(String(r.source_code), r);
  let lines = [];
  if (eom) {
    lines = C.tbAt(db, operEid, eom).map((l) => {
      const m = map.get(String(l.source_code)); const meta = mapMeta.get(String(l.source_code)) || {};
      return { source_code: l.source_code, source_name: l.source_name, ending: r2(l.ending), activity: r2(l.debit - l.credit),
        target_code: m ? m.target_code : null, target_name: m ? m.target_name : null, target_type: m ? m.target_type : null,
        confirmed: m ? (meta.confirmed == null ? 1 : meta.confirmed) : null, source: m ? (meta.source || 'manual') : null,
        notes: m ? (meta.notes || null) : null };
    }).sort((a, b) => String(a.source_code).localeCompare(String(b.source_code)));
  }
  const unmapped = eom ? C.unmappedFor(db, operEid, eom) : [];
  const pending = lines.filter((l) => l.target_code && !Number(l.confirmed));
  return { oper_eid: operEid, as_of: eom || null, months, lines, unmapped, pending };
}

// ── Routes ──
function registerOperatingMapRoutes(app, deps) {
  const { db, auth, requireRole, userHasEntityAccess, workpapersDir } = deps;
  ensureSchema(db);
  const gate = [auth, requireRole('Admin', 'Accountant')];
  const who = (req) => (req.user && (req.user.email || req.user.name)) || 'system';

  // Resolve the book for :oper_eid and check the caller may see it (access is to
  // the PARENT entity — the operating column may be a synthetic id with no row).
  function resolveBook(req, res) {
    const operEid = Number(req.params.oper_eid);
    const bk = bookFor(db, operEid);
    if (!bk) { res.status(404).json({ error: 'Not an operating mapping book' }); return null; }
    const gateEid = bk.parent_entity_id || operEid;
    if (!userHasEntityAccess(req.user.id, req.user.role, gateEid)) { res.status(403).json({ error: 'No access to ' + (bk.parent_name || gateEid) }); return null; }
    return bk;
  }

  // Which books this user can open.
  app.get('/api/operating-map/books', ...gate, (req, res) => {
    try {
      const out = [];
      for (const b of books(db)) {
        const gateEid = b.parent_entity_id || b.oper_eid;
        if (!userHasEntityAccess(req.user.id, req.user.role, gateEid)) continue;
        const months = C.tbMonths(db, b.oper_eid);
        const latest = (months[0] || {}).as_of || null;
        const unmapped = latest ? C.unmappedFor(db, b.oper_eid, latest) : [];
        const pending = db.prepare('SELECT COUNT(*) n FROM operating_tb_map WHERE entity_id = ? AND confirmed = 0').get(b.oper_eid).n;
        out.push({ oper_eid: b.oper_eid, label: b.label, scope_key: b.scope_key,
          parent_entity_id: b.parent_entity_id, parent_name: b.parent_name,
          latest, month_count: months.length,
          map_rows: db.prepare('SELECT COUNT(*) n FROM operating_tb_map WHERE entity_id = ?').get(b.oper_eid).n,
          unmapped_count: unmapped.length, pending_count: pending });
      }
      res.json({ books: out });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // A book + month: TB lines with their mapping, plus unmapped and the
  // auto-mapped-pending set.
  app.get('/api/operating-map/:oper_eid', ...gate, (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      res.json(Object.assign({ label: bk.label, parent_name: bk.parent_name }, viewFor(db, bk.oper_eid, req.query.as_of ? String(req.query.as_of) : null)));
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Upload a TB for a month → parse, store, auto-map new accounts.
  app.post('/api/operating-map/:oper_eid/upload', ...gate, tbUpload.single('file'), (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const as_of = String(req.body.as_of || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(as_of)) return res.status(400).json({ error: 'as_of must be YYYY-MM-DD' });
      const eom = monthEnd(as_of);
      const parsed = parseOperatingTb(req.file.buffer);
      const now = new Date().toISOString();
      db.transaction(() => {
        db.prepare('DELETE FROM operating_tb WHERE entity_id = ? AND as_of = ?').run(bk.oper_eid, eom);
        const ins = db.prepare('INSERT INTO operating_tb (entity_id, as_of, source_code, source_name, forward, debit, credit, ending, uploaded_at, uploaded_by, filename) VALUES (?,?,?,?,?,?,?,?,?,?,?)');
        for (const l of parsed.lines) ins.run(bk.oper_eid, eom, l.source_code, l.source_name, l.forward, l.debit, l.credit, l.ending, now, who(req), req.file.originalname || null);
      })();
      // File the source document under the parent's workpapers for traceability.
      let workpaper = null;
      try {
        if (workpapersDir && bk.parent_entity_id) {
          const path = require('path'); const fs = require('fs');
          const folder = 'Operating TBs/' + eom; const by = who(req);
          for (const fp of ['Operating TBs', folder]) { try { db.prepare('INSERT INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)').run(bk.parent_entity_id, fp, by); } catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; } }
          const ext = (String(req.file.originalname || '').match(/\.[A-Za-z0-9]+$/) || ['.xlsx'])[0];
          const originalName = eom + ' ' + bk.label + ' TB' + ext;
          const dir = path.join(workpapersDir, String(bk.parent_entity_id)); fs.mkdirSync(dir, { recursive: true });
          for (const old of db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?').all(bk.parent_entity_id, folder, originalName)) { try { fs.unlinkSync(path.join(dir, old.stored_filename)); } catch (e) {} db.prepare('DELETE FROM entity_files WHERE id = ?').run(old.id); }
          const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + originalName.replace(/[^a-zA-Z0-9._-]/g, '_');
          fs.writeFileSync(path.join(dir, stored), req.file.buffer);
          db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)').run(bk.parent_entity_id, folder, stored, originalName, req.file.size, req.file.mimetype || null, by);
          workpaper = { folder_path: folder, file_name: originalName };
        }
      } catch (e) { console.error('[operating-map] TB filing failed:', e.message); }

      const proposals = autoMap(db, bk.oper_eid, eom, who(req));
      const unmapped = C.unmappedFor(db, bk.oper_eid, eom);
      res.json({ success: true, as_of: eom, count: parsed.lines.length, balanced: parsed.balanced,
        residual: parsed.residual, sign_mode: parsed.signMode,
        auto_mapped_count: proposals.length, auto_mapped: proposals,
        unmapped_count: unmapped.length, unmapped_nonzero: unmapped.filter((u) => Math.abs(u.ending) > 0.004),
        workpaper });
    } catch (e) { res.status(e.status || 400).json({ error: e.message }); }
  });

  // Re-run auto-mapping for a month on demand (e.g. after adding a base mapping
  // that lets a contra account resolve).
  app.post('/api/operating-map/:oper_eid/auto', ...gate, (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      const eom = monthEnd(String(req.body.as_of || (C.tbMonths(db, bk.oper_eid)[0] || {}).as_of || ''));
      if (!eom) return res.status(400).json({ error: 'No trial balance on file' });
      const proposals = autoMap(db, bk.oper_eid, eom, who(req));
      res.json({ success: true, as_of: eom, auto_mapped_count: proposals.length, auto_mapped: proposals });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Upsert / confirm / remove mappings. Each row: { source_code, target_code,
  // target_name, target_type, confirmed?, _delete? }. Saving a target always
  // marks the row confirmed (a human touched it); _delete removes it.
  app.put('/api/operating-map/:oper_eid/map', ...gate, (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      const rows = Array.isArray(req.body && req.body.rows) ? req.body.rows : [];
      const now = new Date().toISOString(); const by = who(req);
      let saved = 0, removed = 0;
      db.transaction(() => {
        for (const r of rows) {
          const src = String(r.source_code || '').trim(); if (!src) continue;
          const prev = db.prepare('SELECT * FROM operating_tb_map WHERE entity_id = ? AND source_code = ?').get(bk.oper_eid, src);
          if (r._delete) {
            if (prev) { db.prepare('DELETE FROM operating_tb_map WHERE id = ?').run(prev.id);
              db.prepare('INSERT INTO operating_tb_map_log (entity_id, source_code, old_target, new_target, changed_at, changed_by) VALUES (?,?,?,?,?,?)').run(bk.oper_eid, src, prev.target_code, null, now, by); removed++; }
            continue;
          }
          const tc = String(r.target_code || '').trim();
          const tn = String(r.target_name || '').trim();
          const tt = String(r.target_type || '').trim();
          // A bare confirm (no target change) on an existing row.
          if (!tc && prev && r.confirmed) { db.prepare("UPDATE operating_tb_map SET confirmed = 1, source = CASE WHEN source = 'auto' THEN 'auto-confirmed' ELSE source END, updated_at = ?, updated_by = ? WHERE id = ?").run(now, by, prev.id); saved++; continue; }
          if (!tc || !tn || !isType(tt)) throw Object.assign(new Error('Row for ' + src + ' needs target code, name and a type of ' + TYPES.join('/')), { status: 400 });
          if (prev) db.prepare("UPDATE operating_tb_map SET target_code=?, target_name=?, target_type=?, confirmed=1, source=CASE WHEN source='auto' THEN 'auto-confirmed' ELSE source END, updated_at=?, updated_by=? WHERE id=?").run(tc, tn, tt, now, by, prev.id);
          else db.prepare("INSERT INTO operating_tb_map (entity_id, source_code, target_code, target_name, target_type, confirmed, source, updated_at, updated_by) VALUES (?,?,?,?,?,1,'manual',?,?)").run(bk.oper_eid, src, tc, tn, tt, now, by);
          if (!prev || prev.target_code !== tc) db.prepare('INSERT INTO operating_tb_map_log (entity_id, source_code, old_target, new_target, changed_at, changed_by) VALUES (?,?,?,?,?,?)').run(bk.oper_eid, src, prev ? prev.target_code : null, tc, now, by);
          saved++;
        }
      })();
      res.json(Object.assign({ success: true, saved, removed }, viewFor(db, bk.oper_eid, req.body && req.body.as_of)));
    } catch (e) { res.status(e.status || 400).json({ error: e.message }); }
  });

  // Mapping-change history for a book.
  app.get('/api/operating-map/:oper_eid/log', ...gate, (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      res.json({ log: db.prepare('SELECT * FROM operating_tb_map_log WHERE entity_id = ? ORDER BY id DESC LIMIT 200').all(bk.oper_eid) });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Remove a month's uploaded TB.
  app.delete('/api/operating-map/:oper_eid', ...gate, (req, res) => {
    try { const bk = resolveBook(req, res); if (!bk) return;
      const eom = monthEnd(String(req.query.as_of || ''));
      const r = db.prepare('DELETE FROM operating_tb WHERE entity_id = ? AND as_of = ?').run(bk.oper_eid, eom);
      res.json({ success: true, deleted: r.changes });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });
}

module.exports = { registerOperatingMapRoutes, autoMap, ensureSchema, books };
