// ─── Database + files backup (full archive + automatic server-side scheduler) ──
//
// Two things live here:
//
//   1. GET /api/admin/backup/full
//      Streams a single gzip tarball containing a consistent snapshot of the
//      SQLite database PLUS the attachments/ (invoice & bill PDFs) and
//      entity_files/ (workpapers) directories — i.e. EVERYTHING on the data
//      volume. db.backup() takes a live-safe snapshot while writes are in
//      flight; the file directories are read as-is. Auth is identical to the
//      DB-only endpoint in index.js: a BACKUP_TOKEN bearer/query token for
//      unattended automation, or a normal Admin JWT for a browser download.
//      This is what the daily 9am job on Jimmy's PC pulls.
//
//   2. Automatic server-side backup (in-process scheduler).
//      Once per day (after BACKUP_HOUR, Pacific) it writes the same full
//      archive into <data volume>/backups and keeps the most recent
//      BACKUP_RAILWAY_KEEP of them. The data volume is deliberately sized for
//      little more than the live DB, so EVERY write to it is gated by a
//      free-space check: old archives are pruned first, and if a full archive
//      still will not fit, the run falls back to a DB-only snapshot (tiny) so
//      there is always *some* automatic backup and the live app can never be
//      pushed to SQLITE_FULL. Archives are built in the (large, ephemeral) OS
//      temp dir and only the finished file is copied onto the volume.
//
'use strict';

const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');

// Pacific Y-M-D + hour, so "9am" / "2am" mean what Jimmy means.
function pacificNow() {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles', year: 'numeric', month: '2-digit',
    day: '2-digit', hour: '2-digit', hour12: false,
  }).formatToParts(new Date());
  const g = (t) => parts.find((p) => p.type === t).value;
  let hour = parseInt(g('hour'), 10); if (hour === 24) hour = 0;
  return { date: g('year') + '-' + g('month') + '-' + g('day'), hour };
}

// Free bytes on the filesystem that holds `dir`. Prefer fs.statfsSync; fall
// back to `df -kP`; if both fail, return Infinity (don't block on dev/temp).
function freeBytes(dir) {
  try {
    const s = fs.statfsSync(dir);
    return s.bavail * s.bsize;
  } catch (e) {
    try {
      const out = spawnSync('df', ['-kP', dir], { encoding: 'utf8' });
      const line = out.stdout.trim().split('\n').pop().split(/\s+/);
      return parseInt(line[3], 10) * 1024; // available 1K-blocks -> bytes
    } catch (_) { return Infinity; }
  }
}

function cleanupDir(d) { try { fs.rmSync(d, { recursive: true, force: true }); } catch (_) {} }
function cleanupFile(f) { try { fs.unlinkSync(f); } catch (_) {} }

// Create a gzip tarball at outPath: consistent DB snapshot + whichever of the
// file directories exist. Uses the system `tar` (Linux on Railway).
function buildFullArchive(ctx, outPath) {
  const { db, dataDir, attachmentsDir, entityFilesDir } = ctx;
  return new Promise((resolve, reject) => {
    let stage;
    (async () => {
      try {
        stage = fs.mkdtempSync(path.join(os.tmpdir(), 'cl-bkp-'));
        await db.backup(path.join(stage, 'cloudledger.db')); // live-safe snapshot

        const args = ['-czf', outPath, '-C', stage, 'cloudledger.db'];
        const extras = [];
        for (const d of [attachmentsDir, entityFilesDir]) {
          if (d && fs.existsSync(d)) extras.push(path.basename(d));
        }
        // tar's -C is cumulative; the first -C (stage) is absolute, so this one
        // must be absolute too or it resolves relative to the stage dir. dataDir
        // can be a relative path (e.g. "data") on Railway.
        if (extras.length) args.push('-C', path.resolve(dataDir), ...extras);

        const p = spawn('tar', args, { stdio: ['ignore', 'ignore', 'pipe'] });
        let err = '';
        p.stderr.on('data', (b) => { err += b.toString(); });
        p.on('error', (e) => { cleanupDir(stage); reject(e); });
        p.on('close', (code) => {
          cleanupDir(stage);
          // 0 = ok; 1 = "some files changed as we read them" (benign for live
          // attachment dirs) -> treat as success; anything else is a failure.
          if (code === 0 || code === 1) resolve(outPath);
          else reject(new Error('tar exited ' + code + (err ? ': ' + err.slice(0, 300) : '')));
        });
      } catch (e) { if (stage) cleanupDir(stage); reject(e); }
    })();
  });
}

// Backups sorted by name; names are date-stamped so lexical == chronological.
function listBackups(dir, prefix, suffix) {
  try {
    return fs.readdirSync(dir)
      .filter((n) => n.startsWith(prefix) && n.endsWith(suffix))
      .sort()
      .map((n) => ({ name: n, full: path.join(dir, n) }));
  } catch (_) { return []; }
}

// ── scheduler ────────────────────────────────────────────────────────────────
async function tick(ctx) {
  const backupHour = parseInt(process.env.BACKUP_HOUR || '2', 10); // 2am PT (offset from the 9am local pull)
  const keep = Math.max(1, parseInt(process.env.BACKUP_RAILWAY_KEEP || '3', 10));
  const marginBytes = 200 * 1024 * 1024; // keep 200MB head-room on the volume

  try {
    const now = pacificNow();
    if (now.hour < backupHour) return;

    const backupDir = path.join(ctx.dataDir, 'backups');
    fs.mkdirSync(backupDir, { recursive: true });

    const fullName = 'cloudledger-full-' + now.date + '.tar.gz';
    const dbName = 'cloudledger-db-' + now.date + '.db';
    // Already done today? (a full archive OR a db-only fallback both count.)
    if (fs.existsSync(path.join(backupDir, fullName)) || fs.existsSync(path.join(backupDir, dbName))) return;

    // Build the full archive in ephemeral temp first.
    const tmpArchive = path.join(os.tmpdir(), '_cl_full_' + Date.now() + '.tar.gz');
    let built = false;
    try { await buildFullArchive(ctx, tmpArchive); built = true; }
    catch (e) { console.error('[backup] full archive build failed: ' + e.message); cleanupFile(tmpArchive); }

    if (built) {
      const size = fs.statSync(tmpArchive).size;
      // Prune old full archives to (keep-1) to make room for today's.
      const fulls = listBackups(backupDir, 'cloudledger-full-', '.tar.gz');
      while (fulls.length > keep - 1) cleanupFile(fulls.shift().full);
      if (freeBytes(backupDir) > size + marginBytes) {
        fs.copyFileSync(tmpArchive, path.join(backupDir, fullName)); // temp is a different fs -> copy
        cleanupFile(tmpArchive);
        for (const b of listBackups(backupDir, 'cloudledger-db-', '.db')) cleanupFile(b.full); // full supersedes db-only
        console.log('[backup] wrote ' + fullName + ' (' + Math.round(size / 1048576) + ' MB), keeping ' + keep);
        return;
      }
      cleanupFile(tmpArchive);
      console.warn('[backup] full archive (' + Math.round(size / 1048576) + ' MB) will not fit on the volume after pruning; falling back to DB-only');
    }

    // DB-only fallback (tiny, essentially always fits) so there is still a daily backup.
    const tmpDb = path.join(os.tmpdir(), '_cl_dbonly_' + Date.now() + '.db');
    try {
      await ctx.db.backup(tmpDb);
      const size = fs.statSync(tmpDb).size;
      const dbs = listBackups(backupDir, 'cloudledger-db-', '.db');
      while (dbs.length > keep - 1) cleanupFile(dbs.shift().full);
      if (freeBytes(backupDir) > size + marginBytes) {
        fs.copyFileSync(tmpDb, path.join(backupDir, dbName));
        console.log('[backup] wrote ' + dbName + ' (DB-only, ' + Math.round(size / 1048576) + ' MB)');
      } else {
        console.error('[backup] volume too full even for a DB-only snapshot — no automatic backup written');
      }
    } catch (e) { console.error('[backup] DB-only fallback failed: ' + e.message); }
    finally { cleanupFile(tmpDb); }
  } catch (e) {
    console.error('[backup] scheduler error: ' + e.message);
  }
}

function startScheduler(ctx) {
  if (process.env.BACKUP_SCHEDULER === 'off') { console.log('[backup] automatic scheduler disabled by env'); return; }
  setTimeout(() => tick(ctx), 90 * 1000);
  setInterval(() => tick(ctx), 30 * 60 * 1000);
  console.log('[backup] automatic daily backup scheduler started (after '
    + (process.env.BACKUP_HOUR || '2') + ':00 PT, keep ' + (process.env.BACKUP_RAILWAY_KEEP || '3') + ')');
}

// ── auth (identical scheme to the DB-only endpoint in index.js) ───────────────
function makeAuth(ctx) {
  return function backupAuth(req, res, next) {
    const hdr = req.headers['authorization'] || '';
    const token = hdr.startsWith('Bearer ') ? hdr.slice(7) : ((req.query && req.query.token) || '');
    const BACKUP_TOKEN = process.env.BACKUP_TOKEN || '';
    if (BACKUP_TOKEN && token) {
      const a = Buffer.from(String(token)), b = Buffer.from(BACKUP_TOKEN);
      if (a.length === b.length && crypto.timingSafeEqual(a, b)) return next();
    }
    try {
      const payload = ctx.jwt.verify(token, ctx.jwtSecret);
      if (payload.role !== 'Admin') return res.status(403).json({ error: 'Forbidden' });
      req.user = payload; return next();
    } catch (e) { return res.status(401).json({ error: 'Unauthorized' }); }
  };
}

// ── routes ────────────────────────────────────────────────────────────────────
function registerBackupRoutes(app, ctx) {
  const backupAuth = makeAuth(ctx);

  // Full everything-archive (DB snapshot + attachments + entity_files).
  app.get('/api/admin/backup/full', backupAuth, async (req, res) => {
    const stamp = new Date().toISOString().slice(0, 10);
    const tmp = path.join(os.tmpdir(), '_cl_full_dl_' + Date.now() + '.tar.gz');
    try {
      await buildFullArchive(ctx, tmp);
      res.setHeader('Content-Type', 'application/gzip');
      res.setHeader('Content-Disposition', 'attachment; filename="cloudledger-full-' + stamp + '.tar.gz"');
      const stream = fs.createReadStream(tmp);
      stream.pipe(res);
      const done = () => cleanupFile(tmp);
      stream.on('close', done);
      stream.on('error', () => { done(); if (!res.headersSent) res.status(500).end(); });
    } catch (e) {
      cleanupFile(tmp);
      console.error('[backup] full download failed: ' + e.message);
      if (!res.headersSent) res.status(500).json({ error: 'Full backup failed: ' + e.message });
    }
  });

  // Status of the automatic backups currently on the volume.
  app.get('/api/admin/backup/status', backupAuth, (req, res) => {
    const backupDir = path.join(ctx.dataDir, 'backups');
    const list = [...listBackups(backupDir, 'cloudledger-full-', '.tar.gz'), ...listBackups(backupDir, 'cloudledger-db-', '.db')]
      .map((b) => { const s = fs.statSync(b.full); return { name: b.name, size_mb: Math.round(s.size / 1048576), mtime: s.mtime.toISOString() }; });
    res.json({ backup_dir: backupDir, free_mb: Math.round(freeBytes(backupDir) / 1048576), count: list.length, backups: list });
  });

  startScheduler(ctx);
}

module.exports = { registerBackupRoutes };
