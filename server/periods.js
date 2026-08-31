// Period locking: soft-close a month (warn + confirm to post), hard-close a year (block).
//
// Model
//   period_locks rows are per entity. Two levels:
//     level='soft'  period_start..period_end spans one calendar month
//     level='hard'  period_start..period_end spans one calendar year
//   A date is judged against the entity's locks:
//     - inside a hard-closed year            -> HARD_CLOSED (423, blocked in-app)
//     - in a soft-closed month, OR at/before the latest soft-closed month
//       within an otherwise-open year        -> SOFT_CLOSED (409, warn + confirm)
//     - otherwise                            -> postable
//
// Only matched, explicit locks close a period. Nothing is inferred closed that
// wasn't closed by a user. The "at/before latest soft month" rule exists so that
// backdating into an earlier, already-reviewed month of the same open year also
// prompts, not just the specific month someone soft-closed.
//
// Reopening a HARD-closed year is restricted to a named allowlist (below), and is
// a separate admin action never reachable through a posting flow.

const REOPEN_YEAR_ALLOWLIST = [
  'jyun@banyanres.com',
  'ibermudez@banyanres.com',
];

function normEmail(e) { return String(e || '').trim().toLowerCase(); }
function canReopenYear(userEmail) {
  return REOPEN_YEAR_ALLOWLIST.map(normEmail).includes(normEmail(userEmail));
}

// A posting date is 'YYYY-MM-DD'. Derive the month/year window strings.
function monthBounds(dateStr) {
  const y = dateStr.slice(0, 4), m = dateStr.slice(5, 7);
  const start = `${y}-${m}-01`;
  const end = new Date(Date.UTC(+y, +m, 0)).toISOString().slice(0, 10); // last day of month
  return { start, end, year: y, month: m };
}
function yearBounds(year) {
  return { start: `${year}-01-01`, end: `${year}-12-31` };
}

function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS period_locks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
      level TEXT NOT NULL CHECK (level IN ('soft','hard')),
      period_start TEXT NOT NULL,
      period_end TEXT NOT NULL,
      reason TEXT,
      closed_by TEXT NOT NULL,
      closed_at TEXT DEFAULT (datetime('now')),
      UNIQUE (entity_id, level, period_start)
    );
    CREATE TABLE IF NOT EXISTS period_override_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL REFERENCES entities(id) ON DELETE CASCADE,
      entry_date TEXT NOT NULL,
      source TEXT,
      user_email TEXT NOT NULL,
      reason TEXT NOT NULL,
      at TEXT DEFAULT (datetime('now'))
    );
  `);
}

// --- queries -----------------------------------------------------------------

function hardYearLock(db, entityId, year) {
  return db.prepare(
    `SELECT * FROM period_locks WHERE entity_id=? AND level='hard' AND period_start=?`
  ).get(entityId, yearBounds(year).start);
}

function softMonthLock(db, entityId, monthStart) {
  return db.prepare(
    `SELECT * FROM period_locks WHERE entity_id=? AND level='soft' AND period_start=?`
  ).get(entityId, monthStart);
}

// Latest soft-closed month within a given year, as its period_start ('YYYY-MM-01') or null.
function latestSoftMonthInYear(db, entityId, year) {
  const row = db.prepare(
    `SELECT MAX(period_start) AS s FROM period_locks
     WHERE entity_id=? AND level='soft' AND period_start >= ? AND period_start <= ?`
  ).get(entityId, `${year}-01-01`, `${year}-12-31`);
  return row && row.s ? row.s : null;
}

function listLocks(db, entityId) {
  return db.prepare(
    `SELECT id, level, period_start, period_end, reason, closed_by, closed_at
     FROM period_locks WHERE entity_id=? ORDER BY period_start DESC, level`
  ).all(entityId);
}

// --- the single guard --------------------------------------------------------

// Throws { code:'HARD_CLOSED'|'SOFT_CLOSED', period, message } or returns { ok:true }.
// opts: { userEmail, override, reason, source }
function assertPostable(db, entityId, dateStr, opts = {}) {
  entityId = +entityId;
  const date = String(dateStr || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: true }; // malformed dates handled by callers' own validation
  const { year, start: monthStart } = monthBounds(date);

  // 1. Hard-closed year: blocked, no in-app override.
  if (hardYearLock(db, entityId, year)) {
    const e = new Error(`The ${year} fiscal year is closed. Reopening it is a separate admin action.`);
    e.code = 'HARD_CLOSED';
    e.period = { level: 'hard', year };
    throw e;
  }

  // 2. Soft-close: this month, or any month at/before the latest soft-closed month this year.
  const latestSoft = latestSoftMonthInYear(db, entityId, year);
  const inSoftRange = latestSoft && monthStart <= latestSoft;
  if (inSoftRange) {
    if (opts.override) {
      db.prepare(
        `INSERT INTO period_override_log (entity_id, entry_date, source, user_email, reason)
         VALUES (?,?,?,?,?)`
      ).run(entityId, date, opts.source || null, normEmail(opts.userEmail) || 'unknown',
            String(opts.reason || '').trim() || '(no reason given)');
      return { ok: true, overridden: true };
    }
    const closedMonth = softMonthLock(db, entityId, monthStart);
    const e = new Error(
      closedMonth
        ? `${monthStart.slice(0, 7)} is soft-closed. Post anyway?`
        : `${monthStart.slice(0, 7)} is at or before a soft-closed month (${latestSoft.slice(0, 7)}). Post anyway?`
    );
    e.code = 'SOFT_CLOSED';
    e.period = { level: 'soft', month: monthStart.slice(0, 7), latestSoft: latestSoft.slice(0, 7) };
    throw e;
  }

  return { ok: true };
}

// Express helper: turn an assertPostable throw into the right HTTP status.
// Returns true if it handled (responded); false if the error wasn't a period error (rethrow).
function sendPeriodError(res, err) {
  if (err && err.code === 'HARD_CLOSED') {
    res.status(423).json({ error: err.message, code: 'HARD_CLOSED', period: err.period });
    return true;
  }
  if (err && err.code === 'SOFT_CLOSED') {
    res.status(409).json({ error: err.message, code: 'SOFT_CLOSED', period: err.period, needs_override: true });
    return true;
  }
  return false;
}

// --- close / reopen actions --------------------------------------------------

function softClose(db, entityId, monthStr, closedBy, reason) {
  const { start, end } = monthBounds(`${monthStr}-01`);
  db.prepare(
    `INSERT INTO period_locks (entity_id, level, period_start, period_end, reason, closed_by)
     VALUES (?,?,?,?,?,?)
     ON CONFLICT(entity_id, level, period_start)
     DO UPDATE SET period_end=excluded.period_end, reason=excluded.reason,
                   closed_by=excluded.closed_by, closed_at=datetime('now')`
  ).run(+entityId, 'soft', start, end, reason || null, closedBy);
  return { level: 'soft', period_start: start, period_end: end };
}

function reopenSoft(db, entityId, monthStr) {
  const { start } = monthBounds(`${monthStr}-01`);
  const r = db.prepare(
    `DELETE FROM period_locks WHERE entity_id=? AND level='soft' AND period_start=?`
  ).run(+entityId, start);
  return { removed: r.changes };
}

function hardCloseYear(db, entityId, year, closedBy, reason) {
  const { start, end } = yearBounds(year);
  db.prepare(
    `INSERT INTO period_locks (entity_id, level, period_start, period_end, reason, closed_by)
     VALUES (?,?,?,?,?,?)
     ON CONFLICT(entity_id, level, period_start)
     DO UPDATE SET period_end=excluded.period_end, reason=excluded.reason,
                   closed_by=excluded.closed_by, closed_at=datetime('now')`
  ).run(+entityId, 'hard', start, end, reason || null, closedBy);
  return { level: 'hard', period_start: start, period_end: end };
}

function reopenYear(db, entityId, year, userEmail) {
  if (!canReopenYear(userEmail)) {
    const e = new Error('Not authorized to reopen a closed year.');
    e.code = 'FORBIDDEN';
    throw e;
  }
  const { start } = yearBounds(year);
  const r = db.prepare(
    `DELETE FROM period_locks WHERE entity_id=? AND level='hard' AND period_start=?`
  ).run(+entityId, start);
  return { removed: r.changes };
}

module.exports = {
  REOPEN_YEAR_ALLOWLIST,
  canReopenYear,
  ensureSchema,
  assertPostable,
  sendPeriodError,
  softClose,
  reopenSoft,
  hardCloseYear,
  reopenYear,
  listLocks,
};
