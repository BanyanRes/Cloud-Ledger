// ═══════════════════════════════════════════════════════════════════════════
// Consolidation — operating trial balance upload, account mapping, and the
// elimination engine behind the consolidating balance sheet / statement of
// operations.
//
// SCOPE: Braker and HP only (Jimmy, 2026-08-27). The operating trial balance
// upload exists because those two projects hand leasing to a property
// management company that keeps its own general ledger. Every other entity in
// CloudLedger keeps its own books here and needs none of this.
//
// The shape of the problem, using Braker:
//
//   Braker QOZ Business (45)  — the parent. Holds the QOF capital and the
//                               investment in the property company.
//   Braker Prop Co      (56)  — the development ledger. Land, hard costs, soft
//                               costs, the Arbor loan.
//   Braker Operating    (57)  — the property manager's ledger (Foxtail by
//                               Banyan). Lease-up operations only.
//
// During lease-up the development entity funds the property manager's operating
// account: development debits a soft-cost account and credits cash; the property
// manager debits its bank and credits contributed capital. The same dollars are
// therefore a capitalised cost on one ledger and equity on the other, and both
// sides have to come out on consolidation or the cost is counted twice and
// equity shows a contribution that never came from outside the group.
//
// Two eliminations do the whole job:
//
//   1. investment ↔ capital.  The parent's investment account against the
//      development entity's contributed capital from the parent. Reciprocal,
//      one account each side.
//   2. funding ↔ capital.  The operating ledger's contributed capital against
//      the development-entity accounts that were debited when the transfers
//      went out. There is no way to infer which accounts those were — a
//      transfer looks like any other capitalised cost — so the offsetting
//      accounts and amounts are USER-MAINTAINED (consol_funding_accounts),
//      seeded from the accounts already identified on Braker. The operating
//      ledger's contributed-capital balance is the authoritative total: any gap
//      between it and the sum of the funding accounts is reported as
//      unassigned, never plugged.
//
// HP is a different shape, and simpler (Jimmy, 2026-08-27):
//
//   Bridge Banyan HP QOZB (43) — the parent. Holds the investment in the
//                                property company.
//   HP Property Owner     (55) — the development ledger.
//   Highpoint Operating        — the property manager's ledger, uploaded.
//
// No cash passes between development and operating there, so rule 2 does not
// apply and there is nothing to match. Instead the property manager MIRRORS the
// whole development book inside its own trial balance — construction in
// progress, the construction loan, retention, deployed fund capital, the
// mortgage interest. The same dollars therefore appear on the development
// column and again on the operating column, and the mirror is simply removed:
//
//   3. mirrored development accounts.  A user-maintained list of operating
//      accounts that come out IN FULL, for whatever they carry that month. The
//      list is fixed month to month; only the balances move, which is what lets
//      the same elimination run every month with no re-derivation. Unlike rule
//      2 nothing is paired, so the removal can be one-sided — CLA's own July
//      schedule is one-sided by 47,199,612.00 because their operating column is
//      out of balance by the same amount. The rule reports that as its residual
//      rather than hiding it; the schedule still cross-foots either way,
//      because the column carries the same imbalance the elimination does.
//
// Nothing here posts a journal entry. Eliminations are a reporting overlay, so
// the member ledgers stay exactly as they were entered and re-running a prior
// month always reproduces the same schedule.
// ═══════════════════════════════════════════════════════════════════════════


const r2 = n => Math.round((Number(n) || 0) * 100) / 100;
const isDrType = t => t === 'Asset' || t === 'Expense';
const isBsType = t => t === 'Asset' || t === 'Liability' || t === 'Equity';
const isPlType = t => t === 'Revenue' || t === 'Expense';
// Synthetic line that carries the HP development-mirror net-income plug. It is
// not a real GL account: an Expense-typed adjustment on this code folds into
// (revenue - expense) net income on every balance sheet and prints on no
// income statement, which is where a plug measured from stock balances belongs.
const NI_PLUG_CODE = '__NI_MIRROR_PLUG__';
const NI_PLUG_NAME = 'Net Income (Loss) — development mirror';

// ══════════════════════════════ Scope ══════════════════════════════

// Matched on name and code rather than a hardcoded id so a restored or
// renumbered database still resolves. Braker is the parent of the Braker group;
// Bridge Banyan HP QOZB is the parent of the HP group.
const SCOPED_PARENTS = [
  { key: 'braker', name: /braker\s*qoz\s*business/i, codes: ['BRAKERQO1'] },
  { key: 'hp', name: /bridge\s*banyan\s*hp|hp\s*qozb/i, codes: ['BRIDGEBA'] },
];

function scopeKeyFor(ent) {
  if (!ent) return null;
  const n = String(ent.name || '');
  const c = String(ent.code || '').toUpperCase();
  for (const s of SCOPED_PARENTS) if (s.name.test(n) || s.codes.includes(c)) return s.key;
  return null;
}

// ══════════════════════════════ Schema ══════════════════════════════

function ensureSchema(db) {
  db.exec(`
    -- A consolidation group is a REPORT definition: which ledgers form the
    -- columns of the consolidating schedules, in what order, under what
    -- headings. Distinct from org_nodes, which records who owns whom.
    CREATE TABLE IF NOT EXISTS consol_groups (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      parent_entity_id INTEGER NOT NULL UNIQUE,
      scope_key TEXT NOT NULL,
      name TEXT,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT
    );

    -- source = 'ledger'  : column comes from this entity's journal entries
    -- source = 'tb'      : column comes from an uploaded operating trial balance
    CREATE TABLE IF NOT EXISTS consol_members (
      group_id INTEGER NOT NULL,
      entity_id INTEGER NOT NULL,
      label TEXT,
      source TEXT NOT NULL DEFAULT 'ledger',
      sort_order INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (group_id, entity_id)
    );

    -- One uploaded operating trial balance, stored line for line in the
    -- property system's own account codes. All four columns are kept because
    -- each answers a different question:
    --   forward : balance-sheet accounts -> prior month end
    --             income-statement accounts -> year to date through prior month
    --   debit/credit : this month's activity
    --   ending  : balance-sheet accounts -> this month end
    --             income-statement accounts -> year to date through this month
    -- Values are stored DEBIT-POSITIVE exactly as a trial balance reads, so the
    -- file can be tied back to line by line. The sign convention CloudLedger
    -- reports in is applied when the map is used, not on the way in.
    CREATE TABLE IF NOT EXISTS operating_tb (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      as_of TEXT NOT NULL,
      source_code TEXT NOT NULL,
      source_name TEXT,
      forward REAL NOT NULL DEFAULT 0,
      debit REAL NOT NULL DEFAULT 0,
      credit REAL NOT NULL DEFAULT 0,
      ending REAL NOT NULL DEFAULT 0,
      uploaded_at TEXT, uploaded_by TEXT, filename TEXT,
      UNIQUE (entity_id, as_of, source_code)
    );
    CREATE INDEX IF NOT EXISTS idx_otb_entity ON operating_tb(entity_id, as_of);

    -- Property-system account -> statement line. Applied at READ time, so
    -- correcting a mapping fixes every period at once rather than only the
    -- months uploaded after the change. Keyed on the FULL source code including
    -- any sub-account suffix: 51030-000 Bonuses and 51030-001 Quarterly Bonuses
    -- are different statement lines.
    CREATE TABLE IF NOT EXISTS operating_tb_map (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      source_code TEXT NOT NULL,
      target_code TEXT NOT NULL,
      target_name TEXT NOT NULL,
      target_type TEXT NOT NULL,
      notes TEXT,
      updated_at TEXT, updated_by TEXT,
      UNIQUE (entity_id, source_code)
    );

    -- Every change to a mapping, so a restated month can be explained.
    CREATE TABLE IF NOT EXISTS operating_tb_map_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      source_code TEXT NOT NULL,
      old_target TEXT, new_target TEXT,
      changed_at TEXT, changed_by TEXT
    );

    -- The development-entity accounts that were debited when cash went to the
    -- property manager. User-maintained: nothing in the ledger distinguishes a
    -- funding transfer from an ordinary capitalised cost.
    --   mode 'amount' : eliminate the stored cumulative amount
    --   mode 'full'   : eliminate the account's whole balance (use only where
    --                   the account exists solely to carry funding)
    CREATE TABLE IF NOT EXISTS consol_funding_accounts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      group_id INTEGER NOT NULL,
      entity_id INTEGER NOT NULL,
      account_code TEXT NOT NULL,
      account_name TEXT,
      mode TEXT NOT NULL DEFAULT 'amount',
      amount REAL NOT NULL DEFAULT 0,
      notes TEXT,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT,
      UNIQUE (group_id, entity_id, account_code)
    );

    -- The reciprocal investment/capital pair. One row per parent-subsidiary
    -- step, so HP's deeper chain is more rows rather than different code.
    CREATE TABLE IF NOT EXISTS consol_investment_pairs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      group_id INTEGER NOT NULL,
      label TEXT,
      holder_entity_id INTEGER NOT NULL,
      holder_account_code TEXT NOT NULL,
      issuer_entity_id INTEGER NOT NULL,
      issuer_account_code TEXT NOT NULL,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT, created_by TEXT,
      UNIQUE (group_id, holder_entity_id, holder_account_code)
    );

    -- Which operating-ledger account carries the contributed capital that the
    -- funding accounts have to agree to. Stated rather than guessed.
    CREATE TABLE IF NOT EXISTS consol_funding_capital (
      group_id INTEGER NOT NULL,
      entity_id INTEGER NOT NULL,
      account_code TEXT NOT NULL,
      PRIMARY KEY (group_id, entity_id, account_code)
    );

    -- Accounts that MIRROR another column and are removed in full (HP). The
    -- property manager keeps a copy of the whole development book inside its
    -- own trial balance, so the same dollars are already in the development
    -- column. Nothing is paired and no amount is stored: the listed account
    -- comes out for whatever it carries in the window being built, which is why
    -- the same list runs unchanged every month.
    --
    -- account_code is the STATEMENT-LINE code, i.e. the target the operating
    -- account maps to, not the property system's own source code — eliminations
    -- run after the map has been applied.
    CREATE TABLE IF NOT EXISTS consol_full_eliminations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      group_id INTEGER NOT NULL,
      entity_id INTEGER NOT NULL,
      account_code TEXT NOT NULL,
      account_name TEXT,
      notes TEXT,
      -- One row may be flagged the balancer: it absorbs the mirror set's
      -- residual so the elimination column foots and the consolidated balance
      -- sheet balances (HP plugs it into mortgage interest, as CLA does).
      is_balancer INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT,
      UNIQUE (group_id, entity_id, account_code)
    );
  `);
  // Older databases created the table before is_balancer existed.
  try { db.exec('ALTER TABLE consol_full_eliminations ADD COLUMN is_balancer INTEGER NOT NULL DEFAULT 0'); }
  catch (e) { if (!/duplicate column/i.test(e.message)) throw e; }
}

// ══════════════════════════ Date helpers ══════════════════════════

const yearStart = d => String(d).slice(0, 4) + '-01-01';
const monthStart = d => String(d).slice(0, 7) + '-01';

function monthEnd(dateStr) {
  const [y, m] = String(dateStr).split('-').map(Number);
  return new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
}

function prevMonthEnd(dateStr) {
  const [y, m] = String(dateStr).split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, 0)).toISOString().slice(0, 10);
}

// ══════════════════════ Operating TB storage ══════════════════════

function tbMonths(db, entityId) {
  return db.prepare(`SELECT as_of, COUNT(*) AS lines, MAX(uploaded_at) AS uploaded_at,
                            MAX(uploaded_by) AS uploaded_by, MAX(filename) AS filename
                     FROM operating_tb WHERE entity_id = ? GROUP BY as_of ORDER BY as_of DESC`)
    .all(entityId);
}

function tbAt(db, entityId, as_of) {
  return db.prepare('SELECT * FROM operating_tb WHERE entity_id = ? AND as_of = ?').all(entityId, as_of);
}

function hasTbAt(db, entityId, as_of) {
  const r = db.prepare('SELECT 1 FROM operating_tb WHERE entity_id = ? AND as_of = ? LIMIT 1').get(entityId, as_of);
  return !!r;
}

function mapFor(db, entityId) {
  const m = new Map();
  for (const r of db.prepare('SELECT * FROM operating_tb_map WHERE entity_id = ?').all(entityId)) {
    m.set(String(r.source_code), r);
  }
  return m;
}

// Source codes present in an uploaded TB with no mapping. A non-zero unmapped
// account is a real difference against the property manager's statements, not
// noise, so it is reported rather than dropped quietly.
function unmappedFor(db, entityId, as_of) {
  const map = mapFor(db, entityId);
  return tbAt(db, entityId, as_of)
    .filter(l => !map.has(String(l.source_code)))
    .map(l => ({
      source_code: l.source_code, source_name: l.source_name,
      ending: r2(l.ending), activity: r2(l.debit - l.credit),
    }))
    .sort((a, b) => Math.abs(b.ending) - Math.abs(a.ending));
}

// Roll an uploaded TB up to statement lines. Returns rows in exactly the shape
// computeBalances returns, so a caller cannot tell a mapped TB column from a
// real ledger column.
//   pick(line) -> the debit-positive figure to use from that line
function rollUp(db, entityId, lines, pick) {
  const map = mapFor(db, entityId);
  const byTarget = new Map();
  for (const l of lines) {
    const m = map.get(String(l.source_code));
    if (!m) continue;
    const v = Number(pick(l)) || 0;
    if (!v) {
      // Keep the line so the statement can print a zero row where the CPA
      // package prints one, but never invent a target that has no source.
      if (!byTarget.has(m.target_code)) {
        byTarget.set(m.target_code, { code: m.target_code, name: m.target_name, type: m.target_type, balance: 0, total_debit: 0, total_credit: 0 });
      }
      continue;
    }
    // Debit-positive in, natural convention out: assets and expenses keep their
    // sign, liabilities, equity and revenue flip. This is what turns the
    // property system's debit-positive vacancy loss into the negative that
    // prints inside revenue.
    const natural = isDrType(m.target_type) ? v : -v;
    const cur = byTarget.get(m.target_code);
    if (cur) cur.balance = r2(cur.balance + natural);
    else byTarget.set(m.target_code, { code: m.target_code, name: m.target_name, type: m.target_type, balance: r2(natural), total_debit: 0, total_credit: 0 });
  }
  return [...byTarget.values()];
}

// The getBalances contract, served from uploaded trial balances.
//
//   { as_of, close_pl_before } — balance sheet snapshot. Balance-sheet accounts
//       take the ending column; income accounts take activity since
//       close_pl_before. Prior-year earnings need no separate computation: the
//       property system's own retained-earnings account already carries them,
//       which is why the operating column's retained earnings ties.
//   { from, to } — activity in a window.
//
// A window the uploads cannot answer returns null so the caller can say the
// column is unavailable instead of printing zeros that look like real nils.
function operatingBalances(db, entityId, o = {}) {
  const { as_of, from, to, close_pl_before } = o;

  if (as_of) {
    const d = monthEnd(as_of);
    if (!hasTbAt(db, entityId, d)) return null;
    const lines = tbAt(db, entityId, d);
    const plFloor = close_pl_before || yearStart(d);
    // Income accounts: the ending column is year to date, so anything other
    // than a year-start floor needs the earlier month's cumulative figure
    // subtracted.
    let base = null;
    if (plFloor > yearStart(d)) {
      const bd = prevMonthEnd(plFloor);
      if (!hasTbAt(db, entityId, bd)) return null;
      base = new Map(tbAt(db, entityId, bd).map(l => [String(l.source_code), l.ending]));
    }
    const rows = rollUp(db, entityId, lines, l => {
      const code = String(l.source_code);
      const lead = parseInt(code.replace(/[^0-9]/g, '').slice(0, 1), 10);
      const isPl = lead >= 4;
      if (!isPl) return l.ending;
      if (!base) return l.ending;
      return l.ending - (base.get(code) || 0);
    });
    // Closing the books into retained earnings. When close_pl_before falls in a
    // LATER fiscal year than this snapshot — an opening balance sheet built from
    // a prior-year trial balance (e.g. a Dec-2025 operating TB used as the 2026
    // opening) — the pre-floor income was zeroed above, but its net result must
    // be CLOSED INTO RETAINED EARNINGS, or the opening column is short that
    // amount and the cash-flow statement shows a phantom equity movement. The
    // property system leaves its Dec-31 books pre-closing (current-year earnings
    // still sit in the P&L accounts, retained earnings at the prior-year
    // figure), so do the close here. Net income = -(sum of pre-floor income
    // endings): a loss reduces RE, a gain raises it. Runs only on that opening
    // snapshot — the current and prior period statements use a year-start floor
    // (base null) and are untouched, as is any entity with no prior-year TB.
    if (base) {
      const map = mapFor(db, entityId);
      let netIncome = 0;
      for (const [code, ending] of base) {
        const lead = parseInt(String(code).replace(/[^0-9]/g, '').slice(0, 1), 10);
        if (lead < 4) continue;                 // balance-sheet account, not P&L
        if (!map.has(String(code))) continue;   // unmapped — reported elsewhere
        netIncome = r2(netIncome - ending);
      }
      if (Math.abs(netIncome) > 0.004) {
        let re = rows.find(r => r.type === 'Equity' && /retained/i.test(r.name || ''));
        if (!re) { re = { code: '39000', name: 'Retained Earnings', type: 'Equity', balance: 0, total_debit: 0, total_credit: 0 }; rows.push(re); }
        re.balance = r2(re.balance + netIncome);
      }
    }
    return rows;
  }

  if (from && to) {
    const d = monthEnd(to);
    if (!hasTbAt(db, entityId, d)) return null;
    const lines = tbAt(db, entityId, d);
    // A window that lies inside the month of `to` is that month's own
    // debit/credit activity — no second upload needed, and it is the figure the
    // property manager reports. The test is that the window STARTS on or after
    // the first of that month; `from <= monthStart` would swallow year to date.
    if (from >= monthStart(d) && to >= d) {
      return rollUp(db, entityId, lines, l => l.debit - l.credit);
    }
    if (from <= yearStart(d)) {
      // Year to date. Income accounts read straight off the ending column;
      // balance-sheet movement is ending less the opening carry-forward, which
      // for a January upload is the forward column.
      const janFloor = prevMonthEnd(yearStart(d));
      const haveFloor = hasTbAt(db, entityId, janFloor);
      const floor = haveFloor ? new Map(tbAt(db, entityId, janFloor).map(l => [String(l.source_code), l.ending])) : null;
      return rollUp(db, entityId, lines, l => {
        const code = String(l.source_code);
        const lead = parseInt(code.replace(/[^0-9]/g, '').slice(0, 1), 10);
        if (lead >= 4) return l.ending;
        return floor ? (l.ending - (floor.get(code) || 0)) : 0;
      });
    }
    // An arbitrary multi-month window needs the month before it started.
    const bd = prevMonthEnd(from);
    if (!hasTbAt(db, entityId, bd)) return null;
    const base = new Map(tbAt(db, entityId, bd).map(l => [String(l.source_code), l.ending]));
    return rollUp(db, entityId, lines, l => l.ending - (base.get(String(l.source_code)) || 0));
  }

  if (to) return operatingBalances(db, entityId, { from: yearStart(to), to });
  return null;
}

// ══════════════════════════ Group access ══════════════════════════

function groupForParent(db, parentEntityId) {
  return db.prepare('SELECT * FROM consol_groups WHERE parent_entity_id = ?').get(parentEntityId);
}

// Resolve the group an entity belongs to, whether it is the parent OR one of the
// ledger members. Lets a user generate the consolidated package from the
// development entity they work in (e.g. HP Property Owner) and get the parent-
// titled group package, not that member's standalone statements.
function groupForEntity(db, entityId) {
  const asParent = groupForParent(db, entityId);
  if (asParent) return asParent;
  const mem = db.prepare('SELECT group_id FROM consol_members WHERE entity_id = ?').get(entityId);
  if (!mem) return null;
  return db.prepare('SELECT * FROM consol_groups WHERE id = ?').get(mem.group_id);
}

function membersOf(db, groupId) {
  return db.prepare('SELECT * FROM consol_members WHERE group_id = ? ORDER BY sort_order, entity_id').all(groupId);
}

// Column order is parent first, then the members in their configured order —
// the order the CPA package prints.
function columnsOf(db, group) {
  const parent = { entity_id: group.parent_entity_id, label: null, source: 'ledger', sort_order: -1 };
  const rows = membersOf(db, group.id).filter(m => m.entity_id !== group.parent_entity_id);
  return [parent, ...rows];
}

function memberBalances(db, member, o, computeBalances) {
  if (member.source === 'tb') return operatingBalances(db, member.entity_id, o);
  return computeBalances(member.entity_id, o) || [];
}

// ═══════════════════════════ Eliminations ═══════════════════════════

// Look up one account's figure for the requested window, in natural terms.
function figureFor(rows, code) {
  const r = (rows || []).find(x => String(x.code) === String(code));
  return r ? r2(r.balance) : 0;
}

// ── Forward funding derivation (Braker) ────────────────────────────────────
// The funding elimination (rule 2) removes the operating ledger's contributed
// capital against the development accounts that carry the transferred cash.
// Instead of hand-maintaining which development accounts those are as new
// transfers post, derive each new transfer straight from the development ledger
// — exactly the way the accountant does it by hand:
//
//   1. The transfer amount is the increase in operating contributed capital
//      (the gap between it and what the funding map already eliminates).
//   2. Find the development journal entry that recorded the transfer — a credit
//      to the development bank account for that amount (the cash that left).
//   3. The account(s) DEBITED on that entry are what must be eliminated.
//   4. If a debited account was a clearing/holding account later reclassed
//      (e.g. 10119 "Cash - Operating I" → 13425 "Operating Shortfall"), follow
//      the reclass to where it lands, so a live balance is eliminated rather
//      than a zeroed clearing account.
//
// This runs ONLY when contributed capital has grown beyond the mapped funding
// (a new transfer landed) and is purely additive — when everything already ties
// (the recurring shortfall self-captures because 13425 is a `full` account) the
// gap is zero and this does nothing, so prior periods reproduce unchanged. The
// development bank account the transfers leave from, per scoped group:
const FORWARD_FUNDING_CASH = { braker: '10166' };

// Follow a reclass chain. `code` was DEBITED for `amount` on the transfer entry;
// if a LATER entry (through as_of) CREDITED that same account for that amount,
// the cash was reclassed out of it, so recurse into the reclass entry's debit
// legs; otherwise `code` itself is where it landed. Returns [{ code, amount,
// via }] — `via` records the account(s) it passed through, for transparency.
// Depth-capped and it never revisits the account it just left, so it cannot loop.
function followReclass(db, entityId, code, amount, asOf, depth) {
  depth = depth || 0;
  const here = [{ code: String(code), amount: r2(amount), via: null }];
  if (depth > 6) return here;
  const dateClause = asOf ? ' AND je.date <= ?' : '';
  const params = asOf ? [entityId, code, asOf, amount] : [entityId, code, amount];
  const rc = db.prepare(
    `SELECT je.id AS entry_id FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
      WHERE je.entity_id = ? AND jl.account_code = ? AND jl.credit > 0${dateClause}
        AND ABS(jl.credit - ?) < 0.01
      ORDER BY je.date ASC, je.id ASC`
  ).all(...params);
  if (!rc.length) return here;
  const debits = db.prepare(
    `SELECT jl.account_code AS code, jl.debit AS amount FROM journal_lines jl
      WHERE jl.entry_id = ? AND jl.debit > 0 AND jl.account_code <> ?`
  ).all(rc[0].entry_id, code);
  if (!debits.length) return here;
  const out = [];
  for (const d of debits) {
    for (const n of followReclass(db, entityId, d.code, r2(d.amount), asOf, depth + 1)) {
      out.push({ code: n.code, amount: n.amount, via: n.via ? String(code) + '→' + n.via : String(code) });
    }
  }
  return out;
}

// Derive funding legs for the outstanding `gap` (contributed capital not yet
// eliminated by the map) from the development ledger `devEntityId`. Matches the
// gap against credits to the development bank account (newest first, largest
// that still fits the remaining gap), reads each matched entry's debit legs, and
// follows any reclass to the landing account. Returns { legs, matched, residual,
// transfers }: `legs` are extra funding legs to add; `residual` is any part of
// the gap no transfer entry accounted for (reported, never plugged).
function deriveForwardFunding(db, group, o, devEntityId, gap) {
  const out = { legs: [], matched: 0, residual: r2(gap), transfers: [] };
  const cashCode = FORWARD_FUNDING_CASH[group.scope_key];
  if (!cashCode || gap <= 0.01) return out;
  const asOf = o && o.as_of ? monthEnd(o.as_of) : null;
  const dateClause = asOf ? ' AND je.date <= ?' : '';
  const params = asOf ? [devEntityId, cashCode, asOf] : [devEntityId, cashCode];
  const cands = db.prepare(
    `SELECT je.id AS entry_id, je.date AS date, jl.credit AS amount
       FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id
      WHERE je.entity_id = ? AND jl.account_code = ? AND jl.credit > 0${dateClause}
      ORDER BY je.date DESC, je.id DESC`
  ).all(...params);

  // The clean case is one transfer per reporting period: a single credit to the
  // development bank for exactly the gap. Match that, newest first. Requiring an
  // EXACT single-entry match is deliberate — it can never mistake an ordinary
  // vendor payment out of the same bank that merely happens to fit the gap for a
  // funding transfer. If several transfers landed in one period, or the gap does
  // not equal any single entry, nothing is derived and the whole gap is reported
  // as unassigned (visible, never plugged) for the user to resolve.
  const exact = cands.find(c => Math.abs(r2(c.amount) - gap) < 0.01);
  if (exact) {
    const amt = r2(exact.amount);
    const debits = db.prepare(
      `SELECT jl.account_code AS code, jl.debit AS amount FROM journal_lines jl
        WHERE jl.entry_id = ? AND jl.debit > 0`
    ).all(exact.entry_id);
    const landed = [];
    for (const d of debits) {
      for (const t of followReclass(db, devEntityId, d.code, r2(d.amount), asOf, 0)) landed.push(t);
    }
    if (landed.length) {
      for (const l of landed) out.legs.push({ code: l.code, amount: r2(l.amount) });
      out.transfers.push({ entry_id: exact.entry_id, date: exact.date, amount: amt, accounts: landed });
      out.matched = amt;
      out.residual = r2(gap - amt);
    }
  }
  return out;
}

// Build the elimination column for a window.
//
// Each rule computes its own two sides from the ledgers and eliminates the
// MATCHED amount, reporting any difference as a residual. A residual is never
// absorbed into the elimination column — if the two sides disagree the schedule
// still cross-foots and the disagreement is visible.
function computeEliminations(db, group, o, computeBalances, rowsByEntity) {
  const rowsFor = (eid) => {
    if (rowsByEntity && rowsByEntity.has(eid)) return rowsByEntity.get(eid);
    const m = membersOf(db, group.id).find(x => x.entity_id === eid)
      || { entity_id: eid, source: 'ledger' };
    return memberBalances(db, m, o, computeBalances) || [];
  };

  const adjustments = [];   // { entity_id, code, amount }  amount to SUBTRACT
  const rules = [];

  // ── 1. Investment in subsidiary ↔ contributed capital from the parent ──
  for (const p of db.prepare('SELECT * FROM consol_investment_pairs WHERE group_id = ? ORDER BY sort_order, id').all(group.id)) {
    const holder = figureFor(rowsFor(p.holder_entity_id), p.holder_account_code);
    const issuer = figureFor(rowsFor(p.issuer_entity_id), p.issuer_account_code);
    const matched = r2(Math.min(Math.abs(holder), Math.abs(issuer)));
    const typeOf = (eid, code) => {
      const row = (rowsFor(eid) || []).find(x => String(x.code) === String(code));
      return row ? row.type : null;
    };
    if (holder) adjustments.push({ entity_id: p.holder_entity_id, code: p.holder_account_code, type: typeOf(p.holder_entity_id, p.holder_account_code), amount: matched * Math.sign(holder) });
    if (issuer) adjustments.push({ entity_id: p.issuer_entity_id, code: p.issuer_account_code, type: typeOf(p.issuer_entity_id, p.issuer_account_code), amount: matched * Math.sign(issuer) });
    rules.push({
      type: 'investment_capital',
      label: p.label || 'Investment in subsidiary / contributed capital',
      legs: [
        { entity_id: p.holder_entity_id, code: p.holder_account_code, amount: r2(holder) },
        { entity_id: p.issuer_entity_id, code: p.issuer_account_code, amount: r2(issuer) },
      ],
      eliminated: matched,
      residual: r2(Math.abs(holder) - Math.abs(issuer)),
    });
  }

  // ── 2. Operating funding ↔ operating contributed capital ──
  // The funding accounts are user-maintained. The operating ledger's
  // contributed capital is the authoritative total, so anything it carries
  // beyond the listed funding accounts is reported as unassigned — a prompt to
  // tell CloudLedger which development account the newest transfer hit.
  const fundRows = db.prepare('SELECT * FROM consol_funding_accounts WHERE group_id = ? ORDER BY entity_id, account_code').all(group.id);
  const capRows = db.prepare('SELECT * FROM consol_funding_capital WHERE group_id = ?').all(group.id);
  if (fundRows.length || capRows.length) {
    // Both sides are measured BEFORE anything is eliminated, because only the
    // matched amount may come out. Eliminating a funding account with no
    // capital facing it — an operating trial balance not yet uploaded, say —
    // would leave the elimination column one-sided and the schedule would stop
    // cross-footing.
    const legFor = (eid, code, extra) => {
      const row = (rowsFor(eid) || []).find(x => String(x.code) === String(code));
      const bal = row ? r2(row.balance) : 0;
      return Object.assign({ entity_id: eid, code, type: row ? row.type : null, balance: bal }, extra);
    };
    const capLegs = capRows.map(c => {
      const l = legFor(c.entity_id, c.account_code, { side: 'capital' });
      l.available = Math.abs(l.balance);
      return l;
    });
    const fundLegs = fundRows.map(f => {
      const l = legFor(f.entity_id, f.account_code, { side: 'funding', name: f.account_name, mode: f.mode });
      // 'full' self-maintains on an account that carries nothing else;
      // 'amount' is the cumulative figure the user keeps, capped at the
      // account's own balance so an elimination can never exceed what is there.
      let want = f.mode === 'full' ? Math.abs(l.balance) : Math.abs(r2(f.amount));
      if (Math.abs(l.balance) < want) want = Math.abs(l.balance);
      l.available = r2(want);
      l.capped = f.mode !== 'full' && Math.abs(r2(f.amount)) > Math.abs(l.balance);
      return l;
    });
    let fundTotal = r2(fundLegs.reduce((s, l) => s + l.available, 0));
    const capTotal = r2(capLegs.reduce((s, l) => s + l.available, 0));

    // Forward funding derivation (Braker): when the operating ledger's
    // contributed capital exceeds what the map already eliminates, a new
    // transfer has landed that the map has not captured. Derive it from the
    // development ledger (find the recording entry, take its debit legs, follow
    // any reclass to the landing account) and add it as funding legs, so the
    // schedule self-updates each period with no hand maintenance. Capped per
    // account at that account's own balance minus what is already mapped, so a
    // landing account that is also a mapped `full` account is never
    // double-eliminated; anything the derivation cannot account for stays in the
    // residual, reported and never plugged.
    let derivedFunding = null;
    if (group.scope_key === 'braker' && fundRows.length && r2(capTotal - fundTotal) > 0.01) {
      const devEid = fundRows[0].entity_id;
      derivedFunding = deriveForwardFunding(db, group, o, devEid, r2(capTotal - fundTotal));
      const byCode = new Map();
      for (const l of derivedFunding.legs) byCode.set(l.code, r2((byCode.get(l.code) || 0) + l.amount));
      for (const [code, amount] of byCode) {
        const row = (rowsFor(devEid) || []).find(x => String(x.code) === String(code));
        const bal = row ? r2(row.balance) : 0;
        const alreadyMapped = fundLegs
          .filter(l => String(l.code) === String(code) && l.entity_id === devEid)
          .reduce((s, l) => s + Math.abs(l.available), 0);
        const room = Math.max(0, r2(Math.abs(bal) - alreadyMapped));
        const avail = r2(Math.min(Math.abs(amount), room));
        if (avail <= 0.01) continue;   // no live balance left to eliminate on this account
        fundLegs.push({ entity_id: devEid, code, type: row ? row.type : null, name: row ? row.name : null,
          balance: bal, available: avail, side: 'funding', mode: 'derived', derived: true });
      }
      fundTotal = r2(fundLegs.reduce((s, l) => s + l.available, 0));
    }

    const matchedTotal = r2(Math.min(fundTotal, capTotal));

    // Allocate the matched amount down each side in listed order. Whatever is
    // left over stays on the books and is reported, never absorbed.
    const allocate = (legs, budget) => {
      let left = budget;
      for (const l of legs) {
        const amt = r2(Math.min(l.available, Math.max(0, left)));
        l.amount = amt;
        left = r2(left - amt);
        if (amt) adjustments.push({ entity_id: l.entity_id, code: l.code, type: l.type, amount: amt * (Math.sign(l.balance) || 1) });
      }
    };
    allocate(fundLegs, matchedTotal);
    allocate(capLegs, matchedTotal);

    rules.push({
      type: 'funding_capital',
      label: 'Operating funding / contributed capital',
      legs: fundLegs.concat(capLegs),
      eliminated: matchedTotal,
      funding_total: fundTotal,
      capital_total: capTotal,
      residual: r2(capTotal - fundTotal),
      unassigned: r2(capTotal - fundTotal),
      derived: derivedFunding
        ? { matched: derivedFunding.matched, residual: derivedFunding.residual, transfers: derivedFunding.transfers }
        : null,
    });
  }

  // ── 3. Mirrored development accounts on an operating ledger (HP) ──
  // The property manager keeps a copy of the development book inside its own
  // trial balance, so these accounts are the same dollars already carried by
  // the development column. Nothing is matched: the listed account comes out
  // for whatever it holds in this window. That is deliberate — HP moves no cash
  // between development and operating, so there is no second side to agree to,
  // and pairing would only invent a constraint the books do not have.
  //
  // The mirrored set is not automatically a balanced journal entry: the
  // operating book's copy of the development book need not net to zero. What it
  // fails to net to IS the net income embedded in that copy, and it is plugged
  // onto a single Net Income (Loss) line rather than absorbed into any one
  // account (Jimmy, 2026-08-29):
  //
  //   Balance sheet / cumulative window — eliminate the mirrored asset,
  //   liability and equity accounts at their ending balances; the residual
  //   (assets - liabilities - equity across the mirror) is the plug and posts
  //   to net income. The mirrored P&L accounts (mortgage interest) are not
  //   removed line by line here — the plug already carries their whole effect.
  //
  //   Period statement of income — the balance-sheet mirror accounts do not
  //   appear, so only the mirrored P&L accounts are eliminated, each for its
  //   own activity that window (mortgage interest only in the months the
  //   operating ledger actually carries it, and for exactly that amount). No
  //   plug.
  //
  // The legacy `is_balancer` flag is no longer used; the plug is always net
  // income. The column foots and the consolidated balance sheet balances.
  const mirrorRows = db.prepare('SELECT * FROM consol_full_eliminations WHERE group_id = ? ORDER BY sort_order, entity_id, account_code').all(group.id);
  if (mirrorRows.length) {
    const legs = mirrorRows.map(m => {
      const row = (rowsFor(m.entity_id) || []).find(x => String(x.code) === String(m.account_code));
      const bal = row ? r2(row.balance) : 0;
      return {
        entity_id: m.entity_id, code: m.account_code,
        name: (row && row.name) || m.account_name || null,
        type: row ? row.type : null, balance: bal, amount: bal,
        present: !!row, is_balancer: !!m.is_balancer,
      };
    });
    // Two contexts, split by the window being built:
    //
    //   Cumulative (as_of balance sheet, or a year-to-date income range).
    //   Remove the ENTIRE mirror — asset, liability, equity AND P&L — for
    //   whatever it carries this window. The operating book's copy of the
    //   development book does not net exactly to zero (a small reconciling item
    //   between the two books), so removing it in full leaves the elimination
    //   column one-sided by that residual, which is plugged to a single Net
    //   Income (Loss) line. Measured from ending balances so the balance sheet
    //   and the year-to-date income statement carry the SAME plug and tie to
    //   each other. Because the P&L is removed here too, the consolidated
    //   statements never double-count the operating book's mortgage interest.
    //
    //   Single month income statement. Only the mirrored P&L accounts, each for
    //   its OWN activity that month — mortgage interest is removed only in the
    //   months the operating ledger actually carries it, and for exactly that
    //   amount. No plug: the reconciling item is a cumulative (stock) figure,
    //   not a one-month movement.
    let plugAsOf = null;
    if (o && o.as_of) plugAsOf = o.as_of;
    else if (o && o.from && o.to && o.from <= yearStart(o.to)) plugAsOf = o.to;   // year to date
    let plugged = 0;

    if (plugAsOf) {
      // Remove the whole mirror at this window's balances.
      for (const l of legs) {
        if (l.present && l.type) adjustments.push({ entity_id: l.entity_id, code: l.code, type: l.type, amount: l.amount });
      }
      // The plug is the residual of the mirror at ENDING balances (the same
      // figure for the balance sheet and any year-to-date window).
      const snap = new Map();
      for (const eid of new Set(mirrorRows.map(m => m.entity_id))) {
        const mem = membersOf(db, group.id).find(x => x.entity_id === eid) || { entity_id: eid, source: 'tb' };
        snap.set(eid, memberBalances(db, mem, { as_of: plugAsOf, close_pl_before: yearStart(plugAsOf) }, computeBalances) || []);
      }
      let dr = 0, cr = 0;
      for (const l of legs) {
        const row = (snap.get(l.entity_id) || []).find(x => String(x.code) === String(l.code));
        if (!row || !row.type) continue;
        if (isDrType(row.type)) dr = r2(dr + row.balance); else cr = r2(cr + row.balance);
      }
      plugged = r2(dr - cr);
      const plugEid = mirrorRows[0].entity_id;
      if (Math.abs(plugged) > 0.004) {
        // Synthetic Net Income line: Expense-typed with amount = -residual so
        // (revenue - expense) net income moves by +residual. buildColumns
        // creates the line on demand; income-statement renderers fold it into
        // net income without printing it as an expense line.
        adjustments.push({ entity_id: plugEid, code: NI_PLUG_CODE, type: 'Expense', name: NI_PLUG_NAME, amount: r2(-plugged) });
        legs.push({ entity_id: plugEid, code: NI_PLUG_CODE, name: NI_PLUG_NAME, type: 'Net Income', balance: 0, amount: r2(-plugged), present: true, plug: true });
      }
    } else {
      // Single month income statement: only the mirrored P&L accounts, at this
      // window's own activity.
      for (const l of legs) {
        if (!isPlType(l.type)) { l.amount = 0; continue; }
        if (l.present) adjustments.push({ entity_id: l.entity_id, code: l.code, type: l.type, amount: l.amount });
      }
    }

    const drSide = r2(legs.filter(l => l.type && isDrType(l.type)).reduce((s, l) => s + l.amount, 0));
    const crSide = r2(legs.filter(l => l.type && !isDrType(l.type)).reduce((s, l) => s + l.amount, 0));
    rules.push({
      type: 'full_elimination',
      label: 'Development accounts mirrored on the operating ledger',
      legs,
      eliminated: r2(legs.reduce((s, l) => s + Math.abs(l.amount), 0)),
      debit_side: drSide,
      credit_side: crSide,
      residual: 0,                 // foots: the plug closes the balance-sheet side
      plugged: r2(plugged),        // the net-income plug (0 on a period income window)
      plug_code: NI_PLUG_CODE,
      absent: legs.filter(l => !l.present).map(l => l.code),
    });
  }

  // Collapse to one figure per (entity, account). The account's type travels
  // with it so a caller can render the entry as a debit or a credit — removing
  // a positive asset is a credit, removing positive equity is a debit.
  const byKey = new Map();
  for (const a of adjustments) {
    const k = a.entity_id + '|' + a.code;
    const cur = byKey.get(k);
    if (cur) cur.amount = r2(cur.amount + a.amount);
    else byKey.set(k, { entity_id: a.entity_id, code: a.code, type: a.type || null, amount: r2(a.amount) });
  }
  return { rules, adjustments: [...byKey.values()] };
}

// ═══════════════════ Consolidated and consolidating ═══════════════════

// One consolidating schedule: a column per member, an eliminations column, and
// a consolidated column that is nothing but the cross-foot of the others.
function buildColumns(db, group, o, computeBalances) {
  const cols = columnsOf(db, group);
  const rowsByEntity = new Map();
  const unavailable = [];
  for (const c of cols) {
    const rows = memberBalances(db, c, o, computeBalances);
    if (rows === null) { unavailable.push(c.entity_id); rowsByEntity.set(c.entity_id, []); }
    else rowsByEntity.set(c.entity_id, rows);
  }
  const { rules, adjustments } = computeEliminations(db, group, o, computeBalances, rowsByEntity);

  // Union of every account any column touches, so a line prints once with a
  // figure in each column that has one.
  const accounts = new Map();
  const note = (row) => {
    if (!accounts.has(row.code)) {
      accounts.set(row.code, { code: row.code, name: row.name, type: row.type, subtype: row.subtype, bank_acct: row.bank_acct, byEntity: {}, elimination: 0, consolidated: 0 });
    }
    return accounts.get(row.code);
  };
  for (const c of cols) {
    for (const row of rowsByEntity.get(c.entity_id)) {
      const a = note(row);
      a.byEntity[c.entity_id] = r2((a.byEntity[c.entity_id] || 0) + row.balance);
    }
  }
  for (const adj of adjustments) {
    let a = accounts.get(adj.code);
    if (!a) {
      // An adjustment for a code no column carries is the synthetic net-income
      // plug: create an all-zero-member line so it lands in the elimination and
      // consolidated columns. (A regular mirror/funding account that is simply
      // absent this window emits no adjustment, so this only ever fires for the
      // plug.)
      if (Math.abs(adj.amount) < 0.004) continue;
      a = { code: adj.code, name: adj.name || adj.code, type: adj.type || null, subtype: null, bank_acct: null, byEntity: {}, elimination: 0, consolidated: 0 };
      accounts.set(adj.code, a);
    }
    a.elimination = r2(a.elimination - adj.amount);
  }
  for (const a of accounts.values()) {
    let t = 0;
    for (const c of cols) t = r2(t + (a.byEntity[c.entity_id] || 0));
    a.consolidated = r2(t + a.elimination);
  }
  return { columns: cols, accounts: [...accounts.values()], rules, unavailable };
}

// The consolidated set, in the shape buildStatements expects. Dropping this into
// the getBalances slot produces the consolidated face statements out of the
// existing engine.
function consolidatedBalances(db, group, o, computeBalances) {
  const built = buildColumns(db, group, o, computeBalances);
  return built.accounts
    .map(a => ({ code: a.code, name: a.name, type: a.type, subtype: a.subtype, bank_acct: a.bank_acct, balance: a.consolidated, total_debit: 0, total_credit: 0 }))
    .filter(r => Math.abs(r.balance) > 0.004 || true);
}

// A print-ready pair of consolidating schedules for the financial-statement
// package: the balance sheet as of the month end, and the statement of income
// for the month. One column per member plus an eliminations column and the
// consolidated cross-foot — the very same buildColumns the on-screen schedules
// use, so the printed schedules tie to the client (and to CLA) to the penny.
function buildScheduleSet(db, group, asOf, computeBalances) {
  const labelFor = (c) => {
    if (c.label) return c.label;
    const m = db.prepare("SELECT label FROM consol_members WHERE entity_id = ? AND label IS NOT NULL AND label <> '' LIMIT 1").get(c.entity_id);
    if (m && m.label) return m.label;
    const e = db.prepare('SELECT name FROM entities WHERE id = ?').get(c.entity_id);
    return e ? e.name : ('Entity ' + c.entity_id);
  };
  const columns = columnsOf(db, group).map(c => ({ entity_id: c.entity_id, label: labelFor(c), source: c.source }));
  const balanceSheet = buildColumns(db, group, { as_of: asOf, close_pl_before: yearStart(asOf) }, computeBalances);
  const incomeMonth = buildColumns(db, group, { from: monthStart(asOf), to: asOf }, computeBalances);
  const parent = db.prepare('SELECT name FROM entities WHERE id = ?').get(group.parent_entity_id);
  return { parentName: parent ? parent.name : '', asOf, columns, balanceSheet, incomeMonth };
}

module.exports = {
  ensureSchema,
  buildScheduleSet,
  scopeKeyFor,
  groupForParent,
  groupForEntity,
  membersOf,
  columnsOf,
  operatingBalances,
  computeEliminations,
  buildColumns,
  consolidatedBalances,
  tbMonths,
  tbAt,
  unmappedFor,
  mapFor,
  _helpers: { monthEnd, prevMonthEnd, yearStart, monthStart, rollUp, r2, deriveForwardFunding, followReclass },
};
