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
    return rollUp(db, entityId, lines, l => {
      const code = String(l.source_code);
      const lead = parseInt(code.replace(/[^0-9]/g, '').slice(0, 1), 10);
      const isPl = lead >= 4;
      if (!isPl) return l.ending;
      if (!base) return l.ending;
      return l.ending - (base.get(code) || 0);
    });
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
    const fundTotal = r2(fundLegs.reduce((s, l) => s + l.available, 0));
    const capTotal = r2(capLegs.reduce((s, l) => s + l.available, 0));
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
  // operating book's copy of the development book need not net to zero, and on
  // HP it does not — CLA's own July operating column, mirrored, is off by
  // 2,635.75 (a mortgage-interest reconciling item between the two books).
  //
  // One account may be flagged `is_balancer`. When it is, that account absorbs
  // the residual so the elimination column foots and the consolidated balance
  // sheet balances — exactly what CLA does, plugging the difference into the
  // mortgage-interest elimination rather than removing operating's interest in
  // full. Every other account still comes out for its whole balance. With no
  // balancer flagged the rule stays one-sided by design and reports the
  // residual (the earlier behaviour, kept for a column CLA prints one-sided).
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
    // The plug is a BALANCE-SHEET (stock) reconciling item — the amount by
    // which the operating book's copy of the development book fails to net to
    // zero at the period end. It is always measured from ENDING balances, never
    // from a window's movements, so the same figure (2,635.75 on HP at July)
    // applies to the balance sheet and to the year-to-date statement of
    // operations, keeping consolidated net income the same on both — as CLA's
    // package does. A single-month statement carries no plug: it eliminates the
    // operating mortgage interest at that month's own activity, which ties to
    // CLA's month schedule directly.
    const balancer = legs.find(l => l.is_balancer && l.present && l.type);
    let plugged = 0;
    let plugAsOf = null;
    if (o && o.as_of) plugAsOf = o.as_of;
    else if (o && o.from && o.to && o.from <= yearStart(o.to)) plugAsOf = o.to;   // year to date
    if (balancer && plugAsOf) {
      // Mirror accounts all sit on one member; snapshot its ending balances.
      const snap = new Map();
      for (const eid of new Set(mirrorRows.map(m => m.entity_id))) {
        const mem = membersOf(db, group.id).find(x => x.entity_id === eid) || { entity_id: eid, source: 'tb' };
        snap.set(eid, memberBalances(db, mem, { as_of: plugAsOf, close_pl_before: yearStart(plugAsOf) }, computeBalances) || []);
      }
      let bsDr = 0, bsCr = 0;
      for (const m of mirrorRows) {
        const row = (snap.get(m.entity_id) || []).find(x => String(x.code) === String(m.account_code));
        if (!row || !row.type) continue;
        if (isDrType(row.type)) bsDr = r2(bsDr + row.balance); else bsCr = r2(bsCr + row.balance);
      }
      const bsResidual = r2(bsDr - bsCr);
      if (Math.abs(bsResidual) > 0.004) {
        plugged = bsResidual;
        balancer.amount = r2(balancer.balance - (isDrType(balancer.type) ? bsResidual : -bsResidual));
      }
    }

    for (const l of legs) {
      // A listed account absent from the window is normal, not an error: a
      // balance-sheet mirror account has no place on a statement of operations,
      // and an account can carry nothing in a given month. It is reported so a
      // code that has quietly stopped appearing — a renamed mapping target,
      // say — can be seen rather than silently eliminating nothing forever.
      if (l.present) adjustments.push({ entity_id: l.entity_id, code: l.code, type: l.type, amount: l.amount });
    }

    // Sides after any plug — this is what actually hits the elimination column.
    const drSide = r2(legs.filter(l => l.type && isDrType(l.type)).reduce((s, l) => s + l.amount, 0));
    const crSide = r2(legs.filter(l => l.type && !isDrType(l.type)).reduce((s, l) => s + l.amount, 0));
    // Sides at full removal, before the plug — the raw mirror imbalance.
    const drFull = r2(legs.filter(l => l.type && isDrType(l.type)).reduce((s, l) => s + l.balance, 0));
    const crFull = r2(legs.filter(l => l.type && !isDrType(l.type)).reduce((s, l) => s + l.balance, 0));
    rules.push({
      type: 'full_elimination',
      label: 'Development accounts mirrored on the operating ledger',
      legs,
      eliminated: r2(legs.reduce((s, l) => s + Math.abs(l.amount), 0)),
      debit_side: drSide,
      credit_side: crSide,
      residual: r2(drSide - crSide),   // 0 in a plugged window; the imbalance otherwise
      gross_residual: r2(drFull - crFull),   // the raw mirror imbalance, before the plug
      plugged: r2(plugged),
      balancer: balancer ? balancer.code : null,
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
    const a = accounts.get(adj.code);
    if (!a) continue;
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

module.exports = {
  ensureSchema,
  scopeKeyFor,
  groupForParent,
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
  _helpers: { monthEnd, prevMonthEnd, yearStart, monthStart, rollUp, r2 },
};
