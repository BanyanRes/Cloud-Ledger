// ─────────────────────────── Intercompany module ───────────────────────────
//
// Two jobs:
//
//   1. IC MAPPING (setup, human-owned). A table that says, for one entity's GL
//      account, WHICH other entity it faces and WHAT KIND of intercompany
//      balance it is. Account names alone are not trustworthy — CloudLedger
//      copies the same chart of accounts across entities, so entity 39 owns an
//      account literally named "Due from CLR Silsbee Property Owner" while
//      BEING CLR Silsbee Property Owner. The mapping is therefore confirmed by
//      a person; `suggestMappings` only proposes.
//
//   2. IC RECONCILIATION (report). Two independent views:
//
//      • Due from / Due to — transactional. Entity A's net receivable from B
//        must equal entity B's net payable to A. Netted BY COUNTERPARTY PAIR,
//        not by account, because one relationship routinely spans several
//        accounts. Validated against the 6/30/2026 CLIP ↔ CLR Silsbee gap:
//            CLIP 18307 due from Silsbee            215,228.93
//            Silsbee 18378 due from CLIP             26,948.26
//            Silsbee 23375 due to CLIP              202,420.88
//            Silsbee net toward CLIP  26,948.26 - 202,420.88 = (175,472.62)
//            difference 215,228.93 - 175,472.62   =    39,756.31
//        which is the $39,756 reported in the 8/15 statement comparison.
//
//      • Investment / Contributed capital — structural. A parent's
//        "Investment in <child>" should equal the child's "Contributed capital
//        – <parent>".
//
// ── The elimination rule that this module exists to enforce ──
//
// A counterparty OUTSIDE the selected group is never eliminated. It gets a
// "no elim" tag and is reported separately. The CLRFI Midco I Q2 draft got this
// wrong three times — 18310 (due from County Line Rail Fund I, 2,674), 18311
// (due from County Line Rail Operations, 7,671) and 22100 (short-term loan
// payable to CLRF I, 30,073) were all eliminated against parties that are not
// in the consolidation group. Balances facing an outside party are real
// third-party balances and must survive consolidation.
//
// ── The self-referential investment rule ──
//
// Only ONE investment pattern is an error: an entity holding an "Investment in
// [itself]" (ic_type 'investment' AND counterparty_entity_id === entity_id).
// That is the gross-up found at CLIP Property Owner (19041, 1,837,842.67) and
// CLR Silsbee Property Owner (17001, 11,760,052.36) — each carries an
// investment in itself with matching contributed capital, inflating both sides
// of its own balance sheet. A normal parent→child investment is NOT an error:
// County Line Industrial Park (42) legitimately holds 19041 "Investment in CLIP
// Property Owner" of 53,980,295.39, and County Line Rail Silsbee (52)
// legitimately holds 17001 of 11,760,052.36.

const IC_TYPES = ['due_from', 'due_to', 'investment', 'contributed_capital'];
const DEFAULT_TOLERANCE = 0.005;

// ══════════════════════════════ Schema ══════════════════════════════

function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS intercompany_accounts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      account_code TEXT NOT NULL,
      account_name TEXT,
      counterparty_entity_id INTEGER,
      ic_type TEXT NOT NULL,
      is_external INTEGER NOT NULL DEFAULT 0,
      notes TEXT,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT,
      UNIQUE(entity_id, account_code)
    );
    CREATE INDEX IF NOT EXISTS idx_ic_entity ON intercompany_accounts(entity_id);
    CREATE INDEX IF NOT EXISTS idx_ic_cp ON intercompany_accounts(counterparty_entity_id);

    -- Groups were removed from the product: reconciliation now runs for one
    -- entity. These two tables are kept so an existing database is not
    -- rewritten on deploy. Nothing reads them.
    CREATE TABLE IF NOT EXISTS intercompany_groups (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      notes TEXT,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT
    );
    CREATE TABLE IF NOT EXISTS intercompany_group_members (
      group_id INTEGER NOT NULL,
      entity_id INTEGER NOT NULL,
      PRIMARY KEY (group_id, entity_id)
    );
  `);
  // A counterparty may be a company that has no ledger here — a QOZB, a joint
  // venture, a sponsor holdco. Those live in org_nodes (entity_id NULL) and are
  // referenced by counterparty_node_id. Added by migration so existing mappings
  // keep working untouched.
  const icCols = db.prepare("PRAGMA table_info(intercompany_accounts)").all().map(c => c.name);
  if (!icCols.includes('counterparty_node_id')) {
    db.exec("ALTER TABLE intercompany_accounts ADD COLUMN counterparty_node_id INTEGER");
    console.log('[db migrate] intercompany_accounts.counterparty_node_id added');
  }
}

// org_nodes is created by server/orgstructure.js. Both modules register at boot,
// so it exists by the time any request runs — but a read that happens before
// that (or on a database restored without it) must degrade, not throw.
function hasOrgNodes(db) {
  try { db.prepare('SELECT 1 FROM org_nodes LIMIT 1').get(); return true; }
  catch { return false; }
}

// Registered companies: every org_node, whether or not it is backed by a
// CloudLedger entity. This is the list the counterparty picker offers beyond
// the entity list.
function listCompanies(db) {
  if (!hasOrgNodes(db)) return [];
  return db.prepare(`
    SELECT n.id, n.name, n.entity_id, n.node_type, n.notes,
           e.name AS entity_name, e.code AS entity_code
    FROM org_nodes n LEFT JOIN entities e ON e.id = n.entity_id
    WHERE n.node_type <> 'individual'
    ORDER BY n.name COLLATE NOCASE`).all();
}

// Fold a node-based counterparty into the shape the reconciliations expect.
// A node that IS backed by an entity resolves to that entity, so registering a
// company for an entity that already exists changes nothing. A node with no
// entity stays off-ledger: there is no second ledger to match it against, so it
// can never be eliminated, and it is reported as its own case rather than being
// lumped in with "external".
function resolveCounterparties(db, mappings) {
  if (!hasOrgNodes(db)) return mappings.map(m => ({ ...m, off_ledger: false }));
  const byId = new Map(db.prepare('SELECT id, name, entity_id FROM org_nodes').all().map(n => [n.id, n]));
  return mappings.map(m => {
    if (m.counterparty_entity_id != null || !m.counterparty_node_id) return { ...m, off_ledger: false };
    const n = byId.get(Number(m.counterparty_node_id));
    if (!n) return { ...m, off_ledger: false };
    if (n.entity_id != null) {
      return { ...m, counterparty_entity_id: n.entity_id, counterparty_name: n.name, off_ledger: false };
    }
    return { ...m, counterparty_name: n.name, off_ledger: true };
  });
}

// ══════════════════════════ Name → entity matching ══════════════════════════

// Strip the noise that stops "County Line Rail Fund I LP" from matching the
// entity "County Line Rail Fund".
function normName(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[.,'"()]/g, ' ')
    .replace(/\b(llc|l\.l\.c|lp|l\.p|llp|inc|incorporated|corp|corporation|co|company|ltd|limited|the)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

// Hand-maintained aliases for the County Line family, where GL account names
// use short forms ("Due from Buna", "Due to SRN") that no generic matcher can
// resolve. Ordered MOST SPECIFIC FIRST — "county line rail silsbee" has to be
// tested before the bare "silsbee", and "county line industrial park" before
// "clip", or the wrong entity wins.
const ALIASES = [
  [/^clrfi\b.*midco|^midco/,                              'CLRFIMID'],
  [/^county line rail silsbee|^clr silsbee$/,             'COUNTYLI5'],
  [/^county line industrial park/,                        'COUNTYLI2'],
  [/^county line rail operations|^clro\b/,                'COUNTYLI3'],
  [/^county line railroad interest|^clri\b/,              'COUNTYLI'],
  [/^county line rail fund|^clrf\b|^clr fund\b/,          'COUNTYLI1'],
  [/^clr silsbee property owner|^silsbee property owner/, 'CLRSILSB2'],
  [/^clr silsbee sponsor/,                                'CLRSILSB1'],
  [/^clr silsbee manager/,                                'CLRSILSB'],
  [/^clip property owner/,                                'CLIPPROP'],
  [/^clip sponsor/,                                       'CLIPSPON'],
  [/^clr buna property owner|^buna/,                      'CLRBUNAP'],
  [/^sabine river|^srn\b|^county line srn/,               'SABINERI'],
  [/^silsbee/,                                            'CLRSILSB2'],
  [/^clip/,                                               'CLIPPROP'],
  [/^banyan residential$/,                                'BANYANRE1'],
  [/^turn ?key rail/,                                     'TURNKEYR'],
];

// Roman numerals and bare numbers act as series markers in these company names.
// They are the only thing telling "QOZB I" from "QOZB III", so they are compared
// exactly rather than as ordinary tokens.
const ROMAN = /^(i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii)$/;
function seriesTokens(normalized) {
  return new Set(String(normalized || '').split(' ')
    .filter(t => t && (ROMAN.test(t) || /^\d+$/.test(t))));
}
function sameSeries(a, b) {
  if (a.size !== b.size) return false;
  for (const t of a) if (!b.has(t)) return false;
  return true;
}

// Counterparty labels that are explicitly NOT another entity in the ledger.
const EXTERNAL_HINTS = [/outside vendor/i, /third party/i, /^vendor$/i, /affiliates?$/i,
                        /management company/i, /portfolio compan/i, /port co\b/i];

// Resolve a counterparty label from an account name to an entity id.
// Returns { entity_id, confidence, reason } or { entity_id: null, ... }.
// Deliberately conservative: an ambiguous label resolves to null so the setup
// page shows it blank and a person decides, rather than guessing wrong and
// having a wrong guess quietly drive an elimination.
function matchEntity(label, entities) {
  const raw = String(label || '').trim();
  if (!raw) return { entity_id: null, confidence: 'none', reason: 'no counterparty in account name' };
  if (EXTERNAL_HINTS.some(re => re.test(raw))) {
    return { entity_id: null, confidence: 'external', reason: 'looks like a party outside the ledger' };
  }
  const n = normName(raw);
  if (!n) return { entity_id: null, confidence: 'none', reason: 'counterparty name empty after normalization' };

  for (const [re, code] of ALIASES) {
    if (re.test(n)) {
      const ent = entities.find(e => e.code === code);
      if (ent) return { entity_id: ent.id, confidence: 'alias', reason: 'matched alias for ' + ent.name };
    }
  }

  const exact = entities.filter(e => normName(e.name) === n);
  if (exact.length === 1) return { entity_id: exact[0].id, confidence: 'exact', reason: 'exact name match' };
  if (exact.length > 1) return { entity_id: null, confidence: 'ambiguous', reason: exact.length + ' entities share this name' };

  // Token overlap, accepted only when a single entity clearly wins.
  const want = new Set(n.split(' ').filter(Boolean));
  if (!want.size) return { entity_id: null, confidence: 'none', reason: 'no usable tokens' };
  const wantSeries = seriesTokens(n);
  const scored = entities.map(e => {
    const en = normName(e.name);
    // Series guard. Names in these structures are distinguished ONLY by a roman
    // numeral or a number — "Bridge Banyan QOZB I / II / III", "CLRFI Midco I"
    // vs "Midco II", "Milhaus QOZ Fund V / VI / VII". Plain token overlap rates
    // those as near-identical: "Bridge Banyan QOZB III" scores 0.75 against the
    // entity "Bridge Banyan HP QOZB" and would be accepted as a confident
    // match. If the two sides carry different series markers, they are
    // different companies, whatever the rest of the words say.
    if (!sameSeries(wantSeries, seriesTokens(en))) return { e, score: 0 };
    const have = new Set(en.split(' ').filter(Boolean));
    let hit = 0; for (const w of want) if (have.has(w)) hit++;
    return { e, score: hit / Math.max(want.size, have.size || 1) };
  }).filter(s => s.score >= 0.6).sort((a, b) => b.score - a.score);
  if (scored.length === 1 || (scored.length > 1 && scored[0].score - scored[1].score >= 0.15)) {
    return { entity_id: scored[0].e.id, confidence: 'fuzzy', reason: 'token match to ' + scored[0].e.name };
  }
  if (scored.length > 1) return { entity_id: null, confidence: 'ambiguous', reason: 'several entities match "' + raw + '"' };
  return { entity_id: null, confidence: 'none', reason: 'no entity matches "' + raw + '"' };
}

// Resolve a counterparty label to an entity FIRST, then to a registered company.
// Entities win: a company registered for an entity that already exists should
// not shadow the entity itself. Returns the same shape as matchEntity plus
// node_id, and — when nothing matches — the label to offer as a new company.
function matchCompany(label, entities, companies) {
  const viaEntity = matchEntity(label, entities);
  if (viaEntity.entity_id != null || viaEntity.confidence === 'external') {
    return { ...viaEntity, node_id: null };
  }
  const offLedger = (companies || []).filter(c => c.entity_id == null);
  if (offLedger.length) {
    // Reuse the entity matcher by presenting companies in the same shape.
    const asEntities = offLedger.map(c => ({ id: c.id, name: c.name, code: null }));
    const viaNode = matchEntity(label, asEntities);
    if (viaNode.entity_id != null) {
      const hit = offLedger.find(c => c.id === viaNode.entity_id);
      return {
        entity_id: null, node_id: viaNode.entity_id, confidence: 'company',
        reason: 'matched the registered company ' + (hit ? hit.name : ''),
      };
    }
  }
  return { ...viaEntity, node_id: null };
}

// ── Individual investors ──
//
// Outside individuals hold capital in most of these entities ("Contributed
// Capital - John H. Grayson Jr."). They are not intercompany counterparties:
// nothing ever eliminates against a person, and proposing a mapping for each
// one buries the handful of real company counterparties in noise. They are
// therefore kept out of the suggestion list and out of the "not mapped yet"
// panel entirely.
//
// The test is deliberately narrow. It is applied ONLY to contributed-capital
// accounts, because that is the only place a person appears — an investment
// account never names one. That restriction is what makes it safe: two-word
// company names like "North Tempe" and "Mountain Banyan" read exactly like
// personal names, but they sit on investment accounts and are never examined.

// Any of these words means the label is an organisation, whatever else it says.
const COMPANY_HINTS = /\b(llc|l\.l\.c|lp|llp|lllp|inc|incorporated|corp|corporation|co|company|ltd|limited|plc|pllc|pc|trust|fund|funds|capital|holding|holdings|partner|partners|partnership|group|venture|ventures|management|manager|properties|property|investment|investments|investor|investors|associates|sponsor|midco|propco|opco|qozb|qof|jv|joint|bank|reit|gst|endowment|foundation|university|realty|development|enterprises|solutions|services|industries|resources|energy|equity|advisors|advisers|estates|land|rail|railroad|sfr|qoz)\b/i;

// Any of these means the label is a thing rather than a party — "Waived
// Development Fee", "Prior Period Adjustment". Those are not people either, but
// they must not be reported as one.
const NON_PARTY_HINTS = /\b(fee|fees|waived|expense|expenses|reserve|contribution|contributions|adjustment|distribution|distributions|note|notes|loan|interest|misc|miscellaneous|other|various|prior|opening|beginning|retained|earnings|income|syndication|organization|organisation|org|payable|receivable|clearing|suspense)\b/i;

const NAME_SUFFIX = /^(jr|sr|ii|iii|iv|md|phd|esq)$/i;

// True when the label reads as a personal name: two to four alphabetic words,
// allowing a middle initial and a generational suffix, with no organisation
// word, no digits and no ampersand.
function looksIndividual(label) {
  const raw = String(label || '').trim();
  if (!raw) return false;
  if (COMPANY_HINTS.test(raw)) return false;
  if (NON_PARTY_HINTS.test(raw)) return false;
  if (/[&@\/\d]/.test(raw)) return false;
  const toks = raw.replace(/[.,]/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
  if (toks.length < 2 || toks.length > 4) return false;
  let core = 0;
  for (const t of toks) {
    if (NAME_SUFFIX.test(t)) continue;
    if (!/^[A-Za-z'-]+$/.test(t)) return false;
    if (t.length === 1) continue;
    core++;
  }
  return core >= 2;
}

// A parsed account is an individual investor's capital only on the capital side.
function isIndividualInvestor(parsed) {
  return !!parsed && parsed.ic_type === 'contributed_capital' && looksIndividual(parsed.label);
}

// -- People marked by hand --
//
// The automatic test above CANNOT be extended to due-from / due-to accounts.
// Checked against all 141 due-from/to labels in the ledger: the name-shape rule
// flags 30 of them, and 24 are real companies - Banyan Cyrene, Banyan EIV,
// County Line Industrial Park, Milhaus Beckley, Scottsdale Entrada I/II and
// more. Only six are actually people. Applying it there would hide two dozen
// genuine counterparties to catch six, which is the wrong trade every time.
//
// So on those accounts a person decides, once, and the answer is remembered for
// every entity: "Due from Ben Brosseau" is marked a person and never proposed
// again. People are stored in org_nodes with node_type 'individual', so there is
// still one party registry rather than a second table - they are simply excluded
// from the company picker.
function listPeople(db) {
  if (!hasOrgNodes(db)) return [];
  return db.prepare("SELECT id, name, notes FROM org_nodes WHERE node_type = 'individual' ORDER BY name COLLATE NOCASE").all();
}

function isMarkedPerson(people, label) {
  const n = normName(label);
  if (!n) return false;
  return people.some(p => normName(p.name) === n);
}

// Parse a GL account name into { ic_type, counterparty_label }, or null when the
// name is not an intercompany account at all.
function parseAccountName(name) {
  const s = String(name || '').trim();
  let m;
  if ((m = s.match(/^due\s+from\s+(.+)$/i)))                      return { ic_type: 'due_from', label: m[1] };
  if ((m = s.match(/^due\s+to\s+(.+)$/i)))                        return { ic_type: 'due_to', label: m[1] };
  if ((m = s.match(/^contributed\s+capital\s*[-–—:]?\s*(.+)$/i))) return { ic_type: 'contributed_capital', label: m[1] };
  if ((m = s.match(/^investments?\s*(?:in|-|–|—|:)\s*(.+)$/i)))   return { ic_type: 'investment', label: m[1] };
  // CLRF's fund-level naming: "CLIP - Investment Purchase", "Silsbee - Investment Purchase"
  if ((m = s.match(/^(.+?)\s*[-–—]\s*investment\s+purchase$/i)))  return { ic_type: 'investment', label: m[1] };
  return null;
}

// ══════════════════════════════ Groups ══════════════════════════════

// ══════════════════════════════ Mappings ══════════════════════════════

function listMappings(db, { entity_id } = {}) {
  let where = '', params = [];
  if (entity_id) { where = 'WHERE m.entity_id = ?'; params = [entity_id]; }
  // COALESCE so a node-based counterparty shows its company name in the UI
  // exactly like an entity-based one.
  const nodeJoin = hasOrgNodes(db)
    ? 'LEFT JOIN org_nodes n ON n.id = m.counterparty_node_id'
    : '';
  const nodeName = hasOrgNodes(db) ? 'n.name' : 'NULL';
  return db.prepare(`
    SELECT m.*, e.name AS entity_name, e.code AS entity_code,
           COALESCE(c.name, ${nodeName}) AS counterparty_name, c.code AS counterparty_code,
           ${nodeName} AS counterparty_node_name
    FROM intercompany_accounts m
    LEFT JOIN entities e ON e.id = m.entity_id
    LEFT JOIN entities c ON c.id = m.counterparty_entity_id
    ${nodeJoin}
    ${where}
    ORDER BY e.name COLLATE NOCASE, m.account_code`).all(...params);
}

function validateMapping(body) {
  if (!body.entity_id) return 'entity_id is required';
  if (!body.account_code) return 'account_code is required';
  if (!IC_TYPES.includes(body.ic_type)) return 'ic_type must be one of: ' + IC_TYPES.join(', ');
  if (!body.is_external && !body.counterparty_entity_id && !body.counterparty_node_id) {
    return 'Pick a counterparty — a CloudLedger entity, a registered company, or mark it external';
  }
  return null;
}

function createMapping(db, body, who) {
  const now = new Date().toISOString();
  const r = db.prepare(`INSERT INTO intercompany_accounts
    (entity_id, account_code, account_name, counterparty_entity_id, counterparty_node_id, ic_type, is_external, notes, created_at, created_by, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)`).run(
      Number(body.entity_id), String(body.account_code), body.account_name || null,
      body.is_external ? null : (body.counterparty_entity_id ? Number(body.counterparty_entity_id) : null),
      body.is_external ? null : (body.counterparty_node_id ? Number(body.counterparty_node_id) : null),
      body.ic_type, body.is_external ? 1 : 0, body.notes || null, now, who || null, now, who || null);
  return r.lastInsertRowid;
}

function updateMapping(db, id, body, who) {
  const now = new Date().toISOString();
  db.prepare(`UPDATE intercompany_accounts SET
      account_code=?, account_name=?, counterparty_entity_id=?, counterparty_node_id=?, ic_type=?, is_external=?, notes=?, updated_at=?, updated_by=?
    WHERE id=?`).run(
      String(body.account_code), body.account_name || null,
      body.is_external ? null : (body.counterparty_entity_id ? Number(body.counterparty_entity_id) : null),
      body.is_external ? null : (body.counterparty_node_id ? Number(body.counterparty_node_id) : null),
      body.ic_type, body.is_external ? 1 : 0, body.notes || null, now, who || null, id);
}

function deleteMapping(db, id) {
  db.prepare('DELETE FROM intercompany_accounts WHERE id = ?').run(id);
}

// Propose mappings for one entity by reading its chart of accounts. Every
// proposal carries the account's current balance and the reason the
// counterparty was chosen, so a reviewer can see what actually matters (a
// nonzero balance) and where the matcher was unsure. Nothing is saved here —
// the client posts back only the rows a person accepted.
function suggestMappings(db, entityId, { computeBalances, as_of, includeIndividuals }) {
  const hiddenIndividuals = [];
  ensureSchema(db);
  const entities = db.prepare('SELECT id, name, code FROM entities').all();
  const companies = listCompanies(db);
  const people = listPeople(db);
  const existing = new Set(db.prepare('SELECT account_code FROM intercompany_accounts WHERE entity_id = ?')
    .all(entityId).map(r => r.account_code));
  const balances = computeBalances(entityId, as_of ? { as_of } : {});
  const balByCode = new Map(balances.map(b => [String(b.code), b]));
  const accounts = db.prepare('SELECT code, name, type FROM accounts WHERE entity_id = ? ORDER BY code').all(entityId);

  const out = [];
  for (const a of accounts) {
    if (existing.has(String(a.code))) continue;
    const parsed = parseAccountName(a.name);
    if (!parsed) continue;
    if (isIndividualInvestor(parsed) || isMarkedPerson(people, parsed.label)) {
      hiddenIndividuals.push({ account_code: String(a.code), account_name: a.name, label: parsed.label });
      if (!includeIndividuals) continue;
    }
    const match = matchCompany(parsed.label, entities, companies);
    const bal = balByCode.get(String(a.code));
    const self = match.entity_id != null && Number(match.entity_id) === Number(entityId);
    out.push({
      entity_id: Number(entityId),
      account_code: String(a.code),
      account_name: a.name,
      account_type: a.type,
      ic_type: parsed.ic_type,
      counterparty_label: parsed.label,
      counterparty_entity_id: match.entity_id,
      counterparty_node_id: match.node_id || null,
      is_external: match.confidence === 'external' ? 1 : 0,
      confidence: match.confidence,
      reason: match.reason,
      self_referential: self,
      // Nothing in CloudLedger answers to this name. The UI offers a one-click
      // "register this company" so the 36 unresolved counterparties in the
      // portfolio can be entered as they are met, rather than up front.
      can_register: match.entity_id == null && !match.node_id && match.confidence !== 'external',
      // Offered on any unresolved row so a person on a due-from / due-to account
      // can be dismissed for good, which the automatic test cannot do safely.
      can_mark_person: match.entity_id == null && !match.node_id && match.confidence !== 'external',
      balance: bal ? bal.balance : 0,
    });
  }
  // Accounts that carry money first — a zero-balance shell account is noise.
  out.sort((a, b) => Math.abs(b.balance) - Math.abs(a.balance) || a.account_code.localeCompare(b.account_code));
  if (includeIndividuals) for (const o of out) {
    o.individual = isIndividualInvestor({ ic_type: o.ic_type, label: o.counterparty_label })
      || isMarkedPerson(people, o.counterparty_label);
  }
  return { suggestions: out, hidden_individuals: hiddenIndividuals.length, hidden_examples: hiddenIndividuals.slice(0, 5) };
}

// ══════════════════════════ Reconciliation helpers ══════════════════════════


// ════════════════════ Due from / Due to reconciliation ════════════════════
//
// Netted by counterparty PAIR. A's net position toward B is
// (sum of A's due-from B) − (sum of A's due-to B); B's net toward A is the
// same computed from B's side. If the relationship is stated consistently the
// two are equal and opposite, so their SUM is the unreconciled difference.

// ═══════════ Investment / Contributed capital reconciliation ═══════════

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function fmtMoney(n) {
  return '$' + Math.abs(Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ═══════════ Reconciliation for ONE entity ═══════════
//
// The reconciliation is run for a chosen entity, and every row carries the
// account CODE and NAME on BOTH sides — the selected entity's account and the
// counterparty's — because that is what makes a difference actionable. A netted
// pair total tells you the two ledgers disagree; the account numbers tell you
// which two entries to go and look at.
//
// One row per counterparty, holding every account on each side. Legs are kept
// as lists rather than summed away, so the client can print them side by side
// and the export can carry them.

function amountsFor(balances, entityId, code) {
  const row = (balances.get(entityId) || new Map()).get(String(code));
  return row ? Number(row.balance) || 0 : 0;
}

// Turn a mapping into a printable leg. `signed` is the amount in receivable
// terms for due-from/due-to, so a pair's two sides should sum to zero.
function legOf(m, balances) {
  const amount = amountsFor(balances, m.entity_id, m.account_code);
  return {
    mapping_id: m.id,
    account_code: m.account_code,
    account_name: m.account_name,
    ic_type: m.ic_type,
    amount,
    signed: m.ic_type === 'due_to' ? -amount : amount,
    notes: m.notes || null,
  };
}

function reconcileForEntity(db, entityId, opts) {
  ensureSchema(db);
  const eid = Number(entityId);
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;
  const { computeBalances, as_of } = opts;

  const entity = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(eid);
  if (!entity) { const e = new Error('Entity not found'); e.status = 404; throw e; }

  const all = resolveCounterparties(db, listMappings(db, {}));
  const forEntity = all.filter(m => Number(m.entity_id) === eid);

  // Accounts facing a party OUTSIDE the group are not reconciled. There is no
  // counterparty ledger to agree with, so a difference on such a row means
  // nothing — it only buries the rows that do need attention. They are counted
  // below so the excluded balance is still visible, never silently dropped.
  // Off-ledger counterparties are a different thing and stay in: those are
  // inside the group, they just have no ledger in CloudLedger yet.
  const external = forEntity.filter(m => m.is_external);
  const mine = forEntity.filter(m => !m.is_external);

  // Balances are needed for this entity and for every counterparty it names,
  // because the counterparty's own accounts are printed alongside.
  const cpIds = [...new Set(mine
    .filter(m => m.counterparty_entity_id != null && Number(m.counterparty_entity_id) !== eid)
    .map(m => Number(m.counterparty_entity_id)))];
  const balances = new Map();
  for (const id of [eid, ...cpIds]) {
    balances.set(id, new Map(computeBalances(id, as_of ? { as_of } : {}).map(r => [String(r.code), r])));
  }
  const entities = db.prepare('SELECT id, name, code FROM entities').all();
  const entById = new Map(entities.map(e => [e.id, e]));
  const nm = id => (entById.get(Number(id)) || {}).name || ('Entity ' + id);

  // Group this entity's mappings by counterparty. An off-ledger counterparty
  // gets its own row too: it can never match, but it is inside the group and
  // its balance belongs on the entity's intercompany position.
  const build = (types) => {
    const groups = new Map();
    const key = m => m.counterparty_entity_id != null ? 'e' + m.counterparty_entity_id
      : (m.counterparty_node_id != null ? 'n' + m.counterparty_node_id : 'u' + m.account_code);

    for (const m of mine) {
      if (!types.includes(m.ic_type)) continue;
      const k = key(m);
      if (!groups.has(k)) {
        const cpEid = m.counterparty_entity_id != null ? Number(m.counterparty_entity_id) : null;
        groups.set(k, {
          counterparty_entity_id: cpEid,
          counterparty_node_id: m.counterparty_node_id || null,
          counterparty_name: cpEid != null ? nm(cpEid) : (m.counterparty_name || 'Unmapped'),
          off_ledger: !!m.off_ledger,
          self: cpEid === eid,
          our_legs: [], their_legs: [],
        });
      }
      groups.get(k).our_legs.push(legOf(m, balances));
    }

    // The counterparty's own side: its mappings that point back at this entity.
    for (const g of groups.values()) {
      if (g.counterparty_entity_id == null || g.self) continue;
      const theirs = all.filter(m =>
        Number(m.entity_id) === g.counterparty_entity_id &&
        types.includes(m.ic_type) &&
        m.counterparty_entity_id != null &&
        Number(m.counterparty_entity_id) === eid);
      g.their_legs = theirs.map(m => legOf(m, balances));
    }
    return [...groups.values()];
  };

  // ── Due from / Due to ──
  const due = build(['due_from', 'due_to']).map(g => {
    const our_net = g.our_legs.reduce((s, l) => s + l.signed, 0);
    const their_net = g.their_legs.reduce((s, l) => s + l.signed, 0);
    const difference = our_net + their_net;
    const bothEmpty = Math.abs(our_net) < tol && Math.abs(their_net) < tol;
    let status;
    if (g.off_ledger || g.counterparty_entity_id == null) status = 'no_ledger';
    else if (bothEmpty) status = 'settled';
    else if (!g.their_legs.length) status = 'one_sided';
    else if (Math.abs(difference) < tol) status = 'matched';
    else status = 'mismatch';
    return { ...g, our_net: round2(our_net), their_net: round2(their_net), difference: round2(difference), status };
  }).sort((a, b) => Math.abs(b.difference) - Math.abs(a.difference) || Math.abs(b.our_net) - Math.abs(a.our_net));

  // ── Investment / Contributed capital ──
  const invGroups = build(['investment', 'contributed_capital']);
  const sum = (legs, t) => legs.filter(l => l.ic_type === t).reduce((s, l) => s + l.amount, 0);
  const investment = invGroups.map(g => {
    const our_investment = sum(g.our_legs, 'investment');
    const our_capital = sum(g.our_legs, 'contributed_capital');
    const their_investment = sum(g.their_legs, 'investment');
    const their_capital = sum(g.their_legs, 'contributed_capital');
    // Our investment in them answers to their capital from us, and vice versa.
    const inv_difference = our_investment - their_capital;
    const cap_difference = our_capital - their_investment;
    const hasInv = Math.abs(our_investment) >= tol || Math.abs(their_capital) >= tol;
    const hasCap = Math.abs(our_capital) >= tol || Math.abs(their_investment) >= tol;
    let status;
    if (g.self) status = 'self';
    else if (g.off_ledger || g.counterparty_entity_id == null) status = 'no_ledger';
    else if (!hasInv && !hasCap) status = 'settled';
    else if (!g.their_legs.length) status = 'one_sided';
    else if ((!hasInv || Math.abs(inv_difference) < tol) && (!hasCap || Math.abs(cap_difference) < tol)) status = 'matched';
    else status = 'mismatch';
    const difference = (hasInv ? inv_difference : 0) + (hasCap ? cap_difference : 0);
    return {
      ...g,
      our_investment: round2(our_investment), our_capital: round2(our_capital),
      their_investment: round2(their_investment), their_capital: round2(their_capital),
      inv_difference: round2(inv_difference), cap_difference: round2(cap_difference),
      difference: round2(difference), status,
    };
  }).sort((a, b) => Math.abs(b.difference) - Math.abs(a.difference));

  // An entity holding an investment in ITSELF — the CLIP / CLR Silsbee gross-up.
  const findings = mine
    .filter(m => m.ic_type === 'investment' && m.counterparty_entity_id != null && Number(m.counterparty_entity_id) === eid)
    .map(m => {
      const l = legOf(m, balances);
      return { ...l, severity: 'error', finding: 'self_investment',
        message: entity.name + ' holds an investment in itself — this grosses up its own assets and equity by ' + fmtMoney(l.amount) + '.' };
    })
    .filter(f => Math.abs(f.amount) >= tol);

  // Accounts that look intercompany, carry a balance, and are not mapped. Kept
  // per-tab so the investment tab can say "nothing here" honestly.
  // Every mapping the entity has, external included: an account mapped to an
  // outside party is mapped, and must not resurface as "not mapped yet".
  const mappedCodes = new Set(forEntity.map(m => String(m.account_code)));
  const people = listPeople(db);
  const unmappedFor = (types) => db.prepare('SELECT code, name FROM accounts WHERE entity_id = ?').all(eid)
    .filter(a => {
      if (mappedCodes.has(String(a.code))) return false;
      const p = parseAccountName(a.name);
      if (!p || !types.includes(p.ic_type)) return false;
      if (isIndividualInvestor(p) || isMarkedPerson(people, p.label)) return false;
      return Math.abs(amountsFor(balances, eid, a.code)) >= tol;
    })
    .map(a => ({ account_code: String(a.code), account_name: a.name,
      balance: round2(amountsFor(balances, eid, a.code)) }))
    .sort((x, y) => Math.abs(y.balance) - Math.abs(x.balance));

  const tally = rows => ({
    count: rows.length,
    matched: rows.filter(r => r.status === 'matched').length,
    mismatched: rows.filter(r => r.status === 'mismatch').length,
    one_sided: rows.filter(r => r.status === 'one_sided').length,
    no_ledger: rows.filter(r => r.status === 'no_ledger').length,
    settled: rows.filter(r => r.status === 'settled').length,
    abs_difference: round2(rows.filter(r => r.status === 'mismatch')
      .reduce((s, r) => s + Math.abs(r.difference), 0)),
  });

  // Not reconciled, but shown as a one-line note so the balance is accounted
  // for. Only accounts that actually carry one are worth mentioning.
  const excludedFor = (types) => external
    .filter(m => types.includes(m.ic_type))
    .map(m => legOf(m, balances))
    .filter(l => Math.abs(l.amount) >= tol)
    .map(l => ({ account_code: l.account_code, account_name: l.account_name,
      ic_type: l.ic_type, amount: l.amount }))
    .sort((x, y) => Math.abs(y.amount) - Math.abs(x.amount));

  const dueUnmapped = unmappedFor(['due_from', 'due_to']);
  const invUnmapped = unmappedFor(['investment', 'contributed_capital']);
  const dueExternal = excludedFor(['due_from', 'due_to']);
  const invExternal = excludedFor(['investment', 'contributed_capital']);


  return {
    entity: { id: entity.id, name: entity.name, code: entity.code },
    as_of: as_of || null,
    tolerance: tol,
    due: { rows: due, unmapped: dueUnmapped, external: dueExternal, totals: tally(due) },
    investment: {
      rows: investment, unmapped: invUnmapped, findings, totals: tally(investment),
      external: invExternal,
      // Drives an explicit "this entity has no investment accounts" message
      // rather than an empty table that looks like something failed to load.
      has_any: investment.length > 0 || findings.length > 0 || invUnmapped.length > 0
        || invExternal.length > 0,
    },
  };
}

// ══════════════════════════════ Routes ══════════════════════════════

function registerIntercompanyRoutes(app, deps) {
  const { db, auth, requireRole, computeBalances, userHasEntityAccess } = deps;
  ensureSchema(db);

  const who = req => (req.user && (req.user.name || req.user.email)) || null;
  const gate = [auth, requireRole('Admin', 'Accountant')];

  // Every intercompany read spans several entities at once, so entity access is
  // checked per entity rather than by the usual single-:eid middleware. A user
  // who cannot see one member of the group cannot run the group's report.
  function assertEntityAccess(req, entityIds) {
    for (const eid of entityIds) {
      if (!userHasEntityAccess(req.user.id, req.user.role, eid)) {
        const e = new Error('No access to entity ' + eid);
        e.status = 403; throw e;
      }
    }
  }
  const fail = (res, e) => res.status(e.status || 500).json({ error: e.message });

  // ── Registered companies ──
  // Companies that appear on the org charts as counterparties. They live in
  // org_nodes so the same row can later be placed in an ownership tree without
  // being re-entered; a company created here simply has no edges yet. Exposed
  // from this module (not just from /api/org-structure) because the need shows
  // up while mapping accounts, and making someone leave the page to register a
  // QOZB before they can finish a mapping is the wrong shape.
  app.get('/api/intercompany/companies', ...gate, (req, res) => {
    try { res.json(listCompanies(db)); } catch (e) { fail(res, e); }
  });

  app.post('/api/intercompany/companies', ...gate, (req, res) => {
    try {
      const name = req.body && String(req.body.name || '').trim();
      if (!name) return res.status(400).json({ error: 'Company name is required' });
      const dup = db.prepare('SELECT id, name FROM org_nodes WHERE LOWER(name) = LOWER(?)').get(name);
      if (dup) return res.json({ id: dup.id, existing: true, name: dup.name });
      const eid = req.body.entity_id ? Number(req.body.entity_id) : null;
      if (eid && !userHasEntityAccess(req.user.id, req.user.role, eid)) {
        return res.status(403).json({ error: 'No access to entity ' + eid });
      }
      const now = new Date().toISOString();
      const r = db.prepare(`INSERT INTO org_nodes (name, entity_id, node_type, notes, sort_order, created_at, created_by, updated_at, updated_by)
        VALUES (?,?,?,?,0,?,?,?,?)`).run(
          name, eid, eid ? 'company' : 'shell', (req.body.notes || null),
          now, who(req), now, who(req));
      res.json({ id: r.lastInsertRowid, existing: false, name });
    } catch (e) { fail(res, e); }
  });

  // -- People --
  // Names marked as individuals so they stop being proposed as counterparties.
  app.get('/api/intercompany/people', ...gate, (req, res) => {
    try { res.json(listPeople(db)); } catch (e) { fail(res, e); }
  });

  app.post('/api/intercompany/people', ...gate, (req, res) => {
    try {
      const name = req.body && String(req.body.name || '').trim();
      if (!name) return res.status(400).json({ error: 'Name is required' });
      const dup = db.prepare('SELECT id, name, node_type FROM org_nodes WHERE LOWER(name) = LOWER(?)').get(name);
      if (dup) {
        // Re-marking something already registered as a company converts it,
        // rather than leaving two rows disagreeing about what it is.
        if (dup.node_type !== 'individual') {
          db.prepare("UPDATE org_nodes SET node_type = 'individual', entity_id = NULL, updated_at = ?, updated_by = ? WHERE id = ?")
            .run(new Date().toISOString(), who(req), dup.id);
        }
        return res.json({ id: dup.id, existing: true, name: dup.name });
      }
      const now = new Date().toISOString();
      const r = db.prepare("INSERT INTO org_nodes (name, entity_id, node_type, notes, sort_order, created_at, created_by, updated_at, updated_by) VALUES (?, NULL, 'individual', ?, 0, ?, ?, ?, ?)")
        .run(name, req.body.notes || null, now, who(req), now, who(req));
      res.json({ id: r.lastInsertRowid, existing: false, name });
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/intercompany/people/:id', ...gate, (req, res) => {
    try {
      db.prepare("DELETE FROM org_nodes WHERE id = ? AND node_type = 'individual'").run(Number(req.params.id));
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

  // ── Mappings ──
  app.get('/api/intercompany/mappings', ...gate, (req, res) => {
    try {
      const entity_id = req.query.entity_id ? Number(req.query.entity_id) : null;
      if (entity_id) assertEntityAccess(req, [entity_id]);
      const rows = listMappings(db, { entity_id });
      res.json(entity_id
        ? rows
        : rows.filter(r => userHasEntityAccess(req.user.id, req.user.role, r.entity_id)));
    } catch (e) { fail(res, e); }
  });

  app.get('/api/intercompany/mappings/suggest', ...gate, (req, res) => {
    try {
      const entity_id = Number(req.query.entity_id);
      if (!entity_id) return res.status(400).json({ error: 'entity_id is required' });
      assertEntityAccess(req, [entity_id]);
      res.json(suggestMappings(db, entity_id, { computeBalances, as_of: req.query.as_of,
        includeIndividuals: req.query.include_individuals === '1' }));
    } catch (e) { fail(res, e); }
  });

  app.post('/api/intercompany/mappings', ...gate, (req, res) => {
    try {
      // Accepts one mapping or an array (the "accept these suggestions" path).
      const items = Array.isArray(req.body) ? req.body : [req.body];
      for (const it of items) {
        const bad = validateMapping(it);
        if (bad) return res.status(400).json({ error: bad });
        assertEntityAccess(req, [Number(it.entity_id)]);
      }
      const ids = [];
      const tx = db.transaction(() => { for (const it of items) ids.push(createMapping(db, it, who(req))); });
      tx();
      res.json({ ids, count: ids.length, success: true });
    } catch (e) {
      if (/UNIQUE/i.test(e.message)) return res.status(400).json({ error: 'That account is already mapped for this entity' });
      fail(res, e);
    }
  });

  app.put('/api/intercompany/mappings/:id', ...gate, (req, res) => {
    try {
      const cur = db.prepare('SELECT * FROM intercompany_accounts WHERE id = ?').get(req.params.id);
      if (!cur) return res.status(404).json({ error: 'Mapping not found' });
      const body = { ...req.body, entity_id: cur.entity_id };
      const bad = validateMapping(body);
      if (bad) return res.status(400).json({ error: bad });
      assertEntityAccess(req, [cur.entity_id]);
      updateMapping(db, Number(req.params.id), body, who(req));
      res.json({ success: true });
    } catch (e) {
      if (/UNIQUE/i.test(e.message)) return res.status(400).json({ error: 'That account is already mapped for this entity' });
      fail(res, e);
    }
  });

  app.delete('/api/intercompany/mappings/:id', ...gate, (req, res) => {
    try {
      const cur = db.prepare('SELECT * FROM intercompany_accounts WHERE id = ?').get(req.params.id);
      if (!cur) return res.status(404).json({ error: 'Mapping not found' });
      assertEntityAccess(req, [cur.entity_id]);
      deleteMapping(db, Number(req.params.id));
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

  // ── Reconciliation for one entity ──
  // Entity-scoped so each row can carry the account code and name on both
  // sides. Access is checked on the selected entity; a counterparty the caller
  // cannot see still appears, because its accounts are part of THIS entity's
  // intercompany position and hiding them would misstate it.
  app.get('/api/intercompany/reconcile/entity', ...gate, (req, res) => {
    try {
      const entity_id = Number(req.query.entity_id);
      if (!entity_id) return res.status(400).json({ error: 'entity_id is required' });
      assertEntityAccess(req, [entity_id]);
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      res.json(reconcileForEntity(db, entity_id, { computeBalances, as_of, tolerance: req.query.tolerance }));
    } catch (e) { fail(res, e); }
  });
}

module.exports = {
  registerIntercompanyRoutes,
  ensureSchema,
  reconcileForEntity,
  suggestMappings,
  listMappings,
  listCompanies,
  listPeople,
  isMarkedPerson,
  resolveCounterparties,
  matchCompany,
  looksIndividual,
  isIndividualInvestor,
  parseAccountName,
  matchEntity,
  normName,
  IC_TYPES,
};
