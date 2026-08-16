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

function listGroups(db) {
  const groups = db.prepare('SELECT * FROM intercompany_groups ORDER BY name COLLATE NOCASE').all();
  const members = db.prepare(`
    SELECT m.group_id, m.entity_id, e.name, e.code
    FROM intercompany_group_members m LEFT JOIN entities e ON e.id = m.entity_id
    ORDER BY e.name COLLATE NOCASE`).all();
  return groups.map(g => ({ ...g, members: members.filter(m => m.group_id === g.id) }));
}

function getGroupEntityIds(db, groupId) {
  return db.prepare('SELECT entity_id FROM intercompany_group_members WHERE group_id = ?')
    .all(groupId).map(r => r.entity_id);
}

function saveGroup(db, { id, name, notes, entity_ids }, who) {
  const now = new Date().toISOString();
  let gid = id;
  if (gid) {
    db.prepare('UPDATE intercompany_groups SET name=?, notes=?, updated_at=?, updated_by=? WHERE id=?')
      .run(name, notes || null, now, who || null, gid);
  } else {
    gid = db.prepare('INSERT INTO intercompany_groups (name, notes, created_at, created_by) VALUES (?,?,?,?)')
      .run(name, notes || null, now, who || null).lastInsertRowid;
  }
  db.prepare('DELETE FROM intercompany_group_members WHERE group_id = ?').run(gid);
  const ins = db.prepare('INSERT OR IGNORE INTO intercompany_group_members (group_id, entity_id) VALUES (?,?)');
  for (const eid of (entity_ids || [])) ins.run(gid, Number(eid));
  return gid;
}

function deleteGroup(db, id) {
  db.prepare('DELETE FROM intercompany_group_members WHERE group_id = ?').run(id);
  db.prepare('DELETE FROM intercompany_groups WHERE id = ?').run(id);
}

// ══════════════════════════════ Mappings ══════════════════════════════

function listMappings(db, { entity_id, group_id } = {}) {
  let where = '', params = [];
  if (entity_id) { where = 'WHERE m.entity_id = ?'; params = [entity_id]; }
  else if (group_id) {
    const ids = getGroupEntityIds(db, group_id);
    if (!ids.length) return [];
    where = 'WHERE m.entity_id IN (' + ids.map(() => '?').join(',') + ')';
    params = ids;
  }
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

// Scope is every entity that has at least one mapping. An entity with no
// mappings cannot take part in a comparison anyway - there is nothing on its
// side to compare against - so including it would only add empty rows.
function loadContext(db, { computeBalances, as_of, allowEntity }) {
  ensureSchema(db);
  const entities = db.prepare('SELECT id, name, code FROM entities').all();
  const entById = new Map(entities.map(e => [e.id, e]));
  let mappings = resolveCounterparties(db, listMappings(db, {}));
  if (allowEntity) mappings = mappings.filter(m => allowEntity(m.entity_id));
  const memberIds = [...new Set(mappings.map(m => m.entity_id))];
  const inScope = new Set(memberIds);
  const balances = new Map();
  for (const eid of memberIds) {
    const rows = computeBalances(eid, as_of ? { as_of } : {});
    balances.set(eid, new Map(rows.map(r => [String(r.code), r])));
  }
  return { memberIds, inScope, entById, balances, mappings };
}

const nameOf = (entById, id) => (entById.get(id) || {}).name || ('Entity ' + id);

// Amount for one mapping row at the as-of date, in natural sign
// (asset debit-positive, liability/equity credit-positive), which is what
// computeBalances already returns.
function amountFor(balances, m) {
  const row = (balances.get(m.entity_id) || new Map()).get(String(m.account_code));
  return row ? Number(row.balance) || 0 : 0;
}

// Is this mapping's counterparty inside the selected group?
function isInternal(m, inScope) {
  if (m.is_external) return false;
  if (m.counterparty_entity_id == null) return false;
  return inScope.has(Number(m.counterparty_entity_id));
}

// Every intercompany-looking account that carries a balance but has no mapping
// row. Reported so an unmapped account can never silently drop out of the
// reconciliation — the failure mode that lets a real balance disappear.
function findUnmapped(db, memberIds, balances, mappings, wantedTypes) {
  const mapped = new Set(mappings.map(m => m.entity_id + '|' + m.account_code));
  const people = listPeople(db);
  const out = [];
  for (const eid of memberIds) {
    const accounts = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ?').all(eid);
    for (const a of accounts) {
      if (mapped.has(eid + '|' + String(a.code))) continue;
      const parsed = parseAccountName(a.name);
      if (!parsed || !wantedTypes.includes(parsed.ic_type)) continue;
      // An outside individual's balance is not an unmapped intercompany one -
      // it is simply not intercompany.
      if (isIndividualInvestor(parsed) || isMarkedPerson(people, parsed.label)) continue;
      const row = (balances.get(eid) || new Map()).get(String(a.code));
      const bal = row ? Number(row.balance) || 0 : 0;
      if (Math.abs(bal) < DEFAULT_TOLERANCE) continue;
      out.push({ entity_id: eid, account_code: String(a.code), account_name: a.name,
                 ic_type: parsed.ic_type, counterparty_label: parsed.label, balance: bal });
    }
  }
  return out.sort((a, b) => Math.abs(b.balance) - Math.abs(a.balance));
}

// ════════════════════ Due from / Due to reconciliation ════════════════════
//
// Netted by counterparty PAIR. A's net position toward B is
// (sum of A's due-from B) − (sum of A's due-to B); B's net toward A is the
// same computed from B's side. If the relationship is stated consistently the
// two are equal and opposite, so their SUM is the unreconciled difference.

function reconcileDueFromTo(db, opts) {
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;
  const { memberIds, inScope, entById, balances, mappings } = loadContext(db, opts);
  const rows = mappings.filter(m => m.ic_type === 'due_from' || m.ic_type === 'due_to');

  const pairs = new Map();   // "loId|hiId" -> pair record
  const external = [];

  for (const m of rows) {
    const amt = amountFor(balances, m);
    const leg = {
      mapping_id: m.id, entity_id: m.entity_id, entity_name: nameOf(entById, m.entity_id),
      account_code: m.account_code, account_name: m.account_name, ic_type: m.ic_type,
      amount: amt, notes: m.notes || null,
    };

    if (!isInternal(m, inScope)) {
      // No second ledger to compare against, so this can never match and never
      // eliminates. Reported on its own rather than mixed into the pairs - the
      // CLRFI Midco I draft eliminated three of these by mistake.
      external.push({
        ...leg,
        counterparty_entity_id: m.counterparty_entity_id || null,
        counterparty_name: m.counterparty_entity_id
          ? nameOf(entById, m.counterparty_entity_id)
          : (m.notes || 'External party'),
        tag: 'no_elim',
        off_ledger: !!m.off_ledger,
        reason: m.is_external
          ? 'Marked external in IC Mapping'
          : (m.off_ledger
              ? 'Registered company with no ledger in CloudLedger — nothing to match against'
              : (m.counterparty_entity_id == null
                  ? 'No counterparty mapped'
                  : 'That entity has no intercompany mappings of its own yet')),
      });
      continue;
    }

    const a = Number(m.entity_id), b = Number(m.counterparty_entity_id);
    const key = Math.min(a, b) + '|' + Math.max(a, b);
    if (!pairs.has(key)) {
      const lo = Math.min(a, b), hi = Math.max(a, b);
      pairs.set(key, {
        entity_a_id: lo, entity_a_name: nameOf(entById, lo),
        entity_b_id: hi, entity_b_name: nameOf(entById, hi),
        a_legs: [], b_legs: [], a_net: 0, b_net: 0,
      });
    }
    const p = pairs.get(key);
    const signed = m.ic_type === 'due_from' ? amt : -amt;  // receivable positive
    if (a === p.entity_a_id) { p.a_legs.push(leg); p.a_net += signed; }
    else { p.b_legs.push(leg); p.b_net += signed; }
  }

  const pairRows = [...pairs.values()].map(p => {
    const difference = p.a_net + p.b_net;
    const eliminate = (p.a_net * p.b_net < 0) ? Math.min(Math.abs(p.a_net), Math.abs(p.b_net)) : 0;
    // ONE_SIDED is the mapping-completeness signal: this side mapped the
    // relationship and the other side never did, so the comparison cannot be
    // made yet. That is not the same as the two sides disagreeing, and calling
    // it a mismatch would send someone hunting a difference that does not exist.
    const oneSided = !p.a_legs.length || !p.b_legs.length;
    const bothEmpty = Math.abs(p.a_net) < tol && Math.abs(p.b_net) < tol;
    let status;
    if (bothEmpty) status = 'empty';
    else if (oneSided) status = 'one_sided';
    else if (Math.abs(difference) < tol) status = 'matched';
    else status = 'mismatch';
    // Both sides receivable (or both payable) — one ledger has the sign wrong.
    const same_direction = Math.abs(p.a_net) >= tol && Math.abs(p.b_net) >= tol && p.a_net * p.b_net > 0;
    return { ...p, difference, eliminate, status, same_direction };
  }).sort((x, y) => Math.abs(y.difference) - Math.abs(x.difference));
  const settledCount = pairRows.filter(p => p.status === 'empty').length;
  const livePairs = pairRows.filter(p => p.status !== 'empty');

  const unmapped = findUnmapped(db, memberIds, balances, mappings, ['due_from', 'due_to']);

  return {
    as_of: opts.as_of || null,
    tolerance: tol,
    entities: memberIds.map(id => ({ id, name: nameOf(entById, id) })),
    pairs: livePairs,
    external,
    unmapped,
    totals: {
      pair_count: livePairs.length,
      matched: livePairs.filter(p => p.status === 'matched').length,
      mismatched: livePairs.filter(p => p.status === 'mismatch').length,
      one_sided: livePairs.filter(p => p.status === 'one_sided').length,
      settled: settledCount,
      total_eliminate: round2(livePairs.reduce((s, p) => s + p.eliminate, 0)),
      total_difference: round2(livePairs.reduce((s, p) => s + p.difference, 0)),
      // Genuine disagreements only. A one-sided pair is an unfinished mapping,
      // not a difference, and folding it in here would overstate the problem.
      abs_difference: round2(livePairs.filter(p => p.status === 'mismatch')
        .reduce((s, p) => s + Math.abs(p.difference), 0)),
      external_count: external.length,
      external_total: round2(external.reduce((s, r) => s + Math.abs(r.amount), 0)),
      unmapped_count: unmapped.length,
      entity_count: memberIds.length,
    },
  };
}

// ═══════════ Investment / Contributed capital reconciliation ═══════════

function reconcileInvestmentCapital(db, opts) {
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;
  const { memberIds, inScope, entById, balances, mappings } = loadContext(db, opts);
  const rows = mappings.filter(m => m.ic_type === 'investment' || m.ic_type === 'contributed_capital');

  const findings = [];   // self-referential errors
  const pairs = new Map();
  const external = [];

  for (const m of rows) {
    const amt = amountFor(balances, m);
    const leg = {
      mapping_id: m.id, entity_id: m.entity_id, entity_name: nameOf(entById, m.entity_id),
      account_code: m.account_code, account_name: m.account_name, ic_type: m.ic_type,
      amount: amt, notes: m.notes || null,
    };

    // ── The one red flag: an entity holding an investment in ITSELF. ──
    // CLIP Property Owner 19041 (1,837,842.67) and CLR Silsbee Property Owner
    // 17001 (11,760,052.36) both do this, grossing up their own assets and
    // equity. A parent holding an investment in a child is normal and is not
    // flagged here.
    if (m.counterparty_entity_id != null && Number(m.counterparty_entity_id) === Number(m.entity_id)) {
      findings.push({
        ...leg,
        severity: m.ic_type === 'investment' ? 'error' : 'warning',
        finding: m.ic_type === 'investment' ? 'self_investment' : 'self_contributed_capital',
        message: m.ic_type === 'investment'
          ? nameOf(entById, m.entity_id) + ' holds an investment in itself — this grosses up its own assets and equity by ' + fmtMoney(amt) + '.'
          : nameOf(entById, m.entity_id) + ' shows contributed capital from itself (' + fmtMoney(amt) + ').',
      });
      continue;
    }

    if (!isInternal(m, inScope)) {
      external.push({
        ...leg,
        counterparty_entity_id: m.counterparty_entity_id || null,
        counterparty_name: m.counterparty_entity_id
          ? nameOf(entById, m.counterparty_entity_id)
          : (m.notes || 'External party'),
        tag: 'no_elim',
        off_ledger: !!m.off_ledger,
        reason: m.is_external
          ? 'Marked external in IC Mapping'
          : (m.off_ledger
              ? 'Registered company with no ledger in CloudLedger — nothing to match against'
              : (m.counterparty_entity_id == null
                  ? 'No counterparty mapped'
                  : 'That entity has no intercompany mappings of its own yet')),
      });
      continue;
    }

    // Key by (parent, child): the parent is whoever holds the investment.
    const parent = m.ic_type === 'investment' ? Number(m.entity_id) : Number(m.counterparty_entity_id);
    const child  = m.ic_type === 'investment' ? Number(m.counterparty_entity_id) : Number(m.entity_id);
    const key = parent + '>' + child;
    if (!pairs.has(key)) {
      pairs.set(key, {
        parent_id: parent, parent_name: nameOf(entById, parent),
        child_id: child, child_name: nameOf(entById, child),
        investment_legs: [], capital_legs: [], investment: 0, contributed_capital: 0,
      });
    }
    const p = pairs.get(key);
    if (m.ic_type === 'investment') { p.investment_legs.push(leg); p.investment += amt; }
    else { p.capital_legs.push(leg); p.contributed_capital += amt; }
  }

  const pairRows = [...pairs.values()].map(p => {
    const difference = p.investment - p.contributed_capital;
    const has_inv = Math.abs(p.investment) >= tol, has_cap = Math.abs(p.contributed_capital) >= tol;
    let status = 'matched';
    if (!has_inv && !has_cap) status = 'empty';
    else if (!has_inv || !has_cap) status = 'one_sided';
    else if (Math.abs(difference) >= tol) status = 'mismatch';
    const eliminate = Math.min(Math.abs(p.investment), Math.abs(p.contributed_capital));
    return { ...p, difference, eliminate, status };
  }).filter(p => p.status !== 'empty')
    .sort((x, y) => Math.abs(y.difference) - Math.abs(x.difference));

  const unmapped = findUnmapped(db, memberIds, balances, mappings, ['investment', 'contributed_capital']);

  return {
    as_of: opts.as_of || null,
    tolerance: tol,
    entities: memberIds.map(id => ({ id, name: nameOf(entById, id) })),
    findings: findings.sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount)),
    pairs: pairRows,
    external,
    unmapped,
    totals: {
      error_count: findings.filter(f => f.severity === 'error').length,
      warning_count: findings.filter(f => f.severity === 'warning').length,
      self_investment_total: round2(findings.filter(f => f.finding === 'self_investment')
        .reduce((s, f) => s + Math.abs(f.amount), 0)),
      pair_count: pairRows.length,
      entity_count: memberIds.length,
      matched: pairRows.filter(p => p.status === 'matched').length,
      one_sided: pairRows.filter(p => p.status === 'one_sided').length,
      mismatched: pairRows.filter(p => p.status === 'mismatch').length,
      total_eliminate: round2(pairRows.reduce((s, p) => s + p.eliminate, 0)),
      total_difference: round2(pairRows.reduce((s, p) => s + p.difference, 0)),
      external_count: external.length,
      external_total: round2(external.reduce((s, r) => s + Math.abs(r.amount), 0)),
      unmapped_count: unmapped.length,
    },
  };
}

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function fmtMoney(n) {
  return '$' + Math.abs(Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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

  // ── Groups ──
  app.get('/api/intercompany/groups', ...gate, (req, res) => {
    try {
      const all = listGroups(db);
      const visible = all.filter(g => g.members.every(m => userHasEntityAccess(req.user.id, req.user.role, m.entity_id)));
      res.json(visible);
    } catch (e) { fail(res, e); }
  });

  app.post('/api/intercompany/groups', ...gate, (req, res) => {
    try {
      if (!req.body || !req.body.name) return res.status(400).json({ error: 'Group name is required' });
      const ids = (req.body.entity_ids || []).map(Number);
      assertEntityAccess(req, ids);
      const id = saveGroup(db, { name: req.body.name, notes: req.body.notes, entity_ids: ids }, who(req));
      res.json({ id, success: true });
    } catch (e) { fail(res, e); }
  });

  app.put('/api/intercompany/groups/:id', ...gate, (req, res) => {
    try {
      if (!req.body || !req.body.name) return res.status(400).json({ error: 'Group name is required' });
      const ids = (req.body.entity_ids || []).map(Number);
      assertEntityAccess(req, ids);
      saveGroup(db, { id: Number(req.params.id), name: req.body.name, notes: req.body.notes, entity_ids: ids }, who(req));
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/intercompany/groups/:id', ...gate, (req, res) => {
    try {
      assertEntityAccess(req, getGroupEntityIds(db, Number(req.params.id)));
      deleteGroup(db, Number(req.params.id));
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

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
      const group_id = req.query.group_id ? Number(req.query.group_id) : null;
      if (entity_id) assertEntityAccess(req, [entity_id]);
      if (group_id) assertEntityAccess(req, getGroupEntityIds(db, group_id));
      const rows = listMappings(db, { entity_id, group_id });
      res.json(entity_id || group_id
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

  // ── Reconciliation ──
  // No group: the reconciliation covers every entity that has mappings, so the
  // whole picture arrives in one request. Entities the caller cannot see are
  // filtered out rather than refused, so a scoped user still gets an answer.
  const runRecon = (fn) => (req, res) => {
    try {
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      res.json(fn(db, {
        computeBalances, as_of, tolerance: req.query.tolerance,
        allowEntity: eid => userHasEntityAccess(req.user.id, req.user.role, eid),
      }));
    } catch (e) { fail(res, e); }
  };

  app.get('/api/intercompany/reconcile/due', ...gate, runRecon(reconcileDueFromTo));
  app.get('/api/intercompany/reconcile/investment', ...gate, runRecon(reconcileInvestmentCapital));
}

module.exports = {
  registerIntercompanyRoutes,
  ensureSchema,
  reconcileDueFromTo,
  reconcileInvestmentCapital,
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
