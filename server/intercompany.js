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

const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

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
  // The specific account on the counterparty's ledger that answers this one.
  // Optional, and deliberately NOT authoritative for the arithmetic — see
  // counterpartyCandidates below for why a pair is not always one-to-one.
  if (!icCols.includes('counterparty_account_code')) {
    db.exec("ALTER TABLE intercompany_accounts ADD COLUMN counterparty_account_code TEXT");
    console.log('[db migrate] intercompany_accounts.counterparty_account_code added');
  }
  // The other side of an EXTERNAL mapping keeps no ledger here, but its
  // statement still names a number. Entered by hand on the reconciliation
  // page, kept per as-of date so each period stands alone.
  db.exec(`
    CREATE TABLE IF NOT EXISTS intercompany_manual_balances (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      mapping_id INTEGER NOT NULL,
      as_of TEXT NOT NULL DEFAULT '',
      balance REAL NOT NULL,
      updated_at TEXT, updated_by TEXT,
      UNIQUE(mapping_id, as_of)
    );
  `);
  // Uploaded trial balances for counterparties that keep no ledger in
  // CloudLedger (JVs, sponsor holdcos registered as org nodes). One TB per
  // (company, as-of date); re-uploading replaces. Balances stored in NATURAL
  // terms — asset/expense debit-positive, liability/equity credit-positive —
  // so a line reads exactly like a computeBalances row.
  db.exec(`
    CREATE TABLE IF NOT EXISTS external_tb_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      node_id INTEGER NOT NULL,
      as_of TEXT NOT NULL,
      account_code TEXT NOT NULL,
      account_name TEXT,
      type TEXT,
      balance REAL NOT NULL DEFAULT 0,
      uploaded_at TEXT, uploaded_by TEXT, filename TEXT,
      UNIQUE(node_id, as_of, account_code)
    );
    CREATE INDEX IF NOT EXISTS idx_etb_node ON external_tb_lines(node_id, as_of);
  `);
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
           ${nodeName} AS counterparty_node_name,
           ca.name AS counterparty_account_name
    FROM intercompany_accounts m
    LEFT JOIN entities e ON e.id = m.entity_id
    LEFT JOIN entities c ON c.id = m.counterparty_entity_id
    LEFT JOIN accounts ca ON ca.entity_id = m.counterparty_entity_id
      AND ca.code = m.counterparty_account_code
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
    (entity_id, account_code, account_name, counterparty_entity_id, counterparty_node_id, counterparty_account_code, ic_type, is_external, notes, created_at, created_by, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(
      Number(body.entity_id), String(body.account_code), body.account_name || null,
      body.is_external ? null : (body.counterparty_entity_id ? Number(body.counterparty_entity_id) : null),
      body.is_external ? null : (body.counterparty_node_id ? Number(body.counterparty_node_id) : null),
      body.is_external ? null : (body.counterparty_account_code ? String(body.counterparty_account_code) : null),
      body.ic_type, body.is_external ? 1 : 0, body.notes || null, now, who || null, now, who || null);
  return r.lastInsertRowid;
}

function updateMapping(db, id, body, who) {
  const now = new Date().toISOString();
  db.prepare(`UPDATE intercompany_accounts SET
      account_code=?, account_name=?, counterparty_entity_id=?, counterparty_node_id=?, counterparty_account_code=?, ic_type=?, is_external=?, notes=?, updated_at=?, updated_by=?
    WHERE id=?`).run(
      String(body.account_code), body.account_name || null,
      body.is_external ? null : (body.counterparty_entity_id ? Number(body.counterparty_entity_id) : null),
      body.is_external ? null : (body.counterparty_node_id ? Number(body.counterparty_node_id) : null),
      body.is_external ? null : (body.counterparty_account_code ? String(body.counterparty_account_code) : null),
      body.ic_type, body.is_external ? 1 : 0, body.notes || null, now, who || null, id);
}

function deleteMapping(db, id) {
  db.prepare('DELETE FROM intercompany_accounts WHERE id = ?').run(id);
}

// One person confirms a pair ONCE. A mapping that names a CL counterparty and
// that counterparty's GL account fully determines the row the counterparty
// would write — same two accounts, kinds mirrored — so it is written here
// automatically. The other entity's worklist must never ask a second person
// to confirm a fact the first already stated. An account the counterparty has
// already mapped itself is never touched: their own decision wins.
function ensureMirrorMapping(db, mappingId, who) {
  const m = db.prepare('SELECT * FROM intercompany_accounts WHERE id = ?').get(mappingId);
  if (!m || m.is_external || m.counterparty_entity_id == null || !m.counterparty_account_code) return null;
  const cpEid = Number(m.counterparty_entity_id);
  if (cpEid === Number(m.entity_id)) return null;
  const mir = MIRROR_ACCOUNT[m.ic_type];
  if (!mir) return null;
  const acct = db.prepare('SELECT code, name FROM accounts WHERE entity_id = ? AND code = ?')
    .get(cpEid, String(m.counterparty_account_code));
  if (!acct) return null; // broken pin — nothing real to mirror onto
  const existing = db.prepare('SELECT id FROM intercompany_accounts WHERE entity_id = ? AND account_code = ?')
    .get(cpEid, String(m.counterparty_account_code));
  if (existing) return null;
  const src = db.prepare('SELECT name FROM entities WHERE id = ?').get(Number(m.entity_id));
  const now = new Date().toISOString();
  const r = db.prepare(`INSERT INTO intercompany_accounts
    (entity_id, account_code, account_name, counterparty_entity_id, counterparty_node_id, counterparty_account_code, ic_type, is_external, notes, created_at, created_by, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(
      cpEid, String(m.counterparty_account_code), acct.name,
      Number(m.entity_id), null, String(m.account_code),
      mir.ic_type, 0,
      'mirrored automatically — the pair was confirmed on ' + ((src && src.name) || ('entity ' + m.entity_id)),
      now, who || 'auto-mirror', now, who || 'auto-mirror');
  return r.lastInsertRowid;
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

// A shell entity holds no operations — it exists to sit between a fund and the
// company that owns the asset. Its contributed capital is the fund's money
// passing through, not a balance any counterparty ledger in CloudLedger states
// back. Reconciling it produces a permanent one-sided row, so those accounts are
// left out of the unmapped worklist for shells specifically. Due-from / due-to
// and investment accounts on a shell are still reconciled normally.
//
// `entities.entity_type` is one of development / accounting / shell / operating.
// The column arrives by migration, so a database restored without it degrades to
// "not a shell" rather than throwing.
function isShellEntity(entity) {
  return String((entity && entity.entity_type) || '').trim().toLowerCase() === 'shell';
}

// ── Which account on the other side answers this one ──
//
// Account NAMES cannot answer this. A "Due from X" account is an asset, but it
// can carry a credit — BROZ Fund I's "23000 Due to Banyan Residential" sits at
// -2,900, a liability with a debit balance. So the search is by AMOUNT, in
// receivable terms, where the two sides of a healthy relationship sum to zero:
//
//     asset      →  +balance   (computeBalances gives Dr-Cr for an asset)
//     liability  →  -balance   (it gives Cr-Dr, so a payable is negative here)
//
// That is the same convention legOf() uses, and it reproduces the CLIP ↔ CLR
// Silsbee gap: 26,948.26 + (-202,420.88) + 215,228.93 = 39,756.31.
function receivableSigned(acct) {
  if (!acct) return null;
  if (acct.type === 'Asset') return Number(acct.balance) || 0;
  if (acct.type === 'Liability') return -(Number(acct.balance) || 0);
  return null; // equity and P&L are not a due-from/due-to position
}

// Rank the counterparty's accounts as answers to ONE of our accounts.
//
// Amount alone is not enough and must not be the primary key. Measured against
// the real 6/30/2026 ledger, only 2 of 28 mapped relationships tie to the penny
// — the other 26 have the obviously-right counterparty account carrying a
// DIFFERENT number, which is the whole point of reconciling. Ranking by amount
// alone picked "10100 Operating Checking" and "12002 Allowance for Credit
// Losses" as best answers. So: does the account name resolve back to US first,
// closeness of the offset second, and the gap is always reported so a person
// sees immediately whether it ties.
function counterpartyCandidates(db, opts) {
  ensureSchema(db);
  const { entityId, accountCode, counterpartyEntityId, computeBalances, as_of } = opts;
  const eid = Number(entityId), cid = Number(counterpartyEntityId);
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;

  const entities = db.prepare('SELECT id, name, code FROM entities').all();
  const companies = listCompanies(db);
  const us = entities.find(e => Number(e.id) === eid);
  const them = entities.find(e => Number(e.id) === cid);
  if (!us || !them) { const e = new Error('Entity not found'); e.status = 404; throw e; }

  const ourBal = computeBalances(eid, as_of ? { as_of } : {});
  const ourAcct = ourBal.find(b => String(b.code) === String(accountCode));
  const ourSigned = receivableSigned(ourAcct);

  // Accounts on the counterparty already spoken for by a mapping that points
  // somewhere else — flagged, not hidden, because a wrong existing mapping is
  // exactly the kind of thing this screen should let someone notice.
  const theirMaps = new Map(listMappings(db, { entity_id: cid })
    .map(m => [String(m.account_code), m]));

  // EVERY balance-sheet account on the counterparty's chart is offered —
  // zero-activity, equity, bank accounts included: any of them can be the one
  // a person needs to point at. Only P&L stays out (no balance for a
  // counterparty ledger to agree with). Ranking still puts the accounts that
  // name us or offset first, so completeness costs nothing at the top.
  const cpBalances = new Map(computeBalances(cid, as_of ? { as_of } : {}).map(x => [String(x.code), x]));
  const rows = db.prepare('SELECT code, name, type, bank_acct FROM accounts WHERE entity_id = ? ORDER BY code').all(cid)
    .filter(a => !isPnlAccount(a.type, a.code))
    .map(a => {
      const held = cpBalances.get(String(a.code));
      const b = { code: a.code, name: a.name, type: a.type, bank_acct: a.bank_acct,
        balance: held ? held.balance : 0 };
      const signed = receivableSigned(b); // null for equity — kept, just unranked by gap
      const parsed = parseAccountName(b.name);
      // Does THEIR account name point back at US? Reuses the same matcher the
      // suggestions use, so aliases ("Due to SRN", "Due from Buna") resolve.
      let nameMatch = false;
      if (parsed) {
        const m = matchCompany(parsed.label, entities, companies);
        nameMatch = m.entity_id != null && Number(m.entity_id) === eid;
      }
      const gap = (ourSigned == null || signed == null) ? null : round2(Math.abs(signed + ourSigned));
      const mapped = theirMaps.get(String(b.code));
      return {
        account_code: String(b.code), account_name: b.name, type: b.type,
        balance: round2(b.balance), signed: signed == null ? null : round2(signed),
        ic_type: parsed ? parsed.ic_type : null,
        counterparty_label: parsed ? parsed.label : null,
        name_match: nameMatch,
        gap,
        offsets: gap != null && gap < tol,
        already_mapped_to: mapped
          ? (mapped.counterparty_name || null)
          : null,
        already_mapped_to_us: !!(mapped && Number(mapped.counterparty_entity_id) === eid),
      };
    })
    .filter(Boolean);

  // Name first, then how close the offset is. An account that both names us and
  // ties to the penny is the only thing that gets pre-selected without doubt,
  // but naming us while disagreeing on the number is still the right pick — the
  // disagreement is the finding.
  const rank = r => (r.name_match || r.already_mapped_to_us ? 0 : 1);
  rows.sort((a, b) =>
    rank(a) - rank(b)
    || (a.gap == null ? 1 : 0) - (b.gap == null ? 1 : 0)
    || (a.gap - b.gap)
    || String(a.account_code).localeCompare(String(b.account_code)));

  // Only pre-select something defensible: it names us, or it offsets exactly.
  const top = rows[0];
  const best = top && (top.name_match || top.already_mapped_to_us || top.offsets)
    ? top.account_code : null;

  return {
    entity: { id: us.id, name: us.name },
    counterparty: { id: them.id, name: them.name },
    as_of: as_of || null,
    our_account: ourAcct
      ? { account_code: String(ourAcct.code), account_name: ourAcct.name,
          type: ourAcct.type, balance: round2(ourAcct.balance), signed: round2(ourSigned) }
      : { account_code: String(accountCode), account_name: null, type: null, balance: 0, signed: 0 },
    best,
    exact_count: rows.filter(r => r.offsets).length,
    candidates: rows,
  };
}

// The MAPPED accounts of one entity, with both sides' GL accounts on one row.
//
// A mapping is mapped only when it names a counterparty ENTITY *and* that
// entity's GL ACCOUNT. That is what makes it checkable: two account codes and
// two balances that either agree or do not. Naming the entity alone leaves
// nothing to compare, so it is unfinished work and belongs on the to-do list,
// not here.
//
// Follows from the same rule: a counterparty with no ledger in CloudLedger can
// never be mapped, because there is no GL account to name. That covers external
// parties and off-ledger companies (QOZBs, sponsor holdcos). They are not
// mapped and they are not to-do either — nobody can ever finish them — so they
// are excluded from both lists and reported as a count.
//
// The counterparty's balance is read straight from its own ledger, so a pair is
// checkable whether or not the counterparty has mapped anything itself.
// The uploaded trial balance that answers for a ledger-less counterparty at a
// date: the latest one on or before as_of (a June recon reads the June TB; if
// only May exists, May answers and the date is said out loud). Null when the
// company has no TB at all.
function externalTbFor(db, nodeId, as_of) {
  try {
    const r = as_of
      ? db.prepare('SELECT as_of FROM external_tb_lines WHERE node_id = ? AND as_of <= ? ORDER BY as_of DESC LIMIT 1').get(Number(nodeId), String(as_of))
      : db.prepare('SELECT as_of FROM external_tb_lines WHERE node_id = ? ORDER BY as_of DESC LIMIT 1').get(Number(nodeId));
    if (!r) return null;
    const lines = db.prepare('SELECT account_code, account_name, type, balance FROM external_tb_lines WHERE node_id = ? AND as_of = ? ORDER BY account_code')
      .all(Number(nodeId), r.as_of);
    return { as_of: r.as_of, lines, byCode: new Map(lines.map(l => [String(l.account_code), l])) };
  } catch (e) { return null; }
}

function listMappedPairs(db, entityId, { computeBalances, as_of }) {
  ensureSchema(db);
  const eid = Number(entityId);
  const tol = DEFAULT_TOLERANCE;
  const all = listMappings(db, { entity_id: eid });
  const mine = new Map(computeBalances(eid, as_of ? { as_of } : {}).map(b => [String(b.code), b]));

  const cpIds = [...new Set(all
    .filter(m => !m.is_external && m.counterparty_entity_id != null)
    .map(m => Number(m.counterparty_entity_id)))];
  const theirBal = new Map();
  const theirAccts = new Map(); // Track actual accounts, not just ones with balances
  for (const id of cpIds) {
    theirBal.set(id, new Map(computeBalances(id, as_of ? { as_of } : {}).map(b => [String(b.code), b])));
    // Also load all accounts from the counterparty's chart, regardless of balance.
    // A newly created account (zero balance, no transactions) won't appear in
    // computeBalances, so we need a separate check.
    const accts = new Map(db.prepare('SELECT code, name, type FROM accounts WHERE entity_id=?').all(id).map(a => [String(a.code), a]));
    theirAccts.set(id, accts);
  }

  const pairRows = all
    .filter(m => !m.is_external
      && m.counterparty_entity_id != null
      && m.counterparty_account_code)
    .map(m => {
      const ours = mine.get(String(m.account_code)) || null;
      const ourSigned = receivableSigned(ours);
      const acctRec = m.counterparty_account_code
        ? (theirAccts.get(Number(m.counterparty_entity_id)) || new Map()).get(String(m.counterparty_account_code)) || null
        : null;
      // Balance rows only exist for accounts with journal lines. An account
      // that exists but has none IS at zero — say 0.00, not blank/unknown.
      const theirs = m.counterparty_account_code
        ? ((theirBal.get(Number(m.counterparty_entity_id)) || new Map()).get(String(m.counterparty_account_code))
            || (acctRec ? { name: acctRec.name, type: acctRec.type, balance: 0 } : null))
        : null;
      const theirSigned = receivableSigned(theirs);
      const gap = (ourSigned != null && theirSigned != null)
        ? round2(Math.abs(ourSigned + theirSigned)) : null;
      return {
        id: m.id,
        account_code: String(m.account_code),
        account_name: m.account_name,
        ic_type: m.ic_type,
        balance: ours ? round2(ours.balance) : 0,
        signed: ourSigned == null ? null : round2(ourSigned),
        counterparty_entity_id: m.counterparty_entity_id,
        counterparty_node_id: m.counterparty_node_id || null,
        counterparty_name: m.counterparty_name || null,
        self: Number(m.counterparty_entity_id) === eid,
        their_account_code: m.counterparty_account_code || null,
        their_account_name: theirs ? theirs.name : null,
        their_type: theirs ? theirs.type : null,
        their_balance: theirs ? round2(theirs.balance) : null,
        their_signed: theirSigned == null ? null : round2(theirSigned),
        // The counterparty account was named but is not in its ledger at all —
        // renamed, renumbered or deleted since. Worth saying out loud.
        // Check the accounts table directly, not just balances, because a newly
        // created account (zero balance, no transactions) won't appear in computeBalances.
        their_account_missing: !!(m.counterparty_account_code && !acctRec),
        gap,
        offsets: gap != null && gap < tol,
        notes: m.notes || null,
      };
    })
    .sort((a, b) => Math.abs(b.balance) - Math.abs(a.balance)
      || String(a.account_code).localeCompare(String(b.account_code)));

  // Investor capital: contributed capital whose contributor keeps no ledger in
  // CloudLedger. Complete BY RULE without a counterparty account — the only
  // sanctioned exception to "mapped means both sides named" — so it belongs on
  // this list, tagged, rather than on a to-do list nobody can ever finish.
  const investorRows = all
    .filter(m => m.ic_type === 'contributed_capital'
      && !m.counterparty_account_code
      && (m.is_external || (m.counterparty_entity_id == null && m.counterparty_node_id != null)))
    .map(m => {
      const ours = mine.get(String(m.account_code)) || null;
      return {
        id: m.id,
        account_code: String(m.account_code),
        account_name: m.account_name,
        ic_type: m.ic_type,
        balance: ours ? round2(ours.balance) : 0,
        signed: null,
        counterparty_entity_id: null,
        counterparty_node_id: m.counterparty_node_id || null,
        counterparty_name: m.counterparty_name || m.notes || 'Outside investor',
        self: false,
        investor_capital: true,
        their_account_code: null, their_account_name: null, their_type: null,
        their_balance: null, their_signed: null, their_account_missing: false,
        gap: null, offsets: false,
        notes: m.notes || null,
      };
    });
  // Counterparties answered by an UPLOADED trial balance: the mapping names a
  // company node and one of its TB lines. Their balance is read from the
  // latest TB on or before as_of, and the date used is said on the row.
  const tbRows = all
    .filter(m => !m.is_external && m.counterparty_entity_id == null
      && m.counterparty_node_id != null && m.counterparty_account_code)
    .map(m => {
      const ours = mine.get(String(m.account_code)) || null;
      const ourSigned = receivableSigned(ours);
      const tb = externalTbFor(db, m.counterparty_node_id, as_of);
      const line = tb ? tb.byCode.get(String(m.counterparty_account_code)) || null : null;
      const theirSigned = receivableSigned(line);
      const gap = (ourSigned != null && theirSigned != null)
        ? round2(Math.abs(ourSigned + theirSigned)) : null;
      return {
        id: m.id,
        account_code: String(m.account_code),
        account_name: m.account_name,
        ic_type: m.ic_type,
        balance: ours ? round2(ours.balance) : 0,
        signed: ourSigned == null ? null : round2(ourSigned),
        counterparty_entity_id: null,
        counterparty_node_id: m.counterparty_node_id,
        counterparty_name: m.counterparty_name || null,
        self: false,
        from_tb: true,
        tb_as_of: tb ? tb.as_of : null,
        tb_missing: !tb,
        their_account_code: m.counterparty_account_code || null,
        their_account_name: line ? line.account_name : null,
        their_type: line ? line.type : null,
        their_balance: line ? round2(line.balance) : null,
        their_signed: theirSigned == null ? null : round2(theirSigned),
        their_account_missing: !!(tb && !line),
        gap,
        offsets: gap != null && gap < tol,
        notes: m.notes || null,
      };
    });
  const rows = pairRows.concat(investorRows).concat(tbRows)
    .sort((a, b) => Math.abs(b.balance) - Math.abs(a.balance)
      || String(a.account_code).localeCompare(String(b.account_code)));

  return {
    entity_id: eid,
    as_of: as_of || null,
    count: rows.length,
    matched: pairRows.filter(r => r.offsets).length,
    mismatched: pairRows.filter(r => r.gap != null && !r.offsets).length,
    // Named an account that is not in the counterparty's ledger — renamed,
    // renumbered or deleted since. Still listed, flagged, because a broken
    // pin that disappeared would be worse than one that is wrong out loud.
    broken: pairRows.concat(tbRows).filter(r => r.their_account_missing).length,
    investor_capital_count: investorRows.length,
    // Mappable, but not finished: names the entity, not the account.
    incomplete_count: all.filter(m => !m.is_external
      && m.counterparty_entity_id != null && !m.counterparty_account_code).length,
    // Not mappable at all — no ledger in CloudLedger to name an account on.
    // Investor capital is carved out: those are complete by rule and listed above.
    external_count: all.filter(m => m.is_external && m.ic_type !== 'contributed_capital').length,
    off_ledger_count: all.filter(m => !m.is_external && m.counterparty_entity_id == null
      && !(m.ic_type === 'contributed_capital' && m.counterparty_node_id != null)
      && !(m.counterparty_node_id != null && m.counterparty_account_code)).length,
    rows,
  };
}

// An intercompany balance can only live on the BALANCE SHEET. A revenue or
// expense account has no balance for a counterparty ledger to agree with, so
// mapping one would put a row in the reconciliation that can never reconcile —
// it would sit there as a permanent unexplained difference. Management fees and
// interest charged between affiliates are real intercompany TRANSACTIONS, but
// what they leave behind for this module is the due-from / due-to balance, and
// that is what gets mapped. So P&L accounts are kept out of the unmapped lists.
//
// `type` is authoritative. The code range is only the fallback for a chart that
// never set one, and matches how index.js derives type from a code:
// 40000-49999 Revenue, 50000-69999 Expense, 70000+ Revenue.
function isPnlAccount(type, code) {
  const t = String(type || '').trim().toLowerCase();
  if (t) {
    // A type that names the balance sheet wins outright. QuickBooks-style charts
    // say "Other Current Liability", "Fixed Asset", "Accounts Receivable", and
    // one of those must never be read as P&L: hiding a real intercompany account
    // is a far worse failure than listing an expense account someone can ignore.
    if (/\b(assets?|liabilit(y|ies)|equity|receivable|payable|bank)\b/.test(t)) return false;
    return /\b(revenue|income|expenses?)\b|\bcogs\b|cost of (goods|sales)/.test(t);
  }
  const c = String(code == null ? '' : code).trim();
  if (!/^\d+$/.test(c)) return false;
  return Number(c) >= 40000;
}

// ── The unmapped worklist ──
//
// Everything standing between one entity and a fully mapped intercompany
// position, in one list:
//
//   - an account of the four reconcilable kinds, carrying a balance, with no
//     mapping at all; and
//   - a mapping that names the counterparty ENTITY but not that entity's GL
//     ACCOUNT. That is not mapped — two codes and two balances are what make a
//     mapping checkable, and it has one of each — so it is classified as
//     unmapped work here, and Show mapped will not list it.
//
// Each row arrives pre-answered where the data allows: the counterparty its
// name resolves to, and the account on THAT entity's ledger that answers ours
// — an existing one when the ledger holds one, otherwise a DRAFT of the
// account to create. Nothing is created or saved until a person confirms.
//
// Zero balances are excluded throughout: nothing to reconcile, nothing to map.

// What THEIR ledger should hold to answer OUR account of each kind.
const MIRROR_ACCOUNT = {
  due_from:            { ic_type: 'due_to',              type: 'Liability', base: 23000, name: us => 'Due to ' + us },
  due_to:              { ic_type: 'due_from',            type: 'Asset',     base: 18000, name: us => 'Due from ' + us },
  investment:          { ic_type: 'contributed_capital', type: 'Equity',    base: 30100, name: us => 'Contributed Capital - ' + us },
  contributed_capital: { ic_type: 'investment',          type: 'Asset',     base: 19000, name: us => 'Investment in ' + us },
};

// Which of THEIR account kinds can answer OURS. Due accounts answer each
// other in either direction — a due-from carrying a credit is the whole reason
// matching is done by amount — but capital is answered only by investment and
// investment only by capital. Without this, Odyssey's "Due from Banyan
// Residential" was offered as the answer to Banyan Residential's "Contributed
// Capital - Odyssey", purely because the name pointed back at us.
const COMPAT_ANSWER = {
  due_from: ['due_from', 'due_to'],
  due_to: ['due_from', 'due_to'],
  investment: ['contributed_capital'],
  contributed_capital: ['investment'],
};

// Draft the account the counterparty is missing. The name mirrors ours — our
// "Due from X" is answered by "Due to <us>" on X's ledger — and the code
// follows the counterparty's OWN numbering: one past the highest code it
// already uses for that kind, falling back to the portfolio's conventional
// block (18xxx due from, 23xxx due to, 19xxx investment, 30xxx capital) when
// it has none, stepping over collisions either way.
function proposeMirrorAccount(ourEntity, ourIcType, cpAccounts) {
  const mir = MIRROR_ACCOUNT[ourIcType];
  if (!mir) return null;
  const codes = new Set((cpAccounts || []).map(a => String(a.code)));
  let maxNum = -1;
  for (const a of (cpAccounts || [])) {
    if (!/^\d+$/.test(String(a.code))) continue;
    const p = parseAccountName(a.name);
    if (p && p.ic_type === mir.ic_type) maxNum = Math.max(maxNum, Number(a.code));
  }
  let code = maxNum >= 0 ? maxNum + 1 : mir.base;
  while (codes.has(String(code))) code++;
  return { account_code: String(code), account_name: mir.name(ourEntity.name),
    type: mir.type, ic_type: mir.ic_type };
}

function listUnmappedAccounts(db, entityId, { computeBalances, as_of }) {
  ensureSchema(db);
  const eid = Number(entityId);
  // SELECT * rather than naming entity_type: the column is added by migration,
  // and naming a column that does not exist throws instead of degrading.
  const entity = db.prepare('SELECT * FROM entities WHERE id = ?').get(eid);
  if (!entity) { const e = new Error('Entity not found'); e.status = 404; throw e; }
  const shell = isShellEntity(entity);
  const entities = db.prepare('SELECT id, name, code FROM entities').all();
  const entName = new Map(entities.map(e => [Number(e.id), e.name]));
  const companies = listCompanies(db);
  const people = listPeople(db);
  const mapped = new Set(db.prepare('SELECT account_code FROM intercompany_accounts WHERE entity_id = ?')
    .all(eid).map(r => String(r.account_code)));
  const balByCode = new Map(computeBalances(eid, as_of ? { as_of } : {}).map(b => [String(b.code), b]));
  const accounts = db.prepare('SELECT code, name, type FROM accounts WHERE entity_id = ? ORDER BY code').all(eid);

  const out = [];
  let skippedOther = 0, skippedShellCapital = 0, skippedZero = 0;

  // ── Accounts with no mapping at all ──
  for (const a of accounts) {
    if (mapped.has(String(a.code))) continue;
    const parsed = parseAccountName(a.name);
    // Only the four kinds the reconciliation can match belong here. The P&L
    // test stays as a second guard for a chart that names an expense account
    // like a receivable.
    if (!parsed || isPnlAccount(a.type, a.code)) { skippedOther++; continue; }
    if (shell && parsed.ic_type === 'contributed_capital') { skippedShellCapital++; continue; }
    const bal = balByCode.get(String(a.code));
    // A zero balance has nothing for a counterparty ledger to disagree with.
    if (Math.abs(bal ? Number(bal.balance) || 0 : 0) < DEFAULT_TOLERANCE) { skippedZero++; continue; }
    const match = matchCompany(parsed.label, entities, companies);
    out.push({
      source: 'account',
      mapping_id: null,
      entity_id: eid,
      account_code: String(a.code),
      account_name: a.name,
      account_type: a.type || null,
      balance: bal ? round2(bal.balance) : 0,
      individual: isIndividualInvestor(parsed) || isMarkedPerson(people, parsed.label),
      ic_type: parsed.ic_type,
      counterparty_label: parsed.label,
      counterparty_entity_id: match.entity_id,
      counterparty_name: match.entity_id != null ? (entName.get(Number(match.entity_id)) || null) : null,
      counterparty_node_id: match.node_id || null,
      is_external: match.confidence === 'external' ? 1 : 0,
      confidence: match.confidence,
      reason: match.reason,
      can_register: match.entity_id == null && !match.node_id && match.confidence !== 'external',
      notes: null,
    });
  }

  // ── Mappings that name the entity but not its GL account ──
  // Classified as UNMAPPED: naming the entity alone leaves nothing to compare.
  // Show mapped excludes these; this list owns them. External and off-ledger
  // counterparties are neither — no ledger means no account to ever name — and
  // a mapping pointing at the entity itself is a finding, not work.
  for (const m of listMappings(db, { entity_id: eid })) {
    if (m.is_external || m.counterparty_entity_id == null) continue;
    if (Number(m.counterparty_entity_id) === eid) continue;
    if (m.counterparty_account_code) continue;
    const b = balByCode.get(String(m.account_code));
    const balance = b ? round2(b.balance) : 0;
    if (Math.abs(balance) < DEFAULT_TOLERANCE) { skippedZero++; continue; }
    out.push({
      source: 'mapping',
      mapping_id: m.id,
      entity_id: eid,
      account_code: String(m.account_code),
      account_name: m.account_name,
      account_type: b ? b.type : null,
      balance,
      individual: false,
      ic_type: m.ic_type,
      counterparty_label: m.counterparty_name || null,
      counterparty_entity_id: Number(m.counterparty_entity_id),
      counterparty_name: m.counterparty_name || entName.get(Number(m.counterparty_entity_id)) || null,
      counterparty_node_id: null,
      is_external: 0,
      confidence: 'mapped',
      reason: 'the counterparty is chosen — its GL account is not',
      can_register: false,
      notes: m.notes || null,
    });
  }

  // ── External mappings whose party has since gained an uploaded TB ──
  // "External" was the truth when nothing answered for the party. The moment
  // its company is registered and a trial balance is uploaded, there IS a
  // ledger to compare against — so the mapping comes back as work, pre-answered
  // from the TB, and one Confirm re-points it at the TB line. A mapping the
  // user leaves unconfirmed simply stays external.
  for (const m of listMappings(db, { entity_id: eid })) {
    if (!m.is_external) continue;
    const b = balByCode.get(String(m.account_code));
    const balance = b ? round2(b.balance) : 0;
    if (Math.abs(balance) < DEFAULT_TOLERANCE) { skippedZero++; continue; }
    const parsed = parseAccountName(m.account_name || '');
    const label = (parsed && parsed.label) || m.notes || null;
    if (!label) continue;
    const match = matchCompany(label, entities, companies);
    if (match.entity_id != null || !match.node_id) continue; // CL entities never land here; unregistered stays external
    const tb = externalTbFor(db, match.node_id, as_of);
    if (!tb) continue;
    out.push({
      source: 'mapping',
      mapping_id: m.id,
      entity_id: eid,
      account_code: String(m.account_code),
      account_name: m.account_name,
      account_type: b ? b.type : null,
      balance,
      individual: false,
      ic_type: m.ic_type || (parsed && parsed.ic_type) || 'due_from',
      counterparty_label: label,
      counterparty_entity_id: null,
      counterparty_name: null,
      counterparty_node_id: match.node_id,
      is_external: 0,
      confidence: 'tb',
      reason: 'was external — ' + label + ' now has an uploaded TB',
      can_register: false,
      notes: m.notes || null,
    });
  }

  // ── Pre-answer each row from the counterparty's own ledger ──
  // Best EXISTING account first: one whose name resolves back to us, then the
  // closest offsetting balance. A near miss by amount alone is NOT offered —
  // measured against the real ledger that ranking picked operating checking
  // accounts. Only when the ledger holds nothing worth offering does the row
  // carry a draft account to create instead. Balances and charts are read once
  // per counterparty, not once per row.
  const cpBal = new Map(), cpAccts = new Map();
  const balsFor = id => { if (!cpBal.has(id)) cpBal.set(id, computeBalances(id, as_of ? { as_of } : {})); return cpBal.get(id); };
  const acctsFor = id => { if (!cpAccts.has(id)) cpAccts.set(id, db.prepare('SELECT code, name, type FROM accounts WHERE entity_id = ?').all(id)); return cpAccts.get(id); };
  // Answer a row from an UPLOADED trial balance the same way the entity path
  // answers from a ledger: the line whose name points back at us first, then
  // the closest offsetting balance. No drafts — a TB is a statement someone
  // issued, not a chart to add accounts to.
  const bestTbAnswer = (row, tb) => {
    const ourRow = balByCode.get(String(row.account_code));
    const ourSigned = receivableSigned(ourRow);
    const ourBal = ourRow ? (Number(ourRow.balance) || 0) : 0;
    const capital = row.ic_type === 'investment' || row.ic_type === 'contributed_capital';
    const mir = MIRROR_ACCOUNT[row.ic_type];
    let best = null;
    for (const b of tb.lines) {
      if (isPnlAccount(b.type, b.account_code)) continue;
      const p = parseAccountName(b.account_name);
      if (p && COMPAT_ANSWER[row.ic_type] && COMPAT_ANSWER[row.ic_type].indexOf(p.ic_type) === -1) continue;
      let nameMatch = false;
      if (p) {
        const m2 = matchCompany(p.label, entities, companies);
        nameMatch = m2.entity_id != null && Number(m2.entity_id) === eid;
      }
      let gap = null;
      if (capital) {
        if (mir && b.type === mir.type) gap = round2(Math.abs((Number(b.balance) || 0) - ourBal));
        else if (!nameMatch) continue;
      } else {
        const signed = receivableSigned(b);
        if (signed == null) continue;
        gap = ourSigned == null ? null : round2(Math.abs(signed + ourSigned));
      }
      const offsets = gap != null && gap < DEFAULT_TOLERANCE;
      if (!nameMatch && !offsets) continue;
      const cand = { account_code: String(b.account_code), account_name: b.account_name, type: b.type,
        balance: round2(b.balance), gap, offsets, name_match: nameMatch, from_tb: true };
      if (!best) { best = cand; continue; }
      const rc = cand.name_match ? 0 : 1, rb = best.name_match ? 0 : 1;
      const gc = cand.gap == null ? Infinity : cand.gap, gb = best.gap == null ? Infinity : best.gap;
      if (rc < rb || (rc === rb && gc < gb)) best = cand;
    }
    return best;
  };
  for (const row of out) {
    // Contributed capital whose contributor is NOT set up in CloudLedger needs
    // no counterparty account. Per the org charts those contributors are
    // outside investors — Charing Cross, CIG, Marble, Milhaus — and their
    // books are not here, so there is no investment account to match. When the
    // contributor IS a CL entity (Odyssey → Banyan Residential, a link the
    // charts don't even draw), the normal lookup below finds its Investment
    // account and pre-fills it.
    // A company node with an UPLOADED trial balance is answerable after all:
    // pre-answer from its lines exactly as if its ledger were here. Checked
    // before the investor-capital and external rules, so uploading a TB is
    // what upgrades a counterparty from "outside" to "reconciled".
    const nodeTb = (row.counterparty_entity_id == null && row.counterparty_node_id && !row.individual)
      ? externalTbFor(db, row.counterparty_node_id, as_of) : null;
    if (nodeTb) {
      row.tb_as_of = nodeTb.as_of;
      const bestTb = bestTbAnswer(row, nodeTb);
      if (bestTb) row.suggested_existing = bestTb;
      row.reason = 'answered from the uploaded TB as of ' + nodeTb.as_of;
      continue;
    }
    if (row.ic_type === 'contributed_capital' && row.counterparty_entity_id == null && !row.individual) {
      row.investor_capital = true;
      continue;
    }
    if (row.is_external) continue;
    // Any other kind facing a party with no CL entity is EXTERNAL: there is no
    // ledger here to name an account on, so the mapping completes without one.
    // The reconciliation offers a manually entered value to compare against.
    if (row.counterparty_entity_id == null) {
      if (!row.individual) {
        row.is_external = 1;
        row.reason = 'no ledger in CloudLedger — classified external';
      }
      continue;
    }
    const cid = Number(row.counterparty_entity_id);
    const ourRow = balByCode.get(String(row.account_code));
    const ourSigned = receivableSigned(ourRow);
    const ourBal = ourRow ? (Number(ourRow.balance) || 0) : 0;
    const capital = row.ic_type === 'investment' || row.ic_type === 'contributed_capital';
    // Capital and investment agree in EQUAL terms — a parent's investment
    // asset should equal the child's contributed-capital equity — so their gap
    // is a plain difference, not the receivable-signed sum the due accounts use.
    const answerType = row.ic_type === 'investment' ? 'Equity' : 'Asset';
    let best = null;
    for (const b of balsFor(cid)) {
      if (b.bank_acct || isPnlAccount(b.type, b.code)) continue;
      const p = parseAccountName(b.name);
      if (p && COMPAT_ANSWER[row.ic_type] && COMPAT_ANSWER[row.ic_type].indexOf(p.ic_type) === -1) continue;
      let nameMatch = false;
      if (p) {
        const m2 = matchCompany(p.label, entities, companies);
        nameMatch = m2.entity_id != null && Number(m2.entity_id) === eid;
      }
      let gap = null;
      if (capital) {
        if (b.type === answerType) gap = round2(Math.abs((Number(b.balance) || 0) - ourBal));
        else if (!nameMatch) continue;
      } else {
        const signed = receivableSigned(b);
        if (signed == null) continue;
        gap = ourSigned == null ? null : round2(Math.abs(signed + ourSigned));
      }
      const offsets = gap != null && gap < DEFAULT_TOLERANCE;
      if (!nameMatch && !offsets) continue;
      const cand = { account_code: String(b.code), account_name: b.name, type: b.type,
        balance: round2(b.balance), gap, offsets, name_match: nameMatch };
      if (!best) { best = cand; continue; }
      const rc = cand.name_match ? 0 : 1, rb = best.name_match ? 0 : 1;
      const gc = cand.gap == null ? Infinity : cand.gap, gb = best.gap == null ? Infinity : best.gap;
      if (rc < rb || (rc === rb && gc < gb)) best = cand;
    }
    if (best) row.suggested_existing = best;
    else row.suggested_new = proposeMirrorAccount(entity, row.ic_type, acctsFor(cid));
  }

  out.sort((a, b) => (Math.abs(b.balance) - Math.abs(a.balance))
    || a.account_code.localeCompare(b.account_code));

  return {
    entity_id: eid,
    entity: { id: entity.id, name: entity.name, code: entity.code,
      entity_type: entity.entity_type || null },
    shell_entity: shell,
    as_of: as_of || null,
    count: out.length,
    // Reported, not silently dropped — the page says why an account someone
    // went looking for is not in the list.
    skipped_other: skippedOther,
    skipped_shell_capital: skippedShellCapital,
    skipped_zero_balance: skippedZero,
    accounts: out,
  };
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
    counterparty_account_code: m.counterparty_account_code || null,
  };
}

// The counterparty account a person pinned, turned into a leg even though the
// counterparty has no mapping for it.
//
// This is the point of pinning. 26 of the 28 mapped relationships in this
// portfolio are one_sided purely because nobody has set the OTHER entity up —
// the balance is sitting right there in its ledger. Once a person has said
// "their 23000 answers our 18334", waiting for someone to go and mirror the
// mapping adds no information and hides a real difference in the meantime.
//
// ic_type comes from the account TYPE, never the name: a "Due from" account can
// carry a credit, which is why the pin was found by its offsetting balance in
// the first place.
function pinnedLeg(balances, entityId, code, types) {
  const row = (balances.get(Number(entityId)) || new Map()).get(String(code));
  if (!row) return null;
  const amount = Number(row.balance) || 0;
  let ic_type;
  if (row.type === 'Liability') ic_type = 'due_to';
  else if (row.type === 'Equity') ic_type = 'contributed_capital';
  else if (row.type === 'Asset') ic_type = types.includes('investment') ? 'investment' : 'due_from';
  else return null;
  if (!types.includes(ic_type)) return null;
  return {
    mapping_id: null,
    account_code: String(code),
    account_name: row.name,
    ic_type,
    amount,
    signed: ic_type === 'due_to' ? -amount : amount,
    notes: null,
    counterparty_account_code: null,
    // So the page can say this side was asserted here, not mapped over there.
    from_pin: true,
  };
}

function reconcileForEntity(db, entityId, opts) {
  ensureSchema(db);
  const eid = Number(entityId);
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;
  const { computeBalances, as_of } = opts;

  const entity = db.prepare('SELECT * FROM entities WHERE id = ?').get(eid);
  if (!entity) { const e = new Error('Entity not found'); e.status = 404; throw e; }
  const shell = isShellEntity(entity);

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
  // Every counterparty this entity names. A pinned account is read from the
  // counterparty's own ledger, so its balances have to be loaded whether or not
  // that counterparty has any mappings of its own.
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
      // Then any account we pinned that they have not mapped themselves. Deduped
      // by code, so once the counterparty does map it, their own mapping wins
      // and nothing is counted twice.
      const have = new Set(g.their_legs.map(l => String(l.account_code)));
      for (const l of g.our_legs) {
        const code = l.counterparty_account_code;
        if (!code || have.has(String(code))) continue;
        const pin = pinnedLeg(balances, g.counterparty_entity_id, code, types);
        if (pin) { g.their_legs.push(pin); have.add(String(code)); }
      }
    }
    // A counterparty with an UPLOADED trial balance gets a real other side:
    // the TB lines its mappings name become legs, so the pair reconciles like
    // any in-ledger relationship instead of sitting at no_ledger forever. The
    // leg's kind mirrors ours so the investment tab sums it correctly.
    for (const g of groups.values()) {
      if (g.counterparty_entity_id != null || g.counterparty_node_id == null) continue;
      const tb = externalTbFor(db, g.counterparty_node_id, as_of || null);
      if (!tb) continue;
      g.tb_as_of = tb.as_of;
      const have = new Set();
      for (const l of g.our_legs) {
        const code = l.counterparty_account_code;
        if (!code || have.has(String(code))) continue;
        const line = tb.byCode.get(String(code));
        if (!line) continue;
        have.add(String(code));
        const amount = Number(line.balance) || 0;
        const mir = MIRROR_ACCOUNT[l.ic_type];
        const signed = receivableSigned(line);
        g.their_legs.push({ mapping_id: null, account_code: String(line.account_code),
          account_name: line.account_name, ic_type: mir ? mir.ic_type : l.ic_type,
          amount, signed: signed == null ? 0 : signed, notes: null,
          counterparty_account_code: null, from_tb: true });
      }
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
    if ((g.off_ledger || g.counterparty_entity_id == null) && !g.tb_as_of) status = 'no_ledger';
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
    else if ((g.off_ledger || g.counterparty_entity_id == null) && !g.tb_as_of) status = 'no_ledger';
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
  const companies = listCompanies(db);
  // Each unmapped account carries the counterparty its NAME implies, so the
  // reconciliation page can offer to map it where it is found. Sending someone
  // to IC Mapping to retype an account code that is already on the screen is
  // the reason these rows sat unmapped in the first place.
  const unmappedFor = (types) => db.prepare('SELECT code, name, type FROM accounts WHERE entity_id = ?').all(eid)
    .filter(a => {
      if (mappedCodes.has(String(a.code))) return false;
      if (isPnlAccount(a.type, a.code)) return false;
      const p = parseAccountName(a.name);
      if (!p || !types.includes(p.ic_type)) return false;
      // A shell entity's contributed capital is the fund's money passing
      // through, not something a counterparty ledger here states back.
      if (shell && p.ic_type === 'contributed_capital') return false;
      if (isIndividualInvestor(p) || isMarkedPerson(people, p.label)) return false;
      return Math.abs(amountsFor(balances, eid, a.code)) >= tol;
    })
    .map(a => {
      const p = parseAccountName(a.name);
      const match = matchCompany(p.label, entities, companies);
      return {
        account_code: String(a.code), account_name: a.name,
        balance: round2(amountsFor(balances, eid, a.code)),
        ic_type: p.ic_type,
        counterparty_label: p.label,
        counterparty_entity_id: match.entity_id,
        counterparty_node_id: match.node_id || null,
        is_external: match.confidence === 'external' ? 1 : 0,
        confidence: match.confidence,
        reason: match.reason,
        can_register: match.entity_id == null && !match.node_id && match.confidence !== 'external',
      };
    })
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
  const manualFor = (mid) => {
    const r = db.prepare('SELECT balance FROM intercompany_manual_balances WHERE mapping_id = ? AND as_of = ?')
      .get(mid, as_of || '');
    return r ? Number(r.balance) : null;
  };
  const excludedFor = (types) => external
    .filter(m => types.includes(m.ic_type))
    .map(m => legOf(m, balances))
    .filter(l => Math.abs(l.amount) >= tol)
    .map(l => {
      // The counterparty's books are outside CloudLedger, but a person can
      // still state what they say. Positive, in the account's natural terms;
      // the difference is a plain subtraction against our balance.
      const manual = manualFor(l.mapping_id);
      return { mapping_id: l.mapping_id, account_code: l.account_code, account_name: l.account_name,
        ic_type: l.ic_type, amount: l.amount, manual_balance: manual,
        difference: manual == null ? null : round2(l.amount - manual) };
    })
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
  const { db, auth, requireRole, computeBalances, userHasEntityAccess, workpapersDir } = deps;
  ensureSchema(db);

  const who = req => (req.user && (req.user.name || req.user.email)) || null;
  const gate = [auth, requireRole('Admin', 'Accountant')];

  // Pairs confirmed before auto-mirroring existed get their mirrors at boot.
  // Re-running is a no-op: an account already mapped is never touched.
  try {
    const candidates = db.prepare(`SELECT id FROM intercompany_accounts
      WHERE is_external = 0 AND counterparty_entity_id IS NOT NULL AND counterparty_account_code IS NOT NULL`).all();
    let made = 0;
    for (const c of candidates) { if (ensureMirrorMapping(db, c.id, 'auto-mirror')) made++; }
    if (made) console.log('[ic] auto-mirrored ' + made + ' mapping(s) from already-confirmed pairs');
  } catch (e) { console.error('[ic] mirror backfill:', e.message); }

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

  // ── External entity trial balances ──
  // A counterparty with no ledger in CloudLedger can still be reconciled once
  // someone uploads its trial balance. One TB per (company, as-of date);
  // re-uploading the same pair replaces it in full.
  const tbUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

  app.get('/api/intercompany/external-tbs', ...gate, (req, res) => {
    try {
      res.json(db.prepare(`SELECT l.node_id, l.as_of, COUNT(*) AS line_count,
          ROUND(SUM(ABS(l.balance)), 2) AS abs_total,
          MAX(l.uploaded_at) AS uploaded_at, MAX(l.uploaded_by) AS uploaded_by, MAX(l.filename) AS filename,
          n.name AS node_name
        FROM external_tb_lines l JOIN org_nodes n ON n.id = l.node_id
        GROUP BY l.node_id, l.as_of
        ORDER BY n.name COLLATE NOCASE, l.as_of DESC`).all());
    } catch (e) { fail(res, e); }
  });

  app.get('/api/intercompany/external-tb-lines', ...gate, (req, res) => {
    try {
      const node_id = Number(req.query.node_id);
      const as_of = String(req.query.as_of || '');
      if (!node_id || !as_of) return res.status(400).json({ error: 'node_id and as_of are required' });
      res.json(db.prepare(`SELECT account_code, account_name, type, balance FROM external_tb_lines
        WHERE node_id = ? AND as_of = ? ORDER BY account_code`).all(node_id, as_of));
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/intercompany/external-tbs', ...gate, (req, res) => {
    try {
      const node_id = Number(req.query.node_id);
      const as_of = String(req.query.as_of || '');
      if (!node_id || !as_of) return res.status(400).json({ error: 'node_id and as_of are required' });
      const r = db.prepare('DELETE FROM external_tb_lines WHERE node_id = ? AND as_of = ?').run(node_id, as_of);
      res.json({ success: true, deleted: r.changes });
    } catch (e) { fail(res, e); }
  });

  app.post('/api/intercompany/external-tbs', ...gate, tbUpload.single('file'), (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const node_id = Number(req.body.node_id);
      const as_of = String(req.body.as_of || '');
      if (!node_id) return res.status(400).json({ error: 'node_id is required' });
      if (!/^\d{4}-\d{2}-\d{2}$/.test(as_of)) return res.status(400).json({ error: 'as_of must be YYYY-MM-DD' });
      const node = db.prepare('SELECT id, name, entity_id FROM org_nodes WHERE id = ?').get(node_id);
      if (!node) return res.status(404).json({ error: 'Company not found — register it first' });
      if (node.entity_id) return res.status(400).json({ error: node.name + ' is set up as a CloudLedger entity — its ledger is already here, no TB upload needed.' });

      // Same column detection as the entity TB import: code + name columns,
      // then either debit/credit columns or one signed amount column.
      const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!raw.length) return res.status(400).json({ error: 'No data rows found in file' });
      const cols = Object.keys(raw[0]);
      const norm = c => String(c).toLowerCase().trim();
      const findCol = (patterns, exclude = []) => {
        const pool = cols.filter(c => !exclude.includes(c));
        for (const pat of patterns) { const hit = pool.find(c => norm(c) === pat); if (hit) return hit; }
        for (const pat of patterns) { const hit = pool.find(c => norm(c).includes(pat)); if (hit) return hit; }
        return null;
      };
      const codeCol = findCol(['account number','account #','account code','acct number','acct code','acct','code','number']);
      const nameCol = findCol(['account name','account description','acct name','description','name'], [codeCol]);
      // Project managers export from different systems, so the ENDING balance
      // hides under different names — and an activity-format TB (Beginning /
      // Debit / Credit / Ending) must read the Ending column, not the period
      // activity. Beginning/opening/prior/budget columns are never candidates.
      const notEnding = cols.filter(c => /beginn|opening|prior|budget|activity|change/i.test(String(c)));
      const amtCol  = findCol(['ending balance','end balance','closing balance','current balance','balance','net amount','amount','total'],
        [codeCol, nameCol].filter(Boolean).concat(notEnding));
      const drCol   = findCol(['debit'], [codeCol, nameCol, amtCol].filter(Boolean).concat(notEnding));
      const crCol   = findCol(['credit'], [codeCol, nameCol, amtCol].filter(Boolean).concat(notEnding));
      // An explicit ending/closing column outranks debit/credit columns, which
      // in an activity-format export are the month's movement, not the balance.
      const preferAmt = !!(amtCol && /end|closing|current/i.test(String(amtCol)));
      if (!codeCol) return res.status(400).json({ error: 'Could not find an account number/code column. Found: ' + cols.join(', ') });
      if (!nameCol) return res.status(400).json({ error: 'Could not find an account name column. Found: ' + cols.join(', ') });
      if (!amtCol && !drCol && !crCol) return res.status(400).json({ error: 'Could not find amount or debit/credit columns. Found: ' + cols.join(', ') });

      const typeFromCode = (codeStr) => {
        const n = parseInt(String(codeStr).replace(/[^0-9]/g, ''), 10);
        if (isNaN(n)) return null;
        if (n <= 19999) return 'Asset';
        if (n <= 29999) return 'Liability';
        if (n <= 39999) return 'Equity';
        if (n <= 49999) return 'Revenue';
        if (n <= 69999) return 'Expense';
        return 'Revenue';
      };
      const parsed = [];
      for (const row of raw) {
        const code = String(row[codeCol] || '').trim();
        const name = String(row[nameCol] || '').trim();
        if (!code || !name) continue;
        const type = typeFromCode(code);
        if (!type) continue;
        let dr = 0, cr = 0, amt = null;
        if ((drCol || crCol) && !preferAmt) {
          dr = parseFloat(String(row[drCol] || '0').replace(/[,$()]/g, '')) || 0;
          cr = parseFloat(String(row[crCol] || '0').replace(/[,$()]/g, '')) || 0;
        } else {
          const rawAmt = String(row[amtCol] || '').trim();
          const isParen = /^\(.*\)$/.test(rawAmt);
          let v = parseFloat(rawAmt.replace(/[,$()]/g, '')) || 0;
          if (isParen) v = -v;
          amt = v;
        }
        parsed.push({ code, name, type, dr, cr, amt });
      }
      if (!parsed.length) return res.status(400).json({ error: 'No valid rows found. Check that account codes are numeric.' });

      // Single-amount files: debit-positive when the column sums to zero,
      // natural-side otherwise (positive = the account's normal side).
      let signMode = 'debit-positive';
      if (parsed.some(p => p.amt !== null)) {
        const sumSigned = parsed.reduce((s, p) => s + (p.amt || 0), 0);
        signMode = Math.abs(sumSigned) < 0.01 ? 'debit-positive' : 'natural';
      }
      const isDrNatural = t => t === 'Asset' || t === 'Expense';
      const byCode = new Map();
      for (const p of parsed) {
        let natural;
        if (p.amt === null) natural = isDrNatural(p.type) ? (p.dr - p.cr) : (p.cr - p.dr);
        else if (signMode === 'debit-positive') natural = isDrNatural(p.type) ? p.amt : -p.amt;
        else natural = p.amt;
        const cur = byCode.get(p.code);
        if (cur) cur.balance += natural;
        else byCode.set(p.code, { code: p.code, name: p.name, type: p.type, balance: natural });
      }
      const lines = [...byCode.values()];

      // A TB should balance. In debit-positive terms the whole file sums to
      // zero; a residual usually means a subtotal row was read or the ending
      // column was misdetected — said out loud, never swallowed.
      const isDr2 = t => t === 'Asset' || t === 'Expense';
      const residual = round2(lines.reduce((s, l) => s + (isDr2(l.type) ? l.balance : -l.balance), 0));

      const now = new Date().toISOString();
      const tx = db.transaction(() => {
        db.prepare('DELETE FROM external_tb_lines WHERE node_id = ? AND as_of = ?').run(node_id, as_of);
        const ins = db.prepare(`INSERT INTO external_tb_lines
          (node_id, as_of, account_code, account_name, type, balance, uploaded_at, uploaded_by, filename)
          VALUES (?,?,?,?,?,?,?,?,?)`);
        for (const l of lines) ins.run(node_id, as_of, l.code, l.name, l.type, round2(l.balance), now, who(req), req.file.originalname || null);
      });
      tx();

      // File the ORIGINAL upload in Shell Entities Accounting's workpapers so
      // the source document sits next to what was parsed from it. Folder per
      // as-of date, company + date in the file name; re-upload replaces.
      let workpaper = null;
      try {
        const shell = db.prepare(`SELECT id FROM entities WHERE name LIKE '%Shell Entities Accounting%' ORDER BY id LIMIT 1`).get();
        if (shell && workpapersDir) {
          const folder = 'External TBs/' + as_of;
          const by = who(req) || 'system';
          for (const fp of ['External TBs', folder]) {
            try { db.prepare('INSERT INTO entity_folders (entity_id, folder_path, created_by) VALUES (?,?,?)').run(shell.id, fp, by); }
            catch (e) { if (!/UNIQUE/i.test(e.message)) throw e; }
          }
          const ext = (String(req.file.originalname || '').match(/\.[A-Za-z0-9]+$/) || ['.xlsx'])[0];
          const originalName = as_of + ' ' + node.name + ' TB' + ext;
          const dir = path.join(workpapersDir, String(shell.id));
          fs.mkdirSync(dir, { recursive: true });
          // Replace an earlier copy of the same TB rather than piling up versions.
          for (const old of db.prepare('SELECT id, stored_filename FROM entity_files WHERE entity_id = ? AND folder_path = ? AND original_name = ?')
            .all(shell.id, folder, originalName)) {
            try { fs.unlinkSync(path.join(dir, old.stored_filename)); } catch (e) {}
            db.prepare('DELETE FROM entity_files WHERE id = ?').run(old.id);
          }
          const stored = Date.now() + '_' + Math.floor(Math.random() * 1e6) + '_' + originalName.replace(/[^a-zA-Z0-9._-]/g, '_');
          fs.writeFileSync(path.join(dir, stored), req.file.buffer);
          db.prepare('INSERT INTO entity_files (entity_id, folder_path, stored_filename, original_name, size, mime_type, uploaded_by) VALUES (?,?,?,?,?,?,?)')
            .run(shell.id, folder, stored, originalName, req.file.size, req.file.mimetype || null, by);
          workpaper = { entity_id: shell.id, folder_path: folder, file_name: originalName };
        }
      } catch (e) { console.error('[ic] TB workpaper filing failed:', e.message); }

      res.json({ success: true, node_id, node_name: node.name, as_of, count: lines.length,
        sign_mode: signMode, residual, balanced: Math.abs(residual) < 0.02, workpaper });
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

  // Every account on the entity with no mapping yet, recognised by the name
  // parser or not. The suggestion endpoint can only propose what it recognises;
  // this is what you open when the account you are after is not in that list.
  // The finished side of the same job: mappings that face another entity, with
  // both sides' accounts and balances on one row.
  app.get('/api/intercompany/accounts/mapped', ...gate, (req, res) => {
    try {
      const entity_id = Number(req.query.entity_id);
      if (!entity_id) return res.status(400).json({ error: 'entity_id is required' });
      assertEntityAccess(req, [entity_id]);
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      res.json(listMappedPairs(db, entity_id, { computeBalances, as_of }));
    } catch (e) { fail(res, e); }
  });

  app.get('/api/intercompany/accounts/unmapped', ...gate, (req, res) => {
    try {
      const entity_id = Number(req.query.entity_id);
      if (!entity_id) return res.status(400).json({ error: 'entity_id is required' });
      assertEntityAccess(req, [entity_id]);
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      res.json(listUnmappedAccounts(db, entity_id, { computeBalances, as_of }));
    } catch (e) { fail(res, e); }
  });

  // Walk the counterparty's ledger for the account that answers ours. Ranked,
  // never decided: the response says which one it would pick and by how much
  // the two sides disagree, and a person confirms.
  app.get('/api/intercompany/counterparty-accounts', ...gate, (req, res) => {
    try {
      const entity_id = Number(req.query.entity_id);
      const counterparty_entity_id = Number(req.query.counterparty_entity_id);
      const account_code = String(req.query.account_code || '');
      if (!entity_id || !counterparty_entity_id || !account_code) {
        return res.status(400).json({ error: 'entity_id, counterparty_entity_id and account_code are required' });
      }
      assertEntityAccess(req, [entity_id]);
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      res.json(counterpartyCandidates(db, { entityId: entity_id, accountCode: account_code,
        counterpartyEntityId: counterparty_entity_id, computeBalances, as_of }));
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
      // The other half of each pair is written for the counterparty, so its
      // worklist never asks for a confirmation this person just gave.
      let mirrored = 0;
      for (const id of ids) { try { if (ensureMirrorMapping(db, id, who(req))) mirrored++; } catch (e) { console.error('[ic] mirror:', e.message); } }
      res.json({ ids, count: ids.length, mirrored, success: true });
    } catch (e) {
      if (/UNIQUE/i.test(e.message)) return res.status(400).json({ error: 'That account is already mapped for this entity' });
      fail(res, e);
    }
  });

  // Manual counterparty value for an EXTERNAL mapping, keyed by as-of date.
  // A null / blank balance clears the entry.
  app.put('/api/intercompany/manual-balance', ...gate, (req, res) => {
    try {
      const { mapping_id, as_of, balance } = req.body || {};
      const m = db.prepare('SELECT * FROM intercompany_accounts WHERE id = ?').get(Number(mapping_id));
      if (!m) return res.status(404).json({ error: 'Mapping not found' });
      assertEntityAccess(req, [m.entity_id]);
      const key = /^\d{4}-\d{2}-\d{2}$/.test(String(as_of || '')) ? String(as_of) : '';
      if (balance == null || balance === '') {
        db.prepare('DELETE FROM intercompany_manual_balances WHERE mapping_id = ? AND as_of = ?').run(m.id, key);
        return res.json({ success: true, cleared: true });
      }
      const val = Number(balance);
      if (!Number.isFinite(val)) return res.status(400).json({ error: 'balance must be a number' });
      const now = new Date().toISOString();
      db.prepare(`INSERT INTO intercompany_manual_balances (mapping_id, as_of, balance, updated_at, updated_by)
        VALUES (?,?,?,?,?)
        ON CONFLICT(mapping_id, as_of) DO UPDATE SET balance=excluded.balance, updated_at=excluded.updated_at, updated_by=excluded.updated_by`)
        .run(m.id, key, val, now, who(req));
      res.json({ success: true, balance: val });
    } catch (e) { fail(res, e); }
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
      let mirrored = 0;
      try { if (ensureMirrorMapping(db, Number(req.params.id), who(req))) mirrored++; } catch (e) { console.error('[ic] mirror:', e.message); }
      res.json({ success: true, mirrored });
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
  listUnmappedAccounts,
  listMappedPairs,
  listMappings,
  listCompanies,
  listPeople,
  isMarkedPerson,
  isPnlAccount,
  isShellEntity,
  receivableSigned,
  counterpartyCandidates,
  resolveCounterparties,
  matchCompany,
  looksIndividual,
  isIndividualInvestor,
  parseAccountName,
  matchEntity,
  normName,
  IC_TYPES,
};
