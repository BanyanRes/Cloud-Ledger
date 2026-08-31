// ═══════════════════════════════════════════════════════════════════════════
// budget.js — operating budgets and the Budget-to-Actual (B2A) schedule.
//
// Jimmy, 2026-08-28: rail assets carry an annual operations budget workbook
// (CLIP / Silsbee / Buna / SRN, all built off one template). The workbook is
// uploaded ONCE PER YEAR from the Financial Statements page — and again
// whenever the budget is revised — and CL then produces the Budget-to-Actual
// schedule every month with no further input, appending it as the LAST item
// in the statement package.
//
// The design decision that matters: the workbook is parsed ONCE, at upload,
// into budget_lines. The monthly report reads the DATABASE, never the .xlsx.
// Re-reading the file each close would make the report hostage to anybody who
// inserts a row in the spreadsheet. The uploaded file is still filed in the
// entity's Workpapers tree as the source document.
//
// Mapping budget line -> GL account is likewise stored (budget_account_map),
// not re-derived. Several mappings are genuine judgment calls (PPE, Rent
// Expense, Interchange Fees) and must not be silently re-guessed month to
// month.
// ═══════════════════════════════════════════════════════════════════════════
const XLSX = require('xlsx');

const BUDGET_FOLDER = 'Budget';

// ── Schema ─────────────────────────────────────────────────────────────────
function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS budget_versions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      fiscal_year INTEGER NOT NULL,
      version_no INTEGER NOT NULL,
      label TEXT,
      original_name TEXT,
      stored_filename TEXT,
      sheet_name TEXT,
      uploaded_by TEXT,
      uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
      is_active INTEGER NOT NULL DEFAULT 1,
      note TEXT
    );
    CREATE INDEX IF NOT EXISTS ix_budget_versions_ent
      ON budget_versions(entity_id, fiscal_year, is_active);

    CREATE TABLE IF NOT EXISTS budget_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      version_id INTEGER NOT NULL,
      seq INTEGER NOT NULL,
      kind TEXT NOT NULL,
      section TEXT,
      group_name TEXT,
      label TEXT NOT NULL,
      note TEXT,
      m1 REAL DEFAULT 0, m2 REAL DEFAULT 0, m3 REAL DEFAULT 0, m4 REAL DEFAULT 0,
      m5 REAL DEFAULT 0, m6 REAL DEFAULT 0, m7 REAL DEFAULT 0, m8 REAL DEFAULT 0,
      m9 REAL DEFAULT 0, m10 REAL DEFAULT 0, m11 REAL DEFAULT 0, m12 REAL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ix_budget_lines_ver ON budget_lines(version_id, seq);

    CREATE TABLE IF NOT EXISTS budget_account_map (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entity_id INTEGER NOT NULL,
      label_norm TEXT NOT NULL,
      label TEXT NOT NULL,
      account_code TEXT NOT NULL,
      source TEXT,
      created_by TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(entity_id, label_norm, account_code)
    );
    CREATE INDEX IF NOT EXISTS ix_budget_map_ent ON budget_account_map(entity_id, label_norm);
  `);
}

// ── Label normalisation ────────────────────────────────────────────────────
// Budget labels are hand-typed and indented with leading spaces; account names
// vary in punctuation ("Legal Fees - General" vs "Legal Fees – General").
// Normalising collapses whitespace, unifies dashes and ampersands, and drops
// trailing punctuation so the two sides meet.
function norm(s) {
  return String(s == null ? '' : s)
    .replace(/[‐-―]/g, '-')       // en/em dashes -> hyphen
    .replace(/’/g, "'")
    .replace(/&amp;/gi, '&')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[.:;,]+$/, '');
}

const GROUP_HEADINGS = new Set([
  'general and administrative expenses',
  'payroll expenses',
  'utilities and facilities',
  'management fees',
  'taxes and insurance',
]);

const SECTION_HEADINGS = new Map([
  ['revenue', 'Revenue'],
  ['operating expenses', 'Operating Expenses'],
]);

// ── Parser ─────────────────────────────────────────────────────────────────
// Structural, not positional. The four rail workbooks put their line items in
// different ROWS but always in column D, under the same section headings, with
// the twelve month-end dates on one header row starting at column F. Keying off
// the labels means a budget that gains or loses a line still parses.
function parseWorkbook(buffer, opts = {}) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = opts.sheet
    || wb.SheetNames.find(n => /budget\s*detail/i.test(n))
    || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error('Workbook has no readable sheet');

  // Read cells by ABSOLUTE address, never by position in a sheet_to_json row.
  // These sheets start at column C, and sheet_to_json indexes from the range
  // origin rather than column A — so row[5] is column H, not F. Addressing the
  // cells directly removes that whole class of off-by-N bug.
  const COL = { label: 3, first: 5, last: 16, notes: 18 };   // D, F..Q, S
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const at = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell == null ? null : cell.v;
  };

  // Excel dates arrive as UTC instants offset from local midnight, so read the
  // calendar parts rather than toISOString(), which can slide a day either way.
  const isDate = (v) => v instanceof Date && !isNaN(v.getTime());
  const ymd = (d) => d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0')
    + '-' + String(d.getUTCDate()).padStart(2, '0');

  // Header row: the first row carrying a date in column F.
  let hdr = -1;
  const scanTo = Math.min(range.e.r, 24);
  for (let r = range.s.r; r <= scanTo; r++) {
    if (isDate(at(r, COL.first))) { hdr = r; break; }
  }
  if (hdr < 0) {
    throw new Error('Could not find the month header row (expected twelve month-end dates starting in column F of the Budget Detail tab).');
  }
  const periods = [];
  for (let c = COL.first; c <= COL.last; c++) {
    const v = at(hdr, c);
    periods.push(isDate(v) ? ymd(v) : null);
  }
  const firstDate = periods.find(Boolean);
  if (!firstDate) throw new Error('The month header row has no valid dates.');
  const fiscalYear = Number(firstDate.slice(0, 4));
  const monthsFound = periods.filter(Boolean).length;
  if (monthsFound < 12) {
    throw new Error('Expected twelve monthly columns (F:Q) on the header row; found ' + monthsFound + '.');
  }

  const num = (v) => {
    if (v == null || v === '') return 0;
    const n = typeof v === 'number' ? v : Number(String(v).replace(/[$,\s]/g, '').replace(/^\((.*)\)$/, '-$1'));
    return isFinite(n) ? n : 0;
  };

  // A group heading is a row that a later "Total <heading>" row closes. This is
  // the only reliable signal: an all-zero row is NOT a heading — CLIP budgets
  // "Healthcare - Reimbursable" at nil and Buna budgets "Land Leases" at nil,
  // and both are real lines that must keep their place in the schedule.
  const closedByTotal = new Set();
  for (let r = hdr + 1; r <= range.e.r; r++) {
    const v = at(r, COL.label);
    if (v == null) continue;
    const n = norm(v);
    if (/^total /.test(n)) closedByTotal.add(n.replace(/^total /, ''));
  }

  const out = [];
  const warnings = [];
  let section = null;
  let group = null;
  let seq = 0;

  for (let r = hdr + 1; r <= range.e.r; r++) {
    const raw = at(r, COL.label);                        // column D
    const amounts = [];
    for (let c = COL.first; c <= COL.last; c++) amounts.push(num(at(r, c)));
    const hasAmounts = amounts.some(v => Math.abs(v) > 0.0049);
    const rawNote = at(r, COL.notes);
    const note = rawNote == null ? null : (String(rawNote).trim() || null);

    // A row with money on it but NO caption is a budget line whose label was
    // deleted, not an empty row. Buna's 2026 workbook does exactly this: row 47
    // carries $2,000 of December bonuses with an empty column D, and dropping it
    // put the schedule $2,000 below the workbook's own payroll total. Keep it,
    // name it from its note if it has one, and warn.
    if (raw == null || String(raw).trim() === '') {
      if (!hasAmounts) continue;
      const guess = note ? note.split(/[\r\n]/)[0].trim().slice(0, 40) : '';
      const label = '(unlabelled line' + (guess ? ' — ' + guess : '') + ')';
      warnings.push('Row ' + (r + 1) + ' has budget amounts but no line name in column D'
        + (guess ? ' (note: "' + guess + '")' : '') + '. Included as "' + label + '" — add a caption in the workbook and re-upload.');
      out.push({ seq: seq++, kind: 'line', section, group_name: group, label, note, amounts, unlabelled: true });
      continue;
    }
    const label = String(raw).trim();
    const n = norm(label);
    if (n === 'line item') continue;

    // "(Other)" placeholder rows exist so users can add lines; drop the empty ones.
    if (/^\(other\)$/.test(n) && !hasAmounts) continue;

    let kind;
    if (SECTION_HEADINGS.has(n)) {
      section = SECTION_HEADINGS.get(n);
      group = null;
      kind = 'section';
    } else if (n === 'total revenue') {
      kind = 'total';
      group = null;
    } else if (n === 'total operating expenses') {
      kind = 'total';
      group = null;
    } else if (n === 'net operating income') {
      kind = 'noi'; group = null;
    } else if (n === 'projected debt service') {
      kind = 'debt'; group = null;
    } else if (n === 'cash flow after debt service') {
      kind = 'cashflow'; group = null;
    } else if (/^total /.test(n)) {
      kind = 'subtotal';
      group = null;
    } else if (closedByTotal.has(n) || GROUP_HEADINGS.has(n)) {
      // Closed by its own "Total <heading>" row further down (or a heading the
      // template always uses). Amounts play no part in this test.
      group = label.trim();
      kind = 'group';
    } else {
      kind = 'line';
    }

    out.push({
      seq: seq++, kind, section, group_name: kind === 'group' ? null : group,
      label: label.replace(/\s+/g, ' ').trim(), note, amounts,
    });
  }

  const lineCount = out.filter(r => r.kind === 'line').length;
  if (!lineCount) throw new Error('No budget line items were found on the "' + sheetName + '" tab.');

  // ── Self-check: re-add the parsed lines and compare against the workbook's
  // own subtotals and totals, for all twelve months. If the two disagree the
  // parse has dropped or double-counted a row, and the user needs to know at
  // upload time rather than discovering it in a variance three months later.
  {
    const near = (a, b) => Math.abs(a - b) < 0.5;
    const grp = {};
    let cur = null;
    const revSum = Array(12).fill(0), expSum = Array(12).fill(0);
    for (const r of out) {
      if (r.kind === 'group') { cur = norm(r.label); grp[cur] = Array(12).fill(0); continue; }
      if (r.kind === 'line') {
        const target = r.section === 'Revenue' ? revSum : expSum;
        for (let i = 0; i < 12; i++) target[i] += r.amounts[i];
        if (cur && r.section === 'Operating Expenses') for (let i = 0; i < 12; i++) grp[cur][i] += r.amounts[i];
        continue;
      }
      if (r.kind === 'subtotal') {
        const key = norm(r.label).replace(/^total /, '');
        if (grp[key]) {
          for (let i = 0; i < 12; i++) {
            if (!near(grp[key][i], r.amounts[i])) {
              warnings.push(r.label + ' for month ' + (i + 1) + ' is ' + r.amounts[i].toFixed(2)
                + ' in the workbook but the line items add to ' + grp[key][i].toFixed(2) + '.');
            }
          }
        }
        cur = null;
        continue;
      }
      if (r.kind === 'total') {
        const src = /revenue/i.test(r.label) ? revSum : expSum;
        for (let i = 0; i < 12; i++) {
          if (!near(src[i], r.amounts[i])) {
            warnings.push(r.label + ' for month ' + (i + 1) + ' is ' + r.amounts[i].toFixed(2)
              + ' in the workbook but the line items add to ' + src[i].toFixed(2) + '.');
          }
        }
      }
    }
  }

  return { sheetName, fiscalYear, periods, rows: out, warnings };
}

// ── Storing a parsed budget ────────────────────────────────────────────────
function saveVersion(db, { entityId, parsed, originalName, storedFilename, who, label, note }) {
  ensureSchema(db);
  const prev = db.prepare(
    'SELECT MAX(version_no) AS v FROM budget_versions WHERE entity_id=? AND fiscal_year=?'
  ).get(entityId, parsed.fiscalYear);
  const versionNo = (prev && prev.v ? prev.v : 0) + 1;

  const tx = db.transaction(() => {
    // Only one active version per entity + year; older ones are retained so a
    // prior month regenerated later still shows the budget then in force.
    db.prepare('UPDATE budget_versions SET is_active=0 WHERE entity_id=? AND fiscal_year=?')
      .run(entityId, parsed.fiscalYear);
    const ins = db.prepare(`
      INSERT INTO budget_versions
        (entity_id, fiscal_year, version_no, label, original_name, stored_filename, sheet_name, uploaded_by, is_active, note)
      VALUES (?,?,?,?,?,?,?,?,1,?)`);
    const r = ins.run(entityId, parsed.fiscalYear, versionNo,
      label || ('v' + versionNo), originalName || null, storedFilename || null,
      parsed.sheetName || null, who || null, note || null);
    const vid = r.lastInsertRowid;
    const insL = db.prepare(`
      INSERT INTO budget_lines
        (version_id, seq, kind, section, group_name, label, note,
         m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`);
    for (const row of parsed.rows) {
      insL.run(vid, row.seq, row.kind, row.section, row.group_name, row.label, row.note, ...row.amounts);
    }
    return vid;
  });
  const versionId = tx();
  return { versionId, versionNo, fiscalYear: parsed.fiscalYear };
}

function activeVersion(db, entityId, fiscalYear) {
  ensureSchema(db);
  return db.prepare(
    'SELECT * FROM budget_versions WHERE entity_id=? AND fiscal_year=? AND is_active=1 ORDER BY version_no DESC LIMIT 1'
  ).get(entityId, fiscalYear) || null;
}

function listVersions(db, entityId) {
  ensureSchema(db);
  return db.prepare(
    `SELECT v.*, (SELECT COUNT(*) FROM budget_lines l WHERE l.version_id=v.id AND l.kind='line') AS line_count
       FROM budget_versions v WHERE entity_id=? ORDER BY fiscal_year DESC, version_no DESC`
  ).all(entityId);
}

function versionLines(db, versionId) {
  return db.prepare('SELECT * FROM budget_lines WHERE version_id=? ORDER BY seq').all(versionId);
}

// ── Default budget-label -> GL account mapping ─────────────────────────────
// Applied at upload time and stored. Anything here is only a STARTING POINT:
// the stored map is what the report uses, and it is editable.
//
// A label may map to more than one account (SRN's "Payroll Tax" covers both the
// FICA-equivalent RRB tax and payroll taxes proper). Codes that do not exist in
// an entity's chart are skipped, so one table serves all four railroads.
const DEFAULT_MAP = {
  // Revenue
  'land leases': ['40110'],
  'track leases': ['41136'],
  'in & out fees': ['40120'],
  'container fees': ['40125'],
  'storage fees': ['40130'],
  'transloading': ['40140'],
  'intra-plant switch fees': ['40160'],
  'scale fees': ['41135'],
  'interchange fees': ['40135', '40136'],   // BNSF + UP (SRN) — confirm
  'warehouse lease income': ['40150'],
  // General & administrative
  'tax & license': ['68000'],
  'accounting': ['63000'],
  'professional fees': ['63025'],
  'legal fees - general': ['63050'],
  'travel': ['60210'],
  'meals': ['60500'],
  'fuel': ['61164'],
  'dues & subscriptions': ['67100'],
  'office expense': ['67200'],
  'telephone & internet': ['67300'],
  'advertising & marketing': ['67400'],
  'bank fees': ['68110'],
  // Payroll
  'transload salaries - reimbursable': ['63030'],
  'healthcare - reimbursable': ['63032'],
  'switchmen healthcare - reimbursable': ['63032'],
  'switchmen salaries - reimbursable': ['63034'],
  'offsite staff': ['63042'],
  'salaries and wages': ['60000'],
  'health insurance': ['60005'],
  'payroll tax': ['60002', '60012'],        // payroll taxes + RRB employer — confirm
  // Utilities & facilities
  'site/yard maintenance': ['61050'],
  'crossing repairs': ['61052'],
  'locomotive repair': ['61054'],
  'landscape maintenance': ['61056'],
  'rent expense': ['61000'],                // Locomotive Rent — confirm
  'track lease': ['61010'],
  'utilities': ['61150'],
  'security': ['64000', '61100'],
  'ppe': ['61053'],                         // Equipment Supplies — confirm
  // Management fees
  'sponsor management fee': ['63027'],
  'management fees - other': ['63038'],
  'clro management fees': ['63041'],
  // Taxes & insurance
  'insurance': ['68055'],
  'taxes': ['68050'],
  // Debt service. The budget's Projected Debt Service is interest only - its
  // Debt Service tab computes -(ending balance x rate / 12) and the loan
  // balance never amortises - so it compares directly to interest expense.
  // Loan and unused-fee accounts (75005 / 75006) are deliberately NOT
  // included: they are financing costs but not interest, and stay in Other
  // Income (Expense) so this comparison is like for like.
  'projected debt service': ['75000'],
};

// Labels we deliberately do not guess at. Presented as unmapped so somebody
// decides, rather than quietly landing on a plausible-looking account.
const NO_DEFAULT = new Set(['bonuses', 'cam income']);

// Seed the map for an entity from the defaults + exact name matches against its
// own chart. Never overwrites a mapping that already exists (a human may have
// corrected it); returns what it added and what it could not place.
function seedMap(db, entityId, labels, accounts, who) {
  ensureSchema(db);
  const byCode = new Map(accounts.map(a => [String(a.code), a]));
  const byName = new Map();
  for (const a of accounts) {
    const k = norm(a.name);
    if (!byName.has(k)) byName.set(k, a);
  }
  const existing = new Set(
    db.prepare('SELECT label_norm FROM budget_account_map WHERE entity_id=?').all(entityId).map(r => r.label_norm)
  );
  const ins = db.prepare(
    'INSERT OR IGNORE INTO budget_account_map (entity_id, label_norm, label, account_code, source, created_by) VALUES (?,?,?,?,?,?)'
  );
  const added = [];
  const unmapped = [];
  for (const label of labels) {
    const n = norm(label);
    if (existing.has(n)) continue;
    if (NO_DEFAULT.has(n)) { unmapped.push(label); continue; }
    let codes = [];
    let source = null;
    // 1. An account whose NAME is the budget label is the strongest signal.
    if (byName.has(n)) { codes = [String(byName.get(n).code)]; source = 'name'; }
    // 2. Otherwise the curated default table, filtered to codes this entity has.
    if (!codes.length && DEFAULT_MAP[n]) {
      codes = DEFAULT_MAP[n].filter(c => byCode.has(c));
      source = 'default';
    }
    if (!codes.length) { unmapped.push(label); continue; }
    for (const c of codes) ins.run(entityId, n, label, c, source, who || null);
    added.push({ label, codes, source });
  }
  return { added, unmapped };
}

function getMap(db, entityId) {
  ensureSchema(db);
  const rows = db.prepare(
    'SELECT label_norm, label, account_code, source FROM budget_account_map WHERE entity_id=? ORDER BY label, account_code'
  ).all(entityId);
  const m = new Map();
  for (const r of rows) {
    if (!m.has(r.label_norm)) m.set(r.label_norm, []);
    m.get(r.label_norm).push(r.account_code);
  }
  return m;
}

function setMap(db, entityId, label, codes, who) {
  ensureSchema(db);
  const n = norm(label);
  const tx = db.transaction(() => {
    db.prepare('DELETE FROM budget_account_map WHERE entity_id=? AND label_norm=?').run(entityId, n);
    const ins = db.prepare(
      'INSERT OR IGNORE INTO budget_account_map (entity_id, label_norm, label, account_code, source, created_by) VALUES (?,?,?,?,?,?)'
    );
    for (const c of (codes || [])) ins.run(entityId, n, label, String(c).trim(), 'manual', who || null);
  });
  tx();
  return { label, codes: codes || [] };
}

// ── Budget-to-Actual builder ───────────────────────────────────────────────
// Pure given `balancesAt(dateStr) -> [{code,name,type,balance}]`.
//
// Actual activity for a P&L account is the CHANGE in its balance, because CL
// stores Revenue/Expense balances inception-to-date (see the CLRFI tie-out
// note, 2026-08-17):
//     month = balance(asOf) − balance(prior month end)
//     ytd   = balance(asOf) − balance(31 Dec prior year)
//
// Variance is presented favourable-positive: revenue actual − budget, expense
// budget − actual. An accountant reading the column never has to work out the
// sign convention per section.
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const OTHER_CODE = /^(69|70|71|75|82|83)/;   // below the operating line
// Interest capitalised into construction in progress rather than expensed.
// Matched on account NAME, because the code differs by entity (CLIP/Silsbee/
// Buna use 12321, SRN 12325) and these are Asset accounts, so they can never
// collide with a P&L line.
const CAPITALISED_INTEREST = /construction\s+period\s+interest/i;

function priorMonthEnd(asOf) {
  const [y, m] = String(asOf).split('-').map(Number);
  const d = new Date(Date.UTC(y, m - 1, 1));
  d.setUTCDate(0);
  return d.toISOString().slice(0, 10);
}
function priorYearEnd(asOf) { return (Number(String(asOf).slice(0, 4)) - 1) + '-12-31'; }

async function buildBudgetToActual(db, { entityId, asOf, entityName, balancesAt }) {
  ensureSchema(db);
  const fiscalYear = Number(String(asOf).slice(0, 4));
  const monthIdx = Number(String(asOf).slice(5, 7));      // 1..12
  const version = activeVersion(db, entityId, fiscalYear);
  if (!version) return null;                              // no budget on file — caller skips the schedule

  const lines = versionLines(db, version.id);
  const map = getMap(db, entityId);

  const [cur, pm, py] = await Promise.all([
    balancesAt(asOf), balancesAt(priorMonthEnd(asOf)), balancesAt(priorYearEnd(asOf)),
  ]);
  const idx = (arr) => {
    const m = new Map();
    for (const a of arr || []) m.set(String(a.code), a);
    return m;
  };
  const iCur = idx(cur), iPm = idx(pm), iPy = idx(py);
  const meta = new Map();
  for (const a of [].concat(cur || [], pm || [], py || [])) {
    if (!meta.has(String(a.code))) meta.set(String(a.code), a);
  }
  const bal = (m, code) => {
    const a = m.get(String(code));
    return a ? Number(a.balance) || 0 : 0;
  };
  const actOf = (code) => ({
    month: r2(bal(iCur, code) - bal(iPm, code)),
    ytd: r2(bal(iCur, code) - bal(iPy, code)),
  });

  const budOf = (row) => {
    let month = 0, ytd = 0;
    for (let i = 1; i <= 12; i++) {
      const v = Number(row['m' + i]) || 0;
      if (i === monthIdx) month = v;
      if (i <= monthIdx) ytd += v;
    }
    return { month: r2(month), ytd: r2(ytd) };
  };

  const used = new Set();
  const out = [];
  const zero = () => ({ aM: 0, aY: 0, bM: 0, bY: 0 });
  const add = (t, o) => { t.aM = r2(t.aM + o.aM); t.aY = r2(t.aY + o.aY); t.bM = r2(t.bM + o.bM); t.bY = r2(t.bY + o.bY); };

  const revT = zero(), opexT = zero(), otherT = zero();
  let grpT = zero();
  let curSection = null;
  const unmappedBudget = [];
  const tieChecks = [];
  let debtRow = null;

  for (const row of lines) {
    if (row.kind === 'section') {
      curSection = row.section;
      out.push({ kind: 'section', label: row.label });
      continue;
    }
    if (row.kind === 'group') {
      grpT = zero();
      out.push({ kind: 'group', label: row.label });
      continue;
    }
    if (row.kind === 'line') {
      const isRev = curSection === 'Revenue';
      const codes = map.get(norm(row.label)) || [];
      let aM = 0, aY = 0;
      for (const c of codes) { const a = actOf(c); aM += a.month; aY += a.ytd; used.add(String(c)); }
      const b = budOf(row);
      const cell = { aM: r2(aM), aY: r2(aY), bM: b.month, bY: b.ytd };
      if (!codes.length) unmappedBudget.push(row.label);
      out.push({
        kind: 'line', label: row.label, codes, note: row.note,
        sense: isRev ? 'rev' : 'exp', mapped: codes.length > 0, ...cell,
      });
      add(isRev ? revT : opexT, cell);
      if (!isRev) add(grpT, cell);
      continue;
    }
    if (row.kind === 'subtotal') {
      out.push({ kind: 'subtotal', label: row.label, sense: 'exp', ...grpT });
      // The workbook's own subtotal is a cross-check on the parse, not a source.
      const b = budOf(row);
      if (Math.abs(b.month - grpT.bM) > 0.5) {
        tieChecks.push({ label: row.label, workbook: b.month, computed: grpT.bM });
      }
      continue;
    }
    if (row.kind === 'total') {
      const isRev = /revenue/i.test(row.label);
      out.push({ kind: 'total', label: row.label, sense: isRev ? 'rev' : 'exp', ...(isRev ? revT : opexT) });
      const b = budOf(row);
      const t = isRev ? revT : opexT;
      if (Math.abs(b.month - t.bM) > 0.5) tieChecks.push({ label: row.label, workbook: b.month, computed: t.bM });
      continue;
    }
    if (row.kind === 'debt') {
      // Compared directly against actual interest expense. The budget's
      // Projected Debt Service is interest only: every 2026 rail workbook
      // computes it as -(ending balance x rate / 12), and the loan balance is
      // level all year with no payoffs, so there is no principal component to
      // strip out. Stored negative in the workbook; presented positive here,
      // as every other cost on this schedule is.
      const codes = map.get(norm(row.label)) || [];
      let aM = 0, aY = 0;
      for (const c of codes) { const a = actOf(c); aM += a.month; aY += a.ytd; used.add(String(c)); }
      const b = budOf(row);
      if (!codes.length) unmappedBudget.push(row.label);
      debtRow = {
        kind: 'debtline', label: 'Interest Expense', codes, sense: 'exp',
        mapped: codes.length > 0, budgetLabel: row.label,
        aM: r2(aM), aY: r2(aY), bM: r2(Math.abs(b.month)), bY: r2(Math.abs(b.ytd)),
      };
      continue;
    }
    // 'cashflow' and anything else: not presented.
  }

  // ── Actual accounts with activity and no budget line ─────────────────────
  // Split by code: operating accounts join the operating expense section under
  // an explicit heading; 69xxx/7xxxx/8xxxx sit below Net Operating Income in
  // Other Income (Expense), where the budget has nothing to say at all.
  const extraOpex = [], extraOther = [];
  for (const [code, a] of meta) {
    if (a.type !== 'Revenue' && a.type !== 'Expense') continue;
    if (used.has(String(code))) continue;
    const act = actOf(code);
    if (Math.abs(act.month) < 0.005 && Math.abs(act.ytd) < 0.005) continue;
    const isOther = OTHER_CODE.test(String(code)) || a.type === 'Revenue' && OTHER_CODE.test(String(code));
    const row = {
      kind: 'line', label: a.name, codes: [String(code)], mapped: true, unbudgeted: true,
      aM: act.month, aY: act.ytd, bM: 0, bY: 0,
    };
    if (OTHER_CODE.test(String(code))) {
      // Presented as a contribution to income: expenses negative.
      const sign = a.type === 'Expense' ? -1 : 1;
      extraOther.push({ ...row, sense: 'other', aM: r2(sign * act.month), aY: r2(sign * act.ytd) });
    } else if (a.type === 'Revenue') {
      extraOpex.push({ ...row, sense: 'rev', _rev: true });
    } else {
      extraOpex.push({ ...row, sense: 'exp' });
    }
  }
  // Unbudgeted revenue belongs in the revenue section, not with the expenses.
  const extraRev = extraOpex.filter(r => r._rev);
  const extraExp = extraOpex.filter(r => !r._rev);
  extraRev.forEach(r => { delete r._rev; add(revT, r); });
  extraExp.forEach(r => add(opexT, r));
  extraOther.forEach(r => add(otherT, r));
  extraRev.sort((a, b) => a.codes[0].localeCompare(b.codes[0]));
  extraExp.sort((a, b) => a.codes[0].localeCompare(b.codes[0]));
  extraOther.sort((a, b) => a.codes[0].localeCompare(b.codes[0]));

  // Splice the unbudgeted operating rows in ahead of their section totals and
  // recompute the totals rows already emitted (they were built from the budget
  // lines alone).
  const rows = [];
  for (const r of out) {
    if (r.kind === 'total' && /revenue/i.test(r.label) && extraRev.length) {
      rows.push({ kind: 'group', label: 'Other revenue (no budget line)' });
      extraRev.forEach(x => rows.push(x));
    }
    if (r.kind === 'total' && !/revenue/i.test(r.label) && extraExp.length) {
      rows.push({ kind: 'group', label: 'Other operating expenses (no budget line)' });
      extraExp.forEach(x => rows.push(x));
    }
    if (r.kind === 'total') {
      const t = /revenue/i.test(r.label) ? revT : opexT;
      rows.push({ ...r, ...t });
      continue;
    }
    rows.push(r);
  }

  const noi = {
    aM: r2(revT.aM - opexT.aM), aY: r2(revT.aY - opexT.aY),
    bM: r2(revT.bM - opexT.bM), bY: r2(revT.bY - opexT.bY),
  };
  rows.push({ kind: 'noi', label: 'Net Operating Income', sense: 'rev', ...noi });

  // ── Debt service ─────────────────────────────────────────────────────────
  // Projected debt service against actual interest INCURRED, which on a
  // development-stage rail asset is mostly capitalised rather than expensed.
  // Comparing it to interest expense alone reads as a huge favourable variance
  // that is not real: at 6/30/2026 SRN had expensed nothing against 1,373,303 of
  // projected debt service while capitalising 634,330 to construction in
  // progress. So the two components are shown separately and their total is what
  // carries the variance.
  //
  // Net Income (Loss) still deducts only the EXPENSED portion — capitalised
  // interest is added to the asset, not to the period — which is why net income
  // continues to tie to the general ledger.
  const debt = debtRow || { aM: 0, aY: 0, bM: 0, bY: 0 };
  let capInt = { aM: 0, aY: 0, codes: [] };
  for (const [code, a] of meta) {
    if (a.type !== 'Asset' || !CAPITALISED_INTEREST.test(String(a.name || ''))) continue;
    const act = actOf(code);
    if (Math.abs(act.month) < 0.005 && Math.abs(act.ytd) < 0.005) continue;
    capInt.aM = r2(capInt.aM + act.month);
    capInt.aY = r2(capInt.aY + act.ytd);
    capInt.codes.push(String(code));
  }
  const totalInterest = { aM: r2(debt.aM + capInt.aM), aY: r2(debt.aY + capInt.aY), bM: debt.bM, bY: debt.bY };
  if (debtRow || capInt.codes.length) {
    rows.push({ kind: 'section', label: 'Debt Service' });
    if (capInt.codes.length) {
      // Split presentation: components are memo rows (no budget of their own),
      // and the total carries the comparison.
      // Built explicitly rather than spreading debtRow: an entity could
      // capitalise interest with no Projected Debt Service line in its budget,
      // and spreading null would emit a row with no kind at all.
      rows.push({
        kind: 'debtline', role: 'expensed', label: 'Interest expensed', sense: 'exp', memo: true,
        codes: debtRow ? debtRow.codes : [], mapped: debtRow ? debtRow.mapped : false,
        budgetLabel: debtRow ? debtRow.budgetLabel : null,
        aM: debt.aM, aY: debt.aY, bM: 0, bY: 0,
      });
      rows.push({
        kind: 'debtline', role: 'capitalised',
        label: 'Interest capitalised to construction in progress',
        codes: capInt.codes, sense: 'exp', memo: true, mapped: true,
        aM: capInt.aM, aY: capInt.aY, bM: 0, bY: 0,
      });
      rows.push({
        kind: 'debttotal', role: 'total', label: 'Total interest incurred', sense: 'exp',
        mapped: !debtRow || debtRow.mapped,
        aM: totalInterest.aM, aY: totalInterest.aY, bM: totalInterest.bM, bY: totalInterest.bY,
      });
    } else if (debtRow) {
      rows.push({ ...debtRow, role: 'expensed' });
    }
    rows.push({
      kind: 'cashflow', label: 'Cash Flow After Debt Service', sense: 'rev',
      aM: r2(noi.aM - totalInterest.aM), aY: r2(noi.aY - totalInterest.aY),
      bM: r2(noi.bM - totalInterest.bM), bY: r2(noi.bY - totalInterest.bY),
    });
  }

  if (extraOther.length) {
    rows.push({ kind: 'group', label: 'Other Income (Expense) — not in the operating budget' });
    extraOther.forEach(x => rows.push(x));
    rows.push({ kind: 'subtotal', label: 'Total Other Income (Expense)', sense: 'other', ...otherT });
  }
  const net = {
    aM: r2(noi.aM - debt.aM + otherT.aM), aY: r2(noi.aY - debt.aY + otherT.aY),
    bM: r2(noi.bM - debt.bM + otherT.bM), bY: r2(noi.bY - debt.bY + otherT.bY),
  };
  rows.push({ kind: 'net', label: 'Net Income (Loss)', sense: 'rev', ...net });

  return {
    meta: {
      entityName: entityName || null,
      asOf, fiscalYear, monthIdx,
      versionId: version.id, versionNo: version.version_no,
      versionLabel: version.label, uploadedAt: version.uploaded_at,
      sourceFile: version.original_name,
    },
    rows,
    totals: { revenue: revT, opex: opexT, noi, debt, capInt, totalInterest, other: otherT, net },
    debtService: (debtRow || capInt.codes.length) ? {
      expensed: { aM: debt.aM, aY: debt.aY },
      capitalised: { aM: capInt.aM, aY: capInt.aY, codes: capInt.codes },
      total: { aM: totalInterest.aM, aY: totalInterest.aY },
      bM: debt.bM, bY: debt.bY,
      mapped: debtRow ? debtRow.mapped : false,
    } : null,
    unmapped: { budgetLabels: unmappedBudget, unbudgetedAccounts: extraExp.concat(extraRev).map(r => r.codes[0]) },
    tieChecks,
  };
}

module.exports = {
  BUDGET_FOLDER,
  ensureSchema, parseWorkbook, saveVersion, activeVersion, listVersions, versionLines,
  seedMap, getMap, setMap, norm, buildBudgetToActual,
  DEFAULT_MAP, NO_DEFAULT,
  _helpers: { priorMonthEnd, priorYearEnd, r2 },
};
