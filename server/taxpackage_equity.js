// ─── Tax Package: generated Equity Rollforward + Contributed Capital ──────────
//
// Both tabs are derived entirely from the entity's equity-account GL activity so
// they reflect how the equity accounts actually changed during the year and tie
// to the Trial Balance. Net income is NOT allocated to members: the whole year's
// net income flows to a single Retained Earnings column; member-level allocation
// is the tax preparer's job and is not in the books.
//
//   • EQ RF              — inception-to-date rollforward, one column per member
//                          (contributed-capital + distribution accounts grouped by
//                          member), a Retained Earnings column, and a Total that
//                          ties to Members' Equity (A - L, incl. current-year NI)
//                          with a Variance = 0 check. Per year: Beginning ->
//                          Capital Contributed -> Net Income -> Distribution -> Close.
//   • Contributed Capital — per contributed-capital account: prior years rolled
//                          into one "PY" line, current year itemized by date/memo/
//                          amount, summing to the account's TB balance, plus a
//                          "Contributed Capital per GL" reconciliation to the TB.

const FONT = 'Verdana', SIZE = 8;
const ACCT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const F = (o = {}) => Object.assign({ name: FONT, size: SIZE }, o);
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const yrOf = (d) => Number(String(d || '').slice(0, 4)) || 0;

// Entity-specific EQ RF footnotes (mirroring the CLA packages). Keyed by entity id;
// anchored to the Beginning Balance row with a superscript marker.
const EQ_RF_NOTES = {
  38: { marker: 'A', text: 'The $148,117.07 distribution to Ben, the former owner, has been reclassed to retained earnings; this reclassification does not relate to 2025 operating results or distributions.' },
};

// Entity-specific EQ RF start year: roll forward from 1/1 of this year with an
// opening balance, hiding prior-year detail (e.g. an entity acquired that year).
const EQ_RF_START = { 38: 2025 };

// ── Standalone strip ──────────────────────────────────────────────────────────
// Build a child entity's tax package on its OWN (standalone) trial balance by
// removing a parent entity's overlay that has been mingled onto the child's books.
// Keyed by the CHILD entity code. For CLIP Property Owner (CLIPPROP), the parent
// County Line Industrial Park (42) owns it, but County Line's bank accounts, its
// intercompany receivable, and its member contributed-capital accounts are posted
// onto CLIP. On a standalone CLIP TB those belong to the parent, and CLIP's equity
// is a single owner line = the parent's investment in CLIP.
//   remove   — codes dropped from the child outright (parent's investment-in-child
//              asset, and the stale owner line, which is rebuilt below).
//   ccAccount/ccName/ccMember — the single owner contributed-capital line.
//   ccSource — the parent account whose balance becomes that owner line
//              (County Line's "Investment in CLIP Property Owner").
// Every OTHER parent account is subtracted from the child's same-code balance
// (balance sheet AND income statement, per Jimmy's instruction).
const STANDALONE_STRIP = {
  CLIPPROP: {
    parentId: 42,
    parentNameRe: /county line industrial park/i,
    remove: ['19041', '34164'],
    ccAccount: '34164',
    ccName: 'Contributed Capital - County Line Industrial Park LLC',
    ccMember: 'County Line Industrial Park LLC',
    ccSource: '19041',
  },
  // CLR Silsbee Property Owner (39): parent County Line Rail Silsbee LLC (52) holds
  // its investment in the property owner (17001) plus the LLC-level member capital,
  // all mingled onto the child's books. On a standalone CLR Silsbee Property Owner
  // TB those belong to the parent, and the child's equity collapses to a single
  // owner line = the parent's investment in the property owner (17001 -> 34063).
  // Grand Haven is removed outright under BOTH codes: the child's real line is
  // 34144, but it also carries an empty 34251 matching the parent's code, so the
  // same-code net (child 0 - parent 250,000) would otherwise leave a phantom
  // -250,000. The property owner's sole member is the LLC, so no LLC-level
  // investor belongs on its books either way.
  CLRSILSB2: {
    parentId: 52,
    parentNameRe: /county line rail silsbee/i,
    remove: ['17001', '34063', '34144', '34251'],
    ccAccount: '34063',
    ccName: 'Contributed Capital - County Line Rail Silsbee LLC',
    ccMember: 'County Line Rail Silsbee LLC',
    ccSource: '17001',
  },
};

// Resolve + validate the strip config for an entity (looked up by its code).
// Returns the config augmented with the resolved parent row, or null when the
// entity has no standalone strip. Throws if the configured parent is missing or
// its name no longer matches (guards against an entity-id reshuffle).
function resolveStrip(ctx, ent) {
  if (!ent) return null;
  const cfg = STANDALONE_STRIP[String(ent.code)];
  if (!cfg) return null;
  const parent = ctx.db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(cfg.parentId);
  if (!parent || !cfg.parentNameRe.test(String(parent.name || ''))) {
    throw new Error('[tax-package] standalone strip for ' + ent.code + ': parent entity id ' + cfg.parentId + ' not found or its name no longer matches ' + String(cfg.parentNameRe));
  }
  return Object.assign({}, cfg, { parent });
}

// Apply the balance strip to one computeBalances result set. Rows are natural-
// signed (asset debit-positive, liability/equity credit-positive). `opts` are the
// SAME options passed to the child's computeBalances, reused for the parent so it
// is read at the matching as_of. Returns { rows, dropped } — `dropped` is the set
// of child codes removed or netted to ~0, used to prune the matching GL detail.
function applyBalanceStrip(ctx, ent, opts, childRows) {
  let cfg;
  try { cfg = resolveStrip(ctx, ent); } catch (e) { throw e; }
  if (!cfg) return { rows: childRows, dropped: new Set() };
  const parentRows = ctx.computeBalances(cfg.parentId, opts) || [];
  const pMap = new Map(parentRows.map((r) => [String(r.code), Number(r.balance) || 0]));
  const remove = new Set(cfg.remove);
  const dropped = new Set();
  const out = [];
  for (const r of childRows) {
    const code = String(r.code);
    if (remove.has(code)) { dropped.add(code); continue; }
    if (pMap.has(code)) {
      const nb = r2((Number(r.balance) || 0) - pMap.get(code));
      if (Math.abs(nb) < 0.005) dropped.add(code);
      out.push(Object.assign({}, r, { balance: nb }));
    } else {
      out.push(r);
    }
  }
  // Single owner contributed-capital line = the parent's investment in the child.
  const cc = r2(pMap.get(cfg.ccSource) || 0);
  if (Math.abs(cc) >= 0.005) {
    out.push({ code: cfg.ccAccount, name: cfg.ccName, type: 'Equity', subtype: '', bank_acct: 0, balance: cc, total_debit: 0, total_credit: 0 });
  }
  return { rows: out, dropped };
}

function categorize(a) {
  const c = String(a.code || ''); const n = String(a.name || '').toLowerCase();
  if (c.startsWith('39') || c === '31000' || /retained earning/.test(n)) return 're';
  if (c.startsWith('33') || /^distribution/.test(n)) return 'distrib';
  return 'contrib';
}
function memberOf(a) {
  const n = String(a.name || '')
    .replace(/^contributed capital\s*[-\u2013]\s*/i, '')
    .replace(/^distribution\s*[-\u2013]\s*/i, '')
    .replace(/\s*[-\u2013(]*\s*undeployed capital.*$/i, '')
    .replace(/\s*\(retained earnings\)\s*$/i, '')
    .trim();
  return n || String(a.code);
}

// Standalone equity: a child entity (see STANDALONE_STRIP) is presented with a
// single owner member whose contributed capital = the parent's investment in the
// child (ccSource), itemized from that account's GL. CLIP's own retained earnings
// and net income are kept, stripped of the parent's overlay (child minus parent,
// by year) so the rollforward ties to the standalone trial balance. Ties are by
// construction: membersEquityTB = owner ending + RE ending, so EQ RF Variance = 0,
// and the single owner total = the parent's investment = the stripped TB's owner
// line, so Contributed Capital Difference = 0.
function computeEquityStandalone(ctx, ent, eid, p, cfg) {
  const { db } = ctx;
  const curYear = Number(p.year);
  const ownerMember = cfg.ccMember;

  const escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cleanDescr = (s, code) => {
    let t = String(s || '').trim(); if (!t) return '';
    const c = escRe(code);
    t = t.replace(new RegExp('^' + c + '\\s*-\\s*'), '');
    t = t.replace(new RegExp('\\s*-\\s*[^-]*-\\s*GJ\\s*-\\s*' + c + '$', 'i'), '');
    t = t.replace(new RegExp('\\s*-\\s*GJ\\s*-\\s*' + c + '$', 'i'), '');
    return t.trim();
  };

  // Net income (Revenue/Expense, credit-debit) by year for one entity.
  const niByEntity = (entId) => {
    const m = {};
    const rows = db.prepare(
      "SELECT substr(je.date,1,4) AS yr, SUM(jl.credit - jl.debit) AS ni " +
      "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
      "JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code " +
      "WHERE je.entity_id = ? AND a.type IN ('Revenue','Expense') GROUP BY yr").all(entId);
    for (const r of rows) { const y = Number(r.yr); if (y && y <= curYear) m[y] = r2(Number(r.ni) || 0); }
    return m;
  };
  // Direct journal activity (credit-debit) to a set of RE accounts, by year.
  const reDirectByEntity = (entId, reCodes) => {
    const m = {}; if (!reCodes.length) return m;
    const ph = reCodes.map(() => '?').join(',');
    const rows = db.prepare(
      "SELECT substr(je.date,1,4) AS yr, SUM(jl.credit - jl.debit) AS v " +
      "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
      "WHERE je.entity_id = ? AND jl.account_code IN (" + ph + ") GROUP BY yr").all(entId, ...reCodes);
    for (const r of rows) { const y = Number(r.yr); if (y && y <= curYear) m[y] = r2(Number(r.v) || 0); }
    return m;
  };
  const reCodesOf = (entId) => db.prepare("SELECT code, name FROM accounts WHERE entity_id = ? AND type = 'Equity'").all(entId)
    .filter((a) => categorize(a) === 're').map((a) => String(a.code));

  const clipNI = niByEntity(eid), parentNI = niByEntity(cfg.parentId);
  const clipREd = reDirectByEntity(eid, reCodesOf(eid));
  const parentREd = reDirectByEntity(cfg.parentId, reCodesOf(cfg.parentId));

  // Owner contributions = the parent's investment-in-child account, debit-positive
  // (a debit to the investment is a contribution into the child's equity).
  const invLines = db.prepare(
    "SELECT je.date AS date, jl.debit AS debit, jl.credit AS credit, je.memo AS memo, jl.description AS descr " +
    "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
    "WHERE je.entity_id = ? AND jl.account_code = ? ORDER BY je.date, je.entry_num, jl.id").all(cfg.parentId, cfg.ccSource);
  const ownerByYear = {}; const ownerItems = []; let ownerTotal = 0;
  for (const l of invLines) {
    const y = yrOf(l.date); if (!y || y > curYear) continue;
    const v = r2((Number(l.debit) || 0) - (Number(l.credit) || 0));
    ownerByYear[y] = r2((ownerByYear[y] || 0) + v);
    ownerTotal = r2(ownerTotal + v);
    if (Math.abs(v) >= 0.005) ownerItems.push({ date: l.date, memo: cleanDescr(l.descr, cfg.ccSource) || l.memo || '', amt: v });
  }

  let minYear = curYear;
  const allY = new Set([].concat(Object.keys(clipNI), Object.keys(parentNI), Object.keys(clipREd), Object.keys(parentREd), Object.keys(ownerByYear)).map(Number));
  for (const y of allY) if (y && y < minYear) minYear = y;
  const years = []; for (let y = minYear; y <= curYear; y++) years.push(y);

  const act = {}, reAct = {}, niByYear = {};
  for (const y of years) {
    act[y] = { [ownerMember]: { contrib: r2(ownerByYear[y] || 0), distrib: 0 } };
    niByYear[y] = r2((clipNI[y] || 0) - (parentNI[y] || 0));
    reAct[y] = { direct: r2((clipREd[y] || 0) - (parentREd[y] || 0)) };
  }

  const memberNames = [ownerMember];
  const memberEnd = { [ownerMember]: {} };
  const reEnd = {};
  let mPrev = 0, rePrev = 0;
  for (const y of years) {
    memberEnd[ownerMember][y] = r2(mPrev + act[y][ownerMember].contrib + act[y][ownerMember].distrib);
    mPrev = memberEnd[ownerMember][y];
    reEnd[y] = r2(rePrev + (niByYear[y] || 0) + reAct[y].direct);
    rePrev = reEnd[y];
  }

  const lastY = years[years.length - 1];
  const membersEquityTB = r2((memberEnd[ownerMember][lastY] || 0) + (reEnd[lastY] || 0));
  const contribDetail = [{ code: cfg.ccAccount, name: cfg.ccName, member: ownerMember, items: ownerItems, total: r2(memberEnd[ownerMember][lastY] || 0) }];

  const meta = new Map();
  meta.set(String(cfg.ccAccount), { name: cfg.ccName, cat: 'contrib', member: ownerMember });

  return { years, memberNames, memberEnd, reEnd, niByYear, reAct, act, meta, membersEquityTB, contribDetail, note: null, startYear: null };
}

// Pull the equity picture from the ledger. Returns null when the entity has no
// equity accounts (caller then skips both tabs).
function computeEquity(ctx, eid, p) {
  const { db } = ctx;
  const curYear = Number(p.year);
  // Standalone strip: present a child entity's equity as a single owner line.
  const _ent = db.prepare('SELECT id, name, code FROM entities WHERE id = ?').get(eid);
  const _strip = resolveStrip(ctx, _ent);
  if (_strip) return computeEquityStandalone(ctx, _ent, eid, p, _strip);
  const eqAccts = db.prepare("SELECT code, name, type FROM accounts WHERE entity_id = ? AND type = 'Equity'").all(eid);
  if (!eqAccts.length) return null;
  const meta = new Map();
  for (const a of eqAccts) meta.set(String(a.code), { name: a.name || '', cat: categorize(a), member: memberOf(a) });
  const codes = [...meta.keys()];
  const ph = codes.map(() => '?').join(',');
  const lines = db.prepare(
    "SELECT je.date AS date, jl.account_code AS code, jl.debit AS debit, jl.credit AS credit, " +
    "je.memo AS memo, jl.description AS descr " +
    "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
    "WHERE je.entity_id = ? AND jl.account_code IN (" + ph + ") " +
    "ORDER BY je.date, je.entry_num, jl.id").all(eid, ...codes);
  const niRows = db.prepare(
    "SELECT substr(je.date,1,4) AS yr, SUM(jl.credit - jl.debit) AS ni " +
    "FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
    "JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code " +
    "WHERE je.entity_id = ? AND a.type IN ('Revenue','Expense') GROUP BY yr").all(eid);

  const memberNames = [];
  for (const [, m] of meta) if (m.cat === 'contrib' && !memberNames.includes(m.member)) memberNames.push(m.member);
  for (const [, m] of meta) if (m.cat === 'distrib' && !memberNames.includes(m.member)) memberNames.push(m.member);

  let minYear = curYear;
  const act = {}, reAct = {};
  const touch = (y) => { if (y < minYear) minYear = y; if (!act[y]) act[y] = {}; if (!reAct[y]) reAct[y] = { direct: 0 }; };
  for (const ln of lines) {
    const y = yrOf(ln.date); if (!y || y > curYear) continue; touch(y);
    const m = meta.get(String(ln.code)); const v = r2((Number(ln.credit) || 0) - (Number(ln.debit) || 0));
    if (m.cat === 're') { reAct[y].direct = r2(reAct[y].direct + v); continue; }
    if (!act[y][m.member]) act[y][m.member] = { contrib: 0, distrib: 0 };
    if (m.cat === 'distrib') act[y][m.member].distrib = r2(act[y][m.member].distrib + v);
    else act[y][m.member].contrib = r2(act[y][m.member].contrib + v);
  }
  const niByYear = {};
  for (const row of niRows) { const y = Number(row.yr); if (!y || y > curYear) continue; touch(y); niByYear[y] = r2(Number(row.ni) || 0); }

  const years = [];
  for (let y = minYear; y <= curYear; y++) years.push(y);

  const memberEnd = {}; for (const name of memberNames) memberEnd[name] = {};
  const reEnd = {};
  const memberPrev = {}; for (const name of memberNames) memberPrev[name] = 0; let rePrev = 0;
  for (const y of years) {
    for (const name of memberNames) {
      const a = (act[y] && act[y][name]) || { contrib: 0, distrib: 0 };
      memberEnd[name][y] = r2(memberPrev[name] + a.contrib + a.distrib);
      memberPrev[name] = memberEnd[name][y];
    }
    reEnd[y] = r2(rePrev + (niByYear[y] || 0) + ((reAct[y] && reAct[y].direct) || 0));
    rePrev = reEnd[y];
  }

  const tbRow = db.prepare(
    "SELECT SUM(jl.credit - jl.debit) AS v FROM journal_lines jl JOIN journal_entries je ON je.id = jl.entry_id " +
    "JOIN accounts a ON a.entity_id = je.entity_id AND a.code = jl.account_code " +
    "WHERE je.entity_id = ? AND a.type IN ('Equity','Revenue','Expense') AND je.date <= ?").get(eid, p.tbAsOf);
  const membersEquityTB = r2((tbRow && tbRow.v) || 0);

  // Inception-to-date contribution detail per contributed-capital account (CLA
  // format): every entry affecting the account, with the real GL line description
  // (not the generic "GL detail import …" entry memo), signed so the itemized
  // total foots to the account's GL balance for the member-vs-GL tie-out.
  const escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cleanDescr = (s, code) => {
    let t = String(s || '').trim(); if (!t) return '';
    const c = escRe(code);
    t = t.replace(new RegExp('^' + c + '\\s*-\\s*'), '');                            // leading "<code>-"
    t = t.replace(new RegExp('\\s*-\\s*[^-]*-\\s*GJ\\s*-\\s*' + c + '$', 'i'), '');   // trailing "-<doc>-GJ-<code>"
    t = t.replace(new RegExp('\\s*-\\s*GJ\\s*-\\s*' + c + '$', 'i'), '');             // trailing "-GJ-<code>"
    return t.trim();
  };
  const contribDetail = [];
  for (const [code, m] of meta) {
    if (m.cat !== 'contrib') continue;
    const items = []; let total = 0;
    for (const l of lines) {
      if (String(l.code) !== code) continue;
      const y = yrOf(l.date); if (!y || y > curYear) continue;
      const v = r2((Number(l.credit) || 0) - (Number(l.debit) || 0));
      if (Math.abs(v) < 0.005) continue;
      items.push({ date: l.date, memo: cleanDescr(l.descr, code) || l.memo || '', amt: v });
      total = r2(total + v);
    }
    contribDetail.push({ code, name: m.name, member: m.member, items, total: r2(total) });
  }
  contribDetail.sort((a, b) => (Number(a.code) - Number(b.code)) || String(a.code).localeCompare(b.code));

  return { years, memberNames, memberEnd, reEnd, niByYear, reAct, act, meta, membersEquityTB, contribDetail, note: EQ_RF_NOTES[Number(eid)] || null, startYear: EQ_RF_START[Number(eid)] || null };
}

function renderEqRF(wb, d) {
  const eq = d.equity; if (!eq) return;
  const ws = wb.addWorksheet('EQ RF');
  ws.getCell('A1').value = d.ent.name; ws.getCell('A1').font = F({ bold: true });
  ws.getCell('A2').value = 'Equity Rollforward'; ws.getCell('A2').font = F();
  ws.getCell('A3').value = 'Year Ended December 31, ' + d.p.year; ws.getCell('A3').font = F();
  ws.getCell('A4').value = '(Unaudited)'; ws.getCell('A4').font = F();
  const lastY = eq.years[eq.years.length - 1];
  // Only members with a balance at year-end are shown (mirroring CLA — members
  // that fully exited show nothing). Dropping zero-balance members never changes
  // the Total, so the TB tie is preserved.
  const withBal = eq.memberNames.filter((n) => Math.abs(eq.memberEnd[n][lastY] || 0) >= 0.005);
  // Surviving members hold positive contributed capital at year-end. A member whose
  // capital was fully bought out during the year (a former owner) is "exited": its
  // entire equity is folded into the opening balance (net assets acquired), mirroring
  // how the CLA packages present a mid-year buyout — the exited member is not shown.
  const contribEnd = (n) => eq.years.reduce((s, y) => s + ((eq.act[y] && eq.act[y][n] && eq.act[y][n].contrib) || 0), 0);
  const survivors = withBal.filter((n) => contribEnd(n) > 0.005);
  const members = survivors.length ? survivors : withBal;
  const exitedEnding = r2(withBal.filter((n) => !members.includes(n)).reduce((s, n) => s + (eq.memberEnd[n][lastY] || 0), 0));
  // Single surviving member: no separate Retained Earnings column — the sole member
  // column carries net income and the folded-in exited equity.
  const single = members.length === 1;
  // Roll forward from a configured start year (e.g. an entity acquired mid-stream):
  // the opening balance is the cumulative equity through the prior year-end, and
  // only activity from the start year on is shown. Defaults to inception.
  const startYear = (eq.startYear && eq.startYear > eq.years[0]) ? eq.startYear : eq.years[0];
  const priorYear = startYear - 1;
  const hasPrior = priorYear >= eq.years[0];
  const beginMember = (n) => hasPrior ? (eq.memberEnd[n][priorYear] || 0) : 0;
  const beginRE = hasPrior ? (eq.reEnd[priorYear] || 0) : 0;
  const shownYears = eq.years.filter((y) => y >= startYear);
  const colStart = 5;
  const memCols = members.map((_, i) => colStart + i);
  const reCol = single ? null : colStart + members.length;
  const totCol = single ? colStart + members.length : reCol + 1;
  members.forEach((name, i) => { const c = ws.getCell(6, memCols[i]); c.value = name; c.font = F({ bold: true }); c.alignment = { wrapText: true, horizontal: 'center' }; });
  if (reCol) { const c = ws.getCell(6, reCol); c.value = 'Retained Earnings'; c.font = F({ bold: true }); c.alignment = { wrapText: true, horizontal: 'center' }; }
  const tcHead = ws.getCell(6, totCol); tcHead.value = "Total Members' Equity"; tcHead.font = F({ bold: true }); tcHead.alignment = { wrapText: true, horizontal: 'center' };
  const num = (r, col, v) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(); c.numFmt = ACCT; };
  let row = 8;
  ws.getCell(row, 2).value = 'Beginning Balance 01/01/' + startYear; ws.getCell(row, 2).font = F({ bold: true });
  let begTot = 0, runSingle = 0;
  if (single) {
    const v = r2(beginMember(members[0]) + beginRE + exitedEnding);
    num(row, memCols[0], v); begTot = v; runSingle = v;
  } else {
    members.forEach((n, i) => { const v = beginMember(n); num(row, memCols[i], v); begTot += v; });
    const reBeg = r2(beginRE + exitedEnding); num(row, reCol, reBeg); begTot += reBeg;
  }
  num(row, totCol, begTot);
  if (eq.note) { const mk = ws.getCell(row, totCol + 1); mk.value = eq.note.marker; mk.font = F({ bold: true, vertAlign: 'superscript' }); }
  row += 2;
  for (const y of shownYears) {
    const contribs = members.map((n) => (eq.act[y] && eq.act[y][n] && eq.act[y][n].contrib) || 0);
    if (contribs.some((v) => Math.abs(v) >= 0.005)) {
      ws.getCell(row, 3).value = 'Capital Contributed'; ws.getCell(row, 3).font = F();
      let tot = 0; members.forEach((n, i) => { num(row, memCols[i], contribs[i]); tot += contribs[i]; });
      if (single) runSingle = r2(runSingle + tot);
      num(row, totCol, tot); row++;
    }
    const ni = r2((eq.niByYear[y] || 0) + ((eq.reAct[y] && eq.reAct[y].direct) || 0));
    if (Math.abs(ni) >= 0.005) {
      ws.getCell(row, 3).value = 'Net Income (Loss)'; ws.getCell(row, 3).font = F();
      num(row, single ? memCols[0] : reCol, ni); num(row, totCol, ni);
      if (single) runSingle = r2(runSingle + ni); row++;
    }
    const distribs = members.map((n) => (eq.act[y] && eq.act[y][n] && eq.act[y][n].distrib) || 0);
    if (distribs.some((v) => Math.abs(v) >= 0.005)) {
      ws.getCell(row, 3).value = 'Distribution'; ws.getCell(row, 3).font = F();
      let tot = 0; members.forEach((n, i) => { num(row, memCols[i], distribs[i]); tot += distribs[i]; });
      if (single) runSingle = r2(runSingle + tot);
      num(row, totCol, tot); row++;
    }
    ws.getCell(row, 2).value = 'Balance at Closing 12/31/' + y; ws.getCell(row, 2).font = F({ bold: true });
    if (single) {
      num(row, memCols[0], runSingle); num(row, totCol, runSingle);
    } else {
      let tot = 0; members.forEach((n, i) => { const v = eq.memberEnd[n][y]; num(row, memCols[i], v); tot += v; });
      const reV = r2(eq.reEnd[y] + exitedEnding); num(row, reCol, reV); tot += reV; num(row, totCol, tot);
    }
    row += 2;
  }
  const closing = single ? runSingle : r2(members.reduce((s, n) => s + eq.memberEnd[n][lastY], 0) + eq.reEnd[lastY] + exitedEnding);
  num(row, totCol, eq.membersEquityTB); ws.getCell(row, totCol + 1).value = "Members' Equity (TB)"; ws.getCell(row, totCol + 1).font = F();
  num(row + 1, totCol, r2(closing - eq.membersEquityTB)); ws.getCell(row + 1, totCol + 1).value = 'Variance'; ws.getCell(row + 1, totCol + 1).font = F();
  if (eq.note) {
    const nr = row + 3, nc = 2;
    ws.getCell(nr, nc).value = 'Note ' + eq.note.marker; ws.getCell(nr, nc).font = F({ bold: true });
    try { ws.mergeCells(nr + 1, nc, nr + 4, nc + 6); } catch (e) { /* best effort */ }
    const tc = ws.getCell(nr + 1, nc); tc.value = eq.note.text; tc.font = F(); tc.alignment = { wrapText: true, vertical: 'top' };
  }
  ws.getColumn(2).width = 24; ws.getColumn(3).width = 20;
  memCols.concat(reCol ? [reCol, totCol] : [totCol]).forEach((c) => { ws.getColumn(c).width = 16; });
  ws.getColumn(totCol + 1).width = 22;
}

function renderContributedCapital(wb, d) {
  const eq = d.equity; if (!eq) return;
  // Only members with a balance at 12/31 — zero-balance (fully returned) members
  // are dropped.
  const details = eq.contribDetail.filter((cd) => Math.abs(cd.total) >= 0.005);
  if (!details.length) return;
  const ws = wb.addWorksheet('Contributed Capital');
  const num = (r, col, v, bold) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(bold ? { bold: true } : {}); c.numFmt = ACCT; };
  // Contributed-capital balance per the general ledger (computeBalances, via d.tb),
  // for the member-contributions-vs-GL tie-out. The TB closing is debit-positive, so
  // an equity credit balance shows negative; the contributed-capital magnitude is -closing.
  const tbMap = new Map((d.tb || []).map((t) => [String(t.code), Number(t.closing) || 0]));
  const glOf = (code) => r2(-(tbMap.get(String(code)) || 0));
  let row = 1, grand = 0;
  for (const cd of details) { ws.getCell(row, 1).value = cd.name; ws.getCell(row, 1).font = F(); num(row, 2, cd.total); grand = r2(grand + cd.total); row++; }
  ws.getCell(row, 1).value = 'Total'; ws.getCell(row, 1).font = F({ bold: true }); num(row, 2, grand, true); row += 2;
  ws.getCell(row, 1).value = 'Contributions by members thru 12/31/' + d.p.year; ws.getCell(row, 1).font = F({ bold: true }); row += 2;
  for (const cd of details) {
    ws.getCell(row, 2).value = 'Contributions made by ' + cd.member + ':'; ws.getCell(row, 2).font = F({ bold: true }); row++;
    for (const l of cd.items) {
      ws.getCell(row, 1).value = l.date; ws.getCell(row, 1).font = F();
      ws.getCell(row, 2).value = l.memo; ws.getCell(row, 2).font = F();
      num(row, 3, l.amt); row++;
    }
    num(row, 3, cd.total, true); row += 2;
  }
  ws.getCell(row, 2).value = 'Total contributions by members thru 12/31/' + d.p.year; ws.getCell(row, 2).font = F({ bold: true });
  num(row, 3, grand, true); row += 2;
  // Member-contributions-vs-GL tie-out: the itemized contributions (from the GL
  // detail above) vs the contributed-capital balance per the trial balance.
  const glGrand = r2(details.reduce((s, cd) => s + glOf(cd.code), 0));
  ws.getCell(row, 1).value = 'Reconciliation to the general ledger'; ws.getCell(row, 1).font = F({ bold: true }); row++;
  ws.getCell(row, 2).value = 'Total contributions by members'; ws.getCell(row, 2).font = F(); num(row, 3, grand); row++;
  ws.getCell(row, 2).value = 'Contributed Capital per general ledger'; ws.getCell(row, 2).font = F(); num(row, 3, glGrand); row++;
  ws.getCell(row, 2).value = 'Difference'; ws.getCell(row, 2).font = F({ bold: true }); num(row, 3, r2(grand - glGrand), true);
  ws.getColumn(1).width = 18; ws.getColumn(2).width = 60; ws.getColumn(3).width = 16; ws.getColumn(4).width = 20;
}

// Fixed Asset Additions — current-year movement on the capitalized / PP&E asset
// accounts, so the preparer sees what was added during the year. d.fixedAssets is
// computed in taxpackage.js (Beginning + Additions − Disposals = Ending, tying to
// the TB).
function renderFixedAssets(wb, d) {
  const fa = d.fixedAssets; if (!fa || !fa.length) return;
  const ws = wb.addWorksheet('Fixed Asset Additions');
  ws.getCell('A1').value = d.ent.name; ws.getCell('A1').font = F({ bold: true });
  ws.getCell('A2').value = 'Fixed Asset Additions'; ws.getCell('A2').font = F();
  ws.getCell('A3').value = 'Year Ended December 31, ' + d.p.year; ws.getCell('A3').font = F();
  ['Code', 'Account', 'Beginning', 'Additions', 'Disposals', 'Ending'].forEach((h, i) => { const c = ws.getCell(5, i + 1); c.value = h; c.font = F({ bold: true }); });
  const num = (r, col, v, bold) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(bold ? { bold: true } : {}); c.numFmt = ACCT; };
  let row = 6, tB = 0, tA = 0, tD = 0, tE = 0;
  for (const a of fa) {
    ws.getCell(row, 1).value = /^\d+$/.test(String(a.code)) ? Number(a.code) : a.code; ws.getCell(row, 1).font = F();
    ws.getCell(row, 2).value = a.name; ws.getCell(row, 2).font = F();
    num(row, 3, a.beginning); num(row, 4, a.additions); num(row, 5, a.disposals); num(row, 6, a.ending);
    tB = r2(tB + a.beginning); tA = r2(tA + a.additions); tD = r2(tD + a.disposals); tE = r2(tE + a.ending); row++;
  }
  ws.getCell(row, 2).value = 'Total'; ws.getCell(row, 2).font = F({ bold: true });
  num(row, 3, tB, true); num(row, 4, tA, true); num(row, 5, tD, true); num(row, 6, tE, true);
  ws.getColumn(1).width = 14; ws.getColumn(2).width = 46;
  [3, 4, 5, 6].forEach((c) => { ws.getColumn(c).width = 15; });
}

// Amortization Schedule — booked rollforward of capitalized deferred costs (loan
// fees, financing costs) amortized via accumulated amortization into amortization
// expense. Ties to the TB. d.amortization is computed in taxpackage.js.
function renderAmortization(wb, d) {
  const am = d.amortization; if (!am) return;
  const ws = wb.addWorksheet('Goodwill Amortization');
  ws.getCell('A1').value = d.ent.name; ws.getCell('A1').font = F({ bold: true });
  ws.getCell('A2').value = 'Goodwill Amortization'; ws.getCell('A2').font = F();
  ws.getCell('A3').value = (d.p.mode === 'monthly' ? d.p.periodLine : 'Year Ended December 31, ' + d.p.year); ws.getCell('A3').font = F();
  ['Code', 'Account', 'Beginning', 'Additions', 'Amortization', 'Ending'].forEach((h, i) => { const c = ws.getCell(5, i + 1); c.value = h; c.font = F({ bold: true }); });
  const num = (r, col, v, bold) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(bold ? { bold: true } : {}); c.numFmt = ACCT; };
  const codeCell = (r, a) => { const c = ws.getCell(r, 1); c.value = /^\d+$/.test(String(a.code)) ? Number(a.code) : a.code; c.font = F(); ws.getCell(r, 2).value = a.name; ws.getCell(r, 2).font = F(); };
  let row = 6, tBeg = 0, tAdd = 0, tAmort = 0, tEnd = 0;
  for (const a of am.cost) {
    codeCell(row, a); num(row, 3, a.beginning); num(row, 4, a.additions); num(row, 6, a.ending);
    tBeg = r2(tBeg + a.beginning); tAdd = r2(tAdd + a.additions); tEnd = r2(tEnd + a.ending); row++;
  }
  for (const a of am.accum) {
    codeCell(row, a); num(row, 3, a.beginning); num(row, 5, a.amortization); num(row, 6, a.ending);
    tBeg = r2(tBeg + a.beginning); tAmort = r2(tAmort + a.amortization); tEnd = r2(tEnd + a.ending); row++;
  }
  ws.getCell(row, 2).value = 'Net Deferred Cost'; ws.getCell(row, 2).font = F({ bold: true });
  num(row, 3, tBeg, true); num(row, 4, tAdd, true); num(row, 5, tAmort, true); num(row, 6, tEnd, true); row += 2;
  const expTotal = r2(am.expense.reduce((s, e) => s + e.amount, 0));
  const amortRecorded = r2(-am.accum.reduce((s, a) => s + a.amortization, 0));
  ws.getCell(row, 2).value = 'Amortization expense for the period'; ws.getCell(row, 2).font = F(); num(row, 6, expTotal); row++;
  ws.getCell(row, 2).value = 'Accumulated amortization recorded'; ws.getCell(row, 2).font = F(); num(row, 6, amortRecorded); row++;
  ws.getCell(row, 2).value = 'Variance'; ws.getCell(row, 2).font = F(); num(row, 6, r2(expTotal - amortRecorded));
  ws.getColumn(1).width = 14; ws.getColumn(2).width = 42; [3, 4, 5, 6].forEach((c) => { ws.getColumn(c).width = 15; });
}

module.exports = { computeEquity, renderEqRF, renderContributedCapital, renderFixedAssets, renderAmortization, applyBalanceStrip, resolveStrip };
