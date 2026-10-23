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

// Pull the equity picture from the ledger. Returns null when the entity has no
// equity accounts (caller then skips both tabs).
function computeEquity(ctx, eid, p) {
  const { db } = ctx;
  const curYear = Number(p.year);
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

  const contribDetail = [];
  for (const [code, m] of meta) {
    if (m.cat !== 'contrib') continue;
    let pyTotal = 0; const cur = [];
    for (const l of lines) {
      if (String(l.code) !== code) continue;
      const y = yrOf(l.date); const v = r2((Number(l.credit) || 0) - (Number(l.debit) || 0));
      if (y < curYear) pyTotal = r2(pyTotal + v);
      else if (y === curYear) cur.push({ date: l.date, memo: l.memo || l.descr || '', amt: v });
    }
    const total = r2(pyTotal + cur.reduce((s, x) => s + x.amt, 0));
    contribDetail.push({ code, name: m.name, member: m.member, pyTotal: r2(pyTotal), cur, total });
  }
  contribDetail.sort((a, b) => (Number(a.code) - Number(b.code)) || String(a.code).localeCompare(b.code));

  return { years, memberNames, memberEnd, reEnd, niByYear, reAct, act, meta, membersEquityTB, contribDetail };
}

function renderEqRF(wb, d) {
  const eq = d.equity; if (!eq) return;
  const ws = wb.addWorksheet('EQ RF');
  ws.getCell('A1').value = d.ent.name; ws.getCell('A1').font = F({ bold: true });
  ws.getCell('A2').value = 'Equity Rollforward'; ws.getCell('A2').font = F();
  ws.getCell('A3').value = 'Year Ended December 31, ' + d.p.year; ws.getCell('A3').font = F();
  ws.getCell('A4').value = '(Unaudited)'; ws.getCell('A4').font = F();
  const colStart = 5;
  const memCols = eq.memberNames.map((_, i) => colStart + i);
  const reCol = colStart + eq.memberNames.length;
  const totCol = reCol + 1;
  eq.memberNames.forEach((name, i) => { const c = ws.getCell(6, memCols[i]); c.value = name; c.font = F({ bold: true }); c.alignment = { wrapText: true, horizontal: 'center' }; });
  let hc = ws.getCell(6, reCol); hc.value = 'Retained Earnings'; hc.font = F({ bold: true }); hc.alignment = { wrapText: true, horizontal: 'center' };
  hc = ws.getCell(6, totCol); hc.value = "Total Members' Equity"; hc.font = F({ bold: true }); hc.alignment = { wrapText: true, horizontal: 'center' };
  const num = (r, col, v) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(); c.numFmt = ACCT; };
  let row = 8;
  ws.getCell(row, 2).value = 'Beginning Balance 01/01/' + eq.years[0]; ws.getCell(row, 2).font = F({ bold: true });
  eq.memberNames.forEach((n, i) => num(row, memCols[i], 0)); num(row, reCol, 0); num(row, totCol, 0);
  row += 2;
  for (const y of eq.years) {
    const contribs = eq.memberNames.map((n) => (eq.act[y] && eq.act[y][n] && eq.act[y][n].contrib) || 0);
    if (contribs.some((v) => Math.abs(v) >= 0.005)) {
      ws.getCell(row, 3).value = 'Capital Contributed'; ws.getCell(row, 3).font = F();
      let tot = 0; eq.memberNames.forEach((n, i) => { num(row, memCols[i], contribs[i]); tot += contribs[i]; });
      num(row, totCol, tot); row++;
    }
    const ni = r2((eq.niByYear[y] || 0) + ((eq.reAct[y] && eq.reAct[y].direct) || 0));
    if (Math.abs(ni) >= 0.005) {
      ws.getCell(row, 3).value = 'Net Income (Loss)'; ws.getCell(row, 3).font = F();
      num(row, reCol, ni); num(row, totCol, ni); row++;
    }
    const distribs = eq.memberNames.map((n) => (eq.act[y] && eq.act[y][n] && eq.act[y][n].distrib) || 0);
    if (distribs.some((v) => Math.abs(v) >= 0.005)) {
      ws.getCell(row, 3).value = 'Distribution'; ws.getCell(row, 3).font = F();
      let tot = 0; eq.memberNames.forEach((n, i) => { num(row, memCols[i], distribs[i]); tot += distribs[i]; });
      num(row, totCol, tot); row++;
    }
    ws.getCell(row, 2).value = 'Balance at Closing 12/31/' + y; ws.getCell(row, 2).font = F({ bold: true });
    let tot = 0; eq.memberNames.forEach((n, i) => { const v = eq.memberEnd[n][y]; num(row, memCols[i], v); tot += v; });
    num(row, reCol, eq.reEnd[y]); tot += eq.reEnd[y]; num(row, totCol, tot);
    row += 2;
  }
  const lastY = eq.years[eq.years.length - 1];
  const closing = r2(eq.memberNames.reduce((s, n) => s + eq.memberEnd[n][lastY], 0) + eq.reEnd[lastY]);
  num(row, totCol, eq.membersEquityTB); ws.getCell(row, totCol + 1).value = "Members' Equity (TB)"; ws.getCell(row, totCol + 1).font = F();
  num(row + 1, totCol, r2(closing - eq.membersEquityTB)); ws.getCell(row + 1, totCol + 1).value = 'Variance'; ws.getCell(row + 1, totCol + 1).font = F();
  ws.getColumn(2).width = 24; ws.getColumn(3).width = 20;
  memCols.concat([reCol, totCol]).forEach((c) => { ws.getColumn(c).width = 16; });
  ws.getColumn(totCol + 1).width = 22;
}

function renderContributedCapital(wb, d) {
  const eq = d.equity; if (!eq || !eq.contribDetail.length) return;
  const ws = wb.addWorksheet('Contributed Capital');
  const num = (r, col, v, bold) => { const c = ws.getCell(r, col); c.value = r2(v); c.font = F(bold ? { bold: true } : {}); c.numFmt = ACCT; };
  let row = 1, grand = 0;
  for (const cd of eq.contribDetail) { ws.getCell(row, 1).value = cd.name; ws.getCell(row, 1).font = F(); num(row, 2, cd.total); grand = r2(grand + cd.total); row++; }
  ws.getCell(row, 1).value = 'Total'; ws.getCell(row, 1).font = F({ bold: true }); num(row, 2, grand, true); row += 2;
  ws.getCell(row, 1).value = 'Contributions by members thru 12/31/' + d.p.year; ws.getCell(row, 1).font = F({ bold: true }); row += 2;
  for (const cd of eq.contribDetail) {
    ws.getCell(row, 2).value = 'Contributions made by ' + cd.member + ':'; ws.getCell(row, 2).font = F({ bold: true });
    if (Math.abs(cd.pyTotal) >= 0.005) { num(row, 3, cd.pyTotal); ws.getCell(row, 4).value = 'PY'; ws.getCell(row, 4).font = F(); }
    row++;
    for (const l of cd.cur) {
      ws.getCell(row, 1).value = l.date; ws.getCell(row, 1).font = F();
      ws.getCell(row, 2).value = l.memo; ws.getCell(row, 2).font = F();
      num(row, 3, l.amt); row++;
    }
    num(row, 3, cd.total, true); row += 2;
  }
  ws.getCell(row, 2).value = 'Total contributions by members thru 12/31/' + d.p.year; ws.getCell(row, 2).font = F({ bold: true });
  num(row, 3, grand, true); row += 2;
  ws.getCell(row, 1).value = 'Contributed Capital per GL:'; ws.getCell(row, 1).font = F({ bold: true }); row++;
  for (const cd of eq.contribDetail) {
    ws.getCell(row, 1).value = Number(cd.code); ws.getCell(row, 1).font = F();
    ws.getCell(row, 2).value = cd.name; ws.getCell(row, 2).font = F();
    num(row, 3, -cd.total); ws.getCell(row, 4).value = 'Trial Balance'; ws.getCell(row, 4).font = F(); row++;
  }
  ws.getColumn(1).width = 18; ws.getColumn(2).width = 48; ws.getColumn(3).width = 16; ws.getColumn(4).width = 20;
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

module.exports = { computeEquity, renderEqRF, renderContributedCapital, renderFixedAssets };
