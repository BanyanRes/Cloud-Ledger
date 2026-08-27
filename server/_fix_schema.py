p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()

# Rename journal_entry_lines -> journal_lines
src = src.replace('journal_entry_lines', 'journal_lines')

# Replace the postJE function entirely to match real schema
old = """// Journal entry helper (balanced)
function postJE(db, args) {
  const lines = args.lines;
  let totalDr = 0, totalCr = 0;
  for (var i = 0; i < lines.length; i++) {
    totalDr += Number(lines[i].debit || 0);
    totalCr += Number(lines[i].credit || 0);
  }
  if (Math.abs(totalDr - totalCr) > 0.005) {
    throw new Error('Unbalanced JE: Dr ' + totalDr + ' vs Cr ' + totalCr);
  }
  const now = new Date().toISOString();
  const result = db.prepare(
    'INSERT INTO journal_entries (entity_id, date, memo, reference, created_by, created_at) ' +
    'VALUES (?, ?, ?, ?, ?, ?)'
  ).run(
    args.cl_entity_id, args.date, args.memo || null, args.reference || null,
    args.created_by || 'turnkey-sync', now
  );
  const entryId = result.lastInsertRowid;
  const insertLine = db.prepare(
    'INSERT INTO journal_lines (entry_id, account_code, debit, credit, description) ' +
    'VALUES (?, ?, ?, ?, ?)'
  );
  for (var j = 0; j < lines.length; j++) {
    const l = lines[j];
    insertLine.run(entryId, l.account_code, Number(l.debit || 0), Number(l.credit || 0), l.description || null);
  }
  return entryId;
}"""

new = """// Journal entry helper (balanced).
// Schema: journal_entries(entity_id, entry_num, date, memo, created_by, created_at)
//         journal_lines(entry_id, account_code, debit, credit)
// memo is NOT NULL; entry_num auto-incremented per entity.
function postJE(db, args) {
  const lines = args.lines;
  let totalDr = 0, totalCr = 0;
  for (var i = 0; i < lines.length; i++) {
    totalDr += Number(lines[i].debit || 0);
    totalCr += Number(lines[i].credit || 0);
  }
  if (Math.abs(totalDr - totalCr) > 0.005) {
    throw new Error('Unbalanced JE: Dr ' + totalDr + ' vs Cr ' + totalCr);
  }
  // Next entry_num for this entity
  const nextRow = db.prepare(
    'SELECT COALESCE(MAX(entry_num), 0) + 1 AS n FROM journal_entries WHERE entity_id = ?'
  ).get(args.cl_entity_id);
  const entryNum = nextRow.n;

  // Build memo: prefix with reference if provided
  const memo = (args.reference ? '[' + args.reference + '] ' : '') + (args.memo || '');

  const now = new Date().toISOString();
  const result = db.prepare(
    'INSERT INTO journal_entries (entity_id, entry_num, date, memo, created_by, created_at) ' +
    'VALUES (?, ?, ?, ?, ?, ?)'
  ).run(
    args.cl_entity_id, entryNum, args.date, memo,
    args.created_by || 'turnkey-sync', now
  );
  const entryId = result.lastInsertRowid;
  const insertLine = db.prepare(
    'INSERT INTO journal_lines (entry_id, account_code, debit, credit) ' +
    'VALUES (?, ?, ?, ?)'
  );
  for (var j = 0; j < lines.length; j++) {
    const l = lines[j];
    insertLine.run(entryId, l.account_code, Number(l.debit || 0), Number(l.credit || 0));
  }
  return entryId;
}"""

assert old in src, 'postJE pattern not found'
src = src.replace(old, new, 1)

# Also: accounts table — check what's the actual column. The check earlier showed 'code' and 'name', 'type'
# That's fine. But our INSERT into accounts only uses 4 cols; check real schema.

open(p, 'w', encoding='utf-8').write(src)
print('OK')
