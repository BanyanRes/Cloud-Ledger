import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()

# Update postJE to write project_id into journal_lines
old_postJE = """  const entryId = result.lastInsertRowid;
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

new_postJE = """  const entryId = result.lastInsertRowid;
  const insertLine = db.prepare(
    'INSERT INTO journal_lines (entry_id, account_code, debit, credit, project_id) ' +
    'VALUES (?, ?, ?, ?, ?)'
  );
  for (var j = 0; j < lines.length; j++) {
    const l = lines[j];
    insertLine.run(
      entryId, l.account_code,
      Number(l.debit || 0), Number(l.credit || 0),
      l.project_id != null ? String(l.project_id) : null
    );
  }
  return entryId;
}"""

assert old_postJE in src, 'postJE pattern not found'
src = src.replace(old_postJE, new_postJE, 1)
open(p, 'w', encoding='utf-8').write(src)
print('OK postJE updated')
