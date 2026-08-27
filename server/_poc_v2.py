import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()

# Update syncMonthEndPOC: getBal() should filter by project_id, and JE lines must tag project_id
old = """  function getBal(code) {
    const row = db.prepare(
      'SELECT COALESCE(SUM(l.debit), 0) - COALESCE(SUM(l.credit), 0) AS bal ' +
      'FROM journal_lines l ' +
      'JOIN journal_entries e ON e.id = l.entry_id ' +
      'WHERE e.entity_id = ? AND l.account_code = ? AND e.date <= ?'
    ).get(map.cl_entity_id, code, payload.period_end_date);
    return Number(row.bal || 0);
  }"""

new = """  function getBal(code) {
    const row = db.prepare(
      'SELECT COALESCE(SUM(l.debit), 0) - COALESCE(SUM(l.credit), 0) AS bal ' +
      'FROM journal_lines l ' +
      'JOIN journal_entries e ON e.id = l.entry_id ' +
      'WHERE e.entity_id = ? AND l.account_code = ? AND l.project_id = ? AND e.date <= ?'
    ).get(map.cl_entity_id, code, String(payload.turnkey_project_id), payload.period_end_date);
    return Number(row.bal || 0);
  }"""

assert old in src, 'getBal pattern not found'
src = src.replace(old, new, 1)

# Tag the POC JE lines with project_id
old_lines_block = """  const lines = [];
  if (recognizedCost > 0.005) {
    lines.push({ account_code: map.cost_of_construction_code, debit: recognizedCost, credit: 0, description: 'POC cost recognition' });
    lines.push({ account_code: map.cip_code, debit: 0, credit: recognizedCost, description: 'CIP relief' });
  }
  if (earnedRevenue > 0.005) {
    lines.push({ account_code: map.billings_uncompleted_code, debit: earnedRevenue, credit: 0, description: 'Reclass billings to revenue' });
    lines.push({ account_code: map.revenue_code, debit: 0, credit: earnedRevenue, description: 'POC revenue recognition' });
  }
  const diff = Math.round((earnedRevenue - billingsBalance) * 100) / 100;
  if (Math.abs(diff) > 0.005) {
    if (diff > 0) {
      lines.push({ account_code: map.costs_in_excess_code, debit: diff, credit: 0, description: 'Under-billing (costs in excess)' });
      lines.push({ account_code: map.billings_uncompleted_code, debit: 0, credit: diff, description: 'Under-billing offset' });
    } else {
      lines.push({ account_code: map.billings_uncompleted_code, debit: -diff, credit: 0, description: 'Over-billing offset' });
      lines.push({ account_code: map.billings_in_excess_code, debit: 0, credit: -diff, description: 'Over-billing (billings in excess)' });
    }
  }"""

new_lines_block = """  const pid = payload.turnkey_project_id;
  const lines = [];
  if (recognizedCost > 0.005) {
    lines.push({ account_code: map.cost_of_construction_code, debit: recognizedCost, credit: 0, description: 'POC cost recognition', project_id: pid });
    lines.push({ account_code: map.cip_code, debit: 0, credit: recognizedCost, description: 'CIP relief', project_id: pid });
  }
  if (earnedRevenue > 0.005) {
    lines.push({ account_code: map.billings_uncompleted_code, debit: earnedRevenue, credit: 0, description: 'Reclass billings to revenue', project_id: pid });
    lines.push({ account_code: map.revenue_code, debit: 0, credit: earnedRevenue, description: 'POC revenue recognition', project_id: pid });
  }
  const diff = Math.round((earnedRevenue - billingsBalance) * 100) / 100;
  if (Math.abs(diff) > 0.005) {
    if (diff > 0) {
      lines.push({ account_code: map.costs_in_excess_code, debit: diff, credit: 0, description: 'Under-billing (costs in excess)', project_id: pid });
      lines.push({ account_code: map.billings_uncompleted_code, debit: 0, credit: diff, description: 'Under-billing offset', project_id: pid });
    } else {
      lines.push({ account_code: map.billings_uncompleted_code, debit: -diff, credit: 0, description: 'Over-billing offset', project_id: pid });
      lines.push({ account_code: map.billings_in_excess_code, debit: 0, credit: -diff, description: 'Over-billing (billings in excess)', project_id: pid });
    }
  }"""

assert old_lines_block in src
src = src.replace(old_lines_block, new_lines_block, 1)

open(p, 'w', encoding='utf-8').write(src)
print('OK POC handler updated for project_id dimension')
