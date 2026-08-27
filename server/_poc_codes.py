p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()
old = """const POC_ACCOUNTS = [
  { code: '10000', name: 'Cash',                                type: 'Asset' },
  { code: '10100', name: 'Bill.com Clearing',                   type: 'Asset' },
  { code: '12000', name: 'Accounts Receivable - Owner',         type: 'Asset' },
  { code: '12500', name: 'Costs in Excess of Billings',         type: 'Asset' },
  { code: '13000', name: 'Construction-in-Progress',            type: 'Asset' },
  { code: '21000', name: 'Accounts Payable - Subcontractors',   type: 'Liability' },
  { code: '23000', name: 'Billings on Uncompleted Contracts',   type: 'Liability' },
  { code: '24000', name: 'Billings in Excess of Costs',         type: 'Liability' },
  { code: '40000', name: 'Construction Revenue',                type: 'Revenue' },
  { code: '50000', name: 'Cost of Construction',                type: 'Expense' },
];"""
new = """// POC chart of accounts — picks 5-digit codes that don't collide with the
// CloudLedger DEFAULT_COA (which already takes 10000, 10100, 11000, 12000,
// 13000, 15000, 20000, 21000, 40000, 50000, etc.).
const POC_ACCOUNTS = [
  { code: '10000', name: 'Cash',                                type: 'Asset' },     // shared w/ default
  { code: '10500', name: 'Bill.com Clearing',                   type: 'Asset' },     // new (avoids 10100)
  { code: '11500', name: 'Accounts Receivable - Owner',         type: 'Asset' },     // new (avoids 11000)
  { code: '14500', name: 'Costs in Excess of Billings',         type: 'Asset' },     // new (avoids 12500)
  { code: '15500', name: 'Construction-in-Progress',            type: 'Asset' },     // new (avoids 13000)
  { code: '20500', name: 'Accounts Payable - Subcontractors',   type: 'Liability' }, // new (avoids 20000)
  { code: '23000', name: 'Billings on Uncompleted Contracts',   type: 'Liability' }, // new (free slot)
  { code: '24000', name: 'Billings in Excess of Costs',         type: 'Liability' }, // new (free slot)
  { code: '45000', name: 'Construction Revenue',                type: 'Revenue' },   // new (avoids 40000)
  { code: '55000', name: 'Cost of Construction',                type: 'Expense' },   // new (avoids 50000)
];"""
assert old in src
src = src.replace(old, new, 1)

# Also update the code references in linkProject's pick() calls
old_pick = """      pick('10000'), pick('10100'), pick('12000'),
      pick('12500'), pick('13000'), pick('21000'),
      pick('23000'), pick('24000'),
      pick('40000'), pick('50000'),"""
new_pick = """      pick('10000'), pick('10500'), pick('11500'),
      pick('14500'), pick('15500'), pick('20500'),
      pick('23000'), pick('24000'),
      pick('45000'), pick('55000'),"""
assert old_pick in src
src = src.replace(old_pick, new_pick, 1)

open(p, 'w', encoding='utf-8').write(src)
print('OK')
