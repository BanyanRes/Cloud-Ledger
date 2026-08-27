import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/turnkey.js'
src = open(p, encoding='utf-8').read()

# Helper: each sync function's "lines: [...]" block needs project_id added.
# We'll do specific find/replace per function for clarity.

replacements = [
    # Event 1: sub approved
    (
      "lines: [\n      { account_code: map.cip_code, debit: payload.amount, credit: 0, description: 'CIP accrual - ' + payload.vendor_name },\n      { account_code: map.ap_sub_code, debit: 0, credit: payload.amount, description: 'AP - ' + payload.vendor_name },\n    ],\n  });\n\n  logSync(db, {\n    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_approved',",
      "lines: [\n      { account_code: map.cip_code, debit: payload.amount, credit: 0, description: 'CIP accrual - ' + payload.vendor_name, project_id: payload.turnkey_project_id },\n      { account_code: map.ap_sub_code, debit: 0, credit: payload.amount, description: 'AP - ' + payload.vendor_name, project_id: payload.turnkey_project_id },\n    ],\n  });\n\n  logSync(db, {\n    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_approved',"
    ),
    # Event 2: sub paid
    (
      "lines: [\n      { account_code: map.ap_sub_code, debit: payload.amount, credit: 0, description: 'AP clear - ' + payload.vendor_name },\n      { account_code: creditCode, debit: 0, credit: payload.amount, description: 'Cash/Clearing out - ' + payload.payment_method },\n    ],\n  });\n\n  logSync(db, {\n    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_paid',",
      "lines: [\n      { account_code: map.ap_sub_code, debit: payload.amount, credit: 0, description: 'AP clear - ' + payload.vendor_name, project_id: payload.turnkey_project_id },\n      { account_code: creditCode, debit: 0, credit: payload.amount, description: 'Cash/Clearing out - ' + payload.payment_method, project_id: payload.turnkey_project_id },\n    ],\n  });\n\n  logSync(db, {\n    cl_entity_id: map.cl_entity_id, sync_type: 'sub_payapp_paid',"
    ),
    # Event 3: owner issued
    (
      "lines: [\n      { account_code: map.ar_owner_code, debit: payload.amount, credit: 0, description: 'AR - Owner' },\n      { account_code: map.billings_uncompleted_code, debit: 0, credit: payload.amount, description: 'Billings on uncompleted contract' },\n    ],",
      "lines: [\n      { account_code: map.ar_owner_code, debit: payload.amount, credit: 0, description: 'AR - Owner', project_id: payload.turnkey_project_id },\n      { account_code: map.billings_uncompleted_code, debit: 0, credit: payload.amount, description: 'Billings on uncompleted contract', project_id: payload.turnkey_project_id },\n    ],"
    ),
    # Event 4: owner received
    (
      "lines: [\n      { account_code: map.cash_account_code, debit: payload.amount, credit: 0, description: 'Cash from owner' },\n      { account_code: map.ar_owner_code, debit: 0, credit: payload.amount, description: 'AR clear' },\n    ],",
      "lines: [\n      { account_code: map.cash_account_code, debit: payload.amount, credit: 0, description: 'Cash from owner', project_id: payload.turnkey_project_id },\n      { account_code: map.ar_owner_code, debit: 0, credit: payload.amount, description: 'AR clear', project_id: payload.turnkey_project_id },\n    ],"
    ),
]

for old, new in replacements:
    if old not in src:
        print('MISSING:', old[:80])
        sys.exit(1)
    src = src.replace(old, new, 1)

open(p, 'w', encoding='utf-8').write(src)
print('OK: all 4 sync handlers tagged with project_id')
