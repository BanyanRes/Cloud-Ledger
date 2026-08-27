// Scratch test for the AR module: accrual JE, numbering, receipts, aging tie-out,
// void reversal, PDF generation. In-memory DB, never committed.
// better-sqlite3's native binding is only built on Railway, so this test runs on
// node:sqlite with a thin shim exposing the better-sqlite3 surface the module uses.
const { DatabaseSync } = require('node:sqlite');
const ar = require('./ar');

const raw = new DatabaseSync(':memory:');
const db = {
  exec: (sql) => raw.exec(sql),
  prepare: (sql) => {
    const st = raw.prepare(sql);
    return {
      run: (...a) => st.run(...a),
      get: (...a) => st.get(...a),
      all: (...a) => st.all(...a),
    };
  },
  transaction: (fn) => (...a) => fn(...a),
};
db.exec(`
  CREATE TABLE entities (id INTEGER PRIMARY KEY, code TEXT, name TEXT, entity_type TEXT);
  CREATE TABLE accounts (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, code TEXT, name TEXT, type TEXT, bank_acct INTEGER DEFAULT 0);
  CREATE TABLE journal_entries (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, entry_num INTEGER, date TEXT, memo TEXT, created_by TEXT, created_at TEXT);
  CREATE TABLE journal_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, entry_id INTEGER, account_code TEXT, debit REAL DEFAULT 0, credit REAL DEFAULT 0, description TEXT, project_id INTEGER, class_id INTEGER, location_id INTEGER);
  CREATE TABLE entity_files (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, folder_path TEXT, stored_filename TEXT, original_name TEXT, size INTEGER, mime_type TEXT, uploaded_by TEXT, created_at TEXT);
  CREATE TABLE entity_folders (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, folder_path TEXT, created_by TEXT, created_at TEXT, UNIQUE(entity_id, folder_path));
  CREATE TABLE ar_customers (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, name TEXT, email TEXT, address TEXT, terms_days INTEGER DEFAULT 30, active INTEGER DEFAULT 1, created_at TEXT);
  CREATE TABLE ar_invoice_templates (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, customer_id INTEGER, memo TEXT, frequency TEXT, day_of_month INTEGER, next_run TEXT, ar_account_code TEXT, active INTEGER DEFAULT 1, created_at TEXT);
  CREATE TABLE ar_template_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, template_id INTEGER, description TEXT, qty REAL, rate REAL, revenue_account_code TEXT, class_id INTEGER, location_id INTEGER, sort INTEGER);
  CREATE TABLE ar_invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, entity_id INTEGER, customer_id INTEGER, template_id INTEGER, invoice_num TEXT, invoice_date TEXT, due_date TEXT, customer_name TEXT, customer_email TEXT, customer_address TEXT, memo TEXT, subtotal REAL, total REAL, ar_account_code TEXT, status TEXT DEFAULT 'draft', je_id INTEGER, pay_je_id INTEGER, pdf_file_id INTEGER, sent_at TEXT, paid_at TEXT, created_by TEXT, created_at TEXT, UNIQUE(entity_id, invoice_num));
  CREATE TABLE ar_invoice_lines (id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER, description TEXT, qty REAL, rate REAL, amount REAL, revenue_account_code TEXT, class_id INTEGER, location_id INTEGER, sort INTEGER);
`);
ar.ensureSchema(db);

const EID = 37;
db.prepare('INSERT INTO entities (id, code, name, entity_type) VALUES (?,?,?,?)').run(EID, 'SABINERI', 'County Line SRN', 'development');
const acct = db.prepare('INSERT INTO accounts (entity_id, code, name, type, bank_acct) VALUES (?,?,?,?,?)');
acct.run(EID, '10010', 'Operating Cash', 'Asset', 1);
acct.run(EID, '12000', 'Accounts Receivable', 'Asset', 0);
acct.run(EID, '12001', 'Accounts Receivable - Other', 'Asset', 0);
acct.run(EID, '40110', 'Land Leases', 'Revenue', 0);
acct.run(EID, '40130', 'Storage Fees', 'Revenue', 0);
db.prepare('INSERT INTO ar_customers (entity_id, name, email, address, terms_days) VALUES (?,?,?,?,?)')
  .run(EID, 'BNSF Railway', 'ap@bnsf.example', '2650 Lou Menk Dr\nFort Worth, TX 76131', 30);
const CUST = 1;

let pass = 0, fail = 0;
const check = (name, cond, extra) => {
  if (cond) { pass++; console.log('  PASS  ' + name); }
  else { fail++; console.log('  FAIL  ' + name + (extra ? '  -> ' + extra : '')); }
};
const glBal = (code, asOf) => db.prepare(
  'SELECT COALESCE(SUM(jl.debit - jl.credit),0) AS b FROM journal_lines jl JOIN journal_entries je ON je.id=jl.entry_id WHERE je.entity_id=? AND jl.account_code=?'
  + (asOf ? ' AND je.date <= ?' : '')).get(...(asOf ? [EID, code, asOf] : [EID, code])).b;

console.log('\n--- A/R account discovery ---');
check('picks 12000 over 12001 "- Other"', ar.defaultArAccount(db, EID) === '12000', ar.defaultArAccount(db, EID));

console.log('\n--- invoice creation + accrual JE ---');
const id1 = ar.createInvoice(db, EID, {
  customer_id: CUST, invoice_date: '2026-04-30', memo: 'April 2026 land lease + storage',
  lines: [
    { description: 'Land lease - Silsbee yard', qty: 1, rate: 12500, revenue_account_code: '40110' },
    { description: 'Storage fees - 43 cars @ 27.50', qty: 43, rate: 27.5, revenue_account_code: '40130' },
  ],
}, 'Jimmy Yun');
const inv1 = ar.invoiceWithLines(db, EID, id1);
check('invoice number INV-2026-0001', inv1.invoice_num === 'INV-2026-0001', inv1.invoice_num);
check('total = 12500 + 1182.50 = 13682.50', inv1.total === 13682.5, inv1.total);
check('due date = invoice date + Net 30', inv1.due_date === '2026-05-30', inv1.due_date);
check('status draft', inv1.status === 'draft', inv1.status);
check('JE posted and linked', !!inv1.je_id);
const je1 = db.prepare('SELECT * FROM journal_lines WHERE entry_id = ?').all(inv1.je_id);
check('JE has 3 lines (1 A/R + 2 revenue)', je1.length === 3, je1.length);
check('JE balances', Math.abs(je1.reduce((s, l) => s + l.debit - l.credit, 0)) < 0.005);
check('A/R debited 13682.50', glBal('12000') === 13682.5, glBal('12000'));
check('40110 credited 12500', glBal('40110') === -12500, glBal('40110'));
check('40130 credited 1182.50', glBal('40130') === -1182.5, glBal('40130'));

console.log('\n--- numbering increments, per year ---');
const id2 = ar.createInvoice(db, EID, {
  customer_id: CUST, invoice_date: '2026-05-31', memo: 'May 2026',
  lines: [{ description: 'Land lease - Silsbee yard', qty: 1, rate: 12500, revenue_account_code: '40110' }],
}, 'Jimmy Yun');
check('second invoice is 0002', ar.invoiceWithLines(db, EID, id2).invoice_num === 'INV-2026-0002');
check('2027 restarts at 0001', ar.nextInvoiceNum(db, EID, '2027-01-15') === 'INV-2027-0001', ar.nextInvoiceNum(db, EID, '2027-01-15'));

console.log('\n--- rounding: 3 x 33.333 ---');
const id3 = ar.createInvoice(db, EID, {
  customer_id: CUST, invoice_date: '2026-06-01',
  lines: [{ description: 'Odd rate', qty: 3, rate: 33.333, revenue_account_code: '40130' }],
}, 'Jimmy Yun');
const inv3 = ar.invoiceWithLines(db, EID, id3);
check('amount rounds to 100.00', inv3.total === 100, inv3.total);
const je3 = db.prepare('SELECT * FROM journal_lines WHERE entry_id = ?').all(inv3.je_id);
check('rounded JE still balances', Math.abs(je3.reduce((s, l) => s + l.debit - l.credit, 0)) < 0.005);

console.log('\n--- cash receipts (partial then full) ---');
const receipt = (invId, date, amount) => {
  const inv = ar.invoiceWithLines(db, EID, invId);
  const je = ar.postJE(db, EID, date, 'AR Receipt - ' + inv.invoice_num, [
    { account_code: '10010', debit: amount, credit: 0 },
    { account_code: inv.ar_account_code, debit: 0, credit: amount },
  ], 'test');
  db.prepare('INSERT INTO ar_receipts (invoice_id, entity_id, date, amount, bank_account_code, je_id) VALUES (?,?,?,?,?,?)')
    .run(invId, EID, date, amount, '10010', je.id);
  const paid = db.prepare('SELECT COALESCE(SUM(amount),0) AS p FROM ar_receipts WHERE invoice_id=?').get(invId).p;
  if (paid >= inv.total - 0.005) db.prepare("UPDATE ar_invoices SET status='paid', paid_at=? WHERE id=?").run(date, invId);
};
receipt(id1, '2026-06-15', 5000);
let a = ar.invoiceWithLines(db, EID, id1);
check('partial payment leaves 8682.50 open', a.open_amount === 8682.5, a.open_amount);
check('still not marked paid', a.status === 'draft', a.status);
receipt(id1, '2026-07-01', 8682.5);
a = ar.invoiceWithLines(db, EID, id1);
check('fully paid -> status paid', a.status === 'paid', a.status);
check('open amount zero', a.open_amount === 0, a.open_amount);

console.log('\n--- aging + GL tie-out ---');
const ag = ar.buildAging(db, EID, '2026-07-15');
check('aging ties to GL A/R (recon_diff = 0)', ag.recon_diff === 0, 'gl=' + ag.gl_ar_balance + ' aging=' + ag.totals.total);
check('open total = 12500 (May) + 100 (Jun)', ag.totals.total === 12600, ag.totals.total);
// May 31 invoice + Net 30 = due Jun 30; at Jul 15 that is 15 days late.
const may = ag.detail.find(d => d.invoice_num === 'INV-2026-0002');
check('May invoice 15 days past due -> 1-30 bucket', may.bucket === 'd1_30' && may.days_past_due === 15, may.bucket + '/' + may.days_past_due);
const agLate = ar.buildAging(db, EID, '2026-08-15');
const mayLate = agLate.detail.find(d => d.invoice_num === 'INV-2026-0002');
check('same invoice at 8/15 is 46 days late -> 31-60 bucket', mayLate.bucket === 'd31_60' && mayLate.days_past_due === 46, mayLate.bucket + '/' + mayLate.days_past_due);
const agVeryLate = ar.buildAging(db, EID, '2026-10-15');
check('at 10/15 it lands in 90+ bucket', agVeryLate.detail.find(d => d.invoice_num === 'INV-2026-0002').bucket === 'd90_plus');
const jun = ag.detail.find(d => d.invoice_num === 'INV-2026-0003');
check('June invoice due 7/1, 14 days late -> 1-30 bucket', jun.bucket === 'd1_30' && jun.days_past_due === 14, jun.bucket + '/' + jun.days_past_due);
check('paid invoice excluded from aging', !ag.detail.some(d => d.invoice_num === 'INV-2026-0001'));

const agEarly = ar.buildAging(db, EID, '2026-05-15');
check('as-of 5/15 sees only April invoice, still current', agEarly.totals.total === 13682.5 && agEarly.totals.current === 13682.5, JSON.stringify(agEarly.totals));
check('as-of 5/15 ties to GL at that date', agEarly.recon_diff === 0, 'gl=' + agEarly.gl_ar_balance);

console.log('\n--- legacy / imported A/R in the GL column (not aged) ---');
// Simulate an Intacct import: a manual JE debiting 12000 that the AR module does
// not own. It must land in the GL column un-aged, keep the buckets unchanged, and
// leave recon_diff at 0. Dr 12000 / Cr 40110, dated before the as-of.
ar.postJE(db, EID, '2026-04-30', 'GL detail import (2026-04-30)', [
  { account_code: '12000', debit: 50000, credit: 0, description: 'Invoice - Legacy Customer' },
  { account_code: '40110', debit: 0, credit: 50000, description: 'Invoice - Legacy Customer' },
], 'import');
const agGL = ar.buildAging(db, EID, '2026-07-15');
const agedBucketSum = agGL.totals.current + agGL.totals.d1_30 + agGL.totals.d31_60 + agGL.totals.d61_90 + agGL.totals.d90_plus;
check('legacy JE appears in GL column', agGL.gl_total === 50000 && agGL.totals.gl === 50000, 'gl_total=' + agGL.gl_total);
check('legacy JE is listed as a GL row', agGL.gl_rows.length === 1 && agGL.gl_rows[0].amount === 50000, JSON.stringify(agGL.gl_rows));
check('legacy JE does NOT enter any aged bucket', agedBucketSum === 12600, 'agedBucketSum=' + agedBucketSum);
check('grand total = aged 12600 + GL 50000', agGL.totals.total === 62600, agGL.totals.total);
check('still ties to GL 12000 with legacy present', agGL.recon_diff === 0, 'gl=' + agGL.gl_ar_balance + ' total=' + agGL.totals.total);
// A receipt/accrual JE the module DOES own must never be pulled into the GL column.
check('module-owned JEs stay out of GL column', agGL.gl_total === 50000, 'gl_total=' + agGL.gl_total);

console.log('\n--- void reversal (issued invoice) ---');
db.prepare("UPDATE ar_invoices SET status='sent', sent_at=datetime('now') WHERE id=?").run(id3);
const before = glBal('12000');
const inv3b = ar.invoiceWithLines(db, EID, id3);
ar.postJE(db, EID, '2026-07-20', 'Void AR Invoice ' + inv3b.invoice_num, [
  { account_code: inv3b.ar_account_code, debit: 0, credit: inv3b.total },
  ...inv3b.lines.map(l => ({ account_code: l.revenue_account_code, debit: l.amount, credit: 0 })),
], 'test');
db.prepare("UPDATE ar_invoices SET status='void' WHERE id=?").run(id3);
check('reversal removes 100.00 from A/R', glBal('12000') === before - 100, glBal('12000'));
const agAfter = ar.buildAging(db, EID, '2026-07-20');
// Aged buckets now hold only the May invoice (12500). The void JE posted above is
// a manual entry the module does not own (no void_je_id link), so its -100 lands
// and net to zero in the GL column; only the 50000 legacy import remains: total = 62500.
// The whole report still ties to the 12000 balance either way.
const agedAfter = agAfter.totals.current + agAfter.totals.d1_30 + agAfter.totals.d31_60 + agAfter.totals.d61_90 + agAfter.totals.d90_plus;
check('void excluded from aging and still ties', agAfter.recon_diff === 0 && agedAfter === 12500 && agAfter.totals.gl === 50000 && agAfter.totals.total === 62500, 'diff=' + agAfter.recon_diff + ' aged=' + agedAfter + ' gl=' + agAfter.totals.gl + ' total=' + agAfter.totals.total);

console.log('\n--- recurrence roll-forward ---');
check('monthly 2026-01-31 -> 2026-02-28 (clamped)', ar.advanceNextRun('2026-01-31', 'monthly', 31) === '2026-02-28', ar.advanceNextRun('2026-01-31', 'monthly', 31));
check('monthly 2026-04-01 -> 2026-05-01', ar.advanceNextRun('2026-04-01', 'monthly', 1) === '2026-05-01');
check('quarterly 2026-04-01 -> 2026-07-01', ar.advanceNextRun('2026-04-01', 'quarterly', 1) === '2026-07-01');
check('annual 2026-04-01 -> 2027-04-01', ar.advanceNextRun('2026-04-01', 'annual', 1) === '2027-04-01');
check('dec monthly rolls the year', ar.advanceNextRun('2026-12-15', 'monthly', 15) === '2027-01-15', ar.advanceNextRun('2026-12-15', 'monthly', 15));

console.log('\n--- validation guards ---');
const expectThrow = (name, fn) => { try { fn(); check(name, false, 'no error thrown'); } catch (e) { check(name, true); } };
expectThrow('rejects unknown revenue account', () => ar.createInvoice(db, EID, { customer_id: CUST, invoice_date: '2026-06-01', lines: [{ description: 'x', qty: 1, rate: 5, revenue_account_code: '99999' }] }, 't'));
expectThrow('rejects empty line list', () => ar.createInvoice(db, EID, { customer_id: CUST, invoice_date: '2026-06-01', lines: [] }, 't'));
expectThrow('rejects zero-amount line', () => ar.createInvoice(db, EID, { customer_id: CUST, invoice_date: '2026-06-01', lines: [{ description: 'x', qty: 0, rate: 5, revenue_account_code: '40130' }] }, 't'));
expectThrow('rejects wrong-entity customer', () => ar.createInvoice(db, EID, { customer_id: 999, invoice_date: '2026-06-01', lines: [{ description: 'x', qty: 1, rate: 5, revenue_account_code: '40130' }] }, 't'));

console.log('\n--- opening subledger import + residual GL + external cash receipt ---');
const glBalE = (eid, code, asOf) => db.prepare('SELECT COALESCE(SUM(jl.debit - jl.credit),0) AS b FROM journal_lines jl JOIN journal_entries je ON je.id=jl.entry_id WHERE je.entity_id=? AND jl.account_code=?' + (asOf ? ' AND je.date <= ?' : '')).get(...(asOf ? [eid, code, asOf] : [eid, code])).b;
const EID2 = 41;
db.prepare('INSERT INTO entities (id, code, name, entity_type) VALUES (?,?,?,?)').run(EID2, 'BANYANRES', 'Banyan Residential', 'operating');
const a2 = db.prepare('INSERT INTO accounts (entity_id, code, name, type, bank_acct) VALUES (?,?,?,?,?)');
a2.run(EID2, '10010', 'Operating Cash', 'Asset', 1);
a2.run(EID2, '12000', 'Accounts Receivable', 'Asset', 0);
a2.run(EID2, '40110', 'Revenue', 'Revenue', 0);
const OPEN_TOTAL = 66193.49 + 72320.12 + 1029.41 + 1127.24 + (-639.25); // = 140031.01, includes a credit memo
ar.postJE(db, EID2, '2026-04-30', 'Legacy A/R opening balance import', [
  { account_code: '12000', debit: OPEN_TOTAL, credit: 0 },
  { account_code: '40110', debit: 0, credit: OPEN_TOTAL },
], 'import');
check('legacy control balance = 140031.01', glBalE(EID2, '12000') === 140031.01, glBalE(EID2, '12000'));
const items = [
  { customer_name: 'HP Property Owner, LLC', document_no: 'DEV-022026-HP', invoice_date: '2026-02-28', due_date: '2026-02-28', amount: 66193.49 },
  { customer_name: 'Milhaus QOZ Business V, LLC', document_no: 'Dev-032026b-Milhaus- INV 0032', invoice_date: '2026-03-01', due_date: '2026-03-01', amount: 72320.12 },
  { customer_name: 'Scottsdale Entrada I', document_no: 'CC-1125-032026-ENTRADA I', invoice_date: '2026-04-01', due_date: '2026-04-01', amount: 1029.41 },
  { customer_name: 'Scottsdale Entrada II', document_no: 'CC-1125-032026-ENTRADA I', invoice_date: '2026-04-01', due_date: '2026-04-01', amount: 1127.24 },
  { customer_name: 'Pflugerville Property Owner, LLC', document_no: 'CM-00044', invoice_date: '2026-04-01', due_date: '2026-04-01', amount: -639.25 },
];
const imp = ar.importOpeningItems(db, EID2, items, { who: 'test' });
check('import inserted 5 opening items', imp.inserted === 5, imp.inserted);
check('import total ties to legacy GL balance', imp.total === 140031.01, imp.total);
const opens = db.prepare("SELECT invoice_num FROM ar_invoices WHERE entity_id=? AND origin='opening' ORDER BY id").all(EID2).map(r => r.invoice_num);
check('duplicate document number is disambiguated', opens.includes('CC-1125-032026-ENTRADA I') && opens.includes('CC-1125-032026-ENTRADA I-2'), opens.join(' | '));
const agO = ar.buildAging(db, EID2, '2026-04-30');
const bO = agO.totals.current + agO.totals.d1_30 + agO.totals.d31_60 + agO.totals.d61_90 + agO.totals.d90_plus;
check('opening items age into buckets = 140031.01', bO === 140031.01, bO);
check('HP (due 2/28) lands in 61-90 as of 4/30', agO.detail.find(d => d.invoice_num === 'DEV-022026-HP').bucket === 'd61_90');
check('credit memo nets negative in 1-30', agO.detail.find(d => d.invoice_num === 'CM-00044').open === -639.25, agO.detail.find(d => d.invoice_num === 'CM-00044').open);
check('GL residual column is zero (fully itemized)', agO.totals.gl === 0, agO.totals.gl);
check('opening aging ties to GL (recon_diff = 0)', agO.recon_diff === 0, 'diff=' + agO.recon_diff + ' gl=' + agO.gl_ar_balance + ' total=' + agO.totals.total);

// External cash application: a bank deposit JE (Dr cash / Cr 12000) + subledger
// receipt, the way a posted bank transaction records it — NO second JE.
const hp = db.prepare("SELECT id, total FROM ar_invoices WHERE entity_id=? AND invoice_num='DEV-022026-HP'").get(EID2);
const dje = ar.postJE(db, EID2, '2026-05-05', 'Deposit - HighPoint ACH', [
  { account_code: '10010', debit: 66193.49, credit: 0 },
  { account_code: '12000', debit: 0, credit: 66193.49 },
], 'bank');
const rr = ar.recordArReceipt(db, { entity_id: EID2, invoice_id: hp.id, date: '2026-05-05', amount: 66193.49, bank_account_code: '10010', je_id: dje.id, created_by: 'bank' });
check('external receipt clears HP invoice (open = 0)', rr.open === 0, rr.open);
check('HP invoice flips to paid', db.prepare('SELECT status FROM ar_invoices WHERE id=?').get(hp.id).status === 'paid');
const agR = ar.buildAging(db, EID2, '2026-05-05');
check('paid opening item leaves the aging', !agR.detail.some(d => d.invoice_num === 'DEV-022026-HP'));
const bR = Math.round((agR.totals.current + agR.totals.d1_30 + agR.totals.d31_60 + agR.totals.d61_90 + agR.totals.d90_plus) * 100) / 100;
check('buckets drop by the receipt to 73837.52', bR === 73837.52, bR);
check('control balance dropped by the receipt', glBalE(EID2, '12000') === 73837.52, glBalE(EID2, '12000'));
check('aging still ties after external receipt', agR.recon_diff === 0 && agR.totals.gl === 0, 'diff=' + agR.recon_diff + ' gl=' + agR.totals.gl);
expectThrow('receipt exceeding open balance is rejected', () => ar.recordArReceipt(db, { entity_id: EID2, invoice_id: hp.id, date: '2026-05-06', amount: 100, bank_account_code: '10010' }));
expectThrow('re-import blocked when receipts exist (needs force)', () => ar.importOpeningItems(db, EID2, items, { who: 'test' }));
const imp2 = ar.importOpeningItems(db, EID2, items, { who: 'test', force: true });
check('force re-import replaces the opening items', imp2.inserted === 5 && imp2.total === 140031.01, imp2.inserted + '/' + imp2.total);
check('force re-import cleared prior receipts (no orphans)', db.prepare("SELECT COUNT(*) c FROM ar_receipts WHERE entity_id=?").get(EID2).c === 0, db.prepare("SELECT COUNT(*) c FROM ar_receipts WHERE entity_id=?").get(EID2).c);

console.log('\n--- aging basis: GL posting date for opening items ---');
const EID3 = 39;
db.prepare('INSERT INTO entities (id, code, name, entity_type) VALUES (?,?,?,?)').run(EID3, 'SILSBEE', 'CLR Silsbee', 'operating');
const a3 = db.prepare('INSERT INTO accounts (entity_id, code, name, type, bank_acct) VALUES (?,?,?,?,?)');
a3.run(EID3, '12000', 'Accounts Receivable', 'Asset', 0);
a3.run(EID3, '40110', 'Revenue', 'Revenue', 0);
ar.postJE(db, EID3, '2026-04-30', 'Legacy opening balance', [{ account_code: '12000', debit: 922.71, credit: 0 }, { account_code: '40110', debit: 0, credit: 922.71 }], 'import');
ar.importOpeningItems(db, EID3, [{ customer_name: 'Mountain Banyan', document_no: 'AM-042026-WASH-Banyan', posting_date: '2026-04-01', invoice_date: '2026-04-30', due_date: '2026-04-30', amount: 922.71 }], { who: 'test' });
const agP = ar.buildAging(db, EID3, '2026-04-30');
const mb = agP.detail.find(d => d.invoice_num === 'AM-042026-WASH-Banyan');
check('opening item ages by GL posting date (4/1 = 29 days = 1-30), not due date', mb.bucket === 'd1_30' && mb.days_past_due === 29, mb.bucket + '/' + mb.days_past_due);
check('display due date is preserved at 4/30', mb.due_date === '2026-04-30', mb.due_date);
check('posting-date opening still ties to GL', agP.recon_diff === 0 && agP.totals.total === 922.71, 'diff=' + agP.recon_diff);

// An aging file with no invoice-date column: invoice_date falls back to the due
// date, which can be AFTER the report date even though the item posted before it.
// The item must still age (by posting date), not vanish into the residual.
console.log('\n--- opening item due after the as-of date still ages ---');
const EID5 = 88;
db.prepare('INSERT INTO entities (id, code, name, entity_type) VALUES (?,?,?,?)').run(EID5, 'SRN', 'Sabine River & Northern', 'operating');
const a5 = db.prepare('INSERT INTO accounts (entity_id, code, name, type, bank_acct) VALUES (?,?,?,?,?)');
a5.run(EID5, '12000', 'Accounts Receivable', 'Asset', 0);
a5.run(EID5, '40110', 'Revenue', 'Revenue', 0);
ar.postJE(db, EID5, '2026-06-30', 'Legacy A/R opening balance', [
  { account_code: '12000', debit: 92000, credit: 0 },
  { account_code: '40110', debit: 0, credit: 92000 },
], 'import');
const futureDue = [
  // posted 5/31, due 6/30 -> 30 days aged as of 6/30
  { customer_name: 'CP Chem', document_no: '329', posting_date: '2026-05-31', invoice_date: null, due_date: '2026-06-30', amount: 89825 },
  // posted 6/30 but NOT due until 7/30 -> current, and must not be dropped
  { customer_name: 'CP Chem', document_no: '336', posting_date: '2026-06-30', invoice_date: null, due_date: '2026-07-30', amount: 2175 },
];
ar.importOpeningItems(db, EID5, futureDue, { who: 'test', as_of: '2026-06-30' });
const agF = ar.buildAging(db, EID5, '2026-06-30');
check('item due after the as-of date is still on the aging', agF.detail.length === 2, agF.detail.length);
check('it lands in current, not dropped', agF.totals.current === 2175, agF.totals.current);
check('the other ages 30 days into 1-30', agF.totals.d1_30 === 89825, agF.totals.d1_30);
check('buckets total the whole subledger', agF.totals.current + agF.totals.d1_30 === 92000);
check('no phantom residual and a clean tie-out',
  agF.opening_residual === 0 && agF.recon_diff === 0, agF.opening_residual + '/' + agF.recon_diff);
// an item posted AFTER the as-of date is correctly excluded
const agBefore = ar.buildAging(db, EID5, '2026-06-01');
check('item posted after the as-of date is excluded', agBefore.detail.length === 1, agBefore.detail.length);

console.log('\n--- wrong-entity import: overclaim detection ---');
const EID4 = 46;
db.prepare('INSERT INTO entities (id, code, name, entity_type) VALUES (?,?,?,?)').run(EID4, 'SABINE2', 'Sabine stand-in', 'operating');
const a4 = db.prepare('INSERT INTO accounts (entity_id, code, name, type, bank_acct) VALUES (?,?,?,?,?)');
a4.run(EID4, '12000', 'Accounts Receivable', 'Asset', 0);
a4.run(EID4, '40110', 'Revenue', 'Revenue', 0);
// the entity's own legacy A/R control balance
ar.postJE(db, EID4, '2026-06-30', 'Legacy A/R opening balance', [
  { account_code: '12000', debit: 213676.86, credit: 0 },
  { account_code: '40110', debit: 0, credit: 213676.86 },
], 'import');
// ANOTHER entity's aging detail: totals far more than this control account holds
const foreign = [
  { customer_name: 'CLIP', document_no: 'CLRO-0626-90', posting_date: '2026-06-30', invoice_date: '2026-06-30', due_date: '2026-06-30', amount: 178574.23 },
  { customer_name: 'CLIP', document_no: 'CLRO-0626-94', posting_date: '2026-06-30', invoice_date: '2026-06-30', due_date: '2026-06-30', amount: 87729.14 },
  { customer_name: 'SRN',  document_no: 'CLRO-0626-92', posting_date: '2026-06-30', invoice_date: '2026-06-30', due_date: '2026-06-30', amount: 22783.99 },
  { customer_name: 'Buna', document_no: 'CLRO-0626-93', posting_date: '2026-06-30', invoice_date: '2026-06-30', due_date: '2026-06-30', amount: 9493.48 },
];
const fTotal = 298580.84;
expectThrow('import blocked when the detail exceeds the GL control balance',
  () => ar.importOpeningItems(db, EID4, foreign, { who: 'test', as_of: '2026-06-30' }));
check('blocked import left the subledger untouched', db.prepare('SELECT COUNT(*) c FROM ar_invoices WHERE entity_id=?').get(EID4).c === 0);
// no as_of supplied -> the guard falls back to the PEAK balance the control
// account has ever held, so it still fires. A caller that omits the date is the
// least likely to have checked, and must not be handed an exemption.
expectThrow('guard still fires with no as_of (peak-balance basis)',
  () => ar.importOpeningItems(db, EID4, foreign, { who: 'test' }));
check('blocked no-as_of import left the subledger untouched', db.prepare('SELECT COUNT(*) c FROM ar_invoices WHERE entity_id=?').get(EID4).c === 0);
// this is what actually shipped to SRN before the guard existed — force it in to
// reproduce the broken state the aging report then has to expose.
const impNoAsOf = ar.importOpeningItems(db, EID4, foreign, { who: 'test', allow_over_gl: true });
check('override works without an as_of too', impNoAsOf.inserted === 4);
// this is what actually shipped to SRN before the guard existed
const agBad = ar.buildAging(db, EID4, '2026-06-30');
check('overclaiming subledger is NOT absorbed into the GL column', agBad.totals.gl === 0, agBad.totals.gl);
check('aging totals what the subledger actually says', agBad.totals.total === fTotal, agBad.totals.total);
check('recon_diff reports the overclaim instead of a false 0',
  agBad.recon_diff === Math.round((213676.86 - fTotal) * 100) / 100, agBad.recon_diff);
check('overclaim row is not presented as an un-itemized GL balance', agBad.gl_rows.length === 0, agBad.gl_rows.length);
check('opening_overclaim carries the overclaim',
  agBad.opening_overclaim === Math.round((fTotal - 213676.86) * 100) / 100, agBad.opening_overclaim);
check('opening_residual is zero when overclaiming', agBad.opening_residual === 0, agBad.opening_residual);
// explicit override still allowed, for a real reason
const impOver = ar.importOpeningItems(db, EID4, foreign, { who: 'test', as_of: '2026-06-30', allow_over_gl: true });
check('allow_over_gl overrides the guard', impOver.inserted === 4 && impOver.total === fTotal, impOver.inserted + '/' + impOver.total);
// a legitimate PARTIAL itemization is under the control balance and still ties
const impPart = ar.importOpeningItems(db, EID4, [foreign[3]], { who: 'test', as_of: '2026-06-30' });
const agPart = ar.buildAging(db, EID4, '2026-06-30');
check('partial itemization imports', impPart.inserted === 1);
check('positive residual still shows as un-itemized GL balance',
  agPart.totals.gl === Math.round((213676.86 - 9493.48) * 100) / 100, agPart.totals.gl);
check('partial itemization still ties to GL', agPart.recon_diff === 0, agPart.recon_diff);
// recon_diff is 0 here only because the residual is added back. The report has to
// be able to tell "ties AND complete" from "ties BUT partly un-itemized", so the
// residual is surfaced separately instead of hiding behind a green tie-out.
check('opening_residual carries the un-itemized remainder',
  agPart.opening_residual === Math.round((213676.86 - 9493.48) * 100) / 100, agPart.opening_residual);
check('opening_overclaim is zero on a partial itemization', agPart.opening_overclaim === 0, agPart.opening_overclaim);
// a fully-itemized subledger reports neither
const agFull = ar.buildAging(db, EID2, '2026-04-30');
check('fully-itemized opening reports no residual and no overclaim',
  agFull.opening_residual === 0 && agFull.opening_overclaim === 0,
  agFull.opening_residual + '/' + agFull.opening_overclaim);
console.log('\n--- PDF generation ---');
(async () => {
  const pdf = await ar.buildInvoicePdf({
    entityName: 'County Line SRN',
    settings: { bill_from: 'County Line SRN, LLC\n123 Main St\nSilsbee, TX', remit_to: 'Wire: Sunflower Bank\nACH 000123456', footer_note: 'Payment due within 30 days. Late balances accrue 1.5% monthly.' },
    invoice: ar.invoiceWithLines(db, EID, id1),
  });
  check('PDF is a valid non-trivial buffer', pdf.length > 1500 && pdf.slice(0, 5).toString() === '%PDF-', pdf.length + ' bytes');
  const multi = await ar.buildInvoicePdf({
    entityName: 'County Line SRN', settings: {},
    invoice: { invoice_num: 'INV-2026-9999', invoice_date: '2026-07-01', due_date: '2026-07-31', customer_name: 'Test', total: 400, paid_amount: 0, open_amount: 400,
      lines: Array.from({ length: 60 }, (_, i) => ({ description: 'Line item ' + i + ' with a deliberately long description that must wrap across the available column width', qty: 1, rate: 6.6667, amount: 6.67 })) },
  });
  check('long invoice paginates', multi.length > 3000, multi.length + ' bytes');

  console.log('\n=========================================');
  console.log('  ' + pass + ' passed, ' + fail + ' failed');
  console.log('=========================================\n');
  process.exit(fail ? 1 : 0);
})();
