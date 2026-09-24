// ─── A/P aging built from the GL ────────────────────────────────────────────
// Shared by GET /api/billcom/ap-aging/:entity_id (index.js) and the Month-End
// Leadsheets workpaper (leadsheets.js). Read-only. apAccountOverride forces the
// control account (else the entity's Bill.com default, else 202000).
function buildApAging(db, entityId, asOf, apAccountOverride) {
  const cfg = db.prepare('SELECT * FROM billcom_config WHERE entity_id = ?').get(entityId);
  const apAccount = apAccountOverride ? String(apAccountOverride)
    : (cfg && cfg.default_ap_account) ? String(cfg.default_ap_account) : '202000';

  // ── 1. Pull all GL activity on the AP account through the as-of date. This is
  //    the authoritative AP record: credits = bills, debits = payments/relief.
  //    The report is built from here so it ALWAYS ties to the GL balance.
  const glLines = db.prepare(
    `SELECT jl.id AS line_id, je.id AS entry_id, je.entry_num, je.date, je.memo, je.vendor,
            jl.debit, jl.credit, jl.description
       FROM journal_lines jl JOIN journal_entries je ON jl.entry_id = je.id
      WHERE je.entity_id = ? AND jl.account_code = ? AND je.date <= ?
      ORDER BY je.date ASC, je.entry_num ASC, jl.id ASC`
  ).all(entityId, apAccount, asOf);

  const glBalance = glLines.reduce((s, l) => s + (l.credit || 0) - (l.debit || 0), 0);

  // ── 2. Which entries are Bill.com-synced bills? (vs. imported/manual GL entries)
  //    A 202000 credit whose entry is linked in billcom_sync_log is a synced
  //    invoice → aged Bill.com row. Everything else → GL column.
  const syncedEntryIds = new Set();
  const billcomIdByEntry = new Map();
  const invNumByEntry = new Map();
  try {
    const rows = db.prepare(
      "SELECT cl_entry_id, billcom_id, invoice_number FROM billcom_sync_log WHERE entity_id = ? AND sync_type = 'bill' AND status = 'success' AND cl_entry_id IS NOT NULL"
    ).all(entityId);
    for (const r of rows) { syncedEntryIds.add(r.cl_entry_id); billcomIdByEntry.set(r.cl_entry_id, String(r.billcom_id)); if (r.invoice_number) invNumByEntry.set(r.cl_entry_id, String(r.invoice_number)); }
  } catch (e) { /* sync log optional */ }

  // ── 3. Net each payment against the SPECIFIC bill it settled (not FIFO
  //    oldest-first), and treat the imported opening balance as a locked block.
  //    A Bill.com payment JE names its bill ("relieve bill <billcomId>"); we net
  //    it against that bill's own credit and never against the opening GL block.
  //    Imported (non-Bill.com) debits still net FIFO within the imported block.
  //    AP debits paired with cash credits (from bank transaction uploads) reduce
  //    the opening balance directly.
  //    Anything unmatched is carried out as a reconciling line so the report
  //    total still ties to the GL balance exactly.
  const entryIdByBillcomId = new Map(); // billcom_id -> synced-bill cl_entry_id
  for (const [eid, bcid] of billcomIdByEntry.entries()) entryIdByBillcomId.set(String(bcid), eid);
  
  // Identify entries with both AP debit and cash credit (opening balance relief via bank txns)
  const entryHasCashCredit = new Map(); // entry_id -> boolean
  for (const l of glLines) {
    // Cash accounts typically start with 101 or 102
    if ((String(l.account_code).startsWith('101') || String(l.account_code).startsWith('102')) && (l.credit || 0) > 0.005) {
      entryHasCashCredit.set(l.entry_id, true);
    }
  }
  
  const billOpenByEntry = new Map(); // synced-bill entry_id -> { ...line, remaining }
  let glCreditQueue = [];            // FIFO queue of imported/opening credits ONLY
  let glUnappliedDebit = 0;          // imported over-relief carried within the GL block
  let unmatchedPayment = 0;          // Bill.com payment debits not matched to a synced bill
  let openingBalanceRelief = 0;      // AP debits paired with cash credits (opening balance relief)
  const matchedPayments = []; // { targetEntry, debit } - applied AFTER all bill credits are known
  for (const l of glLines) {
    if ((l.credit || 0) > 0.005) {
      if (syncedEntryIds.has(l.entry_id)) {
        const prev = billOpenByEntry.get(l.entry_id);
        billOpenByEntry.set(l.entry_id, { ...l, remaining: (prev ? prev.remaining : 0) + l.credit });
      } else {
        let remaining = l.credit;
        if (glUnappliedDebit > 0.005) { const take = Math.min(remaining, glUnappliedDebit); remaining -= take; glUnappliedDebit -= take; }
        if (remaining > 0.005) glCreditQueue.push({ ...l, remaining });
      }
    }
    if ((l.debit || 0) > 0.005) {
      const mm = /relieve bill (\S+)/.exec(l.memo || '');
      const billId = mm ? mm[1] : null;
      const targetEntry = billId ? entryIdByBillcomId.get(String(billId)) : null;
      if (targetEntry != null) {
        matchedPayments.push({ targetEntry, debit: l.debit }); // net after the pass (a payment can post before its bill's date)
      } else if (billId) {
        unmatchedPayment += l.debit; // payment for a bill not in CL (e.g. pre-cutover) - never touch opening
      } else if (entryHasCashCredit.get(l.entry_id)) {
        // AP debit paired with cash credit (opening balance relief from bank transactions)
        openingBalanceRelief += l.debit;
      } else {
        let pay = l.debit; // imported/manual AP debit: FIFO within the imported block only
        while (pay > 0.005 && glCreditQueue.length) {
          const head = glCreditQueue[0];
          const take = Math.min(head.remaining, pay);
          head.remaining -= take; pay -= take;
          if (head.remaining <= 0.005) glCreditQueue.shift();
        }
        if (pay > 0.005) glUnappliedDebit += pay;
      }
    }
  }
  // Apply matched payments now that every synced bill credit is known - a payment
  // can post before its bill's GL date, so this must run after the pass above.
  for (const mp of matchedPayments) {
    const b = billOpenByEntry.get(mp.targetEntry);
    if (!b) { unmatchedPayment += mp.debit; continue; }
    const take = Math.min(b.remaining, mp.debit);
    b.remaining -= take;
    if (mp.debit - take > 0.005) unmatchedPayment += (mp.debit - take);
  }
  const openItems = []; // { line_id, entry_id, entry_num, date, memo, description, vendor, amount }
  for (const [eid, b] of billOpenByEntry.entries()) {
    if (b.remaining > 0.005) openItems.push({ line_id: b.line_id, entry_id: eid, entry_num: b.entry_num, date: b.date, memo: b.memo || '', description: b.description || '', vendor: b.vendor || '', amount: b.remaining });
  }
  for (const c of glCreditQueue) {
    if (c.remaining > 0.005) openItems.push({ line_id: c.line_id, entry_id: c.entry_id, entry_num: c.entry_num, date: c.date, memo: c.memo || '', description: c.description || '', vendor: c.vendor || '', amount: c.remaining });
  }

  // ── 4. Vendor + invoice number for Bill.com-synced items come from LOCAL data
  //    now (the JE's vendor field, populated at sync, plus the sync log's
  //    invoice_number), so the report builds instantly — no live Bill.com login,
  //    vendor list, or multi-year bill fetch. Due date defaults to the line date
  //    (aging is computed off the line date regardless).
  let billcomError = null;
  const invNumFromMemo = (memo) => { const s = String(memo || ''); const h = s.match(/#\s*([^\s].*?)\s*$/); if (h) return h[1].trim(); const m = s.match(/—\s*(.+?)\s*$/); return m ? m[1].trim() : null; };

  // ── 5. Build buckets for Bill.com invoices; sum GL column for the rest.
  const buckets = ['current', 'd1_30', 'd31_60', 'd61_90', 'd91_plus'];
  const emptyBuckets = () => ({ current: 0, d1_30: 0, d31_60: 0, d61_90: 0, d91_plus: 0, gl: 0, total: 0 });
  const bucketOf = (d) => d <= 0 ? 'current' : d <= 30 ? 'd1_30' : d <= 60 ? 'd31_60' : d <= 90 ? 'd61_90' : 'd91_plus';
  const dayDiff = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);

  const byVendor = new Map();
  const glRows = [];
  const grand = emptyBuckets();

  for (const it of openItems) {
    const num = invNumByEntry.get(it.entry_id) || invNumFromMemo(it.memo) || String(it.entry_num);
    const isBillcom = syncedEntryIds.has(it.entry_id);
    if (isBillcom) {
      const vname = it.vendor || 'Vendor';
      const dueDate = it.date;
      const dpd = dayDiff(asOf, it.date); // age by invoice (GL line) date
      const bk = bucketOf(dpd);
      if (!byVendor.has(vname)) byVendor.set(vname, { vendor: vname, rows: [], subtotal: emptyBuckets() });
      const grp = byVendor.get(vname);
      grp.rows.push({ date: it.date, type: 'Bill', num: String(num), entry_id: it.entry_id, entry_num: it.entry_num, vendor: vname, due_date: dueDate, past_due_days: Math.max(0, dpd), amount: it.amount, bucket: bk });
      grp.subtotal[bk] += it.amount; grp.subtotal.total += it.amount;
      grand[bk] += it.amount; grand.total += it.amount;
    } else {
      // GL column: imported/manual entry, not aged, no vendor/invoice
      glRows.push({ date: it.date, entry_num: it.entry_num, entry_id: it.entry_id, memo: it.memo, description: it.description, amount: it.amount });
      grand.gl += it.amount; grand.total += it.amount;
    }
  }

  // Net overpayment: if payments exceeded all bills (202000 is net-debit at the
  // as-of date), the leftover unapplied debit is a prepaid/overpayment balance.
  // Surface it as a negative GL line so the report still ties to the GL balance.
  if (glUnappliedDebit > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Net prepaid / overpayment (payments exceed open bills)', description: '', amount: -glUnappliedDebit });
    grand.gl -= glUnappliedDebit; grand.total -= glUnappliedDebit;
  }
  if (unmatchedPayment > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Bill.com payment(s) not matched to a synced bill', description: '', amount: -unmatchedPayment });
    grand.gl -= unmatchedPayment; grand.total -= unmatchedPayment;
  }
  if (openingBalanceRelief > 0.005) {
    glRows.push({ date: asOf, entry_num: null, entry_id: null, memo: 'Opening AP balance relief (bank transactions)', description: '', amount: -openingBalanceRelief });
    grand.gl -= openingBalanceRelief; grand.total -= openingBalanceRelief;
  }


  const glTotal = glRows.reduce((s, r) => s + r.amount, 0);
  // Once the GL-sourced A/P nets to zero (e.g. the legacy balance was cleared by a
  // journal entry), drop those GL entries from the report entirely — they no longer
  // represent anything outstanding, so the report shows only real open items.
  if (Math.abs(glTotal) < 0.005) { grand.total -= grand.gl; grand.gl = 0; glRows.length = 0; }
  const vendorsOut = Array.from(byVendor.values())
    .sort((a, b) => a.vendor.localeCompare(b.vendor))
    .map(g => ({ ...g, rows: g.rows.sort((x, y) => String(x.date).localeCompare(String(y.date))) }));

  // Reconciliation: report total should equal the GL balance by construction.
  const reportTotal = grand.total;
  const reconDiff = Math.round((reportTotal - glBalance) * 100) / 100;

  return {
    entity_id: entityId,
    as_of: asOf,
    ap_account: apAccount,
    source: 'gl',
    bucket_labels: { current: 'Current', d1_30: '1-30', d31_60: '31-60', d61_90: '61-90', d91_plus: '91+', gl: 'GL' },
    bucket_order: buckets,
    vendors: vendorsOut,
    gl_rows: glRows.sort((a, b) => String(a.date).localeCompare(String(b.date))),
    gl_total: glRows.reduce((s, r) => s + r.amount, 0),
    grand_total: grand,
    gl_balance: glBalance,
    recon_diff: reconDiff,
    bill_count: vendorsOut.reduce((n, g) => n + g.rows.length, 0),
    gl_entry_count: glRows.length,
    billcom_error: billcomError,
  };
}

module.exports = { buildApAging };
