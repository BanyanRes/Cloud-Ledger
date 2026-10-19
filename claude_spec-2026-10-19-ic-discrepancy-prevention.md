# IC Discrepancy Prevention — Mirror JE Popup + Bill.com Invoice Email

Date: 2026-10-19 · Status: design locked, build pending (deploy pipeline stuck since af4e10a)

## Goal
Stop intercompany discrepancies at the point of entry, instead of catching them later in reconciliation.

## Mechanism 1 — Mirror journal-entry popup (balance-sheet IC: due / investment / contributed capital)
On saving a JE in entity A, scan its lines for accounts that are IC-mapped (`intercompany_accounts`). For each
IC leg, offer to post the mirror leg on the counterparty entity B, with the type flipped:
- due_from (A) -> due_to (B); due_to (A) -> due_from (B)
- investment (A) -> contributed_capital (B); contributed_capital (A) -> investment (B)
Example: Banyan records "Dr Due from CLRO 10,000" -> popup proposes a CLRO JE "Cr Due to Banyan Residential 10,000".

Decisions:
1. Counterparty account: the preparer PICKS the B-side account in the popup (account picker, same entity's chart). Not auto-chosen.
2. On confirm, POST the mirror immediately to B (not a draft). Preparer must have access + period open on B.
3. The SAME user who prepares A's JE confirms the popup. CL prefills ONLY the IC leg on B; the offsetting line
   (cash/expense/etc.) is left blank for that user to complete — CL cannot infer B's offset.
4. Scope: popup fires ONLY when A and B are in the SAME reconciliation group — CLA-CLA or Banyan-Banyan.
   CLRF (fund) is in neither group (it sits in Weaver), so the popup NEVER fires for any transaction involving CLRF,
   as required. Reuse the group-membership helper used by the recon scopes (groupEntitySet / CLA & Banyan sets).

## Mechanism 2 — IC invoice -> counterparty's Bill.com inbox (AR/AP IC)
When A issues an IC invoice to B, a "Send to counterparty" button emails the invoice PDF to B's Bill.com inbox
address. Bill.com ingests it as a bill; on B's next sync it becomes an AP bill that reconciles against A's AR.

Decisions:
5. Trigger: a "Send to counterparty" BUTTON on the IC invoice (not automatic on create).
6. Email transport: REUSE the existing Resend path in server/ar.js (review-then-send, PDF attached). Resend is
   already configured (RESEND_API_KEY / ctx.getResendKey()). No new email provider needed.

Implementation notes:
- Store each entity's Bill.com inbox address: add `bill_inbox_email` to entities (settings field + one-time populate
  from the table below).
- The counterparty's inbox is the "To"; attach CL's generated invoice PDF; reuse ar.js send + templating.

## Bill.com inbox addresses (pulled from BILL > Inbox > "Send to", 10/17/2026)
BILL organization name -> proposed CL entity (CONFIRM the two marked ?):
- Banyan Residential, LLC            -> Banyan Residential            banyanresidential@bill.com
- Banyan SFR GP Investors LLC        -> Banyan SFR GP Investors        banyansfrgp@bill.com
- Banyan Van Buren LLC               -> Banyan Van Buren               vanburen@bill.com
- Braker PropCo, LLC                 -> Braker Prop Co                 brakerpropcollc@bill.com
- CLIP Property Owner, LLC           -> CLIP Property Owner            clippropertyownerllc@bill.com
- CLR Silsbee Property Owner LLC     -> CLR Silsbee Property Owner     silsbeeap@bill.com
- County Line Rail Buna LLC          -> CLR Buna Property Owner  (?)   countylinerailbunallc@bill.com
- County Line Rail Fund I LP         -> County Line Rail Fund          countylinerailfund@bill.com
- County Line Rail Operations LLC    -> County Line Rail Operations    countylinerailoperations@bill.com
- HP Property Owner, LLC             -> HP Property Owner              hppropertyownerllc@bill.com
- Pflugerville Property Owner, LLC   -> Pflugerville Property Owner (?) pflugervillepropertyownerllc@bill.com
- Sabine River & Northern Railroad LLC -> Sabine River & Northern Railroad  srnrailroad@bill.com
- Turnkey Rail LLC                   -> Turnkey Rail                   turnkeyrailap@bill.com

## Build sequence
1. Mirror JE popup (self-contained, biggest discrepancy-killer): JE-save hook -> detect IC legs -> group-scope check
   -> popup with counterparty account picker -> post mirror to B. Build first.
2. Bill.com email: add bill_inbox_email to entities + populate; "Send to counterparty" button on IC invoice -> reuse
   ar.js Resend send to B's inbox with PDF.

## Open items
- Confirm the two "?" entity matches (Buna, Pflugerville).
- Decide settings location for bill_inbox_email (per-entity settings page).
- DEPLOY: Railway has not promoted any build since af4e10a (External rename). 7dc67a7 / 4177db4 / d5a60b3 are queued
  and unverified. Nothing here can ship or be tested until that clears.
