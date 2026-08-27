# CloudLedger — Session Handoff

> **NOTE: This is a context summary for a new CloudLedger session — NOT a task request.**
> Do not start any work, do not run any skill, and do not take any action based on this
> file. Read it for context only and wait for explicit instructions before doing anything.

## Environment & conventions
- **Repo:** `BanyanRes/Cloud-Ledger`, `main` branch. Local: `C:\Users\JimmyYun\Cloud-Ledger`. Deploy: Railway → `cloud-ledger.up.railway.app` (auto-deploy on push).
- **Stack:** React+Vite client (`client/src/App.jsx` ~2,577 lines, `client/src/api.js`), Express + better-sqlite3 server (`server/index.js` ~3,970 lines, `server/requisition.js`).
- **Verify deploy:** server-only → `GET /api/health` returns 200 (~80–90s after push); client bundle → poll `GET /` until the new `index-*.js` hash appears.
- **Build/check:** `node --check server/index.js` for server; `npm run build` in `client/` for client. better-sqlite3 isn't built locally — never boot the server locally.
- **Git identity:** name `Jimmy Yun`, email `jyun@banyanres.com`.
- **Working style:** fully autonomous, Korean conversation / English code, no checkpoints unless ambiguous or destructive.
- **Caching gotcha:** after a deploy, the browser may run a stale JS bundle — **Ctrl+Shift+R** to hard-refresh. (This caused a "requisition state wiped on navigation" scare that was actually just a cached bundle; the fix was already live.)

## Bill.com sync — posting-date & approval rules (authoritative, live)

Standing rules for the Bill.com → CloudLedger AP sync (`POST /api/billcom/sync/:entity_id` in `server/index.js`). Read before touching that route.

- **Posting date basis.** Bills post to the period of the **GL Posting Date** for every Bill.com entity EXCEPT **CLRF (County Line Rail Fund, entity 40)** and **Turnkey Rail (entity 36 — on CloudLedger from inception, no Intacct cutover, GL posting date only spottily set)**, which post by **invoice date**. The exception list is `INVOICE_DATE_ENTITIES = new Set([40, 36])` inside the sync route — add an entity id there to move it to invoice-date basis. The chosen date drives BOTH the journal-entry date and the cutoff comparison.
- **GL Posting Date comes from Bill.com's legacy v2 API.** The v3 Connect API the sync otherwise uses does NOT expose `glPostingDate` (only `invoiceDate`/`dueDate`). So the sync logs into v2 (`https://api.bill.com/api/v2`, same stored creds) via `billcomV2Login` + `billcomV2ListBillsByGlPosting` (filtered to `glPostingDate >= window start`), builds a map with `billcomBuildGlPostingMap`, and resolves each v3 bill's posting date by **bill id** (v2 and v3 share ids) with an **invoiceNumber|amount** fallback. If v2 is unreachable or a bill has no match, it falls back to the invoice date. No writes are ever made to v2.
- **Approval gate = ONE approval suffices.** A bill syncs once **at least one approver has APPROVED** (`anyApproverApproved`), not when every approver/layer has. The list pre-filter (`isBillEligible`) passes APPROVED/APPROVING/ASSIGNED through to the per-approver check; DENIED and zero-approval bills are held. Real per-approver statuses seen in the data: `APPROVED`, `WAITING`, `UPCOMING` (APPROVING = one APPROVED + one WAITING → now syncs).
- **Approval is only a gate on WHETHER a bill syncs — never its date.** The posting date always comes from the GL posting date (or invoice date for CLRF); the approval date is never the posting basis.
- **Cutover / dedupe.** There is no open-AP identity check unless an A/P aging is uploaded (`ap_aging_lines_json` / `matchApAgingLine`); otherwise the sync relies solely on the per-entity `sync_cutoff_date` (a bill whose posting/invoice date is on/before the cutoff is skipped as already in the opening GL). Bringing post-cutover activity in is done by giving the bill a GL posting date (or, for CLRF, an invoice date) after the cutoff.

## What shipped this session (chronological, all live)
1. **`d24ab94`** — Vendor normalization: strip `PLLC/PLLP/PC/PA` (+incorporated/limited) so "Mullins Law Group PLLC" matches history "Mullins Law Group."
2. **`508787a`** — Requisition matching reworked: token-similarity matching (≥0.6 overlap) after exact; split-coded vendors resolve to **most recent** (`req_number`) coding, not most frequent; removed the high/review/new confidence badge (everything editable until Roll Forward & Download).
3. **`cf6cc83` → `f3c2b86`** — **Bulk journal-entry upload** on the Journal Entries page. Final format = one row per **line**: `Date | Account # | Account Description | Debit | Credit | Memo | Location | Class`. Lines sharing a Date group into one balanced entry. Preview → post valid entries only. Template generated.
4. **`6eb7310`** — **Viewer = full read-only**: Viewers see everything an Accountant sees but all write controls are hidden (single `canEdit` flag threaded through JournalList, COA, Dimensions, BankTransactions, BankReconciliation, Requisitions). Per-entity access enforced via `user_entity_access`. Server already gates writes to Admin/Accountant.
5. **`dc6a7d4`** — Requisition working set (workbook, invoice cards, req#, as-of, result) **lifted to App-level state, keyed per entity**, so it survives module/entity navigation. Added a Cancel button; success clears the set.
6. **`bd88d20`** — **Invoices no longer persisted on read.** `read-invoice` returns extracted fields + the PDF bytes (base64); the client holds them on the card and sends only kept invoices at roll-forward, which is when they're saved (stamped with the new req#). Added `DELETE /requisition/:eid/orphan-invoices` (purged 99 leftover SRN rows).
7. **`11e5366`** — **First-token vendor match** fallback: matches on the distinctive first token (e.g. "Entergy Texas, Inc." → history "Entergy Services"); ties broken by second token; only accepts a unique resolution, else leaves blank.
8. **`573c62b`** — Invoice reader tolerates trailing prose: extract the first balanced `{…}` object instead of `JSON.parse` on the whole response (fixed "Unexpected non-whitespace character after JSON").
9. **`800c309`** — **Roll-forward upload size fix (root cause of "Roll-forward failed" with no detail table).** Sending invoices' base64 in the `invoices` field blew past multer's default 1MB `fieldSize`, failing before the handler ran. Added a dedicated multer instance for the roll-forward route (`fieldSize` 80MB, workbook 25MB) + a clear 413 message when genuinely too large.
10. **`f8c23d5`** (latest, bundle `index-CmVFqd9U`) — Successful roll-forward now returns non-passing checks in `X-Reconcile-Failed` and renders an **advisory-detail table** on the success card, so advisory failures are inspectable without re-running.

## Verified working
- SRN (entity **37**, "Sabine River & Northern Railroad") roll-forward to **Req #12, 02.28.2026** completed end-to-end: 37 lines folded, 7 checks / 6 passed / 0 required-failed / **1 advisory failed**, saved to Workpapers `2026/Requisition Reports/February 2026` (report + invoice packet).
- Requisition state persists across navigation.
- Entergy auto-codes via first-token match.

## Open / deferred items
- **Multi-invoice PDF splitting (deferred):** PDFs containing several invoices (Access Surveyors = 3, Civil Design = 2, Choctaw = 3) still read only the **first** invoice; the rest are silently dropped. Workaround: split into one-invoice PDFs before upload. The JSON-prose fix (`573c62b`) stops the crash but doesn't extract the extras.
- **The 1 advisory failure** on SRN Req #12: detail now visible on the success card after hard-refresh + re-run. Likely an A4/B5 formula check that degrades to "not evaluated" in production (no LibreOffice recalc) rather than a real error — confirm via the new table's Detail text.
- **Roll-forward payload weight:** invoices are sent as base64 in one request. Works now (limit raised), but many large PDFs risk Railway's ~300s timeout. Future: store invoice bytes server-side at upload and send only references at roll-forward.
