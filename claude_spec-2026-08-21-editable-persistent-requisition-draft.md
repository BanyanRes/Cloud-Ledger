# CloudLedger — Editable / persistent Requisition Report draft (spec, 2026-08-21)

Origin: CLA/preparer comment — once a Req Report is created it should stay in CloudLedger,
editable, until the next one is created. Concretely: create a Req, and when 5 new (or updated)
invoices arrive several days later, add them in, delete anything wrong, and update everything —
without re-running the whole thing from scratch. This cuts rework.

Decisions (Jimmy, 2026-08-21):
- **Auto-seed** next month's base workbook from the last finalized Req (no re-upload).
- **One open draft per entity** (simplest).
- **Finalize** is an explicit lock step with a confirmation dialog; it files the workbook + packet to
  Workpapers and makes that report the next month's auto-seed base.
- **One copy only** in each month folder: finalize replaces any earlier report/packet for that period
  (pattern-scoped purge of "Requisition Report" / "Invoice Packet" files, so manually-added files are safe);
  if a re-finalize moves the report to a different month folder, the stale copy in the old folder is removed.
- **First-time (no prior Req on file):** the "Start new Req" screen shows a manual upload box for the prior
  month's finalized workbook. Auto-seed handles every subsequent month.
- **Upload guard = Option B (always ask):** when a hand-uploaded base matches the finalized copy in the
  Workpapers folder, CloudLedger stops and asks the user to choose "use uploaded" vs "use filed copy" every
  time — no silent default. Fires only on manual uploads, never on normal auto-seed.
- Spec first, then build.

---

## Current behavior (why editing later is hard today)

The Req Report is **stateless / one-shot**. Each run (`POST /api/requisition/:entity_id/rollforward`,
`server/index.js` ~7919):
1. You upload the prior month's workbook.
2. You code this period's invoices on screen (the coding cards; `newCurrent` + `invoices` in the body).
3. `rollForward()` mutates the uploaded workbook into Req #N+1; `verifyRollforward()` reconciles;
   `finalizeRequisitionWorkbook()` forces recalc.
4. On success the invoices are persisted to `requisition_invoice` **stamped with the new req_number**,
   the workbook + a merged invoice packet auto-save to Workpapers, and the workbook streams to the browser.

So the workbook is an **output, not a record**. The only persisted state is the individual invoices
(`requisition_invoice`) and coding history. There is no "current draft" object to reopen. When new/updated
invoices arrive, you must re-run the entire roll-forward from the prior workbook again.

Key property we exploit: roll-forward is **deterministic and idempotent** from `(base workbook + invoice set)`.
Therefore "edit the report" == "edit the invoice set and re-roll from the same base." We never hand-edit the
workbook, so the B2A / dev-fee / reconciliation guarantees stay intact.

---

## Target model — an open, editable draft per entity

A draft requisition is created once, edited freely (add / delete / update invoices), and re-rolled on each
save so its stored workbook, B2A, dev fee, and packet always reflect the current invoice set. Finalizing
locks it, stamps the req number, saves to Workpapers, and makes it the prior for next month.

### Lifecycle
```
[none] --create--> (open draft) --save*--> (open draft) --finalize--> (finalized)
                        ^  |  add / delete / edit invoices, re-roll               |
                        +--+                                                       |
   next month: new draft auto-seeds its base workbook from this finalized one  <--+
```

- **Create** — starts the entity's single open draft. Base workbook = the last finalized Req's stored
  workbook (auto-seed). For the very first draft on an entity that has no finalized Req yet, the user uploads
  a base workbook once (same as today's first run).
- **Save (re-roll)** — runs `rollForward(base, currentInvoiceSet)` against the **full current set** of the
  draft's invoices, overwrites the stored draft workbook, re-verifies, refreshes the stored packet. This is
  the existing engine, unchanged — only its invoice input now comes from the stored draft instead of a
  one-time upload. Idempotent: re-saving with no changes reproduces the same workbook.
- **Add / delete / update invoices** — operate on the draft's invoice rows (reuse the existing coding-card
  UI, `updateCard`, read-invoice/OCR, orphan purge). Each mutation marks the draft dirty; Save re-rolls.
- **Finalize** — stamps the draft's invoices with the final `req_number`, writes the final workbook + packet
  to Workpapers (today's success path via `saveRequisitionOutputs`), sets status `finalized`, and records this
  workbook as the base for the next draft. After finalize the draft is immutable (reopen only via an explicit
  Admin "reopen" if we choose to add it later — not in v1).

### One-draft-per-entity rule
At most one row with `status='open'` per `entity_id` (enforced by a partial unique index). Create refuses if
one already exists (UI routes you to the open draft instead).

---

## Data model

New table `requisition_draft`:
```sql
CREATE TABLE IF NOT EXISTS requisition_draft (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  entity_id     INTEGER NOT NULL REFERENCES entities(id),
  status        TEXT NOT NULL DEFAULT 'open',   -- 'open' | 'finalized'
  req_number    INTEGER,                        -- target req # (editable while open)
  as_of_date    TEXT,                           -- period end (editable while open)
  base_blob     BLOB,                           -- the prior workbook this draft rolls from
  base_name     TEXT,                           -- original filename of the base (for name bump)
  output_blob   BLOB,                           -- latest rolled-forward workbook
  packet_blob   BLOB,                           -- latest merged invoice packet PDF (optional cache)
  recon_ok      INTEGER,                        -- last verify result (1/0/null)
  recon_summary TEXT,                           -- last verify summary (for the "needs review" banner)
  created_at    TEXT,
  updated_at    TEXT,
  finalized_at  TEXT,
  created_by    TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_reqdraft_open
  ON requisition_draft(entity_id) WHERE status='open';
CREATE INDEX IF NOT EXISTS idx_reqdraft_entity ON requisition_draft(entity_id);
```

`requisition_invoice` gets a nullable link + a lifecycle that matches drafts:
```sql
ALTER TABLE requisition_invoice ADD COLUMN draft_id INTEGER;   -- FK-ish to requisition_draft.id
CREATE INDEX IF NOT EXISTS idx_reqinv_draft ON requisition_invoice(draft_id);
```
Semantics:
- While a draft is open, its invoices carry `draft_id = <draft>` and `req_number = NULL`.
- On **finalize**, the draft's invoices get `req_number = <final>` (today's stamp) and keep `draft_id`
  for provenance. This is backward compatible: existing rows have `draft_id = NULL` and are untouched.
- The existing orphan-purge (`req_number IS NULL`) is narrowed to `req_number IS NULL AND draft_id IS NULL`
  so it never deletes a live draft's invoices.

No change to `requisition_coa_map` / `requisition_coding_history`.

---

## API (new / changed)

All gated `requireRole('Admin','Accountant')` and behind `reqGuards()` (development / rail-assets entities).

- `GET  /api/requisition/:entity_id/draft`
  Returns the open draft (status, req_number, as_of_date, recon_ok/summary, invoice list) or `{draft:null}`.
  Powers "reopen and edit."

- `POST /api/requisition/:entity_id/draft`  — **create**
  Body: optional base workbook upload (only needed when no finalized Req exists to auto-seed from),
  `reqNumber`, `asOfDate`. Auto-seeds `base_blob` from the last finalized Req's stored workbook when present.
  Refuses (409) if an open draft already exists.

- `POST /api/requisition/:entity_id/draft/invoice`  — **add / update**
  Add newly-arrived invoices (with OCR/coding, same payload shape as today's `invoices[]`) or update an
  existing draft invoice's coding/amount. Marks draft dirty.

- `DELETE /api/requisition/:entity_id/draft/invoice/:invoice_id`  — **delete** a wrong line. Marks dirty.

- `POST /api/requisition/:entity_id/draft/roll`  — **save / re-roll**
  Re-runs `rollForward(base_blob, currentSet)` + `verifyRollforward` + `finalizeRequisitionWorkbook`,
  overwrites `output_blob` / `packet_blob` / `recon_*`. Returns recon result + a download link for the
  current workbook. (Can also auto-run on each mutation; explicit endpoint keeps re-roll cheap and testable.)

- `GET  /api/requisition/:entity_id/draft/download`  — stream the current `output_blob` (and packet).

- `POST /api/requisition/:entity_id/draft/finalize`
  Stamps invoices with `req_number`, calls `saveRequisitionOutputs` (workbook + packet → Workpapers exactly
  as today), sets `status='finalized'`, `finalized_at`, and leaves the workbook available to seed next month.

The existing one-shot `POST .../rollforward` stays for backward compatibility, or is internally reimplemented
as create→roll→finalize in one call. TBD during build; not user-visible either way.

---

## Reuse (no logic changes to the engine)

- `rollForward` (`requisition_rollforward.js`) — unchanged; called with the draft's base + current invoice set.
- `verifyRollforward`, `finalizeRequisitionWorkbook` (`requisition_preserve.js`) — unchanged.
- `saveRequisitionOutputs` (`requisition_workpaper_save.js`) — called only at **finalize** (today it runs on
  every roll-forward). While a draft is open, outputs live in `requisition_draft`, not Workpapers, so drafts
  don't clutter the Workpapers tree with intermediate versions.
- Coding cards / OCR read-invoice / coa-map auto-fill / orphan purge — unchanged UI; repointed at the draft.

Auto-seed source: the last finalized Req's workbook. Preference order — (1) `requisition_draft` where
`status='finalized'` for the entity, most recent `finalized_at`, `output_blob`; fallback (2) the newest
`entity_files` row under a `Requisition Reports` folder for the entity (covers Reqs finalized before this
feature shipped). If neither exists → user uploads a base once.

---

## UI (client/src/App.jsx)

Requisition section gains a lightweight draft header:
- If an open draft exists: banner "Draft Req #N (as of MM/DD) — last reconciled ✓ / ⚠ needs review",
  with the invoice list editable in place (existing cards), plus **Save/Re-roll**, **Download current**,
  **Finalize** buttons.
- If none: **Start new Req** (auto-seeds from last finalized; asks for a base upload only if there isn't one).
- Add-invoice reuses the existing upload/OCR/coding flow; delete is a per-card action; edits use `updateCard`.
- Access: its own `canAccess()` section (not `section:'all'`), available to Admin + Accountant per policy.

CRLF caveat: `client/src/App.jsx` is CRLF — use the Python patcher pattern for edits, not `claude-code:Edit`.

---

## Migration / safety

- Additive only: new table + one nullable column. Existing finalized Reqs and their invoices are untouched
  (`draft_id = NULL`).
- Partial unique index guarantees the one-open-draft rule at the DB level.
- Orphan purge narrowed so it can never delete a live draft's invoices.
- No change to reconciliation gating: a draft that fails verify shows the same "needs review" detail; finalize
  can still be forced (today's `force` path) with the failure surfaced.
- `node --check server/index.js` after server edits; esbuild check for the JSX.

---

## Build order (proposed)

1. Schema: `requisition_draft` + `draft_id` column + indexes; narrow orphan purge.
2. Draft CRUD endpoints (create/get/add/update/delete) — no roll yet.
3. `draft/roll` re-roll endpoint wiring the existing engine to the stored set; `draft/download`.
4. `draft/finalize` — stamp + `saveRequisitionOutputs` + status flip + next-month seed source.
5. Auto-seed resolver (finalized draft → entity_files fallback → upload).
6. Client: draft header + editable invoice list repointed at the draft; Save/Download/Finalize.
7. Test against a real entity (e.g. Braker / a CLIP-style dev entity): create → add 5 invoices → delete one →
   edit coding → re-roll → verify B2A + dev fee move correctly → finalize → confirm Workpapers save and that
   next month auto-seeds from it.

Open question for build time: whether Save auto-runs on every invoice mutation (simplest UX, more re-rolls) or
is an explicit button (cheaper, clearer). Default to explicit **Save/Re-roll** button; revisit if it feels clunky.
