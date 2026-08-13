# CLAUDE.md — CloudLedger

Guidance for Claude working in this repository.

## Commit & deploy workflow

- **When an update is made, commit AND push to `origin/main` without stopping to ask.** Do not pause for confirmation before pushing, and never ask "do you want me to commit and push?" — this is a standing, pre-authorized instruction. Just commit and push. `main` auto-deploys to Railway (`cloud-ledger.up.railway.app`, ~2–3 min).
- Stage only named production files by explicit `git add <file>` — never `git add -A`/`.`. Scratch files (`_*.js`, `*.b64`, `*.pdf`, `*.jpg`, `*.png`, `_tmp*`) are never committed.
- Commit-message footer: `Co-Authored-By: Claude <noreply@anthropic.com>`
- Git identity: Jimmy Yun / `jyun@banyanres.com`. Never use `jimmyyun1212@gmail.com` for commits.
- Still pause only for genuinely destructive actions (`git reset --hard`, force push, history rewrite) — but a normal commit + push is not destructive and needs no confirmation.
- After pushing, verify: `git status -sb` should show `main...origin/main` with no "ahead" count.

## Repo & environment

- Local repo: `C:\Users\JimmyYun\Cloud-Ledger` (Git Bash / MINGW64).
- Remote: `BanyanRes/Cloud-Ledger` on GitHub, `main` branch → Railway auto-deploy.
- Stack: React + Vite frontend (`client/src/App.jsx`, `client/src/api.js`); Express + better-sqlite3 backend (`server/*.js`).
- **Never run the server locally** — better-sqlite3 is Railway-only. Test pure modules (e.g. ExcelJS builders) by calling their functions directly with mock data.
- Syntax-check before committing: `node --check server/<file>.js`.

## Verification

- After a deploy, confirm the build landed (poll a live endpoint or check a bundle hash) rather than assuming.
- For spreadsheet/PDF generators, generate a sample artifact and diff it against the target before committing; recalc xlsx with LibreOffice to confirm 0 formula errors.

## Style

- English for all conversation, code, commits, and output by default. Korean only when explicitly requested.
- **Always respond in plain English.** Explain things in clear, everyday language — avoid jargon, dense technical phrasing, and walls of terminology. When a technical term is unavoidable, say what it means in passing. This applies to all replies, not just client-facing ones.
- Proceed end-to-end autonomously; don't ask "should I continue?" between steps.
- Never enter credentials. Never trigger Bill.com syncs.

## Investigation discipline — get the COMPLETE picture before answering (non-negotiable)

When diagnosing data (bank transactions, AR/AP, journal entries, tie-outs, balances),
do NOT reason out loud from partial queries and do NOT state a conclusion until the
full picture is in hand. Repeatedly giving confident but wrong answers from
incomplete data is worse than taking one extra query to be right.

1. **Pull the whole set, not one record.** Before concluding where money "went"
   or whether something is missing/duplicated, query ALL related rows — e.g. every
   ar_receipt across ALL invoices for the entity, not just the one invoice you
   suspect. A payment absent from invoice A may simply be on invoice B. Never infer
   "floating / unapplied / duplicated" from a single-record view.
2. **Trace to the source of truth.** For AR application questions, the ar_receipts
   → invoice mapping is authoritative, not the JE lines alone (a JE credit to the
   A/R control account does not reveal which invoice it cleared).
3. **Check both sides of any tie.** GL control balance AND subledger detail. Only
   claim they disagree after computing both; if unsure, pull the aging-vs-GL
   recon_diff rather than guessing.
4. **State certainty honestly.** Separate "confirmed by the data I just pulled"
   from "not yet verified." If a conclusion depends on a fact not yet queried,
   query it before asserting — don't publish a guess and correct it later.
5. **One clean answer, backed by the complete query.** Prefer running 2–3 queries
   silently and giving one correct summary over narrating a chain of partial
   findings that get revised.

Rule of thumb: if answering would require me to say "actually, correction…" a
moment later, I did not gather enough first. Gather, verify, THEN answer.
