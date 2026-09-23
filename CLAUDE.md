# CLAUDE.md — CloudLedger

Guidance for Claude working in this repository.

## Commit & deploy workflow

- **In the main checkout (`C:\Users\JimmyYun\Cloud-Ledger` on `main`): when an update is made, commit AND push to `origin/main` without stopping to ask.** (Worktree sessions follow "Parallel sessions" below instead.) Do not pause for confirmation before pushing, and never ask "do you want me to commit and push?" — this is a standing, pre-authorized instruction. Just commit and push. `main` auto-deploys to Railway (`cloud-ledger.up.railway.app`, ~2–3 min).
- Stage only named production files by explicit `git add <file>` — never `git add -A`/`.`. Scratch files (`_*.js`, `*.b64`, `*.pdf`, `*.jpg`, `*.png`, `_tmp*`) are never committed.
- Commit-message footer: `Co-Authored-By: Claude <noreply@anthropic.com>`
- Git identity: Jimmy Yun / `jyun@banyanres.com`. Never use `jimmyyun1212@gmail.com` for commits.
- Still pause only for genuinely destructive actions (`git reset --hard`, force push, history rewrite) — but a normal commit + push is not destructive and needs no confirmation.
- After pushing, verify: `git status -sb` should show `main...origin/main` with no "ahead" count.

## Repo & environment

- Local repo: `C:\Users\JimmyYun\Cloud-Ledger` (Git Bash / MINGW64).
- Remote: `BanyanRes/Cloud-Ledger` on GitHub, `main` branch → Railway auto-deploy.
- Stack: React + Vite frontend (`client/src/App.jsx`, `client/src/api.js`); Express + better-sqlite3 backend (`server/*.js`).
- The server runs locally (better-sqlite3 works on node 22.11.0 via nvm-windows; use `npm.cmd`, not `npm`). `npm.cmd run dev` starts the API server (PORT from `.env`) and Vite (CLIENT_PORT from `.env`, default 5173). Pure modules (e.g. ExcelJS builders) can still be tested directly with mock data.
- Syntax-check before committing: `node --check server/<file>.js`.

## Parallel sessions (worktrees)

Several Claude Code sessions can work on CloudLedger at once, each in its own git worktree (Code tab worktree option, or `claude --worktree <name>`). Worktrees live in `.claude/worktrees/` (gitignored).

- **What happens automatically:** `.worktreeinclude` copies `.env` and `data/` into the new worktree; the SessionStart hook (`scripts/worktree-setup.js`) runs `npm ci` for root and `client/` and writes a unique `PORT` (3101+) / `CLIENT_PORT` (5174+) into that worktree's own `.env`. It never touches the main checkout.
- **One narrow task per worktree.** Keep each session to a single, well-scoped change so branches merge cleanly.
- **Own branch.** Commit and push only the worktree's branch (`git push -u origin <branch>`). Never push to `main` from a worktree, and never commit in the main checkout from a worktree session.
- **Own port.** Start the app with `npm.cmd run dev`; use the ports printed at session start (also in the worktree's `.env`). Never use 3000/3001/5173 from a worktree.
- **Own database.** Each worktree has its own snapshot copy of `data/cloudledger.db`. Local test data and schema changes stay in that worktree. Don't point `DB_PATH` at another checkout's DB.
- **Merge back via branch/PR, one at a time.** In the main checkout: `git pull`, then merge the branch (or merge its PR), resolve conflicts showing both sides for anything ambiguous, run `npm.cmd run build` before declaring done.
- **Only `main` deploys.** Pushing `main` auto-deploys Railway production, so merging to main is a deploy: do it only when Jimmy says the task is done.
- **Browser:** only one session can drive Claude in Chrome at a time; other sessions stay code-only (see `C:\Users\JimmyYun\claude_lanes\LANES.md` for the browser pool).
- **Cleanup:** when a task is merged, remove its worktree and delete its branch.
## Verification

- After a deploy, confirm the build landed (poll a live endpoint or check a bundle hash) rather than assuming.
- For spreadsheet/PDF generators, generate a sample artifact and diff it against the target before committing; recalc xlsx with LibreOffice to confirm 0 formula errors.

## Style

- English for all conversation, code, commits, and output by default. Korean only when explicitly requested.
- **Always respond in plain English.** Explain things in clear, everyday language — avoid jargon, dense technical phrasing, and walls of terminology. When a technical term is unavoidable, say what it means in passing. This applies to all replies, not just client-facing ones.
- Proceed end-to-end autonomously; don't ask "should I continue?" between steps.
- Never enter credentials. Never trigger Bill.com syncs.

## Browser / Chrome MCP — ALWAYS attach to an existing browser, NEVER create a tab

Jimmy has THREE browsers already set up and connected for Claude to use (the
"Claude (MCP)" tabs). A chat must drive one of those existing browsers — it must
NOT spin up its own tab. Opening a new MCP tab every session is a known,
repeatedly-reported annoyance. Follow this exactly:

- **NEVER call `tabs_context_mcp` with `createIfEmpty:true`.** That flag is what
  spawns a fresh tab/tab-group. Do not use it.
- **To start any browser work, first `list_connected_browsers`,** then
  `select_browser` (by `deviceId`) to attach to one already open. If a connection
  handshake is needed, use `switch_browser`. Only after attaching do anything else.
- If the attached browser already has the CloudLedger tab open, **reuse it** —
  read its state / `navigate` within it. Do not open a second tab for the same site.
- Only if `list_connected_browsers` returns nothing at all should Claude tell
  Jimmy the extension looks disconnected (and how to reconnect) — never silently
  fall back to creating a tab.
- One `claude-code` stdio pipe / one browser attachment at a time; a second
  concurrent chat can cause false timeouts. After any MCP reconnect, verify with a
  trivial action before resuming.

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
