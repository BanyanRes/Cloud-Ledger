# CLAUDE.md — CloudLedger

Guidance for Claude working in this repository.

## Commit & deploy workflow

- **When an update is made, commit AND push to `origin/main` without stopping to ask.** Do not pause for confirmation before pushing. `main` auto-deploys to Railway (`cloud-ledger.up.railway.app`, ~2–3 min).
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
- Proceed end-to-end autonomously; don't ask "should I continue?" between steps.
- Never enter credentials. Never trigger Bill.com syncs.
