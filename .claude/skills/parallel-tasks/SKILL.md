---
name: parallel-tasks
description: "Coordinate several Claude Code worker agents in parallel inside Herdr. Use when the user asks to run TASKS.md, plan or start parallel tasks, check on workers, or merge worker results. The coordinator plans, dispatches, monitors, and reports; it never edits project files itself (except TASKS.md). Requires HERDR_ENV=1."
---

# Parallel tasks coordinator (Herdr)

You are the COORDINATOR. Workers do the work. You plan, dispatch, monitor, report, and (only with the user's OK) merge.

Also load the `herdr` skill (user-level) for the full CLI reference. When unsure about a flag, run `herdr <command> --help` — never guess.

## 0. Gate: must be inside Herdr

First run: `echo $env:HERDR_ENV` (PowerShell) or `echo $HERDR_ENV` (bash).
If it is not `1`: STOP. Tell the user:
"I'm not running inside Herdr. Close this, open a terminal in the project folder, run `herdr`, start `claude` in the pane, and ask again."
Never try to launch `herdr` yourself (nested launches are blocked).

## Hard rules

1. NEVER approve a permission prompt for a worker. NEVER send `y`, `enter`, `1`, `2`, or any key into a blocked dialog unless the user told you exactly what to answer for that specific prompt.
2. NEVER start workers with bypass/auto-approve: no `--dangerously-skip-permissions`, no `--permission-mode bypassPermissions`, `acceptEdits`, `auto`, or `dontAsk`, no `--allowedTools` wildcards. ALWAYS start workers with exactly these args: `-- --permission-mode manual --setting-sources project,local`. Why: the user-level settings auto-allow Bash/PowerShell/Write/Edit (needed elsewhere), and Claude Code may default to auto mode. Loading only project settings (`.claude/settings.json`: read-only tools allowed, everything else asks) makes every write and command come back to the user. After start, `herdr agent read` the pane and confirm it does NOT say "auto mode" or "accept edits"; if it does, stop the worker and tell the user.
3. You do NOT edit project files. The only file you write is `TASKS.md` (status/agent/notes fields). No code edits, no output files.
4. Max 3 workers at once unless the user says otherwise.
5. Pane IDs, workspace IDs, and paths come ONLY from the JSON that herdr returns. Never predict or hand-type them.
6. Ask before: starting workers (show the plan first), merging into main, removing worktrees, deleting branches.
7. An `idle` or `unknown` status can be wrong when an unfamiliar dialog is showing. Always `herdr agent read` a pane before sending it anything.

## 1. Read and plan

1. Read `TASKS.md` (format is documented at its top). Only consider tasks with `Status: todo`.
2. For each task, classify:
   - **code** — changes files in the repo -> gets its own git worktree + branch.
   - **non-code** — research/writing/data -> no worktree; writes only to `outputs/<task-name>/`.
3. Check overlap using each task's `Scope` (allowed files/folders). Two tasks overlap if any Scope path is the same as, inside, or contains another's. Overlapping tasks run one after the other (order by `Depends on`, then file order). Explain each serialization in one line, e.g. "`api-auth` and `api-logging` both touch `src/api/` -> run sequentially".
   A task with `Depends on: X` waits until X is `merged` (code) or `done` (non-code).
4. Show the plan as a short table: task, type, wave (1 = start now), reason. Then STOP and wait for the user's OK.

## 2. Dispatch (after the user OKs)

Use the PowerShell tool for herdr commands and parse JSON with `ConvertFrom-Json` (jq is not installed). Keep the raw JSON if parsing fails and inspect it rather than guessing field names.

Task names become agent names and branch names: lowercase letters, digits, dashes, start with a letter, max 32 chars (`[a-z][a-z0-9_-]{0,31}`).

### Code task

```powershell
$wt = herdr worktree create --branch <task-name> --label <task-name> --no-focus | ConvertFrom-Json
$wt | ConvertTo-Json -Depth 6   # inspect once; record workspace id, root pane id, checkout path
$pane = $wt.result.root_pane.pane_id
$ws   = $wt.result.workspace.workspace_id
```

If those fields are missing, look at the printed JSON for the workspace and its root pane (or run `herdr pane list --workspace <ws>`). Record the workspace id and checkout path in TASKS.md `Notes` — you need them for cleanup.

**CloudLedger:** before starting the worker, copy `.env` (and `data/cloudledger.db` if the task needs local data) from the main checkout into the new worktree — see the CloudLedger specifics section below. Do this BETWEEN `worktree create` and `agent start`.

### Non-code task

The WORKER creates `outputs/<task-name>/`. You only create a pane for it, in a dedicated tab:

```powershell
# first non-code worker only: make a tab for them
$tab  = herdr tab create --label workers --cwd <repo-root> --no-focus | ConvertFrom-Json
$pane = $tab.result.root_pane.pane_id
# later workers: split an existing worker pane
$sp   = herdr pane split <existing-worker-pane> --direction down --cwd <repo-root> --no-focus | ConvertFrom-Json
$pane = $sp.result.pane.pane_id
```

### Start the worker and send the task

```powershell
herdr agent start <task-name> --kind claude --pane $pane --timeout 60000 -- --permission-mode manual --setting-sources project,local
```

If it returns `agent_not_ready` (e.g. a "trust this folder?" dialog in a new worktree), read the pane and ask the user what to answer. Do not answer it yourself.

Write the prompt into a PowerShell here-string (the closing "@ must be at column 0; escape any literal $ with a backtick), then send it:

```powershell
$msg = @"
You are worker <task-name>. Task: <goal from TASKS.md>.
Rules:
- Only create or modify files inside: <Scope list>  (non-code: only outputs/<task-name>/).
- Do not touch any other file. If you need to, stop and say why instead.
- Done means: <Done when from TASKS.md>.
- (code tasks) When finished, run any relevant tests, then git add only your files and git commit on your branch with a clear message. Do not push. Do not merge.
- End with a section titled SUMMARY: what you did, files changed, how you verified, open questions/risks.
"@
herdr agent prompt <task-name> $msg
```

Update TASKS.md: `Status: in-progress`, `Agent: <task-name>`, Notes: pane id / workspace id / path.

## 3. Monitor

Loop over active workers:

```powershell
herdr agent get <task-name>
herdr agent wait <task-name> --timeout 600000      # returns on idle, done, or blocked
herdr agent read <task-name> --source recent-unwrapped --lines 120
```

- `working` -> keep waiting. Check in with the user every ~10 minutes with a one-line status per worker.
- `blocked` -> read the pane. Set TASKS.md `Status: blocked`. Tell the user: which agent, what it wants to do (the exact command/file from the dialog), and the options shown. WAIT for the user's answer. Only then send exactly the key(s) the user chose with `herdr agent send-keys <task-name> <key>`.
- `idle` / `done` -> read the pane. If there is a SUMMARY section, it finished: set `Status: done`. If there is no summary or it asked a question, relay it to the user. If `agent read` returns `agent_not_idle`, wait and retry.
- If the full summary is cut off, ask the worker: "Write your SUMMARY to outputs/<task-name>/SUMMARY.md (or a temp file) and reply with only the path", then read that file.
- When a slot frees up, offer to start the next wave (ask first).

## 4. Review and merge (code tasks)

For each `done` code task:

1. Show the worker's SUMMARY.
2. Show the diff: `git -C <repo-root> diff main...<task-name> --stat` and then the full diff (long diffs: summarize per file, offer the full text).
3. Point out risks: files outside Scope, deleted files, secrets/.env, large binaries, missing tests, failing tests, conflicts with other branches.
4. Ask the user: "Merge `<task-name>` into main?" Wait for yes.
5. On yes (from the main repo checkout, which must be on `main` with a clean tree):
   `git -C <repo-root> merge --no-ff <task-name> -m "Merge <task-name>"`
   On conflict: stop, run `git merge --abort`, report, and ask how to proceed.
6. Run tests if the project has them (look for a test script/config; if none, say "no tests found").
7. Set `Status: merged`.
   **CloudLedger:** merging does NOT push. After the merge, ask the user separately before pushing `main` (a push deploys to production) — see the CloudLedger specifics section.

Non-code tasks: show the SUMMARY and list files in `outputs/<task-name>/`. Ask whether to commit them to main.

## 5. Cleanup (ask first, per worktree)

Ask: "Remove worktree for `<task-name>`?" On yes:

```powershell
herdr agent get <task-name>   # make sure it is idle first; ask the user to /exit it or approve sending ctrl+c twice
herdr worktree remove --workspace <ws>
```

This never deletes the branch. Ask separately before `git branch -d <task-name>` (never `-D`). Never use `--force` unless the user explicitly says so after you explain it would discard uncommitted work.

## CloudLedger specifics

This repo (`C:\Users\JimmyYun\Cloud-Ledger`, origin `git@github.com:BanyanRes/Cloud-Ledger.git`) has rules that override or extend the generic workflow above. Follow them for every CloudLedger run.

### Railway deploys from main — never push from here

- A push to `main` is a PRODUCTION deploy (Railway builds from `main`). Workers therefore **never push** — they commit on their own branch only. The usual "push after every commit" habit is suspended for worker branches.
- After the user approves a merge into `main` (section 4), do NOT push as part of the merge. Ask the user separately — "Push main now? (this deploys to production)" — and wait for an explicit yes.
- `git push` is already denied in `.claude/settings.json`, but do not rely on that alone — the rule is yours to enforce.

### Copy .env and the local db into each worktree BEFORE starting the worker

Herdr worktrees (`~/.herdr/worktrees/Cloud-Ledger/<branch>`) do NOT honor `.worktreeinclude`, so gitignored runtime files are missing in a fresh worktree. After `herdr worktree create` returns the checkout path, and BEFORE `herdr agent start`:

```powershell
$src = "C:\Users\JimmyYun\Cloud-Ledger"
$dst = "<checkout-path-from-worktree-create-json>"
Copy-Item "$src\.env" "$dst\.env" -Force
# only if the task needs local data:
Copy-Item "$src\data\cloudledger.db" "$dst\data\cloudledger.db" -Force
```

The SessionStart hook (`scripts/worktree-setup.js`) then runs `npm ci` and rewrites the worktree's own ports (server 3101-3120, client 5174-5193) so workers don't collide. `.env` and `data/cloudledger.db` are gitignored — never let a worker `git add` or commit them.

### App.jsx / index.js overlap — the big single-file hotspots

The frontend is essentially one file, `client/src/App.jsx` (~955 KB), and the backend is mostly `server/index.js` (~823 KB). Two tasks that both touch the same one of these overlap by the section-1 rule and are **serialized by default**.

If the user wants two such tasks parallel anyway, they may do so ONLY when each touches clearly separate functions/sections of that file, and the plan must name exactly which. Before planning, `Grep` the actual code to find and name the functions/sections each task will touch (don't infer them from the task title), and warn the user that merge conflicts are still possible.

### Worker conventions to repeat in every worker prompt

- Stage specific files by name — **never `git add -A` / `git add .`** (the checkout carries gitignored runtime files that must not be swept in).
- Commit messages include the Claude co-author tag, matching the repo's existing commits.
- Patches to `client/src/App.jsx` must be CRLF-aware (the repo is `core.autocrlf=true`; the file is CRLF in the working tree).

### Leave the main checkout's scratch files alone

The main checkout has untracked scratch files — `_*.py`, `_*.js`, `Claude outputs/`, `client/dev/`. Do not commit, delete, or move them, and do not let a worker touch them.

## Reporting style

Talk to the user in plain language (Korean is fine). Keep commands, file names, and agent names in English. Short status tables beat long prose.
