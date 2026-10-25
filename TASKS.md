# TASKS

How to use: one task per `##` heading. The heading text is the task name: lowercase letters, numbers, dashes, starts with a letter, max 32 characters. It becomes the worker's agent name (and the git branch name for code tasks).

Fields:
- **Status**: `todo` | `in-progress` | `blocked` | `done` | `merged`
- **Type**: `code` (gets its own git worktree + branch) or `non-code` (writes only to `outputs/<task-name>/`)
- **Agent**: filled in by the coordinator
- **Scope**: the only files/folders the worker may touch (comma-separated). Tasks with overlapping scope run one after another.
- **Depends on**: other task names that must finish first, or `-`
- **Goal**: what to do, in plain words
- **Done when**: how we know it is finished
- **Notes**: coordinator's notes (pane/workspace ids, blockers)

Only edit Status yourself if you want to skip/re-run a task. Add one `## task-name` block per task below.

---
