#!/usr/bin/env node
// Claude Code SessionStart hook: prepares a git worktree so the app runs
// immediately. Safe to run repeatedly; does nothing in the main checkout.
//
//   1. Installs npm dependencies (root + client) if node_modules is missing.
//   2. Gives the worktree its own dev ports by rewriting PORT / CLIENT_PORT in
//      the worktree's own .env (copied in by .worktreeinclude, gitignored).
//
// Ports: server 3101..3120, client 5174..5193. The main checkout keeps its
// own .env untouched. Production is unaffected: the Dockerfile builds from
// .env.example and sets PORT=3000.
//
// stdout on SessionStart is added to Claude's context, so print one short
// status line there and send everything else to stderr.

const fs = require('fs');
const path = require('path');
const { execFileSync, spawnSync } = require('child_process');

const log = (...a) => console.error('[worktree-setup]', ...a);

function readStdin() {
  try { return JSON.parse(fs.readFileSync(0, 'utf8') || '{}'); } catch { return {}; }
}

function git(cwd, ...args) {
  return execFileSync('git', args, { cwd, encoding: 'utf8' }).trim();
}

const input = readStdin();
const startDir = input.cwd || process.cwd();

let root, gitDir, commonDir;
try {
  root = git(startDir, 'rev-parse', '--show-toplevel');
  gitDir = path.resolve(root, git(root, 'rev-parse', '--git-dir'));
  commonDir = path.resolve(root, git(root, 'rev-parse', '--git-common-dir'));
} catch (e) {
  log('not a git checkout, skipping');
  process.exit(0);
}

// Main checkout: its .git dir IS the common dir. Leave it alone.
if (path.normalize(gitDir).toLowerCase() === path.normalize(commonDir).toLowerCase()) {
  process.exit(0);
}

// ---- 1. dependencies -------------------------------------------------------
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
for (const dir of [root, path.join(root, 'client')]) {
  if (fs.existsSync(path.join(dir, 'node_modules'))) continue;
  const cmd = fs.existsSync(path.join(dir, 'package-lock.json')) ? 'ci' : 'install';
  log(`npm ${cmd} in ${dir} ...`);
  const r = spawnSync(npm, [cmd, '--no-audit', '--no-fund'], {
    cwd: dir, stdio: ['ignore', 'pipe', 'pipe'], shell: process.platform === 'win32', encoding: 'utf8',
  });
  if (r.status !== 0) {
    log(`npm ${cmd} failed in ${dir}:\n${(r.stderr || '').slice(-2000)}`);
    console.log(`Worktree setup: npm ${cmd} FAILED in ${path.relative(root, dir) || '.'}; run it manually.`);
    process.exit(0);
  }
}

// ---- 2. per-worktree ports ------------------------------------------------
const envPath = path.join(root, '.env');
const MARK = '# worktree-ports';

function readPorts(file) {
  try {
    const t = fs.readFileSync(file, 'utf8');
    if (!t.includes(MARK)) return null;
    const p = /^PORT=(\d+)/m.exec(t), c = /^CLIENT_PORT=(\d+)/m.exec(t);
    return p && c ? { server: +p[1], client: +c[1] } : null;
  } catch { return null; }
}

let ports = readPorts(envPath);
if (!ports) {
  if (!fs.existsSync(envPath)) {
    console.log('Worktree setup: no .env found (check .worktreeinclude); ports not assigned.');
    process.exit(0);
  }
  // Find slots already used by other worktrees of this repo.
  const used = new Set();
  const list = git(root, 'worktree', 'list', '--porcelain');
  for (const line of list.split(/\r?\n/)) {
    if (!line.startsWith('worktree ')) continue;
    const other = readPorts(path.join(line.slice(9).trim(), '.env'));
    if (other) used.add(other.server - 3100);
  }
  let n = 1;
  while (used.has(n) && n < 20) n++;
  ports = { server: 3100 + n, client: 5173 + n };

  let text = fs.readFileSync(envPath, 'utf8')
    .split(/\r?\n/)
    .filter(l => !/^(PORT|CLIENT_PORT)=/.test(l) && l !== MARK)
    .join('\n').replace(/\n+$/, '');
  text += `\n${MARK}\nPORT=${ports.server}\nCLIENT_PORT=${ports.client}\n`;
  fs.writeFileSync(envPath, text);
  log(`assigned ports server=${ports.server} client=${ports.client}`);
}

console.log(
  `Worktree ready (${path.basename(root)}): server port ${ports.server}, ` +
  `client port ${ports.client}. Start with: npm.cmd run dev. ` +
  `This worktree has its own copy of data/ (sqlite DB).`
);
