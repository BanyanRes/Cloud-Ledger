#!/usr/bin/env node
// CloudLedger - reports runner (runs on Jimmy's PC)
//
// Waits for "Report" clicks on CloudLedger's Admin > Compliance Certificates and
// Admin > Time Allocation pages.
// When one is queued, it builds that quarter's lender packages from the
// OneDrive-synced SharePoint folders on this PC (same code as the server:
// server/compliancecerts.js) and saves them into
//   Accounting - Documents/General/3. Projects/Compliance Certificates/{YYYY}/Q{n}
// OneDrive then syncs the files up to SharePoint. Results go back to CloudLedger.
//
// Settings: %USERPROFILE%\.cloudledger\runner.json
//   { "url": "https://cloud-ledger.up.railway.app", "key": "<same as CC_RUNNER_KEY on Railway>",
//     "local_root": "C:\\Users\\JimmyYun\\banyanres.com" }
// Log: %USERPROFILE%\.cloudledger\runner.log
// Started at sign-in by a hidden launcher in the Windows Startup folder.

const fs = require('fs');
const os = require('os');
const path = require('path');

const HOME = process.env.CC_RUNNER_HOME || path.join(os.homedir(), '.cloudledger');
const CFG_FILE = path.join(HOME, 'runner.json');
const LOG_FILE = path.join(HOME, 'runner.log');
const LOCK_FILE = path.join(HOME, 'runner.lock');
const POLL_MS = 5000, HEARTBEAT_MS = 15000;

fs.mkdirSync(HOME, { recursive: true });
function log(msg) {
  const line = new Date().toISOString().replace('T', ' ').slice(0, 19) + '  ' + msg + '\n';
  try {
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 5 * 1024 * 1024) fs.renameSync(LOG_FILE, LOG_FILE + '.old');
    fs.appendFileSync(LOG_FILE, line);
  } catch (e) { /* ignore */ }
  if (process.stdout.isTTY) process.stdout.write(line);
}

// single instance: exit if another runner is alive
try {
  const pid = +fs.readFileSync(LOCK_FILE, 'utf8');
  if (pid && pid !== process.pid) { try { process.kill(pid, 0); console.log('Runner already running (pid ' + pid + ')'); process.exit(0); } catch (e) { /* stale lock */ } }
} catch (e) { /* no lock */ }
fs.writeFileSync(LOCK_FILE, String(process.pid));
process.on('exit', () => { try { if (+fs.readFileSync(LOCK_FILE, 'utf8') === process.pid) fs.unlinkSync(LOCK_FILE); } catch (e) { /* ignore */ } });
process.on('SIGINT', () => process.exit(0));
process.on('uncaughtException', (e) => log('uncaught: ' + (e && e.stack || e)));
process.on('unhandledRejection', (e) => log('unhandled: ' + (e && e.stack || e)));

const cfg = JSON.parse(fs.readFileSync(CFG_FILE, 'utf8'));
if (!cfg.url || !cfg.key || !cfg.local_root) { log('runner.json needs url, key and local_root'); process.exit(1); }
process.env.CC_STORAGE = 'local';
process.env.CC_LOCAL_ROOT = cfg.local_root;
if (cfg.write_root) process.env.CC_LOCAL_WRITE_ROOT = cfg.write_root; else delete process.env.CC_LOCAL_WRITE_ROOT;

async function api(p, body) {
  const r = await fetch(cfg.url.replace(/\/$/, '') + p, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Runner-Key': cfg.key, 'X-Runner-Host': os.hostname() },
    body: JSON.stringify(body || {}),
  });
  if (!r.ok) throw new Error(p + ' -> HTTP ' + r.status + ' ' + (await r.text()).slice(0, 200));
  return r.json();
}

// load the report module fresh for every job so code updates (git pull) are picked up
function loadModule() {
  const mod = path.join(__dirname, '..', 'server', 'compliancecerts.js');
  delete require.cache[require.resolve(mod)];
  return require(mod);
}
function loadTAModule() {
  const mod = path.join(__dirname, '..', 'server', 'timeallocation.js');
  delete require.cache[require.resolve(mod)];
  return require(mod);
}

async function runJob(job) {
  log('Job ' + job.id + ': quarter end ' + job.quarter_end);
  let buf = [], flushing = Promise.resolve();
  const flush = () => { if (!buf.length) return flushing; const lines = buf; buf = []; flushing = flushing.then(() => api('/api/compliance-runner/runs/' + job.id + '/log', { lines }).catch((e) => log('log post failed: ' + e.message))); return flushing; };
  const timer = setInterval(flush, 2000);
  try {
    const cc = loadModule();
    const sv = job.config || {};
    const conf = Object.assign({}, cc.DEFAULTS, sv, { midco: Object.assign({}, cc.DEFAULTS.midco, sv.midco || {}, { entities: cc.DEFAULTS.midco.entities }) });
    const clOpt = job.cl ? { url: cfg.url, token: job.cl.token, entities: job.cl.entities } : null;
    const uploads = (job.uploads || []).map((u) => {
      const data = u.data ? Buffer.from(u.data, 'base64') : null; let image = null, ext = 'png';
      if (data && /png/i.test(u.mime || u.filename)) image = data;
      else if (data && /jpe?g/i.test(u.mime || u.filename)) { image = data; ext = 'jpeg'; }
      else if (data && /pdf/i.test(u.mime || u.filename)) {
        try {
          const tmp = path.join(os.tmpdir(), 'cc_upload_' + Date.now()); fs.writeFileSync(tmp + '.pdf', data);
          const py = ['python', path.join(os.homedir(), 'AppData', 'Local', 'Programs', 'Python', 'Python313', 'python.exe')];
          let ok = false;
          for (const exe of py) { try { require('child_process').execFileSync(exe, ['-c', 'import fitz,sys; d=fitz.open(sys.argv[1]); d[0].get_pixmap(dpi=150).save(sys.argv[2])', tmp + '.pdf', tmp + '.png'], { timeout: 60000, stdio: 'ignore' }); ok = true; break; } catch (e) { /* next */ } }
          if (ok) image = fs.readFileSync(tmp + '.png'); else log('could not turn ' + u.filename + ' into a picture (Python/PyMuPDF not found)');
        } catch (e) { log('statement conversion failed: ' + e.message); }
      }
      return { kind: u.kind, filename: u.filename, mime: u.mime, balance: u.balance, image, ext };
    });
    const result = await cc.runReport(job.quarter_end, conf, (line) => { buf.push(line); log('  ' + line); }, { cl: clOpt, uploads });
    clearInterval(timer); await flush();
    await api('/api/compliance-runner/runs/' + job.id + '/finish', { ok: true, result });
    log('Job ' + job.id + ' done: ' + JSON.stringify(result.counts || {}));
  } catch (e) {
    clearInterval(timer); await flush();
    log('Job ' + job.id + ' FAILED: ' + (e && e.stack || e));
    await api('/api/compliance-runner/runs/' + job.id + '/finish', { ok: false, error: String(e && e.message || e) }).catch((x) => log('finish post failed: ' + x.message));
  }
}

async function runTAJob(job) {
  log('Time Allocation job ' + job.id + ': month end ' + job.month_end);
  let buf = [], flushing = Promise.resolve();
  const flush = () => { if (!buf.length) return flushing; const lines = buf; buf = []; flushing = flushing.then(() => api('/api/time-allocation-runner/runs/' + job.id + '/log', { lines }).catch((e) => log('log post failed: ' + e.message))); return flushing; };
  const timer = setInterval(flush, 2000);
  try {
    const ta = loadTAModule();
    const result = await ta.runReport(job.month_end, {}, (line) => { buf.push(line); log('  ' + line); });
    clearInterval(timer); await flush();
    await api('/api/time-allocation-runner/runs/' + job.id + '/finish', { ok: true, result });
    log('Time Allocation job ' + job.id + ' done: ' + JSON.stringify(result.counts || {}));
  } catch (e) {
    clearInterval(timer); await flush();
    log('Time Allocation job ' + job.id + ' FAILED: ' + (e && e.stack || e));
    await api('/api/time-allocation-runner/runs/' + job.id + '/finish', { ok: false, error: String(e && e.message || e) }).catch((x) => log('finish post failed: ' + x.message));
  }
}

let lastBeat = 0, busy = false, lastErr = '';
async function tick() {
  if (busy) return;
  busy = true;
  try {
    if (Date.now() - lastBeat > HEARTBEAT_MS) {
      await api('/api/compliance-runner/heartbeat');
      await api('/api/time-allocation-runner/heartbeat').catch(() => {});   // endpoint may not exist until deployed
      lastBeat = Date.now();
    }
    const { job } = await api('/api/compliance-runner/claim');
    if (job) await runJob(job);
    // time-allocation queue (same runner key); tolerate the endpoint being absent
    // during a deploy window so the compliance poll is never disturbed.
    let ta = null;
    try { ta = await api('/api/time-allocation-runner/claim'); }
    catch (e) { if (!/-> HTTP 404\b/.test(e.message)) throw e; }
    if (ta && ta.job) await runTAJob(ta.job);
    if (lastErr) { log('connection restored'); lastErr = ''; }
  } catch (e) {
    if (e.message !== lastErr) { log('poll error: ' + e.message); lastErr = e.message; }   // log each distinct error once
  } finally { busy = false; }
}

log('Runner started (pid ' + process.pid + ', ' + cfg.url + ', OneDrive root ' + cfg.local_root + ')');
tick();
setInterval(tick, POLL_MS);
