// ─── Administration: AM Fee Invoices ─────────────────────────────────────────
//
// Monthly / quarterly asset-management fee invoices issued by Banyan Residential,
// LLC (entity 41) and billed to its managed properties. Replaces the old manual
// Excel -> PDF -> Outlook process.
//
// Five billings, all issued from Banyan Residential (one letterhead, one wiring
// block, one invoice-number series):
//   • 48th and Washington  — Milhaus QOZ Business VIII, LLC   — FIXED, monthly,  issued the 1st  (current month)
//   • Orchards             — CPI/BYN Orchards SFR Owner, LLC  — CALC,  monthly,  issued the 20th (prior month)
//   • South Mountain       — CPI/BYN South Mountain SFR Owner — CALC,  monthly,  issued the 20th (prior month)
//   • Skye Canyon          — CPI/BYN Skye Canyon SFR Owner    — CALC,  monthly,  issued the 20th (prior month)
//   • North Tempe          — North Tempe QOZ Business, LLC    — FIXED, quarterly
//
// CALC fee = round( Net Collections of Rent x 1% ) to the nearest dollar, where
//   Net Collections = Gross Rent + Rent Losses + Other Income - A/R balance + Prepaid balance
// pulled from that month's 12-month P&L (T12) and Balance Sheet in SharePoint.
//
// For each due billing the module, in one idempotent step:
//   1. computes the amount (fixed, or calculated from SharePoint financials),
//   2. creates a NATIVE A/R invoice in Banyan Residential via ar.createInvoice
//      (Dr 12000 Accounts Receivable / Cr 43000 Asset Management Fee Revenue),
//      so it lands in the subledger + aging and posts the accrual JE,
//   3. renders the client-facing PDF in the approved invoice layout,
//   4. emails it (Resend) to the recipients (To = first, everyone else CC),
//   5. files the PDF (and, for CALC billings, the calc workbook) into the
//      AM Fees SharePoint folder via MS Graph,
//   6. records the issue in amfee_issues so it can never go out twice.
//
// An in-process scheduler (every 30 min, from 7am Pacific) issues each billing on
// its due day. Only billings with enabled=1 auto-issue; everything can also be
// previewed and sent manually from the Admin page.

const ar = require('./ar');

const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// ── date helpers (plain YYYY-MM-DD strings, UTC math) ─────────────────────────
const D = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const S = (dt) => dt.toISOString().slice(0, 10);
const addDays = (s, n) => { const x = D(s); x.setUTCDate(x.getUTCDate() + n); return S(x); };
const mdy = (s) => { const [y, m, d] = s.split('-'); return m + '/' + d + '/' + y.slice(2); };
function pacificToday() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Los_Angeles', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', hour12: false }).formatToParts(new Date());
  const g = (t) => (parts.find((p) => p.type === t) || {}).value;
  return { date: g('year') + '-' + g('month') + '-' + g('day'), hour: Number(g('hour')) % 24 };
}

// ── constants ─────────────────────────────────────────────────────────────────
const ISSUER = {
  entity_id: 41,                 // Banyan Residential, LLC
  name: 'Banyan Residential, LLC',
  address: '2381 Rosecrans Ave Suite 425, El Segundo, CA 90245',
  phone: '(240) 620-6298',
  ar_account: '12000',           // Accounts Receivable
  revenue_account: '43000',      // Asset Management Fee Revenue
  wiring: {
    name: 'Banyan Residential, LLC',
    bank: 'Maple Mark Bank',
    account: '10028868',
    routing: '111026135',
    address: '4143 Maple Ave STE 100',
    city: 'Dallas, TX 75219',
  },
};

// Recipient lists (To = first, rest CC). As provided by Jimmy, 2026-10-31.
const CLA = ['Tiffani.Thompson@claconnect.com', 'Dennis.Arada@claconnect.com'];
const BANYAN_CC = ['IBermudez@banyanres.com', 'mkim@banyanres.com', 'jyun@banyanres.com'];
const MILHAUS = {
  to: 'Remie.Magallanes@milhaus.com',
  cc: ['angela.agamata@milhaus.com', 'ap@milhausmanagement.com', 'donna.scott@milhaus.com', CLA[0], ...BANYAN_CC, CLA[1]],
};
const MARKTAYLOR = { to: 'derek.lehn@mark-taylor.com', cc: ['amy.schneider@mark-taylor.com', CLA[0], ...BANYAN_CC, CLA[1]] };
const CUSHWAKE = { to: 'rafael.baptista@cushwake.com', cc: ['chloe.schumacher@cushwake.com', CLA[0], ...BANYAN_CC, CLA[1]] };

const SP_AMFEE_BASE = "General/Invoices/AM Fees";
const SP_FIN_BASE = "General/4. Financials/Carlyle Projects' Financials";

// Seed billing definitions. Stored in amfee_billings; edited there afterwards.
// next_no = the next client-facing invoice number for this billing's stream.
// enabled = whether the scheduler may auto-issue it (1) or it is prepare-only (0).
const SEED_BILLINGS = [
  {
    key: '48th', property: '48th and Washington', client_name: 'Milhaus QOZ Business VIII, LLC',
    customer_email: null, project_hint: '48th|washington',
    cadence: 'monthly', issue_dom: 1, period: 'current', fee_type: 'fixed', fixed_amount: 5504.88,
    rounding: 'cent', next_no: 40, recipients: MILHAUS, enabled: 1,
    sp_pdf_dir: SP_AMFEE_BASE + '/48th and Washington/{YYYY}/{MONTH} {YYYY}',
    pdf_name: 'Milhaus AM Fee_ {MONTH} {YYYY}.pdf',
  },
  {
    key: 'orchards', property: 'Orchards', client_name: 'CPI/BYN Orchards SFR Owner, LLC',
    customer_email: null, project_hint: 'orchard|preserve',
    cadence: 'monthly', issue_dom: 20, period: 'prior', fee_type: 'calc', rate: 0.01, rounding: 'dollar',
    next_no: null, recipients: MARKTAYLOR, enabled: 0,
    fin_dir: SP_FIN_BASE + '/Preserve_Orchards/{YYYY}/{MONFOLDER}',
    calc: { source: 'banyan_mtr', gross_label: 'Total Gross Potential Rent', loss_label: 'Total Rent Losses', other_label: 'Total Other Income', ar_label: 'Accounts Receivable', prepaid_label: 'Prepaid Rent' },
    sp_calc_dir: SP_AMFEE_BASE + '/Orchards/{YYYY}', calc_name: 'Orchards AM Monthly Fee_{MONTH} {YYYY}.xlsx',
    sp_pdf_dir: SP_AMFEE_BASE + '/Orchards/Invoiced/{YYYY}', pdf_name: 'Orchards AM Fee_{MONABBR} {YYYY}.pdf',
  },
  {
    key: 'south_mountain', property: 'South Mountain', client_name: 'CPI/BYN South Mountain SFR Owner, LLC',
    customer_email: null, project_hint: 'south mountain',
    cadence: 'monthly', issue_dom: 20, period: 'prior', fee_type: 'calc', rate: 0.01, rounding: 'dollar',
    next_no: null, recipients: MARKTAYLOR, enabled: 0,
    fin_dir: SP_FIN_BASE + '/South Mountain/{YYYY}/{MONFOLDER}',
    calc: { source: 'banyan_mtr', gross_label: 'Total Gross Potential Rent', loss_label: 'Total Rent Losses', other_label: 'Total Other Income', ar_label: 'Accounts Receivable', prepaid_label: 'Prepaid Rent' },
    sp_calc_dir: SP_AMFEE_BASE + '/South Mountain/{YYYY}', calc_name: 'South Mountain AM Fee_{MONTH} {YYYY}.xlsx',
    sp_pdf_dir: SP_AMFEE_BASE + '/South Mountain/Invoiced/{YYYY}', pdf_name: 'South Mountain AM Fee_{MONABBR} {YYYY}.pdf',
  },
  {
    key: 'skye_canyon', property: 'Skye Canyon', client_name: 'CPI/BYN Skye Canyon SFR Owner, LLC',
    customer_email: null, project_hint: 'skye',
    cadence: 'monthly', issue_dom: 20, period: 'prior', fee_type: 'calc', rate: 0.01, rounding: 'dollar',
    next_no: null, recipients: CUSHWAKE, enabled: 0,
    fin_dir: SP_FIN_BASE + '/Skye Canyon/{YYYY}/{MONFOLDER}',
    calc: { source: 'rp_codes', gross_code: '3165-000', loss_code: '3358-000', other_code: '3650-000', ar_code: '1200-000', prepaid_code: '2670-000' },
    sp_calc_dir: SP_AMFEE_BASE + '/Skye Canyon/{YYYY}', calc_name: 'Skye Canyon AM Fee Monthly Calculation {MONTH} {YYYY}.xlsx',
    sp_pdf_dir: SP_AMFEE_BASE + '/Skye Canyon/Invoiced/{YYYY}', pdf_name: 'Skye Canyon AM Fee_{MONABBR} {YYYY}.pdf',
  },
  {
    key: 'north_tempe', property: 'North Tempe', client_name: 'North Tempe QOZ Business, LLC',
    customer_email: null, project_hint: 'tempe',
    cadence: 'quarterly', fee_type: 'fixed', fixed_amount: 18797.51, rounding: 'cent',
    next_no: 77, recipients: MILHAUS, enabled: 0,
    // Explicit quarter schedule (period_key -> issue date). Seeded with the next one.
    schedule: [{ period: '2026-Q4', issue: '2027-01-01' }],
    sp_pdf_dir: SP_AMFEE_BASE + '/North Tempe', pdf_name: 'Banyan AM Fee_North Tempe {PERIOD}.pdf',
  },
];

const DEFAULTS = {
  scheduler_on: 1,
  terms_days: 0,               // AM fees are due on receipt
  last_invoice_no: 76,         // highest client-facing number used so far (North Tempe Q4 = 0076)
  from_name: ISSUER.name,
  sp_drive_id: '',             // falls back to AMFEE_SP_DRIVE_ID / USC_SP_DRIVE_ID env
};

// ── storage ───────────────────────────────────────────────────────────────────
function ensureTables(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS amfee_config (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      data TEXT NOT NULL,
      updated_by TEXT, updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS amfee_billings (
      key TEXT PRIMARY KEY,
      data TEXT NOT NULL,
      updated_by TEXT, updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS amfee_issues (
      billing_key TEXT NOT NULL,
      period_key TEXT NOT NULL,          -- '2026-11' (monthly) | '2026-Q4' (quarterly)
      status TEXT NOT NULL,              -- issuing | issued | sent | failed | skipped
      period_label TEXT,                 -- 'November 2026' | 'Q4 2026'
      issue_date TEXT,
      invoice_no TEXT,
      amount REAL,
      ar_invoice_id INTEGER,
      calc_json TEXT,
      pdf_url TEXT,
      calc_url TEXT,
      sent_to TEXT,
      trigger TEXT,                      -- auto | manual
      error TEXT,
      by_who TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      sent_at TEXT,
      PRIMARY KEY (billing_key, period_key)
    );`);
  const row = db.prepare('SELECT id FROM amfee_config WHERE id = 1').get();
  if (!row) db.prepare('INSERT INTO amfee_config (id, data, updated_by) VALUES (1, ?, ?)').run(JSON.stringify(DEFAULTS), 'seed');
  for (const b of SEED_BILLINGS) {
    const ex = db.prepare('SELECT key FROM amfee_billings WHERE key = ?').get(b.key);
    if (!ex) db.prepare('INSERT INTO amfee_billings (key, data, updated_by) VALUES (?, ?, ?)').run(b.key, JSON.stringify(b), 'seed');
  }
  // Invoices already issued outside CloudLedger (the old manual Excel process),
  // recorded so the scheduler never re-sends them and the page shows them as done.
  const EXT_SENT = [
    ['48th', '2026-09', 'September 2026', '0038', 5504.88],
    ['48th', '2026-10', 'October 2026', '0039', 5504.88],
    ['north_tempe', '2026-Q4', 'Q4 2026', '0076', 18797.51],
    ['orchards', '2026-09', 'September 2026', '0034', 1040],
    ['south_mountain', '2026-09', 'September 2026', '0034', 1660],
    ['skye_canyon', '2026-09', 'September 2026', '0034', 3038],
  ];
  const insExt = db.prepare("INSERT OR IGNORE INTO amfee_issues (billing_key, period_key, status, period_label, invoice_no, amount, trigger, by_who, sent_at) VALUES (?,?,'sent',?,?,?,'external','sent outside CloudLedger', datetime('now'))");
  for (const row of EXT_SENT) insExt.run(row[0], row[1], row[2], row[3], row[4]);
}

function getConfig(db) {
  ensureTables(db);
  const row = db.prepare('SELECT data FROM amfee_config WHERE id = 1').get();
  let d = {}; try { d = JSON.parse(row.data); } catch (e) { d = {}; }
  return Object.assign({}, DEFAULTS, d);
}
function saveConfig(db, patch, who) {
  const cur = getConfig(db);
  const next = Object.assign({}, cur, patch || {});
  db.prepare("UPDATE amfee_config SET data = ?, updated_by = ?, updated_at = datetime('now') WHERE id = 1").run(JSON.stringify(next), who || 'system');
  return next;
}
function getBillings(db) {
  ensureTables(db);
  return db.prepare('SELECT key, data FROM amfee_billings').all().map((r) => { const b = JSON.parse(r.data); b.key = r.key; return b; });
}
function getBilling(db, key) { return getBillings(db).find((b) => b.key === key) || null; }
function saveBilling(db, key, patch, who) {
  const b = getBilling(db, key); if (!b) throw new Error('No billing ' + key);
  const next = Object.assign({}, b, patch || {}); delete next.key;
  db.prepare("UPDATE amfee_billings SET data = ?, updated_by = ?, updated_at = datetime('now') WHERE key = ?").run(JSON.stringify(next), who || 'system', key);
  return getBilling(db, key);
}

// ── period helpers ────────────────────────────────────────────────────────────
// Monthly billing: returns {key,label,ym:[y,m], issueDate} for the period that
// becomes due when we reach issue_dom of the relevant month.
function monthlyPeriodsDue(b, today) {
  const [ty, tm, td] = today.split('-').map(Number);
  const out = [];
  // consider the current and a couple of prior trigger months so a missed tick recovers
  for (let back = 0; back <= 2; back++) {
    const trig = new Date(Date.UTC(ty, tm - 1 - back, 1));
    const trY = trig.getUTCFullYear(), trM = trig.getUTCMonth() + 1;
    const issueDate = trY + '-' + String(trM).padStart(2, '0') + '-' + String(b.issue_dom).padStart(2, '0');
    if (issueDate > today) continue;
    // period billed = the trigger month (current) or the month before it (prior)
    const pd = new Date(Date.UTC(trY, trM - 1 - (b.period === 'prior' ? 1 : 0), 1));
    const py = pd.getUTCFullYear(), pm = pd.getUTCMonth() + 1;
    out.push({ key: py + '-' + String(pm).padStart(2, '0'), label: MON[pm - 1] + ' ' + py, y: py, m: pm, issueDate });
  }
  return out;
}
function quarterlyPeriodsDue(b, today) {
  return (b.schedule || []).filter((s) => s.issue <= today).map((s) => {
    const [y, q] = s.period.split('-Q'); return { key: s.period, label: 'Q' + q + ' ' + y, y: +y, quarter: +q, issueDate: s.issue };
  });
}
function periodsDue(b, today) { return b.cadence === 'quarterly' ? quarterlyPeriodsDue(b, today) : monthlyPeriodsDue(b, today); }

// ── MS Graph (app-only) ───────────────────────────────────────────────────────
function spDriveId(cfg) { return (cfg && cfg.sp_drive_id) || process.env.AMFEE_SP_DRIVE_ID || process.env.USC_SP_DRIVE_ID || ''; }
async function graphToken() {
  const t = process.env.MS_TENANT_ID, c = process.env.MS_CLIENT_ID, s = process.env.MS_CLIENT_SECRET;
  if (!t || !c || !s) return null;
  const r = await fetch('https://login.microsoftonline.com/' + t + '/oauth2/v2.0/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: c, client_secret: s, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!r.ok) throw new Error('Graph token ' + r.status + ': ' + (await r.text()).slice(0, 200));
  return (await r.json()).access_token;
}
const spEnc = (p) => String(p).split('/').filter(Boolean).map(encodeURIComponent).join('/');
async function graph(cfg, method, pathAfterDrive, opts = {}) {
  const drive = spDriveId(cfg); if (!drive) throw new Error('No SharePoint drive id (set AMFEE_SP_DRIVE_ID or USC_SP_DRIVE_ID)');
  const token = await graphToken(); if (!token) throw new Error('Graph not configured (MS_TENANT_ID/MS_CLIENT_ID/MS_CLIENT_SECRET)');
  const url = 'https://graph.microsoft.com/v1.0/drives/' + drive + pathAfterDrive;
  const headers = Object.assign({ Authorization: 'Bearer ' + token }, opts.headers || {});
  const r = await fetch(url, { method, headers, body: opts.body });
  if (!r.ok && !(opts.okStatuses || []).includes(r.status)) throw new Error('Graph ' + method + ' ' + r.status + ': ' + (await r.text()).slice(0, 250));
  return r;
}
async function spListChildren(cfg, folderPath) {
  const r = await graph(cfg, 'GET', '/root:/' + spEnc(folderPath) + ':/children?$top=400');
  return (await r.json()).value || [];
}
async function spDownload(cfg, filePath) {
  const r = await graph(cfg, 'GET', '/root:/' + spEnc(filePath) + ':/content');
  return Buffer.from(await r.arrayBuffer());
}
async function spEnsureFolder(cfg, folderPath) {
  const parts = folderPath.split('/').filter(Boolean);
  let cur = '';
  for (const part of parts) {
    const parent = cur; cur = cur ? cur + '/' + part : part;
    const body = JSON.stringify({ name: part, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
    const target = parent ? '/root:/' + spEnc(parent) + ':/children' : '/root/children';
    await graph(cfg, 'POST', target, { headers: { 'Content-Type': 'application/json' }, body, okStatuses: [409] });
  }
}
async function spUpload(cfg, folderPath, filename, buf, mime) {
  await spEnsureFolder(cfg, folderPath);
  const r = await graph(cfg, 'PUT', '/root:/' + spEnc(folderPath + '/' + filename) + ':/content',
    { headers: { 'Content-Type': mime || 'application/octet-stream' }, body: buf });
  const j = await r.json().catch(() => ({}));
  return j.webUrl || null;
}
// find the month subfolder inside a year folder regardless of abbreviation ("Sept 2026","Jul 2026",…)
function matchMonthFolder(names, y, m) {
  const full = MON[m - 1].toLowerCase(), abbr = MON_ABBR[m - 1].toLowerCase(), yr = String(y);
  const cand = names.filter((n) => n.toLowerCase().includes(yr) && (n.toLowerCase().includes(full) || n.toLowerCase().includes(abbr) || n.toLowerCase().startsWith(abbr)));
  // prefer an exact "<Mon> <Year>" style
  return cand.sort((a, b) => a.length - b.length)[0] || null;
}

// ── CALC: compute the fee for a Carlyle property from its SharePoint financials ─
// NOTE: wired but gated (enabled=0) until each property is tied out to the last
// manual month. computeCalc downloads the month's T12 + Balance Sheet, extracts
// the five inputs per the billing's calc map, and returns the breakdown.
async function computeCalc(cfg, b, period) {
  const XLSX = require('xlsx');
  const dir = b.fin_dir.replace('{YYYY}', String(period.y)).replace('{MONFOLDER}', '');
  const yearDir = dir.replace(/\/$/, '');
  const children = await spListChildren(cfg, yearDir);
  const monName = matchMonthFolder(children.filter((c) => c.folder).map((c) => c.name), period.y, period.m);
  if (!monName) throw new Error('No month folder for ' + period.label + ' under ' + yearDir);
  const monPath = yearDir + '/' + monName;
  const files = await spListChildren(cfg, monPath);
  const fnames = files.filter((f) => f.file).map((f) => f.name);
  const pick = (re) => fnames.find((n) => re.test(n.toLowerCase()));

  const inputs = {};
  if (b.calc.source === 'banyan_mtr') {
    const t12f = pick(/profit and loss|12 month/);
    const bsf = pick(/balance|trial balance|financial\.xlsx/);
    if (!t12f) throw new Error('No T12 (Profit and Loss) file in ' + monPath);
    const t12 = XLSX.read(await spDownload(cfg, monPath + '/' + t12f), { type: 'buffer' });
    const g = labelValue(t12, b.calc.gross_label), l = labelValue(t12, b.calc.loss_label), o = labelValue(t12, b.calc.other_label);
    inputs.gross = g; inputs.loss = l; inputs.other = o;
    if (bsf) { const bs = XLSX.read(await spDownload(cfg, monPath + '/' + bsf), { type: 'buffer' }); inputs.ar = labelValue(bs, b.calc.ar_label); inputs.prepaid = labelValue(bs, b.calc.prepaid_label); }
  } else if (b.calc.source === 'rp_codes') {
    const merged = pick(/mergedexcel|res\d+/);
    if (!merged) throw new Error('No merged financial export in ' + monPath);
    const wb = XLSX.read(await spDownload(cfg, monPath + '/' + merged), { type: 'buffer' });
    inputs.gross = codeValue(wb, b.calc.gross_code); inputs.loss = codeValue(wb, b.calc.loss_code); inputs.other = codeValue(wb, b.calc.other_code);
    inputs.ar = codeValue(wb, b.calc.ar_code); inputs.prepaid = codeValue(wb, b.calc.prepaid_code);
  }
  const totalRev = r2((inputs.gross || 0) + (inputs.loss || 0) + (inputs.other || 0));
  const netColl = r2(totalRev - (inputs.ar || 0) + (inputs.prepaid || 0));
  const feeRaw = netColl * (b.rate || 0.01);
  const fee = Math.round(feeRaw);     // Carlyle invoices show whole dollars
  return { inputs, totalRev, netColl, rate: b.rate || 0.01, feeRaw: r2(feeRaw), fee, source_folder: monPath };
}
// find the current-month value on a row whose label matches; MTR files put the
// month in the right-most numeric column of that row.
function labelValue(wb, label) {
  const want = String(label).toLowerCase().replace(/\s+/g, ' ').trim();
  for (const sn of wb.SheetNames) {
    const rows = require('xlsx').utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true });
    for (const row of rows) {
      const lab = (row.find((c) => typeof c === 'string' && c.toLowerCase().replace(/\s+/g, ' ').trim() === want));
      if (lab !== undefined) { const nums = row.filter((c) => typeof c === 'number'); if (nums.length) return nums[nums.length - 1]; }
    }
  }
  return null;
}
function codeValue(wb, code) {
  const want = String(code).toLowerCase();
  for (const sn of wb.SheetNames) {
    const rows = require('xlsx').utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true });
    for (const row of rows) {
      if (row.some((c) => typeof c === 'string' && c.toLowerCase().includes(want))) { const nums = row.filter((c) => typeof c === 'number'); if (nums.length) return nums[nums.length - 1]; }
    }
  }
  return null;
}
async function buildCalcWorkbook(b, period, calc) {
  const XLSX = require('xlsx');
  const i = calc.inputs;
  const aoa = [
    [b.property], ['Asset Management Fee Calculation'], [period.label], [],
    ['All cash or cash equivalent revenue:'],
    ['Gross Rent', i.gross], ['Rent Adjustment', i.loss], ['Net Rent revenue', r2((i.gross || 0) + (i.loss || 0))],
    ['Other residential', i.other], ['Total Revenue', calc.totalRev],
    ['Accounts receivable', -(i.ar || 0)], ['Prepaid Rent', i.prepaid], ['Net Collections of Rent', calc.netColl], [],
    ['Asset Management %', calc.rate], [], ['Asset Management Fee', calc.feeRaw], ['Invoiced (rounded)', calc.fee],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Monthly AM Fee Cal');
  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}

// ── PDF (client-facing, matches the approved invoice layout) ──────────────────
async function renderInvoicePdf(inv) {
  const { PDFDocument, StandardFonts } = require('pdf-lib');
  const pdf = await PDFDocument.create();
  pdf.setTitle('AM Fee Invoice ' + inv.invoice_no);
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const W = 612, H = 792, L = 60, R = 552;
  const page = pdf.addPage([W, H]);
  const txt = (s, x, y, f, size) => page.drawText(String(s), { x, y, size, font: f });
  const ctr = (s, y, f, size) => txt(s, (W - f.widthOfTextAtSize(String(s), size)) / 2, y, f, size);
  const rt = (s, xR, y, f, size) => txt(s, xR - f.widthOfTextAtSize(String(s), size), y, f, size);
  const line = (x0, x1, y, th) => page.drawLine({ start: { x: x0, y }, end: { x: x1, y }, thickness: th || 0.6 });
  const money = (v, cents) => '$' + (Number(v) || 0).toLocaleString('en-US', { minimumFractionDigits: cents ? 2 : 0, maximumFractionDigits: cents ? 2 : 0 });

  let y = H - 70;
  ctr(ISSUER.name, y, bold, 16); y -= 20;
  ctr(ISSUER.address + '  Telephone: ' + ISSUER.phone, y, reg, 9); y -= 34;
  txt('Date: ' + mdy(inv.issue_date), L, y, reg, 10); y -= 16;
  txt('Invoice #: ' + inv.invoice_no, L, y, reg, 10); y -= 16;
  txt('Client Name: ' + inv.client_name, L, y, reg, 10); y -= 28;

  txt('Line Item', L, y, bold, 10); rt('Cost', R, y, bold, 10); line(L, R, y - 4, 0.6); y -= 20;
  txt(inv.line_desc, L, y, reg, 10); rt(money(inv.amount, inv.cents), R, y, reg, 10); y -= 20;
  line(L, R, y + 4, 0.4);
  txt('Total Payment Due', L, y, bold, 11); rt(money(inv.amount, inv.cents), R, y, bold, 11); y -= 50;

  const w = ISSUER.wiring;
  ctr('Wiring Instructions', y, bold, 11); y -= 18;
  for (const lnStr of ['Name: ' + w.name, w.bank, 'Account#: ' + w.account, 'Routing#: ' + w.routing, 'Bank Address: ' + w.address, w.city]) { ctr(lnStr, y, reg, 10); y -= 16; }
  return Buffer.from(await pdf.save());
}

// ── email (Resend) ────────────────────────────────────────────────────────────
function senderAddress(fromSetting, displayName) {
  const m = String(fromSetting || '').match(/<([^>]+)>/);
  const addr = m ? m[1] : String(fromSetting || '').trim();
  return displayName + ' <' + addr + '>';
}
async function emailInvoice(ctx, cfg, inv, pdf) {
  const apiKey = ctx.getResendKey && ctx.getResendKey();
  if (!apiKey) return { ok: false, reason: 'RESEND_API_KEY not configured' };
  const to = [inv.recipients.to];
  const cc = (inv.recipients.cc || []).filter((e) => /@/.test(e));
  const body = {
    from: senderAddress(ctx.getFromEmail && ctx.getFromEmail(), cfg.from_name || ISSUER.name),
    to, cc,
    subject: inv.property + ' \u2014 ' + inv.period_label + ' Asset Management Fee \u2014 Invoice ' + inv.invoice_no,
    html: '<p>Please find attached the ' + inv.period_label + ' asset management fee invoice for <b>' + inv.property + '</b> (' + inv.client_name + ').</p>'
      + '<p>Invoice ' + inv.invoice_no + '<br>Amount due: $' + Number(inv.amount).toLocaleString('en-US', { minimumFractionDigits: inv.cents ? 2 : 0 }) + '</p>'
      + '<p>Thank you,<br>Banyan Residential, LLC</p>',
    attachments: [{ filename: inv.pdf_filename, content: pdf.toString('base64') }],
  };
  try {
    const r = await fetch('https://api.resend.com/emails', { method: 'POST', headers: { Authorization: 'Bearer ' + apiKey, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    if (!r.ok) return { ok: false, reason: 'Resend ' + r.status + ': ' + (await r.text()).slice(0, 300) };
    return { ok: true };
  } catch (e) { return { ok: false, reason: e.message }; }
}

// ── customers / projects ──────────────────────────────────────────────────────
function ensureCustomer(db, b) {
  let c = db.prepare('SELECT * FROM ar_customers WHERE entity_id = ? AND name = ?').get(ISSUER.entity_id, b.client_name);
  if (!c) {
    const r = db.prepare('INSERT INTO ar_customers (entity_id, name, email, address, terms_days) VALUES (?, ?, ?, ?, ?)')
      .run(ISSUER.entity_id, b.client_name, b.customer_email || null, null, 0);
    c = db.prepare('SELECT * FROM ar_customers WHERE id = ?').get(r.lastInsertRowid);
  }
  return c;
}
function resolveProject(db, hint) {
  if (!hint) return null;
  try {
    const rows = db.prepare('SELECT id, name FROM dim_projects WHERE entity_id = ?').all(ISSUER.entity_id);
    const re = new RegExp(hint, 'i');
    const hit = rows.find((p) => re.test(p.name || ''));
    return hit ? hit.id : null;
  } catch (e) { return null; }
}

// ── issue one invoice (idempotent) ────────────────────────────────────────────
function claim(db, bKey, pKey, trigger, who) {
  const tx = db.transaction(() => {
    const cur = db.prepare('SELECT status FROM amfee_issues WHERE billing_key = ? AND period_key = ?').get(bKey, pKey);
    if (cur && cur.status !== 'failed') return false;
    db.prepare("INSERT INTO amfee_issues (billing_key, period_key, status, trigger, by_who, created_at) VALUES (?,?,'issuing',?,?,datetime('now')) "
      + "ON CONFLICT(billing_key, period_key) DO UPDATE SET status='issuing', trigger=excluded.trigger, by_who=excluded.by_who, error=NULL").run(bKey, pKey, trigger, who);
    return true;
  });
  return tx();
}
function nextInvoiceNo(db, b) {
  // per-billing stream if it carries its own next_no, else the global counter
  if (Number.isFinite(b.next_no)) { const n = b.next_no; saveBilling(db, b.key, { next_no: n + 1 }, 'amfee'); return String(n).padStart(4, '0'); }
  const cfg = getConfig(db); const n = (cfg.last_invoice_no || 0) + 1; saveConfig(db, { last_invoice_no: n }, 'amfee'); return String(n).padStart(4, '0');
}

async function issueOne(ctx, b, period, trigger, who) {
  const { db } = ctx;
  const cfg = getConfig(db);
  if (!claim(db, b.key, period.key, trigger, who)) return { ok: false, reason: b.property + ' ' + period.label + ' already issued or skipped' };
  const upd = (patch) => {
    const cols = Object.keys(patch); const sets = cols.map((c) => c + ' = ?').join(', ');
    db.prepare('UPDATE amfee_issues SET ' + sets + ' WHERE billing_key = ? AND period_key = ?').run(...cols.map((c) => patch[c]), b.key, period.key);
  };
  try {
    // 1) amount (+ calc artifacts)
    let amount, cents = (b.rounding === 'cent'), calc = null, calcBuf = null;
    if (b.fee_type === 'fixed') { amount = r2(b.fixed_amount); }
    else { calc = await computeCalc(cfg, b, period); amount = calc.fee; calcBuf = await buildCalcWorkbook(b, period, calc); upd({ calc_json: JSON.stringify(calc) }); }
    if (!(amount > 0)) throw new Error('Computed amount is not positive (' + amount + ')');

    // 2) invoice number + native A/R invoice (posts Dr 12000 / Cr 43000)
    const invoiceNo = nextInvoiceNo(db, b);
    const cust = ensureCustomer(db, b);
    const projectId = resolveProject(db, b.project_hint);
    const lineDesc = period.label + ' Asset Management Fee';
    const invId = ar.createInvoice(db, ISSUER.entity_id, {
      customer_id: cust.id, invoice_date: period.issueDate, terms_days: (cfg.terms_days || 0),
      invoice_num: invoiceNo, ar_account_code: ISSUER.ar_account,
      lines: [{ description: lineDesc, qty: 1, rate: amount, revenue_account_code: ISSUER.revenue_account, project_id: projectId }],
      memo: b.property + ' asset management fee \u2014 ' + period.label,
    }, who || 'amfee-scheduler');
    upd({ invoice_no: invoiceNo, amount, ar_invoice_id: invId });

    // 3) PDF
    const pdfName = fill(b.pdf_name, b, period);
    const pdf = await renderInvoicePdf({
      invoice_no: invoiceNo, issue_date: period.issueDate, client_name: b.client_name, property: b.property,
      line_desc: lineDesc, amount, cents, period_label: period.label,
    });

    // 4) email
    const res = await emailInvoice(ctx, cfg, {
      property: b.property, client_name: b.client_name, period_label: period.label, invoice_no: invoiceNo,
      amount, cents, recipients: b.recipients, pdf_filename: pdfName,
    }, pdf);
    if (!res.ok) throw new Error('email: ' + res.reason);
    // mark the A/R invoice as sent
    try { db.prepare("UPDATE ar_invoices SET status='sent', sent_at=datetime('now') WHERE id=?").run(invId); } catch (e) { /* non-fatal */ }

    // 5) file to SharePoint (non-fatal — invoice already sent + booked)
    let pdfUrl = null, calcUrl = null, fileWarn = null;
    try {
      pdfUrl = await spUpload(cfg, fill(b.sp_pdf_dir, b, period), pdfName, pdf, 'application/pdf');
      if (calcBuf) calcUrl = await spUpload(cfg, fill(b.sp_calc_dir, b, period), fill(b.calc_name, b, period), calcBuf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (e) { fileWarn = e.message; }

    upd({ status: 'sent', sent_to: [b.recipients.to].concat(b.recipients.cc || []).join('; '), pdf_url: pdfUrl, calc_url: calcUrl, error: fileWarn, sent_at: new Date().toISOString() });
    console.log('[amfee] ' + trigger + ' issued ' + b.key + ' ' + period.key + ' #' + invoiceNo + ' $' + amount + (fileWarn ? ' (file warning: ' + fileWarn + ')' : ''));
    return { ok: true, invoice_no: invoiceNo, amount, ar_invoice_id: invId, pdf_url: pdfUrl, calc_url: calcUrl, file_warning: fileWarn };
  } catch (e) {
    upd({ status: 'failed', error: e.message });
    console.error('[amfee] FAILED ' + b.key + ' ' + period.key + ': ' + e.message);
    return { ok: false, reason: e.message };
  }
}
function fill(tpl, b, period) {
  return String(tpl || '')
    .replace(/\{YYYY\}/g, String(period.y))
    .replace(/\{MONTH\}/g, period.m ? MON[period.m - 1] : '')
    .replace(/\{MONABBR\}/g, period.m ? MON_ABBR[period.m - 1] : '')
    .replace(/\{MONFOLDER\}/g, '')
    .replace(/\{PERIOD\}/g, period.label || period.key);
}

// ── scheduler ─────────────────────────────────────────────────────────────────
async function tick(ctx) {
  const { db } = ctx;
  try {
    const cfg = getConfig(db); if (!cfg.scheduler_on) return;
    const now = pacificToday(); if (now.hour < 7) return;
    const sent = {};
    for (const r of db.prepare('SELECT billing_key, period_key, status FROM amfee_issues').all()) sent[r.billing_key + '|' + r.period_key] = r.status;
    for (const b of getBillings(db)) {
      if (!b.enabled) continue;
      for (const period of periodsDue(b, now.date)) {
        if (period.issueDate < addDays(now.date, -10)) continue; // don't backfill stale periods automatically
        const st = sent[b.key + '|' + period.key];
        if (st && st !== 'failed') continue;
        await issueOne(ctx, b, period, 'auto', 'scheduler');
      }
    }
  } catch (e) { console.error('[amfee] scheduler error: ' + e.message); }
}
function startScheduler(ctx) {
  if (process.env.AMFEE_SCHEDULER === 'off') { console.log('[amfee] scheduler disabled by env'); return; }
  setTimeout(() => tick(ctx), 90 * 1000);
  setInterval(() => tick(ctx), 30 * 60 * 1000);
}

// ── routes ────────────────────────────────────────────────────────────────────
function registerAmFeeRoutes(app, ctx) {
  const { db, auth, requireRole } = ctx;
  ensureTables(db);
  const who = (req) => (req.user && (req.user.email || req.user.name)) || 'system';

  app.get('/api/admin/am-fee', auth, requireRole('Admin', 'Accountant'), (req, res) => {
    try {
      const cfg = getConfig(db);
      const today = pacificToday().date;
      const issues = {}; for (const r of db.prepare('SELECT * FROM amfee_issues').all()) issues[r.billing_key + '|' + r.period_key] = r;
      const billings = getBillings(db).map((b) => {
        const due = periodsDue(b, today).map((p) => { const i = issues[b.key + '|' + p.key]; return Object.assign({}, p, { status: i ? i.status : 'due', invoice_no: i && i.invoice_no, amount: i && i.amount, error: i && i.error, sent_at: i && i.sent_at }); });
        const history = db.prepare('SELECT period_key, period_label, status, invoice_no, amount, sent_at, error FROM amfee_issues WHERE billing_key = ? ORDER BY period_key DESC LIMIT 12').all(b.key);
        return { key: b.key, property: b.property, client_name: b.client_name, cadence: b.cadence, fee_type: b.fee_type, fixed_amount: b.fixed_amount, enabled: b.enabled, recipients: b.recipients, due, history };
      });
      res.json({ config: cfg, today, issuer: ISSUER, email_configured: !!(ctx.getResendKey && ctx.getResendKey()), graph_configured: !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET && spDriveId(cfg)), billings });
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Preview the PDF for a given billing + period without issuing (no JE, no email, no file).
  app.get('/api/admin/am-fee/:key/preview.pdf', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const b = getBilling(db, req.params.key); if (!b) return res.status(404).json({ error: 'No billing ' + req.params.key });
      const today = pacificToday().date;
      const period = (periodsDue(b, '2999-12-31').find((p) => p.key === req.query.period)) || periodsDue(b, today)[0] || monthlyPeriodsDue(Object.assign({ issue_dom: 1, period: 'current' }, b), today)[0];
      if (!period) return res.status(400).json({ error: 'No period to preview' });
      const cfg = getConfig(db);
      let amount = b.fee_type === 'fixed' ? r2(b.fixed_amount) : (await computeCalc(cfg, b, period)).fee;
      const pdf = await renderInvoicePdf({ invoice_no: '(preview)', issue_date: period.issueDate || today, client_name: b.client_name, property: b.property, line_desc: period.label + ' Asset Management Fee', amount, cents: b.rounding === 'cent', period_label: period.label });
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'inline; filename="preview_' + b.key + '.pdf"');
      res.send(pdf);
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Dry-run the calc for a CALC billing + period (no side effects).
  app.get('/api/admin/am-fee/:key/calc', auth, requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const b = getBilling(db, req.params.key); if (!b) return res.status(404).json({ error: 'No billing ' + req.params.key });
      if (b.fee_type !== 'calc') return res.json({ fee_type: 'fixed', amount: b.fixed_amount });
      const period = periodsDue(b, req.query.today || '2999-12-31').find((p) => p.key === req.query.period) || periodsDue(b, pacificToday().date)[0];
      if (!period) return res.status(400).json({ error: 'No period' });
      res.json(await computeCalc(getConfig(db), b, period));
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  app.put('/api/admin/am-fee/config', auth, requireRole('Admin'), (req, res) => {
    try { res.json({ config: saveConfig(db, req.body || {}, who(req)) }); } catch (e) { res.status(400).json({ error: e.message }); }
  });
  app.put('/api/admin/am-fee/:key', auth, requireRole('Admin'), (req, res) => {
    try { res.json({ billing: saveBilling(db, req.params.key, req.body || {}, who(req)) }); } catch (e) { res.status(400).json({ error: e.message }); }
  });

  // Issue now (manual). force=true clears a prior failed/sent row for re-issue.
  app.post('/api/admin/am-fee/:key/issue', auth, requireRole('Admin'), async (req, res) => {
    try {
      const b = getBilling(db, req.params.key); if (!b) return res.status(404).json({ error: 'No billing ' + req.params.key });
      const today = pacificToday().date;
      const period = periodsDue(b, '2999-12-31').find((p) => p.key === (req.body && req.body.period)) || periodsDue(b, today)[0] || monthlyPeriodsDue(b, today)[0];
      if (!period) return res.status(400).json({ error: 'No period to issue' });
      if (req.body && req.body.force) db.prepare("UPDATE amfee_issues SET status='failed' WHERE billing_key=? AND period_key=? AND status!='sent'").run(b.key, period.key);
      const out = await issueOne(ctx, b, period, 'manual', who(req));
      if (!out.ok) return res.status(400).json({ error: out.reason });
      res.json(Object.assign({ ok: true, period: period.key }, out));
    } catch (e) { res.status(400).json({ error: e.message }); }
  });

  startScheduler(ctx);
}

module.exports = { DEFAULTS, SEED_BILLINGS, ISSUER, registerAmFeeRoutes, computeCalc, renderInvoicePdf, _tick: tick };
