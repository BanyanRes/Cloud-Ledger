'use strict';
// Default per-entity Executive Summary pages for the financial-statement package.
//
// Behavior: generatePackage() uses an uploaded exec-summary PDF when the user
// provides one; otherwise it falls back to the entity's default rendered here.
// The title-block DATE LINE is dynamic (driven by the statement period / asOf);
// the note bodies are the CLA-authored text as written and stay static until
// CLA revises them.
//
// Keyed by entity NAME regex (primary) with optional entity CODE match, resolved
// from statements.meta (entityName is the display name, rawEntityName the raw
// entity name, entityCode the entity code).

const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// ── Page geometry (matches the statements renderer: US Letter, 54pt margins) ──
const PAGE = { w: 612, h: 792, mT: 54, mB: 54, mL: 72, mR: 72 };

// ── Date-line resolver ───────────────────────────────────────────────────────
// Most entities: "For the N Months Ended <asOf>" (meta.monthsEnded, dynamic).
// Some entities present a two-date compare header or a fixed inception stub.
// `mode`:
//   'monthsEnded'      → meta.monthsEnded  (e.g. "For the Six Months Ended June 30, 2026")
//   'monthsCompare'    → "For the Months Ended <asOf> and <prior month>"
//   'quartersCompare'  → "For the Quarters Ended <asOf> and <prior quarter>"
//   'quarterEnded'     → "For the Quarter Ended <asOf>"
//   'periodFrom:M D'   → "For the Period <M D> - <asOf>" (fixed inception start)
function monthName(m) {
  return ['January','February','March','April','May','June','July','August','September','October','November','December'][m - 1];
}
function ymd(asOf) { const [y, m, d] = asOf.split('-').map(Number); return { y, m, d }; }
function longDate(asOf) { const { y, m, d } = ymd(asOf); return monthName(m) + ' ' + d + ', ' + y; }
function endOfPriorMonth(asOf) {
  const { y, m } = ymd(asOf);
  const pm = m === 1 ? 12 : m - 1; const py = m === 1 ? y - 1 : y;
  const last = new Date(py, pm, 0).getDate();
  return py + '-' + String(pm).padStart(2, '0') + '-' + String(last).padStart(2, '0');
}
function endOfPriorQuarter(asOf) {
  const { y, m } = ymd(asOf);
  const q = Math.floor((m - 1) / 3); // 0..3
  const prevQEndMonth = q === 0 ? 12 : q * 3; // 12,3,6,9
  const py = q === 0 ? y - 1 : y;
  const last = new Date(py, prevQEndMonth, 0).getDate();
  return py + '-' + String(prevQEndMonth).padStart(2, '0') + '-' + String(last).padStart(2, '0');
}

function resolveDateLine(mode, meta) {
  const asOf = meta.asOf;
  if (!mode || mode === 'monthsEnded') return meta.monthsEnded;
  if (mode === 'monthsCompare') return 'For the Months Ended ' + longDate(asOf) + ' and ' + longDate(endOfPriorMonth(asOf));
  if (mode === 'quartersCompare') return 'For the Quarters Ended ' + longDate(asOf) + ' and ' + longDate(endOfPriorQuarter(asOf));
  if (mode === 'quarterEnded') return 'For the Quarter Ended ' + longDate(asOf);
  if (mode.startsWith('periodFrom:')) {
    const start = mode.slice('periodFrom:'.length); // e.g. "April 16"
    return 'For the Period ' + start + ' - ' + longDate(asOf);
  }
  return meta.monthsEnded;
}

// ── Entity summaries ─────────────────────────────────────────────────────────
// Each: match (name regex / codes), title (entity display line on the page),
// dateMode, basis word ('income tax basis' / GAAP handled by intro text), and
// the body as an ordered list of blocks. Block types:
//   { p: '...' }          paragraph
//   { bullets: ['...'] }  bulleted list (bullet glyph rendered)
// The intro line and closing paragraph are stored explicitly per entity because
// they vary (income-tax-basis vs GAAP; "the Company" vs the entity name).
const GAAP_INTRO = 'The accompanying financial statements include the following departures from accounting principles generally accepted in the United States of America:';
const GAAP_CLOSE = 'The financial statements are developed by the Company to comply with accounting principles generally accepted in the United States of America ("GAAP"), although there may be departures from GAAP not identified. These statements are primarily intended for use in managing the Company\u2019s operations and may not be suitable for other purposes. Users should be aware of these limitations when utilizing the financial statements.';

const SUMMARIES = [
  {
    key: 'banyan_residential',
    match: { name: /banyan\s*residential/i, codes: ['BANYANRE1'] },
    title: 'Banyan Residential, LLC',
    dateMode: 'monthsCompare',
    blocks: [
      { p: 'The accompanying financial statements include the following departures from the income tax basis of accounting:' },
      { bullets: [
        'The financial statements omit substantially all of the disclosures and the statement of members\u2019 equity ordinarily included in financial statements prepared in accordance with the income tax basis of accounting.',
        'Some intercompany payable balances are reflected as current assets rather than current liabilities.',
        'Banyan Residential, LLC performs annual, rather than monthly, allocation of salaries and wages, insurance expenses, and rent expenses to County Line Railroad Interests.',
      ] },
      { p: 'The financial statements are developed by the Banyan Residential, LLC to comply with the income tax basis of accounting, although there may be departures from the income tax basis of accounting not identified. These statements are primarily intended for use in managing the Banyan Residential, LLC\u2019s operations and may not be suitable for other purposes. Users should be aware of these limitations when utilizing the financial statements.' },
      { p: 'These financial statements follow the income tax basis of accounting, and the 2025 tax return has not yet been completed as of July 30, 2026. Any adjustments resulting from completion of the tax return are not included in these financial statements. Adjustments may be material.' },
    ],
  },
  {
    key: 'banyan_sfr_gp',
    match: { name: /banyan\s*sfr\s*gp/i, codes: [] },
    title: 'Banyan SFR GP Investors, LLC',
    dateMode: 'quartersCompare',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'The Company would typically be required to prepare consolidated financial statements. However, management\u2019s decision not to consolidate represents a departure from principles generally accepted in the United States of America.',
        'Retained earnings as shown on the financial statements indicates a deficit if the related balance is negative, as reflected by balances in parentheses.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'clr_buna',
    match: { name: /buna/i, codes: [] },
    title: 'CLR Buna Property Owner, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'A current portion of the long-term debt has not been calculated and reported as a short-term liability.',
        'The equity section of the balance sheet is not grouped by member class nor has prior activities been closed to each member.',
        'Retained earnings reflected on the balance sheet include distributions made to former members.',
        'Retained earnings as shown on the financial statements indicate a deficit if the related balance is negative, as reflected by balances in parentheses.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'clrfi_midco',
    match: { name: /clrfi|midco/i, codes: [] },
    title: 'CLRFI Midco I LLC',
    dateMode: 'quarterEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures, the statement of changes in members equity, and the statement of cash flows required by accounting principles generally accepted in the United States of America.',
        'Income as shown on the financial statement indicates a loss if the related balance is negative, as reflected by balances in parentheses.',
        'Some current asset balances are reflected as current liabilities rather than current assets.',
        'A current portion of the long-term debt has not been calculated and reported as a short-term liability.',
        'Retained earnings as shown on the financial statements indicates a deficit if the related balance is negative, as reflected by balances in parentheses.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'clro',
    match: { name: /county\s*line\s*rail\s*operations|\bclro\b/i, codes: [] },
    title: 'County Line Rail Operations, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'Payroll expense is recorded on the cash basis of accounting.',
        'Member\u2019s capital is reported as retained earnings.',
        '$88,988 of special management fees is related to 2025.',
        '$28,600 of management fees related to 2025 were reversed in January 2026.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'clip',
    match: { name: /county\s*line\s*industrial\s*park|\bclip\b/i, codes: [] },
    title: 'County Line Industrial Park, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'Income as shown on the financial statements indicates a loss if the related balance is negative, as reflected by balances in parentheses.',
        'Retained earnings as shown on the financial statements indicate a deficit if the related balance is negative, as reflected by balances in parentheses.',
        'A current portion of the long-term debt has not been calculated and reported as a short-term liability.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'clr_silsbee',
    match: { name: /silsbee/i, codes: [] },
    title: 'County Line Rail Silsbee, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'Income as shown on the financial statement indicates a loss if the related balance is negative, as reflected by balances in parentheses.',
        'Retained earnings as shown on the financial statements indicate a deficit if the related balance is negative, as reflected by balances in parentheses.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'srn',
    match: { name: /sabine|(county\s*line\s*)?srn/i, codes: [] },
    title: 'County Line SRN, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'Income as shown on the financial statements indicates a loss if the related balance is negative, as reflected by balances in parentheses.',
        'Retained earnings as shown on the financial statements indicates a deficit if the related balance is negative, as reflected by balances in parentheses.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'turnkey_rail',
    match: { name: /turnkey/i, codes: [] },
    title: 'Turnkey Rail, LLC',
    // Inception stub: fixed start (April 16), dynamic end (period end).
    dateMode: 'periodFrom:April 16',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
      ] },
      { p: 'The financial statements are developed by Turnkey Rail, LLC to comply with accounting principles generally accepted in the United States of America ("GAAP"), although there may be departures from GAAP not identified. These statements are primarily intended for use in managing Turnkey Rail, LLC\u2019s operations and may not be suitable for other purposes. Users should be aware of these limitations when utilizing the financial statements.' },
    ],
  },
  {
    key: 'braker_qozb',
    match: { name: /braker/i, codes: [] },
    title: 'Braker QOZ Business, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'The equity section of the balance sheet is not grouped by member class nor has prior activities been closed to each member.',
        'Retained earnings reflected on the balance sheet include distributions made to former members.',
      ] },
      { p: 'These financial forecasts present, to the best of management\u2019s knowledge and belief, the Company\u2019s expected financial position, results of operations, and cash flows for the forecast periods. Accordingly, the forecasts reflect its judgment as of May 31, 2026, the date these forecasts were prepared, of the expected conditions and its expected course of action. The assumptions disclosed herein are those that management believes are significant to the forecasts. There will usually be differences between the forecasted and actual results, because events and circumstances frequently do not occur as expected, and those differences may be material.' },
      { p: GAAP_CLOSE },
    ],
  },
  {
    key: 'bridge_banyan_hp',
    match: { name: /bridge\s*banyan\s*hp|hp\s*qozb/i, codes: [] },
    title: 'Bridge Banyan HP QOZB, LLC',
    dateMode: 'monthsEnded',
    blocks: [
      { p: GAAP_INTRO },
      { bullets: [
        'The financial statements omit substantially all of the disclosures required by accounting principles generally accepted in the United States of America.',
        'The equity section of the balance sheet is not grouped by member class nor has prior activities been closed to each member.',
        'Some payable balances are reflected as current assets rather than current liabilities.',
        'Some intercompany receivable balances are reflected as current liabilities rather than current assets.',
      ] },
      { p: GAAP_CLOSE },
    ],
  },
];

function resolveSummary(meta) {
  const code = String(meta.entityCode || '').trim().toUpperCase();
  const name = String(meta.rawEntityName || meta.entityName || '');
  if (code) {
    const byCode = SUMMARIES.find(s => (s.match.codes || []).map(c => c.toUpperCase()).includes(code));
    if (byCode) return byCode;
  }
  return SUMMARIES.find(s => s.match.name.test(name)) || null;
}

// ── Renderer ──────────────────────────────────────────────────────────────────
// Left-aligned wrapped prose with a centered title block; footer matches the
// statements' "<entity>, <MonthName YYYY> | See Executive Summary" style plus a
// page number. Returns a single-page (usually) PDF as a Uint8Array.
async function renderExecSummaryPdf(meta) {
  const def = resolveSummary(meta);
  if (!def) return null;

  const pdf = await PDFDocument.create();
  const reg = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  const CONTENT_W = PAGE.w - PAGE.mL - PAGE.mR;
  const TITLE_SZ = 12, DATE_SZ = 11, BODY_SZ = 10.5, FOOT_SZ = 8;
  const LEADING = 14, PARA_GAP = 8, BULLET_INDENT = 16, BULLET_GAP = 10;

  let page = pdf.addPage([PAGE.w, PAGE.h]);
  let y = PAGE.h - PAGE.mT;

  const centre = (str, size, font, yy) => {
    const w = font.widthOfTextAtSize(str, size);
    page.drawText(str, { x: (PAGE.w - w) / 2, y: yy, size, font });
  };

  // Word-wrap a string to a given width; returns array of lines.
  const wrap = (str, size, font, width) => {
    const words = str.split(/\s+/);
    const lines = []; let cur = '';
    for (const w of words) {
      const trial = cur ? cur + ' ' + w : w;
      if (font.widthOfTextAtSize(trial, size) > width && cur) { lines.push(cur); cur = w; }
      else cur = trial;
    }
    if (cur) lines.push(cur);
    return lines;
  };

  const footerAndPage = () => {
    const { y: yr, m } = ymd(meta.asOf);
    const label = def.title + ', ' + monthName(m) + ' ' + yr + '  |  See Executive Summary';
    const w = reg.widthOfTextAtSize(label, FOOT_SZ);
    page.drawText(label, { x: (PAGE.w - w) / 2, y: PAGE.mB - 24, size: FOOT_SZ, font: reg, color: rgb(0.4, 0.4, 0.4) });
    // No top-of-page number: the package draws page references only in the
    // table of contents, and statement pages carry just the centered footer.
  };

  const ensure = (space) => {
    if (y - space < PAGE.mB + 8) {
      footerAndPage();
      page = pdf.addPage([PAGE.w, PAGE.h]);
      y = PAGE.h - PAGE.mT;
    }
  };

  const drawParagraph = (str, size, font, indent = 0, bulletGlyph = null) => {
    const width = CONTENT_W - indent - (bulletGlyph ? BULLET_INDENT : 0);
    const lines = wrap(str, size, font, width);
    for (let i = 0; i < lines.length; i++) {
      ensure(LEADING);
      let x = PAGE.mL + indent;
      if (bulletGlyph && i === 0) {
        page.drawText(bulletGlyph, { x, y, size, font });
      }
      const tx = x + (bulletGlyph ? BULLET_INDENT : 0);
      page.drawText(lines[i], { x: tx, y, size, font });
      y -= LEADING;
    }
  };

  // Title block (centered): entity name, "Executive Summary", dynamic date line.
  centre(def.title, TITLE_SZ, bold, y); y -= LEADING;
  centre('Executive Summary', TITLE_SZ, bold, y); y -= LEADING;
  centre(resolveDateLine(def.dateMode, meta), DATE_SZ, bold, y); y -= LEADING * 2;

  // "Notes to the Reader:" header, left-aligned bold.
  ensure(LEADING);
  page.drawText('Notes to the Reader:', { x: PAGE.mL, y, size: BODY_SZ, font: bold }); y -= LEADING; y -= PARA_GAP;

  for (const blk of def.blocks) {
    if (blk.p) { drawParagraph(blk.p, BODY_SZ, reg); y -= PARA_GAP; }
    else if (blk.bullets) {
      for (const b of blk.bullets) { drawParagraph(b, BODY_SZ, reg, 0, '\u2022'); y -= BULLET_GAP - 4; }
      y -= PARA_GAP;
    }
  }

  footerAndPage();
  return await pdf.save();
}

module.exports = { renderExecSummaryPdf, resolveSummary, resolveDateLine, SUMMARIES };
