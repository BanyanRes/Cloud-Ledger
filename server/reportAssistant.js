// ═══ Reports › Ask Claude — conversational reporting assistant ═══
//
// A chat endpoint (Reports › Ask Claude in the client) that lets a user talk to
// Claude about an entity's financial-statement GROUPINGS and PRESENTATION, and
// have it PROPOSE concrete regrouping changes that the user can review and apply.
//
// "Groupings" in CloudLedger = each account's fs_section / fs_subsection, which
// control where the line lands on the Balance Sheet / Statement of Operations
// (see server/financials.js and the Chart-of-Accounts section picker). This
// module never mutates on /chat — it only proposes. The user reviews the diff in
// the UI and clicks Apply, which hits /apply; /apply re-validates every change
// server-side (account exists; target section/subsection is valid for that
// account's type on this entity's profile) and writes fs_section/fs_subsection
// using the SAME fields the PUT /accounts/:code editor uses. Both routes are
// reachable by the Accountant role, per the CloudLedger access policy.
//
// The Anthropic call reuses aiModel.anthropicFetch (model:'auto' → newest Haiku,
// retirement-retry), keyed off ANTHROPIC_API_KEY, exactly like the other AI
// endpoints (invoice reader, TTM "items needing attention").

const { anthropicFetch } = require('./aiModel');
const financials = require('./financials');

const MESSAGES_URL = 'https://api.anthropic.com/v1/messages';

// ── Build the per-account-type pick-list of valid sections/subsections for an
//    entity, identical in shape to GET /api/entities/:eid/statement-sections.
function buildPickList(db, eid, ent) {
  const profile = financials.entityProfile({ entityName: ent.name, entityCode: ent.code, entityType: ent.entity_type });
  let tax;
  if (profile === 'clrf') {
    tax = financials.fundStatementTaxonomy();
  } else {
    const accounts = db.prepare('SELECT code, name, type, subtype, fs_section, fs_subsection FROM accounts WHERE entity_id=?').all(eid);
    tax = financials.statementLineItems(profile, accounts);
  }
  const byType = {};
  for (const [type, def] of Object.entries(tax)) {
    byType[type] = { statement: def.statement, sections: def.sections.map(sec => ({ section: sec.section, subsections: sec.subsections.slice() })) };
  }
  const customs = db.prepare('SELECT statement, section, subsection FROM statement_subsections WHERE entity_id=? ORDER BY subsection').all(eid);
  for (const c of customs) {
    for (const def of Object.values(byType)) {
      const sec = def.sections.find(s => s.section === c.section);
      if (sec && !sec.subsections.includes(c.subsection)) sec.subsections.push(c.subsection);
    }
  }
  return { profile, byType };
}

// ── Validate one proposed change against the pick-list. `extraSubs` is the set
//    of (section|subsection) pairs introduced by new_subsections in the same
//    proposal, which count as valid targets. Returns { ok, issue }.
function validateChange(acct, pick, extraSubs) {
  if (!acct) return { ok: false, issue: 'account not found' };
  const def = pick.byType[acct.type];
  if (!def) return { ok: false, issue: 'no statement sections for account type ' + acct.type };
  // Clearing the section (null) resets the line to its default heuristic placement.
  // We treat that as valid only when the subsection is also cleared.
  const toSec = acct._toSection;
  const toSub = acct._toSubsection;
  if (toSec == null) {
    return (toSub == null) ? { ok: true, issue: '' } : { ok: false, issue: 'cannot set a subsection without a section' };
  }
  const sec = def.sections.find(s => s.section === toSec);
  if (!sec) return { ok: false, issue: 'section "' + toSec + '" is not valid for ' + acct.type + ' accounts on this entity' };
  if (toSub == null || toSub === '') return { ok: true, issue: '' };
  const known = sec.subsections.includes(toSub) || (extraSubs && extraSubs.has(toSec + '\u0000' + toSub));
  if (!known) return { ok: false, issue: 'subsection "' + toSub + '" is not a valid subsection of "' + toSec + '"' };
  return { ok: true, issue: '' };
}

// ── Find which statement a section belongs to (for new_subsections' statement).
function statementForSection(pick, section) {
  for (const def of Object.values(pick.byType)) {
    if (def.sections.some(s => s.section === section)) return def.statement || 'Statement of Operations';
  }
  return 'Statement of Operations';
}

// ── Extract the first top-level JSON object from a text blob (handles ```json
//    fences, surrounding prose, and nested braces/strings). Returns null if none.
function extractJsonObject(text) {
  if (!text) return null;
  let clean = String(text).replace(/```json/gi, '```').trim();
  const start = clean.indexOf('{');
  if (start === -1) return null;
  let depth = 0, inStr = false, esc = false, end = -1;
  for (let i = start; i < clean.length; i++) {
    const ch = clean[i];
    if (inStr) { if (esc) esc = false; else if (ch === '\\') esc = true; else if (ch === '"') inStr = false; }
    else if (ch === '"') inStr = true;
    else if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) { end = i; break; } }
  }
  if (end === -1) return null;
  try { return JSON.parse(clean.slice(start, end + 1)); } catch (_) { return null; }
}

// ── Remove any fenced code block (``` … ```) from the user-visible reply so the
//    raw proposal JSON isn't shown twice (the UI renders it as a review table).
function stripCodeFences(text) {
  return String(text || '').replace(/```[\s\S]*?```/g, '').replace(/\n{3,}/g, '\n\n').trim();
}

// ── Allowed inbound image types for pasted/attached screenshots, and a size cap
//    on the base64 payload (~5 MB binary) so a huge paste can't blow up a request.
const ALLOWED_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/webp']);
const MAX_IMAGE_B64 = 7 * 1024 * 1024;

// ── Sanitize one message's content. Accepts a plain string, or an array of
//    content blocks (text + base64 image) as the Anthropic Messages API expects.
//    Returns a cleaned string / block-array, or null when nothing is usable.
function sanitizeContent(content) {
  if (typeof content === 'string') { const t = content.slice(0, 8000); return t.trim() ? t : null; }
  if (Array.isArray(content)) {
    const blocks = [];
    let imgCount = 0;
    for (const b of content) {
      if (!b || typeof b !== 'object') continue;
      if (b.type === 'text' && typeof b.text === 'string' && b.text.trim()) {
        blocks.push({ type: 'text', text: b.text.slice(0, 8000) });
      } else if (b.type === 'image' && b.source && b.source.type === 'base64'
          && ALLOWED_IMAGE_TYPES.has(b.source.media_type) && typeof b.source.data === 'string'
          && b.source.data.length <= MAX_IMAGE_B64 && imgCount < 10) {
        blocks.push({ type: 'image', source: { type: 'base64', media_type: b.source.media_type, data: b.source.data } });
        imgCount++;
      }
    }
    return blocks.length ? blocks : null;
  }
  return null;
}

function buildSystemPrompt(ent, pick, accounts) {
  const lines = [];
  lines.push('You are CloudLedger\u2019s reporting assistant for the entity "' + ent.name + '" (reporting profile: ' + pick.profile + ').');
  lines.push('');
  lines.push('CloudLedger is a multi-entity accounting system. You help the user with the FINANCIAL-STATEMENT GROUPINGS and PRESENTATION of this entity\u2019s reports.');
  lines.push('');
  lines.push('How groupings work: every general-ledger account has two optional placement fields, fs_section and fs_subsection. They decide where the account\u2019s balance appears on the financial statements (Balance Sheet / Statement of Operations). When fs_section is blank, the account falls to its default placement by account type and name heuristics. Changing fs_section / fs_subsection is purely a re-grouping/presentation change: it never changes any dollar amount, only where a line is shown and which subtotal it rolls into.');
  lines.push('');
  lines.push('VALID SECTIONS AND SUBSECTIONS for this entity, by account type (you may ONLY target these; you may also introduce a brand-new subsection under an existing valid section):');
  for (const [type, def] of Object.entries(pick.byType)) {
    lines.push('  ' + type + ' (' + def.statement + '):');
    for (const s of def.sections) {
      lines.push('    - ' + s.section + (s.subsections.length ? '  [subsections: ' + s.subsections.join('; ') + ']' : '  [no subsections]'));
    }
  }
  lines.push('');
  lines.push('CURRENT CHART OF ACCOUNTS (code | name | type | current section > subsection):');
  for (const a of accounts) {
    lines.push('  ' + a.code + ' | ' + a.name + ' | ' + a.type + ' | ' + (a.fs_section || '(default)') + (a.fs_subsection ? ' > ' + a.fs_subsection : ''));
  }
  lines.push('');
  lines.push('INSTRUCTIONS:');
  lines.push('- Answer questions about how this entity\u2019s statements are grouped and presented, in clear, concise language a CAO would expect.');
  lines.push('- When the user asks you to CHANGE a grouping or presentation, first explain briefly what you will do, then emit ONE fenced ```json block describing the change set in exactly this shape:');
  lines.push('  {"changes":[{"code":"<account code>","fs_section":"<section or null>","fs_subsection":"<subsection or null>","reason":"<one short sentence>"}],"new_subsections":[{"section":"<existing valid section>","subsection":"<new subsection name>"}]}');
  lines.push('- Use EXACT account codes from the chart above and EXACT section/subsection names from the valid list. Set fs_section to null to reset an account to its default placement. Omit new_subsections (or use []) unless you are introducing a subsection that is not already in the valid list.');
  lines.push('- Only include accounts you are actually changing. Do not invent accounts, sections, or subsections. If a request is ambiguous or cannot be done with grouping changes alone (e.g. it needs a code-level report-format change), say so plainly and do not emit a JSON block.');
  lines.push('- If the user is only asking a question and not requesting a change, do NOT emit a JSON block.');
  return lines.join('\n');
}

function registerReportAssistantRoutes(app, deps) {
  const { db, auth, requireEntityAccess, requireRole } = deps;

  // POST /api/entities/:eid/assistant/chat
  // Body: { messages: [{role:'user'|'assistant', content:string}, ...] }
  // Returns: { reply, proposal|null, model }
  app.post('/api/entities/:eid/assistant/chat', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), async (req, res) => {
    try {
      const eid = +req.params.eid;
      const ent = db.prepare('SELECT id, name, code, entity_type FROM entities WHERE id=?').get(eid);
      if (!ent) return res.status(404).json({ error: 'Entity not found' });
      const apiKey = process.env.ANTHROPIC_API_KEY;
      if (!apiKey) return res.status(503).json({ error: 'The assistant is not configured (ANTHROPIC_API_KEY missing on the server).' });

      const rawMsgs = Array.isArray(req.body && req.body.messages) ? req.body.messages : [];
      let messages = [];
      for (const m of rawMsgs) {
        if (!m || (m.role !== 'user' && m.role !== 'assistant')) continue;
        const c = sanitizeContent(m.content);
        if (c) messages.push({ role: m.role, content: c });
      }
      messages = messages.slice(-20);
      if (!messages.length || messages[messages.length - 1].role !== 'user') {
        return res.status(400).json({ error: 'Expected a conversation ending with a user message.' });
      }

      const accounts = db.prepare('SELECT code, name, type, subtype, fs_section, fs_subsection FROM accounts WHERE entity_id=? ORDER BY type, code').all(eid);
      const pick = buildPickList(db, eid, ent);
      const system = buildSystemPrompt(ent, pick, accounts);

      const apiRes = await anthropicFetch(MESSAGES_URL, {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({ model: 'auto', max_tokens: 2000, system, messages }),
      });
      if (!apiRes.ok) {
        const t = await apiRes.text();
        return res.status(502).json({ error: 'Assistant failed (Anthropic ' + apiRes.status + '): ' + t.slice(0, 300) });
      }
      const data = await apiRes.json();
      const fullText = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');

      // Parse an optional proposal and validate it against live data.
      let proposal = null;
      const parsed = extractJsonObject(fullText);
      if (parsed && (Array.isArray(parsed.changes) || Array.isArray(parsed.new_subsections))) {
        const acctByCode = new Map(accounts.map(a => [String(a.code), a]));
        const newSubs = Array.isArray(parsed.new_subsections) ? parsed.new_subsections
          .filter(s => s && s.section && s.subsection)
          .map(s => ({ section: String(s.section), subsection: String(s.subsection) })) : [];
        const extraSubs = new Set(newSubs.map(s => s.section + '\u0000' + s.subsection));
        const changes = (Array.isArray(parsed.changes) ? parsed.changes : []).map(c => {
          const code = String(c.code == null ? '' : c.code);
          const acct = acctByCode.get(code);
          const toSection = (c.fs_section === null || c.fs_section === '' || c.fs_section === undefined) ? null : String(c.fs_section);
          const toSubsection = (c.fs_subsection === null || c.fs_subsection === '' || c.fs_subsection === undefined) ? null : String(c.fs_subsection);
          const probe = acct ? Object.assign({}, acct, { _toSection: toSection, _toSubsection: toSubsection }) : null;
          const v = validateChange(probe, pick, extraSubs);
          // Drop no-op changes (target identical to current placement).
          const isNoop = acct && (acct.fs_section || null) === toSection && (acct.fs_subsection || null) === toSubsection;
          return {
            code,
            name: acct ? acct.name : '(unknown account)',
            type: acct ? acct.type : '',
            from_section: acct ? (acct.fs_section || null) : null,
            from_subsection: acct ? (acct.fs_subsection || null) : null,
            to_section: toSection,
            to_subsection: toSubsection,
            reason: String(c.reason || '').slice(0, 300),
            valid: v.ok && !isNoop,
            issue: isNoop ? 'no change (already placed there)' : v.issue,
          };
        }).filter(c => !(c.issue === 'no change (already placed there)' && !c.valid) || true); // keep all for transparency
        // Only keep new_subsections whose parent section is real for this entity.
        const keptSubs = newSubs
          .filter(s => Object.values(pick.byType).some(def => def.sections.some(sec => sec.section === s.section)))
          .map(s => ({ section: s.section, subsection: s.subsection, statement: statementForSection(pick, s.section) }));
        if (changes.length || keptSubs.length) {
          proposal = { changes, new_subsections: keptSubs, applicable: changes.some(c => c.valid) };
        }
      }

      res.json({ reply: stripCodeFences(fullText) || '(no response)', proposal, model: data.model || 'claude-haiku' });
    } catch (e) {
      res.status(500).json({ error: 'Assistant error: ' + e.message });
    }
  });

  // POST /api/entities/:eid/assistant/apply
  // Body: { changes:[{code, fs_section, fs_subsection}], new_subsections:[{section, subsection, statement}] }
  // Re-validates everything server-side, then writes fs_section/fs_subsection.
  app.post('/api/entities/:eid/assistant/apply', auth, requireEntityAccess(), requireRole('Admin', 'Accountant'), (req, res) => {
    try {
      const eid = +req.params.eid;
      const ent = db.prepare('SELECT id, name, code, entity_type FROM entities WHERE id=?').get(eid);
      if (!ent) return res.status(404).json({ error: 'Entity not found' });

      const reqChanges = Array.isArray(req.body && req.body.changes) ? req.body.changes : [];
      const reqSubs = Array.isArray(req.body && req.body.new_subsections) ? req.body.new_subsections : [];
      if (!reqChanges.length && !reqSubs.length) return res.status(400).json({ error: 'Nothing to apply.' });

      const accounts = db.prepare('SELECT code, name, type, subtype, fs_section, fs_subsection FROM accounts WHERE entity_id=?').all(eid);
      const acctByCode = new Map(accounts.map(a => [String(a.code), a]));
      const pick = buildPickList(db, eid, ent);

      const normSubs = reqSubs
        .filter(s => s && s.section && s.subsection)
        .map(s => ({ section: String(s.section), subsection: String(s.subsection), statement: String(s.statement || statementForSection(pick, String(s.section))) }))
        .filter(s => Object.values(pick.byType).some(def => def.sections.some(sec => sec.section === s.section)));
      const extraSubs = new Set(normSubs.map(s => s.section + '\u0000' + s.subsection));

      const results = [];
      let applied = 0;
      const run = db.transaction(() => {
        for (const s of normSubs) {
          db.prepare('INSERT OR IGNORE INTO statement_subsections (entity_id, statement, section, subsection, created_by) VALUES (?, ?, ?, ?, ?)')
            .run(eid, s.statement, s.section, s.subsection, (req.user && (req.user.email || req.user.name)) || 'assistant');
        }
        for (const c of reqChanges) {
          const code = String(c.code == null ? '' : c.code);
          const acct = acctByCode.get(code);
          const toSection = (c.fs_section === null || c.fs_section === '' || c.fs_section === undefined) ? null : String(c.fs_section);
          const toSubsection = (c.fs_subsection === null || c.fs_subsection === '' || c.fs_subsection === undefined) ? null : String(c.fs_subsection);
          const probe = acct ? Object.assign({}, acct, { _toSection: toSection, _toSubsection: toSubsection }) : null;
          const v = validateChange(probe, pick, extraSubs);
          if (!v.ok) { results.push({ code, ok: false, error: v.issue }); continue; }
          db.prepare('UPDATE accounts SET fs_section=?, fs_subsection=? WHERE entity_id=? AND code=?')
            .run(toSection, toSubsection, eid, code);
          results.push({ code, ok: true });
          applied++;
        }
      });
      run();
      res.json({ applied, subsections_created: normSubs.length, results });
    } catch (e) {
      res.status(500).json({ error: 'Apply error: ' + e.message });
    }
  });
}

module.exports = { registerReportAssistantRoutes };
