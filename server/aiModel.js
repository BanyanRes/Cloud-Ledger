'use strict';
// Single source of truth for the Claude model CloudLedger calls.
//
// History: on 2026-02-19 Anthropic retired Claude Haiku 3.5, but the
// insurance-coverage extractor in otherworkpapers.js was still pinned to the
// `claude-3-5-haiku-latest` alias. That alias is family-locked -- it never
// advances past the 3.5 generation -- so it kept resolving to the retired
// model and every request 404'd (not_found_error) for ~8 months, silently.
// A dated pin (the fix after that) has the same failure mode, just later.
//
// So the model is no longer hardcoded:
//   * At startup and every 24h we ask the Models API (GET /v1/models) for every
//     available model, keep the ones whose `line` is "haiku", and use the newest
//     by created_at. New Haiku releases are picked up with no code change.
//   * anthropicFetch() wraps every Messages call. If Anthropic answers 404
//     not_found (model retired mid-day), it refreshes the list immediately and
//     retries once on the newest Haiku, so callers never see the retirement.
//   * If the lookup fails, we keep the last good model and retry in 10 minutes.
//
// Env overrides (Railway service vars, effective on restart):
//   CLAUDE_HAIKU_MODEL   pin an exact model id and disable auto-selection
//   CLAUDE_MODEL_LINE    model line to track (default "haiku")
//   CLAUDE_FALLBACK_MODEL model used before the first successful lookup

const PINNED = (process.env.CLAUDE_HAIKU_MODEL || '').trim();
const LINE = (process.env.CLAUDE_MODEL_LINE || 'haiku').trim().toLowerCase();
const FALLBACK = (process.env.CLAUDE_FALLBACK_MODEL || 'claude-haiku-5-5').trim();
const MESSAGES_URL = 'https://api.anthropic.com/v1/messages';
const MODELS_URL = 'https://api.anthropic.com/v1/models';
const REFRESH_MS = 24 * 3600 * 1000;
const RETRY_MS = 10 * 60 * 1000;

let current = PINNED || FALLBACK;
let source = PINNED ? 'pinned' : 'fallback';
let nextCheckAt = 0;
let inflight = null;

// Synchronous: the model to use right now (pinned, last auto-resolved, or fallback).
function getHaikuModel() {
  if (!PINNED && Date.now() >= nextCheckAt && !inflight) refreshHaikuModel().catch(() => {});
  return current;
}

function modelStatus() {
  return { model: current, source, line: LINE, pinned: !!PINNED, nextCheckAt: nextCheckAt ? new Date(nextCheckAt).toISOString() : null };
}

async function listModels(apiKey, fetchImpl) {
  const doFetch = fetchImpl || globalThis.fetch;
  const out = [];
  let after = null;
  for (let page = 0; page < 10; page++) {
    const url = MODELS_URL + '?limit=100' + (after ? '&after_id=' + encodeURIComponent(after) : '');
    const res = await doFetch(url, { headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' } });
    if (!res.ok) throw new Error('Models API ' + res.status + ': ' + (await res.text()).slice(0, 200));
    const j = await res.json();
    out.push(...(j.data || []));
    if (!j.has_more || !j.last_id) break;
    after = j.last_id;
  }
  return out;
}

// Re-resolve the newest model in LINE. force=true bypasses the 24h cache.
async function refreshHaikuModel({ force = false, apiKey = process.env.ANTHROPIC_API_KEY, fetchImpl } = {}) {
  if (PINNED || !apiKey) return current;
  if (!force && Date.now() < nextCheckAt) return current;
  if (inflight) return inflight;
  inflight = (async () => {
    try {
      const models = await listModels(apiKey, fetchImpl);
      const inLine = models.filter(m => (m.line || '').toLowerCase() === LINE ||
        (!m.line && String(m.id || '').toLowerCase().includes('-' + LINE + '-')));
      inLine.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
      if (inLine.length) {
        if (inLine[0].id !== current) console.log('[aiModel] using ' + inLine[0].id + ' (newest "' + LINE + '", was ' + current + ')');
        current = inLine[0].id;
        source = 'auto';
      } else {
        console.warn('[aiModel] Models API listed no "' + LINE + '" models; keeping ' + current);
      }
      nextCheckAt = Date.now() + REFRESH_MS;
    } catch (e) {
      console.warn('[aiModel] model lookup failed, keeping ' + current + ': ' + e.message);
      nextCheckAt = Date.now() + RETRY_MS;
    } finally {
      inflight = null;
    }
    return current;
  })();
  return inflight;
}

async function isModelGone(res) {
  if (res.status !== 404) return false;
  try { return /not_found_error|model/i.test(await res.clone().text()); } catch (_e) { return true; }
}

// Drop-in replacement for fetch() on the Messages endpoint. If the request body
// omits `model` (or sets it to 'auto'), the current Haiku is filled in. A 404 for
// a retired model triggers one refresh + retry, unless the model is pinned.
async function anthropicFetch(url, init = {}, fetchImpl) {
  const doFetch = fetchImpl || globalThis.fetch;
  let body;
  try { body = typeof init.body === 'string' ? JSON.parse(init.body) : null; } catch (_e) { body = null; }
  const auto = body && (!body.model || body.model === 'auto');
  if (!auto) return doFetch(url, init);

  if (!PINNED && !nextCheckAt) await refreshHaikuModel();
  const send = model => doFetch(url, { ...init, body: JSON.stringify({ ...body, model }) });
  const tried = getHaikuModel();
  const res = await send(tried);
  if (PINNED || !(await isModelGone(res))) return res;

  console.warn('[aiModel] ' + tried + ' unavailable (404); refreshing model list and retrying');
  const apiKey = (init.headers && (init.headers['x-api-key'] || init.headers['X-Api-Key'])) || process.env.ANTHROPIC_API_KEY;
  const fresh = await refreshHaikuModel({ force: true, apiKey });
  return fresh !== tried ? send(fresh) : res;
}

// Back-compat: anything still reading HAIKU_MODEL gets the live value.
module.exports = { getHaikuModel, refreshHaikuModel, anthropicFetch, modelStatus, MESSAGES_URL };
Object.defineProperty(module.exports, 'HAIKU_MODEL', { enumerable: true, get: getHaikuModel });
