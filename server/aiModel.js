'use strict';
// Single source of truth for the Claude model CloudLedger calls.
//
// Why this exists: on 2026-02-19 Anthropic retired Claude Haiku 3.5, but the
// insurance-coverage extractor in otherworkpapers.js was still pinned to the
// `claude-3-5-haiku-latest` alias. That alias is family-locked -- it never
// advances past the 3.5 generation -- so it kept resolving to the retired
// model and every request 404'd (not_found_error) for ~8 months, silently,
// because the caller swallowed the error. Lesson: do NOT rely on a `-latest`
// alias to auto-upgrade across model generations.
//
// To move to a newer Haiku: set CLAUDE_HAIKU_MODEL in the Railway service env
// (effective on next restart -- no code change, no redeploy), or bump the
// pinned default below. Always keep the default a dated, known-good model id.
const HAIKU_MODEL = process.env.CLAUDE_HAIKU_MODEL || 'claude-haiku-4-5-20251001';

module.exports = { HAIKU_MODEL };
