// Self-test for tools/docx. Builds one of each document kind, and proves the
// delivery guard actually refuses a non-.docx path.
//
//   cd tools/docx && npm install && npm run selftest
const path = require('path');
const { buildDoc, changelogHead, replyHead, save } = require('../docs.js');
const { P, H2, t, b, m, bullet, codePanel, resultsTable, signOff } = require('../blocks.js');

const OUT = __dirname;
let failures = 0;
const ok = (cond, label) => { console.log((cond ? '  PASS  ' : '  FAIL  ') + label); if (!cond) failures++; };

(async () => {
  console.log('=== the guard refuses non-.docx ===');
  let threw = false;
  try { await save(buildDoc([P('x')]), path.join(OUT, 'nope.md')); }
  catch (e) { threw = /refusing to write/.test(e.message); }
  ok(threw, 'save() throws on a .md path');

  threw = false;
  try { await save(buildDoc([P('x')]), path.join(OUT, 'nope.txt')); }
  catch (e) { threw = /refusing to write/.test(e.message); }
  ok(threw, 'save() throws on a .txt path');

  console.log('\n=== resultsTable enforces DXA widths ===');
  threw = false;
  try { resultsTable(['A', 'B'], [['1', '2']], [1000, 1000]); }
  catch (e) { threw = /must sum to 9360/.test(e.message); }
  ok(threw, 'resultsTable throws when widths do not sum to the content width');
  ok(!!resultsTable(['A', 'B'], [['1', '2']]), 'resultsTable splits evenly when widths are omitted');

  console.log('\n=== builds one of each kind ===');
  await save(buildDoc([
    ...changelogHead({
      title: 'CloudLedger Change Log',
      subtitle: 'Self-test document — not a real change',
      meta: [['Date', 'August 19, 2026'], ['Commit', [m('0000000'), t(' on '), m('main')]]],
    }),
    H2('Problem'), P('Body copy at Calibri 11pt.'),
    codePanel(['before = broken', 'after  = fixed']),
    H2('Results'),
    resultsTable(['Check', 'Result'], [['Something', [t('passed')]], ['Another', [b('also passed')]]], [4680, 4680]),
    bullet('A house bullet.'),
  ]), path.join(OUT, '_selftest_changelog.docx'));

  await save(buildDoc([
    ...replyHead({
      subject: 'Self-test — Reply Layout',
      draftFor: 'Nobody, Nowhere LLP',
      to: 'Nobody (Nowhere LLP)',
      cc: 'Someone Else',
      date: 'August 19, 2026',
    }),
    P('Hello,'), P('This is the approved letter layout.'),
    H2('1. A talking point'), P('With a paragraph under it.'),
    ...signOff('Jimmy'),
  ]), path.join(OUT, '_selftest_reply.docx'));

  console.log('\n' + (failures ? failures + ' FAILED' : 'all guard checks passed'));
  console.log('Now render both files:  node ../verify.js _selftest_changelog.docx');
  process.exit(failures ? 1 : 0);
})();
