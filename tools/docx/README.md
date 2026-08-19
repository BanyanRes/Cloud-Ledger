# tools/docx — house-style Word files

## The rule

**Every prose file delivered to Jimmy is a house-style `.docx`.**

Changelogs, plans, memos, reports, reviews, specs, letters, and **email / reply
drafts**. There is no category of prose that is exempt.

A raw `.md` opened in Word renders the `**` and `#` marks literally, in a
monospace font. Jimmy opens everything in Word. This has been hit three times
(2026-08-18 ×2, 2026-08-19) — the last one because a rule that enumerated
"changelog/memo/report" left room to argue an email draft was different. It
isn't.

The only non-`.docx` deliverables are things that genuinely are not prose:
`.xlsx` workbooks, `.pdf` packages, source code, images.

`save()` in `docs.js` enforces this — it throws on any extension but `.docx`.

## Setup

`docx` is not a server dependency, so it lives here and is installed separately.
Nothing under `tools/` is copied into the Docker image (see the `Dockerfile` —
it only copies `package.json`, `client/`, `server/`, `.env.example`), so this
has no effect on the Railway build.

```bash
cd tools/docx && npm install
```

## Writing a document

A build script is content only. Never re-derive typography, page size, table
widths, or the footer — they live in `house.js` and are already wired by
`buildDoc()`.

```js
const { buildDoc, changelogHead, save } = require('./docs.js');
const { P, H2, t, b, m, bullet, codePanel, resultsTable } = require('./blocks.js');

const children = [
  ...changelogHead({
    title: 'CloudLedger Change Log',
    subtitle: 'What changed, in one line',
    meta: [
      ['Date', 'August 19, 2026'],
      ['Reported by', 'Brad Tarter, CLA'],
      ['Commit', [m('77e8859'), t(' on '), m('main')]],
    ],
  }),
  H2('Problem'),
  P('Prose goes here.'),
  codePanel(['old = broken', 'new = fixed']),
  H2('Results'),
  resultsTable(['Check', 'Result'], [['Order', [t('deterministic')]]], [4680, 4680]),
  bullet('A bullet.'),
];

save(buildDoc(children), 'CloudLedger_Changelog_2026-08-19_Topic.docx');
```

For an email or reply draft, swap the head:

```js
const { buildDoc, replyHead, save } = require('./docs.js');
const { P, H2, signOff } = require('./blocks.js');

const children = [
  ...replyHead({
    subject: 'Silsbee — July Requisition Report',
    draftFor: 'Brad Tarter, CLA',
    to: 'Brad Tarter (CLA)',
    cc: 'Irvin Bermudez · Will Myers',
    subjectLine: 'RE: Silsbee - July Requisition Report',
    date: 'August 19, 2026',
  }),
  P('Brad,'),
  P('Body.'),
  H2('1. First talking point'),
  P('...'),
  ...signOff('Jimmy'),
];

save(buildDoc(children), 'CloudLedger_Reply_2026-08-19_Brad_Topic.docx');
```

## Verifying — not optional

A document that builds is not a document that looks right.

```bash
node tools/docx/verify.js CloudLedger_Changelog_2026-08-19_Topic.docx
```

It renders to PDF, rasterises every page, and prints the image paths. **Then
actually look at them.** Check page count, that the footer page numbers render
at a single size, that custom-styled paragraphs match body copy, and that
nothing is orphaned or running off the right margin.

## File naming

| Kind | Pattern |
|---|---|
| Changelog | `CloudLedger_Changelog_YYYY-MM-DD_<topic>.docx` |
| Plan | `CloudLedger_Plan_YYYY-MM-DD_<topic>.docx` |
| Reply / email draft | `CloudLedger_Reply_YYYY-MM-DD_<recipient>_<topic>.docx` |

## Approved reference documents

- Changelog: `CloudLedger_Changelog_2026-08-18_Bank_Rec_Zero_Balance.docx`
- Reply: `CloudLedger_Reply_2026-08-19_Brad_Silsbee_July_Req.docx`

## The house style itself

Full table of sizes, colours and the non-obvious rules that cost a rebuild each
time (custom styles must declare their own `run`; footer fields take their size
from the paragraph mark; tables need dual DXA widths; `ShadingType.CLEAR` never
`SOLID`; never `\n`) is in `house.js` and in the project doc
`claude/CLAUDE.md` → "Word document house style".
