// CloudLedger — document assembly + the delivery guard.
//
// buildDoc()      wires styles / page / footer so no build script re-derives them
// changelogHead() the approved changelog opening
// replyHead()     the approved letter / email-draft opening
// save()          writes the file AND REFUSES any extension other than .docx
const { Document, Packer } = require('docx');
const fs = require('fs');
const path = require('path');
const H = require('./house.js');
const { P, H1, t, metaBlock } = require('./blocks.js');

// Every house document is built through here.
function buildDoc(children) {
  return new Document({
    styles: H.styles,
    numbering: H.numbering,
    sections: [{
      properties: { page: H.PAGE },
      footers: { default: H.footer() },
      children,
    }],
  });
}

// Changelog / plan / report opening:
//   Title1 · bold one-line subject · metaBlock · rule
// meta = [['Date', 'August 19, 2026'], ['Commit', [m('abc1234'), t(' on main')]], ...]
function changelogHead({ title, subtitle, meta = [] }) {
  const out = [H1(title)];
  if (subtitle) out.push(P([t(subtitle, { bold: true, size: 24, color: H.C.body })], { spacing: { after: 200 } }));
  if (meta.length) out.push(metaBlock(meta));
  out.push(H.rule());
  return out;
}

// Letter / email-draft opening:
//   Title1 = the subject matter
//   bold one-liner "Draft reply to <name>, <firm>"
//   metaBlock of To / Cc / Subject / Date
//   rule
function replyHead({ subject, draftFor, to, cc, subjectLine, date }) {
  const rows = [['To', to]];
  if (cc) rows.push(['Cc', cc]);
  rows.push(['Subject', subjectLine || ('RE: ' + subject)]);
  if (date) rows.push(['Date', date]);
  return [
    H1(subject),
    P([t('Draft reply to ' + draftFor, { bold: true, size: 24, color: H.C.body })], { spacing: { after: 200 } }),
    metaBlock(rows, { labelWidth: 1000 }),
    H.rule(),
  ];
}

// THE GUARD. Prose deliverables for Jimmy are .docx — always, including email and
// reply drafts. A .md or .txt opened in Word shows the ** and # marks in a
// monospace font, which is exactly the bug this tooling exists to prevent. If you
// find yourself wanting an exception, there isn't one.
const ALLOWED = /\.docx$/i;
function save(doc, outPath) {
  if (!ALLOWED.test(outPath)) {
    throw new Error(
      'save(): refusing to write "' + path.basename(outPath) + '".\n' +
      'Prose deliverables are .docx. Markdown/plain text opened in Word renders the\n' +
      '** and # marks literally. See tools/docx/README.md.'
    );
  }
  return Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outPath, buf);
    console.log('wrote ' + outPath + ' (' + buf.length + ' bytes)');
    console.log('NOT DONE YET — run:  node tools/docx/verify.js "' + outPath + '"');
    return outPath;
  });
}

module.exports = { buildDoc, changelogHead, replyHead, save };
