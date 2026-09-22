// ═══════════════════════════════════════════════════════════════════════════
// Assignment of Interest — document generator
//
// Given a Fund, Assignor, Assignee and Effective Date, produces the executed-
// ready assignment paperwork:
//   • ALL funds  → "Short Form Assignment and Assumption Agreement" (.docx),
//                  filled from a placeholder template.
//   • CLRF only  → additionally the CLRF Subscription Documents (.pdf) with the
//                  assignee (incoming investor) name stamped on the cover and the
//                  two exhibit dividers.
//
// Templates are not shipped in the repo. They live on the persistent data volume
// under <dataDir>/templates and are uploaded once through the Assignment page
// (Admin only), so legal can swap a revised template without a code deploy.
//   assignment   → assignment_template.docx   (must contain {{PLACEHOLDER}} tokens)
//   subscription → clrf_subscription_template.pdf
//
// The fill logic here is a straight port of the sandbox-validated routines:
//   - docx: JSZip string-replace of {{TOKENS}} in word/document.xml (the template
//     is authored so every token is a single contiguous run).
//   - pdf: pdf-lib text overlay of the investor name at fixed coordinates on the
//     cover (page 0) and the Exhibit A / Exhibit B dividers (pages 15 / 32).
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

const ASSIGNMENT_FILE = 'assignment_template.docx';
const SUBSCRIPTION_FILE = 'clrf_subscription_template.pdf';

// Placeholders the assignment template must contain (used to fill and to validate
// an uploaded template is the right document).
const REQUIRED_TOKENS = [
  'FUND_NAME', 'EFFECTIVE_DATE',
  'ASSIGNOR_NAME', 'ASSIGNOR_ARTICLE', 'ASSIGNOR_TYPE',
  'ASSIGNEE_NAME', 'ASSIGNEE_ARTICLE', 'ASSIGNEE_TYPE',
  'INTEREST_TYPE', 'GOVERNING_LAW',
  'ASSIGNOR_SIGNATORY', 'ASSIGNOR_TITLE', 'SIG_DATE',
];

// Entity-type → legal descriptor, indefinite article, and default interest label.
const TYPES = {
  individual:  { desc: 'individual',                   article: 'an', interest: 'Limited Partner' },
  llc:         { desc: 'limited liability company',    article: 'a',  interest: 'Member' },
  lp:          { desc: 'limited partnership',          article: 'a',  interest: 'Limited Partner' },
  corporation: { desc: 'corporation',                  article: 'a',  interest: 'Shareholder' },
  scorp:       { desc: 'S corporation',                article: 'an', interest: 'Shareholder' },
  partnership: { desc: 'partnership',                  article: 'a',  interest: 'Partner' },
  trust:       { desc: 'trust',                        article: 'a',  interest: 'Beneficiary' },
  ira:         { desc: 'individual retirement account', article: 'an', interest: 'Limited Partner' },
};

// Investor-name fill locations in the CLRF subscription template (612x792 pt
// pages, pdf-lib origin bottom-left). Cover "Name of Investor:____" line, then
// under the right-aligned "Name of Investor" header on each exhibit divider.
const NAME_SLOTS = [
  { page: 0,  align: 'left',  x: 157, y: 696 },
  { page: 15, align: 'right', x: 540, y: 699 },
  { page: 32, align: 'right', x: 540, y: 699 },
];

function xmlEsc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function articleFor(desc) { return /^[aeiou]/i.test(String(desc).trim()) ? 'an' : 'a'; }

function resolveType(typeKey, customDesc) {
  if (typeKey && TYPES[typeKey]) return { ...TYPES[typeKey] };
  const desc = (customDesc || 'entity').trim();
  return { desc, article: articleFor(desc), interest: 'Limited Partner' };
}

function isClrfName(fundName) {
  return /county\s+line\s+rail\s+fund\s+i\b/i.test(String(fundName || ''));
}

// Fill the assignment .docx template buffer with the supplied data.
async function fillAssignment(templateBuf, d) {
  const ar = resolveType(d.assignorType, d.assignorTypeCustom);
  const ae = resolveType(d.assigneeType, d.assigneeTypeCustom);
  const map = {
    EFFECTIVE_DATE:     d.effectiveDate || '',
    FUND_NAME:          d.fundName || '',
    ASSIGNOR_NAME:      d.assignorName || '',
    ASSIGNOR_ARTICLE:   ar.article,
    ASSIGNOR_TYPE:      ar.desc,
    ASSIGNEE_NAME:      d.assigneeName || '',
    ASSIGNEE_ARTICLE:   ae.article,
    ASSIGNEE_TYPE:      ae.desc,
    INTEREST_TYPE:      d.interestType || ar.interest || 'Limited Partner',
    GOVERNING_LAW:      d.governingLaw || 'Delaware',
    ASSIGNOR_SIGNATORY: (d.assignorSignatory && d.assignorSignatory.trim()) || d.assignorName || '',
    ASSIGNOR_TITLE:     d.assignorTitle || '',
    SIG_DATE:           d.sigDate || '',
  };
  const zip = await JSZip.loadAsync(templateBuf);
  const docXml = zip.file('word/document.xml');
  if (!docXml) throw new Error('Template is not a valid .docx (missing word/document.xml)');
  let xml = await docXml.async('string');
  for (const [k, v] of Object.entries(map)) {
    xml = xml.split('{{' + k + '}}').join(xmlEsc(v));
  }
  const left = xml.match(/\{\{[A-Z_]+\}\}/g);
  if (left) throw new Error('Template still has unfilled placeholders: ' + [...new Set(left)].join(', '));
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// Overlay the incoming investor (assignee) name on the CLRF subscription template.
async function fillSubscription(templateBuf, investorName) {
  const pdf = await PDFDocument.load(templateBuf);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const size = 11;
  const pages = pdf.getPages();
  const name = String(investorName || '');
  for (const s of NAME_SLOTS) {
    const pg = pages[s.page];
    if (!pg) continue; // template shorter than expected — skip missing slot
    const w = font.widthOfTextAtSize(name, size);
    const x = s.align === 'right' ? s.x - w : s.x;
    pg.drawText(name, { x, y: s.y, size, font, color: rgb(0, 0, 0) });
  }
  return Buffer.from(await pdf.save());
}

// Windows/Excel-safe filename fragment.
function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Party';
}

function registerAssignmentRoutes(app, ctx) {
  const { auth, requireRole, memUpload, templatesDir } = ctx;
  try { fs.mkdirSync(templatesDir, { recursive: true }); } catch {}

  const templatePath = (kind) =>
    path.join(templatesDir, kind === 'subscription' ? SUBSCRIPTION_FILE : ASSIGNMENT_FILE);

  function statusFor(kind) {
    const p = templatePath(kind);
    try {
      const st = fs.statSync(p);
      return { installed: true, size: st.size, updated_at: st.mtime.toISOString() };
    } catch { return { installed: false }; }
  }

  // Which templates are installed (drives the page's Templates card).
  app.get('/api/assignments/templates/status', auth, requireRole('Admin', 'Accountant'),
    (req, res) => {
      res.json({ assignment: statusFor('assignment'), subscription: statusFor('subscription') });
    });

  // Upload/replace a template (Admin only). Validates the file is the right kind.
  app.post('/api/assignments/templates/:kind', auth, requireRole('Admin'),
    memUpload.single('file'), async (req, res) => {
      try {
        const kind = req.params.kind;
        if (kind !== 'assignment' && kind !== 'subscription')
          return res.status(400).json({ error: 'Unknown template kind' });
        if (!req.file || !req.file.buffer || !req.file.buffer.length)
          return res.status(400).json({ error: 'No file uploaded' });
        const buf = req.file.buffer;

        if (kind === 'assignment') {
          // Must be a .docx carrying the required placeholder tokens.
          let xml;
          try {
            const zip = await JSZip.loadAsync(buf);
            const f = zip.file('word/document.xml');
            xml = f ? await f.async('string') : '';
          } catch { return res.status(422).json({ error: 'File is not a valid .docx' }); }
          const missing = REQUIRED_TOKENS.filter(t => !xml.includes('{{' + t + '}}'));
          if (missing.length)
            return res.status(422).json({
              error: 'This .docx is missing required placeholders: ' +
                missing.map(m => '{{' + m + '}}').join(', ') +
                '. Upload the placeholder assignment template.',
            });
        } else {
          // Must be a loadable PDF.
          try { await PDFDocument.load(buf); }
          catch { return res.status(422).json({ error: 'File is not a valid .pdf' }); }
        }

        fs.writeFileSync(templatePath(kind), buf);
        res.json({ ok: true, kind, status: statusFor(kind) });
      } catch (e) {
        res.status(500).json({ error: e.message || 'Upload failed' });
      }
    });

  // Generate the assignment paperwork.
  app.post('/api/assignments/generate', auth, requireRole('Admin', 'Accountant'),
    async (req, res) => {
      try {
        const b = req.body || {};
        const required = ['fundName', 'assignorName', 'assigneeName', 'effectiveDate'];
        const miss = required.filter(k => !String(b[k] || '').trim());
        if (miss.length) return res.status(400).json({ error: 'Missing required fields: ' + miss.join(', ') });

        const assignmentTpl = templatePath('assignment');
        if (!fs.existsSync(assignmentTpl))
          return res.status(409).json({ error: 'Assignment template not installed. Upload it in the Templates section.' });

        const clrf = !!b.isCLRF || isClrfName(b.fundName);
        const docxBuf = await fillAssignment(fs.readFileSync(assignmentTpl), b);
        const baseName = safeName(b.assignorName) + ' to ' + safeName(b.assigneeName);

        // Non-CLRF (or CLRF without a subscription template): return the .docx alone.
        const subTpl = templatePath('subscription');
        if (!clrf || !fs.existsSync(subTpl)) {
          if (clrf && !fs.existsSync(subTpl)) res.setHeader('X-Subscription-Missing', '1');
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.setHeader('Content-Disposition',
            'attachment; filename="Assignment of Interest - ' + baseName + '.docx"');
          res.setHeader('X-Assignment-Summary', JSON.stringify({ fund: b.fundName, clrf, subscription: false }));
          return res.send(docxBuf);
        }

        // CLRF: bundle the assignment .docx and the filled subscription .pdf.
        const pdfBuf = await fillSubscription(fs.readFileSync(subTpl), b.assigneeName);
        const zip = new JSZip();
        zip.file('Assignment of Interest - ' + baseName + '.docx', docxBuf);
        zip.file('Subscription Agreement - ' + safeName(b.assigneeName) + '.pdf', pdfBuf);
        const zipBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition',
          'attachment; filename="Assignment Package - ' + baseName + '.zip"');
        res.setHeader('X-Assignment-Summary', JSON.stringify({ fund: b.fundName, clrf: true, subscription: true }));
        return res.send(zipBuf);
      } catch (e) {
        res.status(500).json({ error: e.message || 'Generation failed' });
      }
    });
}

module.exports = { fillAssignment, fillSubscription, isClrfName, registerAssignmentRoutes };
