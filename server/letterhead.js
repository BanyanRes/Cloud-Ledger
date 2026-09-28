// ═══════════════════════════════════════════════════════════════════════════
// Banyan Residential Letterhead — fillable PDF generator (Administration)
//
// Produces a Banyan Residential letterhead as a fillable PDF: the Banyan logo
// centered at the top, then interactive form fields (Date, To, Re, a large
// multi-line Body, and a signature block) the user can type into and save.
//
// When a content file is uploaded (.txt or .docx), its text is reflowed into
// the Body field so the letter comes back pre-filled on the letterhead, still
// editable in any PDF viewer.
//
//   • Logo   → server/assets/banyan-letterhead-logo.jpg is the built-in default.
//              An Admin can upload a replacement (an image, or a .docx whose
//              first image is used) which is stored on the persistent data
//              volume under <dataDir>/templates and reused thereafter.
//   • Output → a single-page US-Letter (612×792 pt) fillable PDF (AcroForm).
//
// No third-party services; everything is built in-process with pdf-lib + JSZip.
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

const DEFAULT_LOGO = path.join(__dirname, 'assets', 'banyan-letterhead-logo.jpg');
// Custom (uploaded) logo lives on the data volume. We keep the original bytes
// under a fixed base name; the extension records the image kind for embedding.
const CUSTOM_LOGO_BASENAME = 'letterhead_logo';

// ── helpers ────────────────────────────────────────────────────────────────

function decodeEntities(s) {
  return String(s == null ? '' : s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

// Plain-text extraction from a .docx buffer: paragraphs -> newlines, tabs/breaks
// preserved, all other tags stripped. Good enough to reflow letter body copy.
async function docxToText(buf) {
  let zip;
  try { zip = await JSZip.loadAsync(buf); }
  catch { throw new Error('That file is not a readable .docx.'); }
  const f = zip.file('word/document.xml');
  if (!f) throw new Error('That .docx has no document body (word/document.xml missing).');
  let xml = await f.async('string');
  xml = xml
    .replace(/<w:tab[^>]*\/?>/g, '\t')
    .replace(/<w:br[^>]*\/?>/g, '\n')
    .replace(/<\/w:p>/g, '\n')
    .replace(/<[^>]+>/g, '');
  return decodeEntities(xml).replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim();
}

// Pull the first embedded image out of a .docx (used when an Admin uploads a
// Word letterhead instead of a bare image).
async function firstImageFromDocx(buf) {
  let zip;
  try { zip = await JSZip.loadAsync(buf); }
  catch { return null; }
  const names = Object.keys(zip.files)
    .filter(n => /^word\/media\/.+\.(png|jpe?g)$/i.test(n))
    .sort();
  if (!names.length) return null;
  const name = names[0];
  const data = await zip.file(name).async('nodebuffer');
  return { data, ext: /\.png$/i.test(name) ? 'png' : 'jpg' };
}

function isPng(buf) {
  return buf && buf.length > 8 &&
    buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47;
}
function isJpg(buf) {
  return buf && buf.length > 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff;
}

// Resolve the active logo: a custom uploaded one if present, else the default.
function resolveLogo(templatesDir) {
  for (const ext of ['png', 'jpg', 'jpeg']) {
    const p = path.join(templatesDir, CUSTOM_LOGO_BASENAME + '.' + ext);
    if (fs.existsSync(p)) return { path: p, source: 'custom' };
  }
  return { path: DEFAULT_LOGO, source: 'default' };
}

function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
}

// ── PDF builder ─────────────────────────────────────────────────────────────

// Build the fillable letterhead PDF. `fields` pre-fills any of date/to/re/body/
// signerName/signerTitle; anything omitted becomes an empty fillable field.
async function buildLetterheadPdf(logoBuf, fields = {}) {
  const pdf = await PDFDocument.create();
  const W = 612, H = 792;                 // US Letter, points
  const L = 72, R = W - 72;               // 1" left/right margins
  const CW = R - L;                       // content width
  const page = pdf.addPage([W, H]);
  const helv = await pdf.embedFont(StandardFonts.Helvetica);
  const helvBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const labelColor = rgb(0.42, 0.42, 0.42);
  const lineColor = rgb(0.78, 0.72, 0.55); // muted gold to match the brand

  // Logo, centered near the top.
  let img;
  if (isPng(logoBuf)) img = await pdf.embedPng(logoBuf);
  else img = await pdf.embedJpg(logoBuf);
  const logoW = 200;
  const logoH = logoW * (img.height / img.width);
  const logoTop = H - 54;                  // 0.75" from the top edge
  page.drawImage(img, { x: (W - logoW) / 2, y: logoTop - logoH, width: logoW, height: logoH });

  // Thin divider under the logo.
  const dividerY = logoTop - logoH - 16;
  page.drawLine({ start: { x: L, y: dividerY }, end: { x: R, y: dividerY }, thickness: 1, color: lineColor });

  const form = pdf.getForm();
  const label = (text, x, y) => page.drawText(text, { x, y, size: 9, font: helvBold, color: labelColor });

  const mkText = (name, x, y, w, h, value, opts = {}) => {
    const tf = form.createTextField('letterhead.' + name);
    if (value) tf.setText(String(value));
    if (opts.multiline) tf.enableMultiline();
    tf.addToPage(page, {
      x, y, width: w, height: h,
      borderWidth: opts.border == null ? 0 : opts.border,
      borderColor: rgb(0.85, 0.85, 0.85),
    });
    tf.setFontSize(opts.size || 11);
    return tf;
  };

  // Date / To / Re — a single label column with fields beside them.
  let y = dividerY - 26;
  const rowGap = 26, fieldH = 16;
  label('DATE', L, y + 3);
  mkText('date', L + 42, y, 190, fieldH, fields.date, { border: 1 });
  y -= rowGap;
  label('TO', L, y + 3);
  mkText('to', L + 42, y, CW - 42, fieldH, fields.to, { border: 1 });
  y -= rowGap;
  label('RE', L, y + 3);
  mkText('re', L + 42, y, CW - 42, fieldH, fields.re, { border: 1 });

  // Body — large multi-line field the uploaded content flows into.
  const bodyTop = y - 18;
  const bodyBottom = 150;                   // leave room for the signature block
  label('BODY', L, bodyTop + 6);
  mkText('body', L, bodyBottom, CW, bodyTop - bodyBottom, fields.body, { multiline: true, border: 1, size: 11 });

  // Signature block.
  let sy = bodyBottom - 26;
  page.drawText('Sincerely,', { x: L, y: sy, size: 11, font: helv, color: rgb(0, 0, 0) });
  sy -= 40;
  mkText('signerName', L, sy, 260, fieldH, fields.signerName, { border: 1 });
  label('NAME', L, sy - 12);
  sy -= 34;
  mkText('signerTitle', L, sy, 260, fieldH, fields.signerTitle, { border: 1 });
  label('TITLE', L, sy - 12);

  // Make the empty look clean: appearance uses Helvetica.
  form.updateFieldAppearances(helv);
  return Buffer.from(await pdf.save());
}

// ── routes ──────────────────────────────────────────────────────────────────

function registerLetterheadRoutes(app, ctx) {
  const { auth, requireRole, memUpload, templatesDir } = ctx;
  try { fs.mkdirSync(templatesDir, { recursive: true }); } catch {}

  function logoStatus() {
    const r = resolveLogo(templatesDir);
    try {
      const st = fs.statSync(r.path);
      return { installed: true, source: r.source, size: st.size, updated_at: st.mtime.toISOString() };
    } catch { return { installed: false, source: r.source }; }
  }

  // Current logo status (drives the Letterhead admin card).
  app.get('/api/letterhead/logo/status', auth, requireRole('Admin', 'Accountant'),
    (req, res) => res.json(logoStatus()));

  // Preview the active logo image itself.
  app.get('/api/letterhead/logo/preview', auth, requireRole('Admin', 'Accountant'),
    (req, res) => {
      const r = resolveLogo(templatesDir);
      if (!fs.existsSync(r.path)) return res.status(404).json({ error: 'No logo installed' });
      res.setHeader('Content-Type', /\.png$/i.test(r.path) ? 'image/png' : 'image/jpeg');
      res.setHeader('Cache-Control', 'no-store');
      return res.send(fs.readFileSync(r.path));
    });

  // Upload/replace the letterhead logo (Admin only). Accepts an image, or a
  // .docx whose first embedded image is used.
  app.post('/api/letterhead/logo', auth, requireRole('Admin'),
    memUpload.single('file'), async (req, res) => {
      try {
        if (!req.file || !req.file.buffer || !req.file.buffer.length)
          return res.status(400).json({ error: 'No file uploaded' });
        let buf = req.file.buffer;
        let ext;
        const nameLower = (req.file.originalname || '').toLowerCase();
        if (nameLower.endsWith('.docx') || (!isPng(buf) && !isJpg(buf))) {
          const img = await firstImageFromDocx(buf);
          if (!img) return res.status(422).json({ error: 'Upload a PNG/JPG image, or a .docx that contains the logo image.' });
          buf = img.data; ext = img.ext;
        } else {
          ext = isPng(buf) ? 'png' : 'jpg';
        }
        // Sanity check it embeds.
        try {
          const test = await PDFDocument.create();
          if (ext === 'png') await test.embedPng(buf); else await test.embedJpg(buf);
        } catch { return res.status(422).json({ error: 'That image could not be read as a PNG or JPG.' }); }
        // Remove any prior custom logo, then write the new one.
        for (const e of ['png', 'jpg', 'jpeg']) {
          const p = path.join(templatesDir, CUSTOM_LOGO_BASENAME + '.' + e);
          try { fs.unlinkSync(p); } catch {}
        }
        fs.writeFileSync(path.join(templatesDir, CUSTOM_LOGO_BASENAME + '.' + ext), buf);
        res.json({ ok: true, status: logoStatus() });
      } catch (e) {
        res.status(500).json({ error: e.message || 'Upload failed' });
      }
    });

  // Revert to the built-in Banyan logo (Admin only).
  app.delete('/api/letterhead/logo', auth, requireRole('Admin'), (req, res) => {
    for (const e of ['png', 'jpg', 'jpeg']) {
      const p = path.join(templatesDir, CUSTOM_LOGO_BASENAME + '.' + e);
      try { fs.unlinkSync(p); } catch {}
    }
    res.json({ ok: true, status: logoStatus() });
  });

  // Generate the fillable letterhead PDF. Optional content file (.txt/.docx)
  // reflows into the Body; optional form fields pre-fill the rest.
  app.post('/api/letterhead/generate', auth, requireRole('Admin', 'Accountant'),
    memUpload.single('file'), async (req, res) => {
      try {
        const b = req.body || {};
        let body = typeof b.body === 'string' ? b.body : '';
        // If a content file was uploaded, it wins as the body source.
        if (req.file && req.file.buffer && req.file.buffer.length) {
          const nm = (req.file.originalname || '').toLowerCase();
          if (nm.endsWith('.docx')) body = await docxToText(req.file.buffer);
          else if (nm.endsWith('.txt') || nm.endsWith('.md') || nm.endsWith('.csv') || (req.file.mimetype || '').startsWith('text/'))
            body = req.file.buffer.toString('utf8').replace(/\r/g, '').trim();
          else if (nm.endsWith('.doc'))
            return res.status(422).json({ error: 'Old .doc files are not supported — save it as .docx or paste the text.' });
          else
            return res.status(422).json({ error: 'Upload a .txt or .docx file (or type the text in the Body box).' });
        }

        const logo = resolveLogo(templatesDir);
        if (!fs.existsSync(logo.path))
          return res.status(409).json({ error: 'No letterhead logo is installed.' });
        const logoBuf = fs.readFileSync(logo.path);

        const pdfBuf = await buildLetterheadPdf(logoBuf, {
          date: b.date, to: b.to, re: b.re, body,
          signerName: b.signerName, signerTitle: b.signerTitle,
        });

        const reTag = safeName(b.re);
        const fname = 'Banyan Letterhead' + (reTag ? ' - ' + reTag : '') + '.pdf';
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename="' + fname + '"');
        return res.send(pdfBuf);
      } catch (e) {
        res.status(500).json({ error: e.message || 'Generation failed' });
      }
    });
}

module.exports = { buildLetterheadPdf, docxToText, registerLetterheadRoutes };
