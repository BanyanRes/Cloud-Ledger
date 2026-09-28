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

async function imageDims(buf, ext) {
  const d = await PDFDocument.create();
  const img = ext === 'png' ? await d.embedPng(buf) : await d.embedJpg(buf);
  return { w: img.width, h: img.height };
}

// ── Word (.docx) letterhead stamper ──────────────────────────────────────────
// Insert the Banyan logo into the uploaded .docx's page header so it appears at
// the top of every page, leaving the document's own content and formatting
// untouched. Returns the modified .docx as a Buffer.
const OOXML = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
};

async function applyLetterheadToDocx(docxBuf, logoBuf, logoExt) {
  const ext = logoExt === 'png' ? 'png' : 'jpg';
  let zip;
  try { zip = await JSZip.loadAsync(docxBuf); }
  catch { throw new Error('That file is not a readable .docx.'); }
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error('That .docx has no document body (word/document.xml missing).');
  let docXml = await docFile.async('string');
  if (!/<w:sectPr[\s>]/.test(docXml))
    throw new Error('That document has no section layout, so a header cannot be added. Open it in Word and re-save it, then try again.');

  // Logo image part (unique name so we never clobber the doc's own media).
  let mediaPath = 'word/media/lhbanyanlogo.' + ext, k = 1;
  while (zip.file(mediaPath)) mediaPath = 'word/media/lhbanyanlogo' + (k++) + '.' + ext;
  const mediaBase = mediaPath.split('/').pop();
  zip.file(mediaPath, logoBuf);

  // Logo sized to 2" wide, aspect preserved (EMU: 914400 per inch).
  const cx = 1828800;
  let cy = 557040;
  try { const d = await imageDims(logoBuf, ext); cy = Math.round(cx * (d.h / d.w)); } catch {}

  // Header part with a centered inline picture.
  let headerPath = 'word/lhbanyan_header.xml', j = 1;
  while (zip.file(headerPath)) headerPath = 'word/lhbanyan_header' + (j++) + '.xml';
  const headerBase = headerPath.split('/').pop();
  const headerXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:hdr xmlns:w="' + OOXML.w + '" xmlns:r="' + OOXML.r + '" xmlns:wp="' + OOXML.wp +
    '" xmlns:a="' + OOXML.a + '" xmlns:pic="' + OOXML.pic + '">' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:docPr id="1001" name="Banyan Letterhead"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="' + OOXML.pic + '">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="1001" name="banyan-logo"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>';
  zip.file(headerPath, headerXml);
  zip.file('word/_rels/' + headerBase + '.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + OOXML.r + '/image" Target="media/' + mediaBase + '"/></Relationships>');

  // Register the header part in the document's relationships.
  const relsPath = 'word/_rels/document.xml.rels';
  const relsFile = zip.file(relsPath);
  if (!relsFile) throw new Error('That .docx is missing its relationships part.');
  let rels = await relsFile.async('string');
  const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1], 10));
  const relId = 'rId' + ((ids.length ? Math.max(...ids) : 0) + 1);
  rels = rels.replace('</Relationships>',
    '<Relationship Id="' + relId + '" Type="' + OOXML.r + '/header" Target="' + headerBase + '"/></Relationships>');
  zip.file(relsPath, rels);

  // Content types: image default + header override.
  const ctPath = '[Content_Types].xml';
  let ct = await zip.file(ctPath).async('string');
  if (!new RegExp('Extension="' + ext + '"', 'i').test(ct))
    ct = ct.replace('</Types>', '<Default Extension="' + ext + '" ContentType="image/' + (ext === 'jpg' ? 'jpeg' : 'png') + '"/></Types>');
  if (!ct.includes('/word/' + headerBase))
    ct = ct.replace('</Types>', '<Override PartName="/word/' + headerBase +
      '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
  zip.file(ctPath, ct);

  // Point every section at our header, on every page: drop any existing header
  // references and the title-page flag so the one default header applies
  // throughout.
  docXml = docXml.replace(/<w:headerReference[^>]*\/>/g, '');
  docXml = docXml.replace(/<w:titlePg[^>]*\/>/g, '');
  docXml = docXml.replace(/(<w:sectPr[^>]*>)/g,
    '$1<w:headerReference w:type="default" r:id="' + relId + '"/>');
  zip.file('word/document.xml', docXml);

  // Even/odd header setting would hide the logo on even pages — remove it.
  const setFile = zip.file('word/settings.xml');
  if (setFile) {
    let s = await setFile.async('string');
    s = s.replace(/<w:evenAndOddHeaders[^>]*\/>/g, '');
    zip.file('word/settings.xml', s);
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
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

  // Add the Banyan letterhead to an uploaded Word (.docx): the logo goes into
  // the page header (top of every page); the document's own content is untouched.
  // Returns the modified .docx.
  app.post('/api/letterhead/apply', auth, requireRole('Admin', 'Accountant'),
    memUpload.single('file'), async (req, res) => {
      try {
        if (!req.file || !req.file.buffer || !req.file.buffer.length)
          return res.status(400).json({ error: 'Upload a Word (.docx) file.' });
        const nm = (req.file.originalname || '').toLowerCase();
        if (!nm.endsWith('.docx')) {
          if (nm.endsWith('.doc'))
            return res.status(422).json({ error: 'Old .doc files are not supported — open it in Word and Save As .docx first.' });
          return res.status(422).json({ error: 'Please upload a Word .docx file.' });
        }
        const logo = resolveLogo(templatesDir);
        if (!fs.existsSync(logo.path))
          return res.status(409).json({ error: 'No letterhead logo is installed.' });
        const ext = /\.png$/i.test(logo.path) ? 'png' : 'jpg';
        const out = await applyLetterheadToDocx(req.file.buffer, fs.readFileSync(logo.path), ext);
        const base = safeName((req.file.originalname || 'document.docx').replace(/\.docx$/i, '')) || 'Document';
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="' + base + ' - Letterhead.docx"');
        return res.send(out);
      } catch (e) {
        res.status(500).json({ error: e.message || 'Failed to add letterhead' });
      }
    });
}

module.exports = { buildLetterheadPdf, docxToText, applyLetterheadToDocx, registerLetterheadRoutes };
