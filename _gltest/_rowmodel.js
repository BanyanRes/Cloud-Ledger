// Shared helper: read a rendered PDF back as a ROW MODEL — one entry per drawn
// baseline, with the row's label, its money cells, and whether a "$" was drawn
// on that baseline (and where).
//
// Why coordinates and not text order: layout.row() draws each value BEFORE its
// "$", so in extraction order a row's "$" trails its own figures and reads as if
// it belonged to the next line. Grouping by y removes the ambiguity.
const MONEY = /^\(?-?[\d,]+\.\d{2}\)?$/;   // 1,234.56 / (1,234.56) / -1,234.56
const DASH = /^-$/;                        // a zero cell renders as a dash

async function readRows(bytes) {
  const pdfParse = require('pdf-parse');
  const pages = [];
  await pdfParse(Buffer.from(bytes), {
    pagerender: (pd) => pd.getTextContent().then(tc => {
      const items = tc.items
        .filter(i => i.str && i.str.trim())
        .map(i => ({ s: i.str.trim(), x: i.transform[4], y: i.transform[5], w: i.width || 0 }));
      // Group into baselines (1.2pt tolerance).
      const rows = [];
      for (const it of items) {
        let r = rows.find(r => Math.abs(r.y - it.y) < 1.2);
        if (!r) { r = { y: it.y, items: [] }; rows.push(r); }
        r.items.push(it);
      }
      for (const r of rows) {
        r.items.sort((a, b) => a.x - b.x);
        r.money = r.items.filter(i => MONEY.test(i.s) || DASH.test(i.s));
        r.dollar = r.items.find(i => i.s === '$') || null;
        // Label = the leftmost run of non-money, non-$ items.
        r.label = r.items
          .filter(i => !MONEY.test(i.s) && !DASH.test(i.s) && i.s !== '$')
          .map(i => i.s).join(' ');
        r.text = r.items.map(i => i.s).join(' ');
        // Right edge of the label text (rightmost non-money, non-$ item).
        const labelItems = r.items.filter(i => !MONEY.test(i.s) && !DASH.test(i.s) && i.s !== '$');
        r.labelRight = labelItems.length
          ? Math.max.apply(null, labelItems.map(i => i.x + i.w))
          : null;
      }
      rows.sort((a, b) => b.y - a.y);   // top of page down
      pages.push(rows);
      return items.map(i => i.s).join(' ');
    }),
  });
  return pages;
}

const pageText = (rows) => rows.map(r => r.text).join(' ');

// Find a FIGURE row (one that actually carries money cells) whose label starts
// with `label`, across every page in `pages`. `which` picks among matches:
// 'first' (default) or 'last'.
function figureRow(pages, label, which = 'first') {
  const hits = [];
  pages.forEach((rows, pi) => {
    rows.forEach(r => {
      if (r.money.length && r.label.startsWith(label)) hits.push(Object.assign({ page: pi }, r));
    });
  });
  if (!hits.length) return null;
  return which === 'last' ? hits[hits.length - 1] : hits[0];
}

// Assert on the "$" for a labeled figure row. Returns { ok, why, detail }.
function dollarOn(pages, label, want, which) {
  const r = figureRow(pages, label, which);
  if (!r) return { ok: false, why: 'no figure row labeled "' + label + '"' };
  const has = !!r.dollar;
  if (want && !has) return { ok: false, why: '$ missing (row: ' + r.text.slice(0, 70) + ')' };
  if (!want && has) return { ok: false, why: 'unexpected $ at x=' + Math.round(r.dollar.x) };
  if (want && has && r.dollar.x > r.money[0].x) {
    return { ok: false, why: '$ at x=' + Math.round(r.dollar.x) + ' is right of the first figure at x=' + Math.round(r.money[0].x) };
  }
  // The "$" must also clear the label to its left. layout.row() anchors it at a
  // fixed inset from the column edge, so a long account name is the thing that
  // could collide with it.
  const MIN_GAP = 4;
  if (want && has && r.labelRight != null && r.dollar.x - r.labelRight < MIN_GAP) {
    return {
      ok: false,
      why: '$ at x=' + Math.round(r.dollar.x) + ' collides with the label ending at x=' +
           Math.round(r.labelRight) + ' (gap ' + Math.round(r.dollar.x - r.labelRight) + 'pt)',
    };
  }
  return {
    ok: true,
    gap: has && r.labelRight != null ? Math.round(r.dollar.x - r.labelRight) : null,
    detail: has ? ('$@' + Math.round(r.dollar.x) + ' figure@' + Math.round(r.money[0].x) +
                   ' labelEnd@' + Math.round(r.labelRight) +
                   ' gap=' + Math.round(r.dollar.x - r.labelRight) + 'pt' +
                   ' cells=' + r.money.length) : 'no $ (as expected)',
  };
}

module.exports = { readRows, pageText, figureRow, dollarOn, MONEY, DASH };
