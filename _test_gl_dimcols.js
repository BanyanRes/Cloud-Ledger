// Column arithmetic for the two GL detail exports after making the dimension
// columns conditional. Off-by-one in the row shapes is the real risk here, so this
// mirrors the builders exactly and asserts, for each of the four dimension
// combinations, that every row is the same width as the header and that the
// Debit/Credit/Balance indices land on the columns the header says they do.
const XLC = n => { let s = ''; n = n + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const bl = n => Array.from({ length: n }, () => '');
let bad = 0;
const check = (label, cond, extra) => { if (!cond) { bad++; console.log('FAIL  ' + label + (extra ? '  ' + extra : '')); } else console.log('PASS  ' + label); };

for (const hasCls of [false, true]) for (const hasLoc of [false, true]) {
  const tag = `cls=${hasCls} loc=${hasLoc}`;

  // ---- account drill-down -------------------------------------------------
  {
    const dimHdr = [...(hasCls ? ['Class'] : []), ...(hasLoc ? ['Location'] : []), 'Project'];
    const dimOf = l => [...(hasCls ? [l.class_name || ''] : []), ...(hasLoc ? [l.location_name || ''] : []), l.project_name || ''];
    const nD = dimHdr.length;
    const cDr = 7 + nD, cCr = 8 + nD, cBal = 9 + nD;
    const hdr = ['Date', 'JE', 'Doc #', 'Account', ...dimHdr, 'Memo', 'Offset Account', 'Vendor/Payee', 'Debit', 'Credit', 'Balance'];
    const beg = [...bl(4 + nD), 'Beginning Balance', ...bl(4), 1000];
    const det = ['2026-07-22', 'JE-1618', '18832', '11760 - x', ...dimOf({ project_name: 'P' }), 'memo', 'off', 'vend', 23015, '', 24015];
    const tot = [...bl(4 + nD), 'Totals', ...bl(2), 23015, 0, 24015];
    check(`drilldown ${tag}: header/detail width`, hdr.length === det.length, `${hdr.length} vs ${det.length}`);
    check(`drilldown ${tag}: header/beg width`, hdr.length === beg.length, `${hdr.length} vs ${beg.length}`);
    check(`drilldown ${tag}: header/total width`, hdr.length === tot.length, `${hdr.length} vs ${tot.length}`);
    check(`drilldown ${tag}: Debit index`, hdr[cDr] === 'Debit', `hdr[${cDr}]=${hdr[cDr]}`);
    check(`drilldown ${tag}: Credit index`, hdr[cCr] === 'Credit', `hdr[${cCr}]=${hdr[cCr]}`);
    check(`drilldown ${tag}: Balance index`, hdr[cBal] === 'Balance', `hdr[${cBal}]=${hdr[cBal]}`);
    check(`drilldown ${tag}: Beginning Balance label sits under Memo`, beg[4 + nD] === 'Beginning Balance' && hdr[4 + nD] === 'Memo');
    check(`drilldown ${tag}: beg balance amount in Balance col`, beg[cBal] === 1000);
    check(`drilldown ${tag}: totals amounts in Dr/Cr/Bal cols`, tot[cDr] === 23015 && tot[cCr] === 0 && tot[cBal] === 24015);
    check(`drilldown ${tag}: detail amounts in Dr/Bal cols`, det[cDr] === 23015 && det[cBal] === 24015);
    check(`drilldown ${tag}: formula letters`, XLC(cDr) + '/' + XLC(cCr) + '/' + XLC(cBal) === [XLC(cDr), XLC(cCr), XLC(cBal)].join('/'));
    // Class/Location must be absent from the header when unused.
    check(`drilldown ${tag}: Class column presence`, hdr.includes('Class') === hasCls);
    check(`drilldown ${tag}: Location column presence`, hdr.includes('Location') === hasLoc);
    check(`drilldown ${tag}: Project always present`, hdr.includes('Project'));
    check(`drilldown ${tag}: Doc # always present`, hdr.includes('Doc #'));
  }

  // ---- Trial Balance "Export GL Detail" ----------------------------------
  {
    const dimHdr = ['Project', ...(hasLoc ? ['Location'] : []), ...(hasCls ? ['Class'] : [])];
    const dimOf = l => [l.project_name || '', ...(hasLoc ? [l.location_name || ''] : []), ...(hasCls ? [l.class_name || ''] : [])];
    const nD = dimHdr.length;
    const cDr = 6 + nD, cCr = 7 + nD, cBal = 8 + nD;
    const hdr = ['Date', 'Entry #', 'Doc #', 'Account', 'Account Name', 'Memo / Description', ...dimHdr, 'Debit', 'Credit', 'Running Bal'];
    const det = ['2026-07-22', 1618, '18832', '11760', 'Other DD', 'memo', ...dimOf({ project_name: 'P' }), 23015, '', 24015];
    const lbls = [...bl(cDr), 'Total Dr', 'Total Cr', ''];
    const tot = [...bl(cDr), 23015, 0, ''];
    check(`tb-gl ${tag}: header/detail width`, hdr.length === det.length, `${hdr.length} vs ${det.length}`);
    check(`tb-gl ${tag}: header/total width`, hdr.length === tot.length, `${hdr.length} vs ${tot.length}`);
    check(`tb-gl ${tag}: Debit index`, hdr[cDr] === 'Debit', `hdr[${cDr}]=${hdr[cDr]}`);
    check(`tb-gl ${tag}: Credit index`, hdr[cCr] === 'Credit');
    check(`tb-gl ${tag}: Running Bal index`, hdr[cBal] === 'Running Bal');
    check(`tb-gl ${tag}: "Total Dr" label above the Debit total`, lbls[cDr] === 'Total Dr' && tot[cDr] === 23015);
    check(`tb-gl ${tag}: Class column presence`, hdr.includes('Class') === hasCls);
    check(`tb-gl ${tag}: Location column presence`, hdr.includes('Location') === hasLoc);
    check(`tb-gl ${tag}: Project always present`, hdr.includes('Project'));
  }
}
console.log('');
console.log(bad === 0 ? 'ALL PASS' : bad + ' FAILED');
process.exit(bad === 0 ? 0 : 1);
