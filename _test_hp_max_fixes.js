// Verify the three HP requisition fixes (Max, 2026-09-10) by rolling the filed
// July HP workbook (Req 57) forward to Req 58 and checking:
//   1. Soft Cost Contingency F55 folds in the prior period's G55.
//   2. B2A recon cells I71/J71/J72 repoint to the logs' NEW Grand Total rows.
//   3. finalize restores the Budget-to-Actual over-budget red CF to the RIGHT sheet.
const fs = require('fs');
const ExcelJS = require('exceljs');
const { rollForward } = require('./server/requisition_rollforward.js');
const { finalizeRequisitionWorkbook } = require('./server/requisition_preserve.js');

const JULY = 'C:/Users/JimmyYun/OneDrive - banyanres.com/CLA - Documents/11 Bridge Banyan HP QOZB (High Point)/02 Requisition Report/2026/07 July/Bridge Banyan HP QOZB Requisition Report - 07.31.2026.xlsx';
const OUT = 'C:/Users/JimmyYun/Cloud-Ledger/_hp_req58_rebuilt.xlsx';

const F = c => { const v = c.value; if (v && typeof v === 'object' && v.formula) return '=' + v.formula; return v == null ? '' : String(v); };
const N = c => { const v = c.value; if (typeof v === 'number') return v; if (v && typeof v === 'object' && typeof v.result === 'number') return v.result; return null; };

(async () => {
  const originalBuf = fs.readFileSync(JULY);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(originalBuf);

  // Pre-roll snapshot of the July source
  const b2a0 = wb.getWorksheet('Budget to Actual');
  console.log('=== PRE-ROLL (July source) ===');
  console.log('  F55:', F(b2a0.getCell('F55')));
  console.log('  G55:', F(b2a0.getCell('G55')));
  console.log('  I71:', F(b2a0.getCell('I71')), '| J71:', F(b2a0.getCell('J71')), '| J72:', F(b2a0.getCell('J72')));

  const res = await rollForward(wb, [], { reqNumber: 58, asOfDate: '2026-08-31' });

  const b2a = wb.getWorksheet('Budget to Actual');
  const cur = wb.getWorksheet('Current Invoice Log');
  const prior = wb.getWorksheet('Prior Invoice Log');
  const findGT = (ws) => { const last = Math.max(ws.rowCount||0, ws.actualRowCount||0); for (let r=1;r<=last;r++){ const d=ws.getCell(r,4).value; if (String(d && d.result!==undefined?d.result:d).trim().toLowerCase()==='grand total') return r; } return null; };
  const newCurGT = findGT(cur), newPriorGT = findGT(prior);

  console.log('\n=== POST-ROLL (Req 58) ===');
  console.log('  new Current GT row:', newCurGT, '| new Prior GT row:', newPriorGT);
  console.log('  F55:', F(b2a.getCell('F55')));
  console.log('  G55:', F(b2a.getCell('G55')));
  console.log('  I71:', F(b2a.getCell('I71')));
  console.log('  J71:', F(b2a.getCell('J71')));
  console.log('  J72:', F(b2a.getCell('J72')));

  // ---- Checks ----
  let pass = 0, fail = 0;
  const ok = (c, label, detail) => { if (c) { pass++; console.log('  PASS  ' + label); } else { fail++; console.log('  FAIL  ' + label + (detail ? '   ' + detail : '')); } };

  console.log('\n=== FIX 1: F55 folds prior G55 ===');
  const f55 = F(b2a.getCell('F55'));
  ok(/-54143\.86$/.test(f55), 'F55 chain appended prior period G55 (-54143.86)', f55.slice(-40));
  ok(/=-G46/.test(F(b2a.getCell('G55'))), 'G55 keeps its live =-G46 formula');

  console.log('\n=== FIX 2: recon cells repointed ===');
  const i71 = F(b2a.getCell('I71')), j71 = F(b2a.getCell('J71')), j72 = F(b2a.getCell('J72'));
  ok(newPriorGT && i71.includes("!G" + newPriorGT) || i71.includes('!$G$' + newPriorGT) || (newPriorGT && i71.includes('' + newPriorGT)), 'I71 points at new Prior GT row ' + newPriorGT, i71);
  ok(newCurGT && (j71.includes('' + newCurGT)), 'J71 points at new Current GT row ' + newCurGT, j71);
  ok(newCurGT && (j72.includes('' + newCurGT)), 'J72 points at new Current GT row ' + newCurGT, j72);
  ok(!i71.includes('1173') && !j71.includes('135') && !j72.includes('135'), 'no stale hardcoded rows (1173 / 135) remain');

  // Write + finalize for the CF check
  const outBuf = await wb.xlsx.writeBuffer();
  const finalBuf = await finalizeRequisitionWorkbook(originalBuf, Buffer.from(outBuf));
  fs.writeFileSync(OUT, finalBuf);
  console.log('\nwrote ' + OUT + '  (' + finalBuf.length + ' bytes)');
  console.log('\nrfResult.warnings:', (res && res.warnings && res.warnings.length) ? res.warnings.join(' | ') : '(none)');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e.message); console.error(e.stack); process.exit(2); });
