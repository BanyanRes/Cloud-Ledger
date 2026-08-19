#!/usr/bin/env node
// CloudLedger — verify a house-style .docx before delivering it.
//
//   node tools/docx/verify.js path/to/file.docx
//
// Renders to PDF, rasterises every page to JPEG, and prints the image paths.
// A document that BUILDS is not a document that LOOKS RIGHT: the last step is a
// human (or Claude) actually reading those images. This script says so and exits
// non-zero if it could not produce them.
const fs = require('fs');
const path = require('path');
const cp = require('child_process');

const file = process.argv[2];
if (!file) { console.error('usage: node tools/docx/verify.js <file.docx>'); process.exit(2); }
if (!/\.docx$/i.test(file)) {
  console.error('verify.js: "' + path.basename(file) + '" is not a .docx.');
  console.error('Prose deliverables for Jimmy are .docx — including email and reply drafts.');
  console.error('See tools/docx/README.md.');
  process.exit(2);
}
if (!fs.existsSync(file)) { console.error('verify.js: no such file: ' + file); process.exit(2); }

const dir = path.dirname(path.resolve(file));
const base = path.basename(file, path.extname(file));
const outDir = path.join(dir, '_verify_' + base.replace(/[^A-Za-z0-9_-]/g, '_'));
fs.mkdirSync(outDir, { recursive: true });

function which(cmd) {
  try {
    const probe = process.platform === 'win32' ? 'where' : 'which';
    return cp.execSync(probe + ' ' + cmd, { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim().split(/\r?\n/)[0];
  } catch (e) { return null; }
}

const soffice = which('soffice') || which('libreoffice')
  || ['C:/Program Files/LibreOffice/program/soffice.exe', '/usr/bin/soffice']
    .find(p => fs.existsSync(p));
if (!soffice) {
  console.error('verify.js: LibreOffice (soffice) not found — cannot render.');
  console.error('Install it, or run this step in the Claude cloud workspace where it is present.');
  process.exit(3);
}

console.log('rendering ' + file + ' ...');
cp.execFileSync(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', outDir, path.resolve(file)],
  { stdio: 'ignore', timeout: 120000 });
const pdf = path.join(outDir, base + '.pdf');
if (!fs.existsSync(pdf)) { console.error('verify.js: conversion produced no PDF'); process.exit(3); }

const pdftoppm = which('pdftoppm');
if (!pdftoppm) {
  console.error('verify.js: pdftoppm not found — PDF is at ' + pdf + ' but pages were not rasterised.');
  process.exit(3);
}
cp.execFileSync(pdftoppm, ['-jpeg', '-r', '100', pdf, path.join(outDir, 'page')], { stdio: 'ignore', timeout: 120000 });

const pages = fs.readdirSync(outDir).filter(f => /^page-?\d+\.jpg$/i.test(f)).sort();
console.log('\n' + pages.length + ' page(s) rendered:');
for (const p of pages) console.log('  ' + path.join(outDir, p));

console.log('\nNOT VERIFIED YET. Open/Read every image above and check:');
console.log('  - page count (a stray empty paragraph adds one)');
console.log('  - the footer page numbers render at a SINGLE size');
console.log('  - custom-styled paragraphs match body copy (no serif fallback)');
console.log('  - no orphaned sign-off, no table running off the right margin');
