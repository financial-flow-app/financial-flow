// Test fixtures, all synthetic:
//   fixtures/ffos-template.xlsx        the blank template as the workbook generator writes it
//   fixtures/ffos-template-excel.xlsx  the same file opened and saved by Excel, then scrubbed:
//     powershell -File scripts/excel-resave.ps1 fixtures/ffos-template.xlsx <tmp.xlsx>
//     node scripts/make-fixtures.mjs scrub <tmp.xlsx> fixtures/ffos-template-excel.xlsx
// Scrubbing removes what Excel stamps into a saved file about the machine and user (author
// names, local path, company), so the fixture holds nothing personal.

import { readFileSync, writeFileSync } from 'node:fs';
import { openWorkbook, partText } from '../src/xlsx.js';
import { writeZip } from '../src/zip.js';

const enc = new TextEncoder();
const [cmd, a, b] = process.argv.slice(2);

if (cmd === 'scrub' && a && b) {
  const wb = await openWorkbook(readFileSync(a));
  const changes = new Map();
  const blank = (xml, tags) => tags.reduce((x, t) => x.replace(new RegExp(`<${t}>[^<]*</${t}>`, 'g'), `<${t}>Synthetic</${t}>`), xml);
  if (wb.byName.has('docProps/core.xml')) {
    changes.set('docProps/core.xml', enc.encode(blank(await partText(wb, 'docProps/core.xml'), ['dc:creator', 'cp:lastModifiedBy'])));
  }
  if (wb.byName.has('docProps/app.xml')) {
    changes.set('docProps/app.xml', enc.encode(blank(await partText(wb, 'docProps/app.xml'), ['Company', 'Manager'])));
  }
  let wx = await partText(wb, 'xl/workbook.xml');
  wx = wx.replace(/<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?x15ac:absPath[\s\S]*?<\/mc:AlternateContent>/g, '');
  changes.set('xl/workbook.xml', enc.encode(wx));
  writeFileSync(b, await writeZip(wb.zip, changes));
  console.log(`wrote ${b}`);
} else {
  console.error('usage: make-fixtures.mjs scrub <in.xlsx> <out.xlsx>');
  process.exit(1);
}
