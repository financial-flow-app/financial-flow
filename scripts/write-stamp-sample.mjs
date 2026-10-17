// Writes test-output/stamp.xlsx (a stamped export with one entry), or with "read",
// prints the stamp found in test-output/stamp-saved.xlsx (after Excel saved it).
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { openWorkbook, exportWorkbook, readStamp } from '../src/xlsx.js';
import { monthEndEdits } from '../src/ffos.js';

const out = new URL('../test-output/', import.meta.url);
mkdirSync(out, { recursive: true });
if (process.argv[2] === 'read') {
  const wb = await openWorkbook(readFileSync(new URL('stamp-saved.xlsx', out)));
  console.log(`stamp after Excel saved: ${await readStamp(wb)}`);
} else {
  const wb = await openWorkbook(readFileSync(new URL('../fixtures/ffos-template.xlsx', import.meta.url)));
  const edits = new Map([['Month-End', monthEndEdits([{ month: '2026-04', row: 7, value: 1234.56 }])]]);
  writeFileSync(new URL('stamp.xlsx', out), await exportWorkbook(wb, edits, { stamp: '2026-10-17T14:32:00.000Z' }));
  console.log('wrote test-output/stamp.xlsx');
}
