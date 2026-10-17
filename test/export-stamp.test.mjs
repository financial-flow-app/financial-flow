// Export file names with date and time, the stamp inside the file, and the older-copy check.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readZip, entryText } from '../src/zip.js';
import { openWorkbook, exportWorkbook, readStamp } from '../src/xlsx.js';
import { importWorkbook } from '../src/import.js';
import { dateSerial } from '../src/ffos.js';
import { setEntry, addGift, workbookEdits, exportFileName, olderCopyCheck, changeSummary } from '../src/entry.js';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));
const at = new Date(2026, 9, 17, 14, 32);

test('export file name: date and time, no colon, no piling up of old stamps', () => {
  assert.equal(exportFileName('FFOS.xlsx', at), 'FFOS 2026-10-17 1432.xlsx');
  assert.equal(exportFileName('FFOS 2026-10-01 0902.xlsx', at), 'FFOS 2026-10-17 1432.xlsx');
  assert.equal(exportFileName('Household budget.XLSX', at), 'Household budget 2026-10-17 1432.xlsx');
  assert.equal(exportFileName(null, at), 'FFOS 2026-10-17 1432.xlsx');
  assert.ok(!exportFileName('FFOS.xlsx', at).includes(':'));
});

for (const f of ['ffos-template.xlsx', 'ffos-template-excel.xlsx']) {
  test(`stamp: written, read back, updated, other parts untouched (${f})`, async () => {
    const original = fixture(f);
    const wb = await openWorkbook(original);
    assert.equal(await readStamp(wb), null);
    const one = await exportWorkbook(wb, new Map(), { stamp: '2026-10-05T14:32:00.000Z' });
    const wb1 = await openWorkbook(one);
    assert.equal(await readStamp(wb1), '2026-10-05T14:32:00.000Z');
    assert.equal((await importWorkbook(one)).exportStamp, '2026-10-05T14:32:00.000Z');
    // A second export replaces the value and keeps one property.
    const two = await exportWorkbook(wb1, new Map(), { stamp: '2026-10-17T09:00:00.000Z' });
    const custom = await entryText(readZip(two).entries.find((e) => e.name === 'docProps/custom.xml'));
    assert.equal((custom.match(/<property\b/g) || []).length, 1);
    assert.equal(await readStamp(await openWorkbook(two)), '2026-10-17T09:00:00.000Z');
    // Only the custom part, the content types and the root relationships differ.
    const before = new Map(readZip(original).entries.map((e) => [e.name, e]));
    for (const e of readZip(one).entries) {
      if (['docProps/custom.xml', '[Content_Types].xml', '_rels/.rels'].includes(e.name)) continue;
      assert.ok(Buffer.from(e.data).equals(Buffer.from(before.get(e.name).data)), `${e.name} changed`);
    }
    const ct = await entryText(readZip(one).entries.find((e) => e.name === '[Content_Types].xml'));
    assert.match(ct, /PartName="\/docProps\/custom\.xml"/);
    const rels = await entryText(readZip(one).entries.find((e) => e.name === '_rels/.rels'));
    assert.match(rels, /Target="docProps\/custom\.xml"/);
  });
}

test('stamp: other custom properties are kept', async () => {
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const one = await openWorkbook(await exportWorkbook(wb, new Map(), { stamp: 'A' }));
  // Pretend Excel added a property of its own next to ours.
  const xml = (await entryText(one.byName.get('docProps/custom.xml'))).replace('</Properties>',
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Owner note"><vt:lpwstr>keep me</vt:lpwstr></property></Properties>');
  const { writeZip } = await import('../src/zip.js');
  const withOther = await openWorkbook(await writeZip(one.zip, new Map([['docProps/custom.xml', new TextEncoder().encode(xml)]])));
  const out = await entryText(readZip(await exportWorkbook(withOther, new Map(), { stamp: 'B' })).entries.find((e) => e.name === 'docProps/custom.xml'));
  assert.match(out, /name="Owner note"><vt:lpwstr>keep me</);
  assert.match(out, /name="FinancialFlowExport"><vt:lpwstr>B</);
});

test('older-copy check: warns for an older copy, not for Excel edits after the latest export', async () => {
  const base = await importWorkbook(fixture('ffos-template.xlsx'));
  // The phone exports October entries on Oct 5.
  const sent = structuredClone(base);
  setEntry(sent, 7, 'balances', 0, 5000);
  setEntry(sent, 7, 'income', 0, 3000);
  addGift(sent, { date: dateSerial('2026-10-03'), amount: 50, recipient: 'Sample Charity' });
  const exportedAt = '2026-10-05T14:32:00.000Z';
  const memory = { exported: { ...sent, exportStamp: exportedAt }, exportedAt };

  // 1. The old file from before that export (stamped Oct 1): warn, with what's missing.
  const old = { ...structuredClone(base), exportStamp: '2026-10-01T10:02:00.000Z' };
  const w1 = olderCopyCheck(memory, old);
  assert.equal(w1.warn, true);
  assert.equal(w1.missing, 3);
  assert.deepEqual(w1.months, ['2026-10']);
  assert.equal(w1.gifts, 1);

  // 2. A file that never went through the app, missing the entries: warn.
  assert.equal(olderCopyCheck(memory, { ...structuredClone(base), exportStamp: null }).warn, true);

  // 3. The exported file, then edited in Excel (stamp kept, a value changed): no warning.
  const edited = { ...structuredClone(sent), exportStamp: exportedAt };
  setEntry(edited, 7, 'balances', 0, 5100);
  assert.equal(olderCopyCheck(memory, edited).warn, false);

  // 4. An unstamped file that has everything: no warning.
  assert.equal(olderCopyCheck(memory, { ...structuredClone(sent), exportStamp: null }).warn, false);

  // 5. A phone that never exported: no warning.
  assert.equal(olderCopyCheck({ exported: null, exportedAt: null }, old).warn, false);
});

test('change count: one per thing done', async () => {
  const base = await importWorkbook(fixture('ffos-template.xlsx'));
  const data = structuredClone(base);
  setEntry(data, 7, 'balances', 0, 5000);
  setEntry(data, 7, 'balances', 2, 400);
  addGift(data, { date: dateSerial('2026-10-03'), amount: 50, recipient: 'Sample Charity', deductible: 'Yes' });
  const s = changeSummary(base, data);
  assert.deepEqual([s.fields, s.gifts, s.total], [2, 1, 3]);
  // The export still writes every cell of that donation.
  assert.equal(workbookEdits(base, data).get('Giving').size, 7);
});
