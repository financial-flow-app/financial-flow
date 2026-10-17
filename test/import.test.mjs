import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { importWorkbook, WorkbookError } from '../src/import.js';
import { openWorkbook, exportWorkbook } from '../src/xlsx.js';
import { round, WEEKDAY, EDATE, dollars } from '../src/excel.js';
import { dateSerial } from '../src/ffos.js';
import { runScenario } from './scenario.mjs';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));

test('import: the template and its Excel-saved copy give the same data', async () => {
  const a = await importWorkbook(fixture('ffos-template.xlsx'));
  const b = await importWorkbook(fixture('ffos-template-excel.xlsx'));
  assert.deepEqual(b, a);
  assert.equal(a.months.length, 178);
  assert.equal(a.months[0].ym, '2026-03');
  assert.equal(a.months.at(-1).ym, '2040-12');
  assert.ok(a.months.every((m) => m.balances.every((c) => c.kind === 'auto')));
  assert.equal(a.giving.length, 0);
  assert.equal(a.settings.openingDate, dateSerial('2026-03-31'));
});

test('import: typed entries come back as typed, cleared ones as automatic', async () => {
  for (const f of ['ffos-template.xlsx', 'ffos-template-excel.xlsx']) {
    const { pass1, pass2 } = await runScenario(fixture(f));
    const one = await importWorkbook(pass1);
    const apr = one.months[1];
    assert.deepEqual(apr.balances[0], { kind: 'typed', value: 1234.56 });
    assert.deepEqual(apr.deposits[1], { kind: 'typed', value: -50 });
    assert.equal(apr.note, 'Synthetic note: rent & <utilities> "test"');
    assert.deepEqual(one.giving.map((g) => g.recipient), ['Sample Charity A', 'Sample Fund B']);
    const two = await importWorkbook(pass2);
    assert.equal(two.months[0].balances[0].kind, 'auto');
    assert.equal(two.months[1].balances[2].kind, 'auto');
    assert.equal(two.giving.length, 1);
  }
});

test('import: a formula the user typed into an input cell is kept apart from the automatic one', async () => {
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const out = await exportWorkbook(wb, new Map([['Month-End', new Map([['D22', { formula: '100+50' }]])]]));
  const data = await importWorkbook(out);
  assert.equal(data.months[1].income[0].kind, 'formula');
});

test('import: reads FFOS 4.0 with its tax option and owed trend', async () => {
  for (const f of ['ffos-template.xlsx', 'ffos-template-excel.xlsx']) {
    const d = await importWorkbook(fixture(f));
    assert.deepEqual([d.format, d.hasTax, d.hasOwedTrend], ['FFOS 4.0', true, true], f);
    assert.ok('taxable' in d.sources[0] && 'withheld' in d.sources[0] && 'deductible' in d.recurring[0]);
  }
});

test('import: an older or unmarked workbook is refused with a clear message', async () => {
  for (const version of [null, 'FFOS 3.9', 'something else']) {
    const wb = await openWorkbook(fixture('ffos-template.xlsx'));
    const older = await exportWorkbook(wb, new Map([['Settings', new Map([['J1', version]])]]));
    await assert.rejects(importWorkbook(older), /older FFOS version.*FFOS 4\.0 or newer/, String(version));
  }
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const newer = await exportWorkbook(wb, new Map([['Settings', new Map([['J1', 'FFOS 4.2']])]]));
  assert.equal((await importWorkbook(newer)).format, 'FFOS 4.2');
});

test('import: rejects files that are not FFOS workbooks', async () => {
  await assert.rejects(importWorkbook(new TextEncoder().encode('not a zip')), WorkbookError);
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const moved = await exportWorkbook(wb, new Map([['Month-End', new Map([['C4', 46000]])]]));
  await assert.rejects(importWorkbook(moved), /different Month-End layout/);
});

test('Excel rules: ROUND half away from zero, WEEKDAY, EDATE, dollar text', () => {
  assert.equal(round(1.005, 2), 1.01);
  assert.equal(round(-2.5, 0), -3);
  assert.equal(round(2.675, 2), 2.68);
  assert.equal(round(-0.001, 2), 0);
  assert.equal(WEEKDAY(dateSerial('2026-03-01')), 1); // a Sunday
  assert.equal(EDATE(dateSerial('2026-03-31'), -1), dateSerial('2026-02-28'));
  assert.equal(dollars(1566.67), '$1,567');
  assert.equal(dollars(-1234.5), '-$1,235');
});
