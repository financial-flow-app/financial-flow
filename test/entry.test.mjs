import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { importWorkbook } from '../src/import.js';
import { openWorkbook, exportWorkbook } from '../src/xlsx.js';
import { calculate } from '../src/calc.js';
import { dateSerial } from '../src/ffos.js';
import {
  parseAmount, setEntry, setNote, addGift, updateGift, deleteGift,
  workbookEdits, changeCount, nextMonthToFill, monthView,
} from '../src/entry.js';
import { SCENARIOS } from './scenarios/index.mjs';
import { buildScenario } from './scenarios/build.mjs';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));
const clone = (x) => structuredClone(x);

test('amounts: what people type', () => {
  assert.deepEqual(parseAmount('1,234.50'), { value: 1234.5 });
  assert.deepEqual(parseAmount(' $1,234 '), { value: 1234 });
  assert.deepEqual(parseAmount('-50'), { value: -50 });
  assert.deepEqual(parseAmount('(50)'), { value: -50 });
  assert.deepEqual(parseAmount('.5'), { value: 0.5 });
  assert.deepEqual(parseAmount('0'), { value: 0 });
  assert.deepEqual(parseAmount(''), { value: null });
  assert.ok(parseAmount('12a').error);
  assert.ok(parseAmount('1.2.3').error);
  assert.ok(parseAmount('2000000000').error);
});

for (const f of ['ffos-template.xlsx', 'ffos-template-excel.xlsx']) {
  test(`entries survive export and re-import (${f})`, async () => {
    const bytes = fixture(f);
    const base = await importWorkbook(bytes);
    const data = clone(base);
    assert.equal(changeCount(base, data), 0);

    setEntry(data, 1, 'balances', 0, 2500.75);
    setEntry(data, 1, 'income', 3, 0);
    setEntry(data, 1, 'family', null, 120);
    setEntry(data, 1, 'deposits', 2, -75);
    setNote(data, 1, '  synthetic note  ');
    const r1 = addGift(data, { date: dateSerial('2026-04-03'), amount: 18, recipient: 'Sample Charity', household: 'No' });
    const r2 = addGift(data, { date: dateSerial('2026-04-09'), amount: 36, recipient: 'Sample Fund', deductible: 'Yes' });
    assert.deepEqual([r1, r2], [7, 8]);
    assert.equal(changeCount(base, data), 5 + 2, '5 month fields and 2 donations: one per thing done, not per cell');

    const out = await exportWorkbook(await openWorkbook(bytes), workbookEdits(base, data));
    const back = await importWorkbook(out);
    assert.deepEqual(back.months, data.months);
    assert.deepEqual(back.giving, data.giving);
    assert.equal(changeCount(back, data), 0);

    // Clear back to automatic, change and delete gifts, export again.
    const data2 = clone(back);
    setEntry(data2, 1, 'balances', 0, null);
    updateGift(data2, 8, { ...data2.giving[1], amount: 40 });
    deleteGift(data2, 7);
    assert.equal(addGift(data2, { date: dateSerial('2026-05-01'), amount: 5 }), 7, 'a freed row is reused');
    const out2 = await exportWorkbook(await openWorkbook(out), workbookEdits(back, data2));
    const back2 = await importWorkbook(out2);
    assert.equal(back2.months[1].balances[0].kind, 'auto');
    assert.deepEqual(back2.giving, data2.giving);
  });
}

test('a formula the user set in Excel is not overwritten', async () => {
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const data = await importWorkbook(await exportWorkbook(wb, new Map([['Month-End', new Map([['D22', { formula: '100+50' }]])]])));
  assert.equal(setEntry(data, 1, 'income', 0, 10), false);
  assert.equal(data.months[1].income[0].kind, 'formula');
});

test('month screen: next month to fill, required first, grey automatic values', async () => {
  const sc = SCENARIOS.find((s) => s.name === 'opening-month');
  const data = await importWorkbook(await buildScenario(fixture('ffos-template.xlsx'), sc));
  const today = dateSerial('2026-10-16');
  const out = calculate(data, today);
  assert.equal(out.months[nextMonthToFill(out, today)].ym, '2026-07'); // June is the latest done month

  const v = monthView(data, out, 4); // July
  assert.deepEqual(v.accounts.map((a) => a.label), ['Checking A', 'Card A', 'Savings A', 'Wallet']);
  assert.equal(v.missing, 2);
  assert.equal(v.accounts[0].auto, null, 'required accounts show no carried value');
  assert.equal(v.accounts[2].auto, 10100, 'optional accounts show last month');
  assert.equal(v.income.find((x) => x.label === 'Salary').auto, out.months[4].Auto0);
  assert.equal(v.family.auto, 300);
  assert.equal(v.status, '⚠ Needs balances');

  const pre = monthView(data, out, 0); // March, before the opening date
  assert.equal(pre.phase, 'pre');
  assert.deepEqual(pre.accounts, []);
});
