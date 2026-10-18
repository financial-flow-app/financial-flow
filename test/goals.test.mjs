// Goal editing: add, change and delete goals on the phone; export writes Settings rows 69-88 (A-G) only.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { importWorkbook } from '../src/import.js';
import { openWorkbook, exportWorkbook, sheetXml, readCells } from '../src/xlsx.js';
import { calculate } from '../src/calc.js';
import { dateSerial, goalEdits, GOAL_TYPES } from '../src/ffos.js';
import {
  goalNeeds, goalChoices, cleanGoal, goalProblem, addGoal, updateGoal, deleteGoal,
  workbookEdits, changeSummary, olderCopyCheck,
} from '../src/entry.js';
import { SCENARIOS } from './scenarios/index.mjs';
import { buildScenario } from './scenarios/build.mjs';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));
const goalsScenario = async () => buildScenario(fixture('ffos-template.xlsx'), SCENARIOS.find((s) => s.name === 'goals'));
const blank = { name: null, type: null, target: null, basedOn: null, currentFrom: null, account: null, targetDate: null };

test('goal form: the fields each type uses', () => {
  const uses = (t) => Object.entries(goalNeeds(t)).filter(([, v]) => v).map(([k]) => k);
  assert.deepEqual(uses('Fixed Amount'), ['target', 'currentFrom']);
  assert.deepEqual(uses('Annual Contribution'), ['target', 'currentFrom']);
  assert.deepEqual(uses('Pay Down To'), ['target', 'currentFrom']);
  assert.deepEqual(uses('Save By Date'), ['target', 'currentFrom', 'targetDate']);
  assert.deepEqual(uses('percentage'), ['target', 'basedOn', 'currentFrom'], 'case is ignored, as in Excel');
  assert.deepEqual(uses('Multiple'), ['target', 'basedOn', 'currentFrom']);
  assert.deepEqual(uses('Tithe Balance'), []);
  assert.equal(GOAL_TYPES.length, 7);
});

test('goal form: unused fields are cleared, the account only with Specific Account', () => {
  const full = { name: '  Cushion ', type: 'Fixed Amount', target: 1000, basedOn: 'Net Worth', currentFrom: 'Liquid Cash', account: 'Checking A', targetDate: 46000 };
  assert.deepEqual(cleanGoal(full), { ...blank, name: 'Cushion', type: 'Fixed Amount', target: 1000, currentFrom: 'Liquid Cash' });
  assert.deepEqual(cleanGoal({ ...full, type: 'Save By Date', currentFrom: 'Specific Account' }),
    { name: 'Cushion', type: 'Save By Date', target: 1000, basedOn: null, currentFrom: 'Specific Account', account: 'Checking A', targetDate: 46000 });
  assert.equal(cleanGoal({ ...full, type: 'Percentage', basedOn: 'Specific Account' }).account, 'Checking A');
  assert.deepEqual(cleanGoal({ ...full, type: 'Tithe Balance' }), { ...blank, name: 'Cushion', type: 'Tithe Balance' });
  assert.equal(cleanGoal({ ...full, name: 2027 }).name, 2027, 'a name Excel holds as a number stays a number');
});

test('goal form: what must be filled in', () => {
  const ok = { name: 'Cushion', type: 'Multiple', target: 6, basedOn: 'Average Monthly True Cost', currentFrom: 'Liquid Cash' };
  assert.equal(goalProblem(ok), null);
  assert.match(goalProblem({ ...ok, name: '  ' }), /name/);
  assert.match(goalProblem({ ...ok, type: null }), /type/);
  assert.match(goalProblem({ ...ok, target: null }), /target/);
  assert.match(goalProblem({ ...ok, basedOn: null }), /based on/);
  assert.match(goalProblem({ ...ok, currentFrom: null }), /current value/);
  assert.match(goalProblem({ ...ok, currentFrom: 'Specific Account' }), /account/);
  assert.equal(goalProblem({ ...ok, currentFrom: 'Specific Account', account: 'Checking A' }), null);
  assert.equal(goalProblem({ name: 'Tithe', type: 'Tithe Balance' }), null, 'Tithe Balance needs only a name');
  assert.equal(goalProblem({ name: 'Trip', type: 'Save By Date', target: 500, currentFrom: 'Liquid Cash' }), null, 'the date may be blank');
});

test('goal choices come from the workbook', async () => {
  const data = await importWorkbook(await goalsScenario());
  const c = goalChoices(data);
  assert.deepEqual(c.accounts, ['Checking A', 'Emergency fund', 'Card A', 'Roth IRA', 'Taxable']);
  assert.ok(c.metrics.includes('Net Worth') && c.metrics.at(-1) === 'Specific Account');
  assert.deepEqual(c.types, GOAL_TYPES);
});

test('add, change and delete: slots, change count and the 20-goal limit', async () => {
  const base = await importWorkbook(await goalsScenario());
  const data = structuredClone(base);
  assert.equal(addGoal(data, { name: 'New one', type: 'Fixed Amount', target: 100, currentFrom: 'Net Worth' }), 9, 'first blank slot');
  updateGoal(data, 0, { ...data.goals[0], target: 60000 });
  deleteGoal(data, 4);
  assert.deepEqual(data.goals[4], blank);
  const s = changeSummary(base, data);
  assert.deepEqual([s.goals, s.total], [3, 3]);
  assert.equal(addGoal(data, { name: 'Reuse', type: 'Tithe Balance' }), 4, 'a deleted slot is reused');
  for (let i = 0; i < 10; i++) assert.notEqual(addGoal(data, { name: `Goal ${i}`, type: 'Tithe Balance' }), null);
  assert.equal(addGoal(data, { name: 'One too many', type: 'Tithe Balance' }), null, 'all 20 used');
  // The engine sees the change at once.
  const out = calculate(data, dateSerial('2026-10-15'));
  assert.equal(out.goals[0].target, 60000);
  assert.equal(out.goals[9].name, 'New one');
});

test('export: only the edited goal cells in Settings change, and they import back the same', async () => {
  const bytes = await goalsScenario();
  const base = await importWorkbook(bytes);
  const data = structuredClone(base);
  addGoal(data, { name: 'Synthetic trip', type: 'Save By Date', target: 2500, currentFrom: 'Specific Account', account: 'Emergency fund', targetDate: dateSerial('2027-03-31') });
  updateGoal(data, 1, { ...data.goals[1], target: 25 });
  deleteGoal(data, 7);
  const edits = workbookEdits(base, data);
  assert.equal(edits.get('Settings').size, 3 * 7, 'every cell of each changed goal row');
  const editedRows = [...new Set([...edits.get('Settings').keys()].map((r) => Number(r.slice(1))))].sort((a, b) => a - b);
  assert.deepEqual(editedRows, [70, 76, 78]);

  const wb0 = await openWorkbook(bytes);
  const out = await exportWorkbook(wb0, edits);
  const back = await importWorkbook(out);
  assert.deepEqual(back.goals, data.goals);
  assert.deepEqual(back.months, base.months);
  assert.deepEqual(back.giving, base.giving);
  assert.equal(changeSummary(back, data).total, 0);

  // Settings: outside the edited rows the XML is identical, and the edited cells keep their styles.
  const wb1 = await openWorkbook(out);
  const x0 = await sheetXml(wb0, 'Settings');
  const x1 = await sheetXml(wb1, 'Settings');
  const rowsOf = (x) => new Map([...x.matchAll(/<row\b[^>]*?\br="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].map((m) => [Number(m[1]), m[0]]));
  const r0 = rowsOf(x0), r1 = rowsOf(x1);
  assert.deepEqual([...r1.keys()], [...r0.keys()]);
  for (const [n, xml] of r0) if (!editedRows.includes(n)) assert.equal(r1.get(n), xml, `row ${n}`);
  const outside = (x) => x.slice(0, x.indexOf('<sheetData')) + x.slice(x.indexOf('</sheetData>'));
  assert.equal(outside(x1), outside(x0), 'validations and everything outside the cells unchanged');
  const c0 = readCells(x0, wb0.sharedStrings), c1 = readCells(x1, wb1.sharedStrings);
  assert.ok(c0.get('A76').value, 'slot 7 had a goal before');
  assert.equal(c1.get('A76').kind, 'empty');
  assert.equal(c1.get('G78').value, dateSerial('2027-03-31'));
  const style = (x, ref) => x.match(new RegExp(`<c r="${ref}"[^>]*?\\bs="(\\d+)"`))?.[1];
  for (const ref of ['A70', 'C70', 'A76', 'B76', 'G78']) assert.equal(style(x1, ref), style(x0, ref), `${ref} style`);
});

test('export into the Excel-saved file (shared strings): a goal added and read back', async () => {
  const bytes = fixture('ffos-template-excel.xlsx');
  const base = await importWorkbook(bytes);
  const data = structuredClone(base);
  assert.equal(addGoal(data, { name: 'Synthetic cushion', type: 'Multiple', target: 3, basedOn: 'Average Monthly True Cost', currentFrom: 'Liquid Cash' }), 0);
  const back = await importWorkbook(await exportWorkbook(await openWorkbook(bytes), workbookEdits(base, data)));
  assert.deepEqual(back.goals, data.goals);
});

test('no goal changes: Settings is left alone', async () => {
  const base = await importWorkbook(fixture('ffos-template.xlsx'));
  assert.equal(workbookEdits(base, structuredClone(base)).get('Settings').size, 0);
});

test('older-copy check counts goals', async () => {
  const base = await importWorkbook(await goalsScenario());
  const exported = structuredClone(base);
  deleteGoal(exported, 0);
  const w = olderCopyCheck({ exported, exportedAt: '2026-10-18T12:00:00.000Z' }, { ...base, exportStamp: null });
  assert.equal(w.warn, true);
  assert.equal(w.goals, 1);
});

test('goalEdits: checks the slot and the target', () => {
  assert.throws(() => goalEdits([{ slot: 20 }]), /outside/);
  assert.throws(() => goalEdits([{ slot: 0, name: 'x', target: 'lots' }]), /number/);
  const e = goalEdits([{ slot: 0, name: 'x', type: 'Fixed Amount', target: 5, targetDate: '2027-01-31' }]);
  assert.deepEqual(Object.fromEntries(e), { A69: 'x', B69: 'Fixed Amount', C69: 5, D69: null, E69: null, F69: null, G69: dateSerial('2027-01-31') });
});
