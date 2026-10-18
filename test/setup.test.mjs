// Accounts, investments and income sources set up on the phone: add (into an unused slot), rename, close and end.
// Export writes only the changed Settings cells; no past month changes.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { importWorkbook } from '../src/import.js';
import { openWorkbook, exportWorkbook, sheetXml, readCells } from '../src/xlsx.js';
import { calculate } from '../src/calc.js';
import { dateSerial, settingsEdits } from '../src/ffos.js';
import {
  addAccount, addInvestment, addSource, renameSetup, closeAccount, endSource, setupProblem, freeSlots,
  workbookEdits, changeSummary, monthView, setEntry,
} from '../src/entry.js';
import { SCENARIOS } from './scenarios/index.mjs';
import { buildScenario } from './scenarios/build.mjs';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));
const goalsScenario = async () => buildScenario(fixture('ffos-template.xlsx'), SCENARIOS.find((s) => s.name === 'goals'));
const TODAY = dateSerial('2026-11-10');
const ym = (data, s) => data.months.findIndex((m) => m.ym === s);

test('adding uses the first unused slot, starts at $0 and is not required', async () => {
  const data = await importWorkbook(await goalsScenario());
  assert.equal(freeSlots(data, 'accounts'), 11);
  assert.equal(addAccount(data, { name: ' Card B ', type: 'Card' }), 3);
  assert.deepEqual(
    (({ name, type, opening, estimate, active, required }) => ({ name, type, opening, estimate, active, required }))(data.accounts[3]),
    { name: 'Card B', type: 'Card', opening: 0, estimate: 'No', active: 'Yes', required: 'No' });
  assert.equal(addInvestment(data, { name: 'Brokerage', type: 'Investment' }), 2);
  assert.deepEqual(
    (({ opening, contributed, contributedThisYear, active }) => ({ opening, contributed, contributedThisYear, active }))(data.investments[2]),
    { opening: 0, contributed: 0, contributedThisYear: 0, active: 'Yes' });
  const oct = ym(data, '2026-10');
  assert.equal(addSource(data, oct, { name: 'Side job', tithe: 'Yes', usual: 400, frequency: 'Monthly', day: 15 }), 1);
  assert.equal(data.sources[1].from, dateSerial('2026-10-01'), 'counts from the month it was added in');
  assert.equal(data.sources[1].until, null);
  assert.equal(addSource(data, oct, { name: 'Gifts', tithe: 'Not Sure' }), 2);
  assert.equal(data.sources[2].usual, null, 'no usual amount: typed each month');
});

test('a new account changes no past month', async () => {
  const base = await importWorkbook(await goalsScenario());
  const data = structuredClone(base);
  addAccount(data, { name: 'Card B', type: 'Card' });
  addInvestment(data, { name: 'Brokerage', type: 'Investment' });
  const before = calculate(base, TODAY), after = calculate(data, TODAY);
  for (const k of ['Status', 'NetCash', 'HHSpend', 'NetWorth', 'Complete']) {
    assert.deepEqual(after.months.map((r) => r[k]), before.months.map((r) => r[k]), k);
  }
});

test('names must be there and differ; types must be picked', async () => {
  const data = await importWorkbook(await goalsScenario());
  assert.match(setupProblem(data, 'accounts', { name: ' ', type: 'Card' }), /name/);
  assert.match(setupProblem(data, 'accounts', { name: 'X', type: null }), /type/);
  assert.match(setupProblem(data, 'investments', { name: 'X', type: 'Card' }), /type/);
  assert.match(setupProblem(data, 'accounts', { name: 'checking a', type: 'Card' }), /already used/);
  assert.match(setupProblem(data, 'accounts', { name: 'Roth IRA', type: 'Savings' }), /already used/, 'accounts and investments share names');
  assert.equal(setupProblem(data, 'accounts', { name: 'Checking A', type: 'Checking' }, 0), null, 'renaming to its own name is fine');
  assert.equal(setupProblem(data, 'sources', { name: 'Side job' }), null);
  assert.match(setupProblem(data, 'sources', { name: 'salary' }), /already used/);
});

test('rename keeps the slot; goals that track the account follow it', async () => {
  const data = await importWorkbook(await goalsScenario());
  renameSetup(data, 'investments', 1, 'Taxable account');
  assert.equal(data.investments[1].name, 'Taxable account');
  assert.equal(data.goals[8].account, 'Taxable account', 'goal "Taxable at 2x cost" still finds it');
  const out = calculate(data, TODAY);
  assert.equal(out.goals[8].base, 8600);
});

test('close: $0 that month, then closed; end: until the last day of the month', async () => {
  const data = await importWorkbook(await goalsScenario());
  const oct = ym(data, '2026-10');
  assert.equal(closeAccount(data, oct, 'accounts', 1), true);
  assert.deepEqual(data.months[oct].balances[1], { kind: 'typed', value: 0 });
  assert.equal(data.accounts[1].active, 'No');
  closeAccount(data, oct, 'investments', 1);
  assert.deepEqual(data.months[oct].invValues[1], { kind: 'typed', value: 0 });
  assert.equal(data.investments[1].active, 'No');
  endSource(data, oct, 0);
  assert.equal(data.sources[0].until, dateSerial('2026-10-31'));
  const nov = calculate(data, TODAY).months[oct + 1];
  assert.equal(nov.Auto0, 0, 'no automatic pay after it ended');
  const v = monthView(data, calculate(data, TODAY), oct);
  assert.ok(v.accounts.find((f) => f.slot === 1).hint.includes('closed'));
  assert.equal(v.income[0].ended, true);
});

test('change count: one per account or source set up, whatever cells it took', async () => {
  const base = await importWorkbook(await goalsScenario());
  const data = structuredClone(base);
  addAccount(data, { name: 'Card B', type: 'Card' });
  addSource(data, ym(data, '2026-10'), { name: 'Side job', tithe: 'Yes' });
  const s = changeSummary(base, data);
  assert.deepEqual([s.setup, s.total], [2, 2]);
});

test('export writes only the changed Settings cells, and they import back the same', async () => {
  const bytes = await goalsScenario();
  const base = await importWorkbook(bytes);
  const data = structuredClone(base);
  const oct = ym(data, '2026-10');
  addAccount(data, { name: 'Card B', type: 'Card' });
  setEntry(data, oct, 'balances', 3, 300);
  addInvestment(data, { name: 'Brokerage', type: 'Investment' });
  addSource(data, oct, { name: 'Side job', tithe: 'Yes', usual: 400, frequency: 'Monthly', day: 15 });
  endSource(data, oct, 0);
  const edits = workbookEdits(base, data);
  const settings = [...edits.get('Settings').keys()].sort();
  assert.deepEqual(settings, ['A22', 'A38', 'A46', 'B22', 'B38', 'B46', 'C22', 'C38', 'C46', 'D22', 'D38', 'D46', 'E22', 'E38', 'E46', 'F22', 'F38', 'F46', 'G38', 'G45'].sort());
  const back = await importWorkbook(await exportWorkbook(await openWorkbook(bytes), edits));
  for (const t of ['accounts', 'investments', 'sources']) assert.deepEqual(back[t], data[t], t);
  assert.deepEqual(back.months, data.months);
  assert.equal(changeSummary(back, data).total, 0);
});
