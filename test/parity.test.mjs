// Parity: the engine must reproduce what Excel computed for each synthetic scenario.
// Excel's results are in test/parity/<name>.excel.json (made by scripts/excel-parity.ps1).
// Each file records the date Excel used as "today", and the engine runs with that date.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { SCENARIOS } from './scenarios/index.mjs';
import { buildScenario } from './scenarios/build.mjs';
import { importWorkbook } from '../src/import.js';
import { calculate } from '../src/calc.js';
import { dashboardView } from '../src/dashboard.js';

const template = readFileSync(new URL('../fixtures/ffos-template.xlsx', import.meta.url));
const norm = (v) => (v === '' || v === undefined ? null : v);

/** Compare one value; returns a message or null. Numbers must agree to 1e-6 (far tighter than a cent). */
function diff(where, excel, app) {
  excel = norm(excel);
  app = norm(app);
  if (typeof excel === 'number' && typeof app === 'number') {
    return Math.abs(excel - app) <= 1e-6 ? null : `${where}: Excel ${excel}, app ${app}`;
  }
  if (typeof excel === 'number' && typeof app === 'boolean') app = app ? 1 : 0;
  return excel === app ? null : `${where}: Excel ${JSON.stringify(excel)}, app ${JSON.stringify(app)}`;
}

/** Every difference between Excel's results and the engine's, as messages. */
function compareAll(xl, out, view = null) {
  const problems = [];
  const check = (where, a, b) => {
    const p = diff(where, a, b);
    if (p) problems.push(p);
  };
  // Month columns are A:DX and ET:FE; DY:ES hold the metric list and the weekday table.
  const headers = xl.calcHeaders[0];
  const monthCol = (c) => c <= 127 || c >= 149;
  xl.calc.forEach((row, i) => {
    headers.forEach((h, c) => {
      if (!monthCol(c) || h === null || h === 'Spare') return;
      check(`${out.months[i].ym} ${h}`, row[c], out.months[i][h]);
    });
  });
  for (const [label, value] of [...xl.metrics, ...(xl.taxMetrics ?? []), ...(xl.owedMetrics ?? [])]) check(`metric "${label}"`, value, out.metrics[label]);
  const acctCols = ['name', 'value', 'lifetime', 'stale', null, null, null, null, 'rank', 'type', null, 'signed', 'estimate'];
  xl.accounts.forEach((row, i) => acctCols.forEach((k, c) => k && check(`account ${i} ${k}`, row[c], out.accounts[i][k])));
  const goalCols = ['name', 'base', 'current', 'target', 'remaining', 'status', 'progress', 'note'];
  xl.goals.forEach((row, i) => goalCols.forEach((k, c) => check(`goal ${i} ${k}`, row[c], out.goals[i][k])));
  xl.health.forEach((row, i) => check(`health ${i}`, row[0], out.health[i]));
  const qCols = ['label', 'start', 'carried', 'due', 'given', 'balance', 'status'];
  xl.history.forEach((row, i) => qCols.forEach((k, c) => check(`quarter ${i} ${k}`, row[c], out.quarters[i][k])));
  if (view) compareDashboard(xl.dashboard, view, check);
  return problems;
}

/** The Dashboard sheet (A1:L44, 0-based [row][col]) against the app's dashboard view. */
function compareDashboard(d, v, check) {
  check('Dashboard A2', d[1][0], v.headline);
  const tiles = [['netWorth', 0], ['liquid', 3], ['investments', 6], ['owed', 9]];
  for (const [k, c] of tiles) {
    check(`Dashboard tile ${k}`, d[4][c], v.tiles[k].value);
    check(`Dashboard tile ${k} line`, d[6][c], v.tiles[k].line);
  }
  v.flow.headers.forEach((t, j) => check(`Dashboard flow header ${j}`, d[9][1 + j], t));
  v.flow.rows.forEach((row, i) => row.forEach((x, j) => check(`Dashboard flow ${row[0].trim()} col ${j}`, d[10 + i][j], x)));
  check('Dashboard tithe balance', d[19][1], v.tithe.balance);
  check('Dashboard tithe line', d[19][2], v.tithe.line);
  if (d.length >= 49) {
    check('Dashboard tax received', d[46][1], v.tax?.received);
    check('Dashboard tax not withheld', d[47][1], v.tax?.unwithheld);
    check('Dashboard tax deductible giving', d[48][1], v.tax?.deductible);
  }
  v.health.forEach((t, i) => check(`Dashboard health ${i}`, d[9 + i][5], t));
  const named = v.goals;
  for (let i = 0; i < 20; i++) {
    // The sheet shows goals by slot; the app lists the named ones in slot order.
    const g = named.find((x) => x.slot === i);
    const cells = g ? [g.name, g.target, g.current, g.progress, g.remaining, g.status, g.note] : Array(7).fill(null);
    cells.forEach((x, j) => check(`Dashboard goal slot ${i} col ${j}`, d[23 + i][j], x));
  }
  for (let i = 0; i < 20; i++) {
    const a = v.accounts[i];
    const cells = a ? [a.name, a.type, a.signed, a.estimate] : Array(4).fill(null);
    cells.forEach((x, j) => check(`Dashboard account ${i} col ${j}`, d[23 + i][8 + j], x));
  }
}

const resultsFile = (sc) => new URL(`./parity/${sc.name}.excel.json`, import.meta.url);

async function run(sc) {
  const xl = JSON.parse(readFileSync(resultsFile(sc), 'utf8'));
  const data = await importWorkbook(await buildScenario(template, sc));
  const out = calculate(data, xl.today);
  return { xl, out, view: dashboardView(out, data, xl.today) };
}

for (const sc of SCENARIOS) {
  test(`parity: ${sc.name}`, { skip: !existsSync(resultsFile(sc)) && 'no Excel results yet (run scripts/excel-parity.ps1)' }, async () => {
    const { xl, out, view } = await run(sc);
    const problems = compareAll(xl, out, view);
    assert.equal(problems.length, 0, `${problems.length} differences:\n${problems.slice(0, 40).join('\n')}`);
  });
}

test('parity check catches a one-cent difference and a changed word', { skip: !existsSync(resultsFile(SCENARIOS[0])) }, async () => {
  const { xl, out, view } = await run(SCENARIOS[0]);
  out.months.find((r) => typeof r.HHSpend === 'number').HHSpend += 0.01;
  out.goals[0].name = 'changed';
  view.tiles.netWorth.line = view.tiles.netWorth.line.replace('vs', 'versus');
  const problems = compareAll(xl, out, view);
  assert.equal(problems.length, 3, problems.join('\n'));
});
