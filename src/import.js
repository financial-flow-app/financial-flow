// Read an FFOS workbook into plain app data.
// Values are kept as Excel holds them (number | string | boolean | null for blank), so the
// calculation engine can apply Excel's own rules (blank vs "", N(), case-insensitive text).

import { openWorkbook, sheetXml, readCells, readStamp } from './xlsx.js';
import { makeRef, numToCol } from './cellref.js';
import { is } from './excel.js';
import {
  SHEET, FIRST_COL, MONTH_COUNT, FIRST_MONTH, MONTH_END_ROWS, GIVING_FIRST_ROW, GIVING_LAST_ROW,
  automaticFormula, dateSerial,
} from './ffos.js';

export const LAYOUT = {
  accounts: { first: 19, count: 14, cols: ['name', 'type', 'opening', 'estimate', 'active', 'required', 'notes'] },
  investments: { first: 36, count: 6, cols: ['name', 'type', 'opening', 'contributed', 'contributedThisYear', 'estimate', 'active', 'notes'] },
  // Columns I-J (sources) and H (recurring) exist from FFOS 3.8 on (the tax option); blank before.
  sources: { first: 45, count: 12, cols: ['name', 'tithe', 'usual', 'frequency', 'day', 'from', 'until', 'notes', 'taxable', 'withheld'] },
  recurring: { first: 60, count: 6, cols: ['name', 'amount', 'day', 'from', 'until', 'household', 'notes', 'deductible'] },
  goals: { first: 69, count: 20, cols: ['name', 'type', 'target', 'basedOn', 'currentFrom', 'account', 'targetDate'] },
};

const SETTINGS_CELLS = {
  openingDate: 'B5', titheStart: 'B6', titheRate: 'B7', openingTithe: 'B8', carryExtra: 'B9',
  avgMonths: 'B10', fallbackHousehold: 'B11', fallbackFamily: 'B12', openingTypedIncome: 'B13',
  familyUsual: 'B14', familyStart: 'B15',
};

function raw(cell) {
  if (!cell) return null;
  if (cell.kind === 'empty' || cell.kind === 'error') return null;
  if (cell.kind === 'text' && cell.value === '') return null;
  return cell.value;
}

/**
 * A Month-End input cell: {kind, value}.
 *   'typed'   a constant the user entered (value = the constant)
 *   'auto'    blank, or the workbook's own automatic formula (value = null)
 *   'formula' some other formula the user put there (value = its last calculated result)
 */
function inputCell(cell, ref) {
  if (!cell) return { kind: 'auto', value: null };
  if (cell.formula) {
    const auto = automaticFormula(ref);
    if (auto !== null && cell.formula.text === auto) return { kind: 'auto', value: null };
    return { kind: 'formula', value: raw(cell) };
  }
  const v = raw(cell);
  return v === null ? { kind: 'auto', value: null } : { kind: 'typed', value: v };
}

export class WorkbookError extends Error {}

const MIN_MAJOR = 4; // FFOS 4.0: neutral wording, tax option, change in amount owed

/** bytes (ArrayBuffer | Uint8Array) -> app data. Throws WorkbookError when it isn't an FFOS 3.6 file. */
export async function importWorkbook(bytes) {
  let wb;
  try {
    wb = await openWorkbook(bytes);
  } catch {
    throw new WorkbookError('This file is not an Excel workbook (.xlsx).');
  }
  for (const name of [SHEET.monthEnd, SHEET.giving, SHEET.settings, 'Calc']) {
    if (!wb.sheets.has(name)) throw new WorkbookError(`This isn't an FFOS workbook: there is no "${name}" sheet.`);
  }
  const ss = wb.sharedStrings;
  const settingsCells = readCells(await sheetXml(wb, SHEET.settings), ss);
  const meCells = readCells(await sheetXml(wb, SHEET.monthEnd), ss);
  const gvCells = readCells(await sheetXml(wb, SHEET.giving), ss);

  const s = (ref) => raw(settingsCells.get(ref));

  // Version check: Settings!J1 holds "FFOS <major>.<minor>". This app reads FFOS 4.0 and later.
  const version = /^FFOS (\d+)\.(\d+)/.exec(String(s('J1') ?? ''));
  if (!version || Number(version[1]) < MIN_MAJOR) {
    throw new WorkbookError(`This workbook is an older FFOS version. Update it with the workbook generator (FFOS ${MIN_MAJOR}.0 or newer), then import it again.`);
  }

  // Layout check: Month-End row 4 holds the months, C = March 2026, FX = December 2040.
  const first = raw(meCells.get(makeRef(FIRST_COL, 4)));
  const last = raw(meCells.get(makeRef(FIRST_COL + MONTH_COUNT - 1, 4)));
  if (first !== dateSerial('2026-03-01') || last !== dateSerial('2040-12-01')) {
    throw new WorkbookError('This workbook has a different Month-End layout than expected (row 4 should run from March 2026 to December 2040).');
  }
  const settings = Object.fromEntries(Object.entries(SETTINGS_CELLS).map(([k, ref]) => [k, s(ref)]));
  const table = ({ first: row0, count, cols }) =>
    Array.from({ length: count }, (_, i) => Object.fromEntries(cols.map((k, j) => [k, s(`${numToCol(j + 1)}${row0 + i}`)])));

  const months = [];
  for (let i = 0; i < MONTH_COUNT; i++) {
    const col = FIRST_COL + i;
    const cell = (row) => {
      const ref = makeRef(col, row);
      return inputCell(meCells.get(ref), ref);
    };
    const y = FIRST_MONTH.year + Math.floor((FIRST_MONTH.month - 1 + i) / 12);
    const m = ((FIRST_MONTH.month - 1 + i) % 12) + 1;
    const note = raw(meCells.get(makeRef(col, MONTH_END_ROWS.note[0])));
    months.push({
      ym: `${y}-${String(m).padStart(2, '0')}`,
      balances: MONTH_END_ROWS.balances.map(cell),
      income: MONTH_END_ROWS.income.map(cell),
      family: cell(MONTH_END_ROWS.familyPaid[0]),
      oneTime: cell(MONTH_END_ROWS.oneTime[0]),
      invValues: MONTH_END_ROWS.investValues.map(cell),
      deposits: MONTH_END_ROWS.investDeposits.map(cell),
      note: note === null ? null : String(note),
    });
  }

  // The tax option (FFOS 3.8): Settings I44 and Giving G6 carry its headers.
  const hasTax = is(s('I44'), 'Taxable?') && is(raw(gvCells.get('G6')), 'Tax-deductible?');
  // FFOS 3.9: the card/loan line shows the change owed, and a 9th health check (Dashboard F18).
  const dash = wb.sheets.has('Dashboard') ? readCells(await sheetXml(wb, 'Dashboard'), ss, (r) => r === 'F18') : new Map();
  const hasOwedTrend = /^Calc!\$EA\$103$/.test(dash.get('F18')?.formula?.text ?? '');

  const giving = [];
  for (let r = GIVING_FIRST_ROW; r <= GIVING_LAST_ROW; r++) {
    const g = (c) => raw(gvCells.get(`${c}${r}`));
    const row = { row: r, date: g('A'), amount: g('B'), recipient: g('C'), paidFrom: g('D'), household: g('E'), note: g('F') };
    if (hasTax) row.deductible = g('G');
    if (Object.entries(row).some(([k, v]) => k !== 'row' && v !== null)) giving.push(row);
  }

  return {
    format: version[0],
    hasOwedTrend,
    exportStamp: await readStamp(wb), // when this app last exported this file (ISO), or null
    hasTax,
    settings,
    accounts: table(LAYOUT.accounts),
    investments: table(LAYOUT.investments),
    sources: table(LAYOUT.sources),
    recurring: table(LAYOUT.recurring),
    goals: table(LAYOUT.goals),
    months,
    giving,
  };
}
