// FFOS workbook layout (version 3.6): which cells are inputs, and how to edit them.

import { makeRef, parseRef, formulaRefs, rewriteRefs } from './cellref.js';
import { FORMULA_MAP } from './formula-map.js';

export const SHEET = { monthEnd: 'Month-End', giving: 'Giving', settings: 'Settings' };

// Month-End: one column per month, C = March 2026 ... FX = December 2040.
export const FIRST_MONTH = { year: 2026, month: 3 };
export const FIRST_COL = 3;
export const MONTH_COUNT = 178;

export const MONTH_END_ROWS = {
  balances: range(7, 20),
  income: range(22, 33),
  familyPaid: [35],
  oneTime: [36],
  investValues: range(38, 43),
  investDeposits: range(45, 50),
  note: [60],
};
export const MONTH_END_INPUT_ROWS = Object.values(MONTH_END_ROWS).flat();

// Giving: rows 7-2006, columns A-F.
export const GIVING_FIRST_ROW = 7;
export const GIVING_LAST_ROW = 2006;
export const GIVING_COLS = { date: 'A', amount: 'B', recipient: 'C', paidFrom: 'D', household: 'E', note: 'F', deductible: 'G' };

function range(a, b) {
  return Array.from({ length: b - a + 1 }, (_, i) => a + i);
}

/** "2026-04" -> Month-End column number. */
export function monthCol(ym) {
  const [y, m] = ym.split('-').map(Number);
  const idx = (y - FIRST_MONTH.year) * 12 + (m - FIRST_MONTH.month);
  if (!(idx >= 0 && idx < MONTH_COUNT)) throw new Error(`Month ${ym} is outside the workbook (2026-03 to 2040-12).`);
  return FIRST_COL + idx;
}

/** Excel date serial (1900 system) for "YYYY-MM-DD". */
export function dateSerial(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

export function serialToIso(n) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return d.toISOString().slice(0, 10);
}

// ---------- automatic formulas ----------

/**
 * Infer how a row's formula changes from one column to the next, from its formulas in two
 * neighbouring columns. Each reference gets a step {dc, dr} applied per column, whether or not
 * it carries $ (the Calc row references are absolute but still move with the month).
 */
export function inferSteps(f1, f2) {
  const a = formulaRefs(f1), b = formulaRefs(f2);
  if (a.length !== b.length) return null;
  const steps = a.map((r, i) => ({ dc: b[i].col - r.col, dr: b[i].row - r.row }));
  const check = rewriteRefs(f1, (r, i) => ({ col: r.col + steps[i].dc, row: r.row + steps[i].dr }));
  return check === f2 ? steps : null;
}

/** The template's automatic formula for a Month-End cell, or null when the template cell is blank. */
export function automaticFormula(ref) {
  const { col, row } = parseRef(ref);
  const entry = FORMULA_MAP[row];
  if (!entry) return null;
  if (col === FIRST_COL) return entry.first;
  if (entry.next == null) return null;
  const k = col - (FIRST_COL + 1);
  return rewriteRefs(entry.next, (r, i) => ({ col: r.col + entry.steps[i][0] * k, row: r.row + entry.steps[i][1] * k }));
}

// ---------- edits ----------

export const AUTO = Symbol('automatic');

/**
 * Month-End edits.
 * changes: [{month: 'YYYY-MM', row, value}] where value is a number, a string (note row only),
 * null (clear to blank) or AUTO (restore the automatic formula; blank if the cell has none).
 */
export function monthEndEdits(changes) {
  const edits = new Map();
  for (const { month, row, value } of changes) {
    if (!MONTH_END_INPUT_ROWS.includes(row)) throw new Error(`Month-End row ${row} is not an input row.`);
    const ref = makeRef(monthCol(month), row);
    if (value === AUTO || value === null) {
      const f = automaticFormula(ref);
      edits.set(ref, f ? { formula: f } : null);
    } else if (row === MONTH_END_ROWS.note[0]) {
      edits.set(ref, String(value));
    } else {
      if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`${ref}: a number is expected.`);
      edits.set(ref, value);
    }
  }
  return edits;
}

/**
 * Giving edits. rows: [{row, date, amount, recipient, paidFrom, household, note}] (row = sheet row).
 * date is "YYYY-MM-DD" or an Excel serial. Missing or null fields become blank cells.
 * tax: the workbook has the tax option (FFOS 3.8), so column G (Tax-deductible?) is written too.
 */
export function givingEdits(rows, { tax = false } = {}) {
  const edits = new Map();
  for (const g of rows) {
    if (!(g.row >= GIVING_FIRST_ROW && g.row <= GIVING_LAST_ROW)) throw new Error(`Giving row ${g.row} is outside the log.`);
    const r = g.row;
    edits.set(`A${r}`, typeof g.date === 'number' ? g.date : g.date ? dateSerial(g.date) : null);
    edits.set(`B${r}`, g.amount ?? null);
    for (const k of ['recipient', 'paidFrom', 'household', 'note', ...(tax ? ['deductible'] : [])]) {
      edits.set(`${GIVING_COLS[k]}${r}`, g[k] ? String(g[k]) : null);
    }
  }
  return edits;
}
