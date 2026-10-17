// Helpers for writing synthetic scenarios as Settings cells and Month-End entries.
import { dateSerial } from '../../src/ffos.js';
import { MONTH_END_ROWS } from '../../src/ffos.js';

export const d = (iso) => dateSerial(iso);

const put = (row, cols, values) =>
  Object.fromEntries(cols.map((c, i) => [`${c}${row}`, values[i] ?? null]).filter(([, v]) => v !== undefined));

/** Everyday account in slot 0-13. */
export const account = (slot, name, type, opening, { estimate = 'No', active = 'Yes', required = 'No' } = {}) =>
  put(19 + slot, ['A', 'B', 'C', 'D', 'E', 'F'], [name, type, opening, estimate, active, required]);

/** Investment account in slot 0-5. */
export const investment = (slot, name, type, opening, { contributed = 0, thisYear = 0, estimate = 'No', active = 'Yes' } = {}) =>
  put(36 + slot, ['A', 'B', 'C', 'D', 'E', 'F', 'G'], [name, type, opening, contributed, thisYear, estimate, active]);

/** Income source in slot 0-11. */
export const source = (slot, name, tithe, { usual = null, frequency = null, day = null, from = null, until = null, taxable = null, withheld = null } = {}) =>
  put(45 + slot, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'I', 'J'], [name, tithe, usual, frequency, day, from && d(from), until && d(until), taxable, withheld]);

/** Recurring donation in slot 0-5. */
export const recurringGift = (slot, name, amount, { day = null, from = null, until = null, household = null, deductible = null } = {}) =>
  put(60 + slot, ['A', 'B', 'C', 'D', 'E', 'F', 'H'], [name, amount, day, from && d(from), until && d(until), household, deductible]);

/** Goal in slot 0-19. */
export const goal = (slot, name, type, target, { basedOn = null, currentFrom = null, account: acct = null, targetDate = null } = {}) =>
  put(69 + slot, ['A', 'B', 'C', 'D', 'E', 'F', 'G'], [name, type, target, basedOn, currentFrom, acct, targetDate && d(targetDate)]);

/**
 * One month's entries. bal/inc/inv/dep map slot -> value.
 * Returns entries for monthEndEdits().
 */
export function month(ym, { bal = {}, inc = {}, inv = {}, dep = {}, family, oneTime, note } = {}) {
  const out = [];
  const add = (rows, map) => Object.entries(map).forEach(([slot, value]) => out.push({ month: ym, row: rows[Number(slot)], value }));
  add(MONTH_END_ROWS.balances, bal);
  add(MONTH_END_ROWS.income, inc);
  add(MONTH_END_ROWS.investValues, inv);
  add(MONTH_END_ROWS.investDeposits, dep);
  if (family !== undefined) out.push({ month: ym, row: MONTH_END_ROWS.familyPaid[0], value: family });
  if (oneTime !== undefined) out.push({ month: ym, row: MONTH_END_ROWS.oneTime[0], value: oneTime });
  if (note !== undefined) out.push({ month: ym, row: MONTH_END_ROWS.note[0], value: note });
  return out;
}

/** Giving log rows starting at sheet row 7. */
export const givingRows = (rows) => rows.map((g, i) => ({ row: 7 + i, ...g }));
