// Monthly entry logic, kept apart from the screen so it can be tested:
// parsing amounts, which month to fill, what each field shows, and turning
// the user's entries into workbook edits for export.

import { isBlank, N, is, DATE, YEAR, MONTH, DAYOF } from './excel.js';
import {
  MONTH_END_ROWS, AUTO, SHEET, monthEndEdits, givingEdits, goalEdits, GIVING_FIRST_ROW, GIVING_LAST_ROW, GOAL_COLS, GOAL_TYPES,
} from './ffos.js';
import { METRIC_NAMES } from './calc.js';

const LIMIT = 1e9; // the workbook's own validation: a number between -1,000,000,000 and 1,000,000,000

/** "1,234.50", "$1,234", "-50", "(50)" -> number. Blank -> null. Anything else -> {error}. */
export function parseAmount(text) {
  let t = String(text ?? '').trim();
  if (t === '') return { value: null };
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  t = t.replace(/[\s$,]/g, '');
  if (t.startsWith('-')) { neg = !neg; t = t.slice(1); }
  if (!/^(\d+\.?\d*|\.\d+)$/.test(t)) return { error: 'Enter a number, like 1234.56' };
  const v = (neg ? -1 : 1) * Number(t);
  if (Math.abs(v) > LIMIT) return { error: 'That number is too large for the workbook.' };
  return { value: v };
}

/** The sections of a month, in Month-End order: data key, sheet rows. */
const SECTIONS = [
  ['balances', MONTH_END_ROWS.balances],
  ['income', MONTH_END_ROWS.income],
  ['family', MONTH_END_ROWS.familyPaid],
  ['oneTime', MONTH_END_ROWS.oneTime],
  ['invValues', MONTH_END_ROWS.investValues],
  ['deposits', MONTH_END_ROWS.investDeposits],
];

const cellsOf = (m, key) => (Array.isArray(m[key]) ? m[key] : [m[key]]);
const sameCell = (a, b) => a.kind === b.kind && a.value === b.value;

/** Set one Month-End input: a number types it, null puts it back to automatic. Returns false if not editable. */
export function setEntry(data, monthIndex, key, slot, value) {
  const m = data.months[monthIndex];
  const cell = cellsOf(m, key)[slot ?? 0];
  if (cell.kind === 'formula') return false; // a formula the user set in Excel: change it there
  const next = value === null ? { kind: 'auto', value: null } : { kind: 'typed', value };
  if (Array.isArray(m[key])) m[key][slot] = next;
  else m[key] = next;
  return true;
}

export function setNote(data, monthIndex, text) {
  const t = String(text ?? '').trim();
  data.months[monthIndex].note = t === '' ? null : t;
}

// ---------- Giving ----------

/** Add a donation in the first empty row of the log. Returns the row, or null if the log is full. */
export function addGift(data, gift) {
  const used = new Set(data.giving.map((g) => g.row));
  for (let r = GIVING_FIRST_ROW; r <= GIVING_LAST_ROW; r++) {
    if (!used.has(r)) {
      data.giving.push(cleanGift({ ...gift, row: r }, data.hasTax));
      data.giving.sort((a, b) => a.row - b.row);
      return r;
    }
  }
  return null;
}

export function updateGift(data, row, gift) {
  const i = data.giving.findIndex((g) => g.row === row);
  if (i >= 0) data.giving[i] = cleanGift({ ...gift, row }, data.hasTax);
}

export function deleteGift(data, row) {
  data.giving = data.giving.filter((g) => g.row !== row);
}

function cleanGift(g, tax) {
  const text = (v) => (isBlank(v) || String(v).trim() === '' ? null : String(v).trim());
  return {
    row: g.row,
    date: typeof g.date === 'number' ? g.date : null,
    amount: typeof g.amount === 'number' ? g.amount : null,
    recipient: text(g.recipient),
    paidFrom: text(g.paidFrom),
    household: text(g.household),
    note: text(g.note),
    ...(tax ? { deductible: text(g.deductible) } : {}),
  };
}

// ---------- Goals (Settings rows 69-88) ----------

const GOAL_FIELDS = Object.keys(GOAL_COLS);
const SPECIFIC = 'Specific Account';

/** Which fields a goal type uses. Tithe Balance takes its target and current value from this quarter's tithe. */
export function goalNeeds(type) {
  const t = String(type ?? '').toLowerCase();
  const tithe = t === 'tithe balance';
  return {
    target: !tithe,
    basedOn: t === 'percentage' || t === 'multiple',
    currentFrom: !tithe,
    targetDate: t === 'save by date',
  };
}

/** The choices the goal form offers, all from the workbook: types, metrics and the named accounts. */
export function goalChoices(data) {
  const accounts = [...data.accounts, ...data.investments].map((a) => a.name).filter((n) => !isBlank(n)).map(String);
  return { types: GOAL_TYPES, metrics: METRIC_NAMES, accounts };
}

/** A goal as the form saves it: fields its type doesn't use are blank; the account only with Specific Account. */
export function cleanGoal(g) {
  const text = (v) => (isBlank(v) || String(v).trim() === '' ? null : typeof v === 'number' ? v : String(v).trim());
  const need = goalNeeds(g.type);
  const basedOn = need.basedOn ? text(g.basedOn) : null;
  const currentFrom = need.currentFrom ? text(g.currentFrom) : null;
  const usesAccount = is(basedOn, SPECIFIC) || is(currentFrom, SPECIFIC);
  return {
    name: text(g.name),
    type: text(g.type),
    target: need.target && typeof g.target === 'number' ? g.target : null,
    basedOn,
    currentFrom,
    account: usesAccount ? text(g.account) : null,
    targetDate: need.targetDate && typeof g.targetDate === 'number' ? g.targetDate : null,
  };
}

/** What's missing before a goal can be saved, or null. A Save By Date goal may leave its date blank. */
export function goalProblem(g) {
  const need = goalNeeds(g.type);
  if (isBlank(g.name) || String(g.name).trim() === '') return 'Give the goal a name.';
  if (isBlank(g.type)) return 'Pick a type.';
  if (need.target && typeof g.target !== 'number') return 'Enter the target.';
  if (need.basedOn && isBlank(g.basedOn)) return 'Pick what the target is based on.';
  if (need.currentFrom && isBlank(g.currentFrom)) return 'Pick where the current value comes from.';
  if ((is(g.basedOn, SPECIFIC) && need.basedOn) || (is(g.currentFrom, SPECIFIC) && need.currentFrom)) {
    if (isBlank(g.account)) return 'Pick the account.';
  }
  return null;
}

/** Add a goal in the first slot with a blank name. Returns the slot, or null when all 20 are used. */
export function addGoal(data, goal) {
  const slot = data.goals.findIndex((g) => isBlank(g.name));
  if (slot < 0) return null;
  data.goals[slot] = cleanGoal(goal);
  return slot;
}

export function updateGoal(data, slot, goal) {
  data.goals[slot] = cleanGoal(goal);
}

/** Clear the goal's row (a blank name means an unused slot). */
export function deleteGoal(data, slot) {
  data.goals[slot] = Object.fromEntries(GOAL_FIELDS.map((k) => [k, null]));
}

// ---------- Changes and export ----------

/** Workbook edits that turn `base` (as imported) into `current`. */
export function workbookEdits(base, current) {
  const { me, gv, gl } = differences(base, current);
  return new Map([
    [SHEET.monthEnd, monthEndEdits(me)],
    [SHEET.giving, givingEdits(gv, { tax: current.hasTax })],
    [SHEET.settings, goalEdits(gl)],
  ]);
}

/**
 * What differs between two copies of the data: Month-End fields (one per field, with its month),
 * Giving rows (one per donation added, changed or removed) and goals (one per goal slot changed).
 */
function differences(base, current) {
  const me = [];
  current.months.forEach((m, i) => {
    const b = base.months[i];
    for (const [key, rows] of SECTIONS) {
      cellsOf(m, key).forEach((cell, slot) => {
        if (sameCell(cell, cellsOf(b, key)[slot])) return;
        me.push({ month: m.ym, row: rows[slot], value: cell.kind === 'typed' ? cell.value : AUTO });
      });
    }
    if (m.note !== b.note) me.push({ month: m.ym, row: MONTH_END_ROWS.note[0], value: m.note });
  });
  const fields = ['date', 'amount', 'recipient', 'paidFrom', 'household', 'note', ...(current.hasTax ? ['deductible'] : [])];
  const before = new Map(base.giving.map((g) => [g.row, g]));
  const after = new Map(current.giving.map((g) => [g.row, g]));
  const gv = [];
  for (const row of new Set([...before.keys(), ...after.keys()])) {
    const a = before.get(row), c = after.get(row);
    if (a && c && fields.every((f) => (a[f] ?? null) === (c[f] ?? null))) continue;
    gv.push(c ?? { row });
  }
  gv.sort((x, y) => x.row - y.row);
  const gl = [];
  current.goals.forEach((g, slot) => {
    const b = base.goals[slot];
    if (!GOAL_FIELDS.every((f) => (g[f] ?? null) === (b?.[f] ?? null))) gl.push({ ...g, slot });
  });
  return { me, gv, gl };
}

/** Things done since `base`: each Month-End field typed or cleared, each donation and each goal added, changed or deleted counts 1. */
export function changeSummary(base, current) {
  const { me, gv, gl } = differences(base, current);
  return {
    fields: me.length, gifts: gv.length, goals: gl.length, total: me.length + gv.length + gl.length,
    months: [...new Set(me.map((x) => x.month))],
  };
}

export const changeCount = (base, current) => changeSummary(base, current).total;

// ---------- Export file name and the older-copy check ----------

// A date and time an earlier export put in the name: "2026-10-17 2.32 PM" (or the older "2026-10-17 1432").
const STAMP_IN_NAME = /\s\d{4}-\d{2}-\d{2} (\d{1,2}\.\d{2} [AP]M|\d{4})$/i;

/**
 * "FFOS.xlsx" or "FFOS 2026-10-01 9.02 AM.xlsx" -> "FFOS 2026-10-18 12.01 AM.xlsx".
 * Local time, 12-hour, with a dot instead of ":" (Windows doesn't allow ":" in file names).
 */
export function exportFileName(name, at = new Date()) {
  const base = String(name || 'FFOS.xlsx').replace(/\.xlsx$/i, '').replace(STAMP_IN_NAME, '') || 'FFOS';
  const p = (n) => String(n).padStart(2, '0');
  const h = at.getHours() % 12 || 12;
  const ampm = at.getHours() < 12 ? 'AM' : 'PM';
  return `${base} ${at.getFullYear()}-${p(at.getMonth() + 1)}-${p(at.getDate())} ${h}.${p(at.getMinutes())} ${ampm}.xlsx`;
}

/** The ± button: "50" -> "-50", "-50" -> "50", "" -> "-" (ready to type), "-" -> "". Text that isn't a number stays as it is. */
export function flipSign(text) {
  const t = String(text ?? '').trim();
  if (t === '') return '-';
  if (t === '-') return '';
  if (parseAmount(t).error) return t;
  if (t.startsWith('-')) return t.slice(1).trim();
  if (/^\(.*\)$/.test(t)) return t.slice(1, -1).trim();
  return `-${t}`;
}

/**
 * Before importing: is the file an older copy than what this phone last exported?
 *   exported    the data as this phone last exported it (null if it never exported)
 *   exportedAt  that export's time (ISO)
 *   incoming    the file being imported; incoming.exportStamp is the stamp inside it (ISO or null)
 * Warns only when the file's stamp is missing or older than the phone's last export AND the file
 * differs from that export. Edits made in Excel after the latest export never warn.
 */
export function olderCopyCheck({ exported, exportedAt }, incoming) {
  if (!exported || !exportedAt) return { warn: false };
  const stamp = incoming.exportStamp ?? null;
  if (stamp && stamp >= exportedAt) return { warn: false, stamp };
  const diff = changeSummary(incoming, exported);
  return { warn: diff.total > 0, stamp, missing: diff.total, months: diff.months, gifts: diff.gifts, goals: diff.goals };
}

// ---------- Which month, and what the screen shows ----------

/** Index of the month containing `today` (clamped to the workbook's range). */
export function currentMonthIndex(months, today) {
  const i = months.findIndex((r) => r.MonthStart <= today && today <= r.MonthEnd);
  return i >= 0 ? i : today < months[0].MonthStart ? 0 : months.length - 1;
}

/** The month the workbook's Dashboard says to fill next ("Next: fill <month>"). */
export function nextMonthToFill(out, today) {
  const asOf = out.metrics['As-of date'];
  const start = DATE(YEAR(asOf), MONTH(asOf) + (DAYOF(asOf + 1) === 1 ? 1 : 0), 1);
  const i = out.months.findIndex((r) => r.MonthStart === start);
  const now = currentMonthIndex(out.months, today);
  return i >= 0 && i <= now ? i : now;
}

/**
 * Everything the month screen needs for month i.
 * Each field: {key, slot, label, cell, auto, required, hint}. `auto` is the grey value shown
 * when nothing is typed (null = none).
 */
export function monthView(data, out, i) {
  const r = out.months[i];
  const m = data.months[i];
  const future = r.Future === 1;
  const pre = r.Phase === 'pre';
  const accounts = [];
  data.accounts.forEach((a, k) => {
    if (isBlank(a.name)) return;
    const required = is(a.required, 'Yes') && !is(a.active, 'No');
    const retired = is(a.active, 'No');
    const owed = is(a.type, 'Card') || is(a.type, 'Loan');
    accounts.push({
      key: 'balances', slot: k, label: String(a.name), cell: m.balances[k], required, retired,
      hint: [a.type, owed ? 'amount owed' : null, retired ? 'retired' : null].filter(Boolean).join(' · '),
      auto: required || future || pre || m.balances[k].kind !== 'auto' ? null : r[`Val${k}`],
    });
  });
  accounts.sort((x, y) => (y.required - x.required) || (x.retired - y.retired) || (x.slot - y.slot));
  const income = [];
  data.sources.forEach((s, k) => {
    if (isBlank(s.name)) return;
    const auto = r[`Auto${k}`];
    income.push({
      key: 'income', slot: k, label: String(s.name), cell: m.income[k],
      hint: [
        is(s.tithe, 'Yes') ? 'tithe-able' : is(s.tithe, 'Not Sure') ? 'tithe: not sure' : null,
        data.hasTax ? taxHint(s) : null,
      ].filter(Boolean).join(' · ') || null,
      auto: m.income[k].kind === 'auto' && auto ? auto : null,
    });
  });
  const investments = [];
  data.investments.forEach((inv, k) => {
    if (isBlank(inv.name)) return;
    investments.push({
      key: 'invValues', slot: k, label: String(inv.name), cell: m.invValues[k], hint: String(inv.type ?? ''),
      auto: future || pre || m.invValues[k].kind !== 'auto' ? null : r[`Val${14 + k}`],
      deposit: { key: 'deposits', slot: k, label: `Moved into ${inv.name}`, cell: m.deposits[k], hint: '+ in / − out', auto: null },
    });
  });
  return {
    ym: m.ym,
    status: r.Status,
    phase: r.Phase,
    accounts: pre ? [] : accounts,
    missing: accounts.filter((f) => f.required && f.cell.kind !== 'typed').length,
    income,
    family: { key: 'family', slot: null, label: 'Family-paid costs', cell: m.family, hint: 'paid for you by family', auto: m.family.kind === 'auto' && r.FamAuto ? r.FamAuto : null },
    oneTime: { key: 'oneTime', slot: null, label: 'One-time costs', cell: m.oneTime, hint: 'part of spending, shown apart', auto: null },
    investments: pre ? [] : investments,
    note: m.note,
    results: {
      income: r.Income, household: r.HHSpend, family: r.Family, trueCost: r.TrueCost,
      given: r.Given, saved: r.Saved, netWorth: r.NetWorth,
    },
  };
}

// ---------- Export reminder ----------

export const REMIND_AFTER_DAYS = 30; // suggest exporting monthly

/**
 * "Last export: N days ago", and whether to remind: there are entries Excel doesn't have yet,
 * and the last export (or the import, if never exported) is a month old.
 * Times are milliseconds; days count calendar days on this phone.
 */
export function exportReminder({ lastExport, importedAt, changes }, now = Date.now()) {
  const day = (ms) => { const d = new Date(ms); return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000; };
  const since = (ms) => (ms ? day(now) - day(ms) : null);
  const days = since(lastExport);
  const text = days === null ? 'Not exported yet' : days === 0 ? 'Last export: today' : `Last export: ${days} day${days === 1 ? '' : 's'} ago`;
  const age = since(lastExport ?? importedAt);
  const due = changes > 0 && age !== null && age >= REMIND_AFTER_DAYS;
  return { days, text, due };
}

/** "taxable, withheld" etc. from a source's tax marks (FFOS 3.8); null when unmarked. */
export function taxHint(s) {
  if (is(s.taxable, 'Yes')) return is(s.withheld, 'Yes') ? 'taxable, withheld' : is(s.withheld, 'No') ? 'taxable, not withheld' : 'taxable';
  if (is(s.taxable, 'No')) return 'not taxable';
  if (is(s.taxable, 'Not Sure')) return 'tax: not sure';
  return null;
}

/** Number of typed entries, for the month list. */
export const typedCount = (m) => SECTIONS.reduce((n, [key]) => n + cellsOf(m, key).filter((c) => c.kind === 'typed').length, 0);

export { N };
