// A synthetic month of entries, applied in two passes:
//   pass 1: type balances, income, costs, investments, a note and two Giving rows;
//   pass 2 (on the pass-1 export): clear some typed cells back to automatic, clear a Giving row.

import { openWorkbook, exportWorkbook } from '../src/xlsx.js';
import { SHEET, AUTO, monthEndEdits, givingEdits } from '../src/ffos.js';

export const PASS1_MONTH_END = [
  { month: '2026-03', row: 7, value: 1000 },
  { month: '2026-04', row: 7, value: 1234.56 },
  { month: '2026-04', row: 8, value: 250 },
  { month: '2026-04', row: 9, value: 75.1 },
  { month: '2026-04', row: 22, value: 3000 },
  { month: '2026-04', row: 35, value: 400 },
  { month: '2026-04', row: 36, value: 150.25 },
  { month: '2026-04', row: 38, value: 10000 },
  { month: '2026-04', row: 45, value: 500 },
  { month: '2026-04', row: 46, value: -50 },
  { month: '2026-04', row: 60, value: 'Synthetic note: rent & <utilities> "test"' },
  { month: '2040-12', row: 20, value: 1 },
];

export const PASS1_GIVING = [
  { row: 7, date: '2026-04-05', amount: 100, recipient: 'Sample Charity A', paidFrom: 'Card 1', note: 'test row' },
  { row: 8, date: '2026-04-20', amount: 36, recipient: 'Sample Fund B', household: 'No' },
];

export const PASS2_MONTH_END = [
  { month: '2026-03', row: 7, value: AUTO }, // first column: its own formula
  { month: '2026-04', row: 9, value: AUTO },
  { month: '2026-04', row: 36, value: null }, // no formula here: becomes blank
];

export const PASS2_GIVING = [{ row: 8 }];

export function pass1Edits() {
  return new Map([
    [SHEET.monthEnd, monthEndEdits(PASS1_MONTH_END)],
    [SHEET.giving, givingEdits(PASS1_GIVING)],
  ]);
}

export function pass2Edits() {
  return new Map([
    [SHEET.monthEnd, monthEndEdits(PASS2_MONTH_END)],
    [SHEET.giving, givingEdits(PASS2_GIVING)],
  ]);
}

export async function runScenario(bytes) {
  const wb0 = await openWorkbook(bytes);
  const pass1 = await exportWorkbook(wb0, pass1Edits());
  const wb1 = await openWorkbook(pass1);
  const pass2 = await exportWorkbook(wb1, pass2Edits());
  return { pass1, pass2 };
}
