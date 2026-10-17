// Write a scenario into the template: Settings cells (test only; the app never writes Settings),
// Month-End entries and Giving rows.
import { openWorkbook, sheetXml, patchSheet } from '../../src/xlsx.js';
import { writeZip } from '../../src/zip.js';
import { SHEET, monthEndEdits, givingEdits, dateSerial } from '../../src/ffos.js';

const enc = new TextEncoder();

export async function buildScenario(templateBytes, sc) {
  const wb = await openWorkbook(templateBytes);
  const changes = new Map();
  const patch = async (sheet, edits) => {
    changes.set(wb.sheets.get(sheet).path, enc.encode(patchSheet(await sheetXml(wb, sheet), edits)));
  };
  await patch(SHEET.settings, new Map(Object.entries(sc.settings)));
  await patch(SHEET.monthEnd, monthEndEdits(sc.monthEnd));
  await patch(SHEET.giving, givingEdits(sc.giving.map((g) => ({ ...g, date: g.date ?? null })), { tax: sc.giving.some((g) => 'deductible' in g) }));
  return writeZip(wb.zip, changes);
}

export { dateSerial };
