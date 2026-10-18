// Write a scenario into the template: Settings cells (test only; the app writes only the goal rows),
// Month-End entries and Giving rows. A scenario with `edit` then goes through the app: import, edit,
// and the app's own export.
import { openWorkbook, sheetXml, patchSheet, exportWorkbook } from '../../src/xlsx.js';
import { importWorkbook } from '../../src/import.js';
import * as entry from '../../src/entry.js';
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
  const bytes = await writeZip(wb.zip, changes);
  if (!sc.edit) return bytes;
  const base = await importWorkbook(bytes);
  const data = structuredClone(base);
  sc.edit(data, entry);
  return exportWorkbook(await openWorkbook(bytes), entry.workbookEdits(base, data));
}

export { dateSerial };
