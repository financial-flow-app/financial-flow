// Import -> edit -> export: only the edited cells change; every other part keeps its exact bytes.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readZip, entryText } from '../src/zip.js';
import { openWorkbook, sheetXml, readCells } from '../src/xlsx.js';
import { SHEET, automaticFormula, dateSerial } from '../src/ffos.js';
import { runScenario, pass1Edits, pass2Edits, PASS1_MONTH_END, PASS1_GIVING } from './scenario.mjs';

const FIXTURES = [
  // [file, how it was made]
  ['ffos-template.xlsx', 'as generated (openpyxl)'],
  ['ffos-template-excel.xlsx', 'saved by Excel'],
];

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));

function sheetDataSplit(xml) {
  const a = xml.indexOf('<sheetData');
  const b = xml.indexOf('</sheetData>') + '</sheetData>'.length;
  return { before: xml.slice(0, a), data: xml.slice(a, b), after: xml.slice(b) };
}

function rowsOf(xml) {
  const m = new Map();
  for (const r of xml.matchAll(/<row\b[^>]*?\br="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) m.set(Number(r[1]), r[0]);
  return m;
}

function cellsXmlOf(xml) {
  const m = new Map();
  for (const c of xml.matchAll(/<c\b[^>]*?\br="([A-Z]+\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) m.set(c[1], c[0]);
  return m;
}

/** Which parts may differ after an export, and why. */
function allowedChanges(before) {
  const names = new Set(before.entries.map((e) => e.name));
  const allowed = new Set(['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml']);
  if (names.has('xl/calcChain.xml')) {
    allowed.add('xl/calcChain.xml');
    allowed.add('[Content_Types].xml');
    allowed.add('xl/_rels/workbook.xml.rels');
  }
  return allowed;
}

async function checkExport(beforeBytes, afterBytes, sheetEdits) {
  const before = readZip(beforeBytes);
  const after = readZip(afterBytes);
  const wbBefore = await openWorkbook(beforeBytes);
  const wbAfter = await openWorkbook(afterBytes);
  const allowed = allowedChanges(before);
  const wbXmlBefore = await entryText(before.entries.find((e) => e.name === 'xl/workbook.xml'));
  if (!/fullCalcOnLoad="1"/.test(wbXmlBefore)) allowed.add('xl/workbook.xml');

  // 1. Every part outside `allowed` is byte-for-byte the same (headers and compressed data).
  const afterByName = new Map(after.entries.map((e) => [e.name, e]));
  for (const e of before.entries) {
    const a = afterByName.get(e.name);
    if (allowed.has(e.name)) continue;
    assert.ok(a, `${e.name} is missing after export`);
    assert.deepEqual(Buffer.from(a.local), Buffer.from(e.local), `${e.name}: local header changed`);
    assert.ok(Buffer.from(a.data).equals(Buffer.from(e.data)), `${e.name}: data changed`);
  }
  assert.deepEqual(after.entries.map((e) => e.name), before.entries.map((e) => e.name).filter((n) => n !== 'xl/calcChain.xml'));
  for (const e of after.entries) await entryText(e); // CRC and size checked on read

  // 2. workbook.xml: only fullCalcOnLoad may change.
  const wbXmlAfter = await entryText(afterByName.get('xl/workbook.xml'));
  assert.match(wbXmlAfter, /<calcPr\b[^>]*fullCalcOnLoad="1"/);
  assert.equal(wbXmlAfter.replace(/ fullCalcOnLoad="1"/, ''), wbXmlBefore.replace(/ fullCalcOnLoad="1"/, ''));

  // 3. In the edited sheets: outside sheetData identical; only edited cells differ.
  for (const [sheet, edits] of sheetEdits) {
    const x0 = await sheetXml(wbBefore, sheet);
    const x1 = await sheetXml(wbAfter, sheet);
    const s0 = sheetDataSplit(x0), s1 = sheetDataSplit(x1);
    assert.equal(s1.before, s0.before, `${sheet}: XML before sheetData changed`);
    assert.equal(s1.after, s0.after, `${sheet}: XML after sheetData changed (validations, conditional formats, merges...)`);
    const r0 = rowsOf(s0.data), r1 = rowsOf(s1.data);
    const editedRows = new Set([...edits.keys()].map((ref) => Number(ref.match(/\d+/)[0])));
    assert.deepEqual([...r1.keys()], [...r0.keys()], `${sheet}: rows added or removed`);
    for (const [n, xml] of r0) if (!editedRows.has(n)) assert.equal(r1.get(n), xml, `${sheet}: row ${n} changed but has no edits`);
    const c0 = cellsXmlOf(s0.data), c1 = cellsXmlOf(s1.data);
    assert.deepEqual([...c1.keys()].sort(), [...c0.keys()].sort(), `${sheet}: cells added or removed`);
    const changed = [...c0.keys()].filter((ref) => c1.get(ref) !== c0.get(ref)).sort();
    const expected = [...edits.keys()].filter((ref) => c1.get(ref) !== c0.get(ref)).sort();
    assert.deepEqual(changed, expected, `${sheet}: cells changed that weren't edited`);
    // The row's opening tag (height, style, spans) is untouched.
    for (const n of editedRows) assert.equal(r1.get(n).match(/^<row\b[^>]*>/)[0], r0.get(n).match(/^<row\b[^>]*>/)[0]);
    // Each edited cell keeps its style.
    for (const ref of edits.keys()) {
      const st = (x) => (x && x.match(/\bs="(\d+)"/)?.[1]) ?? null;
      assert.equal(st(c1.get(ref)), st(c0.get(ref)), `${sheet}!${ref}: style changed`);
    }
  }
  return wbAfter;
}

function countFeatures(xmlBySheet) {
  const n = (x, re) => (x.match(re) || []).length;
  return Object.fromEntries(
    Object.entries(xmlBySheet).map(([k, x]) => [k, {
      validations: n(x, /<dataValidation\b/g),
      conditionalFormats: n(x, /<conditionalFormatting\b/g),
      rules: n(x, /<cfRule\b/g),
      dataBars: n(x, /<dataBar\b/g),
      merges: n(x, /<mergeCell\b/g),
      protection: n(x, /<sheetProtection\b/g),
      drawings: n(x, /<drawing\b/g),
    }]),
  );
}

async function features(bytes) {
  const wb = await openWorkbook(bytes);
  const xml = {};
  for (const name of wb.sheets.keys()) xml[name] = await sheetXml(wb, name);
  const z = readZip(bytes);
  return {
    sheets: countFeatures(xml),
    states: Object.fromEntries([...wb.sheets].map(([k, v]) => [k, v.state])),
    chart: z.entries.some((e) => e.name.startsWith('xl/charts/chart')),
    autoFilter: /_xlnm\._FilterDatabase/.test(await entryText(z.entries.find((e) => e.name === 'xl/workbook.xml'))),
  };
}

for (const [file, how] of FIXTURES) {
  test(`round-trip: ${file} (${how})`, async (t) => {
    const original = fixture(file);
    const { pass1, pass2 } = await runScenario(original);

    await t.test('pass 1: only the typed cells change', async () => {
      const wb = await checkExport(original, pass1, pass1Edits());
      const me = readCells(await sheetXml(wb, SHEET.monthEnd), wb.sharedStrings);
      for (const { month, row, value } of PASS1_MONTH_END) {
        const col = { '2026-03': 'C', '2026-04': 'D', '2040-12': 'FX' }[month];
        const c = me.get(`${col}${row}`);
        assert.equal(c.formula, null, `${col}${row} should be typed`);
        assert.equal(c.value, value, `${col}${row}`);
      }
      const gv = readCells(await sheetXml(wb, SHEET.giving), wb.sharedStrings);
      for (const g of PASS1_GIVING) {
        assert.equal(gv.get(`A${g.row}`).value, dateSerial(g.date));
        assert.equal(gv.get(`B${g.row}`).value, g.amount);
        assert.equal(gv.get(`C${g.row}`).value, g.recipient);
        assert.equal(gv.get(`E${g.row}`).value, g.household ?? null);
      }
    });

    await t.test('pass 2: cleared cells get their automatic formula back', async () => {
      const wb = await checkExport(pass1, pass2, pass2Edits());
      const me = readCells(await sheetXml(wb, SHEET.monthEnd), wb.sharedStrings);
      const orig = await openWorkbook(original);
      const me0 = readCells(await sheetXml(orig, SHEET.monthEnd), orig.sharedStrings);
      for (const ref of ['C7', 'D9']) {
        assert.equal(me.get(ref).formula.text, automaticFormula(ref));
        assert.equal(me.get(ref).formula.text, me0.get(ref).formula.text, `${ref} differs from the original file`);
      }
      assert.equal(me.get('D36').kind, 'empty');
      assert.equal(me.get('D36').formula, null);
      assert.equal(me.get('D7').value, 1234.56, 'untouched typed value survives a second export');
      const gv = readCells(await sheetXml(wb, SHEET.giving), wb.sharedStrings);
      for (const c of 'ABCDEF') assert.equal(gv.get(`${c}8`).kind, 'empty');
      assert.equal(gv.get('C7').value, 'Sample Charity A');
    });

    await t.test('validations, conditional formats, merges, protection, chart and hidden sheet intact', async () => {
      const f0 = await features(original);
      assert.deepEqual(await features(pass1), f0);
      assert.deepEqual(await features(pass2), f0);
      assert.equal(f0.states.Calc, 'hidden');
      assert.ok(f0.chart);
      assert.ok(f0.sheets['Month-End'].validations > 0 && f0.sheets['Month-End'].conditionalFormats > 0);
      assert.ok(f0.sheets.Dashboard.dataBars > 0 && f0.sheets.Dashboard.protection === 1);
    });
  });
}

test('the Excel-saved fixture has the features an Excel-saved file brings', async () => {
  const z = readZip(fixture('ffos-template-excel.xlsx'));
  const names = z.entries.map((e) => e.name);
  assert.ok(names.includes('xl/sharedStrings.xml'));
  assert.ok(names.includes('xl/calcChain.xml'));
  const { pass1 } = await runScenario(fixture('ffos-template-excel.xlsx'));
  const out = readZip(pass1).entries.map((e) => e.name);
  assert.ok(!out.includes('xl/calcChain.xml'), 'calcChain removed on export');
  const ct = await entryText(readZip(pass1).entries.find((e) => e.name === '[Content_Types].xml'));
  assert.ok(!ct.includes('calcChain'));
});
