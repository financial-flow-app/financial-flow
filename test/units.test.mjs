import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readZip, writeZip, entryText } from '../src/zip.js';
import { shiftFormula, formulaRefs, numToCol, colToNum } from '../src/cellref.js';
import { openWorkbook, sheetXml, readCells, patchSheet, cellXml } from '../src/xlsx.js';
import { automaticFormula, MONTH_END_INPUT_ROWS, monthCol, dateSerial, serialToIso, AUTO, monthEndEdits } from '../src/ffos.js';

const fixture = (f) => readFileSync(new URL(`../fixtures/${f}`, import.meta.url));

test('zip: writing with no changes gives the identical file', async () => {
  for (const f of ['ffos-template.xlsx', 'ffos-template-excel.xlsx']) {
    const b = fixture(f);
    assert.ok(Buffer.from(await writeZip(readZip(b))).equals(b), f);
  }
});

test('zip: a changed entry reads back, others keep their bytes', async () => {
  const z = readZip(fixture('ffos-template.xlsx'));
  const text = 'hello <world>';
  const out = readZip(await writeZip(z, new Map([['docProps/app.xml', new TextEncoder().encode(text)]])));
  assert.equal(await entryText(out.entries.find((e) => e.name === 'docProps/app.xml')), text);
  const a = out.entries.find((e) => e.name === 'xl/styles.xml'), b = z.entries.find((e) => e.name === 'xl/styles.xml');
  assert.ok(Buffer.from(a.data).equals(Buffer.from(b.data)));
});

test('columns and months', () => {
  assert.equal(numToCol(180), 'FX');
  assert.equal(colToNum('FX'), 180);
  assert.equal(monthCol('2026-03'), 3);
  assert.equal(monthCol('2040-12'), 180);
  assert.throws(() => monthCol('2041-01'));
  assert.throws(() => monthCol('2026-02'));
  assert.equal(dateSerial('2026-03-01'), 46082); // the template's C4
  assert.equal(serialToIso(46082), '2026-03-01');
});

test('formula shifting follows Excel copy rules', () => {
  assert.equal(shiftFormula('A1+$B$2+C$3+$D4', 1, 1), 'B2+$B$2+D$3+$D5');
  assert.equal(shiftFormula("'Month-End'!C7+Calc!$AE$6", 1, 0), "'Month-End'!D7+Calc!$AE$6");
  assert.equal(shiftFormula('IF(A1="B2",LOG10(A1),A1:B2)', 0, 1), 'IF(A2="B2",LOG10(A2),A2:B3)');
  assert.equal(formulaRefs('EOMONTH(C$4,0)<=Settings!$B$5').length, 2);
});

test('automatic formulas: the map reproduces every input formula in the template', async () => {
  const wb = await openWorkbook(fixture('ffos-template.xlsx'));
  const cells = readCells(await sheetXml(wb, 'Month-End'));
  let n = 0;
  for (const row of MONTH_END_INPUT_ROWS) {
    for (let col = 3; col <= 180; col++) {
      const ref = numToCol(col) + row;
      assert.equal(automaticFormula(ref), cells.get(ref)?.formula?.text ?? null, ref);
      n++;
    }
  }
  assert.equal(n, MONTH_END_INPUT_ROWS.length * 178);
});

test('automatic formulas match the Excel-saved file too', async () => {
  const wb = await openWorkbook(fixture('ffos-template-excel.xlsx'));
  const cells = readCells(await sheetXml(wb, 'Month-End'), wb.sharedStrings);
  for (const row of MONTH_END_INPUT_ROWS) {
    for (let col = 3; col <= 180; col++) {
      const ref = numToCol(col) + row;
      assert.equal(automaticFormula(ref), cells.get(ref)?.formula?.text ?? null, ref);
    }
  }
});

test('edits reject non-input rows and non-numbers', () => {
  assert.throws(() => monthEndEdits([{ month: '2026-04', row: 52, value: 1 }]));
  assert.throws(() => monthEndEdits([{ month: '2026-04', row: 7, value: '12' }]));
  assert.throws(() => monthEndEdits([{ month: '2026-04', row: 7, value: NaN }]));
  assert.deepEqual(monthEndEdits([{ month: '2026-04', row: 45, value: AUTO }]).get('D45'), null);
});

test('text cells are escaped and strip characters XML forbids', () => {
  assert.equal(cellXml('A1', '3', 'a & <b>\u0001'), '<c r="A1" s="3" t="inlineStr"><is><t>a &amp; &lt;b&gt;</t></is></c>');
  assert.equal(cellXml('A1', null, ' x'), '<c r="A1" t="inlineStr"><is><t xml:space="preserve"> x</t></is></c>');
});

test('shared formulas: overwriting the master hands it to the next cell', () => {
  const xml =
    '<worksheet><sheetData>' +
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
    '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
    '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>' +
    '</sheetData></worksheet>';
  const out = patchSheet(xml, new Map([['B1', 99]]));
  assert.ok(out.includes('<c r="B1"><v>99</v></c>'));
  assert.ok(out.includes('<c r="B2"><f t="shared" ref="B2:B3" si="0">A2*2</f><v>4</v></c>'));
  assert.ok(out.includes('<c r="B3"><f t="shared" si="0"/><v>6</v></c>'));
  // Overwriting a non-master member leaves the group alone.
  const out2 = patchSheet(xml, new Map([['B3', 5]]));
  assert.ok(out2.includes('<f t="shared" ref="B1:B3" si="0">A1*2</f>'));
});

test('missing cells and rows are inserted in order', () => {
  const xml = '<worksheet><sheetData><row r="2" spans="1:2"><c r="A2"><v>1</v></c><c r="C2"><v>3</v></c></row></sheetData></worksheet>';
  const out = patchSheet(xml, new Map([['B2', 2], ['A1', 'x'], ['A5', 5], ['E2', 7]]));
  assert.equal(
    out,
    '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>' +
      '<row r="2" spans="1:5"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="E2"><v>7</v></c></row>' +
      '<row r="5"><c r="A5"><v>5</v></c></row></sheetData></worksheet>',
  );
});
