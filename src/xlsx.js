// Read and patch cells in an .xlsx without rewriting the rest of the file.
// Sheet XML is edited as text: only the <row> elements that hold an edited cell are rebuilt,
// and inside them only the edited <c> elements change. Everything else keeps its exact bytes.

import { readZip, writeZip, entryText } from './zip.js';
import { parseRef, makeRef, shiftFormula } from './cellref.js';

const enc = new TextEncoder();

export function xmlEscape(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function xmlUnescape(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, e) => {
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'amp') return '&';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}

function attr(attrs, name) {
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs);
  return m ? xmlUnescape(m[1]) : null;
}

function resolveTarget(target) {
  if (target.startsWith('/')) return target.slice(1);
  return 'xl/' + target;
}

// ---------- workbook ----------

export async function openWorkbook(bytes) {
  const zip = readZip(bytes);
  const byName = new Map(zip.entries.map((e) => [e.name, e]));
  const wbXml = await entryText(byName.get('xl/workbook.xml'));
  const relsXml = await entryText(byName.get('xl/_rels/workbook.xml.rels'));
  const rels = new Map();
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    rels.set(attr(m[1], 'Id'), { type: attr(m[1], 'Type'), target: resolveTarget(attr(m[1], 'Target')) });
  }
  const sheets = new Map();
  for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const name = attr(m[1], 'name');
    const rel = rels.get(attr(m[1], 'r:id'));
    sheets.set(name, { path: rel.target, state: attr(m[1], 'state') || 'visible' });
  }
  let sharedStrings = [];
  const ssRel = [...rels.values()].find((r) => r.type.endsWith('/sharedStrings'));
  if (ssRel && byName.has(ssRel.target)) sharedStrings = parseSharedStrings(await entryText(byName.get(ssRel.target)));
  return { zip, byName, sheets, sharedStrings, texts: new Map() };
}

/** Text of a part, cached. */
export async function partText(wb, path) {
  if (!wb.texts.has(path)) wb.texts.set(path, await entryText(wb.byName.get(path)));
  return wb.texts.get(path);
}

export async function sheetXml(wb, sheetName) {
  const s = wb.sheets.get(sheetName);
  if (!s) throw new Error(`The workbook has no sheet named ${sheetName}.`);
  return partText(wb, s.path);
}

function parseSharedStrings(xml) {
  const out = [];
  for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const body = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let s = '';
    for (const t of body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g)) s += xmlUnescape(t[1] || '');
    out.push(s);
  }
  return out;
}

// ---------- reading cells ----------

const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

/**
 * Parse one <c> element.
 * Returns {ref, style, formula, value, kind}, where kind is
 * 'empty' | 'number' | 'text' | 'bool' | 'error' and `formula` is null for typed (constant) cells.
 */
export function parseCell(attrs, inner = '', sharedStrings = []) {
  const ref = attr(attrs, 'r');
  const t = attr(attrs, 't') || 'n';
  const style = attr(attrs, 's');
  const fm = inner && /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
  const formula = fm ? { text: fm[2] != null ? xmlUnescape(fm[2]) : '', attrs: fm[1] } : null;
  const vm = inner && /<v>([\s\S]*?)<\/v>/.exec(inner);
  const v = vm ? xmlUnescape(vm[1]) : null;
  let kind = 'empty';
  let value = null;
  if (t === 'inlineStr') {
    const im = /<is>([\s\S]*?)<\/is>/.exec(inner || '');
    if (im) {
      kind = 'text';
      value = '';
      for (const tm of im[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) value += xmlUnescape(tm[1]);
    }
  } else if (v !== null && v !== '') {
    if (t === 's') { kind = 'text'; value = sharedStrings[Number(v)]; }
    else if (t === 'str') { kind = 'text'; value = v; }
    else if (t === 'b') { kind = 'bool'; value = v === '1'; }
    else if (t === 'e') { kind = 'error'; value = v; }
    else { kind = 'number'; value = Number(v); }
  }
  return { ref, style, formula, value, kind };
}

/** All cells of a sheet as a Map ref -> parsed cell. Optional filter(ref) keeps only some. */
export function readCells(xml, sharedStrings = [], filter = null) {
  const out = new Map();
  const sd = sheetDataBounds(xml);
  const body = xml.slice(sd.start, sd.end);
  for (const m of body.matchAll(CELL_RE)) {
    const ref = attr(m[1], 'r');
    if (filter && !filter(ref)) continue;
    out.set(ref, parseCell(m[1], m[2], sharedStrings));
  }
  return out;
}

// ---------- editing cells ----------

function sheetDataBounds(xml) {
  const open = /<sheetData\b[^>]*?(\/?)>/.exec(xml);
  if (!open) throw new Error('Sheet has no sheetData.');
  const start = open.index + open[0].length;
  if (open[1] === '/') return { start, end: start, selfClosing: true, openTag: open };
  const end = xml.indexOf('</sheetData>', start);
  return { start, end, selfClosing: false, openTag: open };
}

/** XML for a cell holding `value`. value: number | string | {formula} | null (blank). */
export function cellXml(ref, style, value) {
  const s = style != null ? ` s="${style}"` : '';
  if (value === null || value === undefined || value === '') return `<c r="${ref}"${s}/>`;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error(`Cell ${ref}: not a finite number.`);
    return `<c r="${ref}"${s}><v>${String(value)}</v></c>`;
  }
  if (typeof value === 'string') {
    // Characters XML 1.0 forbids would make the file unreadable.
    const clean = value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
    const sp = /^\s|\s$/.test(clean) ? ' xml:space="preserve"' : '';
    return `<c r="${ref}"${s} t="inlineStr"><is><t${sp}>${xmlEscape(clean)}</t></is></c>`;
  }
  if (typeof value === 'object' && typeof value.formula === 'string') {
    return `<c r="${ref}"${s}><f>${xmlEscape(value.formula)}</f></c>`;
  }
  throw new Error(`Cell ${ref}: unsupported value.`);
}

function splitRows(body) {
  // Returns [{row, start, end, xml}] for each <row> element in sheetData body.
  const rows = [];
  for (const m of body.matchAll(/<row\b([^>]*?)(?:\/>|>[\s\S]*?<\/row>)/g)) {
    rows.push({ row: Number(attr(m[1], 'r')), start: m.index, end: m.index + m[0].length, xml: m[0] });
  }
  return rows;
}

function rowCells(rowXml) {
  const open = /^<row\b[^>]*?(\/?)>/.exec(rowXml);
  const head = open[0].replace(/\/>$/, '>');
  const inner = open[1] === '/' ? '' : rowXml.slice(open[0].length, rowXml.length - '</row>'.length);
  const cells = [];
  for (const m of inner.matchAll(CELL_RE)) {
    cells.push({ ref: attr(m[1], 'r'), col: parseRef(attr(m[1], 'r')).col, xml: m[0], attrs: m[1], inner: m[2] || '' });
  }
  // Text between cells (if any) is kept only when it's whitespace; Excel writes none.
  return { head, cells };
}

function bboxRef(refs) {
  const ps = refs.map(parseRef);
  const c1 = Math.min(...ps.map((p) => p.col)), c2 = Math.max(...ps.map((p) => p.col));
  const r1 = Math.min(...ps.map((p) => p.row)), r2 = Math.max(...ps.map((p) => p.row));
  const a = makeRef(c1, r1), b = makeRef(c2, r2);
  return a === b ? a : `${a}:${b}`;
}

/**
 * Apply cell edits to a sheet's XML.
 * edits: Map ref -> value (number | string | {formula} | null for blank).
 * The cell's style (s) is kept. Shared formulas whose master cell is overwritten are handed
 * to the next cell of the group, so the other cells keep their formulas.
 */
export function patchSheet(xml, edits) {
  if (edits.size === 0) return xml;
  const sd = sheetDataBounds(xml);
  const body = xml.slice(sd.start, sd.end);
  const rows = splitRows(body);
  const rowIndex = new Map(rows.map((r, i) => [r.row, i]));

  // Final XML for each touched cell: ref -> xml string.
  const newCells = new Map();
  const existing = new Map(); // ref -> parsed existing cell (for edited refs)
  for (const ref of edits.keys()) {
    const { row } = parseRef(ref);
    const i = rowIndex.get(row);
    if (i === undefined) continue;
    const c = rowCells(rows[i].xml).cells.find((x) => x.ref === ref);
    if (c) existing.set(ref, c);
  }

  for (const [ref, value] of edits) {
    const old = existing.get(ref);
    const style = old ? attr(old.attrs, 's') : null;
    if (old) {
      const f = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(old.inner);
      if (f && attr(f[1], 't') === 'array') throw new Error(`Cell ${ref} is part of an array formula; not editable.`);
      if (f && attr(f[1], 't') === 'shared' && attr(f[1], 'ref')) {
        handOverSharedFormula(body, rows, ref, attr(f[1], 'si'), xmlUnescape(f[2] || ''), edits, newCells);
      }
    }
    newCells.set(ref, cellXml(ref, style, value));
  }

  // Group touched cells by row and rebuild those rows.
  const byRow = new Map();
  for (const ref of newCells.keys()) {
    const { row } = parseRef(ref);
    if (!byRow.has(row)) byRow.set(row, []);
    byRow.get(row).push(ref);
  }
  const replaced = new Map(); // row index -> new xml
  const added = []; // new rows: {row, xml}
  for (const [row, refs] of byRow) {
    const i = rowIndex.get(row);
    if (i === undefined) {
      const sorted = refs.map((r) => ({ ref: r, col: parseRef(r).col })).sort((a, b) => a.col - b.col);
      added.push({ row, xml: `<row r="${row}">${sorted.map((c) => newCells.get(c.ref)).join('')}</row>` });
      continue;
    }
    const { head, cells } = rowCells(rows[i].xml);
    const out = cells.map((c) => ({ col: c.col, xml: newCells.has(c.ref) ? newCells.get(c.ref) : c.xml }));
    for (const ref of refs) {
      if (!cells.some((c) => c.ref === ref)) out.push({ col: parseRef(ref).col, xml: newCells.get(ref) });
    }
    out.sort((a, b) => a.col - b.col);
    let newHead = head;
    const spans = attr(head, 'spans');
    if (spans) {
      const [a, b] = spans.split(':').map(Number);
      const lo = Math.min(a, out[0].col), hi = Math.max(b, out[out.length - 1].col);
      if (lo !== a || hi !== b) newHead = head.replace(/spans="[^"]*"/, `spans="${lo}:${hi}"`);
    }
    replaced.set(i, newHead + out.map((c) => c.xml).join('') + '</row>');
  }

  // Reassemble sheetData, inserting new rows in order.
  added.sort((a, b) => a.row - b.row);
  let outBody = '';
  let last = 0;
  let k = 0;
  rows.forEach((r, i) => {
    outBody += body.slice(last, r.start);
    while (k < added.length && added[k].row < r.row) outBody += added[k++].xml;
    outBody += replaced.has(i) ? replaced.get(i) : r.xml;
    last = r.end;
  });
  while (k < added.length) outBody += added[k++].xml;
  outBody += body.slice(last);

  if (sd.selfClosing) {
    const tag = sd.openTag[0].replace(/\/>$/, '>');
    return xml.slice(0, sd.openTag.index) + tag + outBody + '</sheetData>' + xml.slice(sd.start);
  }
  return xml.slice(0, sd.start) + outBody + xml.slice(sd.end);
}

function handOverSharedFormula(body, rows, masterRef, si, masterFormula, edits, newCells) {
  // Every other cell in the group that is not itself being overwritten.
  const members = [];
  for (const m of body.matchAll(CELL_RE)) {
    if (!m[2] || !m[2].includes('<f')) continue;
    const f = /<f\b([^>]*?)(?:\/>|>)/.exec(m[2]);
    if (!f || attr(f[1], 't') !== 'shared' || attr(f[1], 'si') !== si) continue;
    const ref = attr(m[1], 'r');
    if (ref !== masterRef && !edits.has(ref)) members.push(ref);
  }
  if (members.length === 0) return;
  members.sort((a, b) => {
    const pa = parseRef(a), pb = parseRef(b);
    return pa.row - pb.row || pa.col - pb.col;
  });
  const newMaster = members[0];
  const m = parseRef(masterRef), n = parseRef(newMaster);
  const formula = shiftFormula(masterFormula, n.col - m.col, n.row - m.row);
  // Rebuild the new master cell: keep its attributes and value, replace its <f>.
  const { row } = n;
  const rowXml = rows.find((r) => r.row === row).xml;
  const cell = rowCells(rowXml).cells.find((c) => c.ref === newMaster);
  const fTag = `<f t="shared" ref="${bboxRef(members)}" si="${si}">${xmlEscape(formula)}</f>`;
  const inner = cell.inner.replace(/<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/, fTag);
  newCells.set(newMaster, `<c${cell.attrs}>${inner}</c>`);
}

// ---------- workbook-level fixes on export ----------

/** Make sure Excel recalculates on open: calcPr fullCalcOnLoad="1". Returns null if unchanged. */
export function ensureFullCalcOnLoad(wbXml) {
  const m = /<calcPr\b([^>]*?)(\/?)>/.exec(wbXml);
  if (m) {
    if (attr(m[1], 'fullCalcOnLoad') === '1') return null;
    const attrs = /fullCalcOnLoad="[^"]*"/.test(m[1])
      ? m[1].replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"')
      : `${m[1]} fullCalcOnLoad="1"`;
    return wbXml.slice(0, m.index) + `<calcPr${attrs}${m[2]}>` + wbXml.slice(m.index + m[0].length);
  }
  // calcPr comes after definedNames and before these elements (schema order).
  const next = /<(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/.exec(wbXml);
  return wbXml.slice(0, next.index) + '<calcPr fullCalcOnLoad="1"/>' + wbXml.slice(next.index);
}

// ---------- the export stamp (a custom document property) ----------

// Excel shows it under File > Info > Properties > Advanced > Custom and keeps it when saving.
// The OOXML namespace and type identifiers below are fixed names, not addresses the app contacts.
export const STAMP_PROPERTY = 'FinancialFlowExport';
const CUSTOM_PART = 'docProps/custom.xml';
const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/';
const CUSTOM_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

/** The stamp in a workbook (an ISO time), or null. */
export async function readStamp(wb) {
  if (!wb.byName.has(CUSTOM_PART)) return null;
  const xml = await partText(wb, CUSTOM_PART);
  const m = new RegExp(`<property\\b[^>]*\\bname="${STAMP_PROPERTY}"[^>]*>\\s*<vt:lpwstr>([^<]*)</vt:lpwstr>`).exec(xml);
  return m ? xmlUnescape(m[1]) : null;
}

/** Changes/additions that set the stamp, keeping any other custom properties. */
async function stampParts(wb, value, changes, additions, text) {
  const prop = (pid) => `<property fmtid="${FMTID}" pid="${pid}" name="${STAMP_PROPERTY}"><vt:lpwstr>${xmlEscape(value)}</vt:lpwstr></property>`;
  if (wb.byName.has(CUSTOM_PART)) {
    let xml = await partText(wb, CUSTOM_PART);
    const existing = new RegExp(`(<property\\b[^>]*\\bname="${STAMP_PROPERTY}"[^>]*>\\s*<vt:lpwstr>)[^<]*(</vt:lpwstr>)`);
    if (existing.test(xml)) xml = xml.replace(existing, `$1${xmlEscape(value)}$2`);
    else {
      const pids = [...xml.matchAll(/\bpid="(\d+)"/g)].map((m) => Number(m[1]));
      const tag = prop(Math.max(1, ...pids) + 1);
      xml = /<\/Properties>/.test(xml) ? xml.replace('</Properties>', `${tag}</Properties>`) : xml.replace(/<Properties\b([^>]*)\/>/, `<Properties$1>${tag}</Properties>`);
    }
    changes.set(CUSTOM_PART, enc.encode(xml));
    return;
  }
  additions.set(CUSTOM_PART, enc.encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<Properties xmlns="${NS}custom-properties" xmlns:vt="${NS}docPropsVTypes">${prop(2)}</Properties>`));
  text.ct = text.ct.replace('</Types>', `<Override PartName="/${CUSTOM_PART}" ContentType="${CUSTOM_TYPE}"/></Types>`);
  const ids = new Set([...text.rootRels.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]));
  let n = 1;
  while (ids.has(`rIdFF${n}`)) n++;
  text.rootRels = text.rootRels.replace('</Relationships>',
    `<Relationship Id="rIdFF${n}" Type="${NS}relationships/custom-properties" Target="${CUSTOM_PART}"/></Relationships>`);
}

/**
 * Build the exported file.
 * sheetEdits: Map sheetName -> Map ref -> value.
 * stamp: an ISO time to record as the export stamp (optional).
 * Rewrites only the edited sheets; removes calcChain.xml if present (Excel rebuilds it);
 * sets fullCalcOnLoad. All other parts are copied byte-for-byte, except the small
 * custom-properties part (and its two references) when a stamp is written.
 */
export async function exportWorkbook(wb, sheetEdits, { stamp = null } = {}) {
  const changes = new Map();
  const additions = new Map();
  for (const [name, edits] of sheetEdits) {
    if (edits.size === 0) continue;
    const path = wb.sheets.get(name).path;
    const xml = patchSheet(await partText(wb, path), edits);
    changes.set(path, enc.encode(xml));
  }
  const ctBefore = await partText(wb, '[Content_Types].xml');
  const relsBefore = wb.byName.has('_rels/.rels') ? await partText(wb, '_rels/.rels') : null;
  const text = { ct: ctBefore, rootRels: relsBefore };

  if (changes.size > 0) {
    const wbXml = ensureFullCalcOnLoad(await partText(wb, 'xl/workbook.xml'));
    if (wbXml) changes.set('xl/workbook.xml', enc.encode(wbXml));
    if (wb.byName.has('xl/calcChain.xml')) {
      changes.set('xl/calcChain.xml', null);
      text.ct = text.ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
      const rels = await partText(wb, 'xl/_rels/workbook.xml.rels');
      changes.set('xl/_rels/workbook.xml.rels', enc.encode(rels.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, '')));
    }
  }
  if (stamp && relsBefore !== null) await stampParts(wb, stamp, changes, additions, text);
  if (text.ct !== ctBefore) changes.set('[Content_Types].xml', enc.encode(text.ct));
  if (text.rootRels !== relsBefore) changes.set('_rels/.rels', enc.encode(text.rootRels));
  return writeZip(wb.zip, changes, additions);
}
