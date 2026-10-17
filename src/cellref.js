// Cell references: "C7" <-> {col: 3, row: 7}, and shifting the references inside a formula.

export function colToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - 1 - r) / 26;
  }
  return s;
}

export function parseRef(ref) {
  const m = /^([A-Z]{1,3})(\d+)$/.exec(ref);
  if (!m) throw new Error(`Bad cell reference ${ref}`);
  return { col: colToNum(m[1]), row: Number(m[2]) };
}

export function makeRef(col, row) {
  return numToCol(col) + row;
}

// A cell reference, optionally with a sheet prefix. Not preceded by a name character,
// not followed by one or by "(" (so LOG10( or a defined name never matches).
const SHEET = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!`;
const REF_RE = new RegExp(
  String.raw`(?<![A-Za-z0-9_.$'])(${SHEET})?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])`,
  'g',
);
const STRING_RE = /"(?:[^"]|"")*"/g;

/** Split a formula into string-literal and code segments; calls fn on code segments only. */
function mapCode(formula, fn) {
  let out = '';
  let last = 0;
  for (const m of formula.matchAll(STRING_RE)) {
    out += fn(formula.slice(last, m.index)) + m[0];
    last = m.index + m[0].length;
  }
  return out + fn(formula.slice(last));
}

/** List of references in a formula, in order: {sheet, colAbs, col, rowAbs, row}. */
export function formulaRefs(formula) {
  const refs = [];
  mapCode(formula, (code) => {
    for (const m of code.matchAll(REF_RE)) {
      refs.push({ sheet: m[1] || '', colAbs: m[2] === '$', col: colToNum(m[3]), rowAbs: m[4] === '$', row: Number(m[5]) });
    }
    return code;
  });
  return refs;
}

/**
 * Rewrite each reference in a formula. fn(ref, index) returns {col, row} for that reference.
 * The $ markers and sheet prefixes are kept.
 */
export function rewriteRefs(formula, fn) {
  let i = 0;
  return mapCode(formula, (code) =>
    code.replace(REF_RE, (_, sheet = '', cAbs, col, rAbs, row) => {
      const r = { sheet, colAbs: cAbs === '$', col: colToNum(col), rowAbs: rAbs === '$', row: Number(row) };
      const n = fn(r, i++);
      if (n.col < 1 || n.row < 1) throw new Error('Formula reference shifted off the sheet.');
      return sheet + cAbs + numToCol(n.col) + rAbs + n.row;
    }),
  );
}

/** What Excel does when a formula is copied by (dCol, dRow): relative parts move, $ parts stay. */
export function shiftFormula(formula, dCol, dRow) {
  return rewriteRefs(formula, (r) => ({
    col: r.colAbs ? r.col : r.col + dCol,
    row: r.rowAbs ? r.row : r.row + dRow,
  }));
}
