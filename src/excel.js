// The few Excel behaviours the workbook's formulas rely on, so the app computes what Excel computes.
// Values: number | string | boolean | null (blank). Dates are Excel serial numbers (1900 system).

export const isBlank = (v) => v === null || v === undefined || v === '';

/** N(): numbers stay, TRUE is 1, everything else (text, blank) is 0. */
export const N = (v) => (typeof v === 'number' ? v : v === true ? 1 : 0);

/** A cell used in arithmetic (=A1*2): blank is 0, numeric text converts, other text is an error (NaN). */
export function num(v) {
  if (typeof v === 'number') return v;
  if (isBlank(v)) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const t = String(v).trim();
  return t !== '' && Number.isFinite(Number(t)) ? Number(t) : NaN;
}

/** Excel's  cell = "text"  (case-insensitive; a blank cell equals ""). */
export function is(v, text) {
  if (typeof v === 'string') return v.toLowerCase() === text.toLowerCase();
  return text === '' && isBlank(v);
}

/** A value on the right of  number <= cell : blank is 0, text sorts above every number. */
export const cmpValue = (v) => (typeof v === 'number' ? v : isBlank(v) ? 0 : typeof v === 'string' ? Infinity : N(v));

/** MAX/MIN over cell references: only numbers count (text, blanks and TRUE/FALSE are skipped). */
export function maxNums(...vs) {
  const n = vs.filter((v) => typeof v === 'number');
  return n.length ? Math.max(...n) : 0;
}
export function minNums(...vs) {
  const n = vs.filter((v) => typeof v === 'number');
  return n.length ? Math.min(...n) : 0;
}

/** ROUND(x, digits): half away from zero, on Excel's 15 significant digits. */
export function round(x, digits = 0) {
  if (!Number.isFinite(x)) return x;
  const f = 10 ** digits;
  const m = Number((Math.abs(x) * f).toPrecision(15)); // 1.005*100 = 100.49999… -> 100.5
  return (Math.sign(x) * Math.round(m)) / f || 0;
}

/** Excel MOD: the result takes the divisor's sign. */
export const mod = (a, b) => ((a % b) + b) % b;

// ---------- dates ----------

const DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
const d = (s) => new Date(EPOCH + Math.floor(s) * DAY);

export const YEAR = (s) => d(s).getUTCFullYear();
export const MONTH = (s) => d(s).getUTCMonth() + 1;
export const DAYOF = (s) => d(s).getUTCDate();
export const DATE = (y, m, day) => (Date.UTC(Math.trunc(y), Math.trunc(m) - 1, Math.trunc(day)) - EPOCH) / DAY;
export const EOMONTH = (s, k) => DATE(YEAR(s), MONTH(s) + k + 1, 0);
export function EDATE(s, k) {
  const last = EOMONTH(s, k);
  return DATE(YEAR(last), MONTH(last), Math.min(DAYOF(s), DAYOF(last)));
}
/** WEEKDAY(serial): Sunday = 1 ... Saturday = 7. */
export const WEEKDAY = (s) => mod(Math.floor(s) - 1, 7) + 1;
export const quarterStart = (s) => DATE(YEAR(s), Math.floor((MONTH(s) - 1) / 3) * 3 + 1, 1);

export function todaySerial(now = new Date()) {
  return (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - EPOCH) / DAY;
}

// ---------- TEXT() for the formats the workbook uses ----------

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/** TEXT(x, "$#,##0") */
export function dollars(x) {
  const r = round(x, 0);
  const s = Math.abs(r).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (r < 0 ? '-$' : '$') + s;
}
/** TEXT(serial, "mmm yyyy") / "mmmm yyyy" / "m/d/yy" / "mmm yy" */
export const monYear = (s) => `${MONTHS[MONTH(s) - 1].slice(0, 3)} ${YEAR(s)}`;
export const monthYear = (s) => `${MONTHS[MONTH(s) - 1]} ${YEAR(s)}`;
export const mdyy = (s) => `${MONTH(s)}/${DAYOF(s)}/${String(YEAR(s) % 100).padStart(2, '0')}`;
export const monYY = (s) => `${MONTHS[MONTH(s) - 1].slice(0, 3)} ${String(YEAR(s) % 100).padStart(2, '0')}`;
