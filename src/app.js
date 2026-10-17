// The screens: Month (monthly entry), Giving (donation log) and File (import / export).
// The DOM is built with createElement only (no HTML strings), so nothing typed can run as code.

import { importWorkbook, WorkbookError } from './import.js';
import { openWorkbook, exportWorkbook } from './xlsx.js';
import { calculate } from './calc.js';
import { todaySerial, monthYear, dollars, YEAR, MONTH, DATE } from './excel.js';
import { dateSerial, serialToIso } from './ffos.js';
import * as entry from './entry.js';
import * as store from './store.js';
import { dashboardView } from './dashboard.js';
import { VERSION } from './version.js';

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const moneyFmt = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
const wholeFmt = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
const centsFmt = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const money = (v) => (typeof v === 'number' && Number.isFinite(v) ? moneyFmt.format(v) : '');
/** 1234 -> "1,234"; 520.4 -> "520.40"; text stays text. */
const plain = (v) =>
  typeof v === 'number' ? (Number.isInteger(v) ? wholeFmt : centsFmt).format(v) : v == null ? '' : String(v);
const dayFmt = new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
const when = (ms) => (ms ? dayFmt.format(new Date(ms)) : 'never');

const state = {
  original: null, base: null, data: null, fileName: null, importedAt: null, lastExport: null,
  exported: null, exportedAt: null, // what this phone last exported, for the older-copy check
  out: null, today: todaySerial(), tab: 'dashboard', month: 0, gift: null, persisted: false, message: null,
  // App updates and install (see setupServiceWorker)
  waiting: null, waitingVersion: null, updating: false, installPrompt: null,
  installed: matchMedia('(display-mode: standalone)').matches,
};

// ---------- DOM helper ----------

function h(tag, props = {}, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v === null || v === undefined || v === false) continue;
    if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'class') el.className = v;
    else if (k === 'value') el.value = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat(Infinity)) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : String(c));
  }
  return el;
}

const toneOf = (status) =>
  !status ? 'tone-muted' : status.startsWith('✓') ? 'tone-good' : status.startsWith('⚠') ? 'tone-bad' : status === 'In progress' ? 'tone-info' : 'tone-muted';

// ---------- state ----------

function recalc() {
  state.today = todaySerial();
  state.out = calculate(state.data, state.today);
}

let saveTimer = null;
function save() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flush, 250);
}
function flush() {
  clearTimeout(saveTimer);
  saveTimer = null;
  if (state.data) return store.setMany({ data: state.data }).catch(showStorageError);
}
function showStorageError(e) {
  state.message = `Couldn't save on this phone (${e?.message || e}). Export soon so nothing is lost.`;
  render();
}
addEventListener('pagehide', flush);
document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && flush());

const changes = () => (state.data ? entry.changeCount(state.base, state.data) : 0);

async function load() {
  const keys = ['original', 'base', 'data', 'fileName', 'importedAt', 'lastExport', 'exported', 'exportedAt'];
  try {
    const values = await Promise.all(keys.map(store.get));
    keys.forEach((k, i) => (state[k] = values[i] ?? null));
  } catch (e) {
    state.message = `This browser won't let the app store data (${e?.message || e}).`;
  }
  if (state.data) {
    recalc();
    state.month = entry.nextMonthToFill(state.out, state.today);
  }
  state.persisted = await store.persisted();
  // Ask again each time: Android grants it once the app is installed or used often.
  if (state.data && !state.persisted) state.persisted = await store.persist();
  render();
}

// ---------- app updates and install ----------

/**
 * The service worker keeps the app's files on the phone. A new version installs in the
 * background and waits; it runs only after the user taps Update (spec: controlled updates).
 */
async function setupServiceWorker() {
  if (!('serviceWorker' in navigator) || !isSecureContext) return;
  let reg;
  try {
    reg = await navigator.serviceWorker.register('sw.js');
  } catch {
    return; // the app still works online without it
  }
  const offer = (worker) => {
    if (!worker || !navigator.serviceWorker.controller) return; // first install: nothing to replace
    state.waiting = worker;
    const channel = new MessageChannel();
    channel.port1.onmessage = (e) => { state.waitingVersion = e.data?.version ?? null; render(); };
    worker.postMessage({ type: 'version' }, [channel.port2]);
    render();
  };
  const watch = (worker) => worker?.addEventListener('statechange', () => worker.state === 'installed' && offer(reg.waiting));
  offer(reg.waiting);
  watch(reg.installing);
  reg.addEventListener('updatefound', () => watch(reg.installing));
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (state.updating) location.reload(); });
  // Look for a new version whenever the app comes back to the front.
  document.addEventListener('visibilitychange', () => document.visibilityState === 'visible' && reg.update().catch(() => {}));
}

async function doUpdate() {
  if (!state.waiting) return;
  await flush();
  state.updating = true;
  state.waiting.postMessage({ type: 'skipWaiting' });
}

addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault(); // show our own "Install app" button instead of the browser's banner
  state.installPrompt = e;
  if (state.tab === 'file') render();
});
addEventListener('appinstalled', async () => {
  state.installPrompt = null;
  state.installed = true;
  state.persisted = await store.persist();
  render();
});

async function doInstall() {
  const p = state.installPrompt;
  if (!p) return;
  state.installPrompt = null;
  await p.prompt();
  render();
}

// ---------- import / export ----------

async function doImport(file) {
  if (!file) return;
  const n = changes();
  if (n > 0 && !confirm(`You have ${n} ${n === 1 ? 'entry' : 'entries'} that ${n === 1 ? "isn't" : "aren't"} exported yet. Importing replaces them. Import anyway?`)) return;
  let data;
  const bytes = new Uint8Array(await file.arrayBuffer());
  try {
    data = await importWorkbook(bytes);
  } catch (e) {
    state.message = e instanceof WorkbookError ? e.message : `Couldn't read that file (${e?.message || e}).`;
    return render();
  }
  // Is this an older copy than what the phone last sent? (cross-checks the stamp inside the file)
  const older = entry.olderCopyCheck({ exported: state.exported, exportedAt: state.exportedAt }, data);
  if (older.warn) {
    const from = older.stamp ? `comes from the export of ${stampText(older.stamp)}` : "wasn't exported by this phone";
    const where = [older.months.length ? older.months.map(ymLabel).join(', ') : null, older.gifts ? 'Giving' : null].filter(Boolean).join('; ');
    if (!confirm(`This file (${file.name}) ${from}, but this phone last exported on ${stampText(state.exportedAt)}. ` +
      `It differs from that export in ${older.missing} ${older.missing === 1 ? 'entry' : 'entries'}${where ? ` (${where})` : ''}. ` +
      'It may be an older copy. Import anyway?')) return;
  }
  Object.assign(state, {
    original: bytes, base: data, data: structuredClone(data), fileName: file.name, importedAt: Date.now(), message: null,
  });
  await store.setMany({ original: state.original, base: state.base, data: state.data, fileName: state.fileName, importedAt: state.importedAt })
    .catch(showStorageError);
  state.persisted = await store.persist();
  recalc();
  state.month = entry.nextMonthToFill(state.out, state.today);
  state.tab = 'month';
  render();
}

async function doExport() {
  await flush();
  const wb = await openWorkbook(state.original);
  const now = new Date();
  const stamp = now.toISOString();
  const bytes = await exportWorkbook(wb, entry.workbookEdits(state.base, state.data), { stamp });
  const file = new File([bytes], entry.exportFileName(state.fileName, now), { type: XLSX_TYPE });
  let sent = false;
  if (navigator.canShare?.({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: file.name });
      sent = true;
    } catch (e) {
      if (e?.name === 'AbortError') return; // the user closed the share sheet
    }
  }
  if (!sent) {
    const url = URL.createObjectURL(file);
    h('a', { href: url, download: file.name }).click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  const exportedData = structuredClone({ ...state.data, exportStamp: stamp });
  Object.assign(state, {
    original: bytes, base: exportedData, data: structuredClone(exportedData), exported: structuredClone(exportedData), exportedAt: stamp,
    fileName: file.name, lastExport: now.getTime(), message: `Exported ${file.name}.`,
  });
  await store.setMany({
    original: state.original, base: state.base, data: state.data, exported: state.exported,
    exportedAt: state.exportedAt, fileName: state.fileName, lastExport: state.lastExport,
  }).catch(showStorageError);
  render();
}

/** "Oct 17, 2026, 2:32 PM" from an ISO time. */
function stampText(iso) {
  return iso ? new Date(iso).toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }) : '—';
}
const ymLabel = (ym) => { const [y, m] = ym.split('-').map(Number); return monthYear(DATE(y, m, 1)); };

async function doForget() {
  if (!confirm('Remove the workbook and all entries from this phone? Export first if you want to keep them.')) return;
  await store.clearAll();
  Object.assign(state, { original: null, base: null, data: null, fileName: null, importedAt: null, lastExport: null, exported: null, exportedAt: null, out: null, message: null });
  render();
}

// ---------- render ----------

function render() {
  for (const b of document.querySelectorAll('.tabs button')) {
    if (b.dataset.tab === state.tab) b.setAttribute('aria-current', 'page');
    else b.removeAttribute('aria-current');
  }
  const main = document.getElementById('main');
  const screen = !state.data && state.tab !== 'file' ? noDataScreen()
    : state.tab === 'dashboard' ? dashboardScreen()
    : state.tab === 'giving' ? givingScreen()
    : state.tab === 'file' ? fileScreen()
    : monthScreen();
  const banners = [];
  if (state.waiting) {
    banners.push(h('div', { class: 'banner tone-info', role: 'status' },
      h('span', {}, state.waitingVersion ? `Version ${state.waitingVersion} is ready.` : 'A new version is ready.'),
      h('button', { type: 'button', class: 'btn primary', onclick: doUpdate }, 'Update')));
  }
  if (state.message) banners.push(h('p', { class: 'status tone-info', role: 'status' }, state.message));
  main.replaceChildren(...banners, screen);
  updateChip();
}

function reminder() {
  return entry.exportReminder({ lastExport: state.lastExport, importedAt: state.importedAt, changes: changes() });
}

function updateChip() {
  const chip = document.getElementById('changes-chip');
  const n = changes();
  chip.hidden = n === 0;
  chip.textContent = `${n} not exported`;
}

function importButton(label, primary = false) {
  return h('label', { class: `btn file-label${primary ? ' primary' : ''}` }, label,
    h('input', { type: 'file', accept: `.xlsx,${XLSX_TYPE}`, onchange: (e) => doImport(e.target.files[0]) }));
}

function noDataScreen() {
  return h('section', { class: 'card' },
    h('h2', {}, 'Start with your workbook'),
    h('p', { class: 'help' }, 'Import your FFOS workbook (.xlsx). Its settings, months and giving log come into the app. Nothing leaves this phone.'),
    h('div', { class: 'btn-row' }, importButton('Import workbook', true)));
}

// ----- Month -----

let fieldRefs = [];

function monthScreen() {
  const out = state.out;
  const now = entry.currentMonthIndex(out.months, state.today);
  const i = Math.min(state.month, now);
  const view = entry.monthView(state.data, out, i);
  fieldRefs = [];
  const go = (j) => { state.month = j; render(); };

  const select = h('select', { 'aria-label': 'Month', id: 'month-select', onchange: (e) => go(Number(e.target.value)) },
    out.months.slice(0, now + 1).map((r, j) => h('option', { value: j, selected: j === i }, monthLabel(r))).reverse());

  const sections = [
    h('div', { class: 'switcher' },
      h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Previous month', disabled: i === 0, onclick: () => go(i - 1) }, '‹'),
      select,
      h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Next month', disabled: i >= now, onclick: () => go(i + 1) }, '›')),
    h('p', { class: `status ${toneOf(view.status)}`, id: 'month-status', role: 'status' }, view.status || ''),
  ];
  if (view.phase === 'pre') {
    sections.push(h('section', { class: 'card' }, h('p', { class: 'help' }, 'Before the opening date, balances aren\'t used. Income and giving still count toward tithe.')));
  } else {
    sections.push(card('Balances', 'Month-end balances. Cards and loans: the amount owed.',
      view.accounts.map((f) => field(i, f)), h('span', { class: 'count', id: 'missing-count' })));
  }
  sections.push(card('Income', view.income.length ? 'Grey = the automatic amount from Settings. Type to change it.' : 'No income sources are set up in Settings.',
    view.income.map((f) => field(i, f))));
  sections.push(card('Costs', null, [field(i, view.family), field(i, view.oneTime)]));
  if (view.investments.length) {
    sections.push(card('Investments', 'Value at month end, and money moved in (+) or out (−).',
      view.investments.flatMap((f) => [field(i, f), field(i, f.deposit, true)])));
  }
  sections.push(h('section', { class: 'card' },
    h('h2', {}, h('label', { for: 'note' }, 'Note')),
    h('textarea', { id: 'note', maxlength: 500, onchange: (e) => { entry.setNote(state.data, i, e.target.value); afterChange(i); } }, view.note ?? '')));
  sections.push(h('section', { class: 'card' }, h('h2', {}, 'This month'), h('dl', { class: 'results', id: 'results' })));

  const frag = h('div', {}, sections);
  queueMicrotask(() => refreshMonth(i));
  return frag;
}

function monthLabel(r) {
  const sym = r.Status?.startsWith('✓') ? ' ✓' : r.Status?.startsWith('⚠') ? ' ⚠' : '';
  return `${monthYear(r.MonthStart)}${sym}`;
}

function card(title, help, fields, extra = null) {
  return h('section', { class: 'card' }, h('h2', {}, title, extra), help ? h('p', { class: 'help' }, help) : null, fields);
}

function field(i, f, sub = false) {
  const id = `f-${f.key}-${f.slot ?? 0}`;
  const isFormula = f.cell.kind === 'formula';
  const input = h('input', {
    id, type: 'text', inputmode: 'decimal', enterkeyhint: 'next', autocomplete: 'off',
    readonly: isFormula, value: f.cell.kind === 'auto' ? '' : plain(f.cell.value),
  });
  const clear = h('button', { type: 'button', class: 'clear', 'aria-label': `Use the automatic value for ${f.label}`, title: 'Use automatic' }, '×');
  const error = h('span', { class: 'error', role: 'alert', hidden: true });
  const hint = isFormula ? 'formula set in Excel; change it there' : f.hint;
  const el = h('div', { class: `field${sub ? ' sub' : ''}` },
    h('label', { for: id }, f.label, f.required ? h('span', { class: 'req' }, 'required') : null, hint ? h('span', { class: 'hint' }, hint) : null),
    h('div', { class: 'input-wrap' }, input, isFormula ? null : clear),
    error);
  const commit = (value) => {
    if (!entry.setEntry(state.data, i, f.key, f.slot, value)) return;
    input.value = value === null ? '' : plain(value);
    afterChange(i);
  };
  input.addEventListener('change', () => {
    const p = entry.parseAmount(input.value);
    input.setAttribute('aria-invalid', p.error ? 'true' : 'false');
    error.hidden = !p.error;
    error.textContent = p.error || '';
    if (!p.error) commit(p.value);
  });
  input.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const all = [...document.querySelectorAll('main input:not([readonly]), main textarea')];
    const next = all[all.indexOf(input) + 1];
    if (next) next.focus();
    else input.blur();
  });
  clear.addEventListener('click', () => { error.hidden = true; input.removeAttribute('aria-invalid'); commit(null); });
  fieldRefs.push({ f, input, clear });
  return el;
}

function afterChange(i) {
  recalc();
  save();
  refreshMonth(i);
}

/** Update what depends on the numbers without rebuilding the inputs (keeps focus and keyboard). */
function refreshMonth(i) {
  const view = entry.monthView(state.data, state.out, i);
  const all = [...view.accounts, ...view.income, view.family, view.oneTime, ...view.investments.flatMap((f) => [f, f.deposit])];
  for (const ref of fieldRefs) {
    const f = all.find((x) => x.key === ref.f.key && x.slot === ref.f.slot);
    if (!f) continue;
    ref.input.placeholder = f.auto !== null ? plain(f.auto) : f.required ? 'required' : '';
    ref.clear.hidden = f.cell.kind !== 'typed';
  }
  // A change can complete this month or the next one, so relabel every month in the picker.
  for (const opt of document.getElementById('month-select')?.options ?? []) {
    opt.textContent = monthLabel(state.out.months[Number(opt.value)]);
  }
  const status = document.getElementById('month-status');
  if (status) {
    status.textContent = view.status || '';
    status.className = `status ${toneOf(view.status)}`;
  }
  const missing = document.getElementById('missing-count');
  if (missing) {
    missing.textContent = view.missing ? `${view.missing} required missing` : view.accounts.length ? '✓ all required in' : '';
    missing.className = `count ${view.missing ? 'tone-bad' : 'tone-good'}`;
  }
  const res = document.getElementById('results');
  if (res) {
    const r = view.results;
    const rows = [
      ['Income', r.income], ['Household-paid spending', r.household], ['Family-paid', r.family],
      ['True cost of living', r.trueCost], ['Given', r.given], ['Saved (after giving)', r.saved], ['Net worth', r.netWorth],
    ];
    res.replaceChildren(...rows.flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, v === null ? '—' : money(v))]));
  }
  updateChip();
}

// ----- Dashboard -----

const whole = (v) => (typeof v === 'number' && Number.isFinite(v) ? dollars(v) : '—');
const lineTone = (t) => (t.startsWith('✓') ? 'tone-good' : t.startsWith('⚠') ? 'tone-bad' : 'tone-info');

function bar(fraction, tone = '') {
  const fill = h('span', { class: `bar-fill ${tone}` });
  fill.style.width = `${Math.round(Math.max(0, Math.min(1, fraction || 0)) * 100)}%`; // CSSOM, allowed by the CSP
  return h('span', { class: 'bar', 'aria-hidden': 'true' }, fill);
}

function goalLine(g) {
  const parts = String(g.type).toLowerCase() === 'pay down to'
    ? [`${whole(g.current)} now`, `target ${whole(g.target)} or less`, g.remaining > 0 ? `${whole(g.remaining)} over` : null]
    : [`${whole(g.current)} of ${whole(g.target)}`, g.remaining > 0 ? `${whole(g.remaining)} to go` : null];
  return [...parts, g.note].filter(Boolean).join(' · ');
}

function dashboardScreen() {
  const v = dashboardView(state.out, state.data, state.today);
  const t = v.tiles;
  const tile = (label, x, extraClass = '') =>
    h('div', { class: `tile ${extraClass}` }, h('span', { class: 'tile-label' }, label), h('span', { class: 'tile-value' }, whole(x.value)),
      x.line ? h('span', { class: `tile-line ${x.tone ? `tone-text-${x.tone}` : ''}` }, x.line) : null);
  const goTo = (i) => { state.month = i; state.tab = 'month'; render(); scrollTo(0, 0); };

  const head = h('section', { class: 'card as-of' },
    h('p', {}, `As of ${v.asOfText}`),
    v.next.index >= 0 && v.next.index <= entry.currentMonthIndex(state.out.months, state.today)
      ? h('button', { type: 'button', class: 'btn primary', onclick: () => goTo(v.next.index) }, `Fill ${v.next.label} →`)
      : null);

  const rem = reminder();
  const remind = rem.due ? h('div', { class: 'banner tone-bad', role: 'status' },
    h('span', {}, `⚠ ${rem.text}. Export so Excel has your ${changes()} new ${changes() === 1 ? 'entry' : 'entries'}.`),
    h('button', { type: 'button', class: 'btn', onclick: () => { state.tab = 'file'; render(); } }, 'Export')) : null;

  const tiles = h('section', { class: 'tiles' },
    tile('Net worth', t.netWorth, 'wide'),
    tile('Liquid cash', t.liquid),
    tile('Investments + retirement', t.investments),
    tile('Owed on cards & loans', t.owed, 'span'));

  const [hLatest, hAvg, hYear] = v.flow.headers;
  const flow = h('section', { class: 'card' },
    h('h2', {}, 'Money flow'),
    h('table', { class: 'flow' },
      h('thead', {}, h('tr', {}, h('th', { scope: 'col' }, h('span', { class: 'sr-only' }, 'Item')),
        h('th', { scope: 'col', title: hLatest }, hLatest.replace(/^Latest month \((.*)\)$/, '$1')),
        h('th', { scope: 'col', title: hAvg }, hAvg.replace(/^Average \(last (.*) full\)$/, 'Avg of $1')),
        h('th', { scope: 'col', title: hYear }, hYear.replace(/^This year \((.*)\)$/, '$1')))),
      h('tbody', {}, v.flow.rows.map(([label, a, b, c]) => h('tr', { class: label.startsWith(' ') ? 'sub-row' : '' },
        h('th', { scope: 'row' }, label.trim()), h('td', {}, a === null ? '' : whole(a)), h('td', {}, b === null ? '' : whole(b)), h('td', {}, c === null ? '' : whole(c)))))));

  const m = v.tithe;
  const tithe = h('section', { class: 'card' },
    h('h2', {}, 'Tithe balance', h('span', { class: `count ${m.balance > 0 ? 'tone-bad' : 'tone-good'}` }, m.balance > 0 ? 'owed' : '✓ settled')),
    h('p', { class: 'big' }, whole(m.balance)),
    h('p', { class: 'help' }, m.line),
    bar(m.owed > 0 ? m.given / m.owed : 1, 'good'));

  const tax = v.tax ? h('section', { class: 'card' },
    h('h2', {}, 'Tax this year'),
    h('dl', { class: 'results' },
      h('dt', {}, 'Taxable income received'), h('dd', {}, whole(v.tax.received)),
      h('dt', { class: 'sub-dt' }, '…of which no tax withheld'), h('dd', {}, whole(v.tax.unwithheld)),
      h('dt', {}, 'Tax-deductible giving'), h('dd', {}, whole(v.tax.deductible))),
    h('p', { class: 'help' }, 'Amounts received: for pay with tax withheld this is take-home pay, not the gross on a W-2. Marks are set in Settings (income sources) and on each donation.')) : null;

  const goals = v.goals.length ? h('section', { class: 'card' },
    h('h2', {}, 'Goals'),
    v.goals.map((g) => {
      const tone = g.status === 'Met' ? 'good' : g.status === 'Above target' ? 'bad' : 'info';
      const sym = g.status === 'Met' ? '✓ ' : g.status === 'Above target' ? '⚠ ' : '';
      return h('div', { class: 'goal' },
        h('div', { class: 'goal-head' }, h('span', { class: 'goal-name' }, String(g.name)), h('span', { class: `count tone-text-${tone}` }, `${sym}${g.status}`)),
        bar(g.progress, tone),
        h('p', { class: 'help' }, goalLine(g)));
    })) : null;

  const accounts = h('section', { class: 'card' },
    h('h2', {}, 'Accounts'),
    v.accounts.length === 0 ? h('p', { class: 'help' }, 'No active accounts in Settings.') : null,
    h('dl', { class: 'results accounts' }, v.accounts.flatMap((a) => [
      h('dt', {}, String(a.name), h('span', { class: 'hint' }, [a.type, a.estimate].filter(Boolean).join(' · '))),
      h('dd', { class: a.signed < 0 ? 'tone-text-bad' : '' }, whole(a.signed)),
    ])));

  const health = h('section', { class: 'card' },
    h('h2', {}, 'Health checks'),
    h('ul', { class: 'checks' }, v.health.map((x) => h('li', { class: lineTone(x) }, x))));

  return h('div', {}, remind, head, tiles, flow, tithe, tax, goals, accounts, health);
}

// ----- Giving -----

function givingScreen() {
  const data = state.data;
  const editing = state.gift !== null ? data.giving.find((g) => g.row === state.gift) : null;
  const g = editing ?? { date: state.today, amount: null, recipient: null, paidFrom: null, household: null, note: null };
  const recipients = [...new Set(data.giving.map((x) => x.recipient).filter(Boolean))].sort();
  const payers = [...new Set([...data.accounts.map((a) => a.name), ...data.giving.map((x) => x.paidFrom)].filter(Boolean))].sort();

  const date = h('input', { type: 'date', id: 'g-date', required: true, value: typeof g.date === 'number' ? serialToIso(g.date) : '' });
  const amount = h('input', { type: 'text', id: 'g-amount', inputmode: 'decimal', autocomplete: 'off', value: plain(g.amount) });
  const recipient = h('input', { type: 'text', id: 'g-recipient', list: 'recipients', autocomplete: 'off', maxlength: 200, value: g.recipient ?? '' });
  const paidFrom = h('input', { type: 'text', id: 'g-paid', list: 'payers', autocomplete: 'off', maxlength: 200, value: g.paidFrom ?? '' });
  const knownHousehold = g.household === null || g.household === 'No';
  const household = h('select', { id: 'g-household' },
    h('option', { value: '', selected: g.household === null }, 'Household money (counts toward tithe)'),
    h('option', { value: 'No', selected: g.household === 'No' }, 'Family money (not your tithe)'),
    knownHousehold ? null : h('option', { value: g.household, selected: true }, String(g.household)));
  const note = h('input', { type: 'text', id: 'g-note', autocomplete: 'off', maxlength: 500, value: g.note ?? '' });
  // Tax option (FFOS 3.8): unchecked keeps an explicit "No" from Excel, otherwise leaves it blank (= No).
  const deductible = data.hasTax ? h('input', { type: 'checkbox', id: 'g-deductible', checked: String(g.deductible ?? '').toLowerCase() === 'yes' }) : null;
  const error = h('p', { class: 'status tone-bad', role: 'alert', hidden: true });

  const saveGift = () => {
    const amt = entry.parseAmount(amount.value);
    const problem = !date.value ? 'Pick a date.' : amt.error ? amt.error : amt.value === null ? 'Enter the amount.' : null;
    if (problem) { error.textContent = problem; error.hidden = false; return; }
    const gift = { date: dateSerial(date.value), amount: amt.value, recipient: recipient.value, paidFrom: paidFrom.value, household: household.value || null, note: note.value };
    if (deductible) gift.deductible = deductible.checked ? 'Yes' : String(g.deductible ?? '').toLowerCase() === 'no' ? g.deductible : null;
    if (editing) entry.updateGift(data, editing.row, gift);
    else if (entry.addGift(data, gift) === null) { error.textContent = 'The Giving log is full (2,000 rows).'; error.hidden = false; return; }
    state.gift = null;
    recalc(); save(); render();
  };

  const m = state.out.dashboard.tithe;
  const form = h('section', { class: 'card' },
    h('h2', {}, editing ? 'Edit donation' : 'Add a donation'),
    h('p', { class: 'help' }, `This quarter: ${money(m.givenThisQuarter)} given of ${money(m.owedThisQuarter)} owed.`),
    h('div', { class: 'form-grid' },
      h('label', { for: 'g-date' }, 'Date', date),
      h('label', { for: 'g-amount' }, 'Amount', amount),
      h('label', { for: 'g-recipient' }, 'Recipient', recipient),
      h('label', { for: 'g-paid' }, 'Paid from (optional)', paidFrom),
      h('label', { for: 'g-household' }, 'Whose money', household),
      h('label', { for: 'g-note' }, 'Note (optional)', note),
      deductible ? h('label', { for: 'g-deductible', class: 'check' }, deductible, 'Tax-deductible') : null),
    h('datalist', { id: 'recipients' }, recipients.map((r) => h('option', { value: r }))),
    h('datalist', { id: 'payers' }, payers.map((p) => h('option', { value: p }))),
    error,
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn primary', onclick: saveGift }, editing ? 'Save' : 'Add'),
      editing ? h('button', { type: 'button', class: 'btn', onclick: () => { state.gift = null; render(); } }, 'Cancel') : null,
      editing ? h('button', { type: 'button', class: 'btn danger', onclick: () => {
        if (!confirm('Delete this donation?')) return;
        entry.deleteGift(data, editing.row); state.gift = null; recalc(); save(); render();
      } }, 'Delete') : null));

  // The log, newest first, grouped by month.
  const sorted = [...data.giving].sort((a, b) => (typeof b.date === 'number' ? b.date : -1) - (typeof a.date === 'number' ? a.date : -1) || b.row - a.row);
  const groups = new Map();
  for (const x of sorted) {
    const key = typeof x.date === 'number' ? `${YEAR(x.date)}-${MONTH(x.date)}` : 'none';
    if (!groups.has(key)) groups.set(key, { label: typeof x.date === 'number' ? monthYear(x.date) : 'No date', total: 0, rows: [] });
    const grp = groups.get(key);
    grp.rows.push(x);
    if (typeof x.amount === 'number') grp.total += x.amount;
  }
  const list = h('section', {}, groups.size === 0 ? h('p', { class: 'notice' }, 'No donations in the log yet.') : null,
    [...groups.values()].map((grp) => [
      h('div', { class: 'gift-month' }, h('span', {}, grp.label), h('span', {}, money(grp.total))),
      grp.rows.map((x) => h('button', { type: 'button', class: 'gift', onclick: () => { state.gift = x.row; render(); scrollTo(0, 0); } },
        h('span', {}, x.recipient ?? '—',
          h('small', {}, [typeof x.date === 'number' ? dayFmt.format(new Date(`${serialToIso(x.date)}T12:00:00`)) : 'no date', x.paidFrom, x.household === 'No' ? 'family money' : null, String(x.deductible ?? '').toLowerCase() === 'yes' ? 'tax-deductible' : null].filter(Boolean).join(' · '))),
        h('span', { class: 'amt' }, typeof x.amount === 'number' ? money(x.amount) : 'no amount'))),
    ]));
  return h('div', {}, form, list);
}

// ----- File -----

function fileScreen() {
  const n = changes();
  const facts = state.data
    ? h('dl', { class: 'facts' },
      h('dt', {}, 'Workbook'), h('dd', {}, state.fileName || '—'),
      h('dt', {}, 'Imported'), h('dd', {}, when(state.importedAt)),
      h('dt', {}, 'Stamp in file'), h('dd', {}, state.base?.exportStamp ? `exported ${stampText(state.base.exportStamp)}` : 'none (not exported by this app yet)'),
      h('dt', {}, 'Not exported'), h('dd', {}, n === 0 ? '✓ nothing' : `${n} ${n === 1 ? 'entry' : 'entries'}`),
      h('dt', {}, 'Last export'), h('dd', { class: reminder().due ? 'tone-text-bad' : '' }, state.lastExport ? `${when(state.lastExport)} (${reminder().text.replace(/^Last export: /, '')})` : 'never'))
    : h('p', { class: 'help' }, 'No workbook on this phone yet.');
  const app = h('section', { class: 'card' },
    h('h2', {}, 'This app'),
    h('dl', { class: 'facts' },
      h('dt', {}, 'Version'), h('dd', {}, state.waiting ? `${VERSION} · ${state.waitingVersion ?? 'a new version'} is ready` : `${VERSION}`),
      h('dt', {}, 'Offline'), h('dd', {}, navigator.serviceWorker?.controller ? '✓ works without internet' : 'available after the next start'),
      h('dt', {}, 'Storage'), h('dd', {}, state.persisted ? '✓ kept safe on this phone' : 'on this phone; Android may clear it when space runs low')),
    !state.persisted && !state.installed ? h('p', { class: 'help' }, 'Install the app to your home screen; Android then keeps its data safe.') : null,
    h('div', { class: 'btn-row' },
      state.waiting ? h('button', { type: 'button', class: 'btn primary', onclick: doUpdate }, 'Update') : null,
      state.installPrompt ? h('button', { type: 'button', class: 'btn', onclick: doInstall }, 'Install app') : null));
  return h('div', {},
    h('section', { class: 'card' },
      h('h2', {}, 'Your workbook'),
      facts,
      h('div', { class: 'btn-row' },
        state.data ? h('button', { type: 'button', class: 'btn primary', onclick: doExport }, 'Export workbook') : null,
        importButton(state.data ? 'Import a newer workbook' : 'Import workbook', !state.data))),
    app,
    h('section', { class: 'card' },
      h('h2', {}, 'How it works'),
      h('p', { class: 'help' }, 'Excel stays the record. Export sends the workbook with your entries written in; choose Quick Share to send it to your PC.'),
      h('p', { class: 'help' }, 'After you change the workbook in Excel, import it here before entering anything on the phone.')),
    state.data ? h('section', { class: 'card' },
      h('h2', {}, 'Remove from this phone'),
      h('p', { class: 'help' }, 'Deletes the workbook and all entries from this phone.'),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn danger', onclick: doForget }, 'Remove'))) : null);
}

// ---------- start ----------

for (const b of document.querySelectorAll('.tabs button')) {
  b.addEventListener('click', () => { state.tab = b.dataset.tab; state.gift = null; state.message = null; render(); scrollTo(0, 0); });
}
document.getElementById('changes-chip').addEventListener('click', () => { state.tab = 'file'; render(); });
load();
setupServiceWorker();
