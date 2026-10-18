// The screens: four tabs (Dashboard, Ledger, Giving, Goals) and two pages reached by tapping what they're about
// (Workbook from the chip in the top bar, Health checks from the Dashboard's warning line).
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
  out: null, today: todaySerial(), tab: 'dashboard' /* dashboard, ledger, giving or goals */, gift: null, goal: null /* null, 'new' or a slot */, entry: null /* the month open in the Ledger, or null for its list */,
  page: null /* 'workbook' or 'checks', open over the current tab */, persisted: false, message: null,
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
  if (state.data) recalc();
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
  if (state.page === 'workbook') render();
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
    const where = [older.months.length ? older.months.map(ymLabel).join(', ') : null, older.gifts ? 'Giving' : null, older.goals ? 'Goals' : null, older.setup ? 'Accounts and income' : null].filter(Boolean).join('; ');
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
  navigate({ tab: 'ledger', entry: entry.nextMonthToFill(state.out, state.today) });
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
  const screen = state.page ? pageScreen()
    : !state.data ? noDataScreen()
    : state.tab === 'dashboard' ? dashboardScreen()
    : state.tab === 'giving' ? givingScreen()
    : state.tab === 'goals' ? goalsScreen()
    : ledgerScreen();
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

/** The chip in the top bar: always there, and the way to the Workbook page (export, import, the app's version). */
function updateChip() {
  const chip = document.getElementById('changes-chip');
  const n = changes();
  chip.textContent = !state.data ? 'Workbook' : n ? `${n} not exported` : '✓ Up to date';
  chip.classList.toggle('quiet', !!state.data && n === 0);
  if (state.page === 'workbook') chip.setAttribute('aria-current', 'page');
  else chip.removeAttribute('aria-current');
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

// ----- Ledger: every month (newest first) and this year's totals; tap a month to fill or fix it -----

/** Open month i's entry screen (the phone's Back button returns to where you were). */
const openEntry = (i) => navigate({ tab: 'ledger', entry: i });

function ledgerScreen() {
  return state.entry === null ? ledgerList() : monthScreen(state.entry);
}

function ledgerList() {
  const out = state.out;
  const v = dashboardView(out, state.data, state.today);
  const now = entry.currentMonthIndex(out.months, state.today);
  const next = entry.nextMonthToFill(out, state.today);
  const nr = out.months[next];
  const fill = h('section', { class: 'card next-fill', id: 'next-fill' },
    h('div', {},
      h('span', { class: 'line-title' }, `Next: ${monthYear(nr.MonthStart)}`),
      nr.Status ? h('span', { class: `line-sub ${toneOf(nr.Status).replace('tone-', 'tone-text-')}` }, nr.Status) : null),
    h('button', { type: 'button', class: 'btn primary', onclick: () => openEntry(next) }, 'Fill →'));

  const [, hAvg, hYear] = v.flow.headers;
  const rows = (col) => v.flow.rows
    .filter(([, ...vals]) => vals[col] !== null)
    .flatMap(([label, ...vals]) => [h('dt', { class: label.startsWith(' ') ? 'sub-dt' : '' }, label.trim()), h('dd', {}, whole(vals[col]))]);
  const year = h('section', { class: 'card', id: 'h-year' },
    h('h2', {}, hYear.replace(/^This year \((.*)\)$/, 'This year ($1)')),
    h('dl', { class: 'results' }, rows(2),
      v.tax ? [h('dt', {}, 'Taxable income received'), h('dd', {}, whole(v.tax.received)),
        h('dt', { class: 'sub-dt' }, '…of which no tax withheld'), h('dd', {}, whole(v.tax.unwithheld))] : null));
  const avg = h('section', { class: 'card', id: 'h-avg' },
    h('h2', {}, hAvg.replace(/^Average \(last (.*) full\)$/, 'Average month (last $1 full)')),
    h('dl', { class: 'results' }, rows(1)));

  // Every month up to this one, newest first (months before the opening date are history only).
  const months = out.months.slice(0, now + 1).map((r, i) => ({ r, i })).reverse();
  const list = h('section', { class: 'card', id: 'h-months' },
    h('h2', {}, 'Month by month'),
    months.map(({ r, i }) => h('button', { type: 'button', class: 'month-row', onclick: () => openEntry(i) },
      h('span', {}, h('span', { class: 'month-name' }, monthLabel(r)),
        h('small', {}, [typeof r.TrueCost === 'number' ? `spent ${whole(r.TrueCost)}` : null, typeof r.Saved === 'number' ? `saved ${whole(r.Saved)}` : null].filter(Boolean).join(' · ') || r.Status || '')),
      h('span', { class: 'amt' }, typeof r.NetWorth === 'number' ? whole(r.NetWorth) : '—'))));
  return h('div', {}, fill, year, avg, list);
}

// ----- A month's entry screen (with adding and closing accounts, investments and income sources) -----

let fieldRefs = [];

function monthScreen(i) {
  const out = state.out;
  const view = entry.monthView(state.data, out, i);
  const ml = monthYear(out.months[i].MonthStart);
  fieldRefs = [];

  const sections = [
    h('div', { class: 'sub-head' },
      h('button', { type: 'button', class: 'back-btn', onclick: goBack }, '‹ Back'),
      h('h2', { id: 'month-title' }, ml)),
    h('p', { class: `status ${toneOf(view.status)}`, id: 'month-status', role: 'status' }, view.status || ''),
  ];
  const manage = (table, f, extra) => ({ table, slot: f.slot, name: f.label, month: ml, i, ...extra });
  if (view.phase === 'pre') {
    sections.push(h('section', { class: 'card' }, h('p', { class: 'help' }, 'Before the opening date, balances aren\'t used. Income and giving still count toward tithe.')));
  } else {
    sections.push(card('Balances', 'Month-end balances. Cards and loans: the amount owed.',
      [view.accounts.map((f) => field(i, f, false, manage('accounts', f, { retired: f.retired }))),
        addForm(i, 'accounts')],
      h('span', { class: 'count', id: 'missing-count' })));
  }
  sections.push(card('Income', view.income.length ? 'Grey = the automatic amount from Settings. Type to change it.' : 'No income sources yet.',
    [view.income.map((f) => field(i, f, false, manage('sources', f, { ended: f.ended }))), addForm(i, 'sources')]));
  sections.push(card('Costs', null, [field(i, view.family), field(i, view.oneTime)]));
  if (view.phase !== 'pre') {
    sections.push(card('Investments', view.investments.length ? 'Value at month end, and money moved in (+) or out (−).' : 'No investment accounts yet.',
      [view.investments.flatMap((f) => [field(i, f, false, manage('investments', f, { retired: f.retired })), field(i, f.deposit, true)]),
        addForm(i, 'investments')]));
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

const SETUP_WORDS = {
  accounts: { add: 'Add account', noun: 'account', full: 'All 14 account rows are used.' },
  investments: { add: 'Add investment', noun: 'investment', full: 'All 6 investment rows are used.' },
  sources: { add: 'Add income source', noun: 'income source', full: 'All 12 income-source rows are used.' },
};

/** "+ Add account" (or investment, or income source) and its small form, at the end of a section. */
function addForm(i, table) {
  const w = SETUP_WORDS[table];
  if (entry.freeSlots(state.data, table) === 0) return h('p', { class: 'help setup-full' }, w.full);
  const id = (k) => `add-${table}-${k}`;
  const box = h('div', { class: 'setup-form', id: id('form'), hidden: true });
  const open = h('button', { type: 'button', class: 'link-btn add-btn', id: id('open') }, `+ ${w.add}`);
  const error = h('p', { class: 'status tone-bad', role: 'alert', hidden: true });
  const name = h('input', { type: 'text', id: id('name'), autocomplete: 'off', maxlength: 100 });
  const select = (k, options, blank = null) => h('select', { id: id(k) },
    blank ? h('option', { value: '' }, blank) : null, options.map((o) => h('option', { value: o }, o)));
  const label = (k, text, el) => h('label', { for: id(k) }, text, el);
  let fields;
  if (table === 'sources') {
    const tax = state.data.hasTax;
    fields = [
      label('name', 'Name', name),
      label('tithe', 'Counts toward tithe?', select('tithe', ['Yes', 'No', 'Not Sure'])),
      label('usual', 'Usual amount (optional, for automatic pay)', h('input', { type: 'text', id: id('usual'), inputmode: 'decimal', autocomplete: 'off' })),
      label('frequency', 'How often (optional)', select('frequency', ['Monthly', 'Weekly'], '—')),
      label('day', 'Payday (day of month, or weekday for weekly)', h('input', { type: 'text', id: id('day'), autocomplete: 'off', maxlength: 10 })),
      tax ? label('taxable', 'Taxable?', select('taxable', ['Yes', 'No', 'Not Sure'], '—')) : null,
      tax ? label('withheld', 'Tax withheld?', select('withheld', ['Yes', 'No'], '—')) : null,
    ];
  } else {
    fields = [label('name', 'Name', name), label('type', 'Type', select('type', table === 'accounts' ? entry.ACCOUNT_TYPES : entry.INVESTMENT_TYPES))];
  }
  const help = table === 'sources'
    ? `Counts from ${monthYear(state.out.months[i].MonthStart)} on.`
    : `Starts at $0. Type its balance from the month you opened it. It isn't “required”, so a blank month counts as unchanged.`;
  const val = (k) => box.querySelector(`#${id(k)}`)?.value ?? '';
  const submit = () => {
    const problem = (msg) => { error.textContent = msg; error.hidden = false; };
    if (table === 'sources') {
      const usual = entry.parseAmount(val('usual'));
      if (usual.error) return problem(usual.error);
      const p = entry.setupProblem(state.data, 'sources', { name: val('name') });
      if (p) return problem(p);
      const dayText = val('day').trim();
      const day = /^\d{1,2}$/.test(dayText) ? Number(dayText) : dayText || null;
      entry.addSource(state.data, i, { name: val('name'), tithe: val('tithe'), usual: usual.value, frequency: val('frequency') || null, day,
        taxable: val('taxable') || null, withheld: val('withheld') || null });
    } else {
      const item = { name: val('name'), type: val('type') };
      const p = entry.setupProblem(state.data, table, item);
      if (p) return problem(p);
      if (table === 'accounts') entry.addAccount(state.data, item);
      else entry.addInvestment(state.data, item);
    }
    recalc(); save(); render();
  };
  box.append(h('div', { class: 'form-grid' }, fields), h('p', { class: 'help' }, help), error,
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn primary', onclick: submit }, 'Add'),
      h('button', { type: 'button', class: 'btn', onclick: () => { box.hidden = true; open.hidden = false; } }, 'Cancel')));
  open.addEventListener('click', () => { box.hidden = false; open.hidden = true; name.focus(); });
  return h('div', { class: 'setup' }, open, box);
}

/** The ⋯ panel under an account, investment or income source: Rename, and Close (or Ended) this month. */
function managePanel(m) {
  const w = SETUP_WORDS[m.table];
  const panel = h('div', { class: 'manage-panel', hidden: true });
  const show = (...kids) => panel.replaceChildren(...kids);
  const done = () => { recalc(); save(); render(); };
  const rename = () => {
    const input = h('input', { type: 'text', value: m.name, maxlength: 100, 'aria-label': `New name for ${m.name}` });
    const error = h('p', { class: 'status tone-bad', role: 'alert', hidden: true });
    show(h('div', { class: 'rename' }, input), error, h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn primary', onclick: () => {
        const p = entry.setupProblem(state.data, m.table, { name: input.value, type: m.table === 'sources' ? null : state.data[m.table][m.slot].type }, m.slot);
        if (p) { error.textContent = p; error.hidden = false; return; }
        entry.renameSetup(state.data, m.table, m.slot, input.value); done();
      } }, 'Save name'),
      h('button', { type: 'button', class: 'btn', onclick: menu }, 'Cancel')));
    input.focus();
  };
  const close = () => show(
    h('p', { class: 'help' }, m.table === 'sources'
      ? `${m.name} ended in ${m.month}: its automatic pay stops after this month. Past months stay as they are.`
      : `Close ${m.name} in ${m.month}? Its balance becomes $0 this month and it's marked closed. Its history stays.`),
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn danger', onclick: () => {
        if (m.table === 'sources') entry.endSource(state.data, m.i, m.slot);
        else entry.closeAccount(state.data, m.i, m.table, m.slot);
        done();
      } }, m.table === 'sources' ? 'It ended this month' : `Close ${w.noun}`),
      h('button', { type: 'button', class: 'btn', onclick: menu }, 'Cancel')));
  const menu = () => show(h('div', { class: 'btn-row' },
    h('button', { type: 'button', class: 'btn', onclick: rename }, 'Rename'),
    m.retired || m.ended ? null : h('button', { type: 'button', class: 'btn', onclick: close }, m.table === 'sources' ? `Ended in ${m.month}` : `Close in ${m.month}`),
    h('button', { type: 'button', class: 'btn', onclick: () => { panel.hidden = true; } }, 'Done')));
  menu();
  return panel;
}

function field(i, f, sub = false, manage = null) {
  const id = `f-${f.key}-${f.slot ?? 0}`;
  const isFormula = f.cell.kind === 'formula';
  const input = h('input', {
    id, type: 'text', inputmode: 'decimal', enterkeyhint: 'next', autocomplete: 'off',
    readonly: isFormula, value: f.cell.kind === 'auto' ? '' : plain(f.cell.value),
  });
  const clear = h('button', { type: 'button', class: 'clear', 'aria-label': `Use the automatic value for ${f.label}`, title: 'Use automatic' }, '×');
  const error = h('span', { class: 'error', role: 'alert', hidden: true });
  const hint = isFormula ? 'formula set in Excel; change it there' : f.hint;
  // Money moved out of an investment is negative, and some phone number pads have no minus key.
  const sign = f.key === 'deposits' && !isFormula
    ? h('button', {
      type: 'button', class: 'sign', 'aria-label': `Plus or minus for ${f.label}`, title: 'Plus or minus',
      onpointerdown: (e) => e.preventDefault(), // keep the keyboard open
    }, '±')
    : null;
  const panel = manage ? managePanel(manage) : null;
  const more = manage ? h('button', { type: 'button', class: 'manage', 'aria-label': `More for ${f.label}`, title: 'Rename or close',
    onclick: () => { panel.hidden = !panel.hidden; } }, '⋯') : null;
  const el = h('div', { class: `field${sub ? ' sub' : ''}` },
    h('div', { class: 'field-name' },
      h('label', { for: id }, f.label, f.required ? h('span', { class: 'req' }, 'required') : null, hint ? h('span', { class: 'hint' }, hint) : null),
      more),
    h('div', { class: `input-wrap${sign ? ' has-sign' : ''}` }, sign, input, isFormula ? null : clear),
    error, panel);
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
    const all = [...document.querySelectorAll('main .field input:not([readonly]), main textarea')];
    const next = all[all.indexOf(input) + 1];
    if (next) next.focus();
    else input.blur();
  });
  clear.addEventListener('click', () => { error.hidden = true; input.removeAttribute('aria-invalid'); commit(null); });
  sign?.addEventListener('click', () => {
    const next = entry.flipSign(input.value);
    input.value = next;
    const p = entry.parseAmount(next);
    if (next !== '-' && next !== '' && !p.error) commit(p.value); // a number: save the flipped value
    else {
      input.focus(); // "-": ready for the digits
      input.setSelectionRange(next.length, next.length);
    }
  });
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

// The Dashboard answers one question: are we OK, and does anything need me? Everything on it fits one phone screen,
// and each part opens the place it's about.
function dashboardScreen() {
  const v = dashboardView(state.out, state.data, state.today);
  const t = v.tiles;
  const tile = (label, x, onclick = null) => {
    const kids = [h('span', { class: 'tile-label' }, label, onclick ? h('span', { class: 'go', 'aria-hidden': 'true' }, ' ›') : null),
      h('span', { class: 'tile-value' }, whole(x.value)),
      x.line ? h('span', { class: `tile-line ${x.tone ? `tone-text-${x.tone}` : ''}` }, x.line) : null];
    return onclick ? h('button', { type: 'button', class: 'tile', onclick }, kids) : h('div', { class: 'tile' }, kids);
  };
  const goTo = (i) => openEntry(i);
  const go = () => h('span', { class: 'go', 'aria-hidden': 'true' }, '›');

  // One row: "As of Sep 30, 2026" and "Fill Oct 2026 →" (the full date and month are in the title).
  const short = (label) => label.replace(/^(\w{3})\w*/, '$1');
  const head = h('section', { class: 'card as-of' },
    h('p', { title: `As of ${v.asOfText}` }, `As of ${short(v.asOfText)}`),
    v.next.index >= 0 && v.next.index <= entry.currentMonthIndex(state.out.months, state.today)
      ? h('button', { type: 'button', class: 'btn primary', title: `Fill ${v.next.label}`, onclick: () => goTo(v.next.index) }, `Fill ${short(v.next.label)} →`)
      : null);

  const rem = reminder();
  const remind = rem.due ? h('div', { class: 'banner tone-bad', role: 'status' },
    h('span', {}, `⚠ ${rem.text}. Export so Excel has your ${changes()} new ${changes() === 1 ? 'entry' : 'entries'}.`),
    h('button', { type: 'button', class: 'btn', onclick: () => openPage('workbook') }, 'Export')) : null;

  // The most urgent health check, with "+N more"; all of them are under More → Health checks.
  const warn = v.health.filter((x) => !x.startsWith('✓'));
  const first = warn.find((x) => x.startsWith('⚠')) ?? warn[0];
  const attention = h('button', { type: 'button', class: `check-line ${warn.length ? lineTone(first) : 'tone-good'}`, id: 'attention', onclick: () => openPage('checks') },
    h('span', {}, warn.length ? first : '✓ Everything checks out'),
    h('span', { class: 'more-count' }, warn.length > 1 ? `+${warn.length - 1} more ` : '', '›'));

  const tiles = h('section', { class: 'tiles' },
    tile('Net worth', t.netWorth, () => switchTab('ledger')),
    tile('Liquid cash', t.liquid),
    tile('Owed', t.owed),
    tile('Investments', t.investments));

  const m = v.tithe;
  const tithe = h('button', { type: 'button', class: 'card line-card', id: 'tithe', onclick: () => switchTab('giving') },
    h('span', { class: 'row' },
      h('span', { class: 'line-title' }, 'Tithe this quarter'),
      h('span', { class: `row-value ${m.balance > 0 ? 'tone-text-bad' : 'tone-text-good'}` }, m.balance > 0 ? `${whole(m.balance)} owed ` : `✓ settled `, go())),
    bar(m.owed > 0 ? m.given / m.owed : 1, 'good'),
    h('span', { class: 'line-sub' }, m.line));

  const offTrack = v.goals.filter((g) => g.status === 'Above target').length;
  const goals = h('button', { type: 'button', class: 'card line-card', id: 'goals-line', onclick: () => switchTab('goals') },
    h('span', { class: 'row' },
      h('span', { class: 'line-title' }, 'Goals'),
      h('span', { class: 'row-sub' },
        v.goals.length === 0 ? 'Add your first goal ' : goalCounts(v.goals, offTrack), go())));

  const accounts = h('section', { class: 'card', id: 'accounts' },
    h('h2', {}, 'Accounts'),
    v.accounts.length === 0 ? h('p', { class: 'help' }, 'No active accounts in Settings.') : null,
    h('dl', { class: 'results accounts compact' }, v.accounts.flatMap((a) => [
      h('dt', {}, String(a.name), h('span', { class: 'hint' }, [a.type, a.estimate].filter(Boolean).join(' · '))),
      h('dd', { class: a.signed < 0 ? 'tone-text-bad' : '' }, whole(a.signed)),
    ])));

  return h('div', { class: 'dashboard' }, remind, head, attention, tiles, tithe, goals, accounts);
}

/** "⚠ 1 off track · 2 met · 6 on the way", leaving out the counts that are 0. */
function goalCounts(goals, offTrack) {
  const met = goals.filter((g) => g.status === 'Met').length;
  const onTheWay = goals.filter((g) => g.status === 'Remaining').length;
  const rest = [met ? `${met} met` : null, onTheWay ? `${onTheWay} on the way` : null].filter(Boolean).join(' · ');
  return [offTrack ? h('span', { class: 'tone-text-bad' }, `⚠ ${offTrack} off track${rest ? ' · ' : ' '}`) : null, rest ? `${rest} ` : null];
}

// ----- Goals (its own tab; written to Settings rows 69-88 on export) -----

/** The short text beside a goal's name: "✓ Met", "⚠ $100 over" or "$4,300 to go". */
function goalShort(g) {
  if (g.status === 'Met') return '✓ Met';
  if (g.status === 'Above target') return `⚠ ${whole(g.remaining)} over`;
  return `${whole(g.remaining)} to go`;
}

function goalsScreen() {
  if (state.goal !== null) return goalForm();
  const list = state.out.dashboard.goals;
  const full = state.data.goals.every((g) => g.name !== null && g.name !== '');
  const open = (goal) => { state.goal = goal; render(); scrollTo(0, 0); };
  return h('section', { class: 'card', id: 'goals' },
    h('h2', {}, 'Goals'),
    list.length === 0 ? h('p', { class: 'help' }, 'No goals yet. A goal you add goes into Settings in the workbook when you export.') : null,
    list.map((g) => {
      const tone = g.status === 'Met' ? 'good' : g.status === 'Above target' ? 'bad' : 'info';
      return h('div', { class: 'goal' },
        h('div', { class: 'goal-head' }, h('span', { class: 'goal-name' }, String(g.name)), h('span', { class: `goal-short tone-text-${tone}` }, goalShort(g))),
        bar(g.progress, tone),
        h('div', { class: 'goal-detail' },
          h('p', { class: 'help' }, goalLine(g)),
          h('button', { type: 'button', class: 'link-btn', 'aria-label': `Edit goal ${g.name}`, onclick: () => open(g.slot) }, 'Edit')));
    }),
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn', disabled: full, onclick: () => open('new') }, 'Add goal')),
    full ? h('p', { class: 'help' }, 'All 20 goal rows are used. Delete one to add another.') : null);
}

const TARGET_LABEL = {
  percentage: 'Target (% of the base)',
  multiple: 'Target (times the base)',
  'pay down to': 'Target (this amount or less)',
};

function goalForm() {
  const data = state.data;
  const editing = typeof state.goal === 'number' ? data.goals[state.goal] : null;
  const g = editing ?? { name: null, type: null, target: null, basedOn: null, currentFrom: null, account: null, targetDate: null };
  const choices = entry.goalChoices(data);
  const same = (x, v) => v !== null && x.toLowerCase() === String(v).toLowerCase();
  // A value Excel holds that isn't in the list stays selectable, so saving doesn't change it.
  const pick = (id, list, value, blank) => h('select', { id },
    h('option', { value: '', selected: value === null }, blank),
    list.map((x) => h('option', { value: x, selected: same(x, value) }, x)),
    value === null || list.some((x) => same(x, value)) ? null : h('option', { value: String(value), selected: true }, String(value)));

  const name = h('input', { type: 'text', id: 'goal-name', autocomplete: 'off', maxlength: 200, value: g.name ?? '' });
  const type = pick('goal-type', choices.types, g.type, 'Pick a type');
  const target = h('input', { type: 'text', id: 'goal-target', inputmode: 'decimal', autocomplete: 'off', value: plain(g.target) });
  const basedOn = pick('goal-based', choices.metrics, g.basedOn, 'Pick a number');
  const currentFrom = pick('goal-current', choices.metrics, g.currentFrom, 'Pick a number');
  const account = pick('goal-account', choices.accounts, g.account, 'Pick an account');
  const date = h('input', { type: 'date', id: 'goal-date', value: typeof g.targetDate === 'number' ? serialToIso(g.targetDate) : '' });
  const error = h('p', { class: 'status tone-bad', role: 'alert', hidden: true });

  const targetText = h('span', {});
  const rows = {
    target: h('label', { for: 'goal-target' }, targetText, target),
    basedOn: h('label', { for: 'goal-based' }, 'Based on', basedOn),
    currentFrom: h('label', { for: 'goal-current' }, 'Current value from', currentFrom),
    account: h('label', { for: 'goal-account' }, 'Account', account),
    targetDate: h('label', { for: 'goal-date' }, 'Target date (optional)', date),
  };
  const titheNote = h('p', { class: 'help' }, 'The target and progress come from this quarter’s tithe.');
  const show = () => {
    const need = entry.goalNeeds(type.value);
    const specific = (sel) => sel.value.toLowerCase() === 'specific account';
    rows.target.hidden = !need.target;
    rows.basedOn.hidden = !need.basedOn;
    rows.currentFrom.hidden = !need.currentFrom;
    rows.targetDate.hidden = !need.targetDate;
    rows.account.hidden = !((need.basedOn && specific(basedOn)) || (need.currentFrom && specific(currentFrom)));
    titheNote.hidden = type.value.toLowerCase() !== 'tithe balance';
    targetText.textContent = TARGET_LABEL[type.value.toLowerCase()] ?? 'Target amount';
  };
  for (const el of [type, basedOn, currentFrom]) el.addEventListener('change', show);
  show();

  const close = () => { state.goal = null; render(); scrollTo(0, 0); };
  const changed = () => { recalc(); save(); close(); };
  const saveGoal = () => {
    const amt = entry.parseAmount(target.value);
    if (!rows.target.hidden && amt.error) { error.textContent = amt.error; error.hidden = false; return; }
    const goal = {
      name: name.value, type: type.value || null, target: amt.value ?? null,
      basedOn: basedOn.value || null, currentFrom: currentFrom.value || null, account: account.value || null,
      targetDate: date.value ? dateSerial(date.value) : null,
    };
    const problem = entry.goalProblem(goal);
    if (problem) { error.textContent = problem; error.hidden = false; return; }
    if (editing) entry.updateGoal(data, state.goal, goal);
    else if (entry.addGoal(data, goal) === null) { error.textContent = 'All 20 goal rows are used.'; error.hidden = false; return; }
    changed();
  };

  return h('section', { class: 'card', id: 'goal-form' },
    h('h2', {}, editing ? 'Edit goal' : 'Add a goal'),
    h('p', { class: 'help' }, 'Saved into Settings in the workbook when you export.'),
    h('div', { class: 'form-grid' },
      h('label', { for: 'goal-name' }, 'Goal', name),
      h('label', { for: 'goal-type' }, 'Type', type),
      titheNote, rows.target, rows.basedOn, rows.currentFrom, rows.account, rows.targetDate),
    error,
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn primary', onclick: saveGoal }, editing ? 'Save' : 'Add'),
      h('button', { type: 'button', class: 'btn', onclick: close }, 'Cancel'),
      editing ? h('button', { type: 'button', class: 'btn danger', onclick: () => {
        if (!confirm(`Delete the goal "${editing.name}"?`)) return;
        entry.deleteGoal(data, state.goal);
        changed();
      } }, 'Delete') : null));
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

  const v = dashboardView(state.out, data, state.today);
  const m = v.tithe;
  const owed = h('section', { class: 'card', id: 'tithe-card' },
    h('h2', {}, 'Tithe balance', h('span', { class: `count ${m.balance > 0 ? 'tone-bad' : 'tone-good'}` }, m.balance > 0 ? 'owed' : '✓ settled')),
    h('p', { class: 'big' }, whole(m.balance)),
    bar(m.owed > 0 ? m.given / m.owed : 1, 'good'),
    h('p', { class: 'help' }, m.line),
    v.tax ? h('dl', { class: 'results' }, h('dt', {}, 'Tax-deductible giving this year'), h('dd', {}, whole(v.tax.deductible))) : null);
  const form = h('section', { class: 'card' },
    h('h2', {}, editing ? 'Edit donation' : 'Add a donation'),
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
  return h('div', {}, editing ? null : owed, form, list);
}

// ----- Pages you reach by tapping what they're about: Workbook (the chip) and Health checks (the warning line) -----

const PAGES = {
  workbook: { title: 'Workbook', screen: () => fileScreen(), needsData: false },
  checks: { title: 'Health checks', screen: () => checksScreen(), needsData: true },
};

// Navigation. Opening a month or a page adds a browser history entry, so the phone's Back button returns
// to where you were; switching tabs replaces the entry (Back from a tab leaves the app, as in other apps).
let navDepth = 0;

function navigate({ tab = state.tab, entry: e = null, page = null }) {
  try { history.replaceState({ tab: state.tab, entry: state.entry, page: state.page }, ''); } catch { /* storage of history refused */ }
  Object.assign(state, { tab, entry: e, page, gift: null, goal: null, message: null });
  try { history.pushState({ tab, entry: e, page }, ''); navDepth++; } catch { /* the ‹ Back button still works */ }
  render();
  scrollTo(0, 0);
}

function switchTab(tab) {
  Object.assign(state, { tab, entry: null, page: null, gift: null, goal: null, message: null });
  try { history.replaceState({ tab, entry: null, page: null }, ''); } catch { /* fine without */ }
  render();
  scrollTo(0, 0);
}

function goBack() {
  if (navDepth > 0) { history.back(); return; } // popstate redraws
  if (state.page) state.page = null;
  else state.entry = null;
  render();
  scrollTo(0, 0);
}

addEventListener('popstate', (e) => {
  navDepth = Math.max(0, navDepth - 1);
  const s = e.state ?? {};
  Object.assign(state, { tab: s.tab ?? 'dashboard', entry: s.entry ?? null, page: s.page ?? null, gift: null, goal: null });
  render();
  scrollTo(0, 0);
});

const openPage = (page) => navigate({ tab: state.tab, entry: state.entry, page });

function pageScreen() {
  const p = PAGES[state.page];
  return h('div', {},
    h('div', { class: 'sub-head' },
      h('button', { type: 'button', class: 'back-btn', onclick: goBack }, '‹ Back'),
      h('h2', {}, p.title)),
    !state.data && p.needsData ? noDataScreen() : p.screen());
}

function checksScreen() {
  const health = state.out.dashboard.health;
  const warn = health.filter((x) => !x.startsWith('✓'));
  const ok = health.filter((x) => x.startsWith('✓'));
  return h('div', {},
    warn.length ? h('section', { class: 'card', id: 'checks-attention' }, h('h2', {}, 'Needs attention'),
      h('ul', { class: 'checks' }, warn.map((x) => h('li', { class: lineTone(x) }, x)))) : null,
    ok.length ? h('section', { class: 'card', id: 'checks-ok' }, h('h2', {}, 'Fine'),
      h('ul', { class: 'checks' }, ok.map((x) => h('li', { class: lineTone(x) }, x)))) : null);
}

// ----- The Workbook page: import and export, and the app itself -----

function aboutScreen() {
  return h('section', { class: 'card' },
    h('h2', {}, 'This app'),
    h('dl', { class: 'facts' },
      h('dt', {}, 'Version'), h('dd', {}, state.waiting ? `${VERSION} · ${state.waitingVersion ?? 'a new version'} is ready` : `${VERSION}`),
      h('dt', {}, 'Offline'), h('dd', {}, navigator.serviceWorker?.controller ? '✓ works without internet' : 'available after the next start'),
      h('dt', {}, 'Storage'), h('dd', {}, state.persisted ? '✓ kept safe on this phone' : 'on this phone; Android may clear it when space runs low')),
    !state.persisted && !state.installed ? h('p', { class: 'help' }, 'Install the app to your home screen; Android then keeps its data safe.') : null,
    h('div', { class: 'btn-row' },
      state.waiting ? h('button', { type: 'button', class: 'btn primary', onclick: doUpdate }, 'Update') : null,
      state.installPrompt ? h('button', { type: 'button', class: 'btn', onclick: doInstall }, 'Install app') : null));
}

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
  return h('div', {},
    h('section', { class: 'card' },
      h('h2', {}, 'Your workbook'),
      facts,
      h('div', { class: 'btn-row' },
        state.data ? h('button', { type: 'button', class: 'btn primary', onclick: doExport }, 'Export workbook') : null,
        importButton(state.data ? 'Import a newer workbook' : 'Import workbook', !state.data))),
    aboutScreen(),
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
  b.addEventListener('click', () => switchTab(b.dataset.tab));
}
document.getElementById('changes-chip').addEventListener('click', () => (state.page === 'workbook' ? goBack() : openPage('workbook')));
load();
setupServiceWorker();
