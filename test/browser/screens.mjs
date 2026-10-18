// The screens in a real browser at phone size (375 x 812), on the local test server: Chromium (the
// engine of Android's Chrome) by default, or WebKit (the engine of every iPhone browser) with
// BROWSER=webkit. It uses the app the way a person would: import a synthetic workbook, read the
// Dashboard, visit the More pages, fill a month, add a donation, reload, export through the share
// sheet, read the shared file back with the app's own importer, reload again, then add and edit a
// goal. It stops at the first case that fails and says which one, after how long, and what the
// screen showed.
//
// The clock is fixed at Nov 10, 2026, noon UTC, so every run sees the same months and statuses.
// Every wait is for a condition with a deadline, never a fixed time (GitHub's machines are slower).
//
// Needs Playwright, which the Check workflow installs for this step only (it is not a dependency):
//   npm install --no-save --no-package-lock playwright@1.63.0
//   npx playwright install chromium webkit
//   node test/browser/screens.mjs                      (Chromium)
//   BROWSER=webkit node test/browser/screens.mjs       (WebKit)

import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual } from 'node:util';
import assert from 'node:assert/strict';
import { chromium, webkit } from 'playwright';
import { importWorkbook } from '../../src/import.js';
import { calculate } from '../../src/calc.js';
import { dashboardView } from '../../src/dashboard.js';
import { dollars } from '../../src/excel.js';
import { dateSerial } from '../../src/ffos.js';

const root = fileURLToPath(new URL('../../', import.meta.url));
const NOW = new Date('2026-11-10T12:00:00Z');
const TODAY = dateSerial('2026-11-10');
const SCENARIO = 'goals'; // test/scenarios/index.mjs: months filled to September 2026, two required balances
const PHONE = { viewport: { width: 375, height: 812 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true };
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const GIFT = { amount: '123.45', recipient: 'Charity D' };

// ---------- waiting for conditions ----------

/** Read until `want` (a value, or a test) holds; fail after `ms` with what was last seen. */
async function soon(what, read, want, ms = 10000) {
  const holds = typeof want === 'function' ? want : (v) => isDeepStrictEqual(v, want);
  const t0 = Date.now();
  let last;
  for (;;) {
    try {
      last = await read();
      if (holds(last)) return last;
    } catch (e) {
      last = `error: ${e.message.split('\n')[0]}`; // e.g. the page is reloading
    }
    if (Date.now() - t0 > ms) {
      const expected = typeof want === 'function' ? '' : ` (expected ${JSON.stringify(want)})`;
      throw new Error(`${what}: not so after ${Date.now() - t0} ms; last seen ${JSON.stringify(last)}${expected}`);
    }
    await new Promise((r) => setTimeout(r, 50));
  }
}

// ---------- the test server ----------

const freePort = () => new Promise((resolve, reject) => {
  const s = createServer();
  s.on('error', reject);
  s.listen(0, '127.0.0.1', () => { const { port } = s.address(); s.close(() => resolve(port)); });
});

async function startServer(port) {
  const proc = spawn(process.execPath, ['scripts/serve.mjs', String(port)], { cwd: root, stdio: ['ignore', 'pipe', 'inherit'] });
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('the test server did not start within 10 s')), 10000);
    proc.stdout.on('data', (b) => { if (String(b).includes(`:${port}`)) { clearTimeout(timer); resolve(); } });
    proc.on('exit', (code) => { clearTimeout(timer); reject(new Error(`the test server stopped (exit code ${code})`)); });
  });
  return proc;
}

/** The synthetic scenario workbook the test server builds (the same file a browser test would get). */
async function fixture(base, name) {
  const js = await (await fetch(`${base}__dev/fixture.js?name=${name}`)).text();
  const { name: fileName, b64 } = JSON.parse(js.replace(/^window\.__fixture = /, '').replace(/;\s*$/, ''));
  return { name: fileName, bytes: Buffer.from(b64, 'base64') };
}

// ---------- the run ----------

const steps = [];
const step = (name, fn) => steps.push({ name, fn });

const port = await freePort();
const server = await startServer(port);
process.on('exit', () => server.kill());
const base = `http://127.0.0.1:${port}/`;
const book = await fixture(base, SCENARIO);
const before = await importWorkbook(book.bytes);
const OCT = before.months.findIndex((m) => m.ym === '2026-10');
const ENGINE = process.env.BROWSER === 'webkit' ? 'WebKit' : 'Chromium';
const browser = await (ENGINE === 'WebKit' ? webkit : chromium).launch();
const context = await browser.newContext({ ...PHONE, locale: 'en-US', timezoneId: 'UTC', colorScheme: 'light' });

// The share sheet: Android's would offer Quick Share; this one keeps the file for the check.
await context.addInitScript(() => {
  const shared = [];
  const can = (d) => Array.isArray(d?.files) && d.files.length > 0;
  Object.defineProperty(window, '__shared', { value: shared });
  Object.defineProperty(navigator, 'canShare', { configurable: true, value: can });
  Object.defineProperty(navigator, 'share', {
    configurable: true,
    value: async (d) => {
      if (!can(d)) throw new TypeError('nothing to share');
      const f = d.files[0];
      const bytes = new Uint8Array(await f.arrayBuffer());
      let bin = '';
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      shared.push({ name: f.name, type: f.type, b64: btoa(bin) });
    },
  });
});

const page = await context.newPage();
page.setDefaultTimeout(10000);
await page.clock.setFixedTime(NOW);

// Anything that goes wrong outside the steps' own checks: errors, dialogs, requests to other addresses.
const problems = [];
page.on('console', (m) => { if (m.type() === 'error' && !m.text().includes('/csp-test')) problems.push(`console error: ${m.text()}`); });
page.on('pageerror', (e) => problems.push(`page error: ${e.message}`));
page.on('dialog', (d) => { problems.push(`unexpected dialog: ${d.message()}`); d.dismiss().catch(() => {}); });
context.on('request', (r) => { if (!r.url().startsWith(base) && !/^(data|blob):/.test(r.url())) problems.push(`request outside the app: ${r.url()}`); });

// ---------- reading the screen ----------

const squash = (s) => (s == null ? null : s.replace(/\s+/g, ' ').trim());
const text = async (sel) => squash(await page.evaluate((s) => document.querySelector(s)?.textContent ?? null, sel));
const value = (sel) => page.evaluate((s) => document.querySelector(s)?.value ?? null, sel);
const chip = () => page.evaluate(() => { const c = document.getElementById('changes-chip'); return c.hidden ? '' : c.textContent; });
const monthShown = () => page.evaluate(() => document.getElementById('month-select')?.selectedOptions[0]?.textContent ?? null);
/** The value next to a label in a <dl> on the screen ("Not exported" -> "3 entries"). */
const fact = async (label) => squash(await page.evaluate((l) =>
  [...document.querySelectorAll('main dt')].find((dt) => dt.textContent.trim() === l)?.nextElementSibling?.textContent ?? null, label));
const gifts = () => page.evaluate(() => [...document.querySelectorAll('main .gift')].map((g) => ({
  who: g.firstElementChild.firstChild.textContent, details: g.querySelector('small').textContent, amount: g.querySelector('.amt').textContent,
})));
/** What the app saved in IndexedDB: a key, then a path inside it ("months.7.deposits.0"). */
const stored = (key, path = '') => page.evaluate(([key, path]) => new Promise((resolve, reject) => {
  const open = indexedDB.open('financial-flow');
  open.onerror = () => reject(open.error);
  open.onsuccess = () => {
    const db = open.result;
    const get = db.transaction('kv').objectStore('kv').get(key);
    get.onsuccess = () => { db.close(); resolve(path.split('.').filter(Boolean).reduce((o, k) => o?.[k], get.result) ?? null); };
    get.onerror = () => { db.close(); reject(get.error); };
  };
}), [key, path]);

const tab = (name) => page.click(`.tabs button[data-tab="${name}"]`);
/** Open a page under ☰ More ("Import & export", "History", ...). */
async function more(label) {
  await page.click('#more-btn');
  await page.locator('.menu-item', { hasText: label }).click();
  await soon(`More → ${label} opens`, () => text('.sub-head h2'), label);
}
async function fitsPhone(screen) {
  const width = await page.evaluate(() => document.documentElement.scrollWidth);
  assert.ok(width <= PHONE.viewport.width, `${screen} is ${width}px wide, so the phone would scroll sideways`);
}
const whole = (v) => (typeof v === 'number' && Number.isFinite(v) ? dollars(v) : '—');

// ---------- the steps ----------

step('opens with no errors, and the browser enforces the page\'s security policy', async () => {
  await page.goto(base);
  await soon('the start screen shows', () => text('main h2'), 'Start with your workbook');
  const csp = await page.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.content ?? null);
  assert.ok(csp, 'the Content-Security-Policy meta tag is missing');
  for (const rule of ["default-src 'self'", "script-src 'self'", "connect-src 'none'", "object-src 'none'", "base-uri 'none'", "form-action 'none'"]) {
    assert.ok(csp.includes(rule), `the security policy lacks ${rule}`);
  }
  // Even a request to the app's own address is refused: the page can't send data anywhere.
  const refused = await page.evaluate(() => new Promise((resolve) => {
    addEventListener('securitypolicyviolation', (e) => resolve(e.effectiveDirective), { once: true });
    fetch('csp-test').then(() => resolve('allowed'), () => {});
    setTimeout(() => resolve('no answer within 5 s'), 5000);
  }));
  assert.equal(refused, 'connect-src', 'a connection from the page should be refused by connect-src');
  await fitsPhone('the start screen');
});

step(`imports a synthetic workbook ("${SCENARIO}") and opens the month to fill`, async () => {
  await page.setInputFiles('main input[type=file]', { name: book.name, mimeType: XLSX_TYPE, buffer: book.bytes });
  await soon('the Month screen opens on October 2026 (past, incomplete)', monthShown, 'October 2026 ⚠');
  assert.equal(await text('#month-status'), '⚠ Needs balances');
  assert.equal(await text('#missing-count'), '2 required missing');
  assert.equal(await chip(), '', 'nothing is waiting to be exported right after an import');
  await fitsPhone('the Month screen');
});

step('Dashboard: one screen answers "are we OK, and does anything need me?"', async () => {
  const v = dashboardView(calculate(structuredClone(before), TODAY), before, TODAY);
  assert.equal(v.asOfText, 'September 30, 2026');
  await tab('dashboard');
  await soon('the as-of line', () => text('.as-of p'), 'As of Sep 30, 2026');
  assert.equal(await text('.as-of button'), 'Fill Oct 2026 →');
  const tiles = await page.evaluate(() => [...document.querySelectorAll('.tile')].map((t) =>
    [t.querySelector('.tile-label').textContent, t.querySelector('.tile-value').textContent]));
  assert.deepEqual(tiles, [
    ['Net worth ›', whole(v.tiles.netWorth.value)],
    ['Liquid cash', whole(v.tiles.liquid.value)],
    ['Owed', whole(v.tiles.owed.value)],
    ['Investments', whole(v.tiles.investments.value)],
  ]);
  assert.ok(tiles.every(([, x]) => /^-?\$[\d,]+$/.test(x)), `every tile shows an amount: ${JSON.stringify(tiles)}`);
  // The most urgent health check, with "+N more".
  const warn = v.health.filter((x) => !x.startsWith('✓'));
  assert.ok(warn.length > 1, 'the scenario has more than one health line that needs attention');
  assert.equal(await text('#attention'), `${warn.find((x) => x.startsWith('⚠'))}+${warn.length - 1} more ›`);
  assert.equal(await text('#tithe .row'), `Tithe this quarter${whole(v.tithe.balance)} owed ›`);
  assert.equal(await text('#goals-line'), 'Goals⚠ 1 off track · 8 on the way ›');
  const accounts = await page.evaluate(() => [...document.querySelectorAll('#accounts dt')].map((dt) => dt.firstChild.textContent));
  assert.deepEqual(accounts, v.accounts.map((a) => a.name));
  // In this order, and all of it above the tabs without scrolling.
  const order = await page.evaluate(() => [...document.querySelector('main .dashboard').children].map((el) => el.id || el.className.split(' ').at(-1)));
  assert.deepEqual(order, ['as-of', 'attention', 'tiles', 'tithe', 'goals-line', 'accounts']);
  const room = await page.evaluate(() => ({ bottom: document.getElementById('accounts').getBoundingClientRect().bottom, top: document.querySelector('.tabs').getBoundingClientRect().top }));
  assert.ok(room.bottom <= room.top + 1, `the Dashboard should end above the tabs: it ends at ${room.bottom}px, the tabs start at ${room.top}px`);
  await fitsPhone('the Dashboard');
});

step('Net worth opens History; the phone\'s Back button returns to the More menu', async () => {
  const v = dashboardView(calculate(structuredClone(before), TODAY), before, TODAY);
  await page.click('.tile:has-text("Net worth")');
  await soon('History opens', () => text('.sub-head h2'), 'History');
  assert.equal(await text('#h-year h2'), 'This year (2026)');
  assert.equal(await fact('Income'), whole(v.flow.rows[0][3]), 'this year\'s income');
  assert.equal(await text('#h-avg h2'), 'Average month (last 3 full)');
  const months = await page.evaluate(() => [...document.querySelectorAll('#h-months .month-name')].map((n) => n.textContent));
  assert.equal(months[0], 'November 2026', 'newest first (the current month, in progress)');
  assert.equal(months.at(-1), 'April 2026 ✓', 'back to the first month after the opening date');
  assert.equal(months.length, 8);
  await fitsPhone('History');
  await page.goBack();
  await soon('Back returns to the More menu', () => page.evaluate(() => [...document.querySelectorAll('.menu-title')].map((t) => t.firstChild.textContent)),
    ['History', 'Import & export', 'Health checks', 'About']);
  await page.locator('.menu-item', { hasText: 'Health checks' }).click();
  await soon('Health checks lists what needs attention, then what is fine',
    () => page.evaluate(() => [...document.querySelectorAll('main .checks li')].map((li) => li.textContent)),
    [...v.health.filter((x) => !x.startsWith('✓')), ...v.health.filter((x) => x.startsWith('✓'))]);
  await page.click('.back-btn');
  await page.locator('.menu-item', { hasText: 'About' }).click();
  await soon('About shows the version', () => fact('Version'), (x) => /^\d+\.\d+\.\d+/.test(x ?? ''));
  await fitsPhone('About');
  await tab('dashboard');
});

step('Month: a required balance changes the status, and × clears it', async () => {
  await page.click('.as-of button'); // "Fill October 2026 →"
  await soon('"Fill October 2026" opens that month', monthShown, 'October 2026 ⚠');
  await page.fill('#f-balances-0', '10000');
  await page.press('#f-balances-0', 'Enter');
  await soon('Checking A shows the typed balance', () => value('#f-balances-0'), '10,000');
  await soon('one required balance is still missing', () => text('#missing-count'), '1 required missing');
  assert.equal(await text('#month-status'), '⚠ Needs balances');
  await page.fill('#f-balances-2', '650');
  await page.press('#f-balances-2', 'Enter');
  await soon('with both required balances, the month is done', () => text('#month-status'), '✓ Done');
  assert.equal(await text('#missing-count'), '✓ all required in');
  assert.equal(await monthShown(), 'October 2026 ✓');
  await page.click('.input-wrap:has(#f-balances-2) .clear');
  await soon('× empties Card A', () => value('#f-balances-2'), '');
  await soon('without Card A the month needs balances again', () => text('#month-status'), '⚠ Needs balances');
  assert.equal(await page.isHidden('.input-wrap:has(#f-balances-2) .clear'), true, '× hides once the field is automatic again');
  await soon('the typed balance is saved on the phone', () => stored('data', `months.${OCT}.balances.0`), { kind: 'typed', value: 10000 });
  assert.equal(await chip(), '1 not exported');
});

step('Month: ± on "Moved into" puts in a minus sign, and flips a number and saves it', async () => {
  assert.equal((await text('label[for="f-deposits-0"]')).startsWith('Moved into Roth IRA'), true);
  const sign = '.input-wrap:has(#f-deposits-0) .sign';
  await page.click(sign);
  await soon('± on the empty field leaves "-" with the cursor in it',
    () => page.evaluate(() => ({ value: document.getElementById('f-deposits-0').value, focused: document.activeElement?.id })),
    { value: '-', focused: 'f-deposits-0' });
  await page.keyboard.type('250');
  await page.keyboard.press('Enter');
  await soon('-250 is saved on the phone', () => stored('data', `months.${OCT}.deposits.0`), { kind: 'typed', value: -250 });
  assert.equal(await value('#f-deposits-0'), '-250');
  await page.click(sign);
  await soon('± flips -250 to 250 and saves it',
    async () => [await value('#f-deposits-0'), await stored('data', `months.${OCT}.deposits.0`)], ['250', { kind: 'typed', value: 250 }]);
  await page.click(sign);
  await soon('± flips it back to -250 and saves it',
    async () => [await value('#f-deposits-0'), await stored('data', `months.${OCT}.deposits.0`)], ['-250', { kind: 'typed', value: -250 }]);
  assert.equal(await chip(), '2 not exported');
});

step('Giving: a tax-deductible donation is added and listed', async () => {
  await tab('giving');
  await soon('the form starts on today\'s date', () => value('#g-date'), '2026-11-10');
  await page.fill('#g-amount', GIFT.amount);
  await page.fill('#g-recipient', GIFT.recipient);
  await page.check('#g-deductible');
  await page.getByRole('button', { name: 'Add', exact: true }).click();
  await soon('the donation is first in the list', async () => (await gifts())[0],
    { who: GIFT.recipient, details: 'Nov 10, 2026 · tax-deductible', amount: '$123.45' });
  assert.equal((await gifts()).length, before.giving.length + 1);
  assert.equal(await text('main .gift-month'), 'November 2026$123.45');
  assert.equal(await value('#g-amount'), '', 'the form is empty again for the next donation');
  assert.equal(await chip(), '3 not exported');
  await fitsPhone('the Giving screen');
  await soon('Giving counts it as tax-deductible giving', () => fact('Tax-deductible giving this year'), '$123');
});

step('reload: the entries stay on the phone (IndexedDB), still waiting to be exported', async () => {
  await soon('the donation is saved on the phone', () => stored('data', 'giving'), (g) => g?.some((x) => x.recipient === GIFT.recipient));
  await page.reload();
  await soon('after the reload, 3 entries are not exported', chip, '3 not exported');
  await tab('month');
  await soon('the balance is still there', () => value('#f-balances-0'), '10,000');
  assert.equal(await value('#f-deposits-0'), '-250');
  assert.equal(await value('#f-balances-2'), '');
  await tab('giving');
  await soon('the donation is still there', async () => (await gifts())[0]?.who, GIFT.recipient);
});

let shared;
step('Import & export: export sends the workbook to the share sheet', async () => {
  await soon('the More button shows no export-due dot (the last export was today)', () => page.isHidden('#more-dot'), true);
  await more('Import & export');
  await soon('Import & export shows what is not exported', () => fact('Not exported'), '3 entries');
  await fitsPhone('the Import & export screen');
  await page.getByRole('button', { name: 'Export workbook' }).click();
  await soon('one file reaches the share sheet', () => page.evaluate(() => window.__shared.length), 1);
  shared = await page.evaluate(() => window.__shared[0]);
  assert.equal(shared.name, 'goals 2026-11-10 12.00 PM.xlsx');
  assert.equal(shared.type, XLSX_TYPE);
  await soon('the app says it exported', () => text('main p.status'), `Exported ${shared.name}.`);
  await soon('nothing is left to export', chip, '');
  assert.equal(await fact('Not exported'), '✓ nothing');
  assert.ok((await fact('Last export')).startsWith('Nov 10, 2026'), `Last export: ${await fact('Last export')}`);
});

step('the shared file, read back with src/import.js, holds exactly the entries made', async () => {
  const after = await importWorkbook(Buffer.from(shared.b64, 'base64'));
  assert.equal(after.exportStamp, NOW.toISOString(), 'the export stamp inside the file');
  const oct = after.months[OCT];
  assert.deepEqual(oct.balances[0], { kind: 'typed', value: 10000 }, 'Checking A in October');
  assert.deepEqual(oct.balances[2], { kind: 'auto', value: null }, 'Card A, cleared with ×, is blank again');
  assert.deepEqual(oct.deposits[0], { kind: 'typed', value: -250 }, 'Moved into Roth IRA in October');
  const gift = after.giving.find((g) => g.recipient === GIFT.recipient);
  assert.ok(gift, 'the donation is in the Giving log');
  const { row, ...rest } = gift;
  assert.deepEqual(rest, { date: TODAY, amount: 123.45, recipient: GIFT.recipient, paidFrom: null, household: null, note: null, deductible: 'Yes' });
  // Everything else is as the imported workbook had it.
  const others = (d) => ({
    ...d,
    exportStamp: null,
    months: d.months.map((m, i) => (i === OCT ? { ...m, balances: m.balances.slice(1), deposits: m.deposits.slice(1) } : m)),
    giving: d.giving.filter((g) => g.recipient !== GIFT.recipient),
  });
  assert.deepEqual(others(after), others(before));
});

step('reload after the export: the export is remembered and the entries stay', async () => {
  await page.reload();
  await more('Import & export');
  await soon('the stamp of the export shows', () => fact('Stamp in file'), 'exported Nov 10, 2026, 12:00 PM');
  assert.equal(await fact('Not exported'), '✓ nothing');
  assert.equal(await fact('Workbook'), shared.name);
  assert.equal(await chip(), '');
  await tab('month');
  await soon('the balance is still there', () => value('#f-balances-0'), '10,000');
});

step('Goals: add one on the Goals tab; the form asks only what the type needs', async () => {
  await tab('goals');
  const names = () => page.evaluate(() => [...document.querySelectorAll('#goals .goal-name')].map((n) => n.textContent));
  await soon('the Goals tab lists the workbook\'s goals', async () => (await names()).length, before.goals.filter((g) => g.name).length);
  await page.getByRole('button', { name: 'Add goal' }).click();
  await soon('the goal form opens', () => text('#goal-form h2'), 'Add a goal');
  const shown = () => page.evaluate(() => ['goal-target', 'goal-based', 'goal-current', 'goal-account', 'goal-date']
    .filter((id) => !document.querySelector(`label[for="${id}"]`).hidden));
  await page.getByRole('button', { name: 'Add', exact: true }).click();
  await soon('a goal without a name is refused', () => text('#goal-form p.status'), 'Give the goal a name.');
  await page.selectOption('#goal-type', 'Save By Date');
  assert.deepEqual(await shown(), ['goal-target', 'goal-current', 'goal-date']);
  await page.selectOption('#goal-current', 'Specific Account');
  assert.deepEqual(await shown(), ['goal-target', 'goal-current', 'goal-account', 'goal-date'], 'Specific Account asks for the account');
  await page.selectOption('#goal-type', 'Multiple');
  assert.deepEqual(await shown(), ['goal-target', 'goal-based', 'goal-current', 'goal-account']);
  assert.equal(await text('label[for="goal-target"] span'), 'Target (times the base)');
  await page.selectOption('#goal-type', 'Save By Date');
  await fitsPhone('the goal form');
  await page.fill('#goal-name', 'Synthetic trip');
  await page.fill('#goal-target', '2,500');
  await page.selectOption('#goal-account', 'Emergency fund');
  await page.fill('#goal-date', '2027-03-31');
  await page.getByRole('button', { name: 'Add', exact: true }).click();
  await soon('the new goal is listed', async () => (await names()).at(-1), 'Synthetic trip');
  assert.equal(await chip(), '1 not exported');
  await soon('the goal is saved on the phone', () => stored('data', 'goals.9'),
    { name: 'Synthetic trip', type: 'Save By Date', target: 2500, basedOn: null, currentFrom: 'Specific Account', account: 'Emergency fund', targetDate: dateSerial('2027-03-31') });
});

step('Goals: Edit changes the type, and fields it no longer uses are cleared', async () => {
  await page.getByRole('button', { name: 'Edit goal Synthetic trip' }).click();
  await soon('the form opens with the goal', () => value('#goal-name'), 'Synthetic trip');
  assert.equal(await value('#goal-target'), '2,500');
  assert.equal(await value('#goal-date'), '2027-03-31');
  await page.selectOption('#goal-type', 'Tithe Balance');
  assert.equal(await page.isHidden('label[for="goal-target"]'), true, 'Tithe Balance needs no target');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await soon('the edited goal is saved on the phone', () => stored('data', 'goals.9'),
    { name: 'Synthetic trip', type: 'Tithe Balance', target: null, basedOn: null, currentFrom: null, account: null, targetDate: null });
  await soon('the card shows it with this quarter\'s tithe', () => page.evaluate(() => [...document.querySelectorAll('#goals .goal .help')].at(-1)?.textContent ?? null),
    (t) => t?.endsWith('This quarter'));
  assert.equal(await chip(), '1 not exported', 'one goal changed since the export');
  await fitsPhone('the Goals tab with the new goal');
});

// ---------- run them, in order, stopping at the first failure ----------

console.log(`Screens check: ${ENGINE} ${browser.version()}, ${PHONE.viewport.width}x${PHONE.viewport.height}, clock fixed at ${NOW.toISOString()}`);
let failed = false;
try {
  for (const s of steps) {
    const t0 = Date.now();
    try {
      await s.fn();
      if (problems.length) throw new Error(problems.join('\n'));
      console.log(`PASS ${s.name} (${Date.now() - t0} ms)`);
    } catch (e) {
      failed = true;
      const ms = Date.now() - t0;
      console.log(`FAIL ${s.name} (after ${ms} ms)\n  ${e.message.replace(/\n/g, '\n  ')}`);
      if (process.env.GITHUB_ACTIONS) console.log(`::error title=Screens check (${ENGINE})::${s.name} failed after ${ms} ms: ${e.message.split('\n')[0]}`);
      const screen = await page.evaluate(() => document.querySelector('main')?.innerText ?? '').catch(() => '');
      console.log(`  The screen showed:\n    ${screen.slice(0, 1500).replace(/\n+/g, '\n    ')}`);
      const dir = new URL('../../test-output/browser/', import.meta.url);
      mkdirSync(dir, { recursive: true });
      await page.screenshot({ path: fileURLToPath(new URL(`failure-${ENGINE.toLowerCase()}.png`, dir)), fullPage: true }).catch(() => {});
      break;
    }
  }
} finally {
  await browser.close();
  server.kill();
}
if (failed) process.exitCode = 1;
else console.log(`All ${steps.length} cases passed.`);
