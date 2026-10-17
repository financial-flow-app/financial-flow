// Offline, install and controlled updates: the service worker, its file list, the manifest, icons,
// and the export reminder.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { VERSION } from '../src/version.js';
import { cachedFiles, fingerprint } from '../scripts/stamp.mjs';
import { exportReminder, REMIND_AFTER_DAYS } from '../src/entry.js';

const root = new URL('../', import.meta.url);
const sw = readFileSync(new URL('sw.js', root), 'utf8');

test('service worker: same version as the app, and a current build stamp', () => {
  assert.equal(sw.match(/const VERSION = '([^']+)'/)[1], VERSION);
  assert.equal(sw.match(/const BUILD = '([^']*)'/)[1], fingerprint(sw), 'an app file changed: run node scripts/stamp.mjs');
});

test('service worker: caches every app file, and only files that exist', () => {
  const files = cachedFiles(sw);
  const expected = [
    'index.html', 'styles.css', 'manifest.webmanifest',
    ...readdirSync(new URL('icons/', root)).filter((f) => f.endsWith('.png')).map((f) => `icons/${f}`),
    ...readdirSync(new URL('src/', root)).filter((f) => f.endsWith('.js')).map((f) => `src/${f}`),
  ];
  for (const f of expected) assert.ok(files.includes(f), `${f} is not cached, so it would be missing offline`);
  for (const f of files) if (f !== './') assert.ok(existsSync(new URL(f, root)), `${f} is listed but doesn't exist`);
});

test('service worker: no silent updates, and only the app\'s own files', () => {
  // skipWaiting only in answer to the page's Update message.
  const skips = [...sw.matchAll(/skipWaiting\(\)/g)];
  assert.equal(skips.length, 1);
  assert.match(sw, /event\.data\?\.type === 'skipWaiting'\) self\.skipWaiting\(\)/);
  assert.ok(!/addEventListener\('install'[\s\S]*?skipWaiting[\s\S]*?\}\);/.test(sw.split("addEventListener('activate'")[0]), 'install must not skip waiting');
  // fetch() only once, after the same-origin check, and no outside URLs at all.
  assert.equal([...sw.matchAll(/\bfetch\(/g)].length, 1);
  assert.match(sw, /url\.origin !== self\.location\.origin\) return;[\s\S]*fetch\(req\)/);
  assert.ok(!/https?:\/\//.test(sw), 'no outside URLs');
  for (const api of ['XMLHttpRequest', 'WebSocket', 'EventSource', 'sendBeacon', 'importScripts']) assert.ok(!sw.includes(api), api);
});

test('manifest: installable on Android, icons are real PNGs of the stated size', () => {
  const m = JSON.parse(readFileSync(new URL('manifest.webmanifest', root), 'utf8'));
  assert.equal(m.name, 'Financial Flow');
  assert.equal(m.display, 'standalone');
  assert.equal(m.start_url, './');
  assert.equal(m.scope, './');
  for (const need of [['192x192', 'any'], ['512x512', 'any'], ['512x512', 'maskable']]) {
    assert.ok(m.icons.some((i) => i.sizes === need[0] && i.purpose === need[1]), `icon ${need.join(' ')}`);
  }
  for (const icon of m.icons) {
    const png = readFileSync(new URL(icon.src, root));
    assert.equal(png.subarray(1, 4).toString('ascii'), 'PNG', icon.src);
    const [w, h] = [png.readUInt32BE(16), png.readUInt32BE(20)];
    assert.equal(`${w}x${h}`, icon.sizes, icon.src);
    assert.ok(cachedFiles(sw).includes(icon.src), `${icon.src} cached`);
  }
  const html = readFileSync(new URL('index.html', root), 'utf8');
  assert.match(html, /<link rel="manifest" href="manifest\.webmanifest">/);
});

test('export reminder: monthly, only when Excel is missing entries', () => {
  const day = 86400000;
  const now = new Date(2026, 9, 17, 12).getTime();
  assert.deepEqual(exportReminder({ lastExport: null, importedAt: now - 2 * day, changes: 0 }, now), { days: null, text: 'Not exported yet', due: false });
  assert.equal(exportReminder({ lastExport: now - 3 * 3600000, importedAt: now - 9 * day, changes: 4 }, now).text, 'Last export: today');
  assert.equal(exportReminder({ lastExport: now - day, importedAt: null, changes: 0 }, now).text, 'Last export: 1 day ago');
  const late = exportReminder({ lastExport: now - REMIND_AFTER_DAYS * day, importedAt: null, changes: 3 }, now);
  assert.deepEqual(late, { days: 30, text: 'Last export: 30 days ago', due: true });
  assert.equal(exportReminder({ lastExport: now - 45 * day, importedAt: null, changes: 0 }, now).due, false, 'nothing new: no reminder');
  assert.equal(exportReminder({ lastExport: null, importedAt: now - 31 * day, changes: 1 }, now).due, true, 'never exported, imported a month ago');
  assert.equal(exportReminder({ lastExport: null, importedAt: now - 5 * day, changes: 9 }, now).due, false);
});
