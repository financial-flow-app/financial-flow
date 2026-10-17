// Source must not reach the network; fixtures must hold nothing about the machine or a person.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { readZip, entryText } from '../src/zip.js';

const root = new URL('../', import.meta.url);
const files = (dir, ext) => readdirSync(new URL(dir, root)).filter((f) => f.endsWith(ext)).map((f) => new URL(`${dir}${f}`, root));

test('source has no network APIs or external URLs', () => {
  // The Office Open XML namespace names (schemas.openxmlformats.org) are identifiers, never contacted.
  const banned = /\b(fetch|XMLHttpRequest|WebSocket|EventSource|sendBeacon|importScripts)\b|https?:\/\/(?!schemas\.openxmlformats\.org\/)/;
  for (const f of files('src/', '.js')) {
    const lines = readFileSync(f, 'utf8').split('\n');
    lines.forEach((line, i) => assert.ok(!banned.test(line), `${f.pathname}:${i + 1}: ${line.trim()}`));
  }
});

test('source builds no HTML from strings and runs no strings as code', () => {
  const banned = /\.innerHTML\b|\.outerHTML\b|insertAdjacentHTML|document\.write|\beval\(|new Function\(/;
  for (const f of files('src/', '.js')) {
    const lines = readFileSync(f, 'utf8').split('\n');
    lines.forEach((line, i) => assert.ok(!banned.test(line), `${f.pathname}:${i + 1}: ${line.trim()}`));
  }
});

test('the page has a Content-Security-Policy at least as strict as required', () => {
  const html = readFileSync(new URL('index.html', root), 'utf8');
  const meta = html.match(/<meta\s+http-equiv="Content-Security-Policy"\s+content="([^"]+)"/i);
  assert.ok(meta, 'CSP meta tag missing');
  assert.ok(html.indexOf(meta[0]) < html.search(/<script|<link/i), 'the CSP must come before any script or stylesheet');
  const csp = Object.fromEntries(meta[1].split(';').map((d) => d.trim().split(/\s+/)).filter((p) => p[0]).map(([k, ...v]) => [k, v]));
  const required = {
    'default-src': ["'self'"], 'connect-src': ["'none'"], 'form-action': ["'none'"],
    'object-src': ["'none'"], 'base-uri': ["'none'"], 'img-src': ["'self'", 'data:'],
  };
  for (const [k, v] of Object.entries(required)) assert.deepEqual(csp[k], v, `${k} must be exactly ${v.join(' ')}`);
  for (const k of ['script-src', 'style-src']) assert.deepEqual(csp[k] ?? ["'self'"], ["'self'"], `${k} must be 'self' only`);
  const all = Object.values(csp).flat();
  for (const bad of ["'unsafe-inline'", "'unsafe-eval'", '*', 'http:', 'https:', 'blob:']) assert.ok(!all.includes(bad), `CSP allows ${bad}`);
});

test('the page has no inline scripts, inline handlers, inline styles or outside URLs', () => {
  const html = readFileSync(new URL('index.html', root), 'utf8');
  for (const m of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)) {
    assert.match(m[1], /\bsrc="(?!https?:|\/\/)[^"]+"/, 'every script is a local file');
    assert.equal(m[2].trim(), '', 'no inline script');
  }
  assert.ok(!/\son[a-z]+\s*=/i.test(html), 'inline event handler');
  assert.ok(!/\sstyle\s*=/i.test(html), 'inline style attribute');
  for (const f of [new URL('index.html', root), new URL('styles.css', root)]) {
    const text = readFileSync(f, 'utf8');
    assert.ok(!/(?:src|href)\s*=\s*"(?:https?:)?\/\//i.test(text) && !/url\(\s*['"]?(?:https?:)?\/\//i.test(text) && !/@import/i.test(text), `${f.pathname}: outside URL`);
  }
});

test('no personal-data files: only the two synthetic workbook fixtures, no exports, backups or keys', () => {
  const skip = new Set(['.git', 'node_modules', 'test-output']);
  const walk = (dir, rel = '') => readdirSync(dir, { withFileTypes: true }).flatMap((d) =>
    skip.has(d.name) ? [] : d.isDirectory() ? walk(new URL(`${d.name}/`, dir), `${rel}${d.name}/`) : [`${rel}${d.name}`]);
  const files = walk(root);
  const allowedBooks = ['fixtures/ffos-template.xlsx', 'fixtures/ffos-template-excel.xlsx'];
  const risky = /\.(xlsx|xlsm|xls|csv|vcf|ods|numbers|zip|bak|db|sqlite|json\.gz|pem|key|p12|env)$|(^|\/)\.env/i;
  const found = files.filter((f) => risky.test(f) && !allowedBooks.includes(f));
  assert.deepEqual(found, [], 'personal-data file types must not be committed');
});

// A list of private words kept OUTSIDE this repo (one folder up). Skipped where it doesn't exist.
const wordsFile = new URL('../privacy-words.txt', root);
test('no private words anywhere in the repo, xlsx contents included', { skip: !existsSync(wordsFile) && 'no word list on this machine' }, async () => {
  const words = readFileSync(wordsFile, 'utf8').split('\n').map((w) => w.trim()).filter((w) => w && !w.startsWith('#'));
  const re = new RegExp(`\\b(${words.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})\\b`, 'i');
  const skip = new Set(['.git', 'node_modules', 'test-output']);
  const walk = (dir) => readdirSync(dir, { withFileTypes: true }).flatMap((d) =>
    skip.has(d.name) ? [] : d.isDirectory() ? walk(new URL(`${d.name}/`, dir)) : [new URL(d.name, dir)]);
  const hits = [];
  for (const f of walk(root)) {
    if (f.pathname.endsWith('.png')) continue;
    const texts = f.pathname.endsWith('.xlsx')
      ? await Promise.all(readZip(readFileSync(f)).entries.map(async (e) => [`${f.pathname} ${e.name}`, await entryText(e)]))
      : [[f.pathname, readFileSync(f, 'utf8')]];
    for (const [where, t] of texts) {
      const m = re.exec(t);
      if (m) hits.push(`${decodeURIComponent(where)}: "${m[0]}"`);
    }
  }
  assert.deepEqual(hits, []);
});

test('fixtures hold no personal metadata, local paths or email addresses', async () => {
  const backslash = String.fromCharCode(92);
  const pathLike = new RegExp(`[A-Za-z]:${backslash}${backslash}|/Users/|/home/`);
  const email = /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+\.[A-Za-z]{2,}/;
  for (const f of files('fixtures/', '.xlsx')) {
    for (const e of readZip(readFileSync(f)).entries) {
      const t = await entryText(e);
      assert.ok(!pathLike.test(t), `${f.pathname} ${e.name}: local path`);
      assert.ok(!email.test(t), `${f.pathname} ${e.name}: email address`);
      if (e.name === 'docProps/core.xml') {
        for (const tag of ['dc:creator', 'cp:lastModifiedBy']) {
          const v = t.match(new RegExp(`<${tag}>([^<]*)</${tag}>`))?.[1];
          assert.ok(v == null || ['openpyxl', 'Synthetic'].includes(v), `${f.pathname}: ${tag} is "${v}"`);
        }
      }
      if (e.name === 'docProps/app.xml') assert.ok(!/<(Company|Manager)>[^<S]/.test(t), `${f.pathname}: company/manager`);
    }
  }
});
