// Writes a fingerprint of every file sw.js caches into sw.js (const BUILD).
// A phone only sees an update when sw.js changes, so any change to an app file must be
// followed by `node scripts/stamp.mjs`; test/offline.test.mjs fails until it is.
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

const root = new URL('../', import.meta.url);

export function cachedFiles(swText) {
  const list = swText.match(/const FILES = \[([\s\S]*?)\];/)[1];
  return [...list.matchAll(/'([^']+)'/g)].map((m) => m[1]);
}

export function fingerprint(swText) {
  const hash = createHash('sha256');
  for (const f of cachedFiles(swText).filter((f) => f !== './').sort()) {
    hash.update(f);
    const bytes = readFileSync(new URL(f, root));
    // Text files: the same fingerprint whether a checkout has LF or CRLF line endings.
    hash.update(f.endsWith('.png') ? bytes : bytes.toString('utf8').replace(/\r\n/g, '\n'));
  }
  return hash.digest('hex').slice(0, 12);
}

if (process.argv[1] && import.meta.url.endsWith(process.argv[1].replace(/\\/g, '/').split('/').pop())) {
  const swUrl = new URL('sw.js', root);
  const sw = readFileSync(swUrl, 'utf8');
  const build = fingerprint(sw);
  writeFileSync(swUrl, sw.replace(/const BUILD = '[^']*';/, `const BUILD = '${build}';`));
  console.log(`sw.js BUILD = ${build}`);
}
