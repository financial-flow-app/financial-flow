// Service worker: keeps the app's own files on the phone so it opens offline.
// - It only ever loads files from the app's own address, never another site, and never user data.
// - A new version installs in the background and then WAITS. It takes over only when the
//   user taps "Update" (the page sends {type: 'skipWaiting'}), so code never changes silently.

const VERSION = '0.8.1'; // same as src/version.js
const BUILD = '4bb950f7e038'; // fingerprint of FILES, written by scripts/stamp.mjs
const CACHE = `financial-flow-${VERSION}-${BUILD}`;
const FILES = [
  './',
  'index.html',
  'styles.css',
  'manifest.webmanifest',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/maskable-512.png',
  'src/app.js',
  'src/calc.js',
  'src/cellref.js',
  'src/dashboard.js',
  'src/entry.js',
  'src/excel.js',
  'src/ffos.js',
  'src/formula-map.js',
  'src/import.js',
  'src/store.js',
  'src/version.js',
  'src/xlsx.js',
  'src/zip.js',
];

const own = (url) => new URL(url, self.registration.scope);

const enc = new TextEncoder();
const dec = new TextDecoder();
const hex = (buf) => [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');

// An update must never keep a mix of old and new files (GitHub's CDN can hand out a file up to 10 minutes old).
// So each file is fetched under its build (…?v=BUILD, which no cache has seen), and the set is kept only if its
// fingerprint is BUILD: the same sha-256 that scripts/stamp.mjs writes. Otherwise the install fails, the version
// already on the phone keeps working, and the next visit tries again.
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    // Each body is read as soon as it arrives: an unread response holds its connection open, and a browser
    // keeps only a few open per server, so waiting for all of them first would never finish.
    const got = new Map(await Promise.all(FILES.map(async (f) => {
      const res = await fetch(new Request(`${own(f)}?v=${BUILD}`, { cache: 'reload' }));
      if (!res.ok) throw new Error(`${f}: HTTP ${res.status}`);
      return [f, { bytes: new Uint8Array(await res.arrayBuffer()), type: res.headers.get('content-type') }];
    })));
    const parts = [];
    for (const f of FILES.filter((x) => x !== './').sort()) {
      const { bytes } = got.get(f);
      parts.push(enc.encode(f), f.endsWith('.png') ? bytes : enc.encode(dec.decode(bytes).replace(/\r\n/g, '\n')));
    }
    const all = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    parts.reduce((at, p) => { all.set(p, at); return at + p.length; }, 0);
    const build = hex(await crypto.subtle.digest('SHA-256', all)).slice(0, 12);
    if (build !== BUILD) throw new Error(`the files don't match build ${BUILD} (got ${build}); trying again next time`);
    const cache = await caches.open(CACHE);
    await Promise.all([...got].map(([f, { bytes, type }]) => cache.put(own(f), new Response(bytes, { headers: { 'content-type': type ?? '' } }))));
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) if (key.startsWith('financial-flow-') && key !== CACHE) await caches.delete(key);
    await self.clients.claim();
  })());
});

self.addEventListener('message', (event) => {
  if (event.data?.type === 'skipWaiting') self.skipWaiting();
  if (event.data?.type === 'version') (event.ports[0] ?? event.source)?.postMessage({ type: 'version', version: VERSION, build: BUILD });
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return; // not ours: leave it alone
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req.mode === 'navigate' ? own('index.html') : req, { ignoreSearch: true });
    return hit || fetch(req); // same-origin only (checked above)
  })());
});
