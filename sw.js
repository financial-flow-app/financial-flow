// Service worker: keeps the app's own files on the phone so it opens offline.
// - It only ever loads files from the app's own address, never another site, and never user data.
// - A new version installs in the background and then WAITS. It takes over only when the
//   user taps "Update" (the page sends {type: 'skipWaiting'}), so code never changes silently.

const VERSION = '0.8.0'; // same as src/version.js
const BUILD = '03079ad122b1'; // fingerprint of FILES, written by scripts/stamp.mjs
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

self.addEventListener('install', (event) => {
  // cache: 'reload' skips the browser's HTTP cache, so the new version's files are the new ones.
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(own(f), { cache: 'reload' })))));
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
