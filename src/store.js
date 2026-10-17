// Everything the app keeps lives in this browser's IndexedDB, on this phone only.
//   original    the workbook bytes as last imported or exported (export patches these)
//   base        that workbook's data, as imported
//   data        the data with the user's entries since then
//   fileName, importedAt, lastExport

const DB = 'financial-flow';
const STORE = 'kv';

function open() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

let dbPromise;
const db = () => (dbPromise ??= open());

function run(mode, fn) {
  return db().then((d) => new Promise((resolve, reject) => {
    const tx = d.transaction(STORE, mode);
    const out = fn(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(out && 'result' in out ? out.result : undefined);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  }));
}

export const get = (key) => run('readonly', (s) => s.get(key));

/** Write several keys in one transaction (all or nothing). */
export const setMany = (entries) => run('readwrite', (s) => { for (const [k, v] of Object.entries(entries)) s.put(v, k); });

export const clearAll = () => run('readwrite', (s) => s.clear());

/** Ask the browser not to clear this data when space runs low. Returns true if granted. */
export async function persist() {
  try {
    if (!navigator.storage?.persist) return false;
    return (await navigator.storage.persisted()) || (await navigator.storage.persist());
  } catch {
    return false;
  }
}

export async function persisted() {
  try {
    return navigator.storage?.persisted ? await navigator.storage.persisted() : false;
  } catch {
    return false;
  }
}
