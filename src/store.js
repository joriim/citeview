// store.js — keeps the source PDFs the user added (IndexedDB) and
// citation→file assignments (localStorage), so they survive restarts.

const DB = 'citeview';
const STORE = 'files';

function open() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'name' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx(mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode);
    const out = fn(t.objectStore(STORE));
    t.oncomplete = () => resolve(out instanceof IDBRequest ? out.result : out);
    t.onerror = () => reject(t.error);
  });
}

/** file: {name, data: ArrayBuffer, dois: string[]} */
export const putFile = (file) => tx('readwrite', s => s.put(file));
export const getFile = (name) => tx('readonly', s => s.get(name));
export const listFiles = () => tx('readonly', s => s.getAll());
export const clearFiles = () => tx('readwrite', s => s.clear());

const MAP_KEY = 'citeview.assignments';

export function getAssignment(address) {
  try { return JSON.parse(localStorage.getItem(MAP_KEY) || '{}')[address] || null; } catch { return null; }
}

export function setAssignment(address, fileName) {
  try {
    const m = JSON.parse(localStorage.getItem(MAP_KEY) || '{}');
    m[address] = fileName;
    localStorage.setItem(MAP_KEY, JSON.stringify(m));
  } catch { /* storage unavailable: assignment lasts for this session only */ }
}
