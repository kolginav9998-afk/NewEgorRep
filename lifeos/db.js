// IndexedDB: одно хранилище "events". Реальное удаление, без soft-delete.
const DB_NAME = 'lifeos';
const STORE = 'events';

function open() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx(mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode);
    const out = fn(t.objectStore(STORE));
    t.oncomplete = () => { db.close(); resolve(out?.result); };
    t.onerror = () => { db.close(); reject(t.error); };
    t.onabort = () => { db.close(); reject(t.error); };
  });
}

export const getAll = () => tx('readonly', s => s.getAll());
export const put = ev => tx('readwrite', s => s.put(ev));
export const remove = id => tx('readwrite', s => s.delete(id));
export const putMany = evs => tx('readwrite', s => { evs.forEach(e => s.put(e)); });
