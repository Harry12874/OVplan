const DB_NAME = "orchard_valley_planner";
const DB_VERSION = 5;
const STORES = [
  "reps",
  "customers",
  "orders",
  "tasks",
  "settings",
  "schedule_events",
  "one_off_items",
];

let dbPromise = null;
let useLocalStorage = false;

const storageKey = (store) => `ov_${store}`;

const localAdapter = {
  async init() {
    STORES.forEach((store) => {
      if (!localStorage.getItem(storageKey(store))) {
        localStorage.setItem(storageKey(store), JSON.stringify({}));
      }
    });
  },
  async getAll(store) {
    const data = JSON.parse(localStorage.getItem(storageKey(store)) || "{}");
    return Object.values(data);
  },
  async get(store, id) {
    const data = JSON.parse(localStorage.getItem(storageKey(store)) || "{}");
    return data[id] || null;
  },
  async put(store, value) {
    const data = JSON.parse(localStorage.getItem(storageKey(store)) || "{}");
    data[value.id] = value;
    localStorage.setItem(storageKey(store), JSON.stringify(data));
    return value;
  },
  async delete(store, id) {
    const data = JSON.parse(localStorage.getItem(storageKey(store)) || "{}");
    delete data[id];
    localStorage.setItem(storageKey(store), JSON.stringify(data));
  },
  async clear(store) {
    localStorage.setItem(storageKey(store), JSON.stringify({}));
  },
};

const idbAdapter = {
  async init() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = (event) => {
          const db = event.target.result;
          STORES.forEach((store) => {
            if (!db.objectStoreNames.contains(store)) {
              db.createObjectStore(store, { keyPath: "id" });
            }
          });
          if (event.oldVersion < 2) {
            const tx = event.target.transaction;
            const customerStore = tx.objectStore("customers");
            customerStore.openCursor().onsuccess = (cursorEvent) => {
              const cursor = cursorEvent.target.result;
              if (cursor) {
                const value = cursor.value;
                if (value.averageOrderValue === undefined) {
                  value.averageOrderValue = null;
                  cursor.update(value);
                }
                cursor.continue();
              }
            };
          }
          if (event.oldVersion < 3) {
            const tx = event.target.transaction;
            const customerStore = tx.objectStore("customers");
            customerStore.openCursor().onsuccess = (cursorEvent) => {
              const cursor = cursorEvent.target.result;
              if (cursor) {
                const value = cursor.value;
                if (value.averageOrderValue === undefined) {
                  value.averageOrderValue = null;
                }
                if (!value.extraFields) {
                  value.extraFields = {};
                }
                cursor.update(value);
                cursor.continue();
              }
            };
          }
          if (event.oldVersion < 4) {
            const tx = event.target.transaction;
            const customerStore = tx.objectStore("customers");
            customerStore.openCursor().onsuccess = (cursorEvent) => {
              const cursor = cursorEvent.target.result;
              if (cursor) {
                const value = cursor.value;
                if (!value.schedule) {
                  value.schedule = {
                    mode: null,
                    frequency: null,
                    orderDay1: null,
                    deliverDay1: null,
                    packDay1: null,
                    isBiWeeklySecondRun: false,
                    orderDay2: null,
                    deliverDay2: null,
                    packDay2: null,
                    customerOrderDays: [],
                    deliverDays: [],
                    packDays: [],
                    anchorDate: null,
                  };
                }
                cursor.update(value);
                cursor.continue();
              }
            };
          }
          if (event.oldVersion < 5) {
            if (!db.objectStoreNames.contains("schedule_events")) {
              db.createObjectStore("schedule_events", { keyPath: "id" });
            }
            if (!db.objectStoreNames.contains("one_off_items")) {
              db.createObjectStore("one_off_items", { keyPath: "id" });
            }
            const tx = event.target.transaction;
            const customerStore = tx.objectStore("customers");
            customerStore.openCursor().onsuccess = (cursorEvent) => {
              const cursor = cursorEvent.target.result;
              if (cursor) {
                const value = cursor.value;
                value.customerNotes = value.customerNotes || "";
                if (value.schedule) {
                  value.schedule.packDay1 = value.schedule.packDay1 ?? null;
                  value.schedule.packDay2 = value.schedule.packDay2 ?? null;
                  value.schedule.packDays = value.schedule.packDays || [];
                }
                cursor.update(value);
                cursor.continue();
              }
            };
          }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }
    return dbPromise;
  },
  async getAll(store) {
    const db = await idbAdapter.init();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readonly");
      const req = tx.objectStore(store).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  },
  async get(store, id) {
    const db = await idbAdapter.init();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readonly");
      const req = tx.objectStore(store).get(id);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  },
  async put(store, value) {
    const db = await idbAdapter.init();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      const req = tx.objectStore(store).put(value);
      req.onsuccess = () => resolve(value);
      req.onerror = () => reject(req.error);
    });
  },
  async delete(store, id) {
    const db = await idbAdapter.init();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      const req = tx.objectStore(store).delete(id);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
  async clear(store) {
    const db = await idbAdapter.init();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      const req = tx.objectStore(store).clear();
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
};

async function initDB() {
  try {
    if (!("indexedDB" in window)) {
      useLocalStorage = true;
      await localAdapter.init();
      return "localStorage";
    }
    await idbAdapter.init();
    return "indexedDB";
  } catch (error) {
    console.warn("IndexedDB unavailable, falling back to localStorage", error);
    useLocalStorage = true;
    await localAdapter.init();
    return "localStorage";
  }
}

const adapter = () => (useLocalStorage ? localAdapter : idbAdapter);

async function getAll(store) {
  return adapter().getAll(store);
}

async function get(store, id) {
  return adapter().get(store, id);
}

async function put(store, value) {
  return adapter().put(store, value);
}

async function deleteItem(store, id) {
  return adapter().delete(store, id);
}

async function clearStore(store) {
  return adapter().clear(store);
}

async function bulkPut(store, items) {
  for (const item of items) {
    await put(store, item);
  }
}

export { initDB, getAll, get, put, deleteItem, clearStore, bulkPut };
