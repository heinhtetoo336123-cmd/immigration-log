// Lightweight, robust browser-native IndexedDB Key-Value Store
// Used to bypass strict 5MB localStorage quotas for base64 PDF attachments and large histories.

export class MiniDB {
  private dbName = 'imm_database_kv_store_v1';
  private dbVersion = 1;
  private db: IDBDatabase | null = null;
  private initPromise: Promise<void> | null = null;

  init(): Promise<void> {
    if (this.initPromise) return this.initPromise;

    this.initPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);

      request.onerror = () => {
        console.error("IndexedDB open error:", request.error);
        reject(request.error);
      };

      request.onsuccess = () => {
        this.db = request.result;
        resolve();
      };

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains('keyvalue')) {
          db.createObjectStore('keyvalue');
        }
      };
    });

    return this.initPromise;
  }

  async get(key: string): Promise<any> {
    await this.init();
    return new Promise((resolve, reject) => {
      if (!this.db) {
        resolve(null);
        return;
      }
      try {
        const txn = this.db.transaction('keyvalue', 'readonly');
        const store = txn.objectStore('keyvalue');
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result !== undefined ? req.result : null);
        req.onerror = () => {
          console.error(`IndexedDB reading error for key "${key}":`, req.error);
          resolve(null);
        };
      } catch (e) {
        console.error(`IndexedDB read crash for key "${key}":`, e);
        resolve(null);
      }
    });
  }

  async set(key: string, value: any): Promise<void> {
    await this.init();
    return new Promise((resolve, reject) => {
      if (!this.db) {
        reject(new Error('IndexedDB not initialized'));
        return;
      }
      try {
        const txn = this.db.transaction('keyvalue', 'readwrite');
        const store = txn.objectStore('keyvalue');
        const req = store.put(value, key);
        req.onsuccess = () => resolve();
        req.onerror = () => {
          console.error(`IndexedDB writing error for key "${key}":`, req.error);
          reject(req.error);
        };
      } catch (e) {
        console.error(`IndexedDB write crash for key "${key}":`, e);
        reject(e);
      }
    });
  }

  async remove(key: string): Promise<void> {
    await this.init();
    return new Promise((resolve, reject) => {
      if (!this.db) {
        reject(new Error('IndexedDB not initialized'));
        return;
      }
      try {
        const txn = this.db.transaction('keyvalue', 'readwrite');
        const store = txn.objectStore('keyvalue');
        const req = store.delete(key);
        req.onsuccess = () => resolve();
        req.onerror = () => reject(req.error);
      } catch (e) {
        reject(e);
      }
    });
  }

  async clear(): Promise<void> {
    await this.init();
    return new Promise((resolve) => {
      if (!this.db) {
        resolve();
        return;
      }
      try {
        const txn = this.db.transaction('keyvalue', 'readwrite');
        const store = txn.objectStore('keyvalue');
        const req = store.clear();
        req.onsuccess = () => resolve();
        req.onerror = () => resolve();
      } catch (e) {
        resolve();
      }
    });
  }
}

export const miniDB = new MiniDB();
