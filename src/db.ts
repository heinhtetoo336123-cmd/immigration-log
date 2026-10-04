// Lightweight, robust browser-native IndexedDB Key-Value Store
// Used to bypass strict 5MB localStorage quotas for large datasets and PDF attachments.

export class MiniDB {
  private dbName = 'imm_database_kv_store_v1';
  private dbVersion = 1;
  private db: IDBDatabase | null = null;
  private openingPromise: Promise<IDBDatabase> | null = null;
  private queue: Promise<any> = Promise.resolve();

  async init(): Promise<void> {
    await this.getDB();
  }

  private async getDB(): Promise<IDBDatabase> {
    if (this.db) {
      return this.db;
    }

    if (this.openingPromise) {
      return this.openingPromise;
    }

    this.openingPromise = new Promise<IDBDatabase>((resolve, reject) => {
      try {
        if (typeof indexedDB === 'undefined') {
          reject(new Error('IndexedDB not supported in this environment'));
          return;
        }

        const request = indexedDB.open(this.dbName, this.dbVersion);

        request.onerror = () => {
          this.openingPromise = null;
          console.warn("IndexedDB open error:", request.error);
          reject(request.error || new Error('Failed to open IndexedDB'));
        };

        request.onblocked = () => {
          console.warn("IndexedDB open blocked by another tab");
        };

        request.onsuccess = () => {
          this.db = request.result;
          this.openingPromise = null;

          this.db.onversionchange = () => {
            if (this.db) {
              this.db.close();
              this.db = null;
            }
          };

          this.db.onclose = () => {
            this.db = null;
          };

          resolve(this.db);
        };

        request.onupgradeneeded = (event: IDBVersionChangeEvent) => {
          const targetDB = (event.target as IDBOpenDBRequest).result;
          if (!targetDB.objectStoreNames.contains('keyvalue')) {
            targetDB.createObjectStore('keyvalue');
          }
        };
      } catch (err) {
        this.openingPromise = null;
        reject(err);
      }
    });

    return this.openingPromise;
  }

  async get(key: string): Promise<any> {
    try {
      const db = await this.getDB();
      return await new Promise((resolve) => {
        try {
          const txn = db.transaction('keyvalue', 'readonly');
          const store = txn.objectStore('keyvalue');
          const req = store.get(key);
          req.onsuccess = () => resolve(req.result !== undefined ? req.result : null);
          req.onerror = () => resolve(null);
          txn.onerror = () => resolve(null);
          txn.onabort = () => resolve(null);
        } catch (e) {
          resolve(null);
        }
      });
    } catch (e) {
      return null;
    }
  }

  async set(key: string, value: any): Promise<void> {
    // Sequence writes in an async queue to avoid transaction collision
    this.queue = this.queue.then(async () => {
      try {
        const db = await this.getDB();
        await new Promise<void>((resolve, reject) => {
          try {
            const txn = db.transaction('keyvalue', 'readwrite');
            const store = txn.objectStore('keyvalue');
            
            // Safe JSON serialization clone to avoid non-serializable object errors
            let clone = value;
            try {
              if (value !== null && typeof value === 'object') {
                clone = JSON.parse(JSON.stringify(value));
              }
            } catch {}

            const req = store.put(clone, key);
            req.onsuccess = () => resolve();
            req.onerror = () => {
              reject(req.error || new Error(`IndexedDB put failed for key "${key}"`));
            };
            txn.onerror = () => {
              reject(txn.error || new Error(`IndexedDB transaction failed for key "${key}"`));
            };
            txn.onabort = () => {
              reject(new Error(`IndexedDB transaction aborted for key "${key}"`));
            };
          } catch (e) {
            reject(e);
          }
        });
      } catch (err) {
        // Fallback: log warning and resolve gracefully without throwing uncaught app errors
        console.warn(`IndexedDB write fallback for key "${key}":`, err);
      }
    }).catch(() => {});

    return this.queue;
  }

  async remove(key: string): Promise<void> {
    this.queue = this.queue.then(async () => {
      try {
        const db = await this.getDB();
        await new Promise<void>((resolve) => {
          try {
            const txn = db.transaction('keyvalue', 'readwrite');
            const store = txn.objectStore('keyvalue');
            const req = store.delete(key);
            req.onsuccess = () => resolve();
            req.onerror = () => resolve();
            txn.onerror = () => resolve();
            txn.onabort = () => resolve();
          } catch (e) {
            resolve();
          }
        });
      } catch (err) {
        console.warn(`IndexedDB delete fallback for key "${key}":`, err);
      }
    }).catch(() => {});

    return this.queue;
  }

  async clear(): Promise<void> {
    this.queue = this.queue.then(async () => {
      try {
        const db = await this.getDB();
        await new Promise<void>((resolve) => {
          try {
            const txn = db.transaction('keyvalue', 'readwrite');
            const store = txn.objectStore('keyvalue');
            const req = store.clear();
            req.onsuccess = () => resolve();
            req.onerror = () => resolve();
            txn.onerror = () => resolve();
            txn.onabort = () => resolve();
          } catch (e) {
            resolve();
          }
        });
      } catch (err) {
        console.warn("IndexedDB clear fallback:", err);
      }
    }).catch(() => {});

    return this.queue;
  }
}

export const miniDB = new MiniDB();
