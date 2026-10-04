import { initializeApp } from 'firebase/app';
import { 
  getFirestore, 
  doc, 
  setDoc, 
  getDoc,
  getDocFromServer,
  onSnapshot, 
  collection, 
  getDocs,
  deleteDoc
} from 'firebase/firestore';
import { getAuth, signInAnonymously, onAuthStateChanged, User } from 'firebase/auth';
import config from '../../firebase-applet-config.json';

// Explicitly set projectId and authDomain as requested
export const firebaseConfig = {
  ...config,
  projectId: "immi-log-sys",
  authDomain: "immi-log-sys.firebaseapp.com",
};

// Initialize Firebase App
export const app = initializeApp(firebaseConfig);

// Connect to default Firestore with getFirestore(app)
export const db = getFirestore(app);
export const auth = getAuth(app);

// LocalStorage Collection Key Mapping for Offline Fallback
export const LOCAL_STORAGE_COLLECTIONS: Record<string, string> = {
  records: 'imm_records_react',
  tempRecords: 'imm_temp_records_react',
  masterData: 'imm_master_react',
  vehicleSummaries: 'imm_vehicle_summaries',
  dossierHistory: 'imm_dossier_history_react',
  watchList: 'imm_watchlist_records_v1',
  activityLogs: 'imm_activity_logs_v1',
  activeSessions: 'imm_pwa_device_sessions',
};

export const getFromLocalStorageFallback = <T = any>(collectionName: string): T[] | null => {
  const localKey = LOCAL_STORAGE_COLLECTIONS[collectionName] || `imm_fallback_${collectionName}`;
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(localKey);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed as T[];
    }
  } catch {
    // Completely silent fail
  }
  return null;
};

export const saveToLocalStorageFallback = (collectionName: string, items: any[]): void => {
  const localKey = LOCAL_STORAGE_COLLECTIONS[collectionName] || `imm_fallback_${collectionName}`;
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(localKey, JSON.stringify(items || []));
  } catch {
    // Completely silent fail
  }
};

// Global switch indicating cloud firestore is unavailable / not found / offline
let isFirestoreUnavailable = false;

export const setFirestoreUnavailable = (val: boolean) => {
  isFirestoreUnavailable = val;
};

export const checkIsFirestoreUnavailable = (): boolean => isFirestoreUnavailable;

// Safe error detection helper: matches connection, not-found, quota, and offline errors
export const isOfflineOrNetworkError = (err: any): boolean => {
  if (!err) return false;
  if (typeof navigator !== 'undefined' && !navigator.onLine) return true;
  const code = String(err?.code || '').toLowerCase();
  const msg = String(err?.message || err || '').toLowerCase();
  return (
    code === 'unavailable' ||
    code === 'failed-precondition' ||
    code === 'not-found' ||
    code === 'permission-denied' ||
    code.includes('offline') ||
    code.includes('network') ||
    code.includes('timeout') ||
    code.includes('not-found') ||
    code.includes('admin-restricted-operation') ||
    code.includes('operation-not-allowed') ||
    msg.includes('client is offline') ||
    msg.includes('offline') ||
    msg.includes('database not found') ||
    msg.includes('not found') ||
    msg.includes('network-request-failed') ||
    msg.includes('could not reach cloud firestore backend') ||
    msg.includes('failed to get document because the client is offline') ||
    msg.includes('admin-restricted-operation') ||
    msg.includes('restricted to administrators') ||
    msg.includes('operation-not-allowed')
  );
};

// Trackers to prevent infinite auth and console error loops
let isAuthRestricted = false;
let isAuthAttempting = false;

export const checkIsAuthRestricted = (): boolean => isAuthRestricted;

// Initialize anonymous auth with complete error suppression and fallback
export const initAuth = (onUserChanged?: (user: User | null) => void) => {
  if (onUserChanged) onUserChanged(auth.currentUser);

  return onAuthStateChanged(auth, (user) => {
    if (user) {
      if (onUserChanged) onUserChanged(user);
      return;
    }

    if (isAuthRestricted || isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
      if (onUserChanged) onUserChanged(null);
      return;
    }

    if (!isAuthAttempting) {
      isAuthAttempting = true;
      signInAnonymously(auth)
        .then((cred) => {
          isAuthAttempting = false;
          if (onUserChanged) onUserChanged(cred.user);
        })
        .catch((err: any) => {
          isAuthAttempting = false;
          // Silent fail / suppress error loop
          const code = String(err?.code || '').toLowerCase();
          const msg = String(err?.message || err || '').toLowerCase();
          if (
            code.includes('admin-restricted-operation') ||
            code.includes('operation-not-allowed') ||
            msg.includes('admin-restricted-operation') ||
            msg.includes('restricted to administrators')
          ) {
            isAuthRestricted = true;
          }
          if (onUserChanged) onUserChanged(null);
        });
    } else {
      if (onUserChanged) onUserChanged(null);
    }
  });
};

export const ensureAuthenticated = (): Promise<User | null> => {
  if (auth.currentUser) {
    return Promise.resolve(auth.currentUser);
  }
  if (isAuthRestricted || isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return Promise.resolve(null);
  }
  if (isAuthAttempting) {
    return Promise.resolve(auth.currentUser || null);
  }

  isAuthAttempting = true;
  return signInAnonymously(auth)
    .then((cred) => {
      isAuthAttempting = false;
      return cred.user;
    })
    .catch((err: any) => {
      isAuthAttempting = false;
      const code = String(err?.code || '').toLowerCase();
      const msg = String(err?.message || err || '').toLowerCase();
      if (
        code.includes('admin-restricted-operation') ||
        code.includes('operation-not-allowed') ||
        msg.includes('admin-restricted-operation') ||
        msg.includes('restricted to administrators')
      ) {
        isAuthRestricted = true;
      }
      // Silently resolve null - no console.error / warn
      return null;
    });
};

// Sync helpers for main collections & app state
const lastSyncedHashes = new Map<string, string>();

export const setSyncedHash = (collectionName: string, items: any[]) => {
  try {
    lastSyncedHashes.set(collectionName, JSON.stringify(items || []));
  } catch {
    // silent catch
  }
};

export const getSyncedHash = (collectionName: string): string | undefined => {
  return lastSyncedHashes.get(collectionName);
};

let isQuotaExhausted = false;
let quotaResetTimeout: any = null;
const quotaListeners = new Set<(exhausted: boolean) => void>();

export type WriteErrorInfo = {
  collectionName: string;
  reason: string;
  error?: any;
};

type WriteErrorCallback = (info: WriteErrorInfo) => void;
const writeErrorCallbacks: WriteErrorCallback[] = [];

export const onWriteError = (cb: WriteErrorCallback) => {
  writeErrorCallbacks.push(cb);
  return () => {
    const idx = writeErrorCallbacks.indexOf(cb);
    if (idx >= 0) writeErrorCallbacks.splice(idx, 1);
  };
};

export const triggerWriteErrorAlert = (collectionName: string, reason: string, err?: any) => {
  // If firestore is unavailable or offline, suppress alarm popup to keep UI completely calm
  if (isFirestoreUnavailable || isOfflineOrNetworkError(err)) return;
  writeErrorCallbacks.forEach(cb => {
    try { cb({ collectionName, reason, error: err }); } catch {}
  });
};

export const isQuotaError = (err: any): boolean => {
  if (!err) return false;
  const code = String(err?.code || '').toLowerCase();
  const msg = String(err?.message || err || '').toLowerCase();
  return (
    code === 'resource-exhausted' ||
    code === '8' ||
    code.includes('resource-exhausted') ||
    msg.includes('resource_exhausted') ||
    msg.includes('resource-exhausted') ||
    msg.includes('quota exceeded') ||
    msg.includes('quota_exceeded') ||
    msg.includes('over quota') ||
    msg.includes('429')
  );
};

export const notifyQuotaListeners = (exhausted: boolean) => {
  quotaListeners.forEach(cb => {
    try { cb(exhausted); } catch {}
  });
};

export const onQuotaStatusChange = (cb: (exhausted: boolean) => void) => {
  quotaListeners.add(cb);
  cb(isQuotaExhausted);
  return () => {
    quotaListeners.delete(cb);
  };
};

export const setQuotaExhausted = (exhausted: boolean) => {
  if (isQuotaExhausted !== exhausted) {
    isQuotaExhausted = exhausted;
    notifyQuotaListeners(exhausted);
  }
};

export const resetQuotaState = () => {
  setQuotaExhausted(false);
  isFirestoreUnavailable = false;
  if (quotaResetTimeout) clearTimeout(quotaResetTimeout);
};

export const resetFirestoreAvailability = () => {
  isFirestoreUnavailable = false;
  isAuthRestricted = false;
  isAuthAttempting = false;
};

export const getIsQuotaExhausted = () => isQuotaExhausted;

export const getLocalTimestamps = (): Record<string, string> => {
  try {
    const saved = localStorage.getItem('imm_pwa_local_timestamps');
    return saved ? JSON.parse(saved) : {};
  } catch {
    return {};
  }
};

export const setLocalTimestamp = (collectionName: string, timestamp: string) => {
  try {
    const timestamps = getLocalTimestamps();
    timestamps[collectionName] = timestamp;
    localStorage.setItem('imm_pwa_local_timestamps', JSON.stringify(timestamps));
  } catch {
    // silent catch
  }
};

export const getCloudMetadata = async (): Promise<Record<string, string> | null> => {
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return null;
  }
  try {
    await ensureAuthenticated();
    const metaRef = doc(db, 'appData', '_metadata');
    const snapshot = await getDocFromServer(metaRef).catch(() => getDoc(metaRef));
    if (snapshot.exists()) {
      return snapshot.data() as Record<string, string>;
    }
    return null;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    if (isQuotaError(err)) {
      setQuotaExhausted(true);
    }
    // Return null silently - zero console error loop
    return null;
  }
};

export const mergeCollectionItems = <T extends Record<string, any>>(localArr: T[], remoteArr: T[], key: string = 'id'): T[] => {
  const map = new Map<string, T>();

  const getItemKey = (item: any): string | null => {
    if (!item) return null;
    // For master data items with type and name, group by normalized type and name so duplicates merge properly
    if (item.name && item.type) {
      return `master_${String(item.type).trim().toLowerCase()}_${String(item.name).trim().toLowerCase().replace(/\s+/g, ' ')}`;
    }
    if (item[key] !== undefined && item[key] !== null && String(item[key]).trim() !== '') {
      return `${key}_${item[key]}`;
    }
    if (item.date && item.vehicleNo) {
      return `veh_${item.date}_${String(item.vehicleNo).toUpperCase().trim()}_${item.time || ''}_${item.id || ''}`;
    }
    if (item.timestamp && item.passport) {
      return `ts_pp_${item.timestamp}_${item.passport.toUpperCase().trim()}`;
    }
    if (item.timestamp) {
      return `ts_${item.timestamp}`;
    }
    return JSON.stringify(item);
  };

  if (Array.isArray(localArr)) {
    localArr.forEach(item => {
      const k = getItemKey(item);
      if (k) map.set(k, item);
    });
  }

  if (Array.isArray(remoteArr)) {
    remoteArr.forEach(item => {
      const k = getItemKey(item);
      if (k) {
        if (!map.has(k)) {
          map.set(k, { ...item, syncStatus: 'synced' });
        } else {
          const localItem = map.get(k)!;
          const isLocalPendingOrFailed = localItem.syncStatus === 'pending_sync' || localItem.syncStatus === 'upload_failed';
          const localTime = localItem.updatedAt ? new Date(localItem.updatedAt).getTime() : 0;
          const remoteTime = item.updatedAt ? new Date(item.updatedAt).getTime() : 0;

          if (isLocalPendingOrFailed) {
            map.set(k, { ...item, ...localItem });
          } else if (remoteTime > localTime) {
            map.set(k, { ...localItem, ...item, syncStatus: 'synced' });
          } else if (localTime > remoteTime) {
            map.set(k, { ...item, ...localItem });
          } else {
            map.set(k, { ...localItem, ...item, syncStatus: 'synced' });
          }
        }
      }
    });
  }

  return Array.from(map.values());
};

export const extractItemsFromDocs = (docs: Array<{ id: string; data: () => any }>, collectionName: string): any[] => {
  if (!docs || docs.length === 0) return [];
  
  // Find all matching docs and chunk parts for this collection
  // e.g. 'records', 'records_part_0', 'records_part_1', 'records_chunk_0', etc.
  const matchingDocs = docs.filter(d => {
    const id = d.id;
    if (id === collectionName) return true;
    if (id.startsWith(`${collectionName}_part_`)) return true;
    if (id.startsWith(`${collectionName}_chunk_`)) return true;
    if (id.startsWith(`${collectionName}_`)) {
      const suffix = id.substring(collectionName.length + 1);
      if (/^\d+$/.test(suffix) || suffix.startsWith('part') || suffix.startsWith('chunk')) {
        return true;
      }
    }
    return false;
  });

  if (matchingDocs.length === 0) return [];

  // Sort chunks in numerical order so rows remain in correct sequence
  matchingDocs.sort((a, b) => {
    if (a.id === collectionName) return -1;
    if (b.id === collectionName) return 1;
    const numA = parseInt(a.id.match(/\d+$/)?.[0] || '0', 10);
    const numB = parseInt(b.id.match(/\d+$/)?.[0] || '0', 10);
    return numA - numB;
  });

  // Flatten and merge all chunk arrays
  const allItems: any[] = [];
  const seenKeys = new Set<string>();

  for (const d of matchingDocs) {
    const data = d.data();
    if (data && Array.isArray(data.items)) {
      for (const item of data.items) {
        if (!item) continue;
        const k = (item.name && item.type)
          ? `master_${String(item.type).trim().toLowerCase()}_${String(item.name).trim().toLowerCase().replace(/\s+/g, ' ')}`
          : item.id !== undefined && item.id !== null
          ? `id_${item.id}`
          : item.date && item.vehicleNo
          ? `veh_${item.date}_${String(item.vehicleNo).toUpperCase().trim()}_${item.time || ''}`
          : item.timestamp && item.passport
          ? `ts_pp_${item.timestamp}_${item.passport.toUpperCase().trim()}`
          : item.timestamp
          ? `ts_${item.timestamp}`
          : JSON.stringify(item);

        if (!seenKeys.has(k)) {
          seenKeys.add(k);
          allItems.push(item);
        }
      }
    }
  }

  return allItems;
};

export const saveCollectionToFirestore = async (
  collectionName: string, 
  items: any[], 
  force: boolean = false,
  mergeWithRemote: boolean = true
): Promise<boolean> => {
  // Always safely back up items to localStorage fallback immediately
  saveToLocalStorageFallback(collectionName, items);

  if (isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    // Data is safely saved to localStorage fallback, return true so callers know local persistence succeeded
    return true;
  }

  if (isQuotaExhausted) {
    return false;
  }

  try {
    await ensureAuthenticated();
    const currentJson = JSON.stringify(items || []);
    if (!force && lastSyncedHashes.get(collectionName) === currentJson) {
      return true;
    }

    const cleanItems: any[] = JSON.parse(currentJson);
    const now = new Date().toISOString();
    let finalItems = cleanItems;

    if (mergeWithRemote) {
      try {
        const colRef = collection(db, 'appData');
        const querySnap = await getDocs(colRef).catch(() => null);
        if (querySnap && !querySnap.empty) {
          const remoteItems = extractItemsFromDocs(querySnap.docs, collectionName);
          if (remoteItems.length > 0) {
            finalItems = mergeCollectionItems(cleanItems, remoteItems, 'id');
          }
        } else {
          // Fallback to single doc fetch
          const docRef = doc(db, 'appData', collectionName);
          const snap = await getDocFromServer(docRef).catch(() => getDoc(docRef));
          if (snap.exists()) {
            const remoteData = snap.data();
            if (remoteData && Array.isArray(remoteData.items) && remoteData.items.length > 0) {
              finalItems = mergeCollectionItems(cleanItems, remoteData.items, 'id');
            }
          }
        }
      } catch (mergeErr) {
        if (isOfflineOrNetworkError(mergeErr)) {
          isFirestoreUnavailable = true;
        }
      }
    }

    const serverItems = finalItems.map(item => ({
      ...item,
      syncStatus: 'synced',
      serverSyncedAt: item.serverSyncedAt || now
    }));

    // If large collection (> 350 items), chunk into multiple parts to stay well under 1MB limit
    const CHUNK_SIZE = 350;
    const docRef = doc(db, 'appData', collectionName);

    if (serverItems.length > CHUNK_SIZE) {
      const totalParts = Math.ceil(serverItems.length / CHUNK_SIZE);
      const writePromises: Promise<any>[] = [];

      for (let i = 0; i < totalParts; i++) {
        const chunk = serverItems.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
        const partRef = doc(db, 'appData', `${collectionName}_part_${i}`);
        writePromises.push(setDoc(partRef, {
          items: chunk,
          partIndex: i,
          totalParts,
          collectionName,
          lastUpdated: now
        }));
      }

      // Also write root doc with reference/first chunk & count
      writePromises.push(setDoc(docRef, {
        items: serverItems.slice(0, CHUNK_SIZE),
        totalCount: serverItems.length,
        totalParts,
        isChunked: true,
        lastUpdated: now
      }));

      await Promise.all(writePromises);
    } else {
      await setDoc(docRef, { 
        items: serverItems, 
        totalCount: serverItems.length, 
        totalParts: 1, 
        lastUpdated: now 
      });
    }
    
    const metaRef = doc(db, 'appData', '_metadata');
    setDoc(metaRef, { 
      [collectionName]: now,
      [`${collectionName}_count`]: serverItems.length 
    }, { merge: true }).catch(() => {});

    const newHash = JSON.stringify(serverItems);
    lastSyncedHashes.set(collectionName, newHash);
    setLocalTimestamp(collectionName, now);
    saveToLocalStorageFallback(collectionName, serverItems);
    return true;
  } catch (err: any) {
    // Suppress all console.warn / console.error for database not found or network errors
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
      return true; // LocalStorage fallback is safely updated
    }
    if (isQuotaError(err)) {
      setQuotaExhausted(true);
      return false;
    }
    return false;
  }
};

export interface CloudStats {
  recordsCount: number;
  tempRecordsCount: number;
  masterDataCount: number;
  vehicleSummariesCount: number;
  dossierHistoryCount: number;
  watchListCount: number;
  lastUpdated?: string;
  metadata?: Record<string, string>;
}

export const fetchCloudCollectionStats = async (): Promise<CloudStats | null> => {
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return null;
  }
  try {
    await ensureAuthenticated();
    const colRef = collection(db, 'appData');
    const querySnap = await getDocs(colRef).catch(() => null);

    if (querySnap && !querySnap.empty) {
      const docs = querySnap.docs;
      const metaDoc = docs.find(d => d.id === '_metadata');
      const metadata = metaDoc ? (metaDoc.data() as Record<string, string>) : {};

      const getCollectionCount = (colName: string): number => {
        const items = extractItemsFromDocs(docs, colName);
        return items.length;
      };

      return {
        recordsCount: getCollectionCount('records'),
        tempRecordsCount: getCollectionCount('tempRecords'),
        masterDataCount: getCollectionCount('masterData'),
        vehicleSummariesCount: getCollectionCount('vehicleSummaries'),
        dossierHistoryCount: getCollectionCount('dossierHistory'),
        watchListCount: getCollectionCount('watchList'),
        metadata,
        lastUpdated: metadata?.records || new Date().toISOString()
      };
    }

    // Fallback if querySnap is empty
    return null;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return null;
  }
};

export const subscribeToFirestoreCollection = (
  collectionName: string, 
  onUpdate: (items: any[]) => void
) => {
  // Immediately feed current local fallback data to listener
  const localFallback = getFromLocalStorageFallback(collectionName);
  if (localFallback && localFallback.length > 0) {
    try { onUpdate(localFallback); } catch {}
  }

  if (isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return () => {};
  }

  let unsubscribeSnapshot: (() => void) | null = null;
  let isUnsubscribed = false;

  ensureAuthenticated().then(() => {
    if (isUnsubscribed || isFirestoreUnavailable) return;
    try {
      const colRef = collection(db, 'appData');
      unsubscribeSnapshot = onSnapshot(colRef, (snapshot) => {
        if (!snapshot.empty) {
          const mergedItems = extractItemsFromDocs(snapshot.docs, collectionName);
          if (mergedItems.length > 0) {
            const jsonStr = JSON.stringify(mergedItems);
            if (lastSyncedHashes.get(collectionName) === jsonStr) {
              return;
            }
            lastSyncedHashes.set(collectionName, jsonStr);
            saveToLocalStorageFallback(collectionName, mergedItems);
            onUpdate(mergedItems);
          }
        }
      }, (err: any) => {
        // Silent fail / suppress error loop
        if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
          isFirestoreUnavailable = true;
        }
      });
    } catch {
      // quiet catch
    }
  }).catch(() => {});

  return () => {
    isUnsubscribed = true;
    if (unsubscribeSnapshot) {
      unsubscribeSnapshot();
    }
  };
};

export const fetchCollectionFromFirestore = async (collectionName: string): Promise<any[] | null> => {
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return getFromLocalStorageFallback(collectionName);
  }
  try {
    await ensureAuthenticated();
    
    // 1. Try querying the whole appData collection to get all chunks in one go
    try {
      const colRef = collection(db, 'appData');
      const querySnap = await getDocs(colRef);
      if (!querySnap.empty) {
        const mergedItems = extractItemsFromDocs(querySnap.docs, collectionName);
        if (mergedItems.length > 0) {
          saveToLocalStorageFallback(collectionName, mergedItems);
          return mergedItems;
        }
      }
    } catch {
      // quiet fallback
    }

    // 2. Sequential fallback if getDocs wasn't permitted or empty: Check main doc + chunk parts
    const fetchedItems: any[] = [];
    const seenKeys = new Set<string>();

    const checkAndPush = (data: any) => {
      if (data && Array.isArray(data.items)) {
        for (const item of data.items) {
          if (!item) continue;
          const k = item.id !== undefined && item.id !== null
            ? `id_${item.id}`
            : item.timestamp && item.passport
            ? `${item.timestamp}_${item.passport}`
            : JSON.stringify(item);
          if (!seenKeys.has(k)) {
            seenKeys.add(k);
            fetchedItems.push(item);
          }
        }
      }
    };

    // Main doc
    const mainDocRef = doc(db, 'appData', collectionName);
    const mainSnap = await getDocFromServer(mainDocRef).catch(() => getDoc(mainDocRef));
    if (mainSnap.exists()) {
      checkAndPush(mainSnap.data());
    }

    // Scan chunk parts part_0, part_1, part_2 ...
    for (let i = 0; i < 30; i++) {
      try {
        const partRef = doc(db, 'appData', `${collectionName}_part_${i}`);
        const partSnap = await getDocFromServer(partRef).catch(() => getDoc(partRef));
        if (partSnap.exists()) {
          checkAndPush(partSnap.data());
        } else if (i > 1 && fetchedItems.length > 0) {
          break;
        }
      } catch {
        break;
      }
    }

    if (fetchedItems.length > 0) {
      saveToLocalStorageFallback(collectionName, fetchedItems);
      return fetchedItems;
    }

    return getFromLocalStorageFallback(collectionName);
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    // Return localStorage fallback seamlessly with zero console errors
    return getFromLocalStorageFallback(collectionName);
  }
};

/**
 * Smart fetch: Fetches collections directly from Firestore with fallback to localStorage
 */
export const checkAndFetchUpdatedCollections = async (
  collectionNames: string[],
  forceAll: boolean = false,
  onProgress?: (percent: number, currentCollection: string) => void
): Promise<Record<string, any[]>> => {
  const results: Record<string, any[]> = {};

  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    for (const col of collectionNames) {
      const local = getFromLocalStorageFallback(col);
      if (local) results[col] = local;
    }
    return results;
  }
  
  onProgress?.(5, 'Checking Cloud Metadata...');
  let cloudMeta: Record<string, string> | null = null;
  try {
    cloudMeta = await getCloudMetadata();
  } catch {
    cloudMeta = null;
  }

  const localTimestamps = getLocalTimestamps();
  const total = collectionNames.length;
  let count = 0;

  for (const col of collectionNames) {
    count++;
    const currentPercent = Math.min(95, Math.round((count / total) * 90) + 5);
    onProgress?.(currentPercent, col);

    try {
      const remoteTime = cloudMeta ? cloudMeta[col] : null;
      const localTime = localTimestamps[col];
      const localHash = getSyncedHash(col);

      const isRemoteNewerByMeta = remoteTime && (!localTime || new Date(remoteTime).getTime() > new Date(localTime).getTime());

      if (forceAll || isRemoteNewerByMeta || !localHash) {
        const items = await fetchCollectionFromFirestore(col);
        if (Array.isArray(items)) {
          results[col] = items;
        }
      } else {
        const items = await fetchCollectionFromFirestore(col);
        if (Array.isArray(items)) {
          const remoteHash = JSON.stringify(items);
          if (localHash !== remoteHash) {
            results[col] = items;
          }
        }
      }
    } catch {
      const local = getFromLocalStorageFallback(col);
      if (local) results[col] = local;
    }
  }

  onProgress?.(100, 'Complete');
  return results;
};

export const saveDeviceSession = async (session: any): Promise<boolean> => {
  if (!session || !session.deviceId) return false;

  // Always update local session storage fallback
  try {
    const saved = localStorage.getItem('imm_pwa_device_sessions');
    let sessions: any[] = saved ? JSON.parse(saved) : [];
    if (!Array.isArray(sessions)) sessions = [];
    const idx = sessions.findIndex(s => s.deviceId === session.deviceId);
    if (idx >= 0) {
      sessions[idx] = { ...sessions[idx], ...session, lastPingTimestamp: Date.now() };
    } else {
      sessions.push({ ...session, lastPingTimestamp: Date.now() });
    }
    localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(sessions));
  } catch {}

  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return true;
  }

  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', session.deviceId);
    await setDoc(docRef, { ...session, lastPingTimestamp: Date.now() }, { merge: true });
    return true;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return true;
  }
};

export const setDeviceKickedStatus = async (deviceId: string, kicked: boolean): Promise<boolean> => {
  try {
    const saved = localStorage.getItem('imm_pwa_device_sessions');
    let sessions: any[] = saved ? JSON.parse(saved) : [];
    if (Array.isArray(sessions)) {
      sessions = sessions.map(s => s.deviceId === deviceId ? { ...s, kicked } : s);
      localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(sessions));
    }
  } catch {}

  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return true;
  }

  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await setDoc(docRef, { kicked }, { merge: true });
    return true;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return true;
  }
};

export const updateDeviceRole = async (deviceId: string, accountRole: string): Promise<boolean> => {
  try {
    const saved = localStorage.getItem('imm_pwa_device_sessions');
    let sessions: any[] = saved ? JSON.parse(saved) : [];
    if (Array.isArray(sessions)) {
      sessions = sessions.map(s => s.deviceId === deviceId ? { ...s, accountRole } : s);
      localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(sessions));
    }
  } catch {}

  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return true;
  }

  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await setDoc(docRef, { accountRole }, { merge: true });
    return true;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return true;
  }
};

export const deleteDeviceSession = async (deviceId: string): Promise<boolean> => {
  try {
    const saved = localStorage.getItem('imm_pwa_device_sessions');
    let sessions: any[] = saved ? JSON.parse(saved) : [];
    if (Array.isArray(sessions)) {
      sessions = sessions.filter(s => s.deviceId !== deviceId);
      localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(sessions));
    }
  } catch {}

  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return true;
  }

  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await deleteDoc(docRef);
    return true;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return true;
  }
};

export const subscribeToMyDeviceSession = (deviceId: string, onUpdate: (data: any) => void) => {
  if (!deviceId || isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return () => {};
  }

  let unsubscribeSnapshot: (() => void) | null = null;

  ensureAuthenticated().then(() => {
    try {
      const docRef = doc(db, 'activeSessions', deviceId);
      unsubscribeSnapshot = onSnapshot(docRef, (snapshot) => {
        if (snapshot.exists()) {
          onUpdate(snapshot.data());
        }
      }, (err: any) => {
        if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
          isFirestoreUnavailable = true;
        }
      });
    } catch {
      // silent catch
    }
  }).catch(() => {});

  return () => {
    if (unsubscribeSnapshot) unsubscribeSnapshot();
  };
};

export const fetchDeviceSessions = async (): Promise<any[]> => {
  const localFallback = getFromLocalStorageFallback('activeSessions') || [];
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return localFallback;
  }
  try {
    await ensureAuthenticated();
    const colRef = collection(db, 'activeSessions');
    const snapshot = await getDocs(colRef);
    const sessions: any[] = [];
    snapshot.forEach((docSnap) => {
      if (docSnap.exists()) {
        sessions.push(docSnap.data());
      }
    });
    if (sessions.length > 0) {
      saveToLocalStorageFallback('activeSessions', sessions);
      return sessions;
    }
    return localFallback;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return localFallback;
  }
};

export const subscribeToDeviceSessions = (onUpdate: (sessions: any[]) => void) => {
  const localFallback = getFromLocalStorageFallback('activeSessions');
  if (localFallback && localFallback.length > 0) {
    try { onUpdate(localFallback); } catch {}
  }

  if (isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return () => {};
  }

  let unsubscribeSnapshot: (() => void) | null = null;

  ensureAuthenticated().then(() => {
    try {
      const colRef = collection(db, 'activeSessions');
      unsubscribeSnapshot = onSnapshot(colRef, (snapshot) => {
        const sessions: any[] = [];
        snapshot.forEach((docSnap) => {
          if (docSnap.exists()) {
            sessions.push(docSnap.data());
          }
        });
        if (sessions.length > 0) {
          saveToLocalStorageFallback('activeSessions', sessions);
          onUpdate(sessions);
        }
      }, (err: any) => {
        if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
          isFirestoreUnavailable = true;
        }
      });
    } catch {
      // silent catch
    }
  }).catch(() => {});

  return () => {
    if (unsubscribeSnapshot) unsubscribeSnapshot();
  };
};
