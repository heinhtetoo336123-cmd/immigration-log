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

// In-memory cache for metadata to prevent repeated Firestore reads
let cachedMetadata: { data: Record<string, string>; fetchedAt: number } | null = null;
const METADATA_CACHE_TTL_MS = 15000; // 15 seconds cache

export const getCloudMetadata = async (force: boolean = false): Promise<Record<string, string> | null> => {
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return null;
  }

  const now = Date.now();
  if (!force && cachedMetadata && (now - cachedMetadata.fetchedAt < METADATA_CACHE_TTL_MS)) {
    return cachedMetadata.data;
  }

  try {
    await ensureAuthenticated();
    const metaRef = doc(db, 'appData', '_metadata');
    const snapshot = await getDocFromServer(metaRef).catch(() => getDoc(metaRef));
    if (snapshot.exists()) {
      const data = snapshot.data() as Record<string, string>;
      cachedMetadata = { data, fetchedAt: Date.now() };
      return data;
    }
    return null;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    if (isQuotaError(err)) {
      setQuotaExhausted(true);
    }
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
  _deprecatedMerge?: boolean
): Promise<boolean> => {
  // Always safely back up items to localStorage fallback immediately (0ms local save)
  saveToLocalStorageFallback(collectionName, items);

  if (isFirestoreUnavailable || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return true;
  }

  if (isQuotaExhausted) {
    return false;
  }

  try {
    const currentJson = JSON.stringify(items || []);
    // Skip network if hash has not changed and not forced
    if (!force && lastSyncedHashes.get(collectionName) === currentJson) {
      return true;
    }

    await ensureAuthenticated();
    const cleanItems: any[] = JSON.parse(currentJson);
    const now = new Date().toISOString();

    const serverItems = cleanItems.map(item => ({
      ...item,
      syncStatus: 'synced',
      serverSyncedAt: item.serverSyncedAt || now
    }));

    // If large collection (> 350 items), chunk into multiple parts
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
    
    // Update lightweight metadata timestamp (1 small write)
    const metaRef = doc(db, 'appData', '_metadata');
    setDoc(metaRef, { 
      [collectionName]: now,
      [`${collectionName}_count`]: serverItems.length 
    }, { merge: true }).catch(() => {});

    // Invalidate local metadata cache so future checks see latest
    cachedMetadata = null;

    const newHash = JSON.stringify(serverItems);
    lastSyncedHashes.set(collectionName, newHash);
    setLocalTimestamp(collectionName, now);
    saveToLocalStorageFallback(collectionName, serverItems);
    return true;
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
      return true;
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
    const meta = await getCloudMetadata(false);
    if (!meta) return null;

    return {
      recordsCount: parseInt(meta['records_count'] || '0', 10),
      tempRecordsCount: parseInt(meta['tempRecords_count'] || '0', 10),
      masterDataCount: parseInt(meta['masterData_count'] || '0', 10),
      vehicleSummariesCount: parseInt(meta['vehicleSummaries_count'] || '0', 10),
      dossierHistoryCount: parseInt(meta['dossierHistory_count'] || '0', 10),
      watchListCount: parseInt(meta['watchList_count'] || '0', 10),
      metadata: meta,
      lastUpdated: meta.records || new Date().toISOString()
    };
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return null;
  }
};

// Deprecated continuous snapshot listener: Serves local data immediately and returns noop cleanup
export const subscribeToFirestoreCollection = (
  collectionName: string, 
  onUpdate: (items: any[]) => void
) => {
  const localFallback = getFromLocalStorageFallback(collectionName);
  if (localFallback && localFallback.length > 0) {
    try { onUpdate(localFallback); } catch {}
  }
  return () => {};
};

// Single-document direct fetcher (no full collection scans)
export const fetchCollectionFromFirestore = async (collectionName: string): Promise<any[] | null> => {
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return getFromLocalStorageFallback(collectionName);
  }
  try {
    await ensureAuthenticated();
    
    const mainDocRef = doc(db, 'appData', collectionName);
    const mainSnap = await getDocFromServer(mainDocRef).catch(() => getDoc(mainDocRef));
    
    if (!mainSnap.exists()) {
      return getFromLocalStorageFallback(collectionName);
    }

    const mainData = mainSnap.data();
    if (!mainData) return getFromLocalStorageFallback(collectionName);

    let allItems: any[] = Array.isArray(mainData.items) ? [...mainData.items] : [];

    // If chunked and totalParts > 1, fetch only existing chunk documents in parallel
    if (mainData.isChunked && typeof mainData.totalParts === 'number' && mainData.totalParts > 1) {
      const partPromises: Promise<any>[] = [];
      for (let i = 1; i < mainData.totalParts; i++) {
        const partRef = doc(db, 'appData', `${collectionName}_part_${i}`);
        partPromises.push(getDocFromServer(partRef).catch(() => getDoc(partRef)));
      }
      const partSnaps = await Promise.all(partPromises);
      for (const pSnap of partSnaps) {
        if (pSnap && pSnap.exists()) {
          const pData = pSnap.data();
          if (pData && Array.isArray(pData.items)) {
            allItems = allItems.concat(pData.items);
          }
        }
      }
    }

    if (allItems.length > 0) {
      saveToLocalStorageFallback(collectionName, allItems);
      setSyncedHash(collectionName, allItems);
      return allItems;
    }

    return getFromLocalStorageFallback(collectionName);
  } catch (err: any) {
    if (isOfflineOrNetworkError(err) || String(err).toLowerCase().includes('not found')) {
      isFirestoreUnavailable = true;
    }
    return getFromLocalStorageFallback(collectionName);
  }
};

/**
 * Smart metadata-based conditional sync:
 * 1. Reads ONLY 1 metadata document to compare timestamps.
 * 2. Only fetches records if the remote timestamp is strictly newer than local timestamp.
 * 3. Skips unchanged collections completely (0 reads).
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
    cloudMeta = await getCloudMetadata(forceAll);
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

      // STRICT CHECK: Only query Firestore document if remote timestamp is strictly newer or forced
      if (forceAll || isRemoteNewerByMeta || !localHash) {
        const items = await fetchCollectionFromFirestore(col);
        if (Array.isArray(items) && items.length > 0) {
          results[col] = items;
          if (remoteTime) setLocalTimestamp(col, remoteTime);
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

// In-memory throttling for device session heartbeats (10 minutes)
let lastDeviceSessionPingTime = 0;
const SESSION_PING_THROTTLE_MS = 10 * 60 * 1000;

export const saveDeviceSession = async (session: any, force: boolean = false): Promise<boolean> => {
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

  const now = Date.now();
  if (!force && (now - lastDeviceSessionPingTime < SESSION_PING_THROTTLE_MS)) {
    return true; // Throttled to save writes
  }

  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', session.deviceId);
    await setDoc(docRef, { ...session, lastPingTimestamp: Date.now() }, { merge: true });
    lastDeviceSessionPingTime = Date.now();
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

// No real-time subscription loop - reads locally
export const subscribeToMyDeviceSession = (deviceId: string, onUpdate: (data: any) => void) => {
  if (!deviceId) return () => {};
  try {
    const saved = localStorage.getItem('imm_pwa_device_sessions');
    if (saved) {
      const sessions = JSON.parse(saved);
      const mySession = sessions.find((s: any) => s.deviceId === deviceId);
      if (mySession) onUpdate(mySession);
    }
  } catch {}
  return () => {};
};

// Throttled in-memory session fetcher (10 minutes cache)
let cachedDeviceSessions: { data: any[]; fetchedAt: number } | null = null;
const SESSIONS_CACHE_TTL_MS = 10 * 60 * 1000;

export const fetchDeviceSessions = async (force: boolean = false): Promise<any[]> => {
  const localFallback = getFromLocalStorageFallback('activeSessions') || [];
  if (isFirestoreUnavailable || isQuotaExhausted || (typeof navigator !== 'undefined' && !navigator.onLine)) {
    return localFallback;
  }

  const now = Date.now();
  if (!force && cachedDeviceSessions && (now - cachedDeviceSessions.fetchedAt < SESSIONS_CACHE_TTL_MS)) {
    return cachedDeviceSessions.data;
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
      cachedDeviceSessions = { data: sessions, fetchedAt: Date.now() };
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
  return () => {};
};
