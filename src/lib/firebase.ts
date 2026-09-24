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
  writeBatch,
  deleteDoc
} from 'firebase/firestore';
import { getAuth, signInAnonymously, onAuthStateChanged, User } from 'firebase/auth';
import config from '../../firebase-applet-config.json';

export const app = initializeApp(config);

const databaseId = config.firestoreDatabaseId || '(default)';
export const db = getFirestore(app, databaseId);
export const auth = getAuth(app);

// Initialize anonymous auth automatically if enabled, fallback gracefully if restricted
export const initAuth = (onUserChanged?: (user: User | null) => void) => {
  if (onUserChanged) onUserChanged(auth.currentUser);
  return onAuthStateChanged(auth, (user) => {
    if (!user) {
      signInAnonymously(auth).catch((err) => {
        console.warn("signInAnonymously failed:", err);
      });
    }
    if (onUserChanged) onUserChanged(user);
  });
};

export const ensureAuthenticated = (): Promise<User | null> => {
  return new Promise((resolve) => {
    if (auth.currentUser) {
      resolve(auth.currentUser);
      return;
    }

    const unsubscribe = onAuthStateChanged(auth, (user) => {
      if (user) {
        unsubscribe();
        resolve(user);
      }
    });

    signInAnonymously(auth).catch((err) => {
      console.warn("signInAnonymously error:", err);
      resolve(null);
    });
  });
};

// Sync helpers for main collections & app state
const lastSyncedHashes = new Map<string, string>();

export const setSyncedHash = (collectionName: string, items: any[]) => {
  try {
    lastSyncedHashes.set(collectionName, JSON.stringify(items || []));
  } catch (e) {
    console.error("Failed to set synced hash:", e);
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
  writeErrorCallbacks.forEach(cb => {
    try { cb({ collectionName, reason, error: err }); } catch (e) { console.error(e); }
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
    try { cb(exhausted); } catch (e) { console.error(e); }
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
  if (quotaResetTimeout) clearTimeout(quotaResetTimeout);
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
  } catch (e) {
    console.error("Failed to save local timestamp:", e);
  }
};

export const getCloudMetadata = async () => {
  if (isQuotaExhausted) return null;
  try {
    await ensureAuthenticated();
    const metaRef = doc(db, 'appData', '_metadata');
    const snapshot = await getDocFromServer(metaRef).catch(() => getDoc(metaRef));
    if (snapshot.exists()) {
      return snapshot.data() as Record<string, string>;
    }
    return null;
  } catch (err: any) {
    if (isQuotaError(err)) {
      setQuotaExhausted(true);
      console.warn("Firestore metadata fetch quota exceeded.");
    }
    return null;
  }
};

export const mergeCollectionItems = <T extends Record<string, any>>(localArr: T[], remoteArr: T[], key: string = 'id'): T[] => {
  const map = new Map<string, T>();

  const getItemKey = (item: any): string | null => {
    if (!item) return null;
    if (item.date && item.vehicleNo) {
      return `veh_${item.date}_${String(item.vehicleNo).toUpperCase().trim()}`;
    }
    if (item.name && item.type) {
      return `master_${String(item.type).trim().toLowerCase()}_${String(item.name).trim().toLowerCase().replace(/\s+/g, ' ')}`;
    }
    if (item[key] !== undefined && item[key] !== null && String(item[key]).trim() !== '') {
      return `${key}_${item[key]}`;
    }
    if (item.timestamp && item.passport) {
      return `ts_pp_${item.timestamp}_${item.passport.toUpperCase().trim()}`;
    }
    if (item.timestamp) {
      return `ts_${item.timestamp}`;
    }
    return JSON.stringify(item);
  };

  // 1. Put local items in map first
  if (Array.isArray(localArr)) {
    localArr.forEach(item => {
      const k = getItemKey(item);
      if (k) map.set(k, item);
    });
  }

  // 2. Merge remote items into map
  if (Array.isArray(remoteArr)) {
    remoteArr.forEach(item => {
      const k = getItemKey(item);
      if (k) {
        if (!map.has(k)) {
          // Item on server is confirmed synced
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

export const saveCollectionToFirestore = async (
  collectionName: string, 
  items: any[], 
  force: boolean = false,
  mergeWithRemote: boolean = true
) => {
  if (isQuotaExhausted) {
    triggerWriteErrorAlert(collectionName, "Write Quota Limit Reached");
    return false;
  }

  try {
    await ensureAuthenticated();
    const currentJson = JSON.stringify(items || []);
    // STRICT check: if data has NOT changed and not forcing, return true
    if (!force && lastSyncedHashes.get(collectionName) === currentJson) {
      return true;
    }

    const cleanItems: any[] = JSON.parse(currentJson);
    const now = new Date().toISOString();
    const docRef = doc(db, 'appData', collectionName);

    let finalItems = cleanItems;

    // Fail-safe merge: fetch existing remote collection so entries from other devices (e.g. Phone or PC) are NEVER wiped out
    if (mergeWithRemote) {
      try {
        const snap = await getDocFromServer(docRef).catch(() => getDoc(docRef));
        if (snap.exists()) {
          const remoteData = snap.data();
          if (remoteData && Array.isArray(remoteData.items) && remoteData.items.length > 0) {
            finalItems = mergeCollectionItems(cleanItems, remoteData.items, 'id');
          }
        }
      } catch (mergeErr) {
        console.warn(`Notice reading remote collection for merge before saving ${collectionName}:`, mergeErr);
      }
    }

    // Normalize all items stored on Cloud Server as 'synced' with server timestamp
    const serverItems = finalItems.map(item => ({
      ...item,
      syncStatus: 'synced',
      serverSyncedAt: item.serverSyncedAt || now
    }));

    await setDoc(docRef, { items: serverItems, lastUpdated: now });
    
    // Update metadata document asynchronously
    const metaRef = doc(db, 'appData', '_metadata');
    setDoc(metaRef, { [collectionName]: now }, { merge: true }).catch((err: any) => {
      if (isQuotaError(err) || String(err).toLowerCase().includes('quota')) {
        setQuotaExhausted(true);
      }
    });

    const newHash = JSON.stringify(serverItems);
    lastSyncedHashes.set(collectionName, newHash);
    setLocalTimestamp(collectionName, now);
    return true;
  } catch (err: any) {
    const quotaExceeded = isQuotaError(err);
    if (quotaExceeded) {
      setQuotaExhausted(true);
      if (quotaResetTimeout) clearTimeout(quotaResetTimeout);
      quotaResetTimeout = setTimeout(() => { setQuotaExhausted(false); }, 10 * 60 * 1000);
    }
    const reason = quotaExceeded ? "Write Quota Limit Exceeded" : (err?.message || "Server / Network Write Error");
    triggerWriteErrorAlert(collectionName, reason, err);
    console.warn(`Firestore write operation failed for ${collectionName}. Data is safely stored in local IndexedDB.`, err);
    return false;
  }
};

export interface CloudStats {
  recordsCount: number;
  tempRecordsCount: number;
  masterDataCount: number;
  vehicleSummariesCount: number;
  checkingHistoryCount: number;
  watchListCount: number;
  lastUpdated?: string;
  metadata?: Record<string, string>;
}

export const fetchCloudCollectionStats = async (): Promise<CloudStats | null> => {
  if (isQuotaExhausted) return null;
  try {
    await ensureAuthenticated();
    const metaRef = doc(db, 'appData', '_metadata');
    const metaSnap = await getDocFromServer(metaRef).catch(() => getDoc(metaRef));
    const metadata = metaSnap.exists() ? (metaSnap.data() as Record<string, string>) : {};

    const [recSnap, tempSnap, masterSnap, checkSnap] = await Promise.all([
      getDocFromServer(doc(db, 'appData', 'records')).catch(() => getDoc(doc(db, 'appData', 'records'))),
      getDocFromServer(doc(db, 'appData', 'tempRecords')).catch(() => getDoc(doc(db, 'appData', 'tempRecords'))),
      getDocFromServer(doc(db, 'appData', 'masterData')).catch(() => getDoc(doc(db, 'appData', 'masterData'))),
      getDocFromServer(doc(db, 'appData', 'checkingHistory')).catch(() => getDoc(doc(db, 'appData', 'checkingHistory')))
    ]);

    const getCount = (snap: any) => (snap.exists() && Array.isArray(snap.data()?.items) ? snap.data().items.length : 0);

    return {
      recordsCount: getCount(recSnap),
      tempRecordsCount: getCount(tempSnap),
      masterDataCount: getCount(masterSnap),
      vehicleSummariesCount: 0,
      checkingHistoryCount: getCount(checkSnap),
      watchListCount: 0,
      metadata,
      lastUpdated: metadata?.records || new Date().toISOString()
    };
  } catch (e) {
    console.warn("fetchCloudCollectionStats error:", e);
    return null;
  }
};

export const subscribeToFirestoreCollection = (
  collectionName: string, 
  onUpdate: (items: any[]) => void
) => {
  let unsubscribeSnapshot: (() => void) | null = null;
  let isUnsubscribed = false;

  ensureAuthenticated().then(() => {
    if (isUnsubscribed) return;
    const docRef = doc(db, 'appData', collectionName);
    unsubscribeSnapshot = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data && Array.isArray(data.items)) {
          const jsonStr = JSON.stringify(data.items);
          // STRICT ECHO PREVENTION: Skip calling onUpdate if local synced hash matches remote snapshot
          if (lastSyncedHashes.get(collectionName) === jsonStr) {
            return;
          }
          lastSyncedHashes.set(collectionName, jsonStr);
          onUpdate(data.items);
        }
      }
    }, (err: any) => {
      if (isQuotaError(err)) {
        setQuotaExhausted(true);
        console.warn(`Firestore snapshot quota limit reached for ${collectionName}. Using local cache/DB.`);
      } else {
        console.warn(`Firestore snapshot notice for ${collectionName}:`, err?.message || err);
      }
    });
  });

  return () => {
    isUnsubscribed = true;
    if (unsubscribeSnapshot) {
      unsubscribeSnapshot();
    }
  };
};

export const fetchCollectionFromFirestore = async (collectionName: string) => {
  if (isQuotaExhausted) return null;
  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'appData', collectionName);
    const snapshot = await getDocFromServer(docRef).catch(() => getDoc(docRef));
    if (snapshot.exists()) {
      const data = snapshot.data();
      if (data && Array.isArray(data.items)) {
        if (data.lastUpdated) {
          setLocalTimestamp(collectionName, data.lastUpdated);
        }
        return data.items;
      }
    }
    return null;
  } catch (err: any) {
    if (isQuotaError(err)) {
      setQuotaExhausted(true);
      console.warn(`Firestore fetch quota limit reached for ${collectionName}. Using local DB.`);
    } else {
      console.warn(`Notice fetching ${collectionName} from Firestore:`, err?.message || err);
    }
    return null;
  }
};

/**
 * Smart fetch: Fetches collections directly from Firestore and downloads items
 * if hash differs or remote timestamp is updated or forceAll is set to true.
 */
export const checkAndFetchUpdatedCollections = async (
  collectionNames: string[],
  forceAll: boolean = false,
  onProgress?: (percent: number, currentCollection: string) => void
) => {
  if (isQuotaExhausted) return {};
  
  onProgress?.(5, 'Checking Cloud Metadata...');
  const cloudMeta = await getCloudMetadata();
  const localTimestamps = getLocalTimestamps();
  const results: Record<string, any[]> = {};

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

      // Check if remote is newer from metadata
      const isRemoteNewerByMeta = remoteTime && (!localTime || new Date(remoteTime).getTime() > new Date(localTime).getTime());

      // If forceAll OR remote is newer by metadata OR localHash is missing/not set, fetch collection directly
      if (forceAll || isRemoteNewerByMeta || !localHash) {
        const items = await fetchCollectionFromFirestore(col);
        if (Array.isArray(items)) {
          results[col] = items;
        }
      } else {
        // Fallback check: fetch document directly to ensure no metadata desync
        const items = await fetchCollectionFromFirestore(col);
        if (Array.isArray(items)) {
          const remoteHash = JSON.stringify(items);
          if (localHash !== remoteHash) {
            results[col] = items;
          }
        }
      }
    } catch (err) {
      console.warn(`Error checking collection ${col}:`, err);
    }
  }

  onProgress?.(100, 'Complete');
  return results;
};

export const saveDeviceSession = async (session: any) => {
  if (isQuotaExhausted) return false;
  try {
    await ensureAuthenticated();
    if (!session || !session.deviceId) return false;
    const docRef = doc(db, 'activeSessions', session.deviceId);
    await setDoc(docRef, { ...session, lastPingTimestamp: Date.now() }, { merge: true });
    return true;
  } catch (err) {
    console.warn("Failed to save device session:", err);
    return false;
  }
};

export const setDeviceKickedStatus = async (deviceId: string, kicked: boolean) => {
  if (isQuotaExhausted) return false;
  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await setDoc(docRef, { kicked }, { merge: true });
    return true;
  } catch (err) {
    console.warn("Failed to update kick status:", err);
    return false;
  }
};

export const updateDeviceRole = async (deviceId: string, accountRole: string) => {
  if (isQuotaExhausted) return false;
  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await setDoc(docRef, { accountRole }, { merge: true });
    return true;
  } catch (err) {
    console.warn("Failed to update device role:", err);
    return false;
  }
};

export const deleteDeviceSession = async (deviceId: string) => {
  if (isQuotaExhausted) return false;
  try {
    await ensureAuthenticated();
    const docRef = doc(db, 'activeSessions', deviceId);
    await deleteDoc(docRef);
    return true;
  } catch (err) {
    console.warn("Failed to delete device session:", err);
    return false;
  }
};

export const subscribeToMyDeviceSession = (deviceId: string, onUpdate: (data: any) => void) => {
  let unsubscribeSnapshot: (() => void) | null = null;
  if (!deviceId) return () => {};
  ensureAuthenticated().then(() => {
    const docRef = doc(db, 'activeSessions', deviceId);
    unsubscribeSnapshot = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        onUpdate(snapshot.data());
      }
    }, (err) => {
      console.warn("subscribeToMyDeviceSession error:", err);
    });
  });

  return () => {
    if (unsubscribeSnapshot) unsubscribeSnapshot();
  };
};

export const fetchDeviceSessions = async (): Promise<any[]> => {
  if (isQuotaExhausted) return [];
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
    return sessions;
  } catch (err) {
    console.warn("fetchDeviceSessions error:", err);
    return [];
  }
};

export const subscribeToDeviceSessions = (onUpdate: (sessions: any[]) => void) => {
  let unsubscribeSnapshot: (() => void) | null = null;
  ensureAuthenticated().then(() => {
    const colRef = collection(db, 'activeSessions');
    unsubscribeSnapshot = onSnapshot(colRef, (snapshot) => {
      const sessions: any[] = [];
      snapshot.forEach((docSnap) => {
        if (docSnap.exists()) {
          sessions.push(docSnap.data());
        }
      });
      onUpdate(sessions);
    }, (err) => {
      console.warn("subscribeToDeviceSessions error:", err);
    });
  });

  return () => {
    if (unsubscribeSnapshot) unsubscribeSnapshot();
  };
};

