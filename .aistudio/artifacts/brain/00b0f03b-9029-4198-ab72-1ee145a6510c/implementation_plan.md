# Direct Collection Cloud Sync & Live Counter Strategy

Ensure the application reliably fetches, synchronizes, and counts records directly from the `(default)` Firestore database without relying on intermediate `_metadata` documents.

## User Review & Critical Decisions

> [!IMPORTANT]
> The following choices were confirmed by the user in Phase 1:

- **Confirmed Data Fetching**: **Direct Collection Fetch**. Read `items` directly from `appData` collection documents (`appData/records`, `appData/tempRecords`, `appData/masterData`, etc.) on the `(default)` Firestore database.
- **Confirmed Cloud Counts Calculation**: **Live Item Count**. Calculate exact Cloud counts by reading the `items` array length directly from each `appData` Cloud document payload rather than depending on cached `_metadata` fields.

---

## 1. Overview & Core Concept

- **Goal**: Restore 100% direct Cloud synchronization with the `(default)` Firestore database.
- **User Experience**: When the user opens the **Cloud Sync Monitor** or clicks **PULL CLOUD**, the app queries the `(default)` Firestore database directly, retrieves document payloads for all 6 collections, displays exact live item counts, and merges cloud logs into local storage.

---

## 2. User Experience & Visual Design

- **Cloud Sync Monitor Dashboard**:
  - Displays real-time live counts for `FFE & FH Logs`, `Movement (M)`, `Master DB (MD)`, `Table Output (TO)`, `Dossier (INV)`, and `Watch-list (WL)`.
  - Shows explicit `Synced` / `Pending Sync` status badges.
- **Visual Feedback**:
  - Progress spinner during active direct document fetches.
  - Toast notification showing exact record count retrieved (`✅ Cloud Database (default) မှ စာရင်း [X] ခုအား အောင်မြင်စွာ ရယူပြီးပါပြီ`).

---

## 3. Key Product Decisions & Trade-Offs

- **Decision 1: Direct Collection Reading vs. Metadata Caching**:
  - *Chosen Approach*: Query `doc(db, 'appData', collectionName)` directly.
  - *Why*: Eliminates dependency on `_metadata` document existence or staleness on the `(default)` database.
- **Decision 2: Live Item Array Count Calculation**:
  - *Chosen Approach*: Inspect `docSnap.data().items.length` directly.
  - *Why*: Guarantees accurate item counts matching the exact payload stored in Firestore.

---

## 4. Technical Architecture & Data Strategy

```
┌───────────────────────────────────────────────────────────────────────┐
│                       Cloud Sync Monitor UI                           │
│  - Displays Live Cloud Item Counts (records, tempRecords, master, etc)│
│  - "PULL CLOUD" Button -> Triggers handleFetchDataFromCloud(true)     │
└──────────────────────────────────┬────────────────────────────────────┘
                                   │
                                   ▼
┌───────────────────────────────────────────────────────────────────────┐
│                    src/lib/firebase.ts Engine                         │
│  - db = getFirestore(app)  [Targets (default) Database]               │
│  - fetchCloudCollectionStats(): Direct getDoc(doc(db, 'appData', col))│
│  - fetchCollectionFromFirestore(): Direct payload array retrieval     │
└──────────────────────────────────┬────────────────────────────────────┘
                                   │
                                   ▼
┌───────────────────────────────────────────────────────────────────────┐
│                    Google Cloud Firestore (default)                   │
│  - /appData/records           -> { items: [...] }                     │
│  - /appData/tempRecords       -> { items: [...] }                     │
│  - /appData/masterData        -> { items: [...] }                     │
│  - /appData/vehicleSummaries  -> { items: [...] }                     │
│  - /appData/dossierHistory    -> { items: [...] }                     │
│  - /appData/watchList         -> { items: [...] }                     │
└───────────────────────────────────────────────────────────────────────┘
```

### Data Flow & Component Mapping
1. `getFirestore(app)` initializes connection to `(default)` database.
2. `fetchCloudCollectionStats(force)` executes parallel `getDocFromServer(doc(db, 'appData', col))` for all 6 collections.
3. Live counts are computed via `snap.data().items.length` and stored in `cloudStats` state.
4. `handleFetchDataFromCloud()` merges fetched `items` into local IndexedDB and React state using composite deduplication keys (`id + passport + timestamp`).
