import { ActivityLogEntry, ActivityActionType, ActivityModule } from '../types';
import { miniDB } from '../db';
import { 
  saveCollectionToFirestore, 
  fetchCollectionFromFirestore, 
  subscribeToFirestoreCollection 
} from '../lib/firebase';
import * as XLSX from 'xlsx';

const ACTIVITY_LOGS_KEY = 'imm_activity_logs_v1';
const MAX_LOGS = 300; // Cap at 300 records as requested

// Format readable date-time in Myanmar standard format DD/MM/YYYY, hh:mm:ss A
function formatReadableDateTime(d: Date): string {
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  let hours = d.getHours();
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const seconds = String(d.getSeconds()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12;
  const strHours = String(hours).padStart(2, '0');
  return `${day}/${month}/${year}, ${strHours}:${minutes}:${seconds} ${ampm}`;
}

let cachedLogs: ActivityLogEntry[] | null = null;

export async function getActivityLogs(fromCloud: boolean = false): Promise<ActivityLogEntry[]> {
  let localLogs: ActivityLogEntry[] = [];
  try {
    const fromDB = await miniDB.get(ACTIVITY_LOGS_KEY);
    if (Array.isArray(fromDB)) {
      localLogs = fromDB;
    }
  } catch (e) {
    console.warn("Could not read activity logs from IndexedDB:", e);
  }

  // Fallback to localStorage if IndexedDB is empty
  if (localLogs.length === 0) {
    try {
      const raw = localStorage.getItem(ACTIVITY_LOGS_KEY);
      if (raw) {
        localLogs = JSON.parse(raw) || [];
      }
    } catch (e) {
      console.warn("Could not read activity logs from localStorage:", e);
    }
  }

  if (fromCloud) {
    try {
      const remoteLogs = await fetchCollectionFromFirestore('activityLogs');
      if (Array.isArray(remoteLogs) && remoteLogs.length > 0) {
        const map = new Map<string, ActivityLogEntry>();
        [...remoteLogs, ...localLogs].forEach(l => {
          if (l && l.id && !map.has(l.id)) {
            map.set(l.id, l);
          }
        });
        const merged = Array.from(map.values())
          .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime())
          .slice(0, MAX_LOGS);

        cachedLogs = merged;
        miniDB.set(ACTIVITY_LOGS_KEY, merged).catch(() => {});
        try {
          localStorage.setItem(ACTIVITY_LOGS_KEY, JSON.stringify(merged.slice(0, 100)));
        } catch {}
        return merged;
      }
    } catch (err) {
      console.warn("Could not fetch remote activity logs:", err);
    }
  }

  if (localLogs.length > 0) {
    cachedLogs = localLogs.slice(0, MAX_LOGS);
    return cachedLogs;
  }

  if (cachedLogs) return cachedLogs.slice(0, MAX_LOGS);
  cachedLogs = [];
  return [];
}

/**
 * Subscribes to live Activity Log changes from Cloud Firestore (for Superadmin monitoring)
 */
export function subscribeToCloudActivityLogs(onUpdate: (logs: ActivityLogEntry[]) => void) {
  return subscribeToFirestoreCollection('activityLogs', (remoteItems) => {
    if (Array.isArray(remoteItems)) {
      const sorted = remoteItems
        .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime())
        .slice(0, MAX_LOGS);
      cachedLogs = sorted;
      miniDB.set(ACTIVITY_LOGS_KEY, sorted).catch(() => {});
      try {
        localStorage.setItem(ACTIVITY_LOGS_KEY, JSON.stringify(sorted.slice(0, 100)));
      } catch {}
      onUpdate(sorted);
    }
  });
}

export async function logActivity(params: {
  action: ActivityActionType;
  module: ActivityModule;
  officerName?: string;
  officerRole?: string;
  deviceId?: string;
  targetId?: string;
  details: string;
}): Promise<ActivityLogEntry> {
  const now = new Date();
  const currentLogs = await getActivityLogs(false);

  // Get current active user identity (Editor, Superadmin, Viewer, or Officer)
  const storedOfficerName = localStorage.getItem('lastCheckedOfficerName') || '';
  const storedOfficerTitle = localStorage.getItem('lastCheckedOfficerTitle') || '';
  const cloudAuthRaw = localStorage.getItem('imm_pwa_cloud_auth_user') || localStorage.getItem('imm_cloud_auth_user');
  let authUsername = 'OFFICER';
  let authRole = 'Editor';
  let authDeviceName = '';
  if (cloudAuthRaw) {
    try {
      const parsed = JSON.parse(cloudAuthRaw);
      authUsername = parsed.username || authUsername;
      authRole = parsed.role || authRole;
      authDeviceName = parsed.deviceName || '';
    } catch {}
  }

  const effectiveOfficerName = params.officerName || 
    (storedOfficerTitle && storedOfficerName ? `${storedOfficerTitle} ${storedOfficerName}` : authUsername);
  const effectiveRole = params.officerRole || authRole;
  const currentDeviceId = params.deviceId || authDeviceName || localStorage.getItem('imm_pwa_device_id') || 'local_terminal';

  const newEntry: ActivityLogEntry = {
    id: `ACT-${Date.now()}-${Math.random().toString(36).substring(2, 7).toUpperCase()}`,
    timestamp: now.toISOString(),
    readableTime: formatReadableDateTime(now),
    action: params.action,
    module: params.module,
    officerName: effectiveOfficerName,
    officerRole: effectiveRole,
    deviceId: currentDeviceId,
    targetId: params.targetId,
    details: params.details,
  };

  // Prepend and strictly maintain max 300 entries
  const updatedLogs = [newEntry, ...currentLogs.filter(l => l.id !== newEntry.id)].slice(0, MAX_LOGS);
  cachedLogs = updatedLogs;

  // Persist locally
  try {
    await miniDB.set(ACTIVITY_LOGS_KEY, updatedLogs);
    localStorage.setItem(ACTIVITY_LOGS_KEY, JSON.stringify(updatedLogs.slice(0, 100)));
  } catch (e) {
    console.error("Failed to save activity log locally:", e);
  }

  // Dispatch custom event so reactive UI components can refresh instantly
  if (typeof window !== 'undefined') {
    window.dispatchEvent(new CustomEvent('imm_activity_logged', { detail: newEntry }));
  }

  // Asynchronously sync new activity log entry to Cloud Firestore across all users
  saveCollectionToFirestore('activityLogs', updatedLogs, true).catch(err => {
    console.warn("Async Cloud save activity log failed (safely stored locally):", err);
  });

  return newEntry;
}

export async function clearActivityLogs(): Promise<void> {
  cachedLogs = [];
  try {
    await miniDB.remove(ACTIVITY_LOGS_KEY);
    localStorage.removeItem(ACTIVITY_LOGS_KEY);
    await saveCollectionToFirestore('activityLogs', [], true);
    if (typeof window !== 'undefined') {
      window.dispatchEvent(new CustomEvent('imm_activity_logged', { detail: null }));
    }
  } catch (e) {
    console.error("Failed to clear activity logs:", e);
  }
}

export function exportActivityLogsToExcel(logs: ActivityLogEntry[]): void {
  if (!logs || logs.length === 0) return;

  const excelData = logs.slice(0, MAX_LOGS).map((log, idx) => ({
    "စဉ်": idx + 1,
    "အချိန် / ရက်စွဲ": log.readableTime,
    "လုပ်ဆောင်ချက် (Action)": log.action,
    "ကဏ္ဍ (Module)": log.module,
    "အရာရှိအမည်": log.officerName,
    "ရာထူး/အခွင့်အရေး": log.officerRole,
    "ပတ်စပို့ / အညွှန်း (Target)": log.targetId || '-',
    "အသေးစိတ်အချက်အလက် (Details)": log.details,
    "စက်အမှတ် (Device ID)": log.deviceId || '-'
  }));

  const worksheet = XLSX.utils.json_to_sheet(excelData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Activity_Audit_Logs");

  const wscols = [
    { wch: 6 },
    { wch: 22 },
    { wch: 18 },
    { wch: 14 },
    { wch: 22 },
    { wch: 16 },
    { wch: 20 },
    { wch: 45 },
    { wch: 20 },
  ];
  worksheet['!cols'] = wscols;

  const todayStr = new Date().toISOString().split('T')[0];
  XLSX.writeFile(workbook, `IMM_Activity_Audit_Logs_${todayStr}.xlsx`);
}
