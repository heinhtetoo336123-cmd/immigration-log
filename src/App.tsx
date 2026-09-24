import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { initAuth, saveCollectionToFirestore, subscribeToFirestoreCollection, fetchCollectionFromFirestore, checkAndFetchUpdatedCollections, resetQuotaState, onQuotaStatusChange, getIsQuotaExhausted, setQuotaExhausted, onWriteError, WriteErrorInfo, setSyncedHash, saveDeviceSession, setDeviceKickedStatus, updateDeviceRole, deleteDeviceSession, subscribeToDeviceSessions, subscribeToMyDeviceSession, fetchDeviceSessions, getCloudMetadata, getLocalTimestamps } from './lib/firebase';
import { 
  Link2,
  Link,
  Users, 
  History, 
  Activity, 
  Layers, 
  GitMerge,
  Calendar, 
  Database, 
  PlusCircle, 
  MapPin, 
  Truck, 
  Search, 
  FileText, 
  Download, 
  Upload, 
  Trash2, 
  Edit,
  Edit2,
  CheckSquare,
  ChevronRight, 
  CheckCircle2, 
  AlertCircle,
  AlertTriangle,
  LogOut,
  LogIn,
  RefreshCcw,
  RefreshCw,
  X,
  Copy,
  Printer,
  ChevronDown,
  Clock,
  Plane,
  ClipboardList,
  Clipboard,
  PieChart,
  Briefcase,
  Building2,
  Home,
  Settings2,
  HelpCircle,
  Table,
  ShieldCheck,
  HardDrive,
  Check,
  ChevronUp,
  Plus,
  FileSpreadsheet,
  Code,
  Shield,
  Key,
  Lock,
  Crown,
  Eye,
  UserCheck,
  Smartphone,
  Laptop,
  Menu,
  User,
  Wifi,
  WifiOff,
  Settings,
  SlidersHorizontal,
  ShieldAlert,
  Maximize2,
  ZoomIn,
  ZoomOut,
  Server,
  UploadCloud,
  Sparkles
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import html2canvas from 'html2canvas';
import { 
  Mode, 
  ImmRecord, 
  MasterItem, 
  Tab,
  MovementData,
  VehicleSummary,
  DossierRecord,
  WatchListRecord,
  PermitStatus,
  normalizePermitStatus
} from './types';
import { DobNumpadInput } from './components/DobNumpadInput';
import { StayPeriodInput } from './components/StayPeriodInput';
import { IndividualSearch } from './components/IndividualSearch';
import { CounterCheckModal } from './components/CounterCheckModal';
import { CloudSyncStatusModal } from './components/CloudSyncStatusModal';
import { WatchList } from './components/WatchList';
import { CustomReportTable } from './components/CustomReportTable';
import { ActivityLogView } from './components/ActivityLogView';
import { PermitManagementConsole } from './components/PermitManagementConsole';
import { logActivity } from './utils/activityLogger';
import { analyzeRecordsForErrors } from './utils/counterCheck';
import { miniDB } from './db';

// --- UTILS ---
export async function hashSHA256(str: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(str);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

export function sanitizeInput(str: string): string {
  if (!str) return '';
  return str
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;')
    .trim();
}

export function normalizeOfficerName(str: string): string {
  if (!str) return '';
  return str.replace(/\s+/g, '').toUpperCase();
}
export const parseTimestamp = (str: any): number => {
  if (!str) return 0;
  try {
    const s = String(str).trim();
    if (!s) return 0;
    // 1. Check if the string is clearly an ISO string or native format with 'T'
    if (s.includes('T')) {
      const p = Date.parse(s);
      if (!isNaN(p) && p > 0) return p;
    }

    // Check for 12-hour AM/PM format
    const isPM = /pm/i.test(s);
    const isAM = /am/i.test(s);

    // Clean string and extract all runs of digits (numbers)
    const nums = s.match(/\d+/g)?.map(Number);
    if (!nums || nums.length < 3) {
      const parsed = Date.parse(s);
      return isNaN(parsed) ? 0 : parsed;
    }

    let y = 0, m = 0, d = 0;
    let h = 0, min = 0, sec = 0;

    // Check if the first token is a 4-digit year format (e.g. YYYY-MM-DD)
    if (nums[0] > 1000 && nums[0] < 3000) {
      y = nums[0];
      m = nums[1];
      d = nums[2];
      h = nums[3] || 0;
      min = nums[4] || 0;
      sec = nums[5] || 0;
    } else if (nums[2] > 1000 && nums[2] < 3000) {
      // DD/MM/YYYY or DD-MM-YYYY format
      d = nums[0];
      m = nums[1];
      y = nums[2];
      h = nums[3] || 0;
      min = nums[4] || 0;
      sec = nums[5] || 0;
    } else {
      // Fallback to standard JS parsing as a last resort
      const parsed = Date.parse(s);
      return isNaN(parsed) ? 0 : parsed;
    }

    if (isPM && h < 12) h += 12;
    if (isAM && h === 12) h = 0;

    // Adjust month for 0-indexed Date constructor in JS
    const dateObj = new Date(y, m - 1, d, h, min, sec);
    const timeVal = dateObj.getTime();
    if (isNaN(timeVal)) {
      const fallback = Date.parse(s);
      return isNaN(fallback) ? 0 : fallback;
    }
    return timeVal;
  } catch (e) {
    return 0;
  }
};

export const getRecordTime = (r: any): number => {
  if (!r) return 0;
  if (r.updatedAt) {
    const p = Date.parse(r.updatedAt);
    if (!isNaN(p) && p > 0) return p;
  }
  const t = parseTimestamp(r.timestamp);
  if (t > 0) return t;
  if (typeof r.id === 'number' && r.id > 1000000000000) return r.id;
  return 0;
};

export const getLatestVerificationRemark = (passport: string, dossierHistoryList: DossierRecord[], recordsList: ImmRecord[]) => {
  try {
    const pp = (passport || '').toUpperCase();
    if (!pp) return '-';
    
    // 1. Check dossierHistory (INV tab investigation dossiers)
    const dHist = (dossierHistoryList || []).filter(d => d && d.remarksMap && typeof d.remarksMap === 'object' && d.remarksMap[pp] && typeof d.remarksMap[pp] === 'string' && d.remarksMap[pp].trim());
    if (dHist.length > 0) {
      const latest = dHist[dHist.length - 1];
      return String(latest.remarksMap![pp]).trim();
    }
    
    // 2. Check recordsList (INV tab investigation notes saved on ImmRecord)
    const userRecs = (recordsList || [])
      .filter(r => r && (r.passport || '').toUpperCase() === pp && r.remarks && typeof r.remarks === 'string' && r.remarks.trim())
      .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
      
    if (userRecs.length > 0) {
      return String(userRecs[0].remarks).trim();
    }
    
    return '-';
  } catch (e) {
    return '-';
  }
};

const formatFriendlyDate = (inputDate: string): string => {
  if (!inputDate) return "";
  if (inputDate.includes('/')) {
    const datePart = inputDate.split(',')[0].trim();
    return datePart.replace(/\//g, '-');
  }
  const parts = inputDate.split('-');
  if (parts.length === 3 && parts[0].length === 4) {
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return inputDate;
};

const getLiveRemainingDays = (toDateStr: string): string => {
  if (!toDateStr) return "-";
  const targetDate = new Date(toDateStr);
  targetDate.setHours(23, 59, 59, 999);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffTime = targetDate.getTime() - today.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return `${diffDays} R`;
};

export const formatToDDMMYYYY = (dateStr: string | undefined | null): string => {
  if (!dateStr || dateStr === 'N/A' || dateStr === '-') return '';
  const trimmed = String(dateStr).trim();
  if (!trimmed) return '';

  // 1. Detect standard ISO/Autofill date format: YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
  const ymdMatch = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:T.*)?$/);
  if (ymdMatch) {
    const [, y, m, d] = ymdMatch;
    return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
  }

  // 2. Detect formatted DD-MM-YYYY, DD/MM/YYYY, DD.MM.YYYY
  const dmyMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const [, d, m, y] = dmyMatch;
    return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
  }

  // 3. Fallback for Date strings
  const d = new Date(trimmed);
  if (!isNaN(d.getTime())) {
    const day = String(d.getUTCDate()).padStart(2, '0');
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const year = d.getUTCFullYear();
    return `${day}-${month}-${year}`;
  }

  return trimmed;
};

const formatDateToDDMMYYYY = (dateStr: string) => {
  if (!dateStr || dateStr === 'N/A') return 'N/A';
  return formatToDDMMYYYY(dateStr) || dateStr;
};

// --- COMPONENTS ---

const Toast = ({ message, visible }: { message: string; visible: boolean }) => (
  <AnimatePresence>
    {visible && (
      <motion.div
        initial={{ opacity: 0, y: 50, x: '-50%' }}
        animate={{ opacity: 1, y: 0, x: '-50%' }}
        exit={{ opacity: 0, y: 50, x: '-50%' }}
        className="fixed bottom-8 left-1/2 -translate-x-1/2 z-50 bg-gray-900 text-white px-6 py-3 rounded-xl shadow-2xl flex items-center gap-3 font-semibold text-sm border border-white/10"
      >
        <CheckCircle2 size={18} className="text-emerald-400" />
        {message}
      </motion.div>
    )}
  </AnimatePresence>
);

// --- STANDALONE COMPONENTS ---

// --- FORM C REGISTRATION COMPONENT ---
const FormCRegistration = () => null;

const FormCRegistration_OLD = ({ 
  records, 
  setRecords, 
  tempRecords,
  setTempRecords,
  masterData, 
  showToast,
  syncMaster,
  deleteRecord,
  preFilledTempId,
  setPreFilledTempId,
  currentUser
}: { 
  records: ImmRecord[]; 
  setRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  tempRecords: ImmRecord[];
  setTempRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  masterData: MasterItem[];
  showToast: (msg: string) => void;
  syncMaster: (val: string, type: MasterItem['type'], linkedValue?: string) => void;
  deleteRecord: (id: number) => void;
  preFilledTempId: number | null;
  setPreFilledTempId: (id: number | null) => void;
  currentUser: { name: string, title: string } | null;
}) => {
  const [search, setSearch] = useState('');
  const [logSearch, setLogSearch] = useState('');
  const [targetRecord, setTargetRecord] = useState<ImmRecord | null>(null);
  const [isCreatingTemp, setIsCreatingTemp] = useState(false);
  const [listDate, setListDate] = useState(new Date().toISOString().split('T')[0]);
  const [formCDeleteId, setFormCDeleteId] = useState<number | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<number | null>(null);
  const [fcrConfirmation, setFcrConfirmation] = useState<{ show: boolean, message: string, data?: any } | null>(null);
  const [fcrSuccessMsg, setFcrSuccessMsg] = useState<string | null>(null);

  useEffect(() => {
    if (currentUser) {
      setForm(prev => ({
        ...prev,
        officialName: currentUser.name || '',
        officialTitle: currentUser.title || ''
      }));
    }
  }, [currentUser]);

  useEffect(() => {
    if (preFilledTempId) {
      const match = tempRecords.find(r => r.id === preFilledTempId);
      if (match) {
        selectPerson(match);
      }
      setPreFilledTempId(null);
    }
  }, [preFilledTempId]);
  
  const [form, setForm] = useState({
    status: '' as 'IN' | 'OUT' | '',
    officialName: '',
    officialTitle: '',
    reporterName: '',
    reporterPhone: '',
    submissionDate: new Date().toISOString().split('T')[0],
    submissionTime: new Date().toTimeString().split(' ')[0].substring(0, 5),
    stayLocation: '',
    stayDescription: '',
    vehicleInfo: '',
    arrivedFrom: '',
    departedTo: '',
    // Temp entry fields
    fullname: '',
    passport: '',
    gender: 'M' as 'M' | 'F' | '',
    nationality: '',
    visaType: '',
    stayFrom: '',
    stayTo: '',
    totalDays: '',
    remarks: '',
    broughtBy: '',
    contactDetails: ''
  });

  const calculateTempStayDates = (field: 'from' | 'days' | 'to', value: string) => {
    let from = field === 'from' ? value : form.stayFrom;
    let to = field === 'to' ? value : form.stayTo;
    let days = field === 'days' ? value : form.totalDays;

    if (field === 'from' || field === 'days') {
      if (from && days) {
        const d = new Date(from);
        d.setDate(d.getDate() + parseInt(days));
        to = d.toISOString().split('T')[0];
      }
    } else if (field === 'to') {
      if (from && to) {
        const d1 = new Date(from);
        const d2 = new Date(to);
        const diffTime = Math.abs(d2.getTime() - d1.getTime());
        days = Math.ceil(diffTime / (1000 * 60 * 60 * 24)).toString();
      }
    }
    setForm(prev => ({ ...prev, stayFrom: from, stayTo: to, totalDays: days }));
  };

  const searchResults = useMemo(() => {
    if (!search.trim()) return [];
    const q = search.toLowerCase();
    const map = new Map<string, ImmRecord>();
    
    // Combine both pools, tempRecords prioritized if they are newer/identical
    const allPool = [...records, ...tempRecords];
    
    allPool.sort((a,b) => parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp)).forEach(r => {
      if (r.fullname.toLowerCase().includes(q) || r.passport.toLowerCase().includes(q)) {
        map.set(r.passport, r);
      }
    });
    return Array.from(map.values()).reverse();
  }, [search, records, tempRecords]);

  const formCLogs = useMemo(() => {
    const all = [...records, ...tempRecords].sort((a,b) => parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp));
    const groups = new Map<string, any>();
    
    all.forEach(r => {
      const passport = r.passport.toUpperCase();
      if (!groups.has(passport)) {
        groups.set(passport, {
          passport: r.passport,
          fullname: r.fullname,
          nationality: r.nationality,
          lastRecord: r,
          logIn: null,
          logOut: null,
          formCIn: null,
          formCOut: null,
        });
      }
      const g = groups.get(passport);
      g.fullname = r.fullname;
      g.nationality = r.nationality;
      g.lastRecord = r;

      if (r.mode === 'IN') g.logIn = r;
      if (r.mode === 'OUT') g.logOut = r;
      if (r.formC?.status === 'IN') g.formCIn = r;
      if (r.formC?.status === 'OUT') g.formCOut = r;
    });

    const logs = Array.from(groups.values()).sort((a,b) => parseTimestamp(b.lastRecord.timestamp) - parseTimestamp(a.lastRecord.timestamp));
    if (!logSearch.trim()) return logs;
    const q = logSearch.toLowerCase();
    return logs.filter(l => 
      l.fullname.toLowerCase().includes(q) || 
      l.passport.toLowerCase().includes(q) ||
      l.nationality.toLowerCase().includes(q)
    );
  }, [records, tempRecords, logSearch]);

  const lastStayMap = useMemo(() => {
    const map = new Map<string, string>();
    [...records, ...tempRecords]
      .sort((a,b) => parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp))
      .forEach(r => {
        if (r.address) map.set(r.passport, r.address);
      });
    return map;
  }, [records, tempRecords]);

  const exportFCRLogsToExcel = () => {
    const data = formCLogs.map((g, i) => ({
      "No": i + 1,
      "Full Name": g.fullname,
      "Passport": g.passport,
      "Nationality": g.nationality,
      "Log IN": g.logIn ? formatDateToDDMMYYYY(g.logIn.timestamp) : 'N/A',
      "Log OUT": g.logOut ? formatDateToDDMMYYYY(g.logOut.timestamp) : 'N/A',
      "Form C IN": g.formCIn ? formatDateToDDMMYYYY(g.formCIn.timestamp) : 'N/A',
      "Form C OUT": g.formCOut ? formatDateToDDMMYYYY(g.formCOut.timestamp) : 'N/A',
      "Stay Location": g.formCIn?.formC?.address || g.formCOut?.formC?.address || g.lastRecord.address || 'N/A',
      "Official": (g.formCOut?.formC || g.formCIn?.formC)?.officialName || 'N/A',
      "Reporter": (g.formCOut?.formC || g.formCIn?.formC)?.reporterName || 'N/A',
    }));

    // Sheet 2: Mismatched Log and Form C entries
    const mismatchedData = formCLogs.filter(g => {
      const hasLogMsg = g.logIn || g.logOut;
      const hasFormCMsg = g.formCIn || g.formCOut;
      if (!hasLogMsg || !hasFormCMsg) return false;
      const lastLogMode = g.logOut ? 'OUT' : 'IN';
      const lastFormCMode = g.formCOut ? 'OUT' : 'IN';
      return lastLogMode !== lastFormCMode;
    }).map((g, i) => ({
      "No": i + 1,
      "Full Name": g.fullname,
      "Passport": g.passport,
      "Nationality": g.nationality,
      "Last Log Mode": g.logOut ? 'OUT' : 'IN',
      "Last Log Date": g.logOut ? formatDateToDDMMYYYY(g.logOut.timestamp) : formatDateToDDMMYYYY(g.logIn?.timestamp || ""),
      "Form C Mode": g.formCOut ? 'OUT' : 'IN',
      "Form C Date": g.formCOut ? formatDateToDDMMYYYY(g.formCOut.timestamp) : formatDateToDDMMYYYY(g.formCIn?.timestamp || ""),
      "Stay Location": g.formCIn?.formC?.address || g.formCOut?.formC?.address || g.lastRecord.address || 'N/A'
    }));

    const ws1 = XLSX.utils.json_to_sheet(data);
    const ws2 = XLSX.utils.json_to_sheet(mismatchedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, "FCR Logs");
    XLSX.utils.book_append_sheet(wb, ws2, "Mismatched Logs");
    XLSX.writeFile(wb, `FCR_Logs_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const resetForm = () => {
    setForm({
      status: '' as 'IN' | 'OUT' | '',
      officialName: '',
      officialTitle: '',
      reporterName: '',
      reporterPhone: '',
      submissionDate: new Date().toISOString().split('T')[0],
      submissionTime: new Date().toTimeString().split(' ')[0].substring(0, 5),
      stayLocation: '',
      stayDescription: '',
      vehicleInfo: '',
      arrivedFrom: '',
      departedTo: '',
      fullname: '',
      passport: '',
      gender: 'M' as 'M' | 'F' | '',
      nationality: '',
      visaType: '',
      stayFrom: '',
      stayTo: '',
      totalDays: '',
      remarks: '',
      broughtBy: '',
      contactDetails: ''
    });
  };

  const deleteFormCData = (id: number) => {
    setRecords(prev => prev.map(r => r.id === id ? { ...r, formC: undefined } : r));
    setTempRecords(prev => prev.map(r => r.id === id ? { ...r, formC: undefined } : r));
    setFormCDeleteId(null);
    showToast("FORM C DATA REMOVED");
  };

  const editLog = (r: ImmRecord) => {
    setTargetRecord(r);
    const stayLoc = r.formC?.address || r.address || '';
    const allRecs = [...records, ...tempRecords].sort((a, b) => getRecordTime(b) - getRecordTime(a));
    const latestRecWithDesc = allRecs.find(rec => (rec.formC?.address?.toLowerCase() === stayLoc.toLowerCase() || rec.address?.toLowerCase() === stayLoc.toLowerCase()) && (rec.formC?.stayDescription || rec.stayDescription));
    const latestMaster = [...masterData].reverse().find(m => m.type === 'Stay' && m.name.toLowerCase() === stayLoc.toLowerCase() && m.linkedValue);
    const stayDesc = r.formC?.stayDescription || r.stayDescription || latestRecWithDesc?.formC?.stayDescription || latestRecWithDesc?.stayDescription || latestMaster?.linkedValue || '';
    setForm({
      status: r.formC?.status || '',
      officialName: r.formC?.officialName || '',
      officialTitle: r.formC?.officialTitle || '',
      reporterName: r.formC?.reporterName || '',
      reporterPhone: r.formC?.reporterPhone || '',
      submissionDate: r.formC?.submissionDate || new Date().toISOString().split('T')[0],
      submissionTime: r.formC?.submissionTime || new Date().toTimeString().split(' ')[0].substring(0, 5),
      stayLocation: stayLoc,
      stayDescription: stayDesc,
      vehicleInfo: r.vehicleInfo || '',
      arrivedFrom: r.arrivedFrom || '',
      departedTo: r.departedTo || ''
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const selectPerson = (r: ImmRecord) => {
    setTargetRecord(r);
    
    // Selecting an existing person (FH or TEH) means we are NOT creating a new temporary entry from scratch
    setIsCreatingTemp(false);
    
    setSearch('');
    setForm(prev => {
      const stayLoc = r.formC?.address || r.address || '';
      const allRecs = [...records, ...tempRecords].sort((a, b) => getRecordTime(b) - getRecordTime(a));
      const latestRecWithDesc = allRecs.find(rec => (rec.formC?.address?.toLowerCase() === stayLoc.toLowerCase() || rec.address?.toLowerCase() === stayLoc.toLowerCase()) && (rec.formC?.stayDescription || rec.stayDescription));
      const latestMaster = [...masterData].reverse().find(m => m.type === 'Stay' && m.name.toLowerCase() === stayLoc.toLowerCase() && m.linkedValue);
      const stayDesc = r.formC?.stayDescription || r.stayDescription || latestRecWithDesc?.formC?.stayDescription || latestRecWithDesc?.stayDescription || latestMaster?.linkedValue || '';
      return {
        ...prev,
        officialName: currentUser?.name || prev.officialName,
        officialTitle: currentUser?.title || prev.officialTitle,
        fullname: r.fullname || '',
        passport: r.passport || '',
        gender: r.gender || 'M',
        nationality: r.nationality || '',
        visaType: r.visaType || '',
        stayFrom: r.stayFrom || '',
        stayTo: r.stayTo || '',
        totalDays: r.totalDays || '',
        remarks: r.remarks || '',
        stayLocation: stayLoc,
        stayDescription: stayDesc,
        vehicleInfo: '', // Clear vehicle info in FCR to avoid incorrect autofill from logs or previous state
        arrivedFrom: prev.arrivedFrom,
        departedTo: prev.departedTo,
        reporterName: r.formC?.reporterName || '',
        reporterPhone: r.formC?.reporterPhone || '',
        status: ''
      };
    });
  };

  const handleOfficialNameChange = (name: string) => {
    setForm(prev => ({ ...prev, officialName: name }));
    if (!name.trim()) return;
    const val = name.trim().toLowerCase().replace(/\s+/g, '');
    
    // 1. Check latest records first for the most recently used title
    const all = [...records, ...tempRecords].sort((a, b) => getRecordTime(b) - getRecordTime(a));
    const latestWithTitle = all.find(r => 
      (r.officialName?.toLowerCase().replace(/\s+/g, '') === val && r.officialTitle && r.officialTitle.trim()) ||
      (r.formC?.officialName?.toLowerCase().replace(/\s+/g, '') === val && r.formC.officialTitle && r.formC.officialTitle.trim())
    );
    if (latestWithTitle) {
      const finalTitle = latestWithTitle.formC?.officialName?.toLowerCase().replace(/\s+/g, '') === val
        ? latestWithTitle.formC.officialTitle
        : latestWithTitle.officialTitle;
      if (finalTitle) setForm(prev => ({ ...prev, officialTitle: finalTitle }));
    } else {
      // 2. Fallback to latest Master Data
      const masterMatch = [...masterData].reverse().find(m => m.type === 'Official' && m.name.toLowerCase().replace(/\s+/g, '') === val && m.linkedValue);
      if (masterMatch && masterMatch.linkedValue) {
        setForm(prev => ({ ...prev, officialTitle: masterMatch.linkedValue || '' }));
      }
    }
  };

  const handleReporterNameChange = (name: string) => {
    setForm(prev => ({ 
      ...prev, 
      reporterName: name,
      // Sync with Agent if creating temp
      ...(isCreatingTemp ? { broughtBy: name } : {})
    }));
    // Auto-fill phone from last known record for this reporter
    if (!name.trim()) return;
    const cleanName = name.trim().toLowerCase();
    const latest = [...records, ...tempRecords]
      .filter(r => r.formC?.reporterName?.toLowerCase() === cleanName && r.formC.reporterPhone)
      .sort((a,b) => getRecordTime(b) - getRecordTime(a))[0];
    if (latest && latest.formC && latest.formC.reporterPhone) {
      setForm(prev => ({ 
        ...prev, 
        reporterPhone: latest.formC.reporterPhone,
        ...(isCreatingTemp ? { contactDetails: latest.formC.reporterPhone } : {})
      }));
    } else {
      const masterMatch = [...masterData].reverse().find(m => m.type === 'Reporter' && m.name.toLowerCase() === cleanName && m.linkedValue);
      if (masterMatch && masterMatch.linkedValue) {
        setForm(prev => ({ 
          ...prev, 
          reporterPhone: masterMatch.linkedValue || '',
          ...(isCreatingTemp ? { contactDetails: masterMatch.linkedValue || '' } : {})
        }));
      }
    }
  };

  const handleReporterPhoneChange = (phone: string) => {
    setForm(prev => ({ 
      ...prev, 
      reporterPhone: phone,
      ...(isCreatingTemp ? { contactDetails: phone } : {})
    }));
  };

  const submitFormC = (e: React.FormEvent) => {
    e.preventDefault();
    if (!targetRecord) return;
    if (!form.status) {
      showToast("PLEASE SELECT STATUS (IN/OUT)");
      return;
    }

    const lastLog = [...records, ...tempRecords]
      .filter(r => r.passport.toUpperCase() === targetRecord.passport.toUpperCase())
      .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];

    const logDate = lastLog ? formatDateToDDMMYYYY(lastLog.timestamp) : 'N/A';
    const logMode = lastLog?.mode || 'IN';
    const formCMode = form.status;

    let message = `လေဆိပ်မှ ${logDate}ရက်နေ့တွင် ${logMode === 'IN' ? 'အဝင်' : 'အထွက်'} ပြထားသူဖြစ်ပါသည်၊ Form C ကို ${formatDateToDDMMYYYY(form.submissionDate)} ရက်စွဲဖြင့် ${form.stayLocation} မှ ${formCMode === 'IN' ? 'အဝင်' : 'အထွက်'} လာရောက်တိုင်ကြားသည်မှာ မှန်ကန်ပါသလား။`;
    
    if (logMode !== formCMode) {
      if (logMode === 'OUT' && formCMode === 'IN') {
        message += " လေဆိပ်မဟုတ်သော အခြားနေရာမှ ဝင်ရောက်လာသူဖြစ်နိုင်ပါသည်။";
      } else if (logMode === 'IN' && formCMode === 'OUT') {
        message += " လေဆိပ်မဟုတ်သော အခြားနေရာမှ ထွက်ခွါသွားသူဖြစ်နိုင်ပါသည်။";
      }
    }

    setFcrConfirmation({ show: true, message, data: { isTemp: false } });
  };

  const executeSubmitFormC = () => {
    if (!targetRecord) return;
    
    const parts = form.submissionDate.split('-');
    const slashDate = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : form.submissionDate;
    const calculatedTimestamp = `${slashDate}, ${form.submissionTime || '00:00'}`;

    const brandNewFcrRecord: ImmRecord = {
      ...targetRecord,
      id: targetRecord.logType === 'FCR' ? targetRecord.id : Date.now(), // Unique ID if creating from FFE
      timestamp: calculatedTimestamp,
      logType: 'FCR', // Always log as standalone FCR
      address: form.stayLocation,
      remarks: form.remarks,
      formC: {
        status: form.status as 'IN' | 'OUT',
        officialName: form.officialName,
        officialTitle: form.officialTitle,
        reporterName: form.reporterName,
        reporterPhone: form.reporterPhone,
        address: form.stayLocation,
        submissionDate: form.submissionDate,
        submissionTime: form.submissionTime
      }
    };

    const existsInTemp = tempRecords.some(r => r.id === targetRecord.id);

    if (existsInTemp) {
      setTempRecords(prev => prev.map(r => r.id === targetRecord.id ? brandNewFcrRecord : r));
    } else {
      setTempRecords(prev => [brandNewFcrRecord, ...prev]);
    }
    
    syncMaster(form.officialName, 'Official');
    syncMaster(form.officialTitle, 'Title');
    syncMaster(form.reporterName, 'Reporter');
    syncMaster(form.reporterPhone, 'Phone');
    syncMaster(form.stayLocation, 'Stay', form.stayDescription);

    setFcrSuccessMsg(`${form.officialName} စာရင်းသွင်းသော Form C ကို စာရင်းသွင်းပြီးပါပြီ`);
    setTimeout(() => setFcrSuccessMsg(null), 3000);

    setFcrConfirmation(null);
    setTargetRecord(null);
    resetForm();
  };

  const submitTempFormC = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.status) {
      showToast("PLEASE SELECT STATUS (IN/OUT)");
      return;
    }

    const lastLog = [...records, ...tempRecords]
      .filter(r => r.passport.toUpperCase() === (targetRecord?.passport?.toUpperCase() || ""))
      .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];

    const logDate = lastLog ? formatDateToDDMMYYYY(lastLog.timestamp) : 'N/A';
    const logMode = lastLog?.mode || 'IN';
    const formCMode = form.status;

    let message = `လေဆိပ်မှ ${logDate}ရက်နေ့တွင် ${logMode === 'IN' ? 'အဝင်' : 'အထွက်'} ပြထားသူဖြစ်ပါသည်၊ Form C ကို ${formatDateToDDMMYYYY(form.submissionDate)} ရက်စွဲဖြင့် ${form.stayLocation} မှ ${formCMode === 'IN' ? 'အဝင်' : 'အထွက်'} လာရောက်တိုင်ကြားသည်မှာ မှန်ကန်ပါသလား။`;
    
    if (logMode !== formCMode) {
      if (logMode === 'OUT' && formCMode === 'IN') {
        message += " လေဆိပ်မဟုတ်သော အခြားနေရာမှ ဝင်ရောက်လာသူဖြစ်နိုင်ပါသည်။";
      } else if (logMode === 'IN' && formCMode === 'OUT') {
        message += " လေဆိပ်မဟုတ်သော အခြားနေရာမှ ထွက်ခွါသွားသူဖြစ်နိုင်ပါသည်။";
      }
    }

    setFcrConfirmation({ show: true, message, data: { isTemp: true } });
  };

  const executeSubmitTempFormC = () => {
    const parts = form.submissionDate.split('-');
    const slashDate = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : form.submissionDate;
    const timestamp = `${slashDate}, ${form.submissionTime || '00:00'}`;
    
    const newTempRecord: ImmRecord = {
      ...(targetRecord || {}),
      logType: 'FCR',
      id: targetRecord && tempRecords.some(tr => tr.id === targetRecord.id) && (!targetRecord.formC || targetRecord.formC.status === form.status) 
        ? targetRecord.id 
        : Date.now(),
      timestamp,
      mode: form.status as 'IN' | 'OUT',
      passport: form.passport.toUpperCase(),
      fullname: form.fullname,
      gender: form.gender as 'M' | 'F',
      nationality: form.nationality,
      address: form.stayLocation,
      visaType: form.visaType,
      stayFrom: form.stayFrom,
      stayTo: form.stayTo,
      totalDays: form.totalDays,
      remarks: form.remarks,
      arrivedFrom: form.status === 'OUT' ? 'MGZ' : form.arrivedFrom,
      departedTo: 'MGZ', // Removed departedTo selection in FCR, default to MGZ
      broughtBy: form.reporterName, 
      contactDetails: form.reporterPhone, 
      formC: {
        status: form.status as 'IN' | 'OUT',
        officialName: form.officialName,
        officialTitle: form.officialTitle,
        reporterName: form.reporterName,
        reporterPhone: form.reporterPhone,
        address: form.stayLocation,
        submissionDate: form.submissionDate,
        submissionTime: form.submissionTime
      }
    };

    if (targetRecord && tempRecords.some(tr => tr.id === targetRecord.id) && (!targetRecord.formC || targetRecord.formC.status === form.status)) {
      setTempRecords(prev => prev.map(tr => tr.id === targetRecord.id ? newTempRecord : tr));
    } else {
      setTempRecords(prev => [newTempRecord, ...prev]);
    }
    
    // Sync to Master
    syncMaster(form.nationality, 'Nationality');
    syncMaster(form.visaType, 'Visa');
    syncMaster(form.broughtBy, 'Agent');
    syncMaster(form.contactDetails, 'Contact');
    syncMaster(form.officialName, 'Official');
    syncMaster(form.officialTitle, 'Title');
    syncMaster(form.reporterName, 'Reporter');
    syncMaster(form.reporterPhone, 'Phone');
    syncMaster(form.vehicleInfo, 'Vehicle');
    syncMaster(form.stayLocation, 'Stay', form.stayDescription);

    setFcrSuccessMsg(`${form.officialName} စာရင်းသွင်းသော Form C ကို စာရင်းသွင်းပြီးပါပြီ`);
    setTimeout(() => setFcrSuccessMsg(null), 3000);

    setFcrConfirmation(null);
    setTargetRecord(null);
    setIsCreatingTemp(false);
    resetForm();
    showToast("TEMPORARY FORM C ENTRY SAVED");
    setSearch('');
  };

  const startTempEntry = () => {
    setIsCreatingTemp(true);
    setForm(prev => ({
      ...prev,
      fullname: search.match(/[a-zA-Z]/) ? search : '',
      passport: search.match(/[0-9]/) ? search : '',
      officialName: currentUser?.name || prev.officialName,
      officialTitle: currentUser?.title || prev.officialTitle
    }));
  };

  return (
    <div className="max-w-[1600px] mx-auto py-8 px-4 space-y-8 relative">
      <AnimatePresence>
        {fcrConfirmation && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 no-print">
             <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/80 backdrop-blur-md" />
             <motion.div 
               initial={{ scale: 0.9, opacity: 0, y: 20 }} 
               animate={{ scale: 1, opacity: 1, y: 0 }} 
               exit={{ scale: 0.9, opacity: 0, y: 20 }} 
               className="relative bg-white w-full max-w-lg rounded-[2.5rem] overflow-hidden shadow-2xl border-4 border-purple-500"
             >
                <div className="p-8 text-center space-y-6">
                  <div className="w-20 h-20 bg-purple-100 rounded-full flex items-center justify-center mx-auto text-purple-600">
                    <ShieldCheck size={40} />
                  </div>
                  <div className="space-y-4">
                    <h3 className="text-2xl font-black text-gray-900 leading-tight">သေချာပါသလား?</h3>
                    <p className="text-gray-600 font-medium leading-relaxed px-4">
                      {fcrConfirmation.message}
                    </p>
                  </div>
                  <div className="flex gap-4 pt-4">
                    <button 
                      onClick={() => setFcrConfirmation(null)}
                      className="flex-1 py-4 px-6 rounded-2xl bg-gray-100 text-gray-500 font-black uppercase text-sm hover:bg-gray-200 transition-colors"
                    >
                      ပြင်ဆင်ရန် (NO)
                    </button>
                    <button 
                      onClick={() => {
                        if (fcrConfirmation.data?.isTemp) executeSubmitTempFormC();
                        else executeSubmitFormC();
                      }}
                      className="flex-1 py-4 px-6 rounded-2xl bg-purple-600 text-white font-black uppercase text-sm hover:bg-purple-900 shadow-xl shadow-purple-200 transition-colors"
                    >
                      မှန်ကန်ပါသည်။ (YES)
                    </button>
                  </div>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {fcrSuccessMsg && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center pointer-events-none no-print">
             <motion.div 
               initial={{ scale: 0.5, opacity: 0, y: 50 }} 
               animate={{ scale: 1, opacity: 1, y: 0 }} 
               exit={{ scale: 1.5, opacity: 0 }} 
               className="bg-indigo-900 text-white px-10 py-6 rounded-full shadow-2xl flex items-center gap-4 border-2 border-white/20 backdrop-blur-xl"
             >
                <div className="w-10 h-10 bg-emerald-400 rounded-full flex items-center justify-center text-indigo-900">
                  <Check size={24} strokeWidth={4} />
                </div>
                <div className="flex flex-col">
                  <span className="text-[10px] font-black uppercase opacity-60 tracking-widest">Registration Successful</span>
                  <span className="text-lg font-black">{fcrSuccessMsg}</span>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      <div className="card shadow-2xl border-t-8 border-purple-600">
        <datalist id="officialList">{masterData.filter(m => m.type === 'Official').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="titleList">{masterData.filter(m => m.type === 'Title').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="reporterList">{masterData.filter(m => m.type === 'Reporter').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="phoneList">{masterData.filter(m => m.type === 'Phone').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="stayList">{masterData.filter(m => m.type === 'Stay').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="vehicleList">{masterData.filter(m => m.type === 'Vehicle').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="natList">{masterData.filter(m => m.type === 'Nationality').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="visaList">{masterData.filter(m => m.type === 'Visa').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="agentList">{masterData.filter(m => m.type === 'Agent').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="contactList">{masterData.filter(m => m.type === 'Contact').map(m => <option key={m.id} value={m.name} />)}</datalist>
        <datalist id="permitDescriptionList">{masterData.filter(m => m.type === 'PermitDescription').map(m => <option key={m.id} value={m.name} />)}</datalist>

        <div className="flex items-center gap-4 mb-8">
          <div className="bg-purple-100 p-3 rounded-full text-purple-700">
            <FileText size={28} />
          </div>
          <div>
            <h2 className="text-2xl font-black uppercase tracking-tight text-gray-800">Form C Registration</h2>
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Update Resident Status & Reporting</p>
          </div>
        </div>

        {isCreatingTemp ? (
          <form onSubmit={submitTempFormC} className="space-y-8 animate-in zoom-in-95 duration-300">
            <div className="bg-amber-50 p-4 rounded-2xl border border-amber-200 flex items-center gap-3">
              <AlertCircle size={24} className="text-amber-600" />
              <div>
                <h3 className="text-sm font-black uppercase text-amber-800">Temporary Entry Mode</h3>
                <p className="text-[10px] font-bold text-amber-600 uppercase">This record will be tagged as TEMPORARY ENTRY</p>
              </div>
              <button type="button" onClick={() => setIsCreatingTemp(false)} className="ml-auto text-amber-400 hover:text-amber-600">
                <X size={20} />
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8 border-t pt-8">
               <div className="space-y-4 p-4 bg-purple-50/30 rounded-2xl border border-purple-100">
                  <h4 className="text-xs font-black uppercase text-purple-700 pb-2 border-b border-purple-200 tracking-widest flex items-center gap-2">
                    <span className="bg-purple-700 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">1</span>
                    Movement & Official
                  </h4>
                  <div className="space-y-4">
                    <div className="flex gap-2">
                      <button 
                        type="button" 
                        onClick={() => setForm(prev => ({...prev, status: 'IN'}))}
                        className={`flex-1 py-3 rounded-lg border-2 font-black text-[10px] transition-all ${form.status === 'IN' ? 'bg-emerald-600 border-emerald-700 text-white shadow-md' : 'bg-white border-emerald-100 text-emerald-600 hover:bg-emerald-50'}`}
                      >
                        ENTRY (IN)
                      </button>
                      <button 
                        type="button" 
                        onClick={() => setForm(prev => ({...prev, status: 'OUT'}))}
                        className={`flex-1 py-3 rounded-lg border-2 font-black text-[10px] transition-all ${form.status === 'OUT' ? 'bg-orange-600 border-orange-700 text-white shadow-md' : 'bg-white border-orange-100 text-orange-600 hover:bg-orange-50'}`}
                      >
                        EXIT (OUT)
                      </button>
                    </div>
                    <div>
                      <label className="input-label">Official Name</label>
                      <input 
                        type="text" 
                        list="officialList" 
                        value={form.officialName} 
                        onChange={e => !currentUser && handleOfficialNameChange(e.target.value)} 
                        className={`input-field shadow-sm ${currentUser ? 'bg-gray-100 font-bold opacity-70' : ''}`}
                        required 
                        readOnly={!!currentUser}
                      />
                    </div>
                    <div>
                      <label className="input-label">Position / Title</label>
                      <input 
                        type="text" 
                        list="titleList" 
                        value={form.officialTitle} 
                        onChange={e => !currentUser && setForm(prev => ({...prev, officialTitle: e.target.value}))} 
                        className={`input-field shadow-sm ${currentUser ? 'bg-gray-100 font-bold opacity-70' : ''}`}
                        readOnly={!!currentUser}
                      />
                    </div>
                  </div>
               </div>

               <div className="space-y-4 p-4 bg-blue-50/30 rounded-2xl border border-blue-100">
                  <h4 className="text-xs font-black uppercase text-blue-700 pb-2 border-b border-blue-200 tracking-widest flex items-center gap-2">
                    <span className="bg-blue-700 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">2</span>
                    Personal Identity
                  </h4>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="col-span-2">
                       <label className="input-label">Full Name</label>
                       <input type="text" value={form.fullname} onChange={e => setForm(prev => ({...prev, fullname: e.target.value}))} className="input-field shadow-sm" required />
                    </div>
                    <div>
                       <label className="input-label">Passport</label>
                       <input type="text" value={form.passport} onChange={e => setForm(prev => ({...prev, passport: e.target.value.toUpperCase()}))} className="input-field shadow-sm font-black" required />
                    </div>
                    <div>
                       <label className="input-label">Nationality</label>
                       <input type="text" list="natList" value={form.nationality} onChange={e => setForm(prev => ({...prev, nationality: e.target.value}))} className="input-field shadow-sm" />
                    </div>
                    <div>
                       <label className="input-label">Gender</label>
                       <select value={form.gender} onChange={e => setForm(prev => ({...prev, gender: e.target.value as any}))} className="input-field shadow-sm">
                          <option value="M">Male</option>
                          <option value="F">Female</option>
                       </select>
                    </div>
                    <div>
                       <label className="input-label">Visa Type</label>
                       <input type="text" list="visaList" value={form.visaType} onChange={e => setForm(prev => ({...prev, visaType: e.target.value}))} className="input-field shadow-sm" />
                    </div>
                  </div>
               </div>

               <div className="space-y-4 p-4 bg-emerald-50/30 rounded-2xl border border-emerald-100">
                  <h4 className="text-xs font-black uppercase text-emerald-700 pb-2 border-b border-emerald-200 tracking-widest flex items-center gap-2">
                    <span className="bg-emerald-700 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">3</span>
                    Route & Stay
                  </h4>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                       <label className="input-label">Origin (From)</label>
                       <input 
                         type="text" 
                         list="stayList" 
                         value={form.status === 'OUT' ? 'MGZ' : form.arrivedFrom} 
                         onChange={e => setForm(prev => ({...prev, arrivedFrom: e.target.value}))} 
                         className={`input-field shadow-sm ${form.status === 'OUT' ? 'bg-gray-100 font-black cursor-not-allowed text-emerald-700' : ''}`}
                         readOnly={form.status === 'OUT'} 
                       />
                    </div>
                    <div>
                       <label className="input-label">Dest (To)</label>
                       <input 
                         type="text" 
                         list="stayList" 
                         value={form.status === 'IN' ? 'MGZ' : form.departedTo} 
                         onChange={e => setForm(prev => ({...prev, departedTo: e.target.value}))} 
                         className={`input-field shadow-sm ${form.status === 'IN' ? 'bg-gray-100 font-black cursor-not-allowed text-orange-700' : ''}`}
                         readOnly={form.status === 'IN'} 
                       />
                    </div>
                    <div className="col-span-2">
                       <label className="input-label">Stay Location</label>
                       <input 
                          type="text" 
                          list="stayList" 
                          value={form.stayLocation} 
                          onChange={e => {
                            const val = e.target.value;
                            const cleanVal = (val || '').trim().toLowerCase();
                            const all = [...records, ...tempRecords].sort((a, b) => getRecordTime(b) - getRecordTime(a));
                            const latestRec = all.find(r => (r.formC?.address?.toLowerCase() === cleanVal || r.address?.toLowerCase() === cleanVal) && (r.formC?.stayDescription || r.stayDescription));
                            const latestMaster = [...masterData].reverse().find(m => m.type === 'Stay' && m.name.toLowerCase() === cleanVal && m.linkedValue);
                            const desc = latestRec?.formC?.stayDescription || latestRec?.stayDescription || latestMaster?.linkedValue || '';
                            setForm(prev => ({
                              ...prev, 
                              stayLocation: val,
                              stayDescription: desc || prev.stayDescription
                            }));
                          }} 
                          className="input-field shadow-sm" 
                          required 
                        />
                     </div>
                     <div className="col-span-2">
                        <label className="input-label text-indigo-900 font-extrabold">🏡 Detail Address / Stay Description</label>
                        <input 
                          type="text" 
                          value={form.stayDescription || ''} 
                          onChange={e => setForm(prev => ({ ...prev, stayDescription: e.target.value }))} 
                          className="input-field shadow-sm bg-indigo-50/20 border-indigo-200" 
                          placeholder="Detail stay address or description..." 
                        />
                     </div>
                     <div className="col-span-2">
                        <label className="input-label pb-1 border-b mb-1 text-[10px]">Stay Period (From - Days - To)</label>
                        <div className="flex gap-2">
                           <input type="date" value={form.stayFrom} onChange={e => calculateTempStayDates('from', e.target.value)} className="input-field py-1 px-2 text-[10px]" />
                           <input type="text" value={form.totalDays} onChange={e => calculateTempStayDates('days', e.target.value)} className="input-field w-16 text-center font-bold text-indigo-700 text-[10px]" placeholder="Days" />
                           <input type="date" value={form.stayTo} onChange={e => calculateTempStayDates('to', e.target.value)} className="input-field py-1 px-2 text-[10px]" />
                        </div>
                     </div>
                  </div>
               </div>

               <div className="space-y-4 p-4 bg-orange-50/30 rounded-2xl border border-orange-100">
                  <h4 className="text-xs font-black uppercase text-orange-700 pb-2 border-b border-orange-200 tracking-widest flex items-center gap-2">
                    <span className="bg-orange-700 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">4</span>
                    Reporter & History
                  </h4>
                  <div className="space-y-4">
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="input-label">Reporter Name</label>
                        <input type="text" list="reporterList" value={form.reporterName} onChange={e => handleReporterNameChange(e.target.value)} className="input-field shadow-sm" required />
                      </div>
                      <div>
                        <label className="input-label">Phone</label>
                        <input type="text" list="phoneList" value={form.reporterPhone} onChange={e => handleReporterPhoneChange(e.target.value)} className="input-field shadow-sm" required />
                      </div>
                    </div>
                    <div>
                       <label className="input-label">Remarks</label>
                       <input type="text" value={form.remarks} onChange={e => setForm(prev => ({...prev, remarks: e.target.value}))} className="input-field shadow-sm" placeholder="Notes..." />
                    </div>
                  </div>
               </div>
            </div>

            <div className="pt-6 border-t flex gap-4">
               <button type="button" onClick={() => setIsCreatingTemp(false)} className="btn flex-1 bg-gray-100 text-gray-600 hover:bg-gray-200 font-black uppercase">Cancel</button>
               <button type="submit" className="btn flex-1 bg-amber-600 text-white hover:bg-black shadow-xl font-black uppercase">Submit Temporary Entry</button>
            </div>
          </form>
        ) : !targetRecord ? (
          <div className="space-y-6">
            <div className="relative">
              <label className="input-label">Search Entry Record (Name or Passport)</label>
              <div className="relative group">
                <input 
                  type="text" 
                  value={search} 
                  onChange={(e) => setSearch(e.target.value)} 
                  className="input-field pl-12 h-14 text-lg font-bold border-purple-200 focus:border-purple-500 transition-all text-gray-800" 
                  placeholder="Type name or passport..." 
                />
                <Search size={24} className="absolute left-4 top-4 text-purple-300 group-focus-within:text-purple-600 transition-colors" />
                {search && (
                  <button 
                    type="button"
                    onClick={() => setSearch('')}
                    className="absolute right-4 top-4 text-gray-400 hover:text-red-500 transition-colors"
                  >
                    <X size={24} strokeWidth={3} />
                  </button>
                )}
              </div>
              
              {searchResults.length > 0 && (
                <div className="absolute z-10 w-full mt-2 bg-white border border-gray-100 shadow-2xl rounded-2xl overflow-hidden animate-in fade-in slide-in-from-top-2">
                  {searchResults.map(r => (
                    <button 
                      key={r.id} 
                      onClick={() => selectPerson(r)}
                      className="w-full p-4 hover:bg-purple-50 text-left border-b last:border-0 transition-colors flex justify-between items-center group"
                    >
                      <div>
                        <div className="font-black text-gray-900 group-hover:text-purple-700">{r.fullname}</div>
                        <div className="text-xs text-gray-400 font-mono">{r.passport} | {r.nationality}</div>
                      </div>
                      <div className="text-right">
                        <div className="text-[10px] font-black text-gray-300 uppercase">{r.timestamp}</div>
                        <div className="text-xs font-bold text-gray-500">{r.address}</div>
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </div>
            {search.length > 0 && searchResults.length === 0 && (
              <div className="text-center py-10 bg-gray-50 rounded-2xl border-2 border-dashed font-bold text-gray-400 space-y-4">
                <p>No matching records found for "{search}"</p>
                <button 
                  onClick={startTempEntry}
                  className="btn bg-purple-600 text-white hover:bg-black uppercase text-xs"
                >
                  Create Temporary Form C Registration Entry
                </button>
              </div>
            )}
          </div>
        ) : (
          <form onSubmit={submitFormC} className="space-y-8 animate-in zoom-in-95 duration-300">
            <div className="bg-purple-50 p-6 rounded-2xl border border-purple-100 flex justify-between items-center">
              <div>
                <h3 className="text-sm font-black uppercase text-purple-800 mb-1">Target Person</h3>
                <div className="text-xl font-black text-gray-900">{targetRecord.fullname}</div>
                <div className="text-xs font-bold text-purple-600">{targetRecord.passport} | {targetRecord.nationality}</div>
                
                {/* Movement Summary in Burmese */}
                <div className="mt-4 p-3 bg-white/80 rounded-xl border border-purple-100 text-[11px] leading-relaxed text-purple-900 font-bold shadow-sm">
                   <div>({targetRecord.timestamp.split(',')[0]}) တွင် ({targetRecord.arrivedFrom || 'N/A'})မှ ({targetRecord.departedTo || 'N/A'})သို့ {targetRecord.mode === 'IN' ? 'ဝင်ရောက်လာ' : 'ထွက်ခွါသွား'}သူဖြစ်ပါသည်။</div>
                   <div className="mt-1">({targetRecord.address || 'N/A'}) တွင် {targetRecord.mode === 'IN' ? 'နေထိုင်မည်' : 'နေထိုင်သူ'} ဟုသိရပါသည်။</div>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setTargetRecord(null)}
                className="btn bg-white text-gray-400 hover:text-red-500 border-0 shadow-none hover:bg-red-50"
              >
                <X size={20} />
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl border border-purple-100 shadow-sm space-y-4">
                  <h4 className="text-xs font-black uppercase text-purple-700 border-b pb-2 tracking-widest flex items-center gap-2">
                    <MapPin size={14} /> Registration Details
                  </h4>
                  <div className="grid grid-cols-1 gap-4">
                    <div>
                      <label className="input-label">Stay Location (Can be updated)</label>
                      <input 
                        type="text" 
                        list="stayList"
                        value={form.stayLocation} 
                        onChange={(e) => {
                          const val = e.target.value;
                          const cleanVal = (val || '').trim().toLowerCase();
                          const all = [...records, ...tempRecords].sort((a, b) => getRecordTime(b) - getRecordTime(a));
                          const latestRec = all.find(r => (r.formC?.address?.toLowerCase() === cleanVal || r.address?.toLowerCase() === cleanVal) && (r.formC?.stayDescription || r.stayDescription));
                          const latestMaster = [...masterData].reverse().find(m => m.type === 'Stay' && m.name.toLowerCase() === cleanVal && m.linkedValue);
                          const desc = latestRec?.formC?.stayDescription || latestRec?.stayDescription || latestMaster?.linkedValue || '';
                          setForm(prev => ({ 
                            ...prev, 
                            stayLocation: val,
                            stayDescription: desc || prev.stayDescription
                          }));
                        }} 
                        className="input-field border-purple-100 placeholder:text-gray-300" 
                        placeholder="Enter location..."
                      />
                    </div>
                    <div>
                      <label className="input-label text-purple-900 font-extrabold">🏡 Detail Address / Stay Description (လိပ်စာအသေးစိတ်)</label>
                      <input 
                        type="text" 
                        value={form.stayDescription || ''} 
                        onChange={(e) => setForm(prev => ({ ...prev, stayDescription: e.target.value }))} 
                        className="input-field border-purple-150 bg-purple-50/20" 
                        placeholder="Detail stay address or description..." 
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="input-label">Submission Date</label>
                        <input 
                          type="date" 
                          value={form.submissionDate} 
                          onChange={(e) => setForm(prev => ({ ...prev, submissionDate: e.target.value }))} 
                          className="input-field border-purple-100" 
                        />
                      </div>
                      <div>
                        <label className="input-label">Submission Time</label>
                        <input 
                          type="time" 
                          value={form.submissionTime} 
                          onChange={(e) => setForm(prev => ({ ...prev, submissionTime: e.target.value }))} 
                          className="input-field border-purple-100" 
                        />
                      </div>
                    </div>
                    <div>
                      <label className="input-label">Form C Status</label>
                      <div className="flex bg-gray-100 p-1 rounded-xl">
                        <button 
                          type="button"
                          onClick={() => setForm(prev => ({ ...prev, status: 'IN' }))}
                          className={`flex-1 py-3 rounded-lg text-xs font-black transition-all ${form.status === 'IN' ? 'bg-white shadow text-emerald-600' : 'text-gray-400 hover:text-gray-600'}`}
                        >
                          FORMC IN
                        </button>
                        <button 
                          type="button"
                          onClick={() => setForm(prev => ({ ...prev, status: 'OUT' }))}
                          className={`flex-1 py-3 rounded-lg text-xs font-black transition-all ${form.status === 'OUT' ? 'bg-white shadow text-orange-600' : 'text-gray-400 hover:text-gray-600'}`}
                        >
                          FORMC OUT
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-4">
                <h4 className="text-xs font-black uppercase text-gray-400 border-b pb-2 tracking-widest">Authority & Reporter</h4>
                <div className="grid grid-cols-1 gap-4">
                  <div>
                    <label className="input-label">Immigration Official Name</label>
                    <input 
                      list="officialList" 
                      type="text" 
                      value={form.officialName} 
                      onChange={(e) => handleOfficialNameChange(e.target.value)} 
                      className={`input-field ${!!currentUser ? 'bg-gray-100 cursor-not-allowed font-black text-slate-700 border-dashed border-gray-300' : ''}`} 
                      readOnly={!!currentUser}
                      required 
                    />
                  </div>
                  <div>
                    <label className="input-label">Position / Title</label>
                    <input 
                      list="titleList" 
                      type="text" 
                      value={form.officialTitle} 
                      onChange={(e) => setForm(prev => ({ ...prev, officialTitle: e.target.value }))} 
                      className={`input-field ${!!currentUser ? 'bg-gray-100 cursor-not-allowed font-black text-slate-700 border-dashed border-gray-300' : ''}`} 
                      readOnly={!!currentUser}
                    />
                  </div>
                  <div>
                    <label className="input-label">Reporter Name</label>
                    <input list="reporterList" type="text" value={form.reporterName} onChange={(e) => handleReporterNameChange(e.target.value)} className="input-field" required />
                  </div>
                  <div>
                    <label className="input-label">Reporter Phone Number</label>
                    <input list="phoneList" type="text" value={form.reporterPhone} onChange={(e) => setForm(prev => ({ ...prev, reporterPhone: e.target.value }))} className="input-field" />
                  </div>
                </div>
              </div>
            </div>

            <div className="pt-6 border-t flex gap-4">
               <button type="button" onClick={() => setTargetRecord(null)} className="btn flex-1 bg-gray-100 text-gray-600 hover:bg-gray-200 uppercase font-black">Cancel</button>
               <button type="submit" className="btn flex-1 bg-purple-700 text-white hover:bg-purple-800 shadow-xl shadow-purple-100 uppercase font-black">Confirm Registration</button>
            </div>
          </form>
        )}
      </div>

      <div className="card shadow-xl border-t-8 border-purple-900 overflow-hidden">
        <div className="flex flex-col md:flex-row justify-between items-center gap-4 mb-8 text-slate-800">
           <div className="flex items-center gap-3">
              <div className="bg-purple-900 text-white p-2.5 rounded-xl shadow-lg">
                 <History size={20} />
              </div>
              <h3 className="text-xl font-black uppercase tracking-tight">Immigration & Form C Complete History</h3>
           </div>

           <div className="flex flex-1 max-w-md w-full relative group">
             <input 
               type="text" 
               placeholder="Search logs by name, passport..." 
               value={logSearch}
               onChange={(e) => setLogSearch(e.target.value)}
               id="fcr-log-search"
               name="fcr-log-search"
               autoComplete="off"
               className="input-field pl-10 pr-10 text-xs h-10 w-full font-bold text-slate-800 bg-slate-50 border-slate-200 focus:border-purple-600 focus:bg-white transition-all outline-none shadow-sm"
             />
             <Search size={16} className="absolute left-3.5 top-3 text-slate-400" />
             {logSearch && (
               <button 
                 onClick={() => setLogSearch('')}
                 className="absolute right-3 top-2.5 text-slate-400 hover:text-red-500 transition-colors"
               >
                 <X size={18} />
               </button>
             )}
           </div>

           <button 
             onClick={exportFCRLogsToExcel} 
             className="btn bg-emerald-600 text-white hover:bg-black text-[10px] font-black uppercase flex items-center gap-2 px-6 shadow-lg shadow-emerald-50"
           >
             <Download size={16} /> Export Excel
           </button>
        </div>

        <div className="overflow-x-auto border rounded-2xl">
           <table className="w-full text-left text-sm">
              <thead className="bg-purple-900 text-white font-black uppercase tracking-widest text-[10px]">
                 <tr>
                    <th className="p-4 border-r border-white/10 text-center">No</th>
                    <th className="p-4 border-r border-white/10">Identity (Passport & Name)</th>
                    <th className="p-4 border-r border-white/10 text-center">Log IN</th>
                    <th className="p-4 border-r border-white/10 text-center">Log OUT</th>
                    <th className="p-4 border-r border-white/10 text-center">Form C IN</th>
                    <th className="p-4 border-r border-white/10 text-center">Form C OUT</th>
                    <th className="p-4 border-white/10">Official / Reporter (Logger)</th>
                  </tr>
               </thead>
               <tbody className="divide-y divide-purple-50">
                 {formCLogs.map((g, i) => (
                   <tr key={g.passport} className="hover:bg-purple-50/50 transition-colors group">
                      <td className="p-4 border-r font-black text-gray-300 text-center">{i + 1}</td>
                      <td className="p-4 border-r">
                         <div className="font-black text-gray-900">{g.fullname}</div>
                         <div className="flex flex-col">
                            <div className="text-[10px] text-gray-400 font-mono italic">{g.passport}</div>
                            <div className="text-[10px] text-indigo-600 font-bold flex items-center gap-1 mt-1">
                               <MapPin size={10} /> {g.formCIn?.formC?.address || g.formCOut?.formC?.address || g.lastRecord.address || 'N/A'}
                            </div>
                         </div>
                      </td>
                      
                      {/* Log IN Cell */}
                      <td className="p-4 border-r text-center group/cell relative min-w-[120px]">
                         {g.logIn ? (
                           <div className="space-y-1">
                              <span className="px-2 py-0.5 rounded text-[10px] font-black border bg-emerald-50 text-emerald-800 border-emerald-100 uppercase">IN</span>
                              <div className="text-[9px] text-gray-500 font-black">{formatDateToDDMMYYYY(g.logIn.timestamp)}</div>
                              <div className="text-[10px] text-emerald-600 font-bold italic line-clamp-1">{g.logIn.address || 'N/A'}</div>
                              
                              <div className="flex flex-col gap-1 mt-2 transition-opacity">
                                 <button onClick={() => editLog(g.logIn)} className={`flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white transition-colors`}>
                                    <Edit size={10} /> Edit
                                 </button>
                                 {deleteConfirmId === g.logIn.id ? (
                                   <div className="flex gap-1">
                                     <button onClick={() => { deleteRecord(g.logIn.id); setDeleteConfirmId(null); }} className="flex-1 text-[8px] font-black bg-red-600 text-white rounded py-1">SURE?</button>
                                     <button onClick={() => setDeleteConfirmId(null)} className="flex-1 text-[8px] font-black bg-gray-200 text-gray-600 rounded py-1">NO</button>
                                   </div>
                                 ) : (
                                   <button onClick={() => setDeleteConfirmId(g.logIn.id)} className="flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-red-50 text-red-600 hover:bg-red-600 hover:text-white transition-colors">
                                      <Trash2 size={10} /> Delete
                                   </button>
                                 )}
                              </div>
                           </div>
                         ) : <span className="text-gray-300 font-bold italic opacity-30">N/A</span>}
                      </td>

                      {/* Log OUT Cell */}
                      <td className="p-4 border-r text-center group/cell relative min-w-[120px]">
                         {g.logOut ? (
                           <div className="space-y-1">
                              <span className="px-2 py-0.5 rounded text-[10px] font-black border bg-orange-50 text-orange-800 border-orange-100 uppercase">OUT</span>
                              <div className="text-[9px] text-gray-500 font-black">{formatDateToDDMMYYYY(g.logOut.timestamp)}</div>
                              <div className="text-[10px] text-orange-600 font-bold italic line-clamp-1">{g.logOut.address || 'N/A'}</div>
                              
                              <div className="flex flex-col gap-1 mt-2 transition-opacity">
                                 <button onClick={() => editLog(g.logOut)} className={`flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white transition-colors`}>
                                    <Edit size={10} /> Edit
                                 </button>
                                 {deleteConfirmId === g.logOut.id ? (
                                   <div className="flex gap-1">
                                     <button onClick={() => { deleteRecord(g.logOut.id); setDeleteConfirmId(null); }} className="flex-1 text-[8px] font-black bg-red-600 text-white rounded py-1">SURE?</button>
                                     <button onClick={() => setDeleteConfirmId(null)} className="flex-1 text-[8px] font-black bg-gray-200 text-gray-600 rounded py-1">NO</button>
                                   </div>
                                 ) : (
                                   <button onClick={() => setDeleteConfirmId(g.logOut.id)} className="flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-red-50 text-red-600 hover:bg-red-600 hover:text-white transition-colors">
                                      <Trash2 size={10} /> Delete
                                   </button>
                                 )}
                              </div>
                           </div>
                         ) : <span className="text-gray-300 font-bold italic opacity-30">N/A</span>}
                      </td>

                      {/* Form C IN Cell */}
                      <td className="p-4 border-r text-center group/cell relative min-w-[120px]">
                         {g.formCIn ? (
                           <div className="space-y-1">
                              <span className="px-2 py-0.5 rounded text-[10px] font-black border bg-purple-900 text-white uppercase shadow-sm">IN</span>
                              <div className="text-[9px] text-purple-900 font-black">{formatDateToDDMMYYYY(g.formCIn.timestamp)}</div>
                              <div className="text-[10px] text-purple-400 font-bold italic line-clamp-1">{g.formCIn.formC?.address || g.formCIn.address || 'N/A'}</div>
                              
                              <div className="flex flex-col gap-1 mt-2 transition-opacity">
                                 <button onClick={() => editLog(g.formCIn)} className={`flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white transition-colors`}>
                                    <Edit size={10} /> Edit
                                 </button>
                                 {formCDeleteId === g.formCIn.id ? (
                                   <div className="flex gap-1">
                                     <button onClick={() => { deleteFormCData(g.formCIn.id); setFormCDeleteId(null); }} className="flex-1 text-[8px] font-black bg-orange-600 text-white rounded py-1">CLEAR?</button>
                                     <button onClick={() => setFormCDeleteId(null)} className="flex-1 text-[8px] font-black bg-gray-200 text-gray-600 rounded py-1">NO</button>
                                   </div>
                                 ) : (
                                   <button onClick={() => setFormCDeleteId(g.formCIn.id)} className="flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-orange-50 text-orange-600 hover:bg-orange-600 hover:text-white transition-colors">
                                      <RefreshCcw size={10} /> Clear
                                   </button>
                                 )}
                              </div>
                           </div>
                         ) : <span className="text-gray-300 font-bold italic opacity-30">N/A</span>}
                      </td>

                      {/* Form C OUT Cell */}
                      <td className="p-4 border-r text-center group/cell relative min-w-[120px]">
                         {g.formCOut ? (
                           <div className="space-y-1">
                              <span className="px-2 py-0.5 rounded text-[10px] font-black border bg-purple-900 text-white uppercase shadow-sm">OUT</span>
                              <div className="text-[9px] text-purple-900 font-black">{formatDateToDDMMYYYY(g.formCOut.timestamp)}</div>
                              <div className="text-[10px] text-purple-400 font-bold italic line-clamp-1">{g.formCOut.formC?.address || g.formCOut.address || 'N/A'}</div>
                              
                              <div className="flex flex-col gap-1 mt-2 transition-opacity">
                                 <button onClick={() => editLog(g.formCOut)} className={`flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white transition-colors`}>
                                    <Edit size={10} /> Edit
                                 </button>
                                 {formCDeleteId === g.formCOut.id ? (
                                   <div className="flex gap-1">
                                     <button onClick={() => { deleteFormCData(g.formCOut.id); setFormCDeleteId(null); }} className="flex-1 text-[8px] font-black bg-orange-600 text-white rounded py-1">CLEAR?</button>
                                     <button onClick={() => setFormCDeleteId(null)} className="flex-1 text-[8px] font-black bg-gray-200 text-gray-600 rounded py-1">NO</button>
                                   </div>
                                 ) : (
                                   <button onClick={() => setFormCDeleteId(g.formCOut.id)} className="flex items-center justify-center gap-1 text-[8px] font-black uppercase py-1 px-2 rounded bg-orange-50 text-orange-600 hover:bg-orange-600 hover:text-white transition-colors">
                                      <RefreshCcw size={10} /> Clear
                                   </button>
                                 )}
                              </div>
                           </div>
                         ) : <span className="text-gray-300 font-bold italic opacity-30">N/A</span>}
                      </td>

                      <td className="p-4">
                        <div className="flex flex-col gap-1.5 min-w-[150px]">
                           { (g.formCOut?.formC || g.formCIn?.formC) ? (
                               <div className="space-y-1">
                                 <div className="flex items-center gap-1.5">
                                   <span className="w-1.5 h-1.5 rounded-full bg-purple-500"></span>
                                   <span className="font-bold text-gray-700">{(g.formCOut?.formC || g.formCIn?.formC).officialName}</span>
                                   <span className="text-[8px] bg-purple-100 text-purple-700 px-1 rounded font-black">{(g.formCOut?.formC || g.formCIn?.formC).officialTitle}</span>
                                 </div>
                                 <div className="flex items-center gap-1.5 opacity-70">
                                   <span className="w-1.5 h-1.5 rounded-full bg-gray-400"></span>
                                   <span className="text-[10px] font-bold">{(g.formCOut?.formC || g.formCIn?.formC).reporterName}</span>
                                   <span className="text-[9px] font-mono">{(g.formCOut?.formC || g.formCIn?.formC).reporterPhone}</span>
                                 </div>
                               </div>
                           ) : (
                             <span className="text-gray-300 font-bold italic">N/A</span>
                           )}
                        </div>
                      </td>
                   </tr>
                 ))}
                  {formCLogs.length === 0 && (
                   <tr>
                     <td colSpan={7} className="p-12 text-center text-gray-400 italic">No activity logs found.</td>
                   </tr>
                 )}
              </tbody>
           </table>
        </div>
      </div>
    </div>
  );
};

const DailyReport = ({ 
  records, 
  showToast, 
  movementMap,
  setActivePrintPreview,
  dailyPdfs,
  setDailyPdfs
}: { 
  records: ImmRecord[]; 
  showToast: (msg: string) => void; 
  movementMap: Record<string, MovementData>;
  setActivePrintPreview: React.Dispatch<React.SetStateAction<any>>;
  dailyPdfs: Record<string, { base64: string; name: string }>;
  setDailyPdfs: React.Dispatch<React.SetStateAction<Record<string, { base64: string; name: string }>>>;
}) => {
  const today = new Date().toISOString().split('T')[0];
  const [reportDate, setReportDate] = useState(today);

  const [y, m, d] = reportDate.split('-');
  const searchStr = `${d}/${m}/${y}`;
  const dailyIn = records.filter(r => r.mode === 'IN' && (r.logType === 'FFE' || !r.logType) && r.timestamp.startsWith(searchStr)).sort((a,b) => parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp));
  const dailyOut = records.filter(r => r.mode === 'OUT' && (r.logType === 'FFE' || !r.logType) && r.timestamp.startsWith(searchStr)).sort((a,b) => parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp));

  const exportDailyToExcel = (data: ImmRecord[], mode: string) => {
    const excelData = data.map((r, i) => {
      const stayedFrom = r.mode === 'IN' ? r.timestamp.split(', ')[0] : (records.filter(rec => (rec.logType === 'FFE' || !rec.logType) && rec.passport === r.passport && rec.mode === 'IN' && parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0]?.timestamp.split(', ')[0] || '-');
      const stayedTo = r.mode === 'IN' ? '' : r.timestamp.split(', ')[0];
      return {
        "No": i + 1,
        "Name": r.fullname,
        "G": r.gender,
        "Nationality": r.nationality,
        "Passport Number": r.passport,
        "Visa Type": r.visaType,
        "Stay Period": r.stayFrom ? `${formatFriendlyDate(r.stayFrom)} to ${formatFriendlyDate(r.stayTo)}` : '-',
        "Stayed From": stayedFrom,
        "Stayed To": stayedTo,
        "Route (From)": r.arrivedFrom,
        "Route (To)": r.departedTo,
        "Vehicle Number": r.vehicleInfo,
        "Stayed Address": r.address,
        "Log Date and Time Stamp": r.timestamp
      };
    });
    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `Daily_${mode}_Report_${reportDate}.xlsx`);
  };

  const combinedExcelExport = () => {
    const wb = XLSX.utils.book_new();

    const generateSheetData = (data: ImmRecord[]) => data.map((r, i) => {
      const stayedFrom = r.mode === 'IN' ? r.timestamp.split(', ')[0] : (records.filter(rec => (rec.logType === 'FFE' || !rec.logType) && rec.passport === r.passport && rec.mode === 'IN' && parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0]?.timestamp.split(', ')[0] || '-');
      const stayedTo = r.mode === 'IN' ? '' : r.timestamp.split(', ')[0];
      return {
        "No": i + 1,
        "Name": r.fullname,
        "G": r.gender,
        "Nationality": r.nationality,
        "Passport Number": r.passport,
        "Visa Type": r.visaType,
        "Stay Period": r.stayFrom ? `${formatFriendlyDate(r.stayFrom)} to ${formatFriendlyDate(r.stayTo)}` : '-',
        "Stayed From": stayedFrom,
        "Stayed To": stayedTo,
        "Route (From)": r.arrivedFrom,
        "Route (To)": r.departedTo,
        "Vehicle Number": r.vehicleInfo,
        "Stayed Address": r.address,
        "Log Date and Time Stamp": r.timestamp
      };
    });

    const wsIn = XLSX.utils.json_to_sheet(generateSheetData(dailyIn));
    XLSX.utils.book_append_sheet(wb, wsIn, "Arrival (IN)");

    const wsOut = XLSX.utils.json_to_sheet(generateSheetData(dailyOut));
    XLSX.utils.book_append_sheet(wb, wsOut, "Departure (OUT)");

    XLSX.writeFile(wb, `Daily_Combined_Report_${reportDate}.xlsx`);
    showToast("COMBINED EXCEL EXPORTED");
  };

  const copyToClipboard = (data: ImmRecord[], mode: string) => {
    let text = `Daily ${mode} Report - ${searchStr}\n\n`;
    text += "No\tName\tG\tNationality\tPassport\tVisa\tStay Period\tStayed From\tStayed To\tRoute(From)\tRoute(To)\tVehicle\tAddress\tLog Stamp\n";
    data.forEach((r, i) => {
      const stayedFrom = r.mode === 'IN' ? r.timestamp.split(', ')[0] : (records.filter(rec => (rec.logType === 'FFE' || !rec.logType) && rec.passport === r.passport && rec.mode === 'IN' && parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0]?.timestamp.split(', ')[0] || '-');
      const stayedTo = r.mode === 'IN' ? '' : r.timestamp.split(', ')[0];
      const stayPeriod = r.stayFrom ? `${formatFriendlyDate(r.stayFrom)} to ${formatFriendlyDate(r.stayTo)}` : '-';
      text += `${i+1}\t${r.fullname}\t${r.gender}\t${r.nationality}\t${r.passport}\t${r.visaType}\t${stayPeriod}\t${stayedFrom}\t${stayedTo}\t${r.arrivedFrom}\t${r.departedTo}\t${r.vehicleInfo}\t${r.address}\t${r.timestamp}\n`;
    });
    navigator.clipboard.writeText(text).then(() => showToast("COPIED TO CLIPBOARD"));
  };

  const generateAnalysisTable = (data: ImmRecord[], title: string, type: 'nationality' | 'visaType') => {
    const analysis: Record<string, { label: string; M: number; F: number; T: number }> = {};
    data.forEach(r => {
      const val = r[type] || 'Unknown';
      if (!analysis[val]) analysis[val] = { label: val, M: 0, F: 0, T: 0 };
      if (r.gender === 'M') analysis[val].M++;
      else if (r.gender === 'F') analysis[val].F++;
      analysis[val].T++;
    });

    const sorted = Object.values(analysis).sort((a, b) => b.T - a.T);
    const total = sorted.reduce((acc, curr) => ({ M: acc.M + curr.M, F: acc.F + curr.F, T: acc.T + curr.T }), { M: 0, F: 0, T: 0 });

    if (sorted.length === 0) return null;

    return (
      <div className="mb-8">
        <h4 className="text-sm font-black uppercase text-gray-500 mb-2 px-1 tracking-widest">{title} Analysis</h4>
        <div className="overflow-x-auto border border-black rounded shadow-sm">
          <table className="report-table text-[11px]">
            <thead>
              <tr className="bg-gray-100 uppercase">
                <th className="text-left py-2 px-3">Category</th>
                <th className="w-16 py-2">M</th>
                <th className="w-16 py-2">F</th>
                <th className="w-16 py-2">Total</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((s, idx) => (
                <tr key={idx}>
                  <td className="text-left py-2 px-3 font-bold">{s.label}</td>
                  <td className="py-2 text-center">{s.M}</td>
                  <td className="py-2 text-center">{s.F}</td>
                  <td className="py-2 text-center font-black bg-indigo-50/30">{s.T}</td>
                </tr>
              ))}
            </tbody>
            <tfoot className="bg-gray-900 text-white font-black">
               <tr>
                 <td className="text-left py-2 px-3 uppercase tracking-widest">TOTAL SUMMARY</td>
                 <td className="py-2 text-center">{total.M}</td>
                 <td className="py-2 text-center">{total.F}</td>
                 <td className="py-2 text-center border-l border-white/20">{total.T}</td>
               </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  };

  const generateTable = (data: ImmRecord[]) => {
    if (data.length === 0) return <div className="p-8 text-center text-gray-400 italic border rounded-xl bg-gray-50">NO DATA LOGGED FOR THIS DATE.</div>;
    
    return (
      <div className="overflow-x-auto border border-black rounded shadow-sm mb-8">
        <table className="report-table">
          <thead>
            <tr>
              <th style={{ width: '40px' }}>No</th>
              <th>Name</th>
              <th style={{ width: '30px' }}>G</th>
              <th>Nationality</th>
              <th>Passport number</th>
              <th>Visa type</th>
              <th style={{ width: '180px' }}>Stay period</th>
              <th>Stayed from</th>
              <th>Stayed to</th>
              <th>Route (From)</th>
              <th>Route (To)</th>
              <th>Vehicle number</th>
              <th style={{ width: '150px' }}>Stayed address</th>
              <th style={{ width: '140px' }}>Log date and time stamp</th>
            </tr>
          </thead>
          <tbody>
            {data.map((r, i) => {
              const stayPeriod = r.stayFrom ? `${formatFriendlyDate(r.stayFrom)} to ${formatFriendlyDate(r.stayTo)}` : '-';
              
              let stayedFrom = '';
              let stayedTo = '';
              
              if (r.mode === 'IN') {
                stayedFrom = r.timestamp.split(', ')[0];
                stayedTo = '';
              } else {
                // For OUT, grab from last log IN date
                const lastIn = records.filter(rec => 
                  (rec.logType === 'FFE' || !rec.logType) &&
                  rec.passport === r.passport && 
                  rec.mode === 'IN' && 
                  parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)
                ).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
                
                stayedFrom = lastIn ? lastIn.timestamp.split(', ')[0] : '-';
                stayedTo = r.timestamp.split(', ')[0];
              }

              return (
                <tr key={r.id}>
                  <td className="font-bold">{i+1}</td>
                  <td className="text-left font-bold">{r.fullname}</td>
                  <td>{r.gender}</td>
                  <td className="font-bold">{r.nationality || '-'}</td>
                  <td className="font-bold">{r.passport}</td>
                  <td>{r.visaType}</td>
                  <td>{stayPeriod}</td>
                  <td>{stayedFrom}</td>
                  <td>{stayedTo}</td>
                  <td>{r.arrivedFrom}</td>
                  <td>{r.departedTo}</td>
                  <td className="font-bold">{r.vehicleInfo}</td>
                  <td className="text-left">{r.address}</td>
                  <td>{r.timestamp}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="max-w-[1600px] mx-auto py-8 px-4 space-y-12">
      <div className="card shadow-xl border-t-8 border-indigo-700 no-print">
        <div className="flex flex-col md:flex-row justify-between items-center gap-6">
          <div className="flex items-center gap-4">
             <div className="bg-indigo-100 p-3 rounded-full text-indigo-700">
                <Calendar size={28} />
             </div>
             <div>
               <h2 className="text-2xl font-black text-gray-800 uppercase tracking-tight">Daily Movement Review</h2>
               <p className="text-gray-400 text-xs font-bold font-mono">GENERATE FINAL LOG REPORTS</p>
             </div>
          </div>
          <div className="flex flex-wrap gap-3">
            <input type="date" value={reportDate} onChange={(e) => setReportDate(e.target.value)} className="input-field border-2 border-indigo-200 font-bold focus:ring-indigo-500 w-48" />
            <button onClick={combinedExcelExport} className="btn bg-emerald-700 text-white hover:bg-emerald-800"><Download size={16} /> Combined Excel</button>
            <button 
               onClick={() => {
                 setActivePrintPreview({
                   title: `Daily Movement Log Report - ${reportDate}`,
                   type: 'DAILY_REPORT',
                   data: [
                     ...dailyIn.map(x => ({ ...x, mode: 'IN' })),
                     ...dailyOut.map(x => ({ ...x, mode: 'OUT' }))
                   ],
                   meta: { reportDate }
                 });
               }} 
               className="btn bg-gray-900 text-white hover:bg-black"
             >
               <Printer size={16} /> Print Report
             </button>
          </div>
        </div>
      </div>

      {/* Daily Raw History Attachment Slot */}
      <div className="card shadow-md border-t-4 border-emerald-600 bg-white p-6 no-print max-w-[1600px] mx-auto">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
          <div>
            <h3 className="text-base font-black text-emerald-800 uppercase flex items-center gap-2">
              <FileText size={18} /> Daily Raw Records History File Link
            </h3>
            <p className="text-gray-400 text-xs font-bold font-mono">UPLOAD THE MAIN SCAN / PDF LOG OF THE ENTIRE DAY RECORD SHEETS FOR CO-RELATING WITH LOG DATA</p>
          </div>
          
          <div className="w-full lg:w-auto flex flex-col sm:flex-row items-stretch sm:items-center gap-4">
            {dailyPdfs && dailyPdfs[reportDate] ? (
              <div className="flex items-center gap-3 bg-indigo-50 border-2 border-indigo-200/60 p-3 rounded-2xl">
                <FileText className="text-indigo-600 shrink-0" size={24} />
                <div className="min-w-0 pr-2">
                  <div className="text-[10px] font-black text-indigo-700 uppercase">Attached raw PDF file</div>
                  <div className="text-xs font-bold text-slate-800 truncate max-w-xs" title={dailyPdfs[reportDate].name}>
                    {dailyPdfs[reportDate].name}
                  </div>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => {
                      const fileObj = dailyPdfs[reportDate];
                      if (fileObj) {
                        const link = document.createElement('a');
                        link.href = fileObj.base64;
                        link.download = fileObj.name;
                        link.click();
                        showToast(`Downloading: ${fileObj.name}`);
                      }
                    }}
                    title="Download attached PDF"
                    className="bg-white hover:bg-emerald-600 hover:text-white text-emerald-600 px-3 py-1.5 rounded-lg border font-black uppercase text-[9px] tracking-wider shrink-0 transition-all flex items-center gap-1"
                  >
                    <Download size={11} /> Download
                  </button>
                  <button
                    onClick={() => {
                      const next = { ...dailyPdfs };
                      delete next[reportDate];
                      setDailyPdfs(next);
                      showToast(`Removed attachment for ${reportDate}`);
                    }}
                    title="Delete attachments"
                    className="bg-red-50 hover:bg-red-650 hover:text-white text-red-600 px-2.5 py-1.5 rounded-lg border border-red-200 font-bold text-[9px] uppercase tracking-wider shrink-0 transition-all"
                  >
                    Remove
                  </button>
                </div>
              </div>
            ) : (
              <div className="flex items-center gap-3 bg-slate-50 border-2 border-dashed border-slate-200 p-3 rounded-2xl">
                <FileText className="text-slate-300 shrink-0" size={24} />
                <div>
                  <div className="text-[10px] font-black text-slate-400 uppercase">No raw scan attachment</div>
                  <div className="text-[9px] font-bold text-slate-500 uppercase">Select a file below to catalog for this day</div>
                </div>
                <input
                  type="file"
                  accept=".pdf"
                  id="daily-pdf-upload"
                  className="hidden"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    if (file.type !== 'application/pdf') {
                      showToast("ONLY PDF FILES ARE PERMITTED");
                      return;
                    }
                    const rdr = new FileReader();
                    rdr.onload = () => {
                      const base64 = rdr.result as string;
                      setDailyPdfs(prev => ({
                        ...prev,
                        [reportDate]: { base64, name: file.name }
                      }));
                      showToast(`Cataloged raw PDF for date ${reportDate} successfully!`);
                    };
                    rdr.readAsDataURL(file);
                  }}
                />
                <button
                  onClick={() => document.getElementById('daily-pdf-upload')?.click()}
                  className="bg-white hover:bg-indigo-600 hover:text-white text-indigo-600 px-3.5 py-1.5 rounded-lg border border-indigo-200 font-black uppercase text-[9px] tracking-wider shrink-0 shadow-xs transition-all flex items-center gap-1.5"
                >
                  <Upload size={11} /> Upload PDF
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="bg-white p-10 border shadow-sm print:p-0 print:border-0 print:shadow-none min-h-[1000px] print-content">
         <div className="text-center space-y-2 mb-10 border-b-2 border-black pb-6">
            <h1 className="text-xl font-black uppercase text-gray-900">Second Division, Immigration Department (Myeik)</h1>
            <h2 className="text-lg font-bold text-gray-700 tracking-widest uppercase border border-black inline-block px-8 py-1">Daily Movement Log Report</h2>
            <div className="flex justify-center gap-8 text-sm font-bold pt-2">
              <span>DATE: {searchStr}</span>
              <span>TOTAL LOGS: {dailyIn.length + dailyOut.length}</span>
            </div>
         </div>

          <div className="space-y-12">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 no-print">
               <div className="space-y-6">
                 <h3 className="text-sm font-black uppercase text-emerald-800 border-b border-emerald-100 pb-2">Arrival (IN) Statistics</h3>
                 <div className="grid grid-cols-2 gap-4">
                   {generateAnalysisTable(dailyIn, "Nationality", "nationality")}
                   {generateAnalysisTable(dailyIn, "Visa Type", "visaType")}
                 </div>
               </div>
               <div className="space-y-6">
                 <h3 className="text-sm font-black uppercase text-orange-700 border-b border-orange-100 pb-2">Departure (OUT) Statistics</h3>
                 <div className="grid grid-cols-2 gap-4">
                   {generateAnalysisTable(dailyOut, "Nationality", "nationality")}
                   {generateAnalysisTable(dailyOut, "Visa Type", "visaType")}
                 </div>
               </div>
            </div>

            <section>
              <div className="flex justify-between items-end mb-3 border-b-2 border-gray-100 pb-1">
                 <h3 className="text-md font-black uppercase text-emerald-800 flex items-center gap-2">
                   <LogIn size={18} /> ARRIVAL LIST (အဝင်စာရင်း)
                 </h3>
                 <div className="flex gap-2 no-print">
                   <button onClick={() => copyToClipboard(dailyIn, "Arrival")} className="btn bg-gray-100 text-gray-700 hover:bg-gray-200 text-[10px] py-1 px-3 flex items-center gap-2"><Copy size={12} /> Copy</button>
                   <button onClick={() => exportDailyToExcel(dailyIn, "Arrival")} className="btn bg-emerald-50 text-emerald-700 border border-emerald-100 hover:bg-emerald-100 text-[10px] py-1 px-3 flex items-center gap-2"><Download size={12} /> Excel</button>
                 </div>
              </div>
              {generateTable(dailyIn)}
            </section>

            <section>
              <div className="flex justify-between items-end mb-3 border-b-2 border-gray-100 pb-1">
                 <h3 className="text-md font-black uppercase text-orange-700 flex items-center gap-2">
                   <LogOut size={18} /> DEPARTURE LIST (အထွက်စာရင်း)
                 </h3>
                 <div className="flex gap-2 no-print">
                   <button onClick={() => copyToClipboard(dailyOut, "Departure")} className="btn bg-gray-100 text-gray-700 hover:bg-gray-200 text-[10px] py-1 px-3 flex items-center gap-2"><Copy size={12} /> Copy</button>
                   <button onClick={() => exportDailyToExcel(dailyOut, "Departure")} className="btn bg-orange-50 text-orange-700 border border-orange-100 hover:bg-orange-100 text-[10px] py-1 px-3 flex items-center gap-2"><Download size={12} /> Excel</button>
                 </div>
              </div>
              {generateTable(dailyOut)}
            </section>
         </div>

         <div className="mt-20 grid grid-cols-2 gap-20 px-10">
            <div className="text-center space-y-12">
              <p className="text-xs font-bold uppercase underline">Prepared By</p>
              <div className="border-t border-black pt-2 text-xs font-bold">Log Unit Officer</div>
            </div>
            <div className="text-center space-y-12">
              <p className="text-xs font-bold uppercase underline">Verified By</p>
              <div className="border-t border-black pt-2 text-xs font-bold">Duty Commander</div>
            </div>
         </div>
      </div>
    </div>
  );
};

const StillInAnalytics = ({ 
  movementMap, 
  records, 
  setRecords, 
  setTempRecords, 
  showToast,
  setActivePrintPreview
}: { 
  movementMap: Record<string, MovementData>; 
  records: ImmRecord[]; 
  setRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  setTempRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  showToast: (msg: string) => void;
  setActivePrintPreview: React.Dispatch<React.SetStateAction<any>>;
}) => {
  const stillInList = (Object.values(movementMap) as MovementData[]).filter(m => !m.out);
  const [searchQuery, setSearchQuery] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [subTab, setSubTab] = useState<'analysis' | 'reports'>('analysis');

  // Convert date picker format (YYYY-MM-DD) to millisecond timestamps safely
  const startMs = startDate ? new Date(startDate + "T00:00:00").getTime() : 0;
  const endMs = endDate ? new Date(endDate + "T23:59:59").getTime() : Infinity;

  // Filter stillInList (Arrival Date check)
  const filteredStillInList = stillInList.filter(m => {
    const entryTime = m.inTime || parseTimestamp(m.in);
    return entryTime >= startMs && entryTime <= endMs;
  });

  // Filter records pool for general movement reports (IN/OUT) within date range
  const filteredRecords = records.filter(r => {
    const recTime = parseTimestamp(r.timestamp);
    return recTime >= startMs && recTime <= endMs;
  });

  // Calculate dynamic reports:
  // 1. Nationality of IN & OUT (Inbound and Outbound)
  const nationalityReportList = useMemo(() => {
    const report: Record<string, { nationality: string; inCount: number; outCount: number; total: number; net: number }> = {};
    filteredRecords.forEach(r => {
      const nat = (r.nationality || 'Unknown').trim().toUpperCase();
      if (!nat) return;
      if (!report[nat]) {
        report[nat] = { nationality: r.nationality || nat, inCount: 0, outCount: 0, total: 0, net: 0 };
      }
      if (r.mode === 'IN') {
        report[nat].inCount++;
        report[nat].net++;
      } else {
        report[nat].outCount++;
        report[nat].net--;
      }
      report[nat].total++;
    });
    return Object.values(report).sort((a, b) => b.total - a.total);
  }, [filteredRecords]);

  // 2. Visa Type Distribution (IN & OUT)
  const visaReportList = useMemo(() => {
    const report: Record<string, { visaType: string; inCount: number; outCount: number; total: number }> = {};
    filteredRecords.forEach(r => {
      const visa = (r.visaType || 'Unknown').trim();
      if (!visa) return;
      const key = visa.toUpperCase();
      if (!report[key]) {
        report[key] = { visaType: visa, inCount: 0, outCount: 0, total: 0 };
      }
      if (r.mode === 'IN') {
        report[key].inCount++;
      } else {
        report[key].outCount++;
      }
      report[key].total++;
    });
    return Object.values(report).sort((a, b) => b.total - a.total);
  }, [filteredRecords]);

  const updateStatusInRecords = (id: number, status: PermitStatus | '', permittedBy: string) => {
    const recordToUpdate = records.find(r => r.id === id);
    if (recordToUpdate) {
      const passport = recordToUpdate.passport.toUpperCase();
      setRecords((prev: ImmRecord[]) => prev.map(r => r.passport.toUpperCase() === passport ? { ...r, stillPermittedStatus: status, permittedBy, syncStatus: 'pending_sync', updatedAt: new Date().toISOString() } : r));
      setTempRecords((prev: ImmRecord[]) => prev.map(r => r.passport.toUpperCase() === passport ? { ...r, stillPermittedStatus: status, permittedBy, syncStatus: 'pending_sync', updatedAt: new Date().toISOString() } : r));
    } else {
      setRecords((prev: ImmRecord[]) => prev.map(r => r.id === id ? { ...r, stillPermittedStatus: status, permittedBy, syncStatus: 'pending_sync', updatedAt: new Date().toISOString() } : r));
      setTempRecords((prev: ImmRecord[]) => prev.map(r => r.id === id ? { ...r, stillPermittedStatus: status, permittedBy, syncStatus: 'pending_sync', updatedAt: new Date().toISOString() } : r));
    }
    showToast("Permit Status updated!");
  };

  // EXPORT UTILITIES:
  // 1. Excel Export (Detailed Current Still-In list matching selected filters)
  const exportStillInExcel = () => {
    const data = filteredStillInList.map((m, i) => {
      const personRecords = records.filter(r => r.passport.toUpperCase() === m.p.toUpperCase())
        .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
      
      const latest = personRecords[0];
      const previousStays = personRecords
        .slice(1)
        .map(r => `${r.timestamp.split(',')[0]}: ${r.address}`)
        .join(" | ");

      return {
        "No": i + 1,
        "Passport Number": m.p,
        "Full Name": m.n,
        "Gender": m.gender,
        "Nationality": m.nat,
        "Visa Type": m.visa,
        "Current Stay Location": latest?.address || m.loc,
        "Detail Stay Description": latest?.stayDescription || '',
        "Stay Period From": m.start,
        "Stay Period To": m.end,
        "Allowed Days": m.allowed,
        "Remaining Days": getLiveRemainingDays(m.end),
        "Arrived From (Origin)": latest?.arrivedFrom || '',
        "Departed To (Destination)": latest?.departedTo || '',
        "Agent Name (Brought By)": m.agent,
        "Agent Contact": latest?.contactDetails || '',
        "Vehicle Info": m.vInfo,
        "Arrival Log Date/Time": latest?.timestamp || m.in,
        "Stay Permit Status": latest?.stillPermittedStatus || 'N/A',
        "Stay Permit Description": latest?.permittedBy || 'N/A',
        "Verified Officer Name": latest?.officialName || '',
        "Verified Officer Title": latest?.officialTitle || '',
        "Log Type": latest?.logType || 'FFE',
        "FormC Official": latest?.formC?.officialName || '',
        "FormC Official Title": latest?.formC?.officialTitle || '',
        "FormC Reporter": latest?.formC?.reporterName || '',
        "FormC Reporter Phone": latest?.formC?.reporterPhone || '',
        "Remarks / Dossier Remarks": latest?.remarks || '',
        "Previous Stays History": previousStays || 'No previous records'
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Still In List Detailed");
    XLSX.writeFile(wb, `Detailed_StillIn_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Excel Live List Exported!");
  };

  // 1b. Individual Nationality Excel Export
  const exportNationalityExcel = () => {
    const natData = nationalityReportList.map((r, i) => ({
      "No": i + 1,
      "Nationality": r.nationality,
      "Inbound Entries (IN)": r.inCount,
      "Outbound Exits (OUT)": r.outCount,
      "Net Stayers (Remaining)": r.net,
      "Total Movements": r.total
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(natData);
    XLSX.utils.book_append_sheet(wb, ws, "Nationality Stats");
    XLSX.writeFile(wb, `SIA_Nationality_Report_${startDate || 'All'}_to_${endDate || 'All'}.xlsx`);
    showToast("Nationality Excel exported successfully!");
  };

  // 1c. Individual Nationality Print Trigger
  const printNationalityReport = () => {
    setActivePrintPreview({
      title: `Nationality Inflow & Outflow Report (${startDate || 'All Time'} to ${endDate || 'All Time'})`,
      type: 'SIA_NATIONALITY_REPORT',
      data: nationalityReportList
    });
  };

  // 1d. Individual Visa Type Excel Export
  const exportVisaExcel = () => {
    const visaData = visaReportList.map((r, i) => ({
      "No": i + 1,
      "Visa Type": r.visaType,
      "Inbound Entries (IN)": r.inCount,
      "Outbound Exits (OUT)": r.outCount,
      "Total Movements": r.total
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(visaData);
    XLSX.utils.book_append_sheet(wb, ws, "Visa Distribution");
    XLSX.writeFile(wb, `SIA_Visa_Distribution_Report_${startDate || 'All'}_to_${endDate || 'All'}.xlsx`);
    showToast("Visa Distribution Excel exported successfully!");
  };

  // 1e. Individual Visa Type Print Trigger
  const printVisaReport = () => {
    setActivePrintPreview({
      title: `Visa Type Distribution Report (${startDate || 'All Time'} to ${endDate || 'All Time'})`,
      type: 'SIA_VISA_REPORT',
      data: visaReportList
    });
  };

  // 1f. Full SIA Print Trigger
  const printFullSiaReport = () => {
    setActivePrintPreview({
      title: `STATISTICAL INVESTIGATION ANALYSIS (SIA) REPORT`,
      type: 'SIA_FULL_REPORT',
      data: {
        nationalityList: nationalityReportList,
        visaList: visaReportList,
        startDate: startDate || 'All Time',
        endDate: endDate || 'All Time'
      }
    });
  };

  // 2. Excel Reports Export (Nationality and Visa reports sheets)
  const exportReportsToExcel = () => {
    const wb = XLSX.utils.book_new();

    // Nationality report data
    const natData = nationalityReportList.map((r, i) => ({
      "No": i + 1,
      "Nationality": r.nationality,
      "Inbound Entries (IN)": r.inCount,
      "Outbound Exits (OUT)": r.outCount,
      "Net Stayers (Remaining)": r.net,
      "Total Movements": r.total
    }));
    const wsNat = XLSX.utils.json_to_sheet(natData);
    XLSX.utils.book_append_sheet(wb, wsNat, "Nationality In-Out Stats");

    // Visa type report data
    const visaData = visaReportList.map((r, i) => ({
      "No": i + 1,
      "Visa Type": r.visaType,
      "Inbound Entries (IN)": r.inCount,
      "Outbound Exits (OUT)": r.outCount,
      "Total Movements": r.total
    }));
    const wsVisa = XLSX.utils.json_to_sheet(visaData);
    XLSX.utils.book_append_sheet(wb, wsVisa, "Visa Type Distribution");

    const dateStr = `SIA_Reports_${startDate || 'All'}_to_${endDate || 'All'}.xlsx`;
    XLSX.writeFile(wb, dateStr);
    showToast("Excel Reports exported successfully!");
  };

  // 3. Word Document Export (SIA statistics with related official headers)
  const exportReportsToWord = () => {
    const headerHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <title>Immigration Statistical Investigation Analysis (SIA) Report</title>
        <style>
          body { font-family: 'Arial', sans-serif; color: #333333; line-height: 1.5; padding: 20px; }
          h1 { color: #1e3a8a; font-size: 18pt; text-align: center; margin-bottom: 5px; font-weight: bold; }
          h2 { color: #0f172a; font-size: 13pt; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; margin-top: 30px; font-weight: bold; }
          .meta-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
          .meta-table td { padding: 4px; font-size: 10pt; }
          table.report-table { width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 25px; }
          table.report-table th { background-color: #1e3a8a; color: #ffffff; border: 1px solid #1e3a8a; padding: 8px; font-size: 10pt; font-weight: bold; text-align: left; }
          table.report-table td { border: 1px solid #dddddd; padding: 8px; font-size: 10pt; }
          table.report-table tr:nth-child(even) { background-color: #f8fafc; }
          .footer { text-align: center; margin-top: 50px; font-size: 9pt; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px; }
          .text-center { text-align: center; }
          .font-bold { font-weight: bold; }
        </style>
      </head>
      <body>
        <h1>SECOND DIVISION • IMMIGRATION DEPARTMENT</h1>
        <p class="text-center" style="font-size: 11pt; font-weight: bold; color: #475569; margin-top: 0;">Statistical Investigation Analysis (SIA) Official Report</p>
        
        <table class="meta-table">
          <tr>
            <td style="width: 50%;"><strong>Report Generated:</strong> ${new Date().toLocaleString()}</td>
            <td style="width: 50%; text-align: right;"><strong>Filtered Date Range:</strong> ${startDate || 'All Time'} to ${endDate || 'All Time'}</td>
          </tr>
        </table>

        <h2>1. Nationality Inflow & Outflow Statistical Breakdown</h2>
        <table class="report-table">
          <thead>
            <tr>
              <th style="width: 8%;">No.</th>
              <th>Nationality</th>
              <th style="text-align: center;">Inbound (IN)</th>
              <th style="text-align: center;">Outbound (OUT)</th>
              <th style="text-align: center;">Net Staying</th>
              <th style="text-align: center;">Total Movements</th>
            </tr>
          </thead>
          <tbody>
            ${nationalityReportList.length === 0 ? '<tr><td colspan="6" class="text-center">No matching records logged.</td></tr>' : nationalityReportList.map((r, i) => `
              <tr>
                <td>${i + 1}</td>
                <td class="font-bold">${r.nationality}</td>
                <td class="text-center">${r.inCount}</td>
                <td class="text-center">${r.outCount}</td>
                <td class="text-center font-bold" style="color: ${r.net >= 0 ? '#15803d' : '#b91c1c'};">${r.net}</td>
                <td class="text-center font-bold">${r.total}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <h2>2. Visa Type Distribution statistical Breakdown</h2>
        <table class="report-table">
          <thead>
            <tr>
              <th style="width: 8%;">No.</th>
              <th>Visa Type</th>
              <th style="text-align: center;">Inbound (IN)</th>
              <th style="text-align: center;">Outbound (OUT)</th>
              <th style="text-align: center;">Total Movements</th>
            </tr>
          </thead>
          <tbody>
            ${visaReportList.length === 0 ? '<tr><td colspan="5" class="text-center">No matching records logged.</td></tr>' : visaReportList.map((r, i) => `
              <tr>
                <td>${i + 1}</td>
                <td class="font-bold">${r.visaType}</td>
                <td class="text-center">${r.inCount}</td>
                <td class="text-center">${r.outCount}</td>
                <td class="text-center font-bold">${r.total}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div class="footer">
          <p>CONFIDENTIAL IMMIGRATION DOCUMENTS • INTERNAL INVESTIGATION USE ONLY</p>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + headerHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `SIA_Immigration_Reports_${startDate || 'All'}_to_${endDate || 'All'}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("DOC Report exported successfully!");
  };

  // 4. HTML Document Export (SIA reports with styled clean layouts)
  const exportReportsToHtml = () => {
    const htmlContent = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Immigration SIA Report</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: #f8fafc; color: #1e293b; padding: 40px 20px; line-height: 1.6; }
          .container { max-width: 950px; margin: 0 auto; background: #ffffff; padding: 40px; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #e2e8f0; }
          .header { text-align: center; border-bottom: 4px solid #1e40af; padding-bottom: 20px; margin-bottom: 30px; }
          .header h1 { font-size: 22px; font-weight: 800; color: #1e3a8a; margin: 0; text-transform: uppercase; letter-spacing: 0.05em; }
          .header h2 { font-size: 15px; font-weight: 600; color: #475569; margin: 6px 0 0 0; }
          .meta-grid { display: flex; justify-content: space-between; background: #f1f5f9; padding: 12px 20px; border-radius: 8px; margin-bottom: 30px; font-size: 13px; font-weight: 500; }
          .section-title { font-size: 16px; font-weight: 800; color: #0f172a; margin-top: 30px; margin-bottom: 14px; text-transform: uppercase; letter-spacing: 0.025em; border-left: 4px solid #1e40af; padding-left: 10px; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
          th { background-color: #1e40af; color: white; font-weight: 700; text-align: left; padding: 10px 14px; font-size: 12px; text-transform: uppercase; }
          td { padding: 10px 14px; font-size: 13.5px; border-bottom: 1px solid #e2e8f0; }
          tr:nth-child(even) td { background-color: #f8fafc; }
          .text-center { text-align: center; }
          .font-bold { font-weight: 700; }
          .footer { text-align: center; font-size: 12px; color: #94a3b8; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 20px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>Second Division • Immigration Department</h1>
            <h2>Statistical Investigation Analysis (SIA) Official Report</h2>
          </div>
          
          <div class="meta-grid">
            <div><strong>Generated Date:</strong> ${new Date().toLocaleString()}</div>
            <div><strong>Date Range Filter:</strong> ${startDate || 'All Time'} to ${endDate || 'All Time'}</div>
          </div>

          <div class="section-title">1. Nationality Inbound & Outbound statistical breakdown</div>
          <table>
            <thead>
              <tr>
                <th style="width: 8%;">No.</th>
                <th>Nationality</th>
                <th style="text-align: center;">Inbound (IN)</th>
                <th style="text-align: center;">Outbound (OUT)</th>
                <th style="text-align: center;">Net Staying</th>
                <th style="text-align: center;">Total Movements</th>
              </tr>
            </thead>
            <tbody>
              ${nationalityReportList.length === 0 ? '<tr><td colspan="6" class="text-center">No matching records logged.</td></tr>' : nationalityReportList.map((r, i) => `
                <tr>
                  <td>${i + 1}</td>
                  <td class="font-bold">${r.nationality}</td>
                  <td class="text-center">${r.inCount}</td>
                  <td class="text-center">${r.outCount}</td>
                  <td class="text-center font-bold" style="color: ${r.net >= 0 ? '#166534' : '#991b1b'};">${r.net}</td>
                  <td class="text-center font-bold">${r.total}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="section-title">2. Visa Type Distribution statistical breakdown</div>
          <table>
            <thead>
              <tr>
                <th style="width: 8%;">No.</th>
                <th>Visa Type</th>
                <th style="text-align: center;">Inbound (IN)</th>
                <th style="text-align: center;">Outbound (OUT)</th>
                <th style="text-align: center;">Total Movements</th>
              </tr>
            </thead>
            <tbody>
              ${visaReportList.length === 0 ? '<tr><td colspan="5" class="text-center">No matching records logged.</td></tr>' : visaReportList.map((r, i) => `
                <tr>
                  <td>${i + 1}</td>
                  <td class="font-bold">${r.visaType}</td>
                  <td class="text-center">${r.inCount}</td>
                  <td class="text-center">${r.outCount}</td>
                  <td class="text-center font-bold">${r.total}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="footer">
            <p>CONFIDENTIAL IMMIGRATION SYSTEMS • FOR AUTHENTICATED INTERNAL USE ONLY</p>
          </div>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `SIA_Immigration_Reports_${startDate || 'All'}_to_${endDate || 'All'}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("HTML Report exported successfully!");
  };

  const generateStillInAnalysis = (title: string, type: keyof MovementData) => {
    const analysis: Record<string, { label: string; M: number; F: number; T: number }> = {};
    filteredStillInList.forEach(m => {
      const val = (m[type] as string) || 'Unknown';
      if (!analysis[val]) analysis[val] = { label: val, M: 0, F: 0, T: 0 };
      if (m.gender === 'M') analysis[val].M++;
      else if (m.gender === 'F') analysis[val].F++;
      analysis[val].T++;
    });

    const sorted = Object.values(analysis).sort((a, b) => b.T - a.T);
    const total = sorted.reduce((acc, curr) => ({ M: acc.M + curr.M, F: acc.F + curr.F, T: acc.T + curr.T }), { M: 0, F: 0, T: 0 });

    if (sorted.length === 0) {
      return (
        <div className="card bg-white border border-slate-100 p-6 rounded-2xl shadow-sm text-center">
          <p className="text-xs font-black uppercase text-gray-400 mb-2 tracking-widest border-b pb-2">{title}</p>
          <p className="text-[11px] font-bold text-slate-400 italic py-4">No records in selected range.</p>
        </div>
      );
    }

    return (
      <div className="card">
        <h4 className="text-xs font-black uppercase text-gray-500 mb-4 tracking-widest border-b pb-2">{title}</h4>
        <div className="overflow-x-auto border border-black rounded bg-white">
          <table className="w-full text-xs text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-black">
                <th className="p-2">Category</th>
                <th className="w-12 p-2 text-center">M</th>
                <th className="w-12 p-2 text-center">F</th>
                <th className="w-12 p-2 text-center">Total</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {sorted.map((s, idx) => (
                <tr key={idx} className="hover:bg-gray-50">
                   <td className="p-2 font-bold truncate max-w-[150px]">{s.label}</td>
                   <td className="p-2 text-center">{s.M}</td>
                   <td className="p-2 text-center">{s.F}</td>
                   <td className="p-2 text-center font-black">{s.T}</td>
                </tr>
              ))}
            </tbody>
            <tfoot className="bg-gray-900 text-white font-black">
               <tr>
                 <td className="p-2">TOTAL</td>
                 <td className="p-2 text-center">{total.M}</td>
                 <td className="p-2 text-center">{total.F}</td>
                 <td className="p-2 text-center">{total.T}</td>
               </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  };

  return (
    <div className="max-w-[1600px] mx-auto py-4 sm:py-8 px-2 sm:px-4 space-y-4 sm:space-y-8">
      {/* Upper header summary panel */}
      <div className="card shadow-xl border-t-8 border-indigo-700 no-print p-4 sm:p-6">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 sm:gap-6">
          <div className="flex items-center gap-3 sm:gap-4">
             <div className="bg-indigo-100 p-2.5 sm:p-3 rounded-full text-indigo-700 shrink-0">
                <Activity size={24} className="sm:w-7 sm:h-7" />
             </div>
             <div>
                <h2 className="text-lg sm:text-2xl font-black text-gray-800 uppercase tracking-tight">STATISTICAL INVESTIGATION ANALYSIS (SIA) - MYEIK</h2>
                <p className="text-gray-400 text-[10px] sm:text-xs font-bold font-mono uppercase">Immigration Intelligence & Traffic Flow (Myeik)</p>
             </div>
          </div>
          
          <div className="flex items-center justify-between sm:justify-center gap-3 sm:gap-6 bg-gray-50 px-4 py-2.5 sm:px-6 sm:py-3 rounded-2xl border border-gray-100 w-full md:w-auto">
             <div className="text-center">
                <p className="text-[9px] sm:text-[10px] font-black text-gray-400 uppercase">MALE STILL IN</p>
                <p className="text-lg sm:text-xl font-black text-blue-600">{filteredStillInList.filter(m => m.gender === 'M').length}</p>
             </div>
             <div className="h-8 w-px bg-gray-200" />
             <div className="text-center">
                <p className="text-[9px] sm:text-[10px] font-black text-gray-400 uppercase">FEMALE STILL IN</p>
                <p className="text-lg sm:text-xl font-black text-pink-600">{filteredStillInList.filter(m => m.gender === 'F').length}</p>
             </div>
             <div className="h-8 w-px bg-gray-200" />
             <div className="text-center">
                <p className="text-[9px] sm:text-[10px] font-black text-gray-400 uppercase">TOTAL STILL IN</p>
                <p className="text-xl sm:text-2xl font-black text-indigo-700">{filteredStillInList.length}</p>
             </div>
          </div>
        </div>
      </div>

      {/* Dynamic Date Range Filter and Section Tabs */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 sm:gap-4 p-3.5 sm:p-4 rounded-2xl bg-slate-50 border border-slate-200/80 shadow-inner no-print">
        <div className="flex flex-wrap items-center gap-2.5 sm:gap-4">
          <div className="flex items-center gap-2">
            <Calendar size={16} className="text-indigo-600 shrink-0" />
            <span className="text-[11px] sm:text-xs font-black uppercase text-slate-500 tracking-wider">Date Range Filter:</span>
          </div>
          <div className="flex items-center gap-1.5 sm:gap-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase">From:</span>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="bg-white px-2 sm:px-2.5 py-1 sm:py-1.5 border border-slate-300 rounded-xl text-[11px] sm:text-xs font-black text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-600"
            />
          </div>
          <div className="flex items-center gap-1.5 sm:gap-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase">To:</span>
            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="bg-white px-2 sm:px-2.5 py-1 sm:py-1.5 border border-slate-300 rounded-xl text-[11px] sm:text-xs font-black text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-600"
            />
          </div>
          {(startDate || endDate) && (
            <button
              onClick={() => { setStartDate(''); setEndDate(''); }}
              className="text-[10px] bg-slate-200 text-slate-700 hover:bg-slate-300 font-bold uppercase px-2.5 py-1 sm:px-3 sm:py-1.5 rounded-lg transition-all"
            >
              Clear Filter
            </button>
          )}
        </div>

        {/* Section selectors */}
        <div className="grid grid-cols-2 w-full sm:w-auto bg-slate-200 p-1 rounded-xl border border-slate-300/60 self-start lg:self-center">
          <button
            onClick={() => setSubTab('analysis')}
            className={`px-3 sm:px-4 py-1.5 rounded-lg text-[11px] sm:text-xs font-black uppercase tracking-tight transition-all text-center ${
              subTab === 'analysis' 
                ? 'bg-indigo-600 text-white shadow-md' 
                : 'text-slate-600 hover:text-slate-800'
            }`}
          >
            Still In Analysis
          </button>
          <button
            onClick={() => setSubTab('reports')}
            className={`px-3 sm:px-4 py-1.5 rounded-lg text-[11px] sm:text-xs font-black uppercase tracking-tight transition-all text-center ${
              subTab === 'reports' 
                ? 'bg-indigo-600 text-white shadow-md' 
                : 'text-slate-600 hover:text-slate-800'
            }`}
          >
            Reports Dashboard
          </button>
        </div>
      </div>

      {subTab === 'analysis' ? (
        <div className="space-y-6 sm:space-y-8 animate-fadeIn">
          {/* Section 1: Still In Analysis visual boxes */}
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 mb-1 no-print">
            <h3 className="text-xs sm:text-sm font-black text-slate-500 uppercase tracking-wider">Live Staying Visitors Categorization</h3>
            <button 
              onClick={() => {
                setActivePrintPreview({
                  title: `Active Stay Inbound Visitors Analytics (${startDate || 'All Time'} to ${endDate || 'All Time'})`,
                  type: 'STILL_IN_ANALYTICS',
                  data: filteredStillInList
                });
              }} 
              className="btn bg-gray-900 text-white hover:bg-black uppercase text-[11px] sm:text-xs px-3 sm:px-4 py-2 flex items-center justify-center gap-2 shadow-sm w-full sm:w-auto"
            >
              <Printer size={14} /> Print Category Summary
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-8 no-print">
            {generateStillInAnalysis("Analysis by Nationality", "nat")}
            {generateStillInAnalysis("Analysis by Visa Type", "visa")}
            {generateStillInAnalysis("Analysis by Current Location", "loc")}
          </div>

          {/* Detailed Individual Live Registry */}
          <motion.div initial={{ y: 20, opacity: 0 }} animate={{ y: 0, opacity: 1 }} className="card shadow-md border border-gray-100 no-print p-4 sm:p-6">
             <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-3 sm:gap-4 mb-4 sm:mb-6">
               <div>
                 <h3 className="text-base sm:text-lg font-black text-slate-900 uppercase">Still In Record Permission Statuses (ခွင့်ပြုချက် စာရင်း)</h3>
                 <p className="text-[11px] sm:text-xs text-slate-500 font-bold">Inspect, verify, and switch (STILL PERMITTED) / (STILL NOT PERMITTED YET) statuses</p>
               </div>
               
               <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2 sm:gap-2.5 w-full lg:w-auto">
                 <button 
                   onClick={exportStillInExcel} 
                   className="btn bg-emerald-700 text-white hover:bg-emerald-800 uppercase text-[11px] sm:text-xs px-3 sm:px-4 py-2 flex items-center justify-center gap-1.5 shadow"
                   title="Download Still In List as Excel with all form entry info"
                 >
                   <Download size={14} /> Download Still In (Excel)
                 </button>
                 <button 
                   onClick={() => {
                     setActivePrintPreview({
                       title: `Active Stay Inbound Visitors List (${startDate || 'All Time'} to ${endDate || 'All Time'})`,
                       type: 'STILL_IN_ANALYTICS',
                       data: filteredStillInList
                     });
                   }} 
                   className="btn bg-gray-900 text-white hover:bg-black uppercase text-[11px] sm:text-xs px-3 sm:px-4 py-2 flex items-center justify-center gap-1.5 shadow"
                   title="Print Still In List"
                 >
                   <Printer size={14} /> Print Still In List
                 </button>
                 
                 <div className="relative w-full sm:w-64">
                   <input
                     type="text"
                     value={searchQuery}
                     onChange={(e) => setSearchQuery(e.target.value)}
                     placeholder="Search Passport or Name..."
                     className="input-field pl-8 text-xs bg-slate-50 w-full"
                   />
                   <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
                 </div>
               </div>
             </div>

             {/* Mobile Cards View for Small Screens */}
             <div className="block lg:hidden space-y-3">
               {stillInList
                 .filter(m => m.p.toLowerCase().includes(searchQuery.toLowerCase()) || m.n.toLowerCase().includes(searchQuery.toLowerCase()))
                 .map((m, idx) => {
                   const personRecords = records.filter(r => r.passport.toUpperCase() === m.p.toUpperCase())
                     .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
                   const latest = personRecords[0];

                   const rawStatus = latest?.stillPermittedStatus || '';
                   const currentStatus = rawStatus === 'STILL PERMITTED' ? 'STAY PERMITTED' : (rawStatus === 'STILL NOT PERMITTED YET' ? 'STAY NOT PERMITTED' : rawStatus);
                   const currentPermittedBy = latest?.permittedBy || '';

                   return (
                     <div key={idx} className="p-3.5 bg-slate-50/80 border border-slate-200/80 rounded-2xl space-y-2.5 text-xs">
                       <div className="flex justify-between items-start gap-2">
                         <div>
                           <div className="font-black text-indigo-950 font-mono text-sm tracking-tight">{m.p}</div>
                           <div className="font-bold text-slate-800 uppercase text-xs">{m.n}</div>
                         </div>
                         <span className="text-[10px] font-black uppercase px-2 py-0.5 rounded bg-indigo-100 text-indigo-800 shrink-0">
                           {m.nat}
                         </span>
                       </div>

                       <div className="text-[11px] space-y-1 text-slate-600">
                         <div><span className="font-bold text-slate-400 uppercase text-[9px]">Location:</span> <span className="font-semibold text-slate-800">{latest?.address || m.loc}</span></div>
                         <div><span className="font-bold text-slate-400 uppercase text-[9px]">Stay Period:</span> <span className="font-bold text-indigo-700">{m.start} to {m.end}</span> <span className="font-mono text-slate-400">({m.allowed} Days / {getLiveRemainingDays(m.end)})</span></div>
                         <div><span className="font-bold text-slate-400 uppercase text-[9px]">Last Log:</span> <span className="font-mono font-bold text-slate-700">{latest?.timestamp || m.in}</span></div>
                       </div>

                       <div className="p-2 bg-white border border-slate-200 rounded-xl space-y-2">
                         <label className="text-[9px] font-black uppercase text-slate-400 block">Permit Status:</label>
                         <select
                           value={currentStatus}
                           onChange={(e) => {
                             const newStatus = e.target.value as any;
                             if (latest) {
                               updateStatusInRecords(latest.id, newStatus, currentPermittedBy);
                             }
                           }}
                           className={`w-full text-xs bg-white border border-slate-200 rounded-lg p-2 font-black uppercase focus:outline-none ${
                             currentStatus === "STAY PERMITTED" 
                               ? "text-emerald-700 border-emerald-300 bg-emerald-50/20" 
                               : currentStatus === "STAY NOT PERMITTED"
                               ? "text-red-700 border-red-300 bg-red-50/20"
                               : "text-slate-700"
                           }`}
                         >
                           <option value="">-- Choose Status --</option>
                           <option value="STAY PERMITTED">STAY PERMITTED (ခွင့်ပြုထားဆဲ)</option>
                           <option value="STAY NOT PERMITTED">STAY NOT PERMITTED (ခွင့်မပြုသေးပါ)</option>
                         </select>

                         {currentStatus === "နေထိုင်ခွင့် မလျှောက်ထားသေးသူ" && latest && (
                           <input
                             type="text"
                             value={currentPermittedBy}
                             onChange={(e) => {
                               updateStatusInRecords(latest.id, currentStatus, e.target.value);
                             }}
                             className="text-xs w-full border border-red-200 focus:border-red-500 rounded-lg p-2 font-bold text-red-800 bg-white"
                             placeholder="Not permitted reason..."
                           />
                         )}
                       </div>
                     </div>
                   );
                 })}
             </div>

             {/* Desktop / Tablet Table View */}
             <div className="hidden lg:block overflow-x-auto border border-slate-100 rounded-xl max-h-[600px] overflow-y-auto">
               <table className="w-full text-left text-xs">
                 <thead className="bg-slate-50 sticky top-0 border-b border-slate-100">
                   <tr className="uppercase text-[9px] text-slate-500 font-black">
                     <th className="p-3">Passport & Name</th>
                     <th className="p-3">Nationality</th>
                     <th className="p-3">Stay Location</th>
                     <th className="p-3">Stay Period</th>
                     <th className="p-3">Last Log Date</th>
                     <th className="p-3">Permit Status & Permitted By</th>
                   </tr>
                 </thead>
                 <tbody className="divide-y divide-slate-100 bg-white">
                   {stillInList
                     .filter(m => m.p.toLowerCase().includes(searchQuery.toLowerCase()) || m.n.toLowerCase().includes(searchQuery.toLowerCase()))
                     .map((m, idx) => {
                       const personRecords = records.filter(r => r.passport.toUpperCase() === m.p.toUpperCase())
                         .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
                       const latest = personRecords[0];

                       const rawStatus = latest?.stillPermittedStatus || '';
                       const currentStatus = rawStatus === 'STILL PERMITTED' ? 'STAY PERMITTED' : (rawStatus === 'STILL NOT PERMITTED YET' ? 'STAY NOT PERMITTED' : rawStatus);
                       const currentPermittedBy = latest?.permittedBy || '';

                       return (
                         <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                           <td className="p-3">
                             <div className="font-black text-slate-900 font-mono">{m.p}</div>
                             <div className="text-[10px] font-bold text-slate-500 uppercase">{m.n}</div>
                           </td>
                           <td className="p-3 font-semibold text-slate-700">{m.nat}</td>
                           <td className="p-3 font-semibold text-slate-800">
                             {latest?.address || m.loc}
                           </td>
                           <td className="p-3">
                             <div className="font-bold text-indigo-700">{m.start} to {m.end}</div>
                             <div className="text-[10px] text-slate-400 font-mono font-bold">Allowed: {m.allowed} Days ({getLiveRemainingDays(m.end)})</div>
                           </td>
                           <td className="p-3 font-mono text-[10px] font-bold text-slate-650">
                             {latest?.timestamp || m.in}
                           </td>
                           <td className="p-3">
                             <div className="flex flex-col gap-1.5 p-1.5 bg-slate-50 border border-slate-200/60 rounded-xl w-fit">
                               <select
                                 value={currentStatus}
                                 onChange={(e) => {
                                   const newStatus = e.target.value as any;
                                   if (latest) {
                                     updateStatusInRecords(latest.id, newStatus, currentPermittedBy);
                                   }
                                 }}
                                 className={`text-[10px] bg-white border border-slate-200 rounded-lg p-1.5 font-black uppercase focus:outline-none ${
                                    currentStatus === "STAY PERMITTED" 
                                      ? "text-emerald-700 border-emerald-300 bg-emerald-50/20" 
                                      : currentStatus === "STAY NOT PERMITTED"
                                      ? "text-red-700 border-red-300 bg-red-50/20"
                                      : "text-slate-700"
                                  }`}
                                >
                                 <option value="">-- Choose Status --</option>
                                 <option value="STAY PERMITTED">STAY PERMITTED (ခွင့်ပြုထားဆဲ)</option>
                                 <option value="STAY NOT PERMITTED">STAY NOT PERMITTED (ခွင့်မပြုသေးပါ)</option>
                               </select>

                               {currentStatus === "နေထိုင်ခွင့် မလျှောက်ထားသေးသူ" && latest && (
                                 <input
                                   type="text"
                                   value={currentPermittedBy}
                                   onChange={(e) => {
                                     updateStatusInRecords(latest.id, currentStatus, e.target.value);
                                   }}
                                   className="text-[10px] w-full border border-red-200 focus:border-red-500 rounded-lg p-1.5 font-bold text-red-800 bg-white mt-1"
                                   placeholder="Not permitted reason..."
                                 />
                               )}
                             </div>
                           </td>
                         </tr>
                       );
                     })}
                 </tbody>
               </table>
             </div>
          </motion.div>
        </div>
      ) : (
        <div className="space-y-6 sm:space-y-8 animate-fadeIn">
          {/* Section 2: Reports generation screen */}
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center p-4 sm:p-6 rounded-2xl bg-indigo-50 border border-indigo-100 shadow-sm gap-4 no-print">
            <div>
              <h3 className="text-sm sm:text-md font-black text-indigo-950 uppercase tracking-tight flex items-center gap-2">
                <FileText className="text-indigo-600 shrink-0" size={18} /> Official Reports generator
              </h3>
              <p className="text-[11px] sm:text-xs text-indigo-700 font-bold mt-1">
                Download fully compiled, formatted reports of Nationality inflows/outflows and Visa distribution.
              </p>
            </div>
            
            <div className="grid grid-cols-2 sm:flex sm:flex-wrap items-center gap-2 w-full sm:w-auto">
              <button
                onClick={printFullSiaReport}
                className="btn bg-gray-900 text-white hover:bg-black uppercase text-[10px] sm:text-xs px-3 sm:px-4 py-2 sm:py-2.5 flex items-center justify-center gap-1.5 sm:gap-2 shadow"
                disabled={filteredRecords.length === 0}
                title="Print Full SIA Report"
              >
                <Printer size={14} /> Print
              </button>
              <button
                onClick={exportReportsToExcel}
                className="btn bg-emerald-700 text-white hover:bg-emerald-800 uppercase text-[10px] sm:text-xs px-3 sm:px-4 py-2 sm:py-2.5 flex items-center justify-center gap-1.5 sm:gap-2 shadow"
                disabled={filteredRecords.length === 0}
                title="Export custom formatted Excel sheets"
              >
                <Download size={14} /> Excel
              </button>
              <button
                onClick={exportReportsToWord}
                className="btn bg-blue-700 text-white hover:bg-blue-800 uppercase text-[10px] sm:text-xs px-3 sm:px-4 py-2 sm:py-2.5 flex items-center justify-center gap-1.5 sm:gap-2 shadow"
                disabled={filteredRecords.length === 0}
                title="Export styled Word Document"
              >
                <FileText size={14} /> Word (DOC)
              </button>
              <button
                onClick={exportReportsToHtml}
                className="btn bg-slate-850 text-white hover:bg-slate-950 uppercase text-[10px] sm:text-xs px-3 sm:px-4 py-2 sm:py-2.5 flex items-center justify-center gap-1.5 sm:gap-2 shadow"
                disabled={filteredRecords.length === 0}
                title="Export clean HTML page"
              >
                <Table size={14} /> Web (HTML)
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 sm:gap-8">
            {/* 1. Nationality of In and Out Report */}
            <div className="card shadow-md border border-slate-200 bg-white p-4 sm:p-6">
              <div className="border-b pb-3 sm:pb-4 mb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                <div>
                  <h4 className="text-xs sm:text-sm font-black text-slate-800 uppercase tracking-tight">1. Nationality Inflow & Outflow</h4>
                  <p className="text-[10px] text-slate-500 font-bold mt-0.5">Calculated entry/exit trends per country</p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportNationalityExcel}
                    className="btn bg-emerald-700 text-white hover:bg-emerald-800 uppercase text-[9px] px-2.5 py-1.5 flex items-center gap-1 shadow"
                    title="Export Nationality Report to Excel"
                  >
                    <Download size={10} /> Excel
                  </button>
                  <button
                    onClick={printNationalityReport}
                    className="btn bg-gray-900 text-white hover:bg-black uppercase text-[9px] px-2.5 py-1.5 flex items-center gap-1 shadow"
                    title="Print Nationality Report"
                  >
                    <Printer size={10} /> Print
                  </button>
                  <span className="text-[10px] bg-slate-100 border border-slate-200 text-slate-700 font-black px-2 py-1 sm:px-2.5 sm:py-1.5 rounded-full shrink-0">
                    {nationalityReportList.length} Countries
                  </span>
                </div>
              </div>

              {nationalityReportList.length === 0 ? (
                <div className="text-center py-12 text-slate-400 font-bold text-xs italic">
                  No records matching selected date range filter.
                </div>
              ) : (
                <div className="overflow-x-auto border border-slate-100 rounded-xl">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-50 border-b border-slate-100">
                      <tr className="uppercase text-[9px] text-slate-500 font-black">
                        <th className="p-2 sm:p-3 w-10 sm:w-12 text-center">No.</th>
                        <th className="p-2 sm:p-3">Nationality</th>
                        <th className="p-2 sm:p-3 text-center">Inbound (IN)</th>
                        <th className="p-2 sm:p-3 text-center">Outbound (OUT)</th>
                        <th className="p-2 sm:p-3 text-center">Net Staying</th>
                        <th className="p-2 sm:p-3 text-center">Total</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white font-bold">
                      {nationalityReportList.map((r, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                          <td className="p-2 sm:p-3 text-center text-slate-400">{idx + 1}</td>
                          <td className="p-2 sm:p-3 text-slate-900 font-extrabold">{r.nationality}</td>
                          <td className="p-2 sm:p-3 text-center text-emerald-700 bg-emerald-50/10">{r.inCount}</td>
                          <td className="p-2 sm:p-3 text-center text-orange-700 bg-orange-50/10">{r.outCount}</td>
                          <td className={`p-2 sm:p-3 text-center text-xs font-black ${r.net >= 0 ? 'text-emerald-800' : 'text-red-700'}`}>
                            {r.net > 0 ? `+${r.net}` : r.net}
                          </td>
                          <td className="p-2 sm:p-3 text-center text-indigo-900 font-black bg-indigo-50/10">{r.total}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* 2. Visa Type distribution Report */}
            <div className="card shadow-md border border-slate-200 bg-white p-4 sm:p-6">
              <div className="border-b pb-3 sm:pb-4 mb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                <div>
                  <h4 className="text-xs sm:text-sm font-black text-slate-800 uppercase tracking-tight">2. Visa Type Distribution</h4>
                  <p className="text-[10px] text-slate-500 font-bold mt-0.5">Categorization of visitor entry types</p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportVisaExcel}
                    className="btn bg-emerald-700 text-white hover:bg-emerald-800 uppercase text-[9px] px-2.5 py-1.5 flex items-center gap-1 shadow"
                    title="Export Visa Report to Excel"
                  >
                    <Download size={10} /> Excel
                  </button>
                  <button
                    onClick={printVisaReport}
                    className="btn bg-gray-900 text-white hover:bg-black uppercase text-[9px] px-2.5 py-1.5 flex items-center gap-1 shadow"
                    title="Print Visa Report"
                  >
                    <Printer size={10} /> Print
                  </button>
                  <span className="text-[10px] bg-slate-100 border border-slate-200 text-slate-700 font-black px-2 py-1 sm:px-2.5 sm:py-1.5 rounded-full shrink-0">
                    {visaReportList.length} Categories
                  </span>
                </div>
              </div>

              {visaReportList.length === 0 ? (
                <div className="text-center py-12 text-slate-400 font-bold text-xs italic">
                  No records matching selected date range filter.
                </div>
              ) : (
                <div className="overflow-x-auto border border-slate-100 rounded-xl">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-50 border-b border-slate-100">
                      <tr className="uppercase text-[9px] text-slate-500 font-black">
                        <th className="p-2 sm:p-3 w-10 sm:w-12 text-center">No.</th>
                        <th className="p-2 sm:p-3">Visa Type</th>
                        <th className="p-2 sm:p-3 text-center">Inbound (IN)</th>
                        <th className="p-2 sm:p-3 text-center">Outbound (OUT)</th>
                        <th className="p-2 sm:p-3 text-center">Total Movements</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white font-bold">
                      {visaReportList.map((r, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                          <td className="p-2 sm:p-3 text-center text-slate-400">{idx + 1}</td>
                          <td className="p-2 sm:p-3 text-slate-900 font-extrabold">{r.visaType}</td>
                          <td className="p-2 sm:p-3 text-center text-emerald-700">{r.inCount}</td>
                          <td className="p-2 sm:p-3 text-center text-orange-700">{r.outCount}</td>
                          <td className="p-2 sm:p-3 text-center text-indigo-900 font-black bg-indigo-50/10">{r.total}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

/* --- CHECKING HISTORY INTERFACE --- */
interface CheckingHistoryEntry {
  id: string;
  passport: string;
  fullname: string;
  nationality: string;
  type: 'CO_LTD' | 'OTHERS';
  checkDate: string; // YYYY-MM-DD
  originalAddress: string;
  confirmedAddress: string;
  confirmedStayDescription?: string;
  status: 'STILL PERMITTED' | 'STILL NOT PERMITTED YET' | 'STILL CONFIRMED' | 'STAY PERMITTED' | 'STAY NOT PERMITTED' | (string & {});
  permittedBy?: string;
  officerName: string;
  officerTitle: string;
  timestamp: string;
  previousPassport?: string;
  dualPassportRemarks?: string;
  linkedPassports?: string[];
  syncStatus?: 'pending_sync' | 'upload_failed' | 'synced';
  updatedAt?: string;
}

interface CheckingRowProps {
  key?: React.Key;
  m?: MovementData;
  h?: CheckingHistoryEntry;
  isCo: boolean;
  currentUser?: { name: string; title: string; } | null;
  cloudAuthUser?: CloudAuthUser | null;
  isViewer?: boolean;
  showToast?: (msg: string) => void;
  onCheckSubmitted: (entry: CheckingHistoryEntry) => void;
  onCheckUpdated?: (entry: CheckingHistoryEntry) => void;
  onCheckRemoved?: (id: string, passport: string) => void;
  records: ImmRecord[];
  syncMaster?: (value: string, type: 'Nationality' | 'Visa' | 'Stay' | 'Vehicle' | 'Agent' | 'Contact' | 'Official' | 'Title' | 'Reporter' | 'Phone' | 'PermitDescription', linkedValue?: string) => void;
  masterData?: MasterItem[];
}

/* --- MOBILE CHECKING CARD COMPONENT --- */
const CheckingCard = ({
  m,
  h,
  isCo,
  currentUser,
  cloudAuthUser,
  isViewer,
  showToast,
  onCheckSubmitted,
  onCheckUpdated,
  onCheckRemoved,
  records,
  syncMaster,
  masterData
}: CheckingRowProps): React.JSX.Element => {
  const [isExpanded, setIsExpanded] = useState(false);
  const [verificationType, setVerificationType] = useState<'CO_LTD' | 'OTHERS'>(() => h ? h.type : (isCo ? 'CO_LTD' : 'OTHERS'));
  
  const pNo = h ? h.passport : (m?.p || '');
  const name = h ? h.fullname : (m?.n || '');
  const nat = h ? h.nationality : (m?.nat || '');
  const loc = h ? h.originalAddress : (m?.loc || '');
  const start = h ? '' : (m?.start || '');
  const end = h ? '' : (m?.end || '');

  const [coStatus, setCoStatus] = useState<PermitStatus>(() => {
    if (h) return h.status as any;
    return 'နေထိုင်ခွင့်ကျထားသောသူ';
  });
  
  const defaultPermittedBy = useMemo(() => {
    if (h) return h.permittedBy || '';
    const matched = records.filter(r => r.passport.toUpperCase() === pNo.toUpperCase())
      .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
    return matched[0]?.permittedBy || '';
  }, [records, pNo, h]);

  const [permittedBy, setPermittedBy] = useState(defaultPermittedBy);
  const [prevPassport, setPrevPassport] = useState(() => (h as any)?.previousPassport || (m as any)?.previousPassport || "");
  const [dualRemarks, setDualRemarks] = useState(() => (h as any)?.dualPassportRemarks || (m as any)?.dualPassportRemarks || "");
  const [confirmedAddress, setConfirmedAddress] = useState(() => h ? h.confirmedAddress : loc);
  const [confirmedStayDescription, setConfirmedStayDescription] = useState(() => {
    if (h) return h.confirmedStayDescription || '';
    const masterMatch = masterData?.find(x => x.type === 'Stay' && x.name.toLowerCase() === loc.toLowerCase());
    if (masterMatch?.linkedValue) return masterMatch.linkedValue;
    const recMatch = records.find(r => r.address?.toLowerCase() === loc.toLowerCase() && r.stayDescription);
    return recMatch?.stayDescription || '';
  });
  const [checkDate, setCheckDate] = useState(() => h ? h.checkDate : new Date().toISOString().split('T')[0]);
  const [officerName, setOfficerName] = useState(() => h ? h.officerName : (currentUser?.name || localStorage.getItem('lastCheckedOfficerName') || ''));
  const [officerTitle, setOfficerTitle] = useState(() => h ? h.officerTitle : (currentUser?.title || localStorage.getItem('lastCheckedOfficerTitle') || 'Officer'));

  useEffect(() => {
    if (currentUser && !h) {
      setOfficerName(currentUser.name);
      setOfficerTitle(currentUser.title);
    }
  }, [currentUser, h]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (isViewer) {
      if (showToast) showToast("🔒 Viewer level ဖြစ်သောကြောင့် Stay Verification သစ် ဖြည့်သွင်း/ပြင်ဆင်၍မရပါ");
      return;
    }
    if (!coStatus) return;

    if (!currentUser) {
      if (officerName) localStorage.setItem('lastCheckedOfficerName', officerName);
      if (officerTitle) localStorage.setItem('lastCheckedOfficerTitle', officerTitle);
    }

    if (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' && permittedBy.trim() && syncMaster) {
      syncMaster(permittedBy.trim(), 'PermitDescription');
    }
    if (confirmedAddress.trim() && syncMaster) {
      syncMaster(confirmedAddress.trim(), 'Stay', confirmedStayDescription);
    }
    if (officerName.trim() && syncMaster) {
      syncMaster(officerName.trim(), 'Official');
    }
    if (officerTitle.trim() && syncMaster) {
      syncMaster(officerTitle.trim(), 'Title');
    }

    if (h) {
      const updatedCheck: CheckingHistoryEntry = {
        ...h,
        confirmedAddress,
        confirmedStayDescription,
        status: coStatus as any,
        permittedBy: (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') ? permittedBy : '',
        officerName: officerName || 'System Officer',
        officerTitle: officerTitle || 'Officer',
        checkDate,
        type: verificationType,
        previousPassport: prevPassport ? prevPassport.toUpperCase().trim() : undefined,
        dualPassportRemarks: dualRemarks ? dualRemarks.trim() : undefined,
        linkedPassports: prevPassport ? Array.from(new Set([pNo.toUpperCase(), prevPassport.toUpperCase().trim()])) : (h as any)?.linkedPassports,
      };
      if (onCheckUpdated) onCheckUpdated(updatedCheck);
    } else {
      const newCheck: CheckingHistoryEntry = {
        id: Date.now().toString() + '_' + Math.random().toString(36).substring(2, 9),
        passport: pNo.toUpperCase(),
        fullname: name,
        nationality: nat,
        type: verificationType,
        checkDate,
        originalAddress: loc,
        confirmedAddress: confirmedAddress,
        confirmedStayDescription: confirmedStayDescription,
        status: coStatus as any,
        permittedBy: (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') ? permittedBy : '',
        officerName: officerName || 'System Officer',
        officerTitle: officerTitle || 'Officer',
        previousPassport: prevPassport ? prevPassport.toUpperCase().trim() : undefined,
        dualPassportRemarks: dualRemarks ? dualRemarks.trim() : undefined,
        linkedPassports: prevPassport ? Array.from(new Set([pNo.toUpperCase(), prevPassport.toUpperCase().trim()])) : undefined,
        timestamp: new Date().toISOString()
      };
      onCheckSubmitted(newCheck);
    }
    setIsExpanded(false);
  };

  return (
    <div className={`p-3.5 sm:p-4 rounded-2xl border shadow-xs transition-all space-y-3 ${
      h ? 'bg-emerald-50/30 border-emerald-200/80' : 'bg-amber-50/30 border-amber-200/80'
    }`}>
      {/* Top row: Status Tag + Passport */}
      <div className="flex items-center justify-between gap-2 border-b border-slate-200/60 pb-2">
        <div>
          {h ? (
            h.type === 'CO_LTD' ? (
              <span className="text-[10px] bg-emerald-100 text-emerald-900 border border-emerald-300/80 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider block w-fit">
                ✅ Corp. Checked
              </span>
            ) : (
              <span className="text-[10px] bg-teal-100 text-teal-900 border border-teal-300/80 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider block w-fit">
                ✅ Std. Checked
              </span>
            )
          ) : (
            verificationType === 'CO_LTD' ? (
              <span className="text-[10px] bg-amber-100 text-amber-900 border border-amber-300/80 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider block w-fit animate-pulse">
                ⏳ Corp. Pending
              </span>
            ) : (
              <span className="text-[10px] bg-indigo-100 text-indigo-900 border border-indigo-300/80 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider block w-fit animate-pulse">
                ⏳ Std. Pending
              </span>
            )
          )}
        </div>
        <div className="font-black text-[#1A365D] text-xs sm:text-sm font-mono tracking-tight bg-white px-2.5 py-1 rounded-lg border border-slate-200 shadow-2xs">
          {pNo}
        </div>
      </div>

      {/* Guest Name & Nationality */}
      <div className="flex items-start justify-between gap-2">
        <div className="space-y-0.5">
          <div className="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-tight">{name}</div>
          <div className="text-[10px] sm:text-[11px] font-extrabold text-slate-500 uppercase">
            Nat: <span className="text-slate-800">{nat}</span>
          </div>
        </div>
        {h ? (
          <div className="text-right text-[10px] sm:text-[11px] font-mono">
            <div className="font-bold text-slate-700">Checked: {checkDate}</div>
            <div className="text-[9px] sm:text-[10px] text-slate-500 font-sans font-medium">{officerTitle} {officerName}</div>
          </div>
        ) : (
          <div className="text-right text-[10px] sm:text-[11px] font-mono">
            <div className="text-slate-400">Exp To:</div>
            <div className="font-black text-[#1A365D]">{end || 'N/A'}</div>
            <div className="text-[9px] font-bold text-amber-600 uppercase">({getLiveRemainingDays(end)})</div>
          </div>
        )}
      </div>

      {/* Address */}
      <div className="bg-white p-2.5 rounded-xl border border-slate-200/80 text-xs">
        <div className="font-extrabold text-slate-900 text-xs flex items-start gap-1.5 leading-snug">
          <span className="shrink-0">📍</span> <span>{confirmedAddress}</span>
        </div>
        {confirmedStayDescription && (
          <div className="text-[10px] text-slate-500 font-mono mt-1 italic pl-5">
            🏡 Detail: {confirmedStayDescription}
          </div>
        )}
      </div>

      {/* Status Details & Actions */}
      <div className="flex items-center justify-between text-xs pt-0.5">
        {h ? (
          <div>
            {coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' ? (
              <span className="text-[9px] font-black uppercase text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-md px-2 py-0.5">
                STAY PERMITTED
              </span>
            ) : coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? (
              <span className="text-[9px] font-black uppercase text-rose-800 bg-rose-50 border border-rose-200 rounded-md px-2 py-0.5">
                STAY NOT PERMITTED
              </span>
            ) : (
              <span className="text-[9px] font-black uppercase text-blue-800 bg-blue-50 border border-blue-200 rounded-md px-2 py-0.5">
                STILL CONFIRMED
              </span>
            )}
            {permittedBy && (
              <div className="text-[9px] text-indigo-900 font-bold mt-0.5">
                Desc: <span className="text-purple-700">{permittedBy}</span>
              </div>
            )}
          </div>
        ) : (
          <span className="text-[10px] bg-amber-50 text-amber-800 font-black rounded-lg border border-amber-200 px-2.5 py-1">
            ⚠️ Awaiting Verification
          </span>
        )}

        <div className="flex items-center gap-1.5">
          {h ? (
            <>
              <button
                type="button"
                onClick={() => setIsExpanded(!isExpanded)}
                className={`text-[10px] font-black uppercase px-3 py-1.5 rounded-xl border transition-all min-h-[34px] ${
                  isExpanded
                    ? 'bg-slate-100 border-slate-300 text-slate-700'
                    : 'bg-amber-500 text-white border-amber-600 hover:bg-amber-600'
                }`}
              >
                {isExpanded ? 'Cancel' : 'Edit Log'}
              </button>
              {onCheckRemoved && (
                <button
                  type="button"
                  onClick={() => onCheckRemoved(h.id, h.passport)}
                  className="text-[10px] font-black uppercase text-rose-700 bg-rose-50 border border-rose-200 px-3 py-1.5 rounded-xl hover:bg-rose-100 min-h-[34px]"
                >
                  Undo
                </button>
              )}
            </>
          ) : (
            <button
              type="button"
              onClick={() => setIsExpanded(!isExpanded)}
              className={`text-[10px] font-black uppercase px-3.5 py-2 rounded-xl transition-all shadow-xs text-center min-h-[36px] ${
                isExpanded 
                  ? 'bg-slate-100 text-slate-700 border border-slate-300' 
                  : (verificationType === 'CO_LTD')
                  ? 'bg-amber-600 text-white hover:bg-amber-700'
                  : 'bg-indigo-600 text-white hover:bg-indigo-700'
              }`}
            >
              {isExpanded ? 'Close Form' : 'Verify Physical'}
            </button>
          )}
        </div>
      </div>

      {/* Mobile Form Overlay */}
      {isExpanded && (
        <motion.div
          initial={{ opacity: 0, height: 0 }}
          animate={{ opacity: 1, height: 'auto' }}
          exit={{ opacity: 0, height: 0 }}
          transition={{ duration: 0.2 }}
          className="pt-2 border-t border-slate-200"
        >
          <form onSubmit={handleSubmit} className="space-y-3 bg-white p-3.5 rounded-xl border border-slate-200 shadow-inner text-xs">
            <div className="text-xs font-black text-slate-800 uppercase flex items-center justify-between border-b pb-2">
              <span>{h ? '✏️ Edit Stay Verification' : '🔍 Stay Verification Check'}</span>
              <span className="font-mono text-indigo-900 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">{pNo}</span>
            </div>

            <div className="space-y-1.5">
              <label className="block text-[10px] font-black uppercase text-slate-600">
                🏢 Classification (အမျိုးအစား ရွေးရန်)
              </label>
              <div className="grid grid-cols-2 gap-2">
                <label className={`flex items-center justify-center p-2.5 rounded-xl border text-[10px] sm:text-xs font-black text-center cursor-pointer transition-all min-h-[40px] ${
                  verificationType === 'CO_LTD'
                    ? 'border-amber-500 bg-amber-50 text-amber-900 shadow-2xs'
                    : 'border-slate-200 bg-slate-50 text-slate-600'
                }`}>
                  <input
                    type="radio"
                    name={`mVerificationType-${pNo}`}
                    value="CO_LTD"
                    checked={verificationType === 'CO_LTD'}
                    onChange={() => {
                      setVerificationType('CO_LTD');
                      setCoStatus('နေထိုင်ခွင့်ကျထားသောသူ');
                    }}
                    className="sr-only"
                  />
                  🏢 CO. LTD. (ကုမ္ပဏီ)
                </label>
                <label className={`flex items-center justify-center p-2.5 rounded-xl border text-[10px] sm:text-xs font-black text-center cursor-pointer transition-all min-h-[40px] ${
                  verificationType === 'OTHERS'
                    ? 'border-indigo-500 bg-indigo-50 text-indigo-900 shadow-2xs'
                    : 'border-slate-200 bg-slate-50 text-slate-600'
                }`}>
                  <input
                    type="radio"
                    name={`mVerificationType-${pNo}`}
                    value="OTHERS"
                    checked={verificationType === 'OTHERS'}
                    onChange={() => {
                      setVerificationType('OTHERS');
                      setCoStatus('နေထိုင်ခွင့်ကျထားသောသူ');
                    }}
                    className="sr-only"
                  />
                  🏨 HOTEL (ဟိုတယ်)
                </label>
              </div>
            </div>

            <div>
              <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                🏡 Address (နေထိုင်ရာလိပ်စာ)
              </label>
              <input
                type="text"
                value={confirmedAddress}
                list="checkingStayList"
                onChange={(e) => {
                  const val = e.target.value;
                  const match = masterData?.find(m => m.type === 'Stay' && m.name.toLowerCase() === val.toLowerCase());
                  setConfirmedAddress(val);
                  setConfirmedStayDescription(match ? (match.linkedValue || '') : confirmedStayDescription);
                }}
                className="input-field bg-slate-50 text-xs font-bold rounded-xl py-2 px-3 border-slate-200"
                placeholder="Address..."
                required
              />
            </div>

            <div>
              <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                🏡 Detail Description (လိပ်စာအသေးစိတ်)
              </label>
              <input
                type="text"
                value={confirmedStayDescription}
                list="checkingStayDescList"
                onChange={(e) => setConfirmedStayDescription(e.target.value)}
                className="input-field bg-slate-50 text-xs font-bold rounded-xl py-2 px-3 border-slate-200"
                placeholder="Detail stay info..."
              />
            </div>

            <div>
              <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                🟢 Status (ခွင့်ပြုချက် အခြေအနေ)
              </label>
              <select
                value={coStatus}
                onChange={(e) => {
                  const newStatus = e.target.value as any;
                  setCoStatus(newStatus);
                  const hist = records.filter(r => r.passport.toUpperCase() === pNo.toUpperCase() && (r.stillPermittedStatus === newStatus || (newStatus === 'STAY PERMITTED' && r.stillPermittedStatus === 'STILL PERMITTED')))
                    .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
                  setPermittedBy(hist?.permittedBy || '');
                }}
                className="input-field bg-slate-50 font-black text-xs uppercase rounded-xl py-2 px-3 border-slate-200"
                required
              >
                <option value="STAY PERMITTED">STAY PERMITTED (ခွင့်ပြုထားဆဲ - သီးခြားနေထိုင်ခွင့်ပြုသူ)</option>
                <option value="STAY NOT PERMITTED">STAY NOT PERMITTED (ခွင့်မပြုသေးပါ)</option>
              </select>
            </div>

            {(coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') && (
              <div>
                <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                  📝 {coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' ? 'Permit Description' : 'Reason / Description'}
                </label>
                <input
                  type="text"
                  value={permittedBy}
                  list={coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' ? "checkingPermitDescriptionList" : undefined}
                  onChange={(e) => setPermittedBy(e.target.value)}
                  className="input-field bg-slate-50 font-bold rounded-xl text-xs py-2 px-3 border-slate-200"
                  placeholder="Reason / Description..."
                  required={coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'}
                />
              </div>
            )}

            {/* Dual / Previous Passport Linking in Card */}
            <div className="p-2.5 bg-purple-50/60 border border-purple-200/80 rounded-xl space-y-2">
              <label className="text-[10px] font-black uppercase text-purple-900 flex items-center gap-1">
                <Link2 size={12} className="text-purple-600" />
                🔗 Linked / Previous Passport (ယခင် Passport နံပါတ်ဟောင်း)
              </label>
              <input
                type="text"
                value={prevPassport}
                onChange={(e) => setPrevPassport(e.target.value.toUpperCase())}
                className="input-field bg-white font-mono font-bold text-xs uppercase rounded-lg py-1.5 px-2.5 border-purple-200 text-purple-950"
                placeholder="e.g. OLD PASSPORT NO."
              />
              <input
                type="text"
                value={dualRemarks}
                onChange={(e) => setDualRemarks(e.target.value)}
                className="input-field bg-white font-medium text-xs rounded-lg py-1.5 px-2.5 border-purple-200 text-slate-800"
                placeholder="Dual Passport Remarks / မှတ်ချက်..."
              />
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                  📅 Check Date
                </label>
                <input
                  type="date"
                  value={checkDate}
                  onChange={(e) => setCheckDate(e.target.value)}
                  className="input-field bg-slate-50 font-bold rounded-xl text-xs py-2 px-3 border-slate-200"
                  required
                />
              </div>
              <div>
                <label className="block text-[10px] font-black uppercase text-slate-600 mb-1">
                  👤 Officer Name
                </label>
                <input
                  type="text"
                  list="checkingOfficialList"
                  value={officerName}
                  onChange={(e) => setOfficerName(e.target.value)}
                  className="input-field bg-slate-50 font-bold text-xs rounded-xl py-2 px-3 border-slate-200"
                  placeholder="Officer"
                  required
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setIsExpanded(false)}
                className="btn bg-slate-100 hover:bg-slate-200 text-slate-700 font-extrabold text-[11px] uppercase px-4 py-2 rounded-xl"
              >
                Close
              </button>
              <button
                type="submit"
                className="btn bg-[#1A365D] hover:bg-black text-white font-black text-[11px] uppercase px-5 py-2 rounded-xl shadow-xs"
              >
                Save
              </button>
            </div>
          </form>
        </motion.div>
      )}
    </div>
  );
};

/* --- DETAILED CHECKING ROW COMPONENT --- */
const CheckingRow = ({
  m,
  h,
  isCo,
  currentUser,
  cloudAuthUser,
  isViewer,
  showToast,
  onCheckSubmitted,
  onCheckUpdated,
  onCheckRemoved,
  records,
  syncMaster,
  masterData
}: CheckingRowProps): React.JSX.Element => {
  const [isExpanded, setIsExpanded] = useState(false);
  const [verificationType, setVerificationType] = useState<'CO_LTD' | 'OTHERS'>(() => h ? h.type : (isCo ? 'CO_LTD' : 'OTHERS'));
  
  const pNo = h ? h.passport : (m?.p || '');
  const name = h ? h.fullname : (m?.n || '');
  const nat = h ? h.nationality : (m?.nat || '');
  const loc = h ? h.originalAddress : (m?.loc || '');
  const start = h ? '' : (m?.start || '');
  const end = h ? '' : (m?.end || '');
  const allowed = h ? '' : (m?.allowed || '');

  const [coStatus, setCoStatus] = useState<PermitStatus>(() => {
    if (h) return h.status as any;
    return 'နေထိုင်ခွင့်ကျထားသောသူ';
  });
  
  // Prefill permittedBy from past records of this guest and map status
  const defaultPermittedBy = useMemo(() => {
    if (h) return h.permittedBy || '';
    const matched = records.filter(r => r.passport.toUpperCase() === pNo.toUpperCase())
      .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
    return matched[0]?.permittedBy || '';
  }, [records, pNo, h]);

  const [permittedBy, setPermittedBy] = useState(defaultPermittedBy);
  const [prevPassport, setPrevPassport] = useState(() => (h as any)?.previousPassport || (m as any)?.previousPassport || "");
  const [dualRemarks, setDualRemarks] = useState(() => (h as any)?.dualPassportRemarks || (m as any)?.dualPassportRemarks || "");
  const [confirmedAddress, setConfirmedAddress] = useState(() => h ? h.confirmedAddress : loc);
  const [confirmedStayDescription, setConfirmedStayDescription] = useState(() => {
    if (h) return h.confirmedStayDescription || '';
    const masterMatch = masterData?.find(x => x.type === 'Stay' && x.name.toLowerCase() === loc.toLowerCase());
    if (masterMatch?.linkedValue) return masterMatch.linkedValue;
    const recMatch = records.find(r => r.address?.toLowerCase() === loc.toLowerCase() && r.stayDescription);
    return recMatch?.stayDescription || '';
  });
  const [checkDate, setCheckDate] = useState(() => h ? h.checkDate : new Date().toISOString().split('T')[0]);
  const [officerName, setOfficerName] = useState(() => h ? h.officerName : (currentUser?.name || localStorage.getItem('lastCheckedOfficerName') || ''));
  const [officerTitle, setOfficerTitle] = useState(() => h ? h.officerTitle : (currentUser?.title || localStorage.getItem('lastCheckedOfficerTitle') || 'Officer'));

  // Sync with currentUser if they just logged in
  useEffect(() => {
    if (currentUser && !h) {
      setOfficerName(currentUser.name);
      setOfficerTitle(currentUser.title);
    }
  }, [currentUser, h]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (isViewer) {
      if (showToast) showToast("🔒 Viewer level ဖြစ်သောကြောင့် Stay Verification သစ် ဖြည့်သွင်း/ပြင်ဆင်၍မရပါ");
      return;
    }
    if (!coStatus) return;

    // Cache typed officer in localStorage for future convenience if not logged in
    if (!currentUser) {
      if (officerName) localStorage.setItem('lastCheckedOfficerName', officerName);
      if (officerTitle) localStorage.setItem('lastCheckedOfficerTitle', officerTitle);
    }

    // Auto save the permit description to Master Data (as per requirement) so it becomes reusable
    if (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' && permittedBy.trim() && syncMaster) {
      syncMaster(permittedBy.trim(), 'PermitDescription');
    }

    // Auto save / contribute the stay location address & description to Master Data
    if (confirmedAddress.trim() && syncMaster) {
      syncMaster(confirmedAddress.trim(), 'Stay', confirmedStayDescription);
    }

    // Auto contribute checking officer and title to master data as well
    if (officerName.trim() && syncMaster) {
      syncMaster(officerName.trim(), 'Official');
    }
    if (officerTitle.trim() && syncMaster) {
      syncMaster(officerTitle.trim(), 'Title');
    }

    if (h) {
      // Editing existing log
      const updatedCheck: CheckingHistoryEntry = {
        ...h,
        confirmedAddress,
        confirmedStayDescription,
        status: coStatus as any,
        permittedBy: (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') ? permittedBy : '',
        officerName: officerName || 'System Officer',
        officerTitle: officerTitle || 'Officer',
        checkDate,
        type: verificationType,
      };
      if (onCheckUpdated) onCheckUpdated(updatedCheck);
    } else {
      // Create new verification entry
      const newCheck: CheckingHistoryEntry = {
        id: Date.now().toString() + '_' + Math.random().toString(36).substring(2, 9),
        passport: pNo.toUpperCase(),
        fullname: name,
        nationality: nat,
        type: verificationType,
        checkDate,
        originalAddress: loc,
        confirmedAddress: confirmedAddress,
        confirmedStayDescription: confirmedStayDescription,
        status: coStatus as any,
        permittedBy: (coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') ? permittedBy : '',
        officerName: officerName || 'System Officer',
        officerTitle: officerTitle || 'Officer',
        previousPassport: prevPassport ? prevPassport.toUpperCase().trim() : undefined,
        dualPassportRemarks: dualRemarks ? dualRemarks.trim() : undefined,
        linkedPassports: prevPassport ? Array.from(new Set([pNo.toUpperCase(), prevPassport.toUpperCase().trim()])) : undefined,
        timestamp: new Date().toISOString()
      };
      onCheckSubmitted(newCheck);
    }
    setIsExpanded(false);
  };

  return (
    <>
      <tr className={`hover:bg-slate-50/50 transition-all border-b border-gray-100 ${h ? 'bg-emerald-50/10' : 'bg-amber-50/5'}`}>
        <td className="p-4 align-middle">
          {h ? (
            h.type === 'CO_LTD' ? (
              <span className="text-[9px] bg-emerald-100 text-emerald-800 border border-emerald-300 px-2.5 py-1 rounded-full font-black uppercase tracking-wider block w-fit">
                ✅ Corp. Checked
              </span>
            ) : (
              <span className="text-[9px] bg-teal-100 text-teal-800 border border-teal-200 px-2.5 py-1 rounded-full font-black uppercase tracking-wider block w-fit">
                ✅ Std. Checked
              </span>
            )
          ) : (
            verificationType === 'CO_LTD' ? (
              <span className="text-[9px] bg-amber-50 text-amber-800 border border-amber-200 px-2.5 py-1 rounded-full font-black uppercase tracking-wider block w-fit animate-pulse">
                ⏳ Corp. Pending
              </span>
            ) : (
              <span className="text-[9px] bg-indigo-50 text-indigo-800 border border-indigo-200 px-2.5 py-1 rounded-full font-black uppercase tracking-wider block w-fit animate-pulse">
                ⏳ Std. Pending
              </span>
            )
          )}
        </td>
        <td className="p-4 align-middle">
          <div className="font-extrabold text-[#1A365D] tracking-tight text-sm font-mono flex items-center gap-1.5 leading-none">
            {pNo}
          </div>
          <div className="text-[11px] font-black text-slate-500 uppercase mt-1 tracking-wide">{name}</div>
        </td>
        <td className="p-4 align-middle font-black text-slate-700 uppercase">{nat}</td>
        <td className="p-4 align-middle leading-tight">
          <div className="font-extrabold text-slate-900 text-[11px]">{confirmedAddress}</div>
          {confirmedStayDescription && (
            <div className="text-[9px] text-gray-500 font-mono mt-1 italic">
              🏡 Detail: {confirmedStayDescription}
            </div>
          )}
        </td>
        <td className="p-4 align-middle">
          {h ? (
            <>
              <div className="mb-1">
                {coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' ? (
                  <span className="text-[9px] font-black uppercase text-emerald-800 bg-emerald-50 border border-emerald-200 rounded px-1.5 py-0.5">
                    STAY PERMITTED
                  </span>
                ) : coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? (
                  <span className="text-[9px] font-black uppercase text-rose-800 bg-rose-50 border border-rose-200 rounded px-1.5 py-0.5">
                    STAY NOT PERMITTED
                  </span>
                ) : (
                  <span className="text-[9px] font-black uppercase text-blue-800 bg-blue-50 border border-blue-200 rounded px-1.5 py-0.5">
                    STILL CONFIRMED
                  </span>
                )}
              </div>
              {permittedBy && (
                <div className="text-[9px] text-indigo-900 font-black tracking-tight mt-0.5">
                  Desc: <span className="text-purple-700">{permittedBy}</span>
                </div>
              )}
            </>
          ) : (
            <span className="text-[10px] bg-amber-50 text-amber-700 font-extrabold rounded border border-amber-200 px-1.5 py-0.5">
              ⚠️ Awaiting Verification
            </span>
          )}
        </td>
        <td className="p-4 align-middle text-xs text-slate-700">
          {h ? (
            <div className="leading-normal">
              <div className="font-extrabold text-indigo-950 flex items-center gap-1">
                <Calendar size={12} className="text-indigo-400" />
                Checked: {checkDate}
              </div>
              <div className="text-[9px] text-slate-400 font-mono font-bold mt-0.5">
                Officer: {officerTitle} {officerName}
              </div>
            </div>
          ) : (
            <div className="leading-normal font-mono text-[10px]">
              <div className="text-gray-400">Exp Stay To:</div>
              <div className="font-extrabold text-[#1A365D]">{end || 'N/A'}</div>
              <div className="text-[8px] font-bold text-amber-600 mt-0.5 uppercase">
                ({getLiveRemainingDays(end)})
              </div>
            </div>
          )}
        </td>
        <td className="p-4 align-middle text-right grid grid-cols-1 gap-1 min-w-[120px]">
          {h ? (
            <div className="flex justify-end gap-1.5">
              <button
                onClick={() => setIsExpanded(!isExpanded)}
                className={`text-[8.5px] font-black uppercase px-2.5 py-1.5 rounded-lg border transition-all ${
                  isExpanded
                    ? 'bg-slate-100 border-slate-300 text-slate-600'
                    : 'bg-orange-55 text-orange-900 border-orange-200 hover:bg-orange-100 hover:text-orange-950'
                }`}
              >
                {isExpanded ? 'Cancel' : 'Edit Details'}
              </button>
              {onCheckRemoved && (
                <button
                  onClick={() => onCheckRemoved(h.id, h.passport)}
                  className="text-[8.5px] font-black uppercase text-rose-700 bg-rose-50 border border-rose-200 px-2.5 py-1.5 rounded-lg hover:bg-rose-100 hover:border-rose-300 transition-all"
                  title="Undo check verification and re-queue"
                >
                  Undo Log
                </button>
              )}
            </div>
          ) : (
            <button
              onClick={() => setIsExpanded(!isExpanded)}
              className={`text-[9px] font-black uppercase px-4 py-2 rounded-xl transition-all shadow-xs w-full text-center ${
                isExpanded 
                  ? 'bg-slate-100 text-slate-600 border border-slate-200 hover:bg-slate-200' 
                  : (verificationType === 'CO_LTD')
                  ? 'bg-amber-600 text-white hover:bg-amber-700 hover:shadow-sm'
                  : 'bg-indigo-600 text-white hover:bg-indigo-700 hover:shadow-sm'
              }`}
            >
              {isExpanded ? 'Close' : 'Verify Physical'}
            </button>
          )}
        </td>
      </tr>
      
      {isExpanded && (
        <tr className="bg-slate-50/50">
          <td colSpan={7} className="p-0 border-b border-slate-100 shadow-inner">
            <motion.div
              initial={{ opacity: 0, height: 0 }}
              animate={{ opacity: 1, height: 'auto' }}
              exit={{ opacity: 0, height: 0 }}
              transition={{ duration: 0.2 }}
              className="p-6 bg-white overflow-hidden"
            >
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="flex items-center justify-between border-b pb-2">
                  <span className="text-xs uppercase font-black text-slate-800 flex items-center gap-2">
                    <span className={`px-2 py-0.5 rounded text-[9px] font-black border ${
                      h 
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-300' 
                        : (verificationType === 'CO_LTD')
                        ? 'bg-amber-50 text-amber-800 border-amber-200' 
                        : 'bg-indigo-50 text-indigo-800 border-indigo-200'
                    }`}>
                      {h ? 'EDIT VERIFICATION DETAILED DIALOG' : (verificationType === 'CO_LTD') ? 'CO LTD VERIFICATION CHECK' : 'STANDARD STAY CHECK'}
                    </span>
                    Stay verification ledger update for passport: <span className="font-mono text-purple-700 font-black">{pNo}</span> ({name})
                  </span>
                  {!h && <span className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">Stay Schedule: {start} to {end}</span>}
                </div>

                {/* Live Category Selector */}
                <div className="bg-slate-50 border border-slate-200/60 p-4 rounded-xl">
                  <label className="block text-[10px] font-black uppercase text-slate-500 tracking-wider mb-2">
                    🏢 Live Classification / Class Verification (နေထိုင်မှု စစ်ဆေးသည့် အမျိုးအစား ရွေးချယ်ရန်)
                  </label>
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className={`flex-1 flex items-center justify-center gap-2 p-2.5 rounded-xl border-2 cursor-pointer text-[11px] font-black tracking-wide transition-all ${
                      verificationType === 'CO_LTD'
                        ? 'border-amber-500 bg-amber-50 text-amber-900 shadow-xs'
                        : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'
                    }`}>
                      <input
                        type="radio"
                        name={`verificationType-${pNo}`}
                        value="CO_LTD"
                        checked={verificationType === 'CO_LTD'}
                        onChange={() => {
                          setVerificationType('CO_LTD');
                          setCoStatus('နေထိုင်ခွင့်ကျထားသောသူ'); // set suitable default
                        }}
                        className="sr-only"
                      />
                      🏢 CORPORATE CO. LTD. (ကုမ္ပဏီ နေထိုင်သူ)
                    </label>
                    <label className={`flex-1 flex items-center justify-center gap-2 p-2.5 rounded-xl border-2 cursor-pointer text-[11px] font-black tracking-wide transition-all ${
                      verificationType === 'OTHERS'
                        ? 'border-indigo-500 bg-indigo-50 text-indigo-900 shadow-xs'
                        : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'
                    }`}>
                      <input
                        type="radio"
                        name={`verificationType-${pNo}`}
                        value="OTHERS"
                        checked={verificationType === 'OTHERS'}
                        onChange={() => {
                          setVerificationType('OTHERS');
                          setCoStatus('နေထိုင်ခွင့်ကျထားသောသူ'); // set suitable default
                        }}
                        className="sr-only"
                      />
                      🏨 STANDARD / HOTEL STAY (ဟိုတယ်/တည်းခိုခန်း နေထိုင်သူ)
                    </label>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                  {/* Left Form: Status Inputs */}
                  <div className="space-y-4">
                    {verificationType === 'CO_LTD' ? (
                      <>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-amber-900 tracking-wider mb-1.5 rounded bg-amber-50/70 px-2.5 py-1 w-fit border border-amber-200">
                            🏡 Confirm/Update Corporate Stay Address (ကုမ္ပဏီ နေထိုင်ရာလိပ်စာ)
                          </label>
                          <input
                            type="text"
                            value={confirmedAddress}
                            list="checkingStayList"
                            onChange={(e) => {
                              const val = e.target.value;
                              const match = masterData?.find(m => m.type === 'Stay' && m.name.toLowerCase() === val.toLowerCase());
                              setConfirmedAddress(val);
                              setConfirmedStayDescription(match ? (match.linkedValue || '') : confirmedStayDescription);
                            }}
                            className="input-field bg-white border-amber-300 text-xs font-bold focus:ring-amber-500 rounded-lg"
                            placeholder="Enter / confirm corporate address stay..."
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-amber-900 tracking-wider mb-1.5 rounded bg-amber-50/70 px-2.5 py-1 w-fit border border-amber-200">
                            🏡 Detail Address / Stay Description (လိပ်စာအသေးစိတ်)
                          </label>
                          <input
                            type="text"
                            value={confirmedStayDescription}
                            list="checkingStayDescList"
                            onChange={(e) => setConfirmedStayDescription(e.target.value)}
                            className="input-field bg-white border-amber-300 text-xs font-bold focus:ring-amber-500 rounded-lg"
                            placeholder="Detail stay address or description..."
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-amber-900 tracking-wider mb-1.5 rounded bg-amber-50/70 px-2.5 py-1 w-fit border border-amber-200">
                            🟢 Corporate Permit Status (ခွင့်ပြုချက် အခြေအနေ)
                          </label>
                          <select
                            value={coStatus}
                            onChange={(e) => {
                              const newStatus = e.target.value as any;
                              setCoStatus(newStatus);
                              const hist = records.filter(r => r.passport.toUpperCase() === pNo.toUpperCase() && (r.stillPermittedStatus === newStatus || (newStatus === 'STAY PERMITTED' && r.stillPermittedStatus === 'STILL PERMITTED')))
                                .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
                              setPermittedBy(hist?.permittedBy || '');
                            }}
                            className="input-field bg-white border-amber-300 font-black text-[11px] uppercase text-amber-800 focus:ring-amber-500 rounded-lg"
                            required
                          >
                            <option value="STAY PERMITTED">STAY PERMITTED (ခွင့်ပြုထားဆဲ - သီးခြားနေထိုင်ခွင့်ပြုသူ)</option>
                            <option value="STAY NOT PERMITTED">STAY NOT PERMITTED (ခွင့်မပြုသေးပါ)</option>
                          </select>
                        </div>
                        {coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' && (
                          <div className="animate-in slide-in-from-top-1 duration-200">
                            <label className="block text-[10px] font-black uppercase text-slate-500 tracking-wider mb-1.5">
                              📝 PERMIT DESCRIPTION (ခွင့်ပြုချက် အကြောင်းအရာ)
                            </label>
                            <input
                              type="text"
                              value={permittedBy}
                              list="checkingPermitDescriptionList"
                              onChange={(e) => setPermittedBy(e.target.value)}
                              className="input-field bg-white focus:ring-amber-500 border-slate-200 font-bold rounded-lg text-xs"
                              placeholder="e.g. Approved sponsor, MD certified, Stay extension granted..."
                              required={coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'}
                            />
                          </div>
                        )}
                        {coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' && (
                          <div className="animate-in slide-in-from-top-1 duration-200">
                            <label className="block text-[10px] font-black uppercase text-red-900 tracking-wider mb-1.5">
                              📝 NOT PERMIT DESCRIPTION (ခွင့်မပြုသည့်အကြောင်းအရာ)
                            </label>
                            <input
                              type="text"
                              value={permittedBy}
                              onChange={(e) => setPermittedBy(e.target.value)}
                              className="input-field bg-white border-red-200/60 font-bold text-red-800 rounded-lg text-xs"
                              placeholder="Not Permitted reason (ခွင့်မပြုသည့်အကြောင်းအရာ)..."
                            />
                          </div>
                        )}
                      </>
                    ) : (
                      <>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-indigo-900 tracking-wider mb-1.5 rounded bg-indigo-50 px-2.5 py-1 w-fit border border-indigo-200">
                            🏡 Confirm Stay Address Location (အတည်ပြုနေထိုင်ရာလိပ်စာ)
                          </label>
                          <input
                            type="text"
                            value={confirmedAddress}
                            list="checkingStayList"
                            onChange={(e) => {
                              const val = e.target.value;
                              const match = masterData?.find(m => m.type === 'Stay' && m.name.toLowerCase() === val.toLowerCase());
                              setConfirmedAddress(val);
                              setConfirmedStayDescription(match ? (match.linkedValue || '') : confirmedStayDescription);
                            }}
                            className="input-field bg-white border-slate-200 text-xs font-bold focus:ring-indigo-500 rounded-lg"
                            placeholder="Enter / confirm local address stay..."
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-indigo-900 tracking-wider mb-1.5 rounded bg-indigo-50 px-2.5 py-1 w-fit border border-indigo-200">
                            🏡 Detail Address / Stay Description (လိပ်စာအသေးစိတ်)
                          </label>
                          <input
                            type="text"
                            value={confirmedStayDescription}
                            list="checkingStayDescList"
                            onChange={(e) => setConfirmedStayDescription(e.target.value)}
                            className="input-field bg-white border-slate-200 text-xs font-bold focus:ring-indigo-500 rounded-lg pb-1 bg-indigo-50/10"
                            placeholder="Detail stay address or description..."
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black uppercase text-indigo-900 tracking-wider mb-1.5 rounded bg-indigo-50 px-2.5 py-1 w-fit border border-indigo-200">
                            🟢 Assigned Verification Status (ခွင့်ပြုချက် အခြေအနေ)
                          </label>
                          <select
                            value={coStatus}
                            onChange={(e) => {
                              const newStatus = e.target.value as any;
                              setCoStatus(newStatus);
                              const hist = records.filter(r => r.passport.toUpperCase() === pNo.toUpperCase() && (r.stillPermittedStatus === newStatus || (newStatus === 'STAY PERMITTED' && r.stillPermittedStatus === 'STILL PERMITTED')))
                                .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
                              setPermittedBy(hist?.permittedBy || '');
                            }}
                            className="input-field bg-white border-slate-200 font-black text-[11px] uppercase text-indigo-800 focus:ring-indigo-500 rounded-lg"
                            required
                          >
                            <option value="STAY PERMITTED">STAY PERMITTED (ခွင့်ပြုထားဆဲ - သီးခြားနေထိုင်ခွင့်ပြုသူ)</option>
                            <option value="STAY NOT PERMITTED">STAY NOT PERMITTED (ခွင့်မပြုသေးပါ)</option>
                          </select>
                        </div>
                        {coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' && (
                          <div className="animate-in slide-in-from-top-1 duration-200">
                            <label className="block text-[10px] font-black uppercase text-slate-500 tracking-wider mb-1.5">
                              📝 PERMIT DESCRIPTION (ခွင့်ပြုချက် အကြောင်းအရာ)
                            </label>
                            <input
                              type="text"
                              value={permittedBy}
                              list="checkingPermitDescriptionList"
                              onChange={(e) => setPermittedBy(e.target.value)}
                              className="input-field bg-white focus:ring-indigo-500 border-slate-200 font-bold rounded-lg text-xs"
                              placeholder="e.g. Approved sponsor, MD certified, Stay extension granted..."
                              required={coStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'}
                            />
                          </div>
                        )}
                        {coStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' && (
                          <div className="animate-in slide-in-from-top-1 duration-200">
                            <label className="block text-[10px] font-black uppercase text-red-900 tracking-wider mb-1.5">
                              📝 NOT PERMIT DESCRIPTION (ခွင့်မပြုသည့်အကြောင်းအရာ)
                            </label>
                            <input
                              type="text"
                              value={permittedBy}
                              onChange={(e) => setPermittedBy(e.target.value)}
                              className="input-field bg-white border-red-200/60 font-bold text-red-800 rounded-lg text-xs"
                              placeholder="Not Permitted reason (ခွင့်မပြုသည့်အကြောင်းအရာ)..."
                            />
                          </div>
                        )}

                        {/* Dual / Previous Passport in Checking Row */}
                        <div className="p-3 bg-purple-50/60 border border-purple-200/80 rounded-xl space-y-2">
                          <label className="text-[10px] font-black uppercase text-purple-900 flex items-center gap-1">
                            <Link2 size={12} className="text-purple-600" />
                            🔗 Linked / Previous Passport (ယခင် Passport နံပါတ်ဟောင်း)
                          </label>
                          <input
                            type="text"
                            value={prevPassport}
                            onChange={(e) => setPrevPassport(e.target.value.toUpperCase())}
                            className="input-field bg-white font-mono font-bold text-xs uppercase rounded-lg py-1.5 px-2.5 border-purple-200 text-purple-950"
                            placeholder="OLD PASSPORT NO."
                          />
                          <input
                            type="text"
                            value={dualRemarks}
                            onChange={(e) => setDualRemarks(e.target.value)}
                            className="input-field bg-white font-medium text-xs rounded-lg py-1.5 px-2.5 border-purple-200 text-slate-800"
                            placeholder="Dual Passport Remarks / မှတ်ချက်..."
                          />
                        </div>
                      </>
                    )}
                  </div>

                  {/* Right Form: Date Picker & Responsible Officer */}
                  <div className="space-y-4">
                    <div>
                      <label className="block text-[10px] font-black uppercase text-slate-500 tracking-wider mb-1.5">
                        📅 Check Date Log (စစ်ဆေးသည့် ရက်စွဲ)
                      </label>
                      <input
                        type="date"
                        value={checkDate}
                        onChange={(e) => setCheckDate(e.target.value)}
                        className="input-field bg-white border-slate-200 font-bold rounded-lg text-xs"
                        required
                      />
                    </div>

                    <div className="p-4 bg-slate-50 border border-slate-200/60 space-y-3 rounded-2xl">
                      <span className="text-[9px] font-black uppercase text-indigo-900 flex items-center gap-1 bg-indigo-50 px-2 py-0.5 border border-indigo-200 rounded w-fit">
                        👤 Responsible Checker Officer (စစ်ဆေးသူ အရာရှိ)
                      </span>
                      <div className="grid grid-cols-2 gap-3 text-xs">
                        <div>
                          <label className="input-label text-purple-850 font-black">Official Name (စာရင်းသွင်းသူ)</label>
                          <input
                            type="text"
                            list="checkingOfficialList"
                            value={officerName}
                            onChange={(e) => setOfficerName(e.target.value)}
                            className="input-field bg-white border-slate-200 text-gray-900 font-bold"
                            placeholder="Officer Name"
                            required
                          />
                        </div>
                        <div>
                          <label className="input-label text-purple-850 font-black">Official Title (ရာထူး)</label>
                          <input
                            type="text"
                            list="checkingTitleList"
                            value={officerTitle}
                            onChange={(e) => setOfficerTitle(e.target.value)}
                            className="input-field bg-white border-slate-200 text-gray-900 font-bold"
                            placeholder="Title"
                            required
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="flex justify-end gap-3 border-t pt-4 mt-2">
                  <button
                    type="button"
                    onClick={() => setIsExpanded(false)}
                    className="btn bg-gray-100 text-gray-600 hover:bg-gray-200 font-bold text-[10px] uppercase px-5 py-2.5 rounded-xl"
                  >
                    Close
                  </button>
                  <button
                    type="submit"
                    className="btn bg-[#1A365D] hover:bg-black text-white font-black text-[10px] uppercase px-6 py-2.5 rounded-xl shadow-md"
                  >
                    {h ? 'Save Log Changes' : 'Confirm Stay Verification & Move to Checked'}
                  </button>
                </div>
              </form>
            </motion.div>
          </td>
        </tr>
      )}
    </>
  );
};

/* --- CHECKING PANEL COMPONENT --- */
interface CheckingOptionProps {
  records: ImmRecord[];
  tempRecords: ImmRecord[];
  setRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  setTempRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  showToast: (msg: string) => void;
  movementMap: Record<string, MovementData>;
  currentUser?: { name: string; title: string } | null;
  cloudAuthUser?: CloudAuthUser | null;
  isViewer?: boolean;
  passportToLatestInfo: Record<string, { fullname: string; nationality: string; gender: string }>;
  setCheckingPrintData: React.Dispatch<React.SetStateAction<{
    title: string;
    type: 'CO_LTD_PENDING' | 'OTHERS_PENDING' | 'CHECKED_HISTORY';
    data: any[];
  } | null>>;
  masterData: MasterItem[];
  syncMaster: (value: string, type: MasterItem['type'], linkedValue?: string) => void;
  setActivePrintPreview: React.Dispatch<React.SetStateAction<any>>;
  checkingHistory: CheckingHistoryEntry[];
  setCheckingHistory: React.Dispatch<React.SetStateAction<CheckingHistoryEntry[]>>;
}

const CheckingOption = ({
  records,
  tempRecords,
  setRecords,
  setTempRecords,
  showToast,
  movementMap,
  currentUser,
  cloudAuthUser,
  isViewer,
  passportToLatestInfo,
  setCheckingPrintData,
  masterData,
  syncMaster,
  setActivePrintPreview,
  checkingHistory,
  setCheckingHistory
}: CheckingOptionProps) => {
  const [searchQuery, setSearchQuery] = useState('');

  const isCoLtdAddress = (addressStr: string) => {
    const clean = (addressStr || '').toUpperCase();
    return clean.includes('CO') && (clean.includes('LTD') || clean.includes('L.T.D'));
  };

  // Derive checklists from active stay-in visitors (movementMap)
  const { coLtdToCheckList, othersToCheckList } = useMemo(() => {
    const checkedPassports = new Set(checkingHistory.map(h => h.passport.toUpperCase()));
    const coLtd: any[] = [];
    const others: any[] = [];
    (Object.values(movementMap) as MovementData[]).filter(m => !m.out).forEach(m => {
      if (checkedPassports.has(m.p.toUpperCase())) return;
      const isCo = isCoLtdAddress(m.loc);
      if (isCo) {
        coLtd.push(m);
      } else {
        others.push(m);
      }
    });
    return { coLtdToCheckList: coLtd, othersToCheckList: others };
  }, [movementMap, checkingHistory]);

  const coLtdCheckedList = useMemo(() => {
    return checkingHistory.filter(h => h.type === 'CO_LTD');
  }, [checkingHistory]);

  const othersCheckedList = useMemo(() => {
    return checkingHistory.filter(h => h.type !== 'CO_LTD');
  }, [checkingHistory]);

  // Helper helper to update records & tempRecords
  const updateStatusInRecords = (passport: string, status: string, permittedBy: string, confirmedAddress: string, confirmedStayDescription: string = '') => {
    const applyUpdate = (r: ImmRecord) => {
      const mappedStatus = status === 'STILL PERMITTED' ? 'STAY PERMITTED' : (status === 'STILL NOT PERMITTED YET' ? 'STAY NOT PERMITTED' : status);
      return {
        ...r,
        stillPermittedStatus: mappedStatus,
        permittedBy: permittedBy,
        address: confirmedAddress ? confirmedAddress : r.address,
        stayDescription: confirmedStayDescription !== undefined ? confirmedStayDescription : r.stayDescription
      };
    };

    setRecords((prev: ImmRecord[]) => {
      return prev.map(r => r.passport.toUpperCase() === passport.toUpperCase() ? applyUpdate(r) : r);
    });

    setTempRecords((prev: ImmRecord[]) => {
      return prev.map(r => r.passport.toUpperCase() === passport.toUpperCase() ? applyUpdate(r) : r);
    });
  };

  // Submits a verified check and persists it
  const handleCheckSubmitted = (newCheck: CheckingHistoryEntry) => {
    setCheckingHistory(prev => [newCheck, ...prev]);
    updateStatusInRecords(
      newCheck.passport, 
      newCheck.status, 
      newCheck.permittedBy || '', 
      newCheck.confirmedAddress, 
      newCheck.confirmedStayDescription || ''
    );
    logActivity({
      action: 'CREATE',
      module: 'CHECKING',
      officerName: newCheck.officerName || cloudAuthUser?.username,
      officerRole: cloudAuthUser?.role || 'Editor',
      targetId: newCheck.passport,
      details: `Stay Verification logged for Passport ${newCheck.passport} (${newCheck.fullname || 'N/A'}, ${newCheck.nationality || 'N/A'}): ${newCheck.status} by ${newCheck.permittedBy || '-'}`
    });
    showToast(`VERIFIED CHECK LOGGED SUCCESS FOR ${newCheck.passport}`);
  };

  // Updates an edited verified check
  const handleCheckUpdated = (updatedCheck: CheckingHistoryEntry) => {
    setCheckingHistory(prev => prev.map(c => c.id === updatedCheck.id ? updatedCheck : c));
    updateStatusInRecords(
      updatedCheck.passport,
      updatedCheck.status,
      updatedCheck.permittedBy || '',
      updatedCheck.confirmedAddress,
      updatedCheck.confirmedStayDescription || ''
    );
    logActivity({
      action: 'UPDATE',
      module: 'CHECKING',
      officerName: updatedCheck.officerName || cloudAuthUser?.username,
      officerRole: cloudAuthUser?.role || 'Editor',
      targetId: updatedCheck.passport,
      details: `Stay Verification updated for Passport ${updatedCheck.passport} (${updatedCheck.fullname || 'N/A'}): ${updatedCheck.status} by ${updatedCheck.permittedBy || '-'}`
    });
    showToast(`UPDATED VERIFIED STAY LOG FOR ${updatedCheck.passport}`);
  };

  // Reverts check log, returns candidate back to check queue
  const removeCheckLog = (id: string, passport: string) => {
    setCheckingHistory(prev => prev.filter(c => c.id !== id));
    updateStatusInRecords(passport, '', '', '');
    logActivity({
      action: 'DELETE',
      module: 'CHECKING',
      officerName: cloudAuthUser?.username,
      officerRole: cloudAuthUser?.role || 'Editor',
      targetId: passport,
      details: `Reverted verification check log for Passport ${passport} back to pending queue`
    });
    showToast("CHECK LOG REVERTED AND VISITOR RE-QUEUED FOR CHECKING");
  };

  // Consolidate PENDING (to check) and CHECKED list into a single unified directory array
  const masterCheckingList = useMemo(() => {
    const list: any[] = [];
    
    // 1. Unchecked Corporate
    coLtdToCheckList.forEach(m => {
      const info = passportToLatestInfo[m.p.toUpperCase()] || { fullname: m.n, nationality: m.nat };
      list.push({
        type: 'PENDING',
        category: 'CO_LTD',
        isCo: true,
        passport: m.p,
        fullname: info.fullname,
        nationality: info.nationality,
        originalAddress: m.loc,
        confirmedAddress: m.loc,
        checkDate: '',
        status: '',
        permittedBy: '',
        rawMovement: m,
      });
    });

    // 2. Unchecked Standard Others
    othersToCheckList.forEach(m => {
      const info = passportToLatestInfo[m.p.toUpperCase()] || { fullname: m.n, nationality: m.nat };
      list.push({
        type: 'PENDING',
        category: 'OTHERS',
        isCo: false,
        passport: m.p,
        fullname: info.fullname,
        nationality: info.nationality,
        originalAddress: m.loc,
        confirmedAddress: m.loc,
        checkDate: '',
        status: '',
        permittedBy: '',
        rawMovement: m,
      });
    });

    // 3. Checked Corporate (CO LTD) history
    coLtdCheckedList.forEach(h => {
      list.push({
        type: 'CHECKED',
        category: 'CO_LTD',
        isCo: true,
        passport: h.passport,
        fullname: h.fullname,
        nationality: h.nationality,
        originalAddress: h.originalAddress || h.confirmedAddress,
        confirmedAddress: h.confirmedAddress,
        confirmedStayDescription: h.confirmedStayDescription || '',
        checkDate: h.checkDate,
        status: h.status,
        permittedBy: h.permittedBy || '',
        officerName: h.officerName,
        officerTitle: h.officerTitle,
        rawHistory: h,
      });
    });

    // 4. Checked Standard Others history
    othersCheckedList.forEach(h => {
      list.push({
        type: 'CHECKED',
        category: 'OTHERS',
        isCo: false,
        passport: h.passport,
        fullname: h.fullname,
        nationality: h.nationality,
        originalAddress: h.originalAddress || h.confirmedAddress,
        confirmedAddress: h.confirmedAddress,
        confirmedStayDescription: h.confirmedStayDescription || '',
        checkDate: h.checkDate,
        status: h.status,
        permittedBy: h.permittedBy || '',
        officerName: h.officerName,
        officerTitle: h.officerTitle,
        rawHistory: h,
      });
    });

    return list;
  }, [coLtdToCheckList, othersToCheckList, coLtdCheckedList, othersCheckedList, passportToLatestInfo]);

  // Unified directory master smart search box
  const filteredMasterCheckingList = useMemo(() => {
    if (!searchQuery.trim()) return masterCheckingList;
    const q = searchQuery.toLowerCase();
    return masterCheckingList.filter(item => {
      const statusText = item.status ? item.status.toLowerCase() : 'awaiting pending';
      const categoryText = item.category === 'CO_LTD' ? 'corporate co ltd hotel company' : 'standard others general guest';
      const stateText = item.type === 'PENDING' ? 'pending unchecked queue' : 'checked verified completed';
      const permittedByText = item.permittedBy ? item.permittedBy.toLowerCase() : '';
      const officerText = item.officerName ? (item.officerName.toLowerCase() + ' ' + (item.officerTitle || '').toLowerCase()) : '';

      return (
        item.passport.toLowerCase().includes(q) ||
        item.fullname.toLowerCase().includes(q) ||
        item.nationality.toLowerCase().includes(q) ||
        item.originalAddress.toLowerCase().includes(q) ||
        item.confirmedAddress.toLowerCase().includes(q) ||
        (item.confirmedStayDescription && item.confirmedStayDescription.toLowerCase().includes(q)) ||
        statusText.includes(q) ||
        categoryText.includes(q) ||
        stateText.includes(q) ||
        permittedByText.includes(q) ||
        officerText.includes(q)
      );
    });
  }, [masterCheckingList, searchQuery]);

  // Performs Excel generation over multiple sheets for maximum data completeness (as per user demand)
  const exportCheckingToExcel = () => {
    try {
      const checkedSheetData = checkingHistory.map((h, i) => {
        const info = passportToLatestInfo[h.passport.toUpperCase()] || { fullname: h.fullname, nationality: h.nationality };
        const m = (movementMap[h.passport.toUpperCase()] || {}) as Partial<MovementData>;
        return {
          "No": i + 1,
          "Passport": h.passport,
          "Full Name": info.fullname,
          "Nationality": info.nationality,
          "Visa Type": m.visa || '',
          "Arrival Date": m.in || '',
          "Expected Stay From": m.start || '',
          "Expected Stay To": m.end || '',
          "Allowed Days": m.allowed || '',
          "Carrier Vehicle": m.vInfo || '',
          "Agent / Brought By": m.agent || '',
          "Contact Details": m.contact || '',
          "Category": h.type === 'CO_LTD' ? 'CORPORATE (CO LTD)' : 'OTHERS',
          "Verification Status": h.status === 'STILL PERMITTED' || h.status === 'STAY PERMITTED' ? 'STAY PERMITTED' : (h.status === 'STILL NOT PERMITTED YET' || h.status === 'STAY NOT PERMITTED' ? 'STAY NOT PERMITTED' : h.status),
          "Permit Description": h.permittedBy || '',
          "Confirmed Address": h.confirmedAddress,
          "Confirmed Stay Detail": h.confirmedStayDescription || '',
          "Checked Date": h.checkDate,
          "Officer Title": h.officerTitle,
          "Officer Name": h.officerName,
          "Log Timestamp": new Date(h.timestamp).toLocaleString()
        };
      });

      const corporateSheetData = coLtdToCheckList.map((m, i) => {
        const info = passportToLatestInfo[m.p.toUpperCase()] || { fullname: m.n, nationality: m.nat };
        return {
          "No": i + 1,
          "Passport": m.p,
          "Full Name": info.fullname,
          "Nationality": info.nationality,
          "Visa Type": m.visa || '',
          "Arrival Date": m.in || '',
          "Active Stay Address": m.loc,
          "Allowed Stay Days": m.allowed,
          "Days Remaining": getLiveRemainingDays(m.end),
          "Start Date": m.start,
          "End Date": m.end,
          "Carrier Vehicle": m.vInfo || '',
          "Agent / Brought By": m.agent,
          "Contact Details": m.contact
        };
      });

      const othersSheetData = othersToCheckList.map((m, i) => {
        const info = passportToLatestInfo[m.p.toUpperCase()] || { fullname: m.n, nationality: m.nat };
        return {
          "No": i + 1,
          "Passport": m.p,
          "Full Name": info.fullname,
          "Nationality": info.nationality,
          "Visa Type": m.visa || '',
          "Arrival Date": m.in || '',
          "Active Stay Address": m.loc,
          "Allowed Stay Days": m.allowed,
          "Days Remaining": getLiveRemainingDays(m.end),
          "Start Date": m.start,
          "End Date": m.end,
          "Carrier Vehicle": m.vInfo || '',
          "Agent / Brought By": m.agent,
          "Contact Details": m.contact
        };
      });

      const wb = XLSX.utils.book_new();
      
      const ws1 = XLSX.utils.json_to_sheet(checkedSheetData);
      XLSX.utils.book_append_sheet(wb, ws1, "Checked History Logs");

      const ws2 = XLSX.utils.json_to_sheet(corporateSheetData);
      XLSX.utils.book_append_sheet(wb, ws2, "Corporate Checklist");

      const ws3 = XLSX.utils.json_to_sheet(othersSheetData);
      XLSX.utils.book_append_sheet(wb, ws3, "Others Checklist");

      XLSX.writeFile(wb, `Checking_Verification_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
      showToast("EXCEL VERIFICATION REPORT EXPORTED SUCCESSFULLY FOR ALL SHEETS!");
    } catch (e) {
      console.error(e);
      showToast("FAILED TO EXPORT EXCEL REPORT");
    }
  };

  const handlePrintReport = () => {
    setActivePrintPreview({
      title: 'Immigration Stay Physical Verification Directory & Checked Ledger',
      type: 'CHECKED_HISTORY',
      data: filteredMasterCheckingList
    });
    showToast("OPENING INTERACTIVE PRINT PREVIEW WITH MASTER FILTERED ENTRIES...");
  };

  return (
    <div className="max-w-[1600px] mx-auto py-4 sm:py-8 px-2.5 sm:px-4 space-y-4 sm:space-y-6 animate-in fade-in duration-300">
      {/* Dynamic Dashboard metrics header inside checking tab */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5 sm:gap-4">
        <div className="bg-gradient-to-br from-indigo-50 to-indigo-100 border border-indigo-200/50 p-3 sm:p-4.5 rounded-2xl flex items-center justify-between">
          <div className="space-y-0.5 sm:space-y-1">
            <span className="text-[9px] sm:text-[10px] text-indigo-950 font-black uppercase tracking-wider">Total Ledger Size</span>
            <div className="text-xl sm:text-2xl font-black text-indigo-950 font-mono">{masterCheckingList.length}</div>
          </div>
          <span className="text-lg sm:text-xl">📋</span>
        </div>
        <div className="bg-gradient-to-br from-amber-50 to-amber-100 border border-amber-200/50 p-3 sm:p-4.5 rounded-2xl flex items-center justify-between">
          <div className="space-y-0.5 sm:space-y-1">
            <span className="text-[9px] sm:text-[10px] text-amber-950 font-black uppercase tracking-wider">Pending Checks</span>
            <div className="text-xl sm:text-2xl font-black text-amber-950 font-mono">
              {coLtdToCheckList.length + othersToCheckList.length}
            </div>
          </div>
          <span className="text-lg sm:text-xl animate-pulse">⏳</span>
        </div>
        <div className="bg-gradient-to-br from-emerald-50 to-emerald-100 border border-emerald-200/50 p-3 sm:p-4.5 rounded-2xl flex items-center justify-between">
          <div className="space-y-0.5 sm:space-y-1">
            <span className="text-[9px] sm:text-[10px] text-emerald-950 font-black uppercase tracking-wider">Corporate Verified</span>
            <div className="text-xl sm:text-2xl font-black text-emerald-950 font-mono">{coLtdCheckedList.length}</div>
          </div>
          <span className="text-lg sm:text-xl">🏢</span>
        </div>
        <div className="bg-gradient-to-br from-teal-50 to-teal-100 border border-teal-200/50 p-3 sm:p-4.5 rounded-2xl flex items-center justify-between">
          <div className="space-y-0.5 sm:space-y-1">
            <span className="text-[9px] sm:text-[10px] text-teal-950 font-black uppercase tracking-wider">Standard Verified</span>
            <div className="text-xl sm:text-2xl font-black text-teal-950 font-mono">{othersCheckedList.length}</div>
          </div>
          <span className="text-lg sm:text-xl">🏡</span>
        </div>
      </div>

      <div className="card shadow-md border-t-4 sm:border-t-8 border-indigo-700 bg-white p-3.5 sm:p-6 no-print">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-3 sm:gap-6">
          <div className="flex items-center gap-3">
             <div className="bg-indigo-100 p-2 sm:p-3 rounded-full text-indigo-700">
                <CheckSquare size={22} className="sm:w-[26px] sm:h-[26px]" />
             </div>
             <div>
                <h2 className="text-base sm:text-xl font-black text-gray-800 uppercase tracking-tight">Immigration Stay Checking Ledger</h2>
                <p className="text-gray-400 text-[10px] sm:text-xs font-bold font-mono uppercase">Liaison and Physical verification operations board</p>
             </div>
          </div>
          
          <div className="flex items-center gap-2">
            <span className="text-[10px] sm:text-[11px] font-black text-emerald-800 bg-emerald-50 border border-emerald-200 uppercase px-3 py-1.5 sm:px-3.5 sm:py-2 rounded-xl">
              ● Unified Directory Active
            </span>
          </div>
        </div>
      </div>

      {/* Main Checking Content list */}
      <div className="card shadow-md border border-gray-100 bg-white p-3.5 sm:p-6">
         <div className="flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4 mb-4 sm:mb-6">
           <div>
              <h3 className="text-sm sm:text-md font-black text-slate-900 uppercase">
                📋 Unified Physical Stay Checklist & Verification Log History
              </h3>
              <p className="text-[11px] sm:text-xs text-slate-500 font-semibold mt-0.5">
                Unified tracking table for Corporate (CO LTD) and Standard (OTHERS) stay lists. Fully searchable and editable logs.
              </p>
           </div>
           
           <div className="flex flex-col sm:flex-row flex-wrap items-center gap-2 w-full xl:w-auto">
             <div className="relative w-full xl:w-80">
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Search Passport, Name, Nationality, Status, Address..."
                  className="input-field pl-8.5 pr-4 text-xs bg-slate-50 border-slate-200 focus:bg-white rounded-xl font-medium py-2.5 sm:py-3"
                />
                <Search size={14} className="absolute left-3 top-3.5 sm:top-4 text-slate-400" />
             </div>

             <div className="flex items-center gap-2 w-full sm:w-auto">
               <button
                  onClick={exportCheckingToExcel}
                  className="btn flex-1 sm:flex-none bg-emerald-600 text-white hover:bg-black text-[10px] font-black uppercase flex items-center justify-center gap-1.5 px-4 py-2.5 sm:py-3 rounded-xl shadow-xs transition-colors"
                  title="Export full checking lists and history into a detailed multi-sheet workbook"
               >
                  <Download size={14} /> EXPORT EXCEL
               </button>

               <button
                  onClick={handlePrintReport}
                  className="btn flex-1 sm:flex-none bg-slate-800 text-white hover:bg-black text-[10px] font-black uppercase flex items-center justify-center gap-1.5 px-4 py-2.5 sm:py-3 rounded-xl shadow-sm transition-colors"
                  title="Print the master verification checklist ledger"
               >
                  <Printer size={14} /> PRINT MASTER
               </button>
             </div>
           </div>
         </div>

         {/* Mobile Card Layout (< md screens) */}
         <div className="block md:hidden space-y-3 max-h-[600px] overflow-y-auto pr-0.5">
           {filteredMasterCheckingList.length === 0 ? (
             <div className="p-8 text-center text-slate-400 font-black uppercase tracking-wider text-xs bg-slate-50 rounded-2xl border border-dashed border-slate-200">
               ⚠️ No checking directory records match current search criteria
             </div>
           ) : (
             filteredMasterCheckingList.map((item, idx) => (
               <CheckingCard
                 key={'m_card_' + item.passport + '_' + item.type + '_' + idx}
                 m={item.type === 'PENDING' ? item.rawMovement : undefined}
                 h={item.type === 'CHECKED' ? item.rawHistory : undefined}
                 isCo={item.isCo}
                 currentUser={currentUser}
                 onCheckSubmitted={handleCheckSubmitted}
                 onCheckUpdated={handleCheckUpdated}
                 onCheckRemoved={removeCheckLog}
                 records={records}
                 syncMaster={syncMaster}
                 masterData={masterData}
               />
             ))
           )}
         </div>

         {/* Desktop Table Layout (>= md screens) */}
         <div className="hidden md:block overflow-x-auto border border-slate-200 rounded-xl max-h-[650px] overflow-y-auto shadow-xs">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-slate-50 sticky top-0 border-b border-slate-200 z-10 font-bold">
                <tr className="uppercase text-[9px] text-slate-500 font-black">
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Verification Status / Class</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Foreigner Identifier</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Nationality</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Stay Address / Confirmed Stay</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Verification Permit Status</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200">Log / Duty Countdown</th>
                  <th className="p-3.5 sm:p-4 border-b border-slate-200 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {filteredMasterCheckingList.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="p-16 text-center text-slate-400 font-black uppercase tracking-wider leading-relaxed">
                      ⚠️ No checking directory records match current search criteria
                    </td>
                  </tr>
                ) : (
                  filteredMasterCheckingList.map((item, idx) => (
                    <CheckingRow 
                      key={item.passport + '_' + item.type + '_' + idx}
                      m={item.type === 'PENDING' ? item.rawMovement : undefined}
                      h={item.type === 'CHECKED' ? item.rawHistory : undefined}
                      isCo={item.isCo}
                      currentUser={currentUser}
                  cloudAuthUser={cloudAuthUser}
                  isViewer={isViewer}
                  showToast={showToast}
                      onCheckSubmitted={handleCheckSubmitted}
                      onCheckUpdated={handleCheckUpdated}
                      onCheckRemoved={removeCheckLog}
                      records={records}
                      syncMaster={syncMaster}
                      masterData={masterData}
                    />
                  ))
                )}
              </tbody>
            </table>
         </div>
      </div>
    </div>
  );
};

/* --- LEGACY SECTION REMOVED FOR COMPILATION --- */

const MasterDB = ({ 
  masterData, 
  setMasterData, 
  showToast,
  backupCombinedJSON,
  restoreCombinedJSON,
  setSearchQuery,
  setActiveTab,
  setRecords,
  setTempRecords,
  handleManualSync,
  handleFetchDataFromCloud,
  records,
  tempRecords
}: { 
  masterData: MasterItem[]; 
  setMasterData: React.Dispatch<React.SetStateAction<MasterItem[]>>; 
  showToast: (msg: string) => void;
  backupCombinedJSON: () => void;
  restoreCombinedJSON: (e: React.ChangeEvent<HTMLInputElement>) => void;
  setSearchQuery: (q: string) => void;
  setActiveTab: (t: string) => void;
  setRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  setTempRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  handleManualSync: () => void;
  handleFetchDataFromCloud: () => void;
  records?: ImmRecord[];
  tempRecords?: ImmRecord[];
}) => {
  const [mName, setMName] = useState('');
  const [mType, setMType] = useState<MasterItem['type']>('Nationality');
  const [mSearch, setMSearch] = useState('');
  const [mLinked, setMLinked] = useState('');
  const [confirmMasterDeleteId, setConfirmMasterDeleteId] = useState<number | null>(null);
  const [isMasterSyncing, setIsMasterSyncing] = useState(false);

  const handleMasterDataUnifiedSync = async () => {
    if (isMasterSyncing) return;
    setIsMasterSyncing(true);
    try {
      showToast(" Cloud နှင့် Master Data Sync ပြုလုပ်နေပါသည် (Upload / Download)...");
      // First, sync any local changes to cloud
      if (typeof handleManualSync === 'function') {
        await handleManualSync();
      }
      // Second, retrieve latest updates from cloud
      if (typeof handleFetchDataFromCloud === 'function') {
        await handleFetchDataFromCloud();
      }
      showToast(" Master Data Cloud Synchronization ပြီးစီးပါပြီ");
    } catch (e: any) {
      console.error("Unified Master Data Sync Error:", e);
      showToast("⚠️ Master Data Sync လုပ်ဆောင်ရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
    } finally {
      setIsMasterSyncing(false);
    }
  };

  const addMaster = () => {
    if (!mName.trim()) return;
    const cleanName = mName.trim();
    const cleanLinked = mLinked.trim();

    if (masterData.some(m => m.name.toLowerCase() === cleanName.toLowerCase() && m.type === mType)) {
      showToast("Already exists in database"); return;
    }

    const newItem: MasterItem = { 
      id: Date.now(), 
      name: cleanName, 
      type: mType,
      linkedValue: cleanLinked || undefined,
      updatedAt: new Date().toISOString()
    };

    setMasterData(prev => {
      const updated = [...prev, newItem];
      
      // Also register the secondary details in its own category for comprehensive lookup support
      if (cleanLinked) {
        if (mType === 'Agent') {
          if (!prev.some(m => m.name.toLowerCase() === cleanLinked.toLowerCase() && m.type === 'Contact')) {
            updated.push({ id: Date.now() + 1, name: cleanLinked, type: 'Contact', updatedAt: new Date().toISOString() });
          }
        } else if (mType === 'Official') {
          if (!prev.some(m => m.name.toLowerCase() === cleanLinked.toLowerCase() && m.type === 'Title')) {
            updated.push({ id: Date.now() + 2, name: cleanLinked, type: 'Title', updatedAt: new Date().toISOString() });
          }
        } else if (mType === 'Reporter') {
          if (!prev.some(m => m.name.toLowerCase() === cleanLinked.toLowerCase() && m.type === 'Phone')) {
            updated.push({ id: Date.now() + 3, name: cleanLinked, type: 'Phone', updatedAt: new Date().toISOString() });
          }
        }
      }
      return updated;
    });

    setMName('');
    setMLinked('');
    logActivity({
      action: 'CREATE',
      module: 'MASTER',
      targetId: cleanName,
      details: `Added Master Reference [${mType}]: "${cleanName}"${cleanLinked ? ` (Linked: "${cleanLinked}")` : ''}`
    });
    showToast("MASTER DATA ADDED SUCCESSFULLY");
  };

  const deleteMaster = (id: number) => {
    const targetM = masterData.find(m => m.id === id);
    setMasterData(prev => prev.filter(m => m.id !== id));
    if (targetM) {
      logActivity({
        action: 'DELETE',
        module: 'MASTER',
        targetId: targetM.name,
        details: `Deleted Master Reference [${targetM.type}]: "${targetM.name}"`
      });
    }
    showToast("DELETED");
    setConfirmMasterDeleteId(null);
  };

  const [editingId, setEditingId] = useState<number | null>(null);
  const [editName, setEditName] = useState('');
  const [editLinked, setEditLinked] = useState('');

  const startEdit = (m: MasterItem) => {
    setEditingId(m.id);
    setEditName(m.name);
    setEditLinked(m.linkedValue || '');
  };

  const saveEdit = () => {
    if (!editName.trim()) return;
    const newName = editName.trim();
    const newLinked = editLinked.trim();

    const originalItem = masterData.find(m => m.id === editingId);
    if (originalItem) {
      const oldName = originalItem.name;
      const oldLinked = originalItem.linkedValue || '';

      const updateRecordFn = (r: ImmRecord): ImmRecord => {
        const updated = { ...r };
        let changed = false;

        if (originalItem.type === 'Nationality') {
          if (updated.nationality === oldName) { updated.nationality = newName; changed = true; }
        } else if (originalItem.type === 'Visa') {
          if (updated.visaType === oldName) { updated.visaType = newName; changed = true; }
        } else if (originalItem.type === 'Stay') {
          if (updated.address === oldName) { updated.address = newName; changed = true; }
          if (updated.formC && updated.formC.address === oldName) {
            updated.formC = { ...updated.formC, address: newName };
            changed = true;
          }
          if (oldLinked && updated.address === oldLinked) { updated.address = newLinked; changed = true; }
          if (oldLinked && updated.formC && updated.formC.address === oldLinked) {
            updated.formC = { ...updated.formC, address: newLinked };
            changed = true;
          }
        } else if (originalItem.type === 'Vehicle') {
          if (updated.vehicleInfo === oldName) { updated.vehicleInfo = newName; changed = true; }
        } else if (originalItem.type === 'Agent') {
          if (updated.broughtBy === oldName) { updated.broughtBy = newName; changed = true; }
        } else if (originalItem.type === 'Contact') {
          if (updated.contactDetails === oldName) { updated.contactDetails = newName; changed = true; }
        } else if (originalItem.type === 'Official') {
          if (updated.officialName === oldName) { updated.officialName = newName; changed = true; }
          if (updated.formC && updated.formC.officialName === oldName) {
            updated.formC = { ...updated.formC, officialName: newName };
            changed = true;
          }
        } else if (originalItem.type === 'Title') {
          if (updated.officialTitle === oldName) { updated.officialTitle = newName; changed = true; }
          if (updated.formC && updated.formC.officialTitle === oldName) {
            updated.formC = { ...updated.formC, officialTitle: newName };
            changed = true;
          }
        } else if (originalItem.type === 'Reporter') {
          if (updated.formC && updated.formC.reporterName === oldName) {
            updated.formC = { ...updated.formC, reporterName: newName };
            changed = true;
          }
        } else if (originalItem.type === 'Phone') {
          if (updated.formC && updated.formC.reporterPhone === oldName) {
            updated.formC = { ...updated.formC, reporterPhone: newName };
            changed = true;
          }
        }

        return updated;
      };

      setRecords(prev => prev.map(updateRecordFn));
      setTempRecords(prev => prev.map(updateRecordFn));

      logActivity({
        action: 'UPDATE',
        module: 'MASTER',
        targetId: newName,
        details: `Updated Master Reference [${originalItem.type}]: "${oldName}" -> "${newName}"`
      });
    }

    setMasterData(prev => prev.map(m => m.id === editingId ? { 
      ...m, 
      name: newName, 
      linkedValue: newLinked || undefined,
      updatedAt: new Date().toISOString()
    } : m));
    setEditingId(null);
    showToast("UPDATED ALL ENTRIES AUTOMATICALLY");
  };

  const [showMergeModal, setShowMergeModal] = useState(false);
  const [selectedTargetIds, setSelectedTargetIds] = useState<Record<string, number>>({});

  // Calculate duplicate clusters in masterData (sorted by newest/updatedAt first)
  const duplicateClusters = useMemo(() => {
    const typeMap: Record<string, MasterItem[]> = {};
    masterData.forEach(item => {
      if (!typeMap[item.type]) typeMap[item.type] = [];
      typeMap[item.type].push(item);
    });

    const clusters: {
      type: MasterItem['type'];
      key: string;
      items: MasterItem[];
    }[] = [];

    const norm = (s: string) => s.toLowerCase().trim().replace(/\s+/g, ' ');

    Object.entries(typeMap).forEach(([type, items]) => {
      const groups: Record<string, MasterItem[]> = {};
      items.forEach(item => {
        const k = norm(item.name);
        if (!groups[k]) groups[k] = [];
        groups[k].push(item);
      });

      Object.entries(groups).forEach(([k, groupItems]) => {
        if (groupItems.length > 1) {
          // Sort items in group so the most recently updated/added item is FIRST (default primary)
          const sorted = [...groupItems].sort((a, b) => {
            const timeA = a.updatedAt ? Date.parse(a.updatedAt) : a.id;
            const timeB = b.updatedAt ? Date.parse(b.updatedAt) : b.id;
            return (timeB || 0) - (timeA || 0);
          });
          clusters.push({
            type: type as MasterItem['type'],
            key: `${type}_${k}`,
            items: sorted
          });
        }
      });
    });

    return clusters;
  }, [masterData]);

  const handleMergeCluster = (clusterKey: string, primaryId: number) => {
    const cluster = duplicateClusters.find(c => c.key === clusterKey);
    if (!cluster) return;

    const primaryItem = cluster.items.find(i => i.id === primaryId) || cluster.items[0];
    const targetName = primaryItem.name.trim();
    const nonPrimaryItems = cluster.items.filter(i => i.id !== primaryId);
    const oldNames = nonPrimaryItems.map(i => i.name);
    const idsToRemove = new Set(nonPrimaryItems.map(i => i.id));

    // Determine latest linkedValue across cluster items
    const latestLinkedFromCluster = cluster.items.find(i => i.linkedValue && i.linkedValue.trim())?.linkedValue?.trim();
    const targetLinked = primaryItem.linkedValue?.trim() || latestLinkedFromCluster || undefined;

    const updateRecordFn = (r: ImmRecord): ImmRecord => {
      let updated = { ...r };

      oldNames.forEach(oldName => {
        if (cluster.type === 'Nationality' && updated.nationality === oldName) {
          updated.nationality = targetName;
        } else if (cluster.type === 'Visa' && updated.visaType === oldName) {
          updated.visaType = targetName;
        } else if (cluster.type === 'Stay') {
          if (updated.address === oldName) {
            updated.address = targetName;
            if (targetLinked && (!updated.stayDescription || updated.stayDescription === oldName)) {
              updated.stayDescription = targetLinked;
            }
          }
          if (updated.formC && updated.formC.address === oldName) {
            updated.formC = { 
              ...updated.formC, 
              address: targetName,
              stayDescription: targetLinked || updated.formC.stayDescription
            };
          }
        } else if (cluster.type === 'Vehicle' && updated.vehicleInfo === oldName) {
          updated.vehicleInfo = targetName;
        } else if (cluster.type === 'Agent' && updated.broughtBy === oldName) {
          updated.broughtBy = targetName;
          if (targetLinked && !updated.contactDetails) updated.contactDetails = targetLinked;
        } else if (cluster.type === 'Contact' && updated.contactDetails === oldName) {
          updated.contactDetails = targetName;
        } else if (cluster.type === 'Official') {
          if (updated.officialName === oldName) {
            updated.officialName = targetName;
            if (targetLinked && !updated.officialTitle) updated.officialTitle = targetLinked;
          }
          if (updated.formC && updated.formC.officialName === oldName) {
            updated.formC = { 
              ...updated.formC, 
              officialName: targetName,
              officialTitle: targetLinked || updated.formC.officialTitle
            };
          }
        } else if (cluster.type === 'Title') {
          if (updated.officialTitle === oldName) updated.officialTitle = targetName;
          if (updated.formC && updated.formC.officialTitle === oldName) {
            updated.formC = { ...updated.formC, officialTitle: targetName };
          }
        } else if (cluster.type === 'Reporter') {
          if (updated.formC && updated.formC.reporterName === oldName) {
            updated.formC = { 
              ...updated.formC, 
              reporterName: targetName,
              reporterPhone: targetLinked || updated.formC.reporterPhone
            };
          }
        } else if (cluster.type === 'Phone') {
          if (updated.formC && updated.formC.reporterPhone === oldName) {
            updated.formC = { ...updated.formC, reporterPhone: targetName };
          }
        }
      });

      return updated;
    };

    if (setRecords) {
      setRecords(prev => {
        const updatedRecs = prev.map(updateRecordFn);
        localStorage.setItem('imm_records_react', JSON.stringify(updatedRecs));
        miniDB.set('records', updatedRecs).catch(() => {});
        saveCollectionToFirestore('records', updatedRecs).catch(() => {});
        return updatedRecs;
      });
    }
    if (setTempRecords) {
      setTempRecords(prev => {
        const updatedTemps = prev.map(updateRecordFn);
        localStorage.setItem('imm_temp_records_react', JSON.stringify(updatedTemps));
        miniDB.set('tempRecords', updatedTemps).catch(() => {});
        saveCollectionToFirestore('tempRecords', updatedTemps).catch(() => {});
        return updatedTemps;
      });
    }

    setMasterData(prev => {
      const mergedList = prev
        .filter(m => !idsToRemove.has(m.id))
        .map(m => m.id === primaryItem.id ? { 
          ...m, 
          name: targetName, 
          linkedValue: targetLinked,
          updatedAt: new Date().toISOString()
        } : m);
      localStorage.setItem('imm_master_react', JSON.stringify(mergedList));
      miniDB.set('masterData', mergedList).catch(() => {});
      saveCollectionToFirestore('masterData', mergedList, true).catch(() => {});
      return mergedList;
    });

    showToast(`MERGED DUPLICATES INTO "${targetName}" SUCCESSFULLY (SYNCED TO CLOUD)`);
  };

  const handleMergeAllClusters = () => {
    if (duplicateClusters.length === 0) return;

    const allIdsToRemove = new Set<number>();
    const clusterUpdates: {
      type: MasterItem['type'];
      oldNames: string[];
      targetName: string;
      targetLinked?: string;
      primaryId: number;
    }[] = [];

    duplicateClusters.forEach(cluster => {
      const primaryId = selectedTargetIds[cluster.key] || cluster.items[0].id;
      const primaryItem = cluster.items.find(i => i.id === primaryId) || cluster.items[0];
      const targetName = primaryItem.name.trim();
      const nonPrimaryItems = cluster.items.filter(i => i.id !== primaryId);
      const oldNames = nonPrimaryItems.map(i => i.name);
      nonPrimaryItems.forEach(i => allIdsToRemove.add(i.id));

      const latestLinked = cluster.items.find(i => i.linkedValue && i.linkedValue.trim())?.linkedValue?.trim();
      const targetLinked = primaryItem.linkedValue?.trim() || latestLinked || undefined;

      clusterUpdates.push({
        type: cluster.type,
        oldNames,
        targetName,
        targetLinked,
        primaryId: primaryItem.id
      });
    });

    const updateRecordBatchFn = (r: ImmRecord): ImmRecord => {
      let updated = { ...r };
      clusterUpdates.forEach(cu => {
        cu.oldNames.forEach(oldName => {
          if (cu.type === 'Nationality' && updated.nationality === oldName) {
            updated.nationality = cu.targetName;
          } else if (cu.type === 'Visa' && updated.visaType === oldName) {
            updated.visaType = cu.targetName;
          } else if (cu.type === 'Stay') {
            if (updated.address === oldName) {
              updated.address = cu.targetName;
              if (cu.targetLinked && (!updated.stayDescription || updated.stayDescription === oldName)) {
                updated.stayDescription = cu.targetLinked;
              }
            }
            if (updated.formC && updated.formC.address === oldName) {
              updated.formC = { 
                ...updated.formC, 
                address: cu.targetName,
                stayDescription: cu.targetLinked || updated.formC.stayDescription
              };
            }
          } else if (cu.type === 'Agent') {
            if (updated.broughtBy === oldName) {
              updated.broughtBy = cu.targetName;
              if (cu.targetLinked && (!updated.contactDetails || updated.contactDetails === oldName)) {
                updated.contactDetails = cu.targetLinked;
              }
            }
          } else if (cu.type === 'Official') {
            if (updated.officialName === oldName) {
              updated.officialName = cu.targetName;
              if (cu.targetLinked && (!updated.officialTitle || updated.officialTitle === oldName)) {
                updated.officialTitle = cu.targetLinked;
              }
            }
            if (updated.formC && updated.formC.officialName === oldName) {
              updated.formC = { 
                ...updated.formC, 
                officialName: cu.targetName,
                officialTitle: cu.targetLinked || updated.formC.officialTitle
              };
            }
          } else if (cu.type === 'Reporter') {
            if (updated.formC && updated.formC.reporterName === oldName) {
              updated.formC = { 
                ...updated.formC, 
                reporterName: cu.targetName,
                reporterPhone: cu.targetLinked || updated.formC.reporterPhone
              };
            }
          } else if (cu.type === 'Phone') {
            if (updated.formC && updated.formC.reporterPhone === oldName) {
              updated.formC = { ...updated.formC, reporterPhone: cu.targetName };
            }
          }
        });
      });
      return updated;
    };

    if (setRecords) {
      setRecords(prev => {
        const updatedRecs = prev.map(updateRecordBatchFn);
        localStorage.setItem('imm_records_react', JSON.stringify(updatedRecs));
        miniDB.set('records', updatedRecs).catch(() => {});
        saveCollectionToFirestore('records', updatedRecs).catch(() => {});
        return updatedRecs;
      });
    }
    if (setTempRecords) {
      setTempRecords(prev => {
        const updatedTemps = prev.map(updateRecordBatchFn);
        localStorage.setItem('imm_temp_records_react', JSON.stringify(updatedTemps));
        miniDB.set('tempRecords', updatedTemps).catch(() => {});
        saveCollectionToFirestore('tempRecords', updatedTemps).catch(() => {});
        return updatedTemps;
      });
    }

    const primaryUpdateMap = new Map<number, { targetName: string; targetLinked?: string }>();
    clusterUpdates.forEach(cu => {
      primaryUpdateMap.set(cu.primaryId, { targetName: cu.targetName, targetLinked: cu.targetLinked });
    });

    setMasterData(prev => {
      const mergedList = prev
        .filter(m => !allIdsToRemove.has(m.id))
        .map(m => {
          if (primaryUpdateMap.has(m.id)) {
            const up = primaryUpdateMap.get(m.id)!;
            return {
              ...m,
              name: up.targetName,
              linkedValue: up.targetLinked,
              updatedAt: new Date().toISOString()
            };
          }
          return m;
        });
      localStorage.setItem('imm_master_react', JSON.stringify(mergedList));
      miniDB.set('masterData', mergedList).catch(() => {});
      saveCollectionToFirestore('masterData', mergedList, true).catch(() => {});
      return mergedList;
    });

    setShowMergeModal(false);
    showToast(`MERGED ALL ${duplicateClusters.length} GROUPS SUCCESSFULLY (SYNCED TO CLOUD)`);
  };

  // Dedicated Auto-Sync: Scans recent records to sync & refresh all master data with latest matched values
  const handleSyncMasterFromLatestRecords = () => {
    const allRecs = [...(records || []), ...(tempRecords || [])].sort((a, b) => getRecordTime(b) - getRecordTime(a));
    if (allRecs.length === 0) {
      showToast("စနစ်အတွင်း မှတ်တမ်းများ မရှိသေးပါ (No records found)");
      return;
    }

    const nowIso = new Date().toISOString();
    const updatedMap = new Map<string, MasterItem>();

    // 1. First seed with existing masterData (reversed so newer items override older)
    [...(masterData || [])].reverse().forEach(m => {
      if (m && m.name && m.type) {
        const key = `${m.type}__${m.name.trim().toLowerCase()}`;
        if (!updatedMap.has(key)) {
          updatedMap.set(key, { ...m });
        }
      }
    });

    // 2. Scan records from newest to oldest to pull latest names and latest linkedValues
    allRecs.forEach(r => {
      if (!r) return;
      
      // Stay Location & Stay Description
      const stayName = (r.formC?.address || r.address || '').trim();
      const stayDesc = (r.formC?.stayDescription || r.stayDescription || '').trim();
      if (stayName) {
        const key = `Stay__${stayName.toLowerCase()}`;
        const existing = updatedMap.get(key);
        if (existing) {
          if (stayDesc && (!existing.linkedValue || existing.linkedValue !== stayDesc)) {
            existing.linkedValue = stayDesc;
            existing.updatedAt = nowIso;
          }
        } else {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: stayName,
            type: 'Stay',
            linkedValue: stayDesc || undefined,
            updatedAt: nowIso
          });
        }
      }

      // Agent & Contact
      const agent = (r.broughtBy || '').trim();
      const contact = (r.contactDetails || '').trim();
      if (agent) {
        const key = `Agent__${agent.toLowerCase()}`;
        const existing = updatedMap.get(key);
        if (existing) {
          if (contact && (!existing.linkedValue || existing.linkedValue !== contact)) {
            existing.linkedValue = contact;
            existing.updatedAt = nowIso;
          }
        } else {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: agent,
            type: 'Agent',
            linkedValue: contact || undefined,
            updatedAt: nowIso
          });
        }
      }

      // Official & Title
      const offName = (r.formC?.officialName || r.officialName || '').trim();
      const offTitle = (r.formC?.officialTitle || r.officialTitle || '').trim();
      if (offName) {
        const key = `Official__${offName.toLowerCase()}`;
        const existing = updatedMap.get(key);
        if (existing) {
          if (offTitle && (!existing.linkedValue || existing.linkedValue !== offTitle)) {
            existing.linkedValue = offTitle;
            existing.updatedAt = nowIso;
          }
        } else {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: offName,
            type: 'Official',
            linkedValue: offTitle || undefined,
            updatedAt: nowIso
          });
        }
      }

      // Reporter & Phone
      const repName = (r.formC?.reporterName || '').trim();
      const repPhone = (r.formC?.reporterPhone || '').trim();
      if (repName) {
        const key = `Reporter__${repName.toLowerCase()}`;
        const existing = updatedMap.get(key);
        if (existing) {
          if (repPhone && (!existing.linkedValue || existing.linkedValue !== repPhone)) {
            existing.linkedValue = repPhone;
            existing.updatedAt = nowIso;
          }
        } else {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: repName,
            type: 'Reporter',
            linkedValue: repPhone || undefined,
            updatedAt: nowIso
          });
        }
      }

      // Nationality
      const nat = (r.nationality || '').trim();
      if (nat) {
        const key = `Nationality__${nat.toLowerCase()}`;
        if (!updatedMap.has(key)) {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: nat,
            type: 'Nationality',
            updatedAt: nowIso
          });
        }
      }

      // Visa
      const visa = (r.visaType || '').trim();
      if (visa) {
        const key = `Visa__${visa.toLowerCase()}`;
        if (!updatedMap.has(key)) {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: visa,
            type: 'Visa',
            updatedAt: nowIso
          });
        }
      }

      // Vehicle
      const veh = (r.vehicleInfo || '').trim();
      if (veh) {
        const key = `Vehicle__${veh.toLowerCase()}`;
        if (!updatedMap.has(key)) {
          updatedMap.set(key, {
            id: Date.now() + Math.floor(Math.random() * 10000),
            name: veh,
            type: 'Vehicle',
            updatedAt: nowIso
          });
        }
      }
    });

    const newMasterList = Array.from(updatedMap.values());
    setMasterData(newMasterList);
    showToast(`မာစတာဒေတာ (${newMasterList.length} ခု) အား နောက်ဆုံးမှတ်တမ်းများ (Latest Records) နှင့် Auto-Sync ချိတ်ဆက်ပြီးပါပြီ`);
  };

  return (
    <div className="max-w-7xl mx-auto py-8 px-4 space-y-8">
      <div className="card shadow-md flex flex-wrap justify-between items-center gap-4 border-l-8 border-blue-600">
        <div className="flex items-center gap-3">
          <Database className="text-blue-600" size={24} />
          <div>
            <h3 className="text-lg font-black uppercase">Master Database Management</h3>
            <p className="text-xs text-slate-500 font-bold">Category Labels, Auto-Extraction & Cloud Server Synchronization</p>
          </div>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            onClick={handleSyncMasterFromLatestRecords}
            className="btn bg-cyan-600 hover:bg-cyan-700 text-white text-[11px] font-black uppercase px-4 py-2 flex items-center gap-1.5 shadow-md shadow-cyan-100 cursor-pointer"
            title="လက်ရှိမှတ်တမ်းများထဲမှ နိုင်ငံသား၊ ဗီဇာ၊ ဟိုတယ် စသည့် မာစတာအချက်အလက်များကို အလိုအလျောက် စုစည်းထုတ်ယူမည် (ဆာဗာ Sync မဟုတ်ပါ)"
          >
            <Sparkles size={14} /> Auto-Extract from Records
          </button>
          <button
            onClick={() => {
              if (duplicateClusters.length === 0) {
                showToast("တူညီသော မာစတာဒေတာ မရှိပါ (No duplicates detected)");
              } else {
                setShowMergeModal(true);
              }
            }}
            className="btn bg-purple-600 hover:bg-purple-700 text-white text-[11px] font-black uppercase px-4 py-2 flex items-center gap-1.5 shadow-md shadow-purple-100 cursor-pointer"
            title="Master Data မှ တူညီသော ဒေတာများကို Semi-Auto Merge ပေးနိုင်သော စနစ် (Default Target is LATEST data)"
          >
            <GitMerge size={14} /> Semi-Auto Merge ({duplicateClusters.length})
          </button>
          <button 
            onClick={handleMasterDataUnifiedSync} 
            disabled={isMasterSyncing}
            className={`btn ${isMasterSyncing ? 'bg-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 cursor-pointer'} text-white text-[11px] font-black uppercase px-4 py-2 flex items-center gap-1.5 shadow-md shadow-blue-100`}
            title="ဒေတာအားလုံး (မှတ်တမ်းများ၊ မာစတာဒေတာ၊ စစ်ဆေးမှုများ) ကို Cloud Server သို့ တိုက်ရိုက် Sync ပြုလုပ်မည်"
          >
            <UploadCloud size={14} className={isMasterSyncing ? "animate-bounce" : ""} /> {isMasterSyncing ? "Syncing All Data..." : "Cloud Sync All Data"}
          </button>
          <button onClick={backupCombinedJSON} className="btn bg-indigo-600 text-white hover:bg-indigo-700 text-[11px] uppercase px-4 py-2 flex items-center gap-1.5 border border-indigo-500 shadow-md shadow-indigo-100 cursor-pointer">
            <Download size={14} /> Combined Backup
          </button>
          <label className="btn bg-indigo-800/10 hover:bg-indigo-800/20 text-indigo-700 border border-indigo-200 text-[11px] uppercase px-4 py-2 cursor-pointer flex items-center gap-1.5 shadow-md shadow-indigo-100/50">
            <Upload size={14} /> Combined Restore
            <input type="file" className="hidden" accept=".json" onChange={restoreCombinedJSON} />
          </label>
        </div>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        <motion.div initial={{ x: -20, opacity: 0 }} animate={{ x: 0, opacity: 1 }} className="card h-fit space-y-6 shadow-xl border-t-8 border-indigo-700">
         <div className="flex items-center gap-3 mb-2">
           <div className="bg-indigo-100 p-2.5 rounded-lg text-indigo-700">
             <PlusCircle size={24} />
           </div>
           <h2 className="text-xl font-black uppercase tracking-tight">Register Master Info</h2>
         </div>
         
         <div className="space-y-4">
           <div>
             <label className="input-label">Data Name (ဥပမာ- Thai, Business...)</label>
             <input type="text" value={mName} onChange={(e) => setMName(e.target.value)} className="input-field" placeholder="Enter Label Name..." />
           </div>
           <div>
             <label className="input-label">Categorize As</label>
             <select value={mType} onChange={(e) => { setMType(e.target.value as MasterItem['type']); setMLinked(''); }} className="input-field">
               <option value="Nationality">Nationality</option>
               <option value="Visa">Visa Type</option>
               <option value="Stay">Stay Location</option>
               <option value="Vehicle">Vehicle Info</option>
               <option value="Agent">Agent Name</option>
               <option value="Contact">Contact Details</option>
               <option value="Official">Official Name</option>
               <option value="Title">Official Title</option>
               <option value="Reporter">Reporter Name</option>
               <option value="Phone">Reporter Phone</option>
             </select>
           </div>

           {mType === 'Agent' && (
             <div>
               <label className="input-label font-black text-indigo-600">Contact Details (ဖုန်းနံပါတ် သို့မဟုတ် ဆက်သွယ်ရန်)</label>
               <input 
                 type="text" 
                 value={mLinked} 
                 onChange={(e) => setMLinked(e.target.value)} 
                 className="input-field border-indigo-200" 
                 placeholder="Enter Contact Info..." 
               />
             </div>
           )}

           {mType === 'Official' && (
             <div>
               <label className="input-label font-black text-indigo-600">Official Title (ရာထူး)</label>
               <input 
                 type="text" 
                 value={mLinked} 
                 onChange={(e) => setMLinked(e.target.value)} 
                 className="input-field border-indigo-200" 
                 placeholder="Enter Official Title..." 
               />
             </div>
           )}

           {mType === 'Reporter' && (
             <div>
               <label className="input-label font-black text-indigo-600">Reporter Phone (ဖုန်းနံပါတ်)</label>
               <input 
                 type="text" 
                 value={mLinked} 
                 onChange={(e) => setMLinked(e.target.value)} 
                 className="input-field border-indigo-200" 
                 placeholder="Enter Reporter Phone..." 
               />
             </div>
           )}

           <button onClick={addMaster} className="btn w-full bg-indigo-700 text-white hover:bg-black font-black uppercase tracking-widest shadow-lg">Save to Database</button>

            {mType === 'Stay' && (
              <div className="p-4 bg-indigo-50 rounded-xl border border-indigo-100 space-y-2 animate-in fade-in duration-300">
                <label className="input-label font-black text-indigo-700">Address Box (လိပ်စာအပြည့်အစုံ)</label>
                <input 
                  type="text" 
                  value={mLinked} 
                  onChange={(e) => setMLinked(e.target.value)} 
                  className="input-field border-indigo-200 focus:border-indigo-500 bg-white text-xs font-bold" 
                  placeholder="Enter Address/Location details..." 
                />
              </div>
            )}
         </div>
      </motion.div>

      <motion.div initial={{ x: 20, opacity: 0 }} animate={{ x: 0, opacity: 1 }} className="card">
        <div className="flex justify-between items-center mb-6">
          <h2 className="text-xl font-black uppercase flex items-center gap-2"><Database size={20} /> Master Database</h2>
          <div className="relative group">
            <input 
              type="text" 
              value={mSearch} 
              onChange={(e) => setMSearch(e.target.value)} 
              className="input-field pl-8 text-xs h-9 w-48 bg-slate-50 border-slate-200 focus:bg-white transition-all text-gray-800" 
              placeholder="Filter Master..." 
            />
            <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400 group-focus-within:text-slate-600" />
            {mSearch && (
              <button 
                onClick={() => setMSearch('')}
                className="absolute right-2 top-2 text-slate-400 hover:text-red-500 transition-colors"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>

        <div className="max-h-[500px] overflow-y-auto border border-gray-100 rounded-xl">
           <table className="w-full text-left text-sm">
              <thead className="bg-gray-50 sticky top-0 z-10">
                 <tr className="border-b">
                   <th className="p-3 text-[10px] font-black uppercase text-slate-500">Label Name</th>
                   <th className="p-3 text-[10px] font-black uppercase text-slate-500">Category</th>
                   <th className="p-3 text-right"></th>
                 </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {masterData.filter(m => m.name.toLowerCase().includes(mSearch.toLowerCase())).map(m => (
                  <tr key={m.id} className="hover:bg-slate-50 group transition-colors">
                    <td className="p-3 text-sm text-slate-800">
                      <div className="font-black text-slate-900">{m.name}</div>
                      {m.linkedValue && (
                        <div className="text-[10px] text-indigo-500 font-bold mt-0.5">
                          {m.type === 'Agent' && `🔗 Contact: ${m.linkedValue}`}
                          {m.type === 'Official' && `🔗 Title: ${m.linkedValue}`}
                          {m.type === 'Reporter' && `🔗 Phone: ${m.linkedValue}`}
                          {m.type === 'Stay' && `🔗 Address: ${m.linkedValue}`}
                        </div>
                      )}
                    </td>
                    <td className="p-3">
                       <span className="text-[10px] font-black px-2 py-0.5 rounded-full bg-slate-100 text-slate-500 border border-slate-200 uppercase tracking-tighter shadow-sm">
                          {m.type}
                       </span>
                    </td>
                    <td className="p-3 text-right">
                      <div className="flex items-center justify-end gap-2">
                         {editingId === m.id ? (
                           <div className="flex flex-col md:flex-row items-start md:items-center gap-1.5 p-1.5 bg-amber-50 rounded-xl border border-amber-200 text-left">
                             <div className="flex flex-col gap-1 w-full min-w-[150px]">
                               <input 
                                 type="text" 
                                 value={editName} 
                                 onChange={(e) => setEditName(e.target.value)}
                                 className="px-2 py-1.5 border border-amber-300 rounded-lg text-xs w-full focus:ring-1 focus:ring-amber-500 outline-none bg-white font-black"
                                 placeholder="Name"
                                 autoFocus
                               />
                               {(m.type === 'Agent' || m.type === 'Official' || m.type === 'Reporter' || m.type === 'Stay') && (
                                 <input 
                                   type="text" 
                                   value={editLinked} 
                                   onChange={(e) => setEditLinked(e.target.value)}
                                   className="px-2 py-1.5 border border-amber-300 rounded-lg text-xs w-full focus:ring-1 focus:ring-amber-500 outline-none bg-white font-bold"
                                   placeholder={
                                     m.type === 'Agent' ? 'Contact details' :
                                     m.type === 'Official' ? 'Official title' :
                                     m.type === 'Stay' ? 'Address Box / Details' : 'Phone number'
                                   }
                                 />
                               )}
                             </div>
                             <div className="flex gap-1.5">
                               <button onClick={saveEdit} className="text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg text-[9px] font-black px-2.5 py-1.5 shadow-sm uppercase flex items-center gap-1"><Check size={12} strokeWidth={3} /> Save</button>
                               <button onClick={() => setEditingId(null)} className="text-gray-600 bg-gray-200 hover:bg-gray-300 rounded-lg text-[9px] font-black px-2.5 py-1.5 uppercase flex items-center gap-1"><X size={12} strokeWidth={3} /> Cancel</button>
                             </div>
                           </div>
                         ) : (
                           <>
                             <button 
                               onClick={() => startEdit(m)}
                               className="text-[9px] font-black uppercase bg-amber-500 text-white px-2.5 py-1 rounded-lg hover:bg-amber-600 shadow-sm transition-all flex items-center gap-1"
                               title="Edit Label Name"
                             >
                               <Edit2 size={11} strokeWidth={2.5} /> Edit
                             </button>
                             <button 
                               onClick={() => {
                                 setSearchQuery(m.name);
                                 setActiveTab('data');
                                 window.scrollTo({ top: 0, behavior: 'smooth' });
                                }}
                               className="text-[9px] font-black uppercase bg-[#2C6CB0] text-white px-2.5 py-1 rounded-lg shadow-sm hover:bg-black flex items-center gap-1.5 transition-all"
                             >
                               <Search size={10} strokeWidth={3} /> History
                             </button>
                             {confirmMasterDeleteId === m.id ? (
                               <div className="flex gap-1 animate-in zoom-in-95">
                                 <button 
                                   onClick={() => deleteMaster(m.id)} 
                                   className="text-[9px] font-black uppercase bg-red-600 text-white px-2.5 py-1 rounded-lg shadow-sm hover:bg-red-700 transition-all"
                                 >
                                   YES
                                 </button>
                                 <button 
                                   onClick={() => setConfirmMasterDeleteId(null)} 
                                   className="text-[9px] font-black uppercase bg-gray-200 text-gray-600 px-2.5 py-1 rounded-lg shadow-sm hover:bg-gray-300 transition-all"
                                 >
                                   NO
                                 </button>
                               </div>
                             ) : (
                               <button 
                                 onClick={() => setConfirmMasterDeleteId(m.id)} 
                                 className="text-[9px] font-black uppercase bg-red-600 text-white px-2.5 py-1 rounded-lg hover:bg-red-700 shadow-sm transition-all flex items-center gap-1"
                                 title="Delete Master Item"
                               >
                                 <Trash2 size={11} strokeWidth={2.5} /> Delete
                               </button>
                             )}
                           </>
                         )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
           </table>
        </div>
      </motion.div>
    </div>

    {/* SEMI-AUTO MERGE DEDUPLICATION MODAL */}
    {showMergeModal && (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm no-print">
        <div className="relative bg-white w-full max-w-3xl rounded-[2rem] overflow-hidden shadow-2xl border-4 border-purple-900 flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
          <div className="bg-purple-900 p-5 text-white flex justify-between items-center">
            <div className="flex items-center gap-3">
              <Layers size={22} className="text-purple-300" />
              <div>
                <h3 className="font-black text-base uppercase tracking-wider">Master Data Semi-Automatic Deduplication & Merge</h3>
                <p className="text-[11px] text-purple-200 font-bold mt-0.5">
                  တူညီသော မာစတာဒေတာများကို ရှာဖွေပြီး စနစ်တစ်ခုလုံး၏ မှတ်တမ်းများနှင့်တကွ ပေါင်းစည်းပေးမည့် စနစ်
                </p>
              </div>
            </div>
            <button
              onClick={() => setShowMergeModal(false)}
              className="hover:bg-white/10 p-1.5 rounded-full text-white transition-colors cursor-pointer"
            >
              <X size={20} />
            </button>
          </div>

          <div className="p-6 overflow-y-auto space-y-6 text-xs">
            <div className="bg-purple-50 p-4 rounded-2xl border border-purple-200 flex justify-between items-center flex-wrap gap-3">
              <div>
                <p className="font-black text-purple-950 text-sm">
                  တွေ့ရှိရသော တူညီမှု အုပ်စုပေါင်း: <span className="text-purple-600 font-extrabold">{duplicateClusters.length}</span> ခု
                </p>
                <p className="text-[11px] text-purple-800 font-medium">
                  အောက်ပါ အုပ်စုများအနက် မူလ (Primary) အဖြစ် ထားရှိလိုသော အမည်ကို ရွေးချယ်ပြီး ပေါင်းစည်းနိုင်ပါသည်။
                </p>
              </div>
              <button
                onClick={handleMergeAllClusters}
                className="btn bg-purple-700 hover:bg-purple-800 text-white font-black text-xs px-4 py-2 flex items-center gap-1.5 shadow-md shadow-purple-200 cursor-pointer"
              >
                <Layers size={14} /> Merge All Groups ({duplicateClusters.length})
              </button>
            </div>

            <div className="space-y-4">
              {duplicateClusters.map((cluster, cIdx) => {
                const currentTargetId = selectedTargetIds[cluster.key] || cluster.items[0].id;

                return (
                  <div key={cluster.key + '_' + cIdx} className="bg-white p-4 rounded-2xl border-2 border-slate-200 shadow-sm space-y-3 hover:border-purple-300 transition-colors">
                    <div className="flex justify-between items-center border-b border-slate-100 pb-2">
                      <div className="flex items-center gap-2">
                        <span className="bg-purple-100 text-purple-900 font-black px-2.5 py-0.5 rounded-full text-[10px] uppercase">
                          {cluster.type}
                        </span>
                        <span className="font-bold text-slate-500 text-[11px]">
                          Duplicate Cluster #{cIdx + 1} ({cluster.items.length} items)
                        </span>
                      </div>
                      <button
                        onClick={() => handleMergeCluster(cluster.key, currentTargetId)}
                        className="btn bg-emerald-600 hover:bg-emerald-700 text-white font-black text-[10px] uppercase px-3 py-1.5 flex items-center gap-1 shadow-sm cursor-pointer"
                      >
                        Merge This Group
                      </button>
                    </div>

                    <div className="space-y-2">
                      <p className="font-bold text-slate-700 text-[11px]">Select Primary Target Name to Keep:</p>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        {cluster.items.map(item => (
                          <label
                            key={item.id}
                            className={`flex items-center gap-2 p-2.5 rounded-xl border cursor-pointer transition-all ${
                              item.id === currentTargetId
                                ? 'bg-purple-50 border-purple-500 ring-2 ring-purple-500/20 font-black text-purple-900'
                                : 'bg-slate-50/50 border-slate-200 hover:bg-slate-100 text-slate-700 font-medium'
                            }`}
                          >
                            <input
                              type="radio"
                              name={`target_${cluster.key}`}
                              checked={item.id === currentTargetId}
                              onChange={() => setSelectedTargetIds(prev => ({ ...prev, [cluster.key]: item.id }))}
                              className="accent-purple-600"
                            />
                            <span className="truncate">{item.name}</span>
                            {item.linkedValue && (
                              <span className="text-[10px] text-slate-400 font-normal">({item.linkedValue})</span>
                            )}
                          </label>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    )}
  </div>
);
};

interface TelegraphRow {
  id: string;
  flightName: string;
  time: string;
  myanmarInM: number;
  myanmarInF: number;
  myanmarOutM: number;
  myanmarOutF: number;
  foreignerInM: number;
  foreignerInF: number;
  foreignerOutM: number;
  foreignerOutF: number;
  remarks: string;
}


export const formatDobInput = (val: string): string => {
  if (!val) return '';
  // Replace slashes or dots with hyphens
  let clean = val.replace(/[\/\.]/g, '-');
  // Only allow digits and hyphens
  clean = clean.replace(/[^0-9\-]/g, '');
  
  // If user typed 8 pure digits without hyphens (e.g., 25121995)
  if (/^\d{8}$/.test(clean)) {
    return `${clean.substring(0, 2)}-${clean.substring(2, 4)}-${clean.substring(4, 8)}`;
  }
  return clean;
};

export const formatDateTimeToDDMMYYYY = (timestampStr: string | undefined | null): string => {
  if (!timestampStr || timestampStr === '-') return '-';
  const s = String(timestampStr).trim();
  if (!s || s === '-') return '-';
  
  let timePortion = '';
  if (s.includes(',')) {
    timePortion = s.split(',')[1]?.trim() || '';
  } else {
    const timeMatch = s.match(/(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)/i);
    if (timeMatch && !s.startsWith(timeMatch[1])) timePortion = timeMatch[1];
  }

  const ddmmyyyy = formatToDDMMYYYY(s);
  if (timePortion) {
    return `${ddmmyyyy}, ${timePortion}`;
  }
  return ddmmyyyy;
};

const toBurmeseDigits = (numStr: string | number) => {
  const burmeseDigits = ['၀', '၁', '၂', '၃', '၄', '၅', '၆', '၇', '၈', '၉'];
  return String(numStr).replace(/[0-9]/g, (match) => burmeseDigits[parseInt(match)]);
};

const toBurmeseDate = (dateStr: string) => {
  if (!dateStr) return '';
  const ddmmyyyy = formatToDDMMYYYY(dateStr);
  const parts = ddmmyyyy.split('-');
  if (parts.length !== 3) return dateStr;
  const day = toBurmeseDigits(parts[0]);
  const month = toBurmeseDigits(parts[1]);
  const year = toBurmeseDigits(parts[2]);
  return `${year} ခု၊ ${month} လ ၊ ${day} ရက်`;
};

const toBurmeseSlashDate = (dateStr: string) => {
  if (!dateStr) return '';
  const ddmmyyyy = formatToDDMMYYYY(dateStr);
  const parts = ddmmyyyy.split('-');
  if (parts.length !== 3) return dateStr;
  const day = toBurmeseDigits(parts[0]);
  const month = toBurmeseDigits(parts[1]);
  const year = toBurmeseDigits(parts[2]);
  return `${day} - ${month} - ${year}`;
};

const burmeseMonthsShort: Record<number, string> = {
  1: 'ဇန်',
  2: 'ဖေ',
  3: 'မတ်',
  4: 'ဧ',
  5: 'မေ',
  6: 'ဇွန်',
  7: 'ဇူ',
  8: 'သြ',
  9: 'စက်',
  10: 'အောက်',
  11: 'နို',
  12: 'ဒီ'
};

const toBurmeseShortDateTime = (dateStr: string, timeStr: string) => {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  const year = toBurmeseDigits(parts[0]);
  const monthNum = parseInt(parts[1]);
  const monthShort = burmeseMonthsShort[monthNum] || parts[1];
  const day = toBurmeseDigits(parseInt(parts[2]).toString());
  // The time portion is written by hand, so leave space for it (approx 14 non-breaking spaces)
  return `${year} ${monthShort} ${day}\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0`;
};

const AutoFitSheetWrapper = ({
  children,
  isLandscape = true,
  autoFit = true,
  zoom = 1,
  className = ""
}: {
  children: React.ReactNode;
  isLandscape?: boolean;
  autoFit?: boolean;
  zoom?: number;
  className?: string;
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  const [contentHeight, setContentHeight] = useState<number | null>(null);
  const innerRef = useRef<HTMLDivElement>(null);

  const sheetWidthMm = isLandscape ? 297 : 210;
  const targetWidthPx = Math.round(sheetWidthMm * 3.7795); // 96 DPI standard

  useEffect(() => {
    const updateScale = () => {
      if (!containerRef.current) return;
      const containerWidth = containerRef.current.clientWidth;
      if (autoFit) {
        const availWidth = Math.max(320, containerWidth - 32);
        const calculatedScale = Math.min(1, availWidth / targetWidthPx);
        setScale(Number((calculatedScale * zoom).toFixed(3)));
      } else {
        setScale(zoom);
      }
      if (innerRef.current) {
        setContentHeight(innerRef.current.scrollHeight);
      }
    };

    updateScale();
    window.addEventListener('resize', updateScale);
    let ro: ResizeObserver | null = null;
    if (typeof ResizeObserver !== 'undefined') {
      ro = new ResizeObserver(updateScale);
      if (containerRef.current) ro.observe(containerRef.current);
      if (innerRef.current) ro.observe(innerRef.current);
    }

    return () => {
      window.removeEventListener('resize', updateScale);
      if (ro) ro.disconnect();
    };
  }, [autoFit, zoom, targetWidthPx]);

  const scaledHeight = contentHeight && scale < 1 ? Math.round(contentHeight * scale) : undefined;

  return (
    <div 
      ref={containerRef}
      className={`w-full overflow-x-auto bg-slate-100 p-2 sm:p-5 rounded-2xl border border-slate-200 flex justify-center items-start shadow-inner ${className}`}
      style={{ minHeight: scaledHeight ? `${scaledHeight + 32}px` : undefined }}
    >
      <div
        style={{
          transform: scale !== 1 ? `scale(${scale})` : undefined,
          transformOrigin: 'top center',
          width: `${sheetWidthMm}mm`,
          minWidth: `${sheetWidthMm}mm`,
          marginBottom: contentHeight && scale < 1 ? `-${Math.round(contentHeight * (1 - scale))}px` : undefined,
          transition: 'transform 0.15s ease-out'
        }}
      >
        <div ref={innerRef}>
          {children}
        </div>
      </div>
    </div>
  );
};

const TelegraphTables = ({
  records,
  summaries,
  currentUser,
  setActivePrintPreview,
  showToast,
  movementMap = {},
  checkingHistory = [],
  dossierHistory = [],
  masterData = [],
  tempRecords = []
}: {
  records: ImmRecord[];
  summaries: VehicleSummary[];
  currentUser: any;
  setActivePrintPreview: React.Dispatch<React.SetStateAction<any>>;
  showToast: (msg: string) => void;
  movementMap?: Record<string, MovementData>;
  checkingHistory?: CheckingHistoryEntry[];
  dossierHistory?: DossierRecord[];
  masterData?: MasterItem[];
  tempRecords?: ImmRecord[];
}) => {

  const isRecordOnDate = (r: ImmRecord, targetDateStr: string) => {
    if (!targetDateStr || !r || !r.timestamp) return false;
    const [y, m, d] = targetDateStr.split('-');
    const slashDate = `${d}/${m}/${y}`;
    const dashDate = `${d}-${m}-${y}`;
    const isoDate = `${y}-${m}-${d}`;
    const cleanTs = String(r.timestamp).trim();
    if (cleanTs.startsWith(slashDate) || cleanTs.startsWith(dashDate) || cleanTs.startsWith(isoDate)) {
      return true;
    }
    const tTime = parseTimestamp(r.timestamp);
    if (tTime === 0) return false;
    const dt = new Date(tTime);
    const yr = dt.getFullYear();
    const mo = String(dt.getMonth() + 1).padStart(2, '0');
    const da = String(dt.getDate()).padStart(2, '0');
    return `${yr}-${mo}-${da}` === targetDateStr;
  };

  const aggregateTelegraphRows = (recordsForDate: ImmRecord[]): TelegraphRow[] => {
    // Check if we have summaries for this date first!
    const dateSummaries = (summaries || []).filter(s => s.date === selectedDate);
    if (dateSummaries.length > 0) {
      return dateSummaries.map((s, idx) => {
        // get foreigner counts from records
        const [y, m, d] = selectedDate.split('-');
        const dateStrForRecords = `${d}/${m}/${y}`;
        const fInMatches = records.filter(r => 
          r.logType !== 'FCR' &&
          r.vehicleInfo?.toUpperCase().trim() === s.vehicleNo.toUpperCase().trim() && 
          r.mode === 'IN' && 
          r.timestamp.startsWith(dateStrForRecords)
        );
        const fInM = fInMatches.filter(r => r.gender === 'M').length;
        const fInF = fInMatches.filter(r => r.gender === 'F').length;

        const fOutMatches = records.filter(r => 
          r.logType !== 'FCR' &&
          r.vehicleInfo?.toUpperCase().trim() === s.vehicleNo.toUpperCase().trim() && 
          r.mode === 'OUT' && 
          r.timestamp.startsWith(dateStrForRecords)
        );
        const fOutM = fOutMatches.filter(r => r.gender === 'M').length;
        const fOutF = fOutMatches.filter(r => r.gender === 'F').length;

        // Citizen = Total - Foreigner
        const myanmarInM = Math.max(0, (s.totalInM || 0) - fInM);
        const myanmarInTotal = Math.max(0, (s.totalInT || 0) - (fInM + fInF));
        const myanmarInF = Math.max(0, myanmarInTotal - myanmarInM);

        const myanmarOutM = Math.max(0, (s.totalOutM || 0) - fOutM);
        const myanmarOutTotal = Math.max(0, (s.totalOutT || 0) - (fOutM + fOutF));
        const myanmarOutF = Math.max(0, myanmarOutTotal - myanmarOutM);

        return {
          id: s.id.toString(),
          flightName: s.vehicleNo,
          time: s.time || '0000',
          myanmarInM,
          myanmarInF,
          myanmarOutM,
          myanmarOutF,
          foreignerInM: fInM,
          foreignerInF: fInF,
          foreignerOutM: fOutM,
          foreignerOutF: fOutF,
          remarks: s.remark || 'MGZ-YGN-MGZ'
        };
      });
    }

    const flightMap: Record<string, {
      flightName: string;
      time: string;
      myanmarInM: number;
      myanmarInF: number;
      myanmarOutM: number;
      myanmarOutF: number;
      foreignerInM: number;
      foreignerInF: number;
      foreignerOutM: number;
      foreignerOutF: number;
      routes: string[];
      remarksList: string[];
    }> = {};

    recordsForDate.forEach(r => {
      const flight = (r.vehicleInfo || '').trim().toUpperCase() || 'UNSPECIFIED';
      
      let flightTime = '';
      if (r.timestamp) {
        const timeMatch = r.timestamp.match(/(\d{2}):(\d{2})/);
        if (timeMatch) {
          flightTime = `${timeMatch[1]}${timeMatch[2]}`;
        }
      }

      if (!flightMap[flight]) {
        flightMap[flight] = {
          flightName: flight,
          time: flightTime || '0000',
          myanmarInM: 0,
          myanmarInF: 0,
          myanmarOutM: 0,
          myanmarOutF: 0,
          foreignerInM: 0,
          foreignerInF: 0,
          foreignerOutM: 0,
          foreignerOutF: 0,
          routes: [],
          remarksList: []
        };
      }

      const item = flightMap[flight];
      if (flightTime && item.time === '0000') {
        item.time = flightTime;
      }

      const isMMR = (r.nationality || '').trim().toUpperCase() === 'MYANMAR' || (r.nationality || '').trim().toUpperCase() === 'MMR' || (r.nationality || '').trim().toUpperCase() === 'BURMESE';
      const isM = r.gender === 'M';

      if (r.mode === 'IN') {
        if (isMMR) {
          if (isM) item.myanmarInM++;
          else item.myanmarInF++;
        } else {
          if (isM) item.foreignerInM++;
          else item.foreignerInF++;
        }
      } else {
        if (isMMR) {
          if (isM) item.myanmarOutM++;
          else item.myanmarOutF++;
        } else {
          if (isM) item.foreignerOutM++;
          else item.foreignerOutF++;
        }
      }

      const rStart = r.arrivedFrom || '';
      const rEnd = r.departedTo || '';
      if (rStart && rEnd) {
        const routeStr = `${rStart}-${rEnd}`;
        if (!item.routes.includes(routeStr)) {
          item.routes.push(routeStr);
        }
      }
      if (r.remarks && !item.remarksList.includes(r.remarks)) {
        item.remarksList.push(r.remarks);
      }
    });

    return Object.values(flightMap).map((item, idx) => {
      let finalRemark = '';
      if (item.routes.length > 0) {
        const stations: string[] = [];
        item.routes.forEach(rt => {
          rt.split('-').forEach(s => {
            const st = s.trim().toUpperCase();
            if (st && !stations.includes(st)) {
              stations.push(st);
            }
          });
        });
        if (stations.length >= 2) {
          if (stations.includes('MYEIK') || stations.includes('MGZ')) {
            const nonMgz = stations.filter(s => s !== 'MYEIK' && s !== 'MGZ');
            finalRemark = `MGZ-${nonMgz.join('-')}-MGZ`;
          } else {
            finalRemark = stations.join('-');
          }
        } else {
          finalRemark = item.routes.join(', ');
        }
      }
      if (item.remarksList.length > 0) {
        finalRemark = finalRemark ? `${finalRemark} (${item.remarksList.join(', ')})` : item.remarksList.join(', ');
      }

      return {
        id: `${idx}_${item.flightName}`,
        flightName: item.flightName,
        time: item.time,
        myanmarInM: item.myanmarInM,
        myanmarInF: item.myanmarInF,
        myanmarOutM: item.myanmarOutM,
        myanmarOutF: item.myanmarOutF,
        foreignerInM: item.foreignerInM,
        foreignerInF: item.foreignerInF,
        foreignerOutM: item.foreignerOutM,
        foreignerOutF: item.foreignerOutF,
        remarks: finalRemark || 'MGZ-YGN-MGZ'
      };
    });
  };

  const latestRecordDate = useMemo(() => {
    if (records.length === 0) {
      const todayObj = new Date();
      return `${todayObj.getFullYear()}-${String(todayObj.getMonth() + 1).padStart(2, '0')}-${String(todayObj.getDate()).padStart(2, '0')}`;
    }
    const sorted = [...records].sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
    const t = parseTimestamp(sorted[0].timestamp);
    if (t === 0) return '2026-07-20';
    const d = new Date(t);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }, [records]);

  const [selectedDate, setSelectedDate] = useState(latestRecordDate);
  const [reportTime, setReportTime] = useState("1400");
  const [rows, setRows] = useState<TelegraphRow[]>([]);
  const [useBurmeseDigits, setUseBurmeseDigits] = useState(true);
  const [isHeaderConfigOpen, setIsHeaderConfigOpen] = useState(false);
  const [isAutoFit, setIsAutoFit] = useState(true);
  const [zoomLevel, setZoomLevel] = useState(1);

  // Template States initialized with local storage persistence
  const [chief1, setChief1] = useState(() => localStorage.getItem('tt_chief1') || "(ပ) ဒုဌနခမ ၊ လဝက ၊ မြိတ်");
  const [chief2, setChief2] = useState(() => localStorage.getItem('tt_chief2') || "(လ) တစမ ၊ လဝက ၊ ထားဝယ်");
  const [chief3, setChief3] = useState(() => localStorage.getItem('tt_chief3') || "(တ) ခရမ၊ လဝက ၊ မြိတ်");
  const [letterNo, setLetterNo] = useState(() => localStorage.getItem('tt_letterNo') || "\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0/လဝက ॥ ယ ॥");
  const [urgentLevel, setUrgentLevel] = useState(() => localStorage.getItem('tt_urgentLevel') || "အမြန်");
  const [officerTitle, setOfficerTitle] = useState(() => localStorage.getItem('tt_officerTitle') || "ဒုတိယဌာနခွဲမှူး");
  const [officerName, setOfficerName] = useState(() => localStorage.getItem('tt_officerName') || "လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်)");
  const [customIntroText, setCustomIntroText] = useState(() => localStorage.getItem('tt_customIntroText') || "(တစ်) [DATE] ရက်နေ့တွင် ရန်ကုန်လေဆိပ်မှ မြိတ်မြို့၊ လေဆိပ်သို့ လေယာဉ် ( [FLIGHT_COUNT] ) စီး ဝင်ရောက်/ထွက်ခွာခဲ့ပြီး ခရီးသည်စာရင်းမှာ အောက်ပါအတိုင်းဖြစ် -");
  const [telegraphTitle, setTelegraphTitle] = useState(() => localStorage.getItem('tt_telegraphTitle') || "လေယာဉ် ရောက်ရှိ/ထွက်ခွါမှု တင်ပြခြင်း");
  const [activeSubTab, setActiveSubTab] = useState<'ALL' | 'TELEGRAPH' | 'FOREIGNER_IN' | 'FOREIGNER_OUT' | 'COMPANY_RESIDENTS' | 'OTHER_RESIDENTS' | 'ADDRESS_SUMMARY'>('ALL');

  const telegraphRef = React.useRef<HTMLDivElement>(null);
  const inwardRef = React.useRef<HTMLDivElement>(null);
  const outwardRef = React.useRef<HTMLDivElement>(null);
  const companyRef = React.useRef<HTMLDivElement>(null);
  const otherRef = React.useRef<HTMLDivElement>(null);
  const addressSummaryRef = React.useRef<HTMLDivElement>(null);

  // Persistence helpers
  const handleChief1Change = (val: string) => { setChief1(val); localStorage.setItem('tt_chief1', val); };
  const handleChief2Change = (val: string) => { setChief2(val); localStorage.setItem('tt_chief2', val); };
  const handleChief3Change = (val: string) => { setChief3(val); localStorage.setItem('tt_chief3', val); };
  const handleLetterNoChange = (val: string) => { setLetterNo(val); localStorage.setItem('tt_letterNo', val); };
  const handleUrgentLevelChange = (val: string) => { setUrgentLevel(val); localStorage.setItem('tt_urgentLevel', val); };
  const handleOfficerTitleChange = (val: string) => { setOfficerTitle(val); localStorage.setItem('tt_officerTitle', val); };
  const handleOfficerNameChange = (val: string) => { setOfficerName(val); localStorage.setItem('tt_officerName', val); };
  const handleCustomIntroTextChange = (val: string) => { setCustomIntroText(val); localStorage.setItem('tt_customIntroText', val); };
  const handleTelegraphTitleChange = (val: string) => { setTelegraphTitle(val); localStorage.setItem('tt_telegraphTitle', val); };

  const oklchToRgb = (l: number, c: number, h: number, alpha?: number): string => {
    const hRad = (h * Math.PI) / 180;
    const a = c * Math.cos(hRad);
    const b = c * Math.sin(hRad);

    const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
    const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
    const s_ = l - 0.0894841775 * a - 1.2914855480 * b;

    const l3 = l_ * l_ * l_;
    const m3 = m_ * m_ * m_;
    const s3 = s_ * s_ * s_;

    const rLinear = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
    const gLinear = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
    const bLinear = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

    const convertComponent = (val: number) => {
      const clamped = Math.max(0, Math.min(1, val));
      const res = clamped <= 0.0031308
        ? 12.92 * clamped
        : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
      return Math.round(res * 255);
    };

    const r = convertComponent(rLinear);
    const g = convertComponent(gLinear);
    const db = convertComponent(bLinear);

    if (alpha !== undefined) {
      return `rgba(${r}, ${g}, ${db}, ${alpha})`;
    }
    return `rgb(${r}, ${g}, ${db})`;
  };

  const oklabToRgb = (l: number, a: number, b: number, alpha?: number): string => {
    const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
    const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
    const s_ = l - 0.0894841775 * a - 1.2914855480 * b;

    const l3 = l_ * l_ * l_;
    const m3 = m_ * m_ * m_;
    const s3 = s_ * s_ * s_;

    const rLinear = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
    const gLinear = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
    const bLinear = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

    const convertComponent = (val: number) => {
      const clamped = Math.max(0, Math.min(1, val));
      const res = clamped <= 0.0031308
        ? 12.92 * clamped
        : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
      return Math.round(res * 255);
    };

    const r = convertComponent(rLinear);
    const g = convertComponent(gLinear);
    const db = convertComponent(bLinear);

    if (alpha !== undefined) {
      return `rgba(${r}, ${g}, ${db}, ${alpha})`;
    }
    return `rgb(${r}, ${g}, ${db})`;
  };

  const parseAndConvertOklchAndOklab = (colorStr: string): string => {
    if (typeof colorStr !== 'string') return colorStr;
    if (!colorStr.includes('oklch') && !colorStr.includes('oklab') && !colorStr.includes('color-mix')) return colorStr;

    const oklchRegex = /oklch\(\s*([\d.%-]+)\s+([\d.%-]+)\s+([\d.%-]+)(?:\s*\/\s*([\d.%-]+))?\s*\)/gi;
    const oklabRegex = /oklab\(\s*([\d.%-]+)\s+([\d.%-]+)\s+([\d.%-]+)(?:\s*\/\s*([\d.%-]+))?\s*\)/gi;

    let res = colorStr.replace(oklchRegex, (match, lStr, cStr, hStr, aStr) => {
      try {
        let l = parseFloat(lStr);
        if (lStr.includes('%')) l = l / 100;

        let c = parseFloat(cStr);
        if (cStr.includes('%')) c = c / 100;

        let h = parseFloat(hStr);
        if (hStr.includes('%')) h = (h / 100) * 360;

        let alpha: number | undefined = undefined;
        if (aStr) {
          alpha = parseFloat(aStr);
          if (aStr.includes('%')) alpha = alpha / 100;
        }

        return oklchToRgb(l, c, h, alpha);
      } catch (e) {
        return 'rgb(0, 0, 0)';
      }
    });

    res = res.replace(oklabRegex, (match, lStr, aValStr, bValStr, aStr) => {
      try {
        let l = parseFloat(lStr);
        if (lStr.includes('%')) l = l / 100;

        let aVal = parseFloat(aValStr);
        if (aValStr.includes('%')) aVal = aVal / 100;

        let bVal = parseFloat(bValStr);
        if (bValStr.includes('%')) bVal = bVal / 100;

        let alpha: number | undefined = undefined;
        if (aStr) {
          alpha = parseFloat(aStr);
          if (aStr.includes('%')) alpha = alpha / 100;
        }

        return oklabToRgb(l, aVal, bVal, alpha);
      } catch (e) {
        return 'rgb(0, 0, 0)';
      }
    });

    // Replace color-mix functions
    res = res.replace(/color-mix\([^)]+\)/gi, 'rgba(0, 0, 0, 0.05)');
    // Fallback for any lingering oklab/oklch strings
    res = res.replace(/oklab\([^)]+\)/gi, 'rgb(240, 240, 240)');
    res = res.replace(/oklch\([^)]+\)/gi, 'rgb(240, 240, 240)');

    return res;
  };

  const triggerDirectPrint = (element: HTMLElement | null, isLandscape: boolean = false) => {
    if (!element) {
      showToast("Print target element not found");
      return;
    }
    let printArea = document.getElementById('temp-print-area');
    if (!printArea) {
      printArea = document.createElement('div');
      printArea.id = 'temp-print-area';
      document.body.appendChild(printArea);
    }
    printArea.innerHTML = '';

    const styleEl = document.createElement('style');
    styleEl.innerHTML = `@media print { @page { size: A4 ${isLandscape ? 'landscape' : 'portrait'}; margin: 0.5in; } }`;

    const clone = element.cloneNode(true) as HTMLElement;
    clone.style.display = 'block';
    clone.style.visibility = 'visible';

    printArea.appendChild(styleEl);
    printArea.appendChild(clone);

    window.print();
  };

  const exportElementToPhoto = async (targetRef: React.RefObject<HTMLDivElement>, filename: string) => {
    const element = targetRef.current;
    if (!element) {
      showToast("ဇယားမတွေ့ရှိပါ။");
      return;
    }

    const styleBackups: { el: HTMLStyleElement; content: string }[] = [];

    try {
      showToast("ပုံ (PNG) အဖြစ် သိမ်းဆည်းနေပါသည်...");

      // 1. Sanitize oklch and oklab in all <style> elements on page
      const styles = Array.from(document.querySelectorAll('style'));
      styles.forEach((style) => {
        if (style.textContent && (style.textContent.includes('oklch') || style.textContent.includes('oklab'))) {
          styleBackups.push({ el: style, content: style.textContent });
          style.textContent = parseAndConvertOklchAndOklab(style.textContent);
        }
      });

      const isLandscape = filename.includes('foreigner_inward') || filename.includes('foreigner_outward');
      const widthPx = isLandscape ? 1123 : 794;

      // 2. Clone to fixed position container at top-left (0,0) so scroll position won't cut off or blank out table headers
      const cloneWrapper = document.createElement('div');
      cloneWrapper.style.position = 'fixed';
      cloneWrapper.style.left = '0px';
      cloneWrapper.style.top = '0px';
      cloneWrapper.style.width = isLandscape ? '297mm' : '210mm';
      cloneWrapper.style.backgroundColor = '#ffffff';
      cloneWrapper.style.zIndex = '99999';
      cloneWrapper.style.opacity = '1';
      cloneWrapper.style.pointerEvents = 'none';

      const clonedNode = element.cloneNode(true) as HTMLElement;
      clonedNode.style.display = 'block';
      clonedNode.style.visibility = 'visible';
      clonedNode.style.transform = 'none';
      clonedNode.style.margin = '0';
      clonedNode.style.padding = '0.5in';
      clonedNode.style.boxSizing = 'border-box';
      clonedNode.style.backgroundColor = '#ffffff';
      clonedNode.style.fontFamily = "'Pyidaungsu', 'Pyidaungsu Number', sans-serif";
      clonedNode.style.color = '#000000';
      clonedNode.style.width = isLandscape ? '297mm' : '210mm';

      // Ensure all cloned child elements carry explicit Pyidaungsu font and clean text formatting
      const allCloned = Array.from(clonedNode.querySelectorAll('*')) as HTMLElement[];
      allCloned.forEach((el) => {
        el.style.fontFamily = "'Pyidaungsu', 'Pyidaungsu Number', sans-serif";
        el.style.letterSpacing = 'normal';
        el.style.wordSpacing = 'normal';
        
        if (el.style && el.style.cssText) {
          if (el.style.cssText.includes('oklch') || el.style.cssText.includes('oklab') || el.style.cssText.includes('color-mix')) {
            el.style.cssText = parseAndConvertOklchAndOklab(el.style.cssText);
          }
        }
        if (el.tagName === 'TABLE') {
          el.style.borderCollapse = 'collapse';
          el.style.border = '1px solid #000000';
          el.style.width = '100%';
          el.style.letterSpacing = 'normal';
          el.style.position = 'static';
        }
        if (el.tagName === 'TR') {
          el.style.backgroundColor = 'transparent';
          el.style.background = 'none';
          el.style.position = 'static';
        }
        if (el.tagName === 'TH' || el.tagName === 'TD') {
          el.style.border = '1px solid #000000';
          el.style.color = '#000000';
          el.style.lineHeight = '1.15';
          el.style.verticalAlign = 'middle';
          const isLeft = el.classList.contains('text-left') || el.style.textAlign === 'left';
          const isRight = el.classList.contains('text-right') || el.style.textAlign === 'right';
          el.style.textAlign = isLeft ? 'left' : isRight ? 'right' : 'center';
          el.style.paddingTop = '4px';
          el.style.paddingBottom = '10px';
          el.style.paddingLeft = '5px';
          el.style.paddingRight = '5px';
          el.style.letterSpacing = 'normal';
          el.style.wordBreak = 'break-word';
          el.style.boxSizing = 'border-box';
          el.style.position = 'static';
          el.style.zIndex = 'auto';
        }
        if (el.tagName === 'TH') {
          el.style.backgroundColor = '#f8fafc';
          el.style.fontWeight = 'bold';
        }
      });

      cloneWrapper.appendChild(clonedNode);
      document.body.appendChild(cloneWrapper);

      await new Promise(r => setTimeout(r, 150));

      const canvas = await html2canvas(clonedNode, {
        scale: 2,
        useCORS: true,
        allowTaint: true,
        backgroundColor: "#ffffff",
        logging: false,
        scrollX: 0,
        scrollY: 0,
        x: 0,
        y: 0,
        windowWidth: widthPx,
        windowHeight: clonedNode.scrollHeight + 50,
        onclone: (clonedDoc) => {
          const clonedStyles = Array.from(clonedDoc.querySelectorAll('style'));
          clonedStyles.forEach((style) => {
            if (style.textContent) {
              style.textContent = parseAndConvertOklchAndOklab(style.textContent);
            }
          });

          const elements = Array.from(clonedDoc.querySelectorAll('*')) as HTMLElement[];
          elements.forEach((el) => {
            el.style.fontFamily = "'Pyidaungsu', 'Pyidaungsu Number', sans-serif";
            if (el.tagName === 'TABLE') {
              el.style.borderCollapse = 'collapse';
              el.style.border = '1px solid #000000';
              el.style.width = '100%';
              el.style.position = 'static';
            }
            if (el.tagName === 'TR') {
              el.style.backgroundColor = 'transparent';
              el.style.background = 'none';
              el.style.position = 'static';
            }
            if (el.tagName === 'TH' || el.tagName === 'TD') {
              el.style.border = '1px solid #000000';
              el.style.color = '#000000';
              el.style.verticalAlign = 'middle';
              const isLeft = el.classList.contains('text-left') || el.style.textAlign === 'left';
              const isRight = el.classList.contains('text-right') || el.style.textAlign === 'right';
              el.style.textAlign = isLeft ? 'left' : isRight ? 'right' : 'center';
              el.style.lineHeight = '1.15';
              el.style.paddingTop = '4px';
              el.style.paddingBottom = '10px';
              el.style.paddingLeft = '5px';
              el.style.paddingRight = '5px';
              el.style.boxSizing = 'border-box';
              el.style.position = 'static';
              el.style.zIndex = 'auto';
            }
            if (el.tagName === 'TH') {
              el.style.backgroundColor = '#f8fafc';
              el.style.fontWeight = 'bold';
            }
          });
        }
      });

      if (document.body.contains(cloneWrapper)) {
        document.body.removeChild(cloneWrapper);
      }

      const downloadFile = (url: string) => {
        const link = document.createElement('a');
        link.href = url;
        link.download = `${filename}.png`;
        link.setAttribute('download', `${filename}.png`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      };

      if (canvas.toBlob) {
        canvas.toBlob((blob) => {
          if (blob) {
            const blobUrl = URL.createObjectURL(blob);
            downloadFile(blobUrl);
            setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
            showToast("ပုံ (PNG) အဖြစ် အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။");
          } else {
            const dataUrl = canvas.toDataURL('image/png');
            downloadFile(dataUrl);
            showToast("ပုံ (PNG) အဖြစ် အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။");
          }
        }, 'image/png');
      } else {
        const dataUrl = canvas.toDataURL('image/png');
        downloadFile(dataUrl);
        showToast("ပုံ (PNG) အဖြစ် အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။");
      }
    } catch (err) {
      console.error("Export photo error:", err);
      showToast("ပုံသိမ်းဆည်းရာတွင် အမှားအယွင်းရှိပါသည်");
    } finally {
      styleBackups.forEach(({ el, content }) => {
        el.textContent = content;
      });
    }
  };

  const isCoLtdAddress = (addressStr: string) => {
    const clean = (addressStr || '').toUpperCase().trim();
    if (!clean) return false;
    if (masterData && masterData.length > 0) {
      const match = masterData.find(m => {
        const itemAny = m as any;
        const addr = (itemAny.address || m.name || '').toUpperCase().trim();
        return addr && (addr === clean || clean.includes(addr));
      });
      if (match) {
        const itemAny = match as any;
        if (
          itemAny.category === 'COMPANY' || 
          String(itemAny.type).toUpperCase() === 'COMPANY' || 
          itemAny.isCompany === true ||
          (itemAny.linkedValue && String(itemAny.linkedValue).toUpperCase().includes('COMPANY'))
        ) {
          return true;
        }
      }
    }
    return (
      (clean.includes('CO') && (clean.includes('LTD') || clean.includes('L.T.D'))) ||
      clean.includes('COMPANY') ||
      clean.includes('CORPORATION') ||
      clean.includes('ENTERPRISE') ||
      clean.includes('MINING') ||
      clean.includes('CO.,') ||
      clean.includes('CO.,LTD') ||
      clean.includes('ကုမ္ပဏီ') ||
      clean.includes('လုပ်ငန်း')
    );
  };

  const getElapsedDays = (m: MovementData, targetDateStr: string) => {
    const arrMs = m.inTime || parseTimestamp(m.in);
    if (!arrMs) return 0;
    let refMs = Date.now();
    if (targetDateStr) {
      refMs = new Date(targetDateStr + "T23:59:59").getTime();
    }
    const diffTime = refMs - arrMs;
    if (diffTime < 0) return 1;
    return Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;
  };

  const getLatestVerificationRemark = (passport: string, dossierHistoryList: DossierRecord[], recordsList: ImmRecord[]) => {
    try {
      const pp = (passport || '').toUpperCase();
      if (!pp) return '-';
      
      // 1. Check dossierHistory (INV tab investigation dossiers)
      const dHist = (dossierHistoryList || []).filter(d => d && d.remarksMap && typeof d.remarksMap === 'object' && d.remarksMap[pp] && typeof d.remarksMap[pp] === 'string' && d.remarksMap[pp].trim());
      if (dHist.length > 0) {
        const latest = dHist[dHist.length - 1];
        return String(latest.remarksMap![pp]).trim();
      }
      
      // 2. Check recordsList (INV tab investigation notes saved on ImmRecord)
      const userRecs = (recordsList || [])
        .filter(r => r && (r.passport || '').toUpperCase() === pp && r.remarks && typeof r.remarks === 'string' && r.remarks.trim())
        .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
        
      if (userRecs.length > 0) {
        return String(userRecs[0].remarks).trim();
      }
      
      return '-';
    } catch (e) {
      return '-';
    }
  };

  const stillInList = useMemo(() => {
    const list: MovementData[] = [];
    const seenPassports = new Set<string>();
    const allRecs = [...(records || []), ...(tempRecords || [])];

    // 1. Process movementMap entries
    Object.entries(movementMap || {}).forEach(([pass, movVal]) => {
      const mov = movVal as MovementData;
      const pUpper = pass.trim().toUpperCase();
      if (!pUpper || seenPassports.has(pUpper)) return;
      const inTime = Number(mov.inTime || 0);
      const outTime = Number(mov.outTime || 0);
      const hasDeparted = Boolean(mov.out || mov.isStillIn === false || (outTime > 0 && inTime > 0 && outTime >= inTime));
      if (hasDeparted || !mov.in) return;

      // check if latest record in allRecs is OUT
      const matchRecords = allRecs
        .filter(r => (r.passport || '').trim().toUpperCase() === pUpper)
        .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp) || (b.id || 0) - (a.id || 0));
      const latestRec = matchRecords[0];
      if (latestRec && latestRec.mode === 'OUT') return;

      seenPassports.add(pUpper);
      list.push({
        ...mov,
        loc: latestRec?.address || mov.loc || '',
        n: latestRec?.fullname || mov.n || '',
        nat: latestRec?.nationality || mov.nat || '',
        visa: latestRec?.visaType || mov.visa || '',
        gender: (latestRec?.gender || mov.gender || 'M') as 'M' | 'F'
      });
    });

    // 2. Fallback for any foreigners in allRecs not captured in movementMap
    const passportMap = new Map<string, ImmRecord[]>();
    allRecs.forEach(r => {
      if (!r || !r.passport || r.logType === 'FCR') return;
      const pUpper = r.passport.trim().toUpperCase();
      if (seenPassports.has(pUpper)) return;
      if (!passportMap.has(pUpper)) passportMap.set(pUpper, []);
      passportMap.get(pUpper)!.push(r);
    });

    passportMap.forEach((pRecs, pUpper) => {
      pRecs.sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp) || (b.id || 0) - (a.id || 0));
      const latest = pRecs[0];
      if (latest && latest.mode === 'IN') {
        seenPassports.add(pUpper);
        list.push({
          p: pUpper,
          n: latest.fullname || '',
          nat: latest.nationality || '',
          loc: latest.address || '',
          visa: latest.visaType || '',
          visaNumber: latest.visaNumber || '',
          dob: latest.dob || '',
          gender: (latest.gender || 'M') as 'M' | 'F',
          start: latest.stayFrom || '',
          end: latest.stayTo || '',
          allowed: latest.totalDays || '',
          vInfo: latest.vehicleInfo || '',
          agent: latest.broughtBy || '',
          contact: latest.contactDetails || '',
          offName: latest.officialName || '',
          offTitle: latest.officialTitle || '',
          repName: '',
          repPhone: '',
          in: latest.timestamp,
          out: '',
          inTime: parseTimestamp(latest.timestamp),
          outTime: 0,
          latestTime: parseTimestamp(latest.timestamp),
          lastId: latest.id
        });
      }
    });

    return list;
  }, [movementMap, records, tempRecords]);

  const companyResidentsList = useMemo(() => {
    return stillInList
      .filter(m => isCoLtdAddress(m.loc))
      .sort((a, b) => getElapsedDays(b, selectedDate) - getElapsedDays(a, selectedDate));
  }, [stillInList, selectedDate]);

  const otherResidentsList = useMemo(() => {
    return stillInList
      .filter(m => !isCoLtdAddress(m.loc))
      .sort((a, b) => getElapsedDays(b, selectedDate) - getElapsedDays(a, selectedDate));
  }, [stillInList, selectedDate]);

  const addressSummaryList = useMemo(() => {
    const map: Record<string, { address: string; male: number; female: number; total: number; isCo: boolean }> = {};
    stillInList.forEach(m => {
      const addr = (m.loc || '').trim() || 'အခြား/လိပ်စာမဖော်ပြထားသူများ';
      if (!map[addr]) {
        map[addr] = {
          address: addr,
          male: 0,
          female: 0,
          total: 0,
          isCo: isCoLtdAddress(addr)
        };
      }
      if (m.gender === 'M') {
        map[addr].male += 1;
      } else if (m.gender === 'F') {
        map[addr].female += 1;
      }
      map[addr].total += 1;
    });

    const list = Object.values(map);
    const companyAddrs = list.filter(item => item.isCo).sort((a, b) => a.address.localeCompare(b.address));
    const otherAddrs = list.filter(item => !item.isCo).sort((a, b) => a.address.localeCompare(b.address));
    
    return [...companyAddrs, ...otherAddrs];
  }, [stillInList]);

  const totalAddressSummary = useMemo(() => {
    return addressSummaryList.reduce((acc, curr) => ({
      male: acc.male + curr.male,
      female: acc.female + curr.female,
      total: acc.total + curr.total
    }), { male: 0, female: 0, total: 0 });
  }, [addressSummaryList]);

  const exportTelegraphToPhoto = () => exportElementToPhoto(telegraphRef, `telegraph_report_${selectedDate}`);
  const exportInwardToPhoto = () => exportElementToPhoto(inwardRef, `foreigner_inward_report_${selectedDate}`);
  const exportOutwardToPhoto = () => exportElementToPhoto(outwardRef, `foreigner_outward_report_${selectedDate}`);
  const exportCompanyToPhoto = () => exportElementToPhoto(companyRef, `company_residents_${selectedDate}`);
  const exportOtherToPhoto = () => exportElementToPhoto(otherRef, `other_residents_${selectedDate}`);
  const exportAddressSummaryToPhoto = () => exportElementToPhoto(addressSummaryRef, `address_summary_${selectedDate}`);

  const exportInwardToExcel = () => {
    const data = foreignerInRecords.map((r, idx) => ({
      "စဉ်": idx + 1,
      "အမည်": r.fullname,
      "လိင်": r.gender,
      "နိုင်ငံသား": r.nationality || '-',
      "နိုင်ငံကူးလက်မှတ်": r.passport,
      "ပြည်ဝင်ဗီဇာအမျိုးအစား": r.visaType || '-',
      "နေခွင့်ကာလ": (r.stayFrom || r.stayTo) ? `${formatToDDMMYYYY(r.stayFrom)} to ${formatToDDMMYYYY(r.stayTo)}` : '-',
      "တည်းခိုမည့်နေ့ရက်(မှ)": formatToDDMMYYYY(r.timestamp),
      "တည်းခိုမည့်နေ့ရက်(ထိ)": formatToDDMMYYYY(r.stayTo),
      "ရောက်ရှိသည့်နေရာ": r.arrivedFrom || '-',
      "ထွက်ခွာသည့်နေရာ": r.departedTo || '-',
      "ယာဉ်အမှတ်": r.vehicleInfo || '-',
      "တည်းခိုလိပ်စာ": r.address || '-',
      "ရက်စွဲ": formatDateTimeToDDMMYYYY(r.timestamp)
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "နိုင်ငံခြားသား အဝင်မှတ်တမ်း");
    XLSX.writeFile(wb, `Foreigner_Inward_${selectedDate}.xlsx`);
  };

  const exportOutwardToExcel = () => {
    const data = foreignerOutRecords.map((r, idx) => {
      const lastIn = records.filter(rec => 
        (rec.logType === 'FFE' || !rec.logType) &&
        rec.passport === r.passport && 
        rec.mode === 'IN' && 
        parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)
      ).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
      const stayedFrom = lastIn ? lastIn.timestamp.split(', ')[0] : '-';

      return {
        "စဉ်": idx + 1,
        "အမည်": r.fullname,
        "လိင်": r.gender,
        "နိုင်ငံသား": r.nationality || '-',
        "နိုင်ငံကူးလက်မှတ်": r.passport,
        "ပြည်ဝင်ဗီဇာအမျိုးအစား": r.visaType || '-',
        "နေခွင့်ကာလ": (r.stayFrom || r.stayTo) ? `${formatToDDMMYYYY(r.stayFrom)} to ${formatToDDMMYYYY(r.stayTo)}` : '-',
        "တည်းခိုမည့်နေ့ရက်(မှ)": stayedFrom !== '-' ? formatToDDMMYYYY(stayedFrom) : '-',
        "တည်းခိုမည့်နေ့ရက်(ထိ)": formatToDDMMYYYY(r.timestamp),
        "ရောက်ရှိသည့်နေရာ": r.arrivedFrom || '-',
        "ထွက်ခွာသည့်နေရာ": r.departedTo || '-',
        "ယာဉ်အမှတ်": r.vehicleInfo || '-',
        "တည်းခိုလိပ်စာ": r.address || '-',
        "ရက်စွဲ": formatDateTimeToDDMMYYYY(r.timestamp)
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "နိုင်ငံခြားသား အထွက်မှတ်တမ်း");
    XLSX.writeFile(wb, `Foreigner_Outward_${selectedDate}.xlsx`);
  };

  const exportCompanyToExcel = () => {
    const data = companyResidentsList.map((m, idx) => ({
      "စဉ်": idx + 1,
      "နိုင်ငံကူးလက်မှတ်အမှတ်": m.p,
      "နိုင်ငံအမည်": m.nat || '-',
      "အမည်": m.n,
      "ကျား/မ": m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-',
      "ဗီဇာအမျိုးအစား": m.visa || '-',
      "ဗီဇာသက်တမ်း(မှ/ထိ)": (m.start || m.end) ? `${formatToDDMMYYYY(m.start)} မှ ${formatToDDMMYYYY(m.end)}` : '-',
      "နောက်ဆုံးရောက်ရှိရက်စွဲ": m.in ? formatToDDMMYYYY(m.in) : '-',
      "နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း": getElapsedDays(m, selectedDate),
      "တည်းခိုလိပ်စာ (ကုမ္မဏီအမည်)": m.loc || '-',
      "စစ်ဆေးမှုမှ နောက်ဆုံးမှတ်ချက်": getLatestVerificationRemark(m.p, dossierHistory, records)
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ကုမ္မဏီလိပ်စာ နေထိုင်သူများ");
    XLSX.writeFile(wb, `Company_Residents_${selectedDate}.xlsx`);
  };

  const exportOtherToExcel = () => {
    const data = otherResidentsList.map((m, idx) => ({
      "စဉ်": idx + 1,
      "နိုင်ငံကူးလက်မှတ်အမှတ်": m.p,
      "နိုင်ငံအမည်": m.nat || '-',
      "အမည်": m.n,
      "ကျား/မ": m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-',
      "ဗီဇာအမျိုးအစား": m.visa || '-',
      "ဗီဇာသက်တမ်း(မှ/ထိ)": (m.start || m.end) ? `${formatToDDMMYYYY(m.start)} မှ ${formatToDDMMYYYY(m.end)}` : '-',
      "နောက်ဆုံးရောက်ရှိရက်စွဲ": m.in ? formatToDDMMYYYY(m.in) : '-',
      "နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း": getElapsedDays(m, selectedDate),
      "တည်းခိုလိပ်စာ": m.loc || '-',
      "စစ်ဆေးမှုမှ နောက်ဆုံးမှတ်ချက်": getLatestVerificationRemark(m.p, dossierHistory, records)
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "အခြားလိပ်စာ နေထိုင်သူများ");
    XLSX.writeFile(wb, `Other_Residents_${selectedDate}.xlsx`);
  };

  const exportAddressSummaryToExcel = () => {
    const data = addressSummaryList.map((item, idx) => ({
      "စဉ်": idx + 1,
      "တည်းခိုလိပ်စာ": item.address,
      "ကျား": item.male,
      "မ": item.female,
      "ပေါင်း": item.total
    }));
    data.push({
      "စဉ်": "စုစုပေါင်း" as any,
      "တည်းခိုလိပ်စာ": "",
      "ကျား": totalAddressSummary.male,
      "မ": totalAddressSummary.female,
      "ပေါင်း": totalAddressSummary.total
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "လိပ်စာအလိုက် စာရင်းချုပ်");
    XLSX.writeFile(wb, `Address_Summary_${selectedDate}.xlsx`);
  };

  const exportAllToPhotos = async () => {
    showToast("ဇယား (၆) ခုလုံးအား ပုံအဖြစ် သိမ်းဆည်းနေပါသည်...");
    try {
      await exportElementToPhoto(telegraphRef, `1_telegraph_report_${selectedDate}`);
      await new Promise(r => setTimeout(r, 500));
      await exportElementToPhoto(inwardRef, `2_foreigner_inward_report_${selectedDate}`);
      await new Promise(r => setTimeout(r, 500));
      await exportElementToPhoto(outwardRef, `3_foreigner_outward_report_${selectedDate}`);
      await new Promise(r => setTimeout(r, 500));
      await exportElementToPhoto(companyRef, `4_company_residents_${selectedDate}`);
      await new Promise(r => setTimeout(r, 500));
      await exportElementToPhoto(otherRef, `5_other_residents_${selectedDate}`);
      await new Promise(r => setTimeout(r, 500));
      await exportElementToPhoto(addressSummaryRef, `6_address_summary_${selectedDate}`);
      showToast("ဇယား (၆) ခုလုံး ပုံ (PNG) အဖြစ် အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။");
    } catch (e) {
      console.error(e);
      showToast("ဇယားအားလုံး ပုံသိမ်းဆည်းရာတွင် အမှားအယွင်းရှိပါသည်");
    }
  };

  useEffect(() => {
    const recordsForDate = records.filter(r => isRecordOnDate(r, selectedDate));
    const aggregated = aggregateTelegraphRows(recordsForDate);
    setRows(aggregated);
  }, [selectedDate, records, summaries]);

  const formatNum = (val: number | string) => {
    return useBurmeseDigits ? toBurmeseDigits(val) : val;
  };

  const handleAddRow = () => {
    const newRow: TelegraphRow = {
      id: `custom_${Date.now()}`,
      flightName: 'MNAUB365',
      time: '1200',
      myanmarInM: 0,
      myanmarInF: 0,
      myanmarOutM: 0,
      myanmarOutF: 0,
      foreignerInM: 0,
      foreignerInF: 0,
      foreignerOutM: 0,
      foreignerOutF: 0,
      remarks: 'MGZ-YGN-MGZ'
    };
    setRows(prev => [...prev, newRow]);
    showToast('Manual row added to Telegraph.');
  };

  const handleDeleteRow = (id: string) => {
    setRows(prev => prev.filter(r => r.id !== id));
  };

  const handleCellChange = (id: string, field: keyof TelegraphRow, value: any) => {
    setRows(prev => prev.map(r => {
      if (r.id === id) {
        return { ...r, [field]: value };
      }
      return r;
    }));
  };

  const handleRecalculate = () => {
    const recordsForDate = records.filter(r => isRecordOnDate(r, selectedDate));
    const aggregated = aggregateTelegraphRows(recordsForDate);
    setRows(aggregated);
    showToast('Telegraph data recompiled from the core database.');
  };

  const handlePrint = () => {
    if (activeSubTab === 'TELEGRAPH') {
      triggerDirectPrint(telegraphRef.current, false);
    } else if (activeSubTab === 'FOREIGNER_IN') {
      triggerDirectPrint(inwardRef.current, true);
    } else if (activeSubTab === 'FOREIGNER_OUT') {
      triggerDirectPrint(outwardRef.current, true);
    } else if (activeSubTab === 'COMPANY_RESIDENTS') {
      triggerDirectPrint(companyRef.current, true);
    } else if (activeSubTab === 'OTHER_RESIDENTS') {
      triggerDirectPrint(otherRef.current, true);
    } else if (activeSubTab === 'ADDRESS_SUMMARY') {
      triggerDirectPrint(addressSummaryRef.current, false);
    } else {
      triggerDirectPrint(telegraphRef.current, false);
    }
  };

  // Calculations for display totals
  const totalMMInM = rows.reduce((acc, row) => acc + (Number(row.myanmarInM) || 0), 0);
  const totalMMInF = rows.reduce((acc, row) => acc + (Number(row.myanmarInF) || 0), 0);
  const totalMMOutM = rows.reduce((acc, row) => acc + (Number(row.myanmarOutM) || 0), 0);
  const totalMMOutF = rows.reduce((acc, row) => acc + (Number(row.myanmarOutF) || 0), 0);
  const totalFRNInM = rows.reduce((acc, row) => acc + (Number(row.foreignerInM) || 0), 0);
  const totalFRNInF = rows.reduce((acc, row) => acc + (Number(row.foreignerInF) || 0), 0);
  const totalFRNOutM = rows.reduce((acc, row) => acc + (Number(row.foreignerOutM) || 0), 0);
  const totalFRNOutF = rows.reduce((acc, row) => acc + (Number(row.foreignerOutF) || 0), 0);

  const allRecords = useMemo(() => {
    return [...(records || []), ...(tempRecords || [])];
  }, [records, tempRecords]);

  const foreignerInRecords = useMemo(() => {
    return allRecords.filter(r => isRecordOnDate(r, selectedDate) && r.mode === 'IN' && (r.logType === 'FFE' || !r.logType));
  }, [allRecords, selectedDate]);

  const foreignerOutRecords = useMemo(() => {
    return allRecords.filter(r => isRecordOnDate(r, selectedDate) && r.mode === 'OUT' && (r.logType === 'FFE' || !r.logType));
  }, [allRecords, selectedDate]);

  return (
    <div className="max-w-[1600px] mx-auto py-6 px-3 sm:px-6 space-y-6 select-text">
      {/* Sub-tab Navigation & Controls */}
      <div className="bg-white rounded-2xl shadow-xl border border-slate-100 p-3 sm:p-4 flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
        <div className="flex items-center gap-1.5 sm:gap-2 overflow-x-auto pb-1 max-w-full flex-wrap">
          <button
            onClick={() => setActiveSubTab('ALL')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'ALL'
                ? 'bg-slate-900 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <Layers size={15} />
            <span>📑 ဇယားအားလုံး (All Tables)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('TELEGRAPH')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'TELEGRAPH'
                ? 'bg-slate-900 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <FileText size={15} />
            <span>၁။ ကြေးနန်းစာအစီရင်ခံစာ</span>
          </button>
          <button
            onClick={() => setActiveSubTab('FOREIGNER_IN')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'FOREIGNER_IN'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <Table size={15} />
            <span>၂။ နိုင်ငံခြားသား (အဝင်)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('FOREIGNER_OUT')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'FOREIGNER_OUT'
                ? 'bg-rose-600 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <Table size={15} />
            <span>၃။ နိုင်ငံခြားသား (အထွက်)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('COMPANY_RESIDENTS')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'COMPANY_RESIDENTS'
                ? 'bg-blue-600 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <Building2 size={15} />
            <span>၄။ ကုမ္ပဏီလိပ်စာ နေထိုင်သူများ</span>
          </button>
          <button
            onClick={() => setActiveSubTab('OTHER_RESIDENTS')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'OTHER_RESIDENTS'
                ? 'bg-amber-600 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <Home size={15} />
            <span>၅။ အခြားလိပ်စာ နေထိုင်သူများ</span>
          </button>
          <button
            onClick={() => setActiveSubTab('ADDRESS_SUMMARY')}
            className={`px-3 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shrink-0 ${
              activeSubTab === 'ADDRESS_SUMMARY'
                ? 'bg-purple-600 text-white shadow-md'
                : 'text-slate-600 hover:bg-slate-100 bg-slate-50 border border-slate-200'
            }`}
          >
            <PieChart size={15} />
            <span>၆။ လိပ်စာအလိုက် စာရင်းချုပ်</span>
          </button>
        </div>

        <div className="flex flex-wrap items-center justify-between lg:justify-end gap-2 shrink-0">
          {/* Auto-Fit Toggle & Zoom Controller */}
          <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl p-1 shadow-sm">
            <button
              onClick={() => {
                setIsAutoFit(!isAutoFit);
                if (!isAutoFit) setZoomLevel(1);
              }}
              className={`px-2.5 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                isAutoFit
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
              }`}
              title={isAutoFit ? "Auto-Fit ဖွင့်ထားပါသည် (ဇယားအားလုံး စခရင်နှင့် အံဝင်ခွင်ကျဖြစ်စေသည်)" : "Auto-Fit ပိတ်ထားပါသည် (မူရင်းအရွယ်အစား)"}
            >
              <Maximize2 size={13} />
              <span>{isAutoFit ? 'Auto-Fit: ON' : 'Auto-Fit: OFF'}</span>
            </button>
            <div className="flex items-center gap-0.5 px-1">
              <button
                onClick={() => {
                  setZoomLevel(prev => Math.max(0.4, Number((prev - 0.1).toFixed(2))));
                  setIsAutoFit(false);
                }}
                className="p-1 hover:bg-slate-200 rounded text-slate-700 font-bold transition-all cursor-pointer"
                title="Zoom Out"
              >
                <ZoomOut size={13} />
              </button>
              <span className="text-[11px] font-bold text-slate-600 min-w-[34px] text-center">
                {isAutoFit ? 'Auto' : `${Math.round(zoomLevel * 100)}%`}
              </span>
              <button
                onClick={() => {
                  setZoomLevel(prev => Math.min(1.5, Number((prev + 0.1).toFixed(2))));
                  setIsAutoFit(false);
                }}
                className="p-1 hover:bg-slate-200 rounded text-slate-700 font-bold transition-all cursor-pointer"
                title="Zoom In"
              >
                <ZoomIn size={13} />
              </button>
            </div>
          </div>

          {/* Global Batch Export Options */}
          <button
            onClick={exportAllToPhotos}
            title="ဇယား (၆) ခုလုံးအား တစ်ခါတည်း Photo (PNG) အဖြစ် သိမ်းမည်"
            className="px-3 py-2 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-sm transition-all flex items-center gap-1.5 cursor-pointer"
          >
            <Download size={14} />
            <span>Save ALL Photos</span>
          </button>

          <div className="flex items-center bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 shadow-sm">
            <span className="text-slate-400 mr-2 text-xs font-bold uppercase font-sans">Date</span>
            <input 
              type="date" 
              value={selectedDate}
              onChange={(e) => setSelectedDate(e.target.value)}
              className="bg-transparent border-none text-xs font-bold text-slate-700 outline-none focus:ring-0"
            />
          </div>
          <button
            onClick={() => setUseBurmeseDigits(!useBurmeseDigits)}
            className={`px-3 py-2 text-xs font-bold rounded-xl border transition-all ${
              useBurmeseDigits 
                ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' 
                : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
            }`}
          >
            <span>{useBurmeseDigits ? 'မြန်မာဂဏန်း' : 'အင်္ဂလိပ်ဂဏန်း'}</span>
          </button>
        </div>
      </div>

      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'TELEGRAPH' ? 'block' : 'hidden'}`}>
        <div className="space-y-6">
          {/* Accordion Config Panel */}
      <div className="bg-white rounded-2xl shadow-lg border border-slate-100 overflow-hidden">
        <button
          onClick={() => setIsHeaderConfigOpen(!isHeaderConfigOpen)}
          className="w-full flex items-center justify-between px-6 py-4 bg-slate-50/50 hover:bg-slate-50 transition-colors"
        >
          <div className="flex items-center gap-2 text-slate-700">
            <span className="text-sm font-black uppercase tracking-wide">Configure Header & Metadata</span>
            <span className="text-[10px] px-2 py-0.5 bg-indigo-50 text-indigo-600 font-bold rounded-full">Editable</span>
          </div>
          {isHeaderConfigOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
        </button>

        {isHeaderConfigOpen && (
          <div className="p-6 border-t border-slate-100 grid grid-cols-1 md:grid-cols-3 gap-6 bg-slate-50/20">
            {/* Line 1 Chiefs */}
            <div className="space-y-4">
              <h3 className="text-xs font-black uppercase text-slate-600 border-b pb-1">Header Chief Officers</h3>
              <div className="space-y-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">(ပ) Chief 1</label>
                  <input 
                    type="text" 
                    value={chief1}
                    onChange={(e) => handleChief1Change(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">(လ) Chief 2</label>
                  <input 
                    type="text" 
                    value={chief2}
                    onChange={(e) => handleChief2Change(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">(တ) Chief 3</label>
                  <input 
                    type="text" 
                    value={chief3}
                    onChange={(e) => handleChief3Change(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
              </div>
            </div>

            {/* Letter info */}
            <div className="space-y-4">
              <h3 className="text-xs font-black uppercase text-slate-600 border-b pb-1">Letter Details</h3>
              <div className="space-y-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Letter Number (စာအမှတ်)</label>
                  <input 
                    type="text" 
                    value={letterNo}
                    onChange={(e) => handleLetterNoChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Urgency Level (အဆင့်အတန်း)</label>
                  <input 
                    type="text" 
                    value={urgentLevel}
                    onChange={(e) => handleUrgentLevelChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Telegraph Title (ခေါင်းစဉ်)</label>
                  <input 
                    type="text" 
                    value={telegraphTitle}
                    onChange={(e) => handleTelegraphTitleChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500 font-bold text-black"
                  />
                </div>
              </div>
            </div>

            {/* Sign and Intro */}
            <div className="space-y-4">
              <h3 className="text-xs font-black uppercase text-slate-600 border-b pb-1">Signature & Introduction</h3>
              <div className="space-y-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Officer Title (ရာထူး)</label>
                  <input 
                    type="text" 
                    value={officerTitle}
                    onChange={(e) => handleOfficerTitleChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Officer Office (ဌာနခွဲ)</label>
                  <input 
                    type="text" 
                    value={officerName}
                    onChange={(e) => handleOfficerNameChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Custom Intro Text Paragraph</label>
                  <textarea 
                    rows={2}
                    value={customIntroText}
                    onChange={(e) => handleCustomIntroTextChange(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-1 text-[11px] font-medium focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Realistic Document Canvas Live Preview with Mobile Fit Controller */}
      <div className="space-y-3">
        <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
          <div className="flex items-center gap-2">
            <span className="text-[10px] uppercase font-black tracking-wider bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 px-2 py-0.5 rounded-md">
              A4 Portrait
            </span>
            <span className="text-xs sm:text-sm font-black text-slate-100">
              ၁။ ကြေးနန်းစာအစီရင်ခံစာ ပုံစံ
            </span>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={exportTelegraphToPhoto}
              className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
            >
              <Download size={13} />
              <span>Save as Photo</span>
            </button>
            <button
              onClick={handlePrint}
              className="px-3 py-1.5 text-xs font-bold rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
            >
              <Printer size={13} />
              <span>Print Report</span>
            </button>
          </div>
        </div>
        <AutoFitSheetWrapper isLandscape={false} autoFit={isAutoFit} zoom={zoomLevel}>
        <div 
          ref={telegraphRef}
          className="p-[0.5in] shadow-2xl rounded-sm w-[210mm] min-w-[210mm] border min-h-[297mm] relative font-pyidaungsu leading-relaxed text-[13px] select-text box-border"
          style={{ backgroundColor: '#ffffff', color: '#000000', borderColor: '#cbd5e1' }}
        >
          {/* Absolute Top Title */}
          <div className="text-center mb-4">
            <span className="text-[14px] font-black block text-center font-pyidaungsu" style={{ color: '#000000', letterSpacing: 'normal', fontWeight: 'bold' }}>
              မှတ်ပုံတင်ကြေးနန်းစာပုံစံအစား
            </span>
          </div>

          {/* Letter Header */}
          <div className="flex justify-between items-start text-[13px] font-bold pb-2">
            <div className="space-y-1">
              <div>(ပ) {chief1}</div>
              <div>(လ) {chief2}</div>
              <div>(တ) {chief3}</div>
            </div>
            <div className="text-center space-y-0.5">
              <div>ရက်စွဲ / အချိန်</div>
              <div className="font-light leading-none -mt-1" style={{ color: '#94a3b8' }}>—————-</div>
              <div className="font-bold">
                {toBurmeseShortDateTime(selectedDate, reportTime)}
              </div>
            </div>
          </div>

          {/* Opening Line above title */}
          <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

          {/* Letter Info Line with Side-by-Side Letter No and Telegraph Title with one tab space separation */}
          <div className="flex items-center gap-12 text-[13px] font-bold py-1" style={{ color: '#000000' }}>
            <div>{letterNo}</div>
            <div className="font-black text-[13px]">{telegraphTitle}</div>
          </div>

          {/* Paragraph 1 */}
          <p className="text-[13px] leading-relaxed font-medium whitespace-pre-line my-4" style={{ color: '#000000' }}>
            {customIntroText
              .replace('[DATE]', toBurmeseSlashDate(selectedDate))
              .replace('[FLIGHT_COUNT]', toBurmeseDigits(rows.length))
            }
          </p>

          {/* Table */}
          <table className="w-full text-center text-[13px] border-collapse font-pyidaungsu my-4" style={{ border: '1px solid #000000', borderColor: '#000000', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ color: '#000000' }}>
                <th rowSpan={3} className="p-1 text-center align-middle w-10 font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>စဉ်</th>
                <th rowSpan={3} className="p-1 text-center align-middle w-[110px] min-w-[110px] max-w-[110px] font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>လေကြောင်းလိုင်းအမည်</th>
                <th rowSpan={3} className="p-1 text-center align-middle w-16 font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>ဝင်/ထွက်ချိန်</th>
                <th colSpan={6} className="p-1 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မြန်မာနိုင်ငံသား</th>
                <th colSpan={6} className="p-1 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>နိုင်ငံခြားသား</th>
                <th rowSpan={3} className="p-1 text-center align-middle font-black" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မှတ်ချက်</th>
              </tr>
              <tr style={{ color: '#000000' }}>
                <th colSpan={3} className="p-1 text-center text-[13px] font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>အဝင်</th>
                <th colSpan={3} className="p-1 text-center text-[13px] font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>အထွက်</th>
                <th colSpan={3} className="p-1 text-center text-[13px] font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>အဝင်</th>
                <th colSpan={3} className="p-1 text-center text-[13px] font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>အထွက်</th>
              </tr>
              <tr className="text-[13px] font-bold" style={{ color: '#000000' }}>
                {/* MMR In */}
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>က</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မ</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f1f5f9' }}>ပ</th>
                {/* MMR Out */}
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>က</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မ</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f1f5f9' }}>ပ</th>
                {/* FRN In */}
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>က</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မ</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f1f5f9' }}>ပ</th>
                {/* FRN Out */}
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>က</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>မ</th>
                <th className="p-1 w-8" style={{ border: '1px solid #000000', backgroundColor: '#f1f5f9' }}>ပ</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, idx) => {
                const myanmarInTotal = row.myanmarInM + row.myanmarInF;
                const myanmarOutTotal = row.myanmarOutM + row.myanmarOutF;
                const foreignerInTotal = row.foreignerInM + row.foreignerInF;
                const foreignerOutTotal = row.foreignerOutM + row.foreignerOutF;

                return (
                  <tr key={row.id}>
                    <td className="p-1.5 text-center font-bold" style={{ border: '1px solid #000000' }}>{formatNum(idx + 1)}</td>
                    <td className="p-1.5 font-extrabold text-left uppercase w-[110px] min-w-[110px] max-w-[110px] break-words whitespace-normal text-wrap leading-tight text-[12px]" style={{ border: '1px solid #000000' }}>{row.flightName}</td>
                    <td className="p-1.5 text-center font-mono" style={{ border: '1px solid #000000' }}>{formatNum(row.time)}</td>
                    
                    {/* MMR In */}
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.myanmarInM)}</td>
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.myanmarInF)}</td>
                    <td className="p-1.5 text-center font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>{formatNum(myanmarInTotal)}</td>
                    
                    {/* MMR Out */}
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.myanmarOutM)}</td>
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.myanmarOutF)}</td>
                    <td className="p-1.5 text-center font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>{formatNum(myanmarOutTotal)}</td>
                    
                    {/* FRN In */}
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.foreignerInM)}</td>
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.foreignerInF)}</td>
                    <td className="p-1.5 text-center font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>{formatNum(foreignerInTotal)}</td>
                    
                    {/* FRN Out */}
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.foreignerOutM)}</td>
                    <td className="p-1.5 text-center" style={{ border: '1px solid #000000' }}>{formatNum(row.foreignerOutF)}</td>
                    <td className="p-1.5 text-center font-bold" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}>{formatNum(foreignerOutTotal)}</td>
                    
                    <td className="p-1.5 text-left text-[13px] font-bold" style={{ border: '1px solid #000000', color: '#334155' }}>{row.remarks}</td>
                  </tr>
                );
              })}

              {/* Totals Row */}
              {rows.length > 0 && (
                <tr className="font-bold text-[13px]" style={{ backgroundColor: '#f1f5f9', color: '#000000' }}>
                  <td colSpan={3} className="p-2 text-center font-black" style={{ border: '1px solid #000000' }}>Total</td>
                  
                  {/* MMR In Total */}
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalMMInM)}</td>
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalMMInF)}</td>
                  <td className="p-2 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#e2e8f0' }}>{formatNum(totalMMInM + totalMMInF)}</td>
                  
                  {/* MMR Out Total */}
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalMMOutM)}</td>
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalMMOutF)}</td>
                  <td className="p-2 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#e2e8f0' }}>{formatNum(totalMMOutM + totalMMOutF)}</td>
                  
                  {/* FRN In Total */}
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalFRNInM)}</td>
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalFRNInF)}</td>
                  <td className="p-2 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#e2e8f0' }}>{formatNum(totalFRNInM + totalFRNInF)}</td>
                  
                  {/* FRN Out Total */}
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalFRNOutM)}</td>
                  <td className="p-2 text-center" style={{ border: '1px solid #000000' }}>{formatNum(totalFRNOutF)}</td>
                  <td className="p-2 text-center font-black" style={{ border: '1px solid #000000', backgroundColor: '#e2e8f0' }}>{formatNum(totalFRNOutM + totalFRNOutF)}</td>
                  
                  <td className="p-2" style={{ border: '1px solid #000000', backgroundColor: '#f8fafc' }}></td>
                </tr>
              )}
            </tbody>
          </table>

          {/* Paragraph 2 - (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... / */}
          <div className="text-[13px] font-normal mt-4" style={{ color: '#000000' }}>
            (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... /
          </div>

          {/* ပိတ်မျဉ်း အပြည့်တစ်ကြောင်း below သိရှိနိုင်ပါရန်တင်ပြအပ် */}
          <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

          {/* Footer Text */}
          <div className="flex justify-between items-start text-[13px] font-bold leading-relaxed" style={{ color: '#000000' }}>
            <div className="space-y-1">
              <div className="font-medium">သင့်ရာနည်းဖြင့်</div>
            </div>
            <div className="text-center space-y-0.5 pr-8">
              <div>အဆင့်အတန်း</div>
              <div className="font-light leading-none -mt-1" style={{ color: '#94a3b8' }}>—————</div>
              <div className="font-black" style={{ color: '#be123c' }}>{urgentLevel}</div>
            </div>
          </div>

          {/* Signature block */}
          <div className="flex justify-start mt-10 pl-4">
            <div className="text-center space-y-1 font-pyidaungsu max-w-sm flex flex-col items-center">
              <div className="text-[13px] font-extrabold" style={{ color: '#000000' }}>
                {officerTitle}
              </div>
              <div className="text-[13px] font-bold leading-tight" style={{ color: '#000000' }}>
                {officerName}
              </div>
            </div>
          </div>
        </div>
        </AutoFitSheetWrapper>
      </div>
      </div>
    </div>

      {/* Inbound Foreigner Daily Record View */}
      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'FOREIGNER_IN' ? 'block' : 'hidden'}`}>
        <div className="space-y-3">
          <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-black tracking-wider bg-rose-500/20 text-rose-300 border border-rose-500/30 px-2 py-0.5 rounded-md">
                A4 Landscape
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-100">
                ၂။ နိုင်ငံခြားသားများ နေ့စဉ် (ဝင်ရောက်)မှတ်တမ်း
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={exportInwardToExcel}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <FileSpreadsheet size={13} />
                <span>Export Excel</span>
              </button>
              <button
                onClick={exportInwardToPhoto}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Download size={13} />
                <span>Save as Photo</span>
              </button>
              <button
                onClick={() => triggerDirectPrint(inwardRef.current, true)}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Printer size={13} />
                <span>Print Table</span>
              </button>
            </div>
          </div>
          <AutoFitSheetWrapper isLandscape={true} autoFit={isAutoFit} zoom={zoomLevel}>
          <div
            ref={inwardRef}
            className="bg-white text-black p-[0.5in] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] border border-slate-300 min-h-[210mm] relative font-pyidaungsu text-[12px] select-text box-border"
          >
            <div className="text-center font-bold text-[14px] mb-1 text-black font-pyidaungsu">
              တနင်္သာရီတိုင်းဒေသကြီး၊ မြိတ်ခရိုင်အတွင်း နိုင်ငံခြားသားများ နေ့စဉ် (ဝင်ရောက်)မှတ်တမ်း
            </div>
            <div className="text-right font-bold text-[12px] mb-4 text-black font-pyidaungsu pr-1">
              ရက်စွဲ၊ {toBurmeseSlashDate(selectedDate)}
            </div>

            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
              <colgroup>
                <col className="w-[30px]" />
                <col className="w-[125px]" />
                <col className="w-[30px]" />
                <col className="w-[55px]" />
                <col className="w-[85px]" />
                <col className="w-[85px]" />
                <col className="w-[70px]" />
                <col className="w-[60px]" />
                <col className="w-[60px]" />
                <col className="w-[80px]" />
                <col className="w-[80px]" />
                <col className="w-[70px]" />
                <col className="w-[125px]" />
                <col className="w-[70px]" />
              </colgroup>
              <thead>
                <tr className="border-b border-black align-middle text-center">
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>စဉ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>အမည်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>လိင်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နိုင်ငံသား</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နိုင်ငံကူး<br/>လက်မှတ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ပြည်ဝင်ဗီဇာ<br/>အမျိုးအစား</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နေခွင့်ကာလ</th>
                  <th colSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>တည်းခိုမည့်<br/>နေ့ရက်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ရောက်ရှိသည့်<br/>နေရာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ထွက်ခွာသည့်<br/>နေရာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ယာဉ်အမှတ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>တည်းခိုလိပ်စာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ရက်စွဲ</th>
                </tr>
                <tr className="border-b border-black align-middle text-center">
                  <th className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>မှ</th>
                  <th className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ထိ</th>
                </tr>
              </thead>
              <tbody>
                {foreignerInRecords.length === 0 ? (
                  <tr>
                    <td colSpan={14} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                      ရွေးချယ်ထားသော ရက်စွဲတွင် နိုင်ငံခြားသား အဝင်မှတ်တမ်း မရှိပါ။
                    </td>
                  </tr>
                ) : (
                  foreignerInRecords.map((r, idx) => (
                    <tr key={r.id} className="border-b border-black">
                      <td className="p-1 border border-black text-center font-bold">
                        {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                      </td>
                      <td className="p-1 border border-black text-left font-bold break-words">{r.fullname}</td>
                      <td className="p-1 border border-black text-center">{r.gender}</td>
                      <td className="p-1 border border-black text-center font-bold break-words">{r.nationality || '-'}</td>
                      <td className="p-1 border border-black text-center font-bold font-mono break-words">{r.passport}</td>
                      <td className="p-1 border border-black text-center break-words">{r.visaType || '-'}</td>
                      <td className="p-1 border border-black text-center break-words">{(r.stayFrom || r.stayTo) ? `${formatToDDMMYYYY(r.stayFrom)} to ${formatToDDMMYYYY(r.stayTo)}` : '-'}</td>
                      <td className="p-1 border border-black text-center">{formatToDDMMYYYY(r.timestamp)}</td>
                      <td className="p-1 border border-black text-center">{formatToDDMMYYYY(r.stayTo)}</td>
                      <td className="p-1 border border-black text-center break-words">{r.arrivedFrom || '-'}</td>
                      <td className="p-1 border border-black text-center break-words">{r.departedTo || '-'}</td>
                      <td className="p-1 border border-black text-center font-bold break-words">{r.vehicleInfo || '-'}</td>
                      <td className="p-1 border border-black text-left break-words">{r.address || '-'}</td>
                      <td className="p-1 border border-black text-center font-mono text-[10px] break-words">{formatDateTimeToDDMMYYYY(r.timestamp)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>

            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px]">
                <div className="h-10" />
                <div className="text-[13px] font-extrabold text-black">{officerTitle}</div>
                <div className="text-[13px] font-bold text-black leading-tight">{officerName}</div>
              </div>
            </div>
          </div>
          </AutoFitSheetWrapper>
        </div>
      </div>

      {/* Outbound Foreigner Daily Record View */}
      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'FOREIGNER_OUT' ? 'block' : 'hidden'}`}>
                <div className="space-y-3">
          <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-black tracking-wider bg-rose-500/20 text-rose-300 border border-rose-500/30 px-2 py-0.5 rounded-md">
                A4 Landscape
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-100">
                ၃။ နိုင်ငံခြားသားများ နေ့စဉ် (ထွက်ခွာ)မှတ်တမ်း
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={exportOutwardToExcel}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <FileSpreadsheet size={13} />
                <span>Export Excel</span>
              </button>
              <button
                onClick={exportOutwardToPhoto}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Download size={13} />
                <span>Save as Photo</span>
              </button>
              <button
                onClick={() => triggerDirectPrint(outwardRef.current, true)}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-rose-600 hover:bg-rose-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Printer size={13} />
                <span>Print Table</span>
              </button>
            </div>
          </div>
          <AutoFitSheetWrapper isLandscape={true} autoFit={isAutoFit} zoom={zoomLevel}>
          <div
            ref={outwardRef}
            className="bg-white text-black p-[0.5in] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] border border-slate-300 min-h-[210mm] relative font-pyidaungsu text-[12px] select-text box-border"
          >
            <div className="text-center font-bold text-[14px] mb-1 text-black font-pyidaungsu">
              တနင်္သာရီတိုင်းဒေသကြီး၊ မြိတ်ခရိုင်အတွင်း နိုင်ငံခြားသားများ နေ့စဉ် (ထွက်ခွာ)မှတ်တမ်း
            </div>
            <div className="text-right font-bold text-[12px] mb-4 text-black font-pyidaungsu pr-1">
              ရက်စွဲ၊ {toBurmeseSlashDate(selectedDate)}
            </div>

            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
              <colgroup>
                <col className="w-[30px]" />
                <col className="w-[125px]" />
                <col className="w-[30px]" />
                <col className="w-[55px]" />
                <col className="w-[85px]" />
                <col className="w-[85px]" />
                <col className="w-[70px]" />
                <col className="w-[60px]" />
                <col className="w-[60px]" />
                <col className="w-[80px]" />
                <col className="w-[80px]" />
                <col className="w-[70px]" />
                <col className="w-[125px]" />
                <col className="w-[70px]" />
              </colgroup>
              <thead>
                <tr className="border-b border-black align-middle text-center">
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>စဉ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>အမည်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>လိင်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နိုင်ငံသား</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နိုင်ငံကူး<br/>လက်မှတ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ပြည်ဝင်ဗီဇာ<br/>အမျိုးအစား</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>နေခွင့်ကာလ</th>
                  <th colSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>တည်းခိုမည့်<br/>နေ့ရက်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ရောက်ရှိသည့်<br/>နေရာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ထွက်ခွာသည့်<br/>နေရာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ယာဉ်အမှတ်</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>တည်းခိုလိပ်စာ</th>
                  <th rowSpan={2} className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ရက်စွဲ</th>
                </tr>
                <tr className="border-b border-black align-middle text-center">
                  <th className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>မှ</th>
                  <th className="bg-slate-50 border border-black p-1 align-middle text-center font-bold text-[11px] leading-snug" style={{ verticalAlign: 'middle', backgroundColor: '#f8fafc', color: '#000000' }}>ထိ</th>
                </tr>
              </thead>
              <tbody>
                {foreignerOutRecords.length === 0 ? (
                  <tr>
                    <td colSpan={14} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                      ရွေးချယ်ထားသော ရက်စွဲတွင် နိုင်ငံခြားသား အထွက်မှတ်တမ်း မရှိပါ။
                    </td>
                  </tr>
                ) : (
                  foreignerOutRecords.map((r, idx) => {
                    const lastIn = records.filter(rec => 
                      (rec.logType === 'FFE' || !rec.logType) &&
                      rec.passport === r.passport && 
                      rec.mode === 'IN' && 
                      parseTimestamp(rec.timestamp) < parseTimestamp(r.timestamp)
                    ).sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
                    const stayedFrom = lastIn ? lastIn.timestamp.split(', ')[0] : '-';

                    return (
                      <tr key={r.id} className="border-b border-black">
                        <td className="p-1 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                        </td>
                        <td className="p-1 border border-black text-left font-bold break-words">{r.fullname}</td>
                        <td className="p-1 border border-black text-center">{r.gender}</td>
                        <td className="p-1 border border-black text-center font-bold break-words">{r.nationality || '-'}</td>
                        <td className="p-1 border border-black text-center font-bold font-mono break-words">{r.passport}</td>
                        <td className="p-1 border border-black text-center break-words">{r.visaType || '-'}</td>
                        <td className="p-1 border border-black text-center break-words">{(r.stayFrom || r.stayTo) ? `${formatToDDMMYYYY(r.stayFrom)} to ${formatToDDMMYYYY(r.stayTo)}` : '-'}</td>
                        <td className="p-1 border border-black text-center">{stayedFrom !== '-' ? formatToDDMMYYYY(stayedFrom) : '-'}</td>
                        <td className="p-1 border border-black text-center">{formatToDDMMYYYY(r.timestamp)}</td>
                        <td className="p-1 border border-black text-center break-words">{r.arrivedFrom || '-'}</td>
                        <td className="p-1 border border-black text-center break-words">{r.departedTo || '-'}</td>
                        <td className="p-1 border border-black text-center font-bold break-words">{r.vehicleInfo || '-'}</td>
                        <td className="p-1 border border-black text-left break-words">{r.address || '-'}</td>
                        <td className="p-1 border border-black text-center font-mono text-[10px] break-words">{formatDateTimeToDDMMYYYY(r.timestamp)}</td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>

            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px]">
                <div className="h-10" />
                <div className="text-[13px] font-extrabold text-black">{officerTitle}</div>
                <div className="text-[13px] font-bold text-black leading-tight">{officerName}</div>
              </div>
            </div>
          </div>
          </AutoFitSheetWrapper>
        </div>
      </div>

      {/* Table 1: Company Address Residents (Landscape A4) */}
      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'COMPANY_RESIDENTS' ? 'block' : 'hidden'}`}>
                <div className="space-y-3">
          <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-black tracking-wider bg-rose-500/20 text-rose-300 border border-rose-500/30 px-2 py-0.5 rounded-md">
                A4 Landscape
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-100">
                ၄။ ကုမ္မဏီလိပ်စာ နေထိုင်သူများ စာရင်း
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={exportCompanyToExcel}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <FileSpreadsheet size={13} />
                <span>Export Excel</span>
              </button>
              <button
                onClick={exportCompanyToPhoto}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Download size={13} />
                <span>Save as Photo</span>
              </button>
              <button
                onClick={() => triggerDirectPrint(companyRef.current, true)}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Printer size={13} />
                <span>Print Table</span>
              </button>
            </div>
          </div>
          <AutoFitSheetWrapper isLandscape={true} autoFit={isAutoFit} zoom={zoomLevel}>
          <div
            ref={companyRef}
            className="bg-white text-black p-[0.5in] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] border border-slate-300 min-h-[210mm] relative font-pyidaungsu text-[12px] select-text box-border"
          >
            <div className="text-center font-bold text-[14px] text-black font-pyidaungsu mb-1">
              လက်ရှိနေထိုင်သူများမှ ကုမ္မဏီလိပ်စာဖြင့်နေထိုင်သူများ
            </div>
            <div className="text-right font-bold text-[12px] text-black font-pyidaungsu mb-3 pr-1">
              ရက်စွဲ၊ {toBurmeseSlashDate(selectedDate)}
            </div>

            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
              <colgroup>
                <col className="w-[30px]" />
                <col className="w-[85px]" />
                <col className="w-[50px]" />
                <col className="w-[115px]" />
                <col className="w-[40px]" />
                <col className="w-[80px]" />
                <col className="w-[105px]" />
                <col className="w-[80px]" />
                <col className="w-[65px]" />
                <col className="w-[140px]" />
                <col className="w-[125px]" />
              </colgroup>
              <thead>
                <tr className="border-b border-black align-middle text-center bg-slate-50">
                  <th className="border border-black p-1.5 font-bold text-[11px]">စဉ်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နိုင်ငံကူးလက်မှတ်အမှတ်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နိုင်ငံအမည်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">အမည်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ကျား/မ</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ဗီဇာအမျိုးအစား</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ဗီဇာသက်တမ်း (မှ/ထိ)</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နောက်ဆုံးရောက်ရှိရက်စွဲ</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">တည်းခိုလိပ်စာ (ကုမ္မဏီအမည်)</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">စစ်ဆေးမှုမှ နောက်ဆုံးမှတ်ချက်</th>
                </tr>
              </thead>
              <tbody>
                {companyResidentsList.length === 0 ? (
                  <tr>
                    <td colSpan={11} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                      ကုမ္မဏီလိပ်စာဖြင့် နေထိုင်သူ စာရင်းမရှိပါ။
                    </td>
                  </tr>
                ) : (
                  companyResidentsList.map((m, idx) => {
                    const elapsed = getElapsedDays(m, selectedDate);
                    const remark = getLatestVerificationRemark(m.p, dossierHistory, records);
                    return (
                      <tr key={m.p} className="border-b border-black">
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                        </td>
                        <td className="p-1.5 border border-black text-center font-bold font-mono break-words">{m.p}</td>
                        <td className="p-1.5 border border-black text-center font-bold break-words">{m.nat || '-'}</td>
                        <td className="p-1.5 border border-black text-left font-bold break-words">{m.n}</td>
                        <td className="p-1.5 border border-black text-center">{m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-'}</td>
                        <td className="p-1.5 border border-black text-center break-words">{m.visa || '-'}</td>
                        <td className="p-1.5 border border-black text-center break-words">
                          {(m.start || m.end) ? `${formatToDDMMYYYY(m.start)} မှ ${formatToDDMMYYYY(m.end)}` : '-'}
                        </td>
                        <td className="p-1.5 border border-black text-center">{m.in ? formatToDDMMYYYY(m.in) : '-'}</td>
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(elapsed) : elapsed}
                        </td>
                        <td className="p-1.5 border border-black text-left font-bold break-words">{m.loc || '-'}</td>
                        <td className="p-1.5 border border-black text-left break-words text-[10px]">{remark}</td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>

            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px]">
                <div className="h-10" />
                <div className="text-[13px] font-extrabold text-black">{officerTitle}</div>
                <div className="text-[13px] font-bold text-black leading-tight">{officerName}</div>
              </div>
            </div>
          </div>
          </AutoFitSheetWrapper>
        </div>
      </div>

      {/* Table 2: Other Address Residents (Landscape A4) */}
      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'OTHER_RESIDENTS' ? 'block' : 'hidden'}`}>
                <div className="space-y-3">
          <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-black tracking-wider bg-rose-500/20 text-rose-300 border border-rose-500/30 px-2 py-0.5 rounded-md">
                A4 Landscape
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-100">
                ၅။ အခြားလိပ်စာ နေထိုင်သူများ စာရင်း
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={exportOtherToExcel}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <FileSpreadsheet size={13} />
                <span>Export Excel</span>
              </button>
              <button
                onClick={exportOtherToPhoto}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Download size={13} />
                <span>Save as Photo</span>
              </button>
              <button
                onClick={() => triggerDirectPrint(otherRef.current, true)}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-amber-600 hover:bg-amber-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Printer size={13} />
                <span>Print Table</span>
              </button>
            </div>
          </div>
          <AutoFitSheetWrapper isLandscape={true} autoFit={isAutoFit} zoom={zoomLevel}>
          <div
            ref={otherRef}
            className="bg-white text-black p-[0.5in] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] border border-slate-300 min-h-[210mm] relative font-pyidaungsu text-[12px] select-text box-border"
          >
            <div className="text-center font-bold text-[14px] text-black font-pyidaungsu mb-1">
              လက်ရှိနေထိုင်သူများမှ အခြားလိပ်စာဖြင့်နေထိုင်သူများ
            </div>
            <div className="text-right font-bold text-[12px] text-black font-pyidaungsu mb-3 pr-1">
              ရက်စွဲ၊ {toBurmeseSlashDate(selectedDate)}
            </div>

            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
              <colgroup>
                <col className="w-[30px]" />
                <col className="w-[85px]" />
                <col className="w-[50px]" />
                <col className="w-[115px]" />
                <col className="w-[40px]" />
                <col className="w-[80px]" />
                <col className="w-[105px]" />
                <col className="w-[80px]" />
                <col className="w-[65px]" />
                <col className="w-[140px]" />
                <col className="w-[125px]" />
              </colgroup>
              <thead>
                <tr className="border-b border-black align-middle text-center bg-slate-50">
                  <th className="border border-black p-1.5 font-bold text-[11px]">စဉ်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နိုင်ငံကူးလက်မှတ်အမှတ်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နိုင်ငံအမည်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">အမည်</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ကျား/မ</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ဗီဇာအမျိုးအစား</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">ဗီဇာသက်တမ်း (မှ/ထိ)</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နောက်ဆုံးရောက်ရှိရက်စွဲ</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">တည်းခိုလိပ်စာ</th>
                  <th className="border border-black p-1.5 font-bold text-[11px]">စစ်ဆေးမှုမှ နောက်ဆုံးမှတ်ချက်</th>
                </tr>
              </thead>
              <tbody>
                {otherResidentsList.length === 0 ? (
                  <tr>
                    <td colSpan={11} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                      အခြားလိပ်စာဖြင့် နေထိုင်သူ စာရင်းမရှိပါ။
                    </td>
                  </tr>
                ) : (
                  otherResidentsList.map((m, idx) => {
                    const elapsed = getElapsedDays(m, selectedDate);
                    const remark = getLatestVerificationRemark(m.p, dossierHistory, records);
                    return (
                      <tr key={m.p} className="border-b border-black">
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                        </td>
                        <td className="p-1.5 border border-black text-center font-bold font-mono break-words">{m.p}</td>
                        <td className="p-1.5 border border-black text-center font-bold break-words">{m.nat || '-'}</td>
                        <td className="p-1.5 border border-black text-left font-bold break-words">{m.n}</td>
                        <td className="p-1.5 border border-black text-center">{m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-'}</td>
                        <td className="p-1.5 border border-black text-center break-words">{m.visa || '-'}</td>
                        <td className="p-1.5 border border-black text-center break-words">
                          {(m.start || m.end) ? `${formatToDDMMYYYY(m.start)} မှ ${formatToDDMMYYYY(m.end)}` : '-'}
                        </td>
                        <td className="p-1.5 border border-black text-center">{m.in ? formatToDDMMYYYY(m.in) : '-'}</td>
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(elapsed) : elapsed}
                        </td>
                        <td className="p-1.5 border border-black text-left font-bold break-words">{m.loc || '-'}</td>
                        <td className="p-1.5 border border-black text-left break-words text-[10px]">{remark}</td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>

            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px]">
                <div className="h-10" />
                <div className="text-[13px] font-extrabold text-black">{officerTitle}</div>
                <div className="text-[13px] font-bold text-black leading-tight">{officerName}</div>
              </div>
            </div>
          </div>
          </AutoFitSheetWrapper>
        </div>
      </div>

      {/* Table 3: Summary by Address (Portrait A4) */}
      <div className={`w-full ${activeSubTab === 'ALL' || activeSubTab === 'ADDRESS_SUMMARY' ? 'block' : 'hidden'}`}>
                <div className="space-y-3">
          <div className="bg-slate-900 text-white rounded-2xl p-3 sm:p-4 shadow-md flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-black tracking-wider bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 px-2 py-0.5 rounded-md">
                A4 Portrait
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-100">
                ၆။ တည်းခိုလိပ်စာအလိုက် စာရင်းချုပ်
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                onClick={exportAddressSummaryToExcel}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <FileSpreadsheet size={13} />
                <span>Export Excel</span>
              </button>
              <button
                onClick={exportAddressSummaryToPhoto}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Download size={13} />
                <span>Save as Photo</span>
              </button>
              <button
                onClick={() => triggerDirectPrint(addressSummaryRef.current, false)}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-purple-600 hover:bg-purple-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer shrink-0"
              >
                <Printer size={13} />
                <span>Print Table</span>
              </button>
            </div>
          </div>
          <AutoFitSheetWrapper isLandscape={false} autoFit={isAutoFit} zoom={zoomLevel}>
          <div
            ref={addressSummaryRef}
            className="bg-white text-black p-[0.6in] shadow-2xl rounded-sm w-[210mm] min-w-[210mm] border border-slate-300 min-h-[297mm] relative font-pyidaungsu text-[12px] select-text box-border"
          >
            <div className="text-center font-bold text-[15px] text-black font-pyidaungsu mb-1">
              တည်းခိုလိပ်စာအလိုက် နေထိုင်လျက်ရှိသော နိုင်ငံခြားသား စာရင်းချုပ်
            </div>
            <div className="text-right font-bold text-[12px] text-black font-pyidaungsu mb-3 pr-1">
              ရက်စွဲ၊ {toBurmeseSlashDate(selectedDate)}
            </div>

            <table className="w-full text-center text-[12px] border-collapse border border-black font-pyidaungsu table-fixed">
              <colgroup>
                <col className="w-[45px]" />
                <col className="w-[320px]" />
                <col className="w-[80px]" />
                <col className="w-[80px]" />
                <col className="w-[90px]" />
              </colgroup>
              <thead>
                <tr className="border-b border-black align-middle text-center bg-slate-50">
                  <th className="border border-black p-2 font-bold text-[12px]">စဉ်</th>
                  <th className="border border-black p-2 font-bold text-[12px]">တည်းခိုလိပ်စာ / ကုမ္မဏီ / ဟိုတယ်</th>
                  <th className="border border-black p-2 font-bold text-[12px]">ကျား</th>
                  <th className="border border-black p-2 font-bold text-[12px]">မ</th>
                  <th className="border border-black p-2 font-bold text-[12px]">ပေါင်း</th>
                </tr>
              </thead>
              <tbody>
                {addressSummaryList.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                      လက်ရှိနေထိုင်သူ စာရင်းမရှိပါ။
                    </td>
                  </tr>
                ) : (
                  <>
                    {addressSummaryList.map((item, idx) => (
                      <tr key={item.address} className="border-b border-black">
                        <td className="p-2 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                        </td>
                        <td className="p-2 border border-black text-left font-bold break-words">{item.address}</td>
                        <td className="p-2 border border-black text-center">{useBurmeseDigits ? toBurmeseDigits(item.male) : item.male}</td>
                        <td className="p-2 border border-black text-center">{useBurmeseDigits ? toBurmeseDigits(item.female) : item.female}</td>
                        <td className="p-2 border border-black text-center font-bold bg-slate-50">
                          {useBurmeseDigits ? toBurmeseDigits(item.total) : item.total}
                        </td>
                      </tr>
                    ))}
                    <tr className="border-t-2 border-black font-bold bg-slate-100">
                      <td colSpan={2} className="p-2.5 border border-black text-center font-extrabold text-[13px]">
                        စုစုပေါင်း
                      </td>
                      <td className="p-2.5 border border-black text-center font-extrabold text-[13px]">
                        {useBurmeseDigits ? toBurmeseDigits(totalAddressSummary.male) : totalAddressSummary.male}
                      </td>
                      <td className="p-2.5 border border-black text-center font-extrabold text-[13px]">
                        {useBurmeseDigits ? toBurmeseDigits(totalAddressSummary.female) : totalAddressSummary.female}
                      </td>
                      <td className="p-2.5 border border-black text-center font-black text-[13px] bg-slate-200">
                        {useBurmeseDigits ? toBurmeseDigits(totalAddressSummary.total) : totalAddressSummary.total}
                      </td>
                    </tr>
                  </>
                )}
              </tbody>
            </table>

            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px]">
                <div className="h-10" />
                <div className="text-[13px] font-extrabold text-black">{officerTitle}</div>
                <div className="text-[13px] font-bold text-black leading-tight">{officerName}</div>
              </div>
            </div>
          </div>
          </AutoFitSheetWrapper>
        </div>
      </div>
    </div>
  );
};

const TableOutput = ({ 
  records, 
  summaries, 
  setSummaries, 
  masterData,
  showToast,
  onEditRecord
}: { 
  records: ImmRecord[]; 
  summaries: VehicleSummary[]; 
  setSummaries: React.Dispatch<React.SetStateAction<VehicleSummary[]>>; 
  masterData: MasterItem[];
  showToast: (msg: string) => void;
  onEditRecord: (r: ImmRecord) => void;
}) => {
  const today = new Date().toISOString().split('T')[0];
  const [selectedDate, setSelectedDate] = useState(today);
  const [startDate, setStartDate] = useState(today);
  const [endDate, setEndDate] = useState(today);
  const [showRangeExport, setShowRangeExport] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState<Omit<VehicleSummary, 'id' | 'date'>>({
    vehicleNo: '',
    time: '',
    totalInM: 0,
    totalInT: 0,
    totalOutM: 0,
    totalOutT: 0,
    remark: ''
  });

  const filteredSummaries = useMemo(() => 
    summaries.filter(s => s.date === selectedDate).sort((a, b) => a.time.localeCompare(b.time)),
  [summaries, selectedDate]);

  const saveOrUpdateSummary = () => {
    if (!form.vehicleNo) {
      showToast("Vehicle Number required");
      return;
    }
    const nowIso = new Date().toISOString();
    if (editingId !== null) {
      setSummaries(prev => prev.map(s => s.id === editingId ? { ...s, ...form, updatedAt: nowIso, syncStatus: 'pending_sync' } : s));
      setEditingId(null);
      showToast("ENTRY UPDATED SUCCESSFULLY");
    } else {
      const newSummary: VehicleSummary = {
        ...form,
        id: Date.now(),
        date: selectedDate,
        updatedAt: nowIso,
        syncStatus: 'pending_sync'
      };
      setSummaries(prev => [...prev, newSummary]);
      showToast("TABLE ENTRY ADDED");
    }
    setForm({
      vehicleNo: '',
      time: '',
      totalInM: 0,
      totalInT: 0,
      totalOutM: 0,
      totalOutT: 0,
      remark: ''
    });
  };

  const startEditSummary = (s: VehicleSummary) => {
    setEditingId(s.id);
    setForm({
      vehicleNo: s.vehicleNo,
      time: s.time,
      totalInM: s.totalInM || 0,
      totalInT: s.totalInT || 0,
      totalOutM: s.totalOutM || 0,
      totalOutT: s.totalOutT || 0,
      remark: s.remark || ''
    });
    showToast(`EDITING VEHICLE: ${s.vehicleNo}`);
  };

  const cancelEditSummary = () => {
    setEditingId(null);
    setForm({
      vehicleNo: '',
      time: '',
      totalInM: 0,
      totalInT: 0,
      totalOutM: 0,
      totalOutT: 0,
      remark: ''
    });
  };

  const deleteSummary = (id: number) => {
    if (editingId === id) {
      cancelEditSummary();
    }
    setSummaries(prev => prev.filter(s => s.id !== id));
    showToast("ENTRY REMOVED");
  };

  const getForeignerData = (vehicleNo: string, mode: Mode) => {
    const [y, m, d] = selectedDate.split('-');
    const dateStr = `${d}/${m}/${y}`;
    const matches = records.filter(r => 
      r.logType !== 'FCR' &&
      r.vehicleInfo?.toUpperCase().trim() === vehicleNo.toUpperCase().trim() && 
      r.mode === mode && 
      r.timestamp.startsWith(dateStr)
    );
    const M = matches.filter(r => r.gender === 'M').length;
    const F = matches.filter(r => r.gender === 'F').length;
    const T = matches.length;
    return { M, F, T };
  };

  const totals = useMemo(() => {
    const res = {
      cInM: 0, cInF: 0, cInT: 0,
      cOutM: 0, cOutF: 0, cOutT: 0,
      fInM: 0, fInF: 0, fInT: 0,
      fOutM: 0, fOutF: 0, fOutT: 0
    };
    filteredSummaries.forEach(s => {
      const fIn = getForeignerData(s.vehicleNo, 'IN');
      const fOut = getForeignerData(s.vehicleNo, 'OUT');
      
      // Citizen = Total Input - Foreigner Logic
      const cInM = Math.max(0, s.totalInM - fIn.M);
      const cInT = Math.max(0, s.totalInT - fIn.T);
      const cInF = Math.max(0, cInT - cInM);

      const cOutM = Math.max(0, s.totalOutM - fOut.M);
      const cOutT = Math.max(0, s.totalOutT - fOut.T);
      const cOutF = Math.max(0, cOutT - cOutM);

      res.cInM += cInM;
      res.cInF += cInF;
      res.cInT += cInT;
      
      res.cOutM += cOutM;
      res.cOutF += cOutF;
      res.cOutT += cOutT;

      res.fInM += fIn.M;
      res.fInF += fIn.F;
      res.fInT += fIn.T;

      res.fOutM += fOut.M;
      res.fOutF += fOut.F;
      res.fOutT += fOut.T;
    });
    return res;
  }, [filteredSummaries, records, selectedDate]);

  const exportToExcel = () => {
    const data = filteredSummaries.map((s, i) => {
      const fIn = getForeignerData(s.vehicleNo, 'IN');
      const fOut = getForeignerData(s.vehicleNo, 'OUT');

      const cInM = Math.max(0, s.totalInM - fIn.M);
      const cInT = Math.max(0, s.totalInT - fIn.T);
      const cInF = Math.max(0, cInT - cInM);

      const cOutM = Math.max(0, s.totalOutM - fOut.M);
      const cOutT = Math.max(0, s.totalOutT - fOut.T);
      const cOutF = Math.max(0, cOutT - cOutM);

      return {
        "No": i + 1,
        "Vehicle number": s.vehicleNo,
        "Time": s.time,
        "Citizen IN (Male)": cInM,
        "Citizen IN (Female)": cInF,
        "Citizen IN (Total)": cInT,
        "Citizen OUT (Male)": cOutM,
        "Citizen OUT (Female)": cOutF,
        "Citizen OUT (Total)": cOutT,
        "Foreigner IN (Male)": fIn.M,
        "Foreigner IN (Female)": fIn.F,
        "Foreigner IN (Total)": fIn.T,
        "Foreigner OUT (Male)": fOut.M,
        "Foreigner OUT (Female)": fOut.F,
        "Foreigner OUT (Total)": fOut.T,
        "Remark": s.remark
      };
    });

    data.push({
      "No": "Total",
      "Vehicle number": "",
      "Time": "",
      "Citizen IN (Male)": totals.cInM,
      "Citizen IN (Female)": totals.cInF,
      "Citizen IN (Total)": totals.cInT,
      "Citizen OUT (Male)": totals.cOutM,
      "Citizen OUT (Female)": totals.cOutF,
      "Citizen OUT (Total)": totals.cOutT,
      "Foreigner IN (Male)": totals.fInM,
      "Foreigner IN (Female)": totals.fInF,
      "Foreigner IN (Total)": totals.fInT,
      "Foreigner OUT (Male)": totals.fOutM,
      "Foreigner OUT (Female)": totals.fOutF,
      "Foreigner OUT (Total)": totals.fOutT,
      "Remark": ""
    } as any);

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Table Output");
    XLSX.writeFile(wb, `Table_Output_${selectedDate}.xlsx`);
  };

  const exportRangeToExcel = () => {
    const rangeData = summaries
      .filter(s => s.date >= startDate && s.date <= endDate)
      .sort((a, b) => a.date.localeCompare(b.date) || a.time.localeCompare(b.time));

    if (rangeData.length === 0) {
      showToast("No data found for selected range");
      return;
    }

    const data = rangeData.map((s, i) => {
      const fIn = getForeignerDataForDate(s.vehicleNo, 'IN', s.date);
      const fOut = getForeignerDataForDate(s.vehicleNo, 'OUT', s.date);

      const cInM = Math.max(0, s.totalInM - fIn.M);
      const cInT = Math.max(0, s.totalInT - fIn.T);
      const cInF = Math.max(0, cInT - cInM);

      const cOutM = Math.max(0, s.totalOutM - fOut.M);
      const cOutT = Math.max(0, s.totalOutT - fOut.T);
      const cOutF = Math.max(0, cOutT - cOutM);

      return {
        "Date": s.date,
        "No": i + 1,
        "Vehicle number": s.vehicleNo,
        "Time": s.time,
        "Citizen IN (Male)": cInM,
        "Citizen IN (Female)": cInF,
        "Citizen IN (Total)": cInT,
        "Citizen OUT (Male)": cOutM,
        "Citizen OUT (Female)": cOutF,
        "Citizen OUT (Total)": cOutT,
        "Foreigner IN (Male)": fIn.M,
        "Foreigner IN (Female)": fIn.F,
        "Foreigner IN (Total)": fIn.T,
        "Foreigner OUT (Male)": fOut.M,
        "Foreigner OUT (Female)": fOut.F,
        "Foreigner OUT (Total)": fOut.T,
        "Remark": s.remark
      };
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Monthly Summary");
    XLSX.writeFile(wb, `IMM_Summary_Range_${startDate}_to_${endDate}.xlsx`);
    showToast("RANGE EXPORT COMPLETE");
  };

  const getForeignerDataForDate = (vehicleNo: string, mode: Mode, specificDate: string) => {
    const [y, m, d] = specificDate.split('-');
    const dateStr = `${d}/${m}/${y}`;
    const matches = records.filter(r => 
      r.logType !== 'FCR' &&
      r.vehicleInfo?.toUpperCase().trim() === vehicleNo.toUpperCase().trim() && 
      r.mode === mode && 
      r.timestamp.startsWith(dateStr)
    );
    const M = matches.filter(r => r.gender === 'M').length;
    const F = matches.filter(r => r.gender === 'F').length;
    const T = matches.length;
    return { M, F, T };
  };

  const unmatchedRecords = useMemo(() => {
    const [y, m, d] = selectedDate.split('-');
    const dateStr = `${d}/${m}/${y}`;
    const vehicleSet = new Set(filteredSummaries.map(s => s.vehicleNo.toUpperCase().trim()));
    
    return records.filter(r => {
      if (r.logType === 'FCR') return false;
      if (!r.timestamp.startsWith(dateStr)) return false;
      const v = r.vehicleInfo?.toUpperCase().trim() || '';
      return v && !vehicleSet.has(v);
    });
  }, [records, selectedDate, filteredSummaries]);

  return (
    <div className="max-w-[1600px] mx-auto py-8 px-4 space-y-8">
      <div className="card shadow-xl border-t-8 border-indigo-700">
        <datalist id="vList">
          {Array.from(new Set(
            records
              .filter(r => {
                if (r.logType === 'FCR') return false;
                const [y, m, d] = selectedDate.split('-');
                const dStr = `${d}/${m}/${y}`;
                return r.timestamp.startsWith(dStr);
              })
              .map(r => r.vehicleInfo)
              .filter(Boolean)
          )).map(v => <option key={v} value={v} />)}
        </datalist>
        <div className="flex flex-col md:flex-row justify-between items-center gap-6 mb-8 border-b pb-6">
          <div className="flex items-center gap-4">
            <div className="bg-indigo-100 p-3 rounded-full text-indigo-700">
              <Layers size={28} />
            </div>
            <div>
              <h2 className="text-2xl font-black text-gray-800 uppercase tracking-tight">Table Output</h2>
              <p className="text-gray-400 text-xs font-bold font-mono">VEHICLE SUMMARY & CROSS-CHECK</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <input 
              type="date" 
              value={selectedDate} 
              onChange={(e) => setSelectedDate(e.target.value)} 
              className="input-field border-2 border-indigo-100 font-black w-48" 
            />
            <div className="flex gap-2">
              <button 
                onClick={() => setShowRangeExport(!showRangeExport)} 
                className={`btn ${showRangeExport ? 'bg-orange-600' : 'bg-blue-600'} text-white hover:opacity-90`}
              >
                <Calendar size={18} /> {showRangeExport ? 'Hide Range' : 'Range Export'}
              </button>
              <button onClick={exportToExcel} className="btn bg-emerald-600 text-white hover:bg-emerald-700">
                <Download size={18} /> Export Excel
              </button>
            </div>
          </div>
        </div>

        {showRangeExport && (
          <div className="mb-8 p-6 bg-indigo-50 border-2 border-indigo-100 rounded-2xl animate-in slide-in-from-top-4 duration-300">
            <div className="flex flex-col md:flex-row items-end gap-6">
              <div className="flex-1 space-y-2">
                <label className="input-label">Start Date</label>
                <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="input-field" />
              </div>
              <div className="flex-1 space-y-2">
                <label className="input-label">End Date</label>
                <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className="input-field" />
              </div>
              <button onClick={exportRangeToExcel} className="btn bg-indigo-900 text-white hover:bg-black w-full md:w-auto h-[48px] uppercase font-black">
                Process Range Export
              </button>
            </div>
            <p className="text-[10px] text-indigo-400 font-bold mt-4 italic uppercase tracking-wider">
              * Exports all vehicle summaries and calculated foreigner totals for the selected period
            </p>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 p-6 bg-gray-50 rounded-2xl border border-gray-100 mb-8">
           <div className="lg:col-span-2 grid grid-cols-2 gap-4">
              <div>
                <label className="input-label">Vehicle Info</label>
                <input 
                  type="text" 
                  list="vList"
                  value={form.vehicleNo} 
                  onChange={(e) => setForm(prev => ({...prev, vehicleNo: e.target.value}))} 
                  className="input-field" 
                  placeholder="e.g. 1A/1234" 
                />
              </div>
              <div>
                <label className="input-label">Time (hhmm)</label>
                <input 
                  type="text" 
                  value={form.time} 
                  onChange={(e) => setForm(prev => ({...prev, time: e.target.value}))} 
                  className="input-field" 
                  placeholder="24hr format" 
                />
              </div>
           </div>
           <div>
              <label className="input-label">TOTAL IN (Male / Total)</label>
              <div className="flex gap-2">
                <input type="number" value={form.totalInM || ''} onChange={(e) => setForm(prev => ({...prev, totalInM: parseInt(e.target.value) || 0}))} className="input-field" placeholder="M" />
                <input type="number" value={form.totalInT || ''} onChange={(e) => setForm(prev => ({...prev, totalInT: parseInt(e.target.value) || 0}))} className="input-field" placeholder="T" />
              </div>
           </div>
           <div>
              <label className="input-label">TOTAL OUT (Male / Total)</label>
              <div className="flex gap-2">
                <input type="number" value={form.totalOutM || ''} onChange={(e) => setForm(prev => ({...prev, totalOutM: parseInt(e.target.value) || 0}))} className="input-field" placeholder="M" />
                <input type="number" value={form.totalOutT || ''} onChange={(e) => setForm(prev => ({...prev, totalOutT: parseInt(e.target.value) || 0}))} className="input-field" placeholder="T" />
              </div>
           </div>
           <div className="lg:col-span-3">
              <label className="input-label">Remark</label>
              <input type="text" value={form.remark} onChange={(e) => setForm(prev => ({...prev, remark: e.target.value}))} className="input-field" placeholder="Notes..." />
           </div>
           <div className="flex items-end gap-2 w-full">
              <button 
                onClick={saveOrUpdateSummary} 
                className={`btn flex-1 text-white uppercase text-xs font-black h-[48px] transition-all duration-200 ${
                  editingId !== null ? 'bg-amber-600 hover:bg-amber-700 shadow-md ring-2 ring-amber-300' : 'bg-indigo-600 hover:bg-black'
                }`}
              >
                {editingId !== null ? '✓ Update Entry' : 'Add Entry'}
              </button>
              {editingId !== null && (
                <button 
                  onClick={cancelEditSummary} 
                  className="btn bg-gray-200 hover:bg-gray-300 text-gray-700 uppercase text-xs font-black h-[48px] px-4"
                >
                  Cancel
                </button>
              )}
           </div>
        </div>

        <div className="overflow-x-auto border rounded-xl">
          <table className="w-full text-[10px] text-center border-collapse">
            <thead className="bg-gray-900 text-white uppercase font-black tracking-tighter">
              <tr>
                <th rowSpan={2} className="p-2 border-r border-white/10">No</th>
                <th rowSpan={2} className="p-2 border-r border-white/10">Vehicle number</th>
                <th rowSpan={2} className="p-2 border-r border-white/10">Time</th>
                <th colSpan={3} className="p-1 border-b border-white/10 border-r border-white/10 bg-emerald-800/50">Citizen IN</th>
                <th colSpan={3} className="p-1 border-b border-white/10 border-r border-white/10 bg-orange-800/50">Citizen OUT</th>
                <th colSpan={3} className="p-1 border-b border-white/10 border-r border-white/10 bg-indigo-800/50">Foreigner IN</th>
                <th colSpan={3} className="p-1 border-b border-white/10 border-r border-white/10 bg-purple-800/50">Foreigner OUT</th>
                <th rowSpan={2} className="p-2 border-r border-white/10">Remark</th>
                <th rowSpan={2} className="p-2 no-print">Action</th>
              </tr>
              <tr className="bg-gray-800">
                <th className="p-1 border-r border-white/10">M</th>
                <th className="p-1 border-r border-white/10">F</th>
                <th className="p-1 border-r border-white/20 font-black">T</th>
                <th className="p-1 border-r border-white/10">M</th>
                <th className="p-1 border-r border-white/10">F</th>
                <th className="p-1 border-r border-white/20 font-black">T</th>
                <th className="p-1 border-r border-white/10">M</th>
                <th className="p-1 border-r border-white/10">F</th>
                <th className="p-1 border-r border-white/20 font-black">T</th>
                <th className="p-1 border-r border-white/10">M</th>
                <th className="p-1 border-r border-white/10">F</th>
                <th className="p-1 border-r border-white/20 font-black">T</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-200">
              {filteredSummaries.map((s, i) => {
                const fIn = getForeignerData(s.vehicleNo, 'IN');
                const fOut = getForeignerData(s.vehicleNo, 'OUT');

                const cInM = Math.max(0, s.totalInM - fIn.M);
                const cInT = Math.max(0, s.totalInT - fIn.T);
                const cInF = Math.max(0, cInT - cInM);

                const cOutM = Math.max(0, s.totalOutM - fOut.M);
                const cOutT = Math.max(0, s.totalOutT - fOut.T);
                const cOutF = Math.max(0, cOutT - cOutM);

                return (
                  <tr key={s.id} className="hover:bg-gray-50 transition-colors">
                    <td className="p-2 border-r">{i + 1}</td>
                    <td className="p-2 border-r font-bold text-gray-800">{s.vehicleNo}</td>
                    <td className="p-2 border-r font-mono">{s.time}</td>
                    
                    <td className="p-2 border-r bg-emerald-50/30">{cInM}</td>
                    <td className="p-2 border-r bg-emerald-50/30">{cInF}</td>
                    <td className="p-2 border-r bg-emerald-100/50 font-black">{cInT}</td>
                    
                    <td className="p-2 border-r bg-orange-50/30">{cOutM}</td>
                    <td className="p-2 border-r bg-orange-50/30">{cOutF}</td>
                    <td className="p-2 border-r bg-orange-100/50 font-black">{cOutT}</td>

                    <td className="p-2 border-r bg-indigo-50/30">{fIn.M}</td>
                    <td className="p-2 border-r bg-indigo-50/30">{fIn.F}</td>
                    <td className="p-2 border-r bg-indigo-100/50 font-black">{fIn.T}</td>

                    <td className="p-2 border-r bg-purple-50/30">{fOut.M}</td>
                    <td className="p-2 border-r bg-purple-50/30">{fOut.F}</td>
                    <td className="p-2 border-r bg-purple-100/50 font-black">{fOut.T}</td>

                    <td className="p-2 border-r text-left italic text-gray-500">{s.remark}</td>
                    <td className="p-1 no-print">
                      <div className="flex items-center justify-center gap-1.5">
                        <button 
                          onClick={() => startEditSummary(s)} 
                          className="text-indigo-600 hover:text-indigo-900 hover:bg-indigo-50 p-1.5 rounded-lg transition-colors"
                          title="Edit vehicle summary row"
                        >
                          <Edit size={13} strokeWidth={2.5} />
                        </button>
                        <button 
                          onClick={() => deleteSummary(s.id)} 
                          className="text-red-500 hover:text-red-700 hover:bg-red-50 p-1.5 rounded-lg transition-colors"
                          title="Delete vehicle summary row"
                        >
                          <Trash2 size={13} strokeWidth={2.5} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filteredSummaries.length === 0 && (
                <tr>
                  <td colSpan={18} className="p-12 text-gray-400 italic">No vehicle summaries logged for this date.</td>
                </tr>
              )}
            </tbody>
            <tfoot className="bg-gray-900 text-white font-black">
              <tr>
                <td colSpan={3} className="p-2 text-right uppercase tracking-widest text-[11px]">Total Summary</td>
                <td className="p-2 border-r border-white/10">{totals.cInM}</td>
                <td className="p-2 border-r border-white/10">{totals.cInF}</td>
                <td className="p-2 border-r border-white/20 bg-emerald-700">{totals.cInT}</td>
                <td className="p-2 border-r border-white/10">{totals.cOutM}</td>
                <td className="p-2 border-r border-white/10">{totals.cOutF}</td>
                <td className="p-2 border-r border-white/20 bg-orange-700">{totals.cOutT}</td>
                <td className="p-2 border-r border-white/10">{totals.fInM}</td>
                <td className="p-2 border-r border-white/10">{totals.fInF}</td>
                <td className="p-2 border-r border-white/20 bg-indigo-700">{totals.fInT}</td>
                <td className="p-2 border-r border-white/10">{totals.fOutM}</td>
                <td className="p-2 border-r border-white/10">{totals.fOutF}</td>
                <td className="p-2 border-r border-white/20 bg-purple-700">{totals.fOutT}</td>
                <td colSpan={2}></td>
              </tr>
            </tfoot>
          </table>
        </div>

        {unmatchedRecords.length > 0 && (
          <div className="mt-8 p-6 bg-red-50 border-2 border-red-100 rounded-2xl">
            <div className="flex items-center gap-3 mb-4 text-red-700">
              <AlertCircle size={24} className="animate-bounce" />
              <h3 className="font-black uppercase tracking-tight">Warning: Miss-match Vehicle Data Detected</h3>
            </div>
            <p className="text-red-600 text-sm mb-4 font-bold">
              The following foreign movement records found for this date do not have corresponding entries in the summary table above:
            </p>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {unmatchedRecords.map(r => (
                <div key={r.id} className="bg-white p-3 rounded-xl border border-red-200 shadow-sm flex flex-col gap-1 hover:border-red-400 transition-colors group">
                  <div className="flex justify-between items-center text-[10px]">
                    <span className="font-black text-red-600">{r.vehicleInfo}</span>
                    <span className={`px-2 py-0.5 rounded-full font-black text-white ${r.mode === 'IN' ? 'bg-emerald-500' : 'bg-orange-500'}`}>{r.mode}</span>
                  </div>
                  <div className="font-bold text-gray-800 truncate flex justify-between items-center">
                    <span>{r.fullname}</span>
                    <button 
                      onClick={() => onEditRecord(r)}
                      className="text-[10px] bg-indigo-600 text-white px-2 py-1 rounded-lg shadow-sm hover:bg-black transition-colors font-black uppercase"
                    >
                      Quick Edit
                    </button>
                  </div>
                  <div className="text-[9px] text-gray-400 font-mono">{r.timestamp}</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export type CloudRole = 'Superadmin' | 'Admin' | 'Editor' | 'Viewer';

export interface CloudAuthUser {
  username: string;
  role: CloudRole;
  deviceName?: string;
  deviceId?: string;
}

export const AUTH_ACCOUNTS: { username: string; role: CloudRole; description: string; badgeBg: string; icon: string }[] = [
  { username: 'SUPERADMIN', role: 'Superadmin', description: 'Full System Control, Database Management & Device Kick Permissions', badgeBg: 'bg-purple-600 text-purple-100 border-purple-400', icon: '👑' },
  { username: 'EDITOR', role: 'Editor', description: 'Flight Entry Form Logging, History Updates & Data Editing', badgeBg: 'bg-emerald-600 text-emerald-100 border-emerald-400', icon: '✏️' },
  { username: 'VIEWER', role: 'Viewer', description: 'Read-Only Mode (အချက်အလက် ကြည့်ရှုခွင့်သာ ရရှိမည် - ဖြည့်သွင်းခွင့် ပိတ်ထားသည်)', badgeBg: 'bg-slate-600 text-slate-100 border-slate-400', icon: '👁️' },
];

const mergeByUniqueKey = <T extends Record<string, any>>(localArr: T[], remoteArr: T[], key: string = 'id'): T[] => {
  const map = new Map<string, T>();

  const getItemKey = (item: any): string | null => {
    if (!item) return null;
    if (item.date && item.vehicleNo) {
      return `veh_${item.date}_${String(item.vehicleNo).toUpperCase().trim()}`;
    }
    // For masterData items with name and type, unify by type and normalized name so duplicate entries across devices don't split
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

  // 1. Put local items in map first so local data is NEVER lost
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
          map.set(k, item);
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
            map.set(k, { ...localItem, ...item });
          }
        }
      }
    });
  }

  return Array.from(map.values());
};

export default function App() {
    const [isQuotaExhausted, setIsQuotaExhausted] = useState<boolean>(() => getIsQuotaExhausted());
    const [isAutoSyncing, setIsAutoSyncing] = useState<boolean>(false);
    const [isSyncModalOpen, setIsSyncModalOpen] = useState<boolean>(false);

    const [uploadErrorAlarm, setUploadErrorAlarm] = useState<{
      show: boolean;
      reason: string;
      collectionName?: string;
      timestamp: string;
    } | null>(null);

    useEffect(() => {
      const unsub = onQuotaStatusChange((status) => {
        setIsQuotaExhausted(status);
      });
      const unsubWriteError = onWriteError((info) => {
        setUploadErrorAlarm({
          show: true,
          reason: info.reason,
          collectionName: info.collectionName,
          timestamp: new Date().toLocaleTimeString()
        });
      });
      return () => { 
        unsub(); 
        unsubWriteError();
      };
    }, []);

    const [cloudAuthUser, setCloudAuthUser] = useState<CloudAuthUser | null>(() => {
      const saved = localStorage.getItem('imm_pwa_cloud_auth_user');
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (parsed && parsed.username && parsed.username !== 'LOCAL_OFFLINE') {
            return parsed;
          }
        } catch { return null; }
      }
      return null;
    });

    const isViewer = cloudAuthUser?.role === 'Viewer';

    const [showCloudLoginModal, setShowCloudLoginModal] = useState<boolean>(() => {
      const saved = localStorage.getItem('imm_pwa_cloud_auth_user');
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (parsed && parsed.username) {
            return false;
          }
        } catch {}
      }
      return true;
    });
    const [loginUsernameInput, setLoginUsernameInput] = useState('');
    const [loginPasswordInput, setLoginPasswordInput] = useState('');
    const [loginDeviceInput, setLoginDeviceInput] = useState('');
    const [loginAuthError, setLoginAuthError] = useState<string | null>(null);

    // Active Device Sessions state with local storage fallback
    const [deviceSessions, setDeviceSessions] = useState<any[]>(() => {
      const saved = localStorage.getItem('imm_pwa_device_sessions');
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed) && parsed.length > 0) return parsed;
        } catch {}
      }
      return [
        {
          deviceId: 'dev_myeik_pad_01',
          deviceName: 'iPad Safari PWA (Myeik Terminal)',
          accountRole: 'Editor',
          username: 'EDITOR',
          lastActive: new Date().toLocaleString(),
          lastPingTimestamp: Date.now(),
          kicked: false
        }
      ];
    });

    // Initialize Firebase Anonymous Authentication on App startup
    useEffect(() => {
      initAuth();
    }, []);

    // Subscribe to Active Sessions for SuperAdmin or targeted device kick check & real-time role changes
    useEffect(() => {
      if (!cloudAuthUser || cloudAuthUser.username === 'LOCAL_OFFLINE') return;
      const currentDevId = cloudAuthUser.deviceId || localStorage.getItem('imm_pwa_device_id');
      if (!currentDevId) return;

      // Single-document listener for current device kick status and role sync
      const unsub = subscribeToMyDeviceSession(currentDevId, (data) => {
        if (!data) return;
        if (data.kicked) {
          setCloudAuthUser(null);
          localStorage.removeItem('imm_pwa_cloud_auth_user');
          setShowCloudLoginModal(true);
          setLoginAuthError("⚠️ သင်၏ စက်အသုံးပြုခွင့် (Device Session) အား SuperAdmin မှ ပိတ်သိမ်း (Revoke/Kick) လိုက်ပါပြီ။ အကောင့်အသစ်ဖြင့် ပြန်လည်ဝင်ရောက်ပါ");
          showToast("⛔ သင်၏ စက်အသုံးပြုခွင့်အား SuperAdmin မှ ပိတ်သိမ်း (Kick) လိုက်ပါပြီ");
          return;
        }

        // Live Role Sync: If Superadmin changed this device's role in Cloud
        if (data.accountRole && cloudAuthUser && cloudAuthUser.role !== data.accountRole) {
          const updatedUser: CloudAuthUser = {
            ...cloudAuthUser,
            role: data.accountRole
          };
          setCloudAuthUser(updatedUser);
          localStorage.setItem('imm_pwa_cloud_auth_user', JSON.stringify(updatedUser));
          if (data.accountRole === 'Viewer') {
            showToast("🔒 Superadmin မှ ဤစက်အား Viewer (View-Only Mode) သို့ ပြောင်းလဲသတ်မှတ်လိုက်ပါပြီ");
          } else if (data.accountRole === 'Editor') {
            showToast("✏️ Superadmin မှ ဤစက်အား Editor အဆင့်သို့ ပြောင်းလဲလိုက်ပါပြီ");
          } else if (data.accountRole === 'Superadmin') {
            showToast("👑 Superadmin မှ ဤစက်အား Superadmin အဆင့်သို့ တိုးမြှင့်လိုက်ပါပြီ");
          }
        }
      });

      return () => {
        if (unsub) unsub();
      };
    }, [cloudAuthUser?.deviceId, cloudAuthUser?.username, cloudAuthUser?.role]);

    // Subscribe to all connected device sessions for SuperAdmin live monitoring
    useEffect(() => {
      if (cloudAuthUser?.role !== 'Superadmin') return;

      const unsub = subscribeToDeviceSessions((sessions) => {
        if (Array.isArray(sessions) && sessions.length > 0) {
          setDeviceSessions(sessions);
          localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(sessions));
        }
      });

      return () => {
        if (unsub) unsub();
      };
    }, [cloudAuthUser?.role]);

    // Heartbeat Ping helper function (Sends active device status to Cloud)
    const sendHeartbeatPing = async (userObj?: CloudAuthUser) => {
      const activeUser = userObj || cloudAuthUser;
      if (!activeUser || !activeUser.username || activeUser.username === 'LOCAL_OFFLINE') return;
      const devId = activeUser.deviceId || localStorage.getItem('imm_pwa_device_id') || `dev_${Date.now()}`;
      const devName = activeUser.deviceName || (navigator.userAgent?.includes('iPad') ? 'iPad Safari PWA' : 'PWA Client Device');
      const now = new Date();

      const newSession = {
        deviceId: devId,
        deviceName: devName,
        accountRole: activeUser.role,
        username: activeUser.username,
        lastActive: now.toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' }),
        lastPingTimestamp: now.getTime(),
        kicked: false
      };

      await saveDeviceSession(newSession);
    };

    // Hourly Heartbeat Timer Effect (1 Hour Interval = 3,600,000ms to minimize Cloud Read/Write Quota)
    useEffect(() => {
      if (!cloudAuthUser || cloudAuthUser.username === 'LOCAL_OFFLINE') return;

      sendHeartbeatPing(cloudAuthUser);

      // Heartbeat ping every 1 Hour (3,600,000 ms)
      const ONE_HOUR_MS = 60 * 60 * 1000;
      const heartbeatInterval = setInterval(() => {
        sendHeartbeatPing();
      }, ONE_HOUR_MS);

      return () => clearInterval(heartbeatInterval);
    }, [cloudAuthUser?.deviceId, cloudAuthUser?.username]);

    // Font Scaling & UI Zoom state (11 granular scaling levels)
    const [fontScale, setFontScale] = useState<string>(() => {
      return localStorage.getItem('imm_pwa_font_scale') || '100';
    });

    // Online & Activity tracking
    const [isOnline, setIsOnline] = useState<boolean>(() => typeof navigator !== 'undefined' ? navigator.onLine : true);
    const lastActivityRef = useRef<number>(Date.now());
    const autoSyncTimerRef = useRef<any>(null);

    useEffect(() => {
      const handleOnline = () => setIsOnline(true);
      const handleOffline = () => setIsOnline(false);
      window.addEventListener('online', handleOnline);
      window.addEventListener('offline', handleOffline);
      return () => {
        window.removeEventListener('online', handleOnline);
        window.removeEventListener('offline', handleOffline);
      };
    }, []);

    // Navigation Dropdowns & Popovers state
    const [isMainMenuOpen, setIsMainMenuOpen] = useState(false);
    const [isStatusPopoverOpen, setIsStatusPopoverOpen] = useState(false);
    const [isUserDropdownOpen, setIsUserDropdownOpen] = useState(false);
    const [isSettingsOpen, setIsSettingsOpen] = useState(false);

    useEffect(() => {
      localStorage.setItem('imm_pwa_font_scale', fontScale);
      const root = document.documentElement;
      const numScale = parseInt(fontScale, 10) || 100;
      root.style.fontSize = (16 * (numScale / 100)) + 'px';
    }, [fontScale]);

    // Persistent Storage initialization
    useEffect(() => {
      const pwaVer = localStorage.getItem('imm_pwa_version');
      if (pwaVer !== 'v3_cloud_first') {
        localStorage.setItem('imm_pwa_version', 'v3_cloud_first');
      }

      if (navigator.storage && navigator.storage.persist) {
        navigator.storage.persist();
      }
    }, []);

    const showToast = (msg: string) => {
      setToast({ message: msg, visible: true });
      setTimeout(() => setToast(prev => ({ ...prev, visible: false })), 3500);
    };

    const markSynced = (col: string) => {
      if (col === 'records') {
        setRecords(prev => prev.map(r => (r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed') ? { ...r, syncStatus: 'synced' } : r));
      } else if (col === 'tempRecords') {
        setTempRecords(prev => prev.map(r => (r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed') ? { ...r, syncStatus: 'synced' } : r));
      } else if (col === 'vehicleSummaries') {
        setVehicleSummaries(prev => prev.map(s => (s.syncStatus === 'pending_sync' || s.syncStatus === 'upload_failed') ? { ...s, syncStatus: 'synced' } : s));
      } else if (col === 'checkingHistory') {
        setCheckingHistory(prev => prev.map(r => (r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed') ? { ...r, syncStatus: 'synced' } : r));
      }
    };

    const markUploadFailed = (col: string) => {
      if (col === 'records') {
        setRecords(prev => prev.map(r => (r.syncStatus === 'pending_sync' || !r.syncStatus) ? { ...r, syncStatus: 'upload_failed' } : r));
      } else if (col === 'tempRecords') {
        setTempRecords(prev => prev.map(r => (r.syncStatus === 'pending_sync' || !r.syncStatus) ? { ...r, syncStatus: 'upload_failed' } : r));
      } else if (col === 'vehicleSummaries') {
        setVehicleSummaries(prev => prev.map(s => (s.syncStatus === 'pending_sync' || !s.syncStatus) ? { ...s, syncStatus: 'upload_failed' } : s));
      } else if (col === 'checkingHistory') {
        setCheckingHistory(prev => prev.map(r => (r.syncStatus === 'pending_sync' || !r.syncStatus) ? { ...r, syncStatus: 'upload_failed' } : r));
      }
    };

    const handleManualSync = async () => {
      if (!isOnline) {
        showToast("⚠️ အင်တာနက် လိုင်းမရှိသေးပါ။ အချက်အလက်များ Local DB တွင် လုံခြုံစွာ ရှိနေပါသည်");
        return;
      }
      try {
        resetQuotaState();
        showToast(" Cloud သို့ ဒေတာများ တိုက်ရိုက် Upload Sync ပြုလုပ်နေပါသည်...");
        const res1 = await saveCollectionToFirestore('records', records, true);
        const res2 = await saveCollectionToFirestore('tempRecords', tempRecords, true);
        const res3 = await saveCollectionToFirestore('masterData', masterData, true);
        const res4 = await saveCollectionToFirestore('vehicleSummaries', vehicleSummaries, true);
        const res5 = await saveCollectionToFirestore('checkingHistory', checkingHistory, true);
        const res6 = await saveCollectionToFirestore('dossierHistory', dossierHistory, true);
        const res7 = await saveCollectionToFirestore('watchList', watchList, true);

        if (res1 && res2 && res3 && res4 && res5 && res6 && res7) {
          setIsCloudSynced(true);
          updateLastSyncTimestamp();
          setUploadErrorAlarm(null);
          markSynced('records');
          markSynced('tempRecords');
          markSynced('masterData');
          markSynced('vehicleSummaries');
          markSynced('checkingHistory');
          markSynced('dossierHistory');
          markSynced('watchList');
          showToast(` Cloud သို့ အချက်အလက်များ အားလုံး တိုက်ရိုက် Upload Sync ပြီးပါပြီ (Records ${records.length} ခု၊ WatchList ${watchList.length} ခု)`);
        } else {
          setIsCloudSynced(false);
          if (res1) markSynced('records'); else markUploadFailed('records');
          if (res2) markSynced('tempRecords'); else markUploadFailed('tempRecords');
          if (res3) markSynced('masterData'); else markUploadFailed('masterData');
          if (res4) markSynced('vehicleSummaries'); else markUploadFailed('vehicleSummaries');
          if (res5) markSynced('checkingHistory'); else markUploadFailed('checkingHistory');
          if (res6) markSynced('dossierHistory'); else markUploadFailed('dossierHistory');
          if (res7) markSynced('watchList'); else markUploadFailed('watchList');

          const reason = getIsQuotaExhausted() ? "Write Quota Limit Reached" : "Write Error / Server Rejection";
          setUploadErrorAlarm({
            show: true,
            reason,
            timestamp: new Date().toLocaleTimeString()
          });
          showToast("⚠️ Upload Failure Alert: Local DB တွင် ဒေတာလုံခြုံစွာ သိမ်းဆည်းထားပါသည်");
        }
      } catch (err: any) {
        setIsCloudSynced(false);
        markUploadFailed('records');
        markUploadFailed('tempRecords');
        setUploadErrorAlarm({
          show: true,
          reason: err?.message || "Upload Failure / Server Error",
          timestamp: new Date().toLocaleTimeString()
        });
        showToast("⚠️ Sync လုပ်ဆောင်ရာတွင် ချို့ယွင်းချက်ရှိပါသည် (Local Data လုံခြုံစွာ ရှိပါသည်)");
      }
    };

    const handleFetchDataFromCloud = async (force: boolean = false, silent: boolean = false) => {
      if (!isOnline) {
        if (!silent) showToast("⚠️ အင်တာနက် လိုင်းမရှိသေးပါ။ လတ်တလော Local DB မှ အချက်အလက်များကို အသုံးပြုနေပါသည်");
        return;
      }
      try {
        isRemoteSyncingRef.current = true;
        if (!silent) {
          setDownloadProgress({ isDownloading: true, percent: 5, currentCol: 'Metadata Check' });
        }
        
        // Smart fetch: check Firestore collections for updates
        const updatedCollections = await checkAndFetchUpdatedCollections([
          'records', 'tempRecords', 'masterData', 'vehicleSummaries', 'checkingHistory', 'dossierHistory', 'watchList'
        ], force, (pct, col) => {
          if (!silent) {
            setDownloadProgress({ isDownloading: true, percent: pct, currentCol: col });
          }
        });

        let updatedAny = false;
        let fetchedAny = false;

        if (Array.isArray(updatedCollections.records) && updatedCollections.records.length > 0) {
          fetchedAny = true;
          setRecords(currentRecords => {
            const merged = mergeByUniqueKey(currentRecords, updatedCollections.records, 'id');
            if (merged.length !== currentRecords.length || JSON.stringify(merged) !== JSON.stringify(currentRecords)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_records_react', JSON.stringify(merged));
            miniDB.set('records', merged).catch(() => {});
            setSyncedHash('records', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.tempRecords) && updatedCollections.tempRecords.length > 0) {
          fetchedAny = true;
          setTempRecords(currentTemp => {
            const merged = mergeByUniqueKey(currentTemp, updatedCollections.tempRecords, 'id');
            if (merged.length !== currentTemp.length || JSON.stringify(merged) !== JSON.stringify(currentTemp)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_temp_records_react', JSON.stringify(merged));
            miniDB.set('tempRecords', merged).catch(() => {});
            setSyncedHash('tempRecords', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.masterData) && updatedCollections.masterData.length > 0) {
          fetchedAny = true;
          setMasterData(currentMaster => {
            const merged = mergeByUniqueKey(currentMaster, updatedCollections.masterData, 'id');
            if (merged.length !== currentMaster.length || JSON.stringify(merged) !== JSON.stringify(currentMaster)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_master_react', JSON.stringify(merged));
            miniDB.set('masterData', merged).catch(() => {});
            setSyncedHash('masterData', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.vehicleSummaries) && updatedCollections.vehicleSummaries.length > 0) {
          fetchedAny = true;
          setVehicleSummaries(currentVehicle => {
            const merged = mergeByUniqueKey(currentVehicle, updatedCollections.vehicleSummaries, 'id');
            if (merged.length !== currentVehicle.length || JSON.stringify(merged) !== JSON.stringify(currentVehicle)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_vehicle_summaries', JSON.stringify(merged));
            miniDB.set('vehicleSummaries', merged).catch(() => {});
            setSyncedHash('vehicleSummaries', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.checkingHistory) && updatedCollections.checkingHistory.length > 0) {
          fetchedAny = true;
          setCheckingHistory(currentHistory => {
            const merged = mergeByUniqueKey(currentHistory, updatedCollections.checkingHistory, 'id');
            if (merged.length !== currentHistory.length || JSON.stringify(merged) !== JSON.stringify(currentHistory)) {
              updatedAny = true;
            }
            localStorage.setItem('checking_verification_history_logs_v1', JSON.stringify(merged));
            miniDB.set('checkingHistory', merged).catch(() => {});
            setSyncedHash('checkingHistory', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.dossierHistory) && updatedCollections.dossierHistory.length > 0) {
          fetchedAny = true;
          setDossierHistory(currentDossier => {
            const merged = mergeByUniqueKey(currentDossier, updatedCollections.dossierHistory, 'id');
            if (merged.length !== currentDossier.length || JSON.stringify(merged) !== JSON.stringify(currentDossier)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_dossier_history_react', JSON.stringify(merged));
            miniDB.set('dossierHistory', merged).catch(() => {});
            setSyncedHash('dossierHistory', merged);
            return merged;
          });
        }

        if (Array.isArray(updatedCollections.watchList) && updatedCollections.watchList.length > 0) {
          fetchedAny = true;
          setWatchList(currentWatch => {
            const merged = mergeByUniqueKey(currentWatch, updatedCollections.watchList, 'id');
            if (merged.length !== currentWatch.length || JSON.stringify(merged) !== JSON.stringify(currentWatch)) {
              updatedAny = true;
            }
            localStorage.setItem('imm_watchlist_records_v1', JSON.stringify(merged));
            miniDB.set('watchList', merged).catch(() => {});
            setSyncedHash('watchList', merged);
            return merged;
          });
        }

        setIsCloudSynced(true);
        updateLastSyncTimestamp();
        if (silent) {
          if (updatedAny) {
            showToast("📱 အခြားစက်မှ အချက်အလက်သစ်များ ရောက်ရှိပါပြီ (Auto-Synced from Server)");
          }
        } else {
          if (updatedAny || fetchedAny) {
            showToast(" Cloud မှ အချက်အလက်သစ်များ အောင်မြင်စွာ ရယူပြီးပါပြီ");
          } else {
            showToast(" Cloud အချက်အလက်နှင့် လက်ရှိ Local အချက်အလက်များ တူညီနေပါသည်");
          }
        }
      } catch (err) {
        console.error("Fetch cloud data error:", err);
        if (!silent) showToast("⚠️ Local Data ကို ဆက်လက်အသုံးပြုနေပါသည်");
      } finally {
        setIsInitialSyncing(false);
        if (!silent) {
          setDownloadProgress({ isDownloading: false, percent: 100, currentCol: 'Complete' });
          setTimeout(() => {
            setDownloadProgress(null);
            isRemoteSyncingRef.current = false;
          }, 1200);
        } else {
          isRemoteSyncingRef.current = false;
        }
      }
    };

    // 25-Second Periodic Background Smart Poll & Window Focus Sync (Auto-receives updates from other devices/phones)
    useEffect(() => {
      if (!cloudAuthUser || cloudAuthUser.username === 'LOCAL_OFFLINE') return;

      const backgroundSyncInterval = setInterval(() => {
        if (isOnline && document.visibilityState === 'visible' && !isRemoteSyncingRef.current && !isAutoSyncing) {
          handleFetchDataFromCloud(false, true).catch(() => {});
        }
      }, 25000);

      const handleFocus = () => {
        if (isOnline && !isRemoteSyncingRef.current && !isAutoSyncing) {
          handleFetchDataFromCloud(false, true).catch(() => {});
        }
      };

      window.addEventListener('focus', handleFocus);

      return () => {
        clearInterval(backgroundSyncInterval);
        window.removeEventListener('focus', handleFocus);
      };
    }, [cloudAuthUser, isOnline, isAutoSyncing]);

    // Startup Launch Sequence: Check if Cloud Auth exists
    useEffect(() => {
      const savedCloud = localStorage.getItem('imm_pwa_cloud_auth_user');
      if (savedCloud) {
        try {
          const parsed = JSON.parse(savedCloud);
          if (parsed && parsed.username) {
            setCloudAuthUser(parsed);
            setShowCloudLoginModal(false);
          } else {
            localStorage.removeItem('imm_pwa_cloud_auth_user');
            setCloudAuthUser(null);
            setShowCloudLoginModal(true);
          }
        } catch {
          localStorage.removeItem('imm_pwa_cloud_auth_user');
          setCloudAuthUser(null);
          setShowCloudLoginModal(true);
        }
      } else {
        setCloudAuthUser(null);
        setShowCloudLoginModal(true);
      }
    }, []);

    // PWA Install Prompt State
    const [pwaInstallPrompt, setPwaInstallPrompt] = useState<any>(null);
    const [isPwaInstalled, setIsPwaInstalled] = useState(false);

    useEffect(() => {
      const handleBeforeInstall = (e: any) => {
        e.preventDefault();
        setPwaInstallPrompt(e);
      };
      const handleAppInstalled = () => {
        setIsPwaInstalled(true);
        setPwaInstallPrompt(null);
      };
      window.addEventListener('beforeinstallprompt', handleBeforeInstall);
      window.addEventListener('appinstalled', handleAppInstalled);
      return () => {
        window.removeEventListener('beforeinstallprompt', handleBeforeInstall);
        window.removeEventListener('appinstalled', handleAppInstalled);
      };
    }, []);

    // Cloud Login with SHA-256 Hashing & Zero Plaintext Evaluation
    const handleCloudLogin = async (customUsername?: string, customPassword?: string) => {
      const u = (customUsername !== undefined ? customUsername : loginUsernameInput).trim().toUpperCase();
      const rawPass = (customPassword !== undefined ? customPassword : loginPasswordInput).trim();
      const devName = loginDeviceInput.trim() || (navigator.userAgent.includes('iPad') ? 'iPad Safari PWA' : 'PWA Client Device');

      if (!u || !rawPass) {
        setLoginAuthError("ကျေးဇူးပြု၍ Username နှင့် Password ရိုက်ထည့်ပါ");
        return;
      }

      const inputHash = await hashSHA256(rawPass);
      const superadminHash = await hashSHA256('HeinHtet@3612');
      const editorHash = await hashSHA256('09799395503');
      const viewerHash = await hashSHA256('viewer');
      const viewerHash2 = await hashSHA256('123456');

      let matchedRole: CloudRole | null = null;
      if (u === 'SUPERADMIN' && inputHash === superadminHash) {
        matchedRole = 'Superadmin';
      } else if (u === 'EDITOR' && inputHash === editorHash) {
        matchedRole = 'Editor';
      } else if (u === 'VIEWER' && (inputHash === viewerHash || inputHash === viewerHash2 || rawPass === 'viewer' || rawPass === '123456')) {
        matchedRole = 'Viewer';
      }

      if (matchedRole) {
        const deviceId = localStorage.getItem('imm_pwa_device_id') || `dev_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
        localStorage.setItem('imm_pwa_device_id', deviceId);

        const authObj: CloudAuthUser = {
          username: u,
          role: matchedRole,
          deviceName: devName,
          deviceId
        };

        setCloudAuthUser(authObj);
        localStorage.setItem('imm_pwa_cloud_auth_user', JSON.stringify(authObj));

        // Add to active device sessions and sync to Firestore
        const now = new Date();
        const freshSession = {
          deviceId,
          deviceName: devName,
          accountRole: matchedRole,
          username: u,
          lastActive: now.toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' }),
          lastPingTimestamp: now.getTime(),
          kicked: false
        };

        saveDeviceSession(freshSession).catch(() => {});

        setLoginAuthError(null);
        setLoginUsernameInput('');
        setLoginPasswordInput('');
        setShowCloudLoginModal(false);
        setIsInitialSyncing(true);
        showToast(`မင်္ဂလာပါ ${u} (${matchedRole}) အနေဖြင့် အသုံးပြုနိုင်ပါပြီ`);

        // Perform Initial Cloud Sync for the new device
        handleFetchDataFromCloud(true)
          .then(() => {
            showToast("Cloud မှ ဒေတာများ အောင်မြင်စွာ Sync လုပ်ပြီးပါပြီ");
          })
          .catch(err => console.warn("Background cloud check on login:", err))
          .finally(() => {
            setIsInitialSyncing(false);
          });
      } else {
        setLoginAuthError("Username သို့မဟုတ် Password မှားယွင်းနေပါသည်။ (လုံခြုံရေးအရ ငြင်းပယ်သည်)");
      }
    };

    const dismissCloudLogin = () => {
      if (cloudAuthUser) {
        setShowCloudLoginModal(false);
      } else {
        showToast("⚠️ အက်ပလီကေးရှင်းကို သုံးနိုင်ရန် Cloud Role Account မဖြစ်မနေ ဝင်ရောက်ရန် လိုအပ်ပါသည်");
      }
    };

    const handleCloudLogout = () => {
      setCloudAuthUser(null);
      localStorage.removeItem('imm_pwa_cloud_auth_user');
      setShowCloudLoginModal(true);
      showToast("Cloud Sync အကောင့်မှ ထွက်လိုက်ပါပြီ။ ဆက်လက်အသုံးပြုရန် Cloud Login မဖြစ်မနေ ဝင်ပါ။");
    };

    const handleKickDevice = async (targetDeviceId: string) => {
      const ok = await setDeviceKickedStatus(targetDeviceId, true);
      if (ok) {
        setDeviceSessions(prev => {
          const updated = prev.map(s => s.deviceId === targetDeviceId ? { ...s, kicked: true } : s);
          localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(updated));
          return updated;
        });
        logActivity({
          action: 'DELETE',
          module: 'SYSTEM',
          officerName: cloudAuthUser?.username || 'SUPERADMIN',
          officerRole: 'Superadmin',
          deviceId: targetDeviceId,
          targetId: targetDeviceId,
          details: `Device (${targetDeviceId}) အား Superadmin မှ Remote Kick (ပိတ်ဆို့) ခဲ့ပါသည်`
        });
        showToast("🚫 အဆိုပါ စက်အသုံးပြုမှုအား Remote Kick (ပိတ်ဆို့) လိုက်ပါပြီ");
      } else {
        showToast("⚠️ Remote Kick ပြုလုပ်ရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
      }
    };

    const handleReauthorizeDevice = async (targetDeviceId: string) => {
      const ok = await setDeviceKickedStatus(targetDeviceId, false);
      if (ok) {
        setDeviceSessions(prev => {
          const updated = prev.map(s => s.deviceId === targetDeviceId ? { ...s, kicked: false } : s);
          localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(updated));
          return updated;
        });
        logActivity({
          action: 'UPDATE',
          module: 'SYSTEM',
          officerName: cloudAuthUser?.username || 'SUPERADMIN',
          officerRole: 'Superadmin',
          deviceId: targetDeviceId,
          targetId: targetDeviceId,
          details: `Device (${targetDeviceId}) အား Superadmin မှ Re-Authorize (ပြန်လည်ခွင့်ပြု) ခဲ့ပါသည်`
        });
        showToast("✅ စက်အား အသုံးပြုခွင့် ပြန်လည် ခွင့်ပြုလိုက်ပါပြီ (Re-Authorized)");
      } else {
        showToast("⚠️ Re-Authorize ပြုလုပ်ရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
      }
    };

    const handleUpdateDeviceRole = async (targetDeviceId: string, newRole: CloudRole) => {
      const ok = await updateDeviceRole(targetDeviceId, newRole);
      if (ok) {
        setDeviceSessions(prev => {
          const updated = prev.map(s => s.deviceId === targetDeviceId ? { ...s, accountRole: newRole } : s);
          localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(updated));
          return updated;
        });
        const currentDevId = localStorage.getItem('imm_pwa_device_id');
        if (targetDeviceId === currentDevId && cloudAuthUser) {
          const updatedUser = { ...cloudAuthUser, role: newRole };
          setCloudAuthUser(updatedUser);
          localStorage.setItem('imm_pwa_cloud_auth_user', JSON.stringify(updatedUser));
        }
        logActivity({
          action: 'UPDATE',
          module: 'SYSTEM',
          officerName: cloudAuthUser?.username || 'SUPERADMIN',
          officerRole: 'Superadmin',
          deviceId: targetDeviceId,
          targetId: targetDeviceId,
          details: `Device (${targetDeviceId}) ၏ Role အား ${newRole} သို့ Superadmin မှ ပြောင်းလဲသတ်မှတ်ခဲ့ပါသည်`
        });
        showToast(`✅ Device Role အား ${newRole} သို့ ပြောင်းလဲလိုက်ပါပြီ (သက်ဆိုင်ရာစက်တွင် ချက်ချင်း အသက်ဝင်ပါမည်)`);
      } else {
        showToast("⚠️ Role ပြောင်းလဲရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
      }
    };

    const handleDeleteDeviceSession = async (targetDeviceId: string) => {
      const ok = await deleteDeviceSession(targetDeviceId);
      if (ok) {
        setDeviceSessions(prev => {
          const updated = prev.filter(s => s.deviceId !== targetDeviceId);
          localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(updated));
          return updated;
        });
        showToast("🗑️ Kicked device session အား စနစ်မှ ဖျက်ထုတ်လိုက်ပါပြီ");
      } else {
        showToast("⚠️ Device session ဖျက်ရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
      }
    };

    const [activeTab, setActiveTab] = useState<Tab>('entry');
    const [language, setLanguage] = useState<'ENG' | 'BUR'>('ENG');
    const [records, setRecords] = useState<ImmRecord[]>([]);
    const [tempRecords, setTempRecords] = useState<ImmRecord[]>([]);
    const [preFilledTempId, setPreFilledTempId] = useState<number | null>(null);
    const [masterData, setMasterData] = useState<MasterItem[]>([]);
    const [currentMode, setCurrentMode] = useState<Mode>('IN');
    const [toast, setToast] = useState({ message: '', visible: false });
    const [editTarget, setEditTarget] = useState<ImmRecord | null>(null);
    const [mobileDetailRecord, setMobileDetailRecord] = useState<ImmRecord | null>(null);
    const [globalDateRange, setGlobalDateRange] = useState({ from: '', to: '' });
    const [searchQuery, setSearchQuery] = useState('');
    const [analyzerTarget, setAnalyzerTarget] = useState<string | null>(null);
    const [tehAnalysisTarget, setTehAnalysisTarget] = useState<ImmRecord | null>(null);
    const [movementFilter, setMovementFilter] = useState<'ALL' | 'IN' | 'OUT'>('ALL');
    const [movementSearchQuery, setMovementSearchQuery] = useState('');
    const [selectedIds, setSelectedIds] = useState<number[]>([]);
    const [selectedTempIds, setSelectedTempIds] = useState<number[]>([]);
    const [confirmDeleteId, setConfirmDeleteId] = useState<number | null>(null);
    const [confirmTempDeleteId, setConfirmTempDeleteId] = useState<number | null>(null);
    const [vehicleSummaries, setVehicleSummaries] = useState<VehicleSummary[]>([]);
    const [isCounterCheckOpen, setIsCounterCheckOpen] = useState<boolean>(false);
    const counterCheckSuspectCount = useMemo(() => {
      return analyzeRecordsForErrors(records).length;
    }, [records]);
    const [tempAlert, setTempAlert] = useState<string | null>(null);
    const [passportSearchResults, setPassportSearchResults] = useState<ImmRecord[]>([]);
    const [nameSearchResults, setNameSearchResults] = useState<ImmRecord[]>([]);
    const [ffeConfirmation, setFfeConfirmation] = useState<{ show: boolean, message: string } | null>(null);
    const [ffeSuccessMsg, setFfeSuccessMsg] = useState<string | null>(null);
    const [currentUser, setCurrentUser] = useState<{ name: string, title: string } | null>(null);
    const [loginName, setLoginName] = useState('');
    const [showOfficerLoginModal, setShowOfficerLoginModal] = useState<boolean>(false);

    // 5-Minute Inactivity Timeout (Officer Session Reset)
    useEffect(() => {
      if (!currentUser) return;

      const resetActivity = () => {
        lastActivityRef.current = Date.now();
      };

      const activityEvents = ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => window.addEventListener(evt, resetActivity, { passive: true }));

      const handleVis = () => {
        if (!document.hidden) resetActivity();
      };
      document.addEventListener('visibilitychange', handleVis);

      const checkInterval = setInterval(() => {
        const elapsed = Date.now() - lastActivityRef.current;
        if (elapsed >= 5 * 60 * 1000) { // 5 minutes = 300,000ms
          setCurrentUser(null);
          localStorage.removeItem('lastCheckedOfficerName');
          localStorage.removeItem('lastCheckedOfficerTitle');
          showToast("5 မိနစ်အတွင်း မလှုပ်ရှားမှုကြောင့် စာရင်းသွင်းအရာရှိ အမည် ရွေးချယ်မှု ပိတ်သွားပါသည်");
        }
      }, 10000);

      return () => {
        activityEvents.forEach(evt => window.removeEventListener(evt, resetActivity));
        document.removeEventListener('visibilitychange', handleVis);
        clearInterval(checkInterval);
      };
    }, [currentUser]);
    const [officerNameInput, setOfficerNameInput] = useState('');
    const [officerTitleInput, setOfficerTitleInput] = useState('');
    const [showBackupReminderModal, setShowBackupReminderModal] = useState(false);
    const [tehQuery, setTehQuery] = useState('');
    const [isDBCardLoaded, setIsDBCardLoaded] = useState(false);
    const [isInitialSyncing, setIsInitialSyncing] = useState<boolean>(false);
    const [waitingElapsedSeconds, setWaitingElapsedSeconds] = useState(0);

    useEffect(() => {
      if (isDBCardLoaded && !isInitialSyncing) return;
      const timer = setInterval(() => {
        setWaitingElapsedSeconds(prev => prev + 1);
      }, 1000);
      return () => clearInterval(timer);
    }, [isDBCardLoaded, isInitialSyncing]);
    const [isCloudSynced, setIsCloudSynced] = useState(false);
    const [lastSyncTime, setLastSyncTime] = useState<string>(() => {
      return localStorage.getItem('imm_last_sync_time') || '';
    });

    const updateLastSyncTimestamp = async (overrideIso?: string) => {
      let targetIso = overrideIso;

      if (!targetIso) {
        try {
          const meta = await getCloudMetadata();
          if (meta) {
            const times = Object.values(meta).filter((v): v is string => typeof v === 'string' && !isNaN(new Date(v).getTime()));
            if (times.length > 0) {
              times.sort((a, b) => new Date(b).getTime() - new Date(a).getTime());
              targetIso = times[0];
            }
          }
        } catch (e) {
          console.warn("Could not fetch cloud metadata timestamp:", e);
        }
      }

      if (!targetIso) {
        const localTs = getLocalTimestamps();
        const times = Object.values(localTs).filter((v): v is string => typeof v === 'string' && !isNaN(new Date(v).getTime()));
        if (times.length > 0) {
          times.sort((a, b) => new Date(b).getTime() - new Date(a).getTime());
          targetIso = times[0];
        }
      }

      if (targetIso) {
        const dateObj = new Date(targetIso);
        const formatted = `${dateObj.toLocaleDateString('en-GB')} ${dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`;
        setLastSyncTime(formatted);
        localStorage.setItem('imm_last_sync_time', formatted);
      }
    };
    const [downloadProgress, setDownloadProgress] = useState<{ isDownloading: boolean; percent: number; currentCol: string } | null>(null);
    const isRemoteSyncingRef = useRef(false);
    const prevLocalHashesRef = useRef<Record<string, string>>({});

    const [dailyPdfs, setDailyPdfs] = useState<Record<string, { base64: string; name: string }>>({});
    const [checkingHistory, setCheckingHistory] = useState<CheckingHistoryEntry[]>([]);
    const [dossierHistory, setDossierHistory] = useState<DossierRecord[]>([]);
    const [watchList, setWatchList] = useState<WatchListRecord[]>([]);

    const [checkingPrintData, setCheckingPrintData] = useState<{
      title: string;
      type: 'CO_LTD_PENDING' | 'OTHERS_PENDING' | 'CHECKED_HISTORY';
      data: any[];
    } | null>(null);

    const [activePrintPreview, setActivePrintPreview] = useState<{
      title: string;
      type: 'CO_LTD_PENDING' | 'OTHERS_PENDING' | 'CHECKED_HISTORY' | 'DAILY_REPORT' | 'STILL_IN_ANALYTICS' | 'DOSSIER';
      data: any[];
    } | null>(null);

    useEffect(() => {
      if (checkingPrintData) {
        const timer = setTimeout(() => {
          window.focus();
          window.print();
          setCheckingPrintData(null);
        }, 500);
        return () => clearTimeout(timer);
      }
    }, [checkingPrintData]);

    useEffect(() => {
      if (activePrintPreview) {
        const triggerPrint = () => {
          window.focus();
          window.print();
        };
        const handleAfterPrint = () => {
          setActivePrintPreview(null);
        };
        const timer = setTimeout(triggerPrint, 350);
        window.addEventListener('afterprint', handleAfterPrint);
        return () => {
          clearTimeout(timer);
          window.removeEventListener('afterprint', handleAfterPrint);
        };
      }
    }, [activePrintPreview]);

    const passportToLatestInfo = useMemo(() => {
      const map: Record<string, { fullname: string; nationality: string; gender: string }> = {};
      (records || []).forEach(r => {
        if (!r || !r.passport || typeof r.passport !== 'string') return;
        const p = r.passport.toUpperCase();
        map[p] = { fullname: r.fullname || '', nationality: r.nationality || '', gender: r.gender || 'M' };
      });
      return map;
    }, [records]);

  // Form State
  const [formData, setFormData] = useState<Partial<ImmRecord>>({
    passport: '',
    previousPassport: '',
    dualPassportRemarks: '',
    fullname: '',
    gender: 'M',
    dob: '',
    nationality: '',
    address: '',
    stayDescription: '',
    visaType: '',
    visaNumber: '',
    vehicleInfo: '',
    stayFrom: '',
    stayTo: '',
    totalDays: '',
    arrivedFrom: '',
    departedTo: '',
    broughtBy: '',
    contactDetails: '',
    officialName: '',
    officialTitle: '',
    stillPermittedStatus: '',
    permittedBy: ''
  });

  // Load Initial Data from IndexedDB (merged with LocalStorage for fail-safe persistence)
  useEffect(() => {
    const loadAllData = async () => {
      try {
        await miniDB.init();

        // 1. Records
        const dbRecs: ImmRecord[] = (await miniDB.get('records')) || [];
        let lRecs: ImmRecord[] = [];
        try {
          const saved = localStorage.getItem('imm_records_react');
          if (saved) lRecs = JSON.parse(saved);
        } catch {}
        const recs = mergeByUniqueKey(dbRecs, lRecs, 'id');
        setRecords(recs);

        // 2. Temp Records
        const dbTRecs: ImmRecord[] = (await miniDB.get('tempRecords')) || [];
        let lTRecs: ImmRecord[] = [];
        try {
          const saved = localStorage.getItem('imm_temp_records_react');
          if (saved) lTRecs = JSON.parse(saved);
        } catch {}
        const tRecs = mergeByUniqueKey(dbTRecs, lTRecs, 'id');
        setTempRecords(tRecs);

        // 3. Master Data
        const dbMData: MasterItem[] = (await miniDB.get('masterData')) || [];
        let lMData: MasterItem[] = [];
        try {
          const saved = localStorage.getItem('imm_master_react');
          if (saved) lMData = JSON.parse(saved);
        } catch {}
        const mData = mergeByUniqueKey(dbMData, lMData, 'id');
        setMasterData(mData);

        // 4. Vehicle Summaries
        const dbVSums: VehicleSummary[] = (await miniDB.get('vehicleSummaries')) || [];
        let lVSums: VehicleSummary[] = [];
        try {
          const saved = localStorage.getItem('imm_vehicle_summaries');
          if (saved) lVSums = JSON.parse(saved);
        } catch {}
        const vSums = mergeByUniqueKey(dbVSums, lVSums, 'id');
        setVehicleSummaries(vSums);

        // 5. Daily PDFs
        let pdfs: Record<string, { base64: string; name: string }> = {};
        const dbPdfs = await miniDB.get('dailyPdfs');
        if (dbPdfs && typeof dbPdfs === 'object') {
          pdfs = dbPdfs;
        } else {
          const lPdfs = localStorage.getItem('imm_daily_pdfs_react');
          if (lPdfs) {
            try { pdfs = JSON.parse(lPdfs); } catch { pdfs = {}; }
          }
        }
        setDailyPdfs(pdfs);

        // 6. Checking History
        const dbCheckHist: CheckingHistoryEntry[] = (await miniDB.get('checkingHistory')) || [];
        let lCheckHist: CheckingHistoryEntry[] = [];
        try {
          const saved = localStorage.getItem('checking_verification_history_logs_v1');
          if (saved) lCheckHist = JSON.parse(saved);
        } catch {}
        let checkHist = mergeByUniqueKey(dbCheckHist, lCheckHist, 'id');
        // Normalize checking history statuses
        checkHist = checkHist.map(h => ({
          ...h,
          status: h.status === 'STILL PERMITTED' ? 'STAY PERMITTED' : (h.status === 'STILL NOT PERMITTED YET' ? 'STAY NOT PERMITTED' : h.status)
        }));
        setCheckingHistory(checkHist);

        // 7. Dossier History
        const dbDHist: DossierRecord[] = (await miniDB.get('dossierHistory')) || [];
        let lDHist: DossierRecord[] = [];
        try {
          const saved = localStorage.getItem('imm_dossier_history_react');
          if (saved) lDHist = JSON.parse(saved);
        } catch {}
        const dHist = mergeByUniqueKey(dbDHist, lDHist, 'id');
        setDossierHistory(dHist);

        // 8. Watch List
        const dbWList: WatchListRecord[] = (await miniDB.get('watchList')) || [];
        let lWList: WatchListRecord[] = [];
        try {
          const saved = localStorage.getItem('imm_watchlist_records_v1');
          if (saved) lWList = JSON.parse(saved);
        } catch {}
        const wList = mergeByUniqueKey(dbWList, lWList, 'id');
        setWatchList(wList);

        // Initialize synced hashes to baseline so initial auto-sync won't write unchanged collections
        setSyncedHash('records', recs);
        setSyncedHash('tempRecords', tRecs);
        setSyncedHash('masterData', mData);
        setSyncedHash('vehicleSummaries', vSums);
        setSyncedHash('checkingHistory', checkHist);
        setSyncedHash('dossierHistory', dHist);
        setSyncedHash('watchList', wList);

        // Pre-fill FFE with latest official info & vehicle
        if (recs.length > 0) {
          const latest = recs.sort((a: any, b: any) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp))[0];
          setFormData(prev => ({
            ...prev,
            officialName: latest.officialName || '',
            officialTitle: latest.officialTitle || '',
            vehicleInfo: latest.vehicleInfo || ''
          }));
        }

      } catch (e) {
        console.error("IndexedDB initial loading failed, falling back to clean arrays:", e);
      } finally {
        setIsDBCardLoaded(true);
        setIsInitialSyncing(false);
        if (localStorage.getItem('imm_pwa_cloud_auth_user')) {
          setTimeout(() => handleFetchDataFromCloud().catch(() => {}), 300);
        }
      }
    };

    loadAllData();
  }, []);

  // Firebase Auth Initialization Effect
  useEffect(() => {
    const unsubAuth = initAuth(() => {
      setIsCloudSynced(true);
    });

    return () => {
      if (unsubAuth) unsubAuth();
    };
  }, []);

  // Sync with Storage (IndexedDB + safe localStorage backup)
  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('records', records).catch(err => console.error("IndexedDB save records failed:", err));
    try {
      localStorage.setItem('imm_records_react', JSON.stringify(records));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [records, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('tempRecords', tempRecords).catch(err => console.error("IndexedDB save tempRecords failed:", err));
    try {
      localStorage.setItem('imm_temp_records_react', JSON.stringify(tempRecords));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [tempRecords, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('masterData', masterData).catch(err => console.error("IndexedDB save masterData failed:", err));
    try {
      localStorage.setItem('imm_master_react', JSON.stringify(masterData));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [masterData, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('vehicleSummaries', vehicleSummaries).catch(err => console.error("IndexedDB save vehicleSummaries failed:", err));
    try {
      localStorage.setItem('imm_vehicle_summaries', JSON.stringify(vehicleSummaries));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [vehicleSummaries, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('dailyPdfs', dailyPdfs).catch(err => console.error("IndexedDB save dailyPdfs failed:", err));
  }, [dailyPdfs, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('checkingHistory', checkingHistory).catch(err => console.error("IndexedDB save checkingHistory failed:", err));
    try {
      localStorage.setItem('checking_verification_history_logs_v1', JSON.stringify(checkingHistory));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [checkingHistory, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('dossierHistory', dossierHistory).catch(err => console.error("IndexedDB save dossierHistory failed:", err));
    try {
      localStorage.setItem('imm_dossier_history_react', JSON.stringify(dossierHistory));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [dossierHistory, isDBCardLoaded]);

  useEffect(() => {
    if (!isDBCardLoaded) return;
    miniDB.set('watchList', watchList).catch(err => console.error("IndexedDB save watchList failed:", err));
    try {
      localStorage.setItem('imm_watchlist_records_v1', JSON.stringify(watchList));
    } catch (err) {
      console.error("LocalStorage write failed:", err);
    }
  }, [watchList, isDBCardLoaded]);

  // 1. Immediate Auto-Sync to Cloud whenever local data is added/edited by user (Removed 20-sec delay)
  useEffect(() => {
    if (!isDBCardLoaded || !isOnline) return;

    const currentHashes: Record<string, string> = {
      records: JSON.stringify(records),
      tempRecords: JSON.stringify(tempRecords),
      masterData: JSON.stringify(masterData),
      vehicleSummaries: JSON.stringify(vehicleSummaries),
      checkingHistory: JSON.stringify(checkingHistory),
      dossierHistory: JSON.stringify(dossierHistory),
      watchList: JSON.stringify(watchList)
    };

    // Skip if change came from remote sync
    if (isRemoteSyncingRef.current) {
      prevLocalHashesRef.current = currentHashes;
      return;
    }

    // On initial DB load, set baseline hashes without starting sync
    if (Object.keys(prevLocalHashesRef.current).length === 0) {
      prevLocalHashesRef.current = currentHashes;
      return;
    }

    // Compare with previous local hashes
    const hasChanged = Object.keys(currentHashes).some(
      key => currentHashes[key] !== prevLocalHashesRef.current[key]
    );

    if (hasChanged) {
      prevLocalHashesRef.current = currentHashes;
      setIsCloudSynced(false);

      // Perform immediate sync (with a micro debounce 300ms to batch rapid keystrokes/state transitions smoothly)
      const immediateSyncTimer = setTimeout(async () => {
        if (!isOnline || isAutoSyncing) return;
        setIsAutoSyncing(true);
        try {
          const res1 = await saveCollectionToFirestore('records', records, false);
          const res2 = await saveCollectionToFirestore('tempRecords', tempRecords, false);
          const res3 = await saveCollectionToFirestore('masterData', masterData, false);
          const res4 = await saveCollectionToFirestore('vehicleSummaries', vehicleSummaries, false);
          const res5 = await saveCollectionToFirestore('checkingHistory', checkingHistory, false);
          const res6 = await saveCollectionToFirestore('dossierHistory', dossierHistory, false);
          const res7 = await saveCollectionToFirestore('watchList', watchList, false);

          if (res1 && res2 && res3 && res4 && res5 && res6 && res7) {
            setIsCloudSynced(true);
            updateLastSyncTimestamp();
            markSynced('records');
            markSynced('tempRecords');
            markSynced('masterData');
            markSynced('vehicleSummaries');
            markSynced('checkingHistory');
            markSynced('dossierHistory');
            markSynced('watchList');
          } else {
            setIsCloudSynced(false);
            if (res1) markSynced('records'); else markUploadFailed('records');
            if (res2) markSynced('tempRecords'); else markUploadFailed('tempRecords');
            if (res3) markSynced('masterData'); else markUploadFailed('masterData');
            if (res4) markSynced('vehicleSummaries'); else markUploadFailed('vehicleSummaries');
            if (res5) markSynced('checkingHistory'); else markUploadFailed('checkingHistory');
            if (res6) markSynced('dossierHistory'); else markUploadFailed('dossierHistory');
            if (res7) markSynced('watchList'); else markUploadFailed('watchList');
          }
        } catch (err) {
          console.warn("Immediate auto sync error:", err);
          setIsCloudSynced(false);
        } finally {
          setIsAutoSyncing(false);
        }
      }, 300);

      return () => clearTimeout(immediateSyncTimer);
    }
  }, [records, tempRecords, masterData, vehicleSummaries, checkingHistory, dossierHistory, watchList, isDBCardLoaded, isOnline]);

  // Emergency Unload & Instant Local Save Engine (Zero Data Loss)
  useEffect(() => {
    const emergencySave = () => {
      try {
        if (records.length > 0) {
          localStorage.setItem('imm_records_react', JSON.stringify(records));
          miniDB.set('records', records).catch(() => {});
        }
        if (tempRecords.length > 0) {
          localStorage.setItem('imm_temp_records_react', JSON.stringify(tempRecords));
          miniDB.set('tempRecords', tempRecords).catch(() => {});
        }
        if (masterData.length > 0) {
          localStorage.setItem('imm_master_react', JSON.stringify(masterData));
          miniDB.set('masterData', masterData).catch(() => {});
        }
        if (vehicleSummaries.length > 0) {
          localStorage.setItem('imm_vehicle_summaries', JSON.stringify(vehicleSummaries));
          miniDB.set('vehicleSummaries', vehicleSummaries).catch(() => {});
        }
        if (checkingHistory.length > 0) {
          localStorage.setItem('checking_verification_history_logs_v1', JSON.stringify(checkingHistory));
          miniDB.set('checkingHistory', checkingHistory).catch(() => {});
        }
        if (dossierHistory.length > 0) {
          localStorage.setItem('imm_dossier_history_react', JSON.stringify(dossierHistory));
          miniDB.set('dossierHistory', dossierHistory).catch(() => {});
        }
        if (watchList.length > 0) {
          localStorage.setItem('imm_watchlist_records_v1', JSON.stringify(watchList));
          miniDB.set('watchList', watchList).catch(() => {});
        }
      } catch (e) {
        console.error("Emergency save error:", e);
      }
    };

    const handleVis = () => {
      if (document.visibilityState === 'hidden') {
        emergencySave();
      }
    };

    window.addEventListener('beforeunload', emergencySave);
    window.addEventListener('pagehide', emergencySave);
    document.addEventListener('visibilitychange', handleVis);

    return () => {
      window.removeEventListener('beforeunload', emergencySave);
      window.removeEventListener('pagehide', emergencySave);
      document.removeEventListener('visibilitychange', handleVis);
    };
  }, [records, tempRecords, masterData, vehicleSummaries, checkingHistory, dossierHistory, watchList]);

  const analyzeAllTEH = () => {
    if (tempRecords.length === 0) return;
    
    let matchCount = 0;
    const passportsToKeep: number[] = [];

    tempRecords.forEach(tr => {
      const matchIndex = records.findIndex(r => r.passport.toUpperCase() === tr.passport.toUpperCase());
      if (matchIndex !== -1) {
        matchCount++;
      } else {
        passportsToKeep.push(tr.id);
      }
    });

    if (matchCount > 0) {
      // Remove matching ones from TEH
      setTempRecords(prev => prev.filter(r => passportsToKeep.includes(r.id)));
      showToast(`SUCCESSFULLY REMOVED ${matchCount} MATCHING RECORDS FROM TEH`);
    } else {
      showToast("NO MATCHING FH RECORDS FOUND IN TEH");
    }
  };

  // Helper: Auto-save to Master and link/update stay description/address
  const syncMaster = (value: string | undefined, type: MasterItem['type'], linkedValue?: string) => {
    if (!value || typeof value !== 'string') return;
    const val = value.trim();
    if (!val) return;
    const nowIso = new Date().toISOString();
    const cleanLinked = linkedValue && typeof linkedValue === 'string' ? linkedValue.trim() : undefined;
    const existingIndex = (masterData || []).findIndex(m => m && m.name && typeof m.name === 'string' && m.name.toLowerCase() === val.toLowerCase() && m.type === type);
    if (existingIndex !== -1) {
      if (cleanLinked !== undefined) {
        const existing = masterData[existingIndex];
        if (existing && existing.linkedValue !== cleanLinked) {
          setMasterData(prev => prev.map((m, idx) => idx === existingIndex ? { ...m, linkedValue: cleanLinked || undefined, updatedAt: nowIso, syncStatus: 'pending_sync' } : m));
        }
      }
    } else {
      setMasterData(prev => [...prev, { 
        id: Date.now() + Math.floor(Math.random() * 1000), 
        name: val, 
        type, 
        linkedValue: cleanLinked || undefined,
        updatedAt: nowIso,
        syncStatus: 'pending_sync'
      }]);
    }
  };

  // Periodic Backup File Reminder (Every 30 mins)
  useEffect(() => {
    if (!currentUser) return;
    const interval = setInterval(() => {
      setShowBackupReminderModal(true);
    }, 30 * 60 * 1000);

    return () => clearInterval(interval);
  }, [currentUser]);

  // Auto-populate INV tab investigation note into FFE form when passport is active
  useEffect(() => {
    if (formData.passport && formData.passport.trim().length >= 2) {
      const invRemark = getLatestVerificationRemark(formData.passport, dossierHistory, records);
      if (invRemark && invRemark !== '-' && !formData.remarks) {
        setFormData(prev => ({ ...prev, remarks: invRemark }));
      }
    }
  }, [formData.passport, dossierHistory, records]);

  const handlePassportInput = (val: string) => {
    try {
      const pp = (val || '').toUpperCase();
      setFormData(prev => ({ ...prev, passport: pp }));

      if (pp.length >= 2) {
        const allPool = [...(records || []), ...(tempRecords || [])];
        const matchMap = new Map<string, ImmRecord>();
        
        allPool.sort((a,b) => getRecordTime(a) - getRecordTime(b)).forEach(r => {
          if (r && r.passport && typeof r.passport === 'string' && r.passport.toUpperCase().includes(pp)) {
            matchMap.set(r.passport.toUpperCase(), r);
          }
        });
        
        const results = Array.from(matchMap.values()).reverse().slice(0, 5);
        setPassportSearchResults(results);

        // Check for exact match in history to prefill
        const exactMatch = allPool
          .sort((a, b) => getRecordTime(b) - getRecordTime(a))
          .find(r => r && r.passport && typeof r.passport === 'string' && r.passport.toUpperCase() === pp);
        const latestCheck = (checkingHistory || []).find(c => c && c.passport && typeof c.passport === 'string' && c.passport.toUpperCase() === pp);
        const invRemark = getLatestVerificationRemark(pp, dossierHistory, allPool);
        const permitStatus = latestCheck?.status 
          ? (latestCheck.status === 'STAY PERMITTED' || latestCheck.status === 'STILL PERMITTED' ? 'STAY PERMITTED' : 'STAY NOT PERMITTED')
          : (exactMatch?.stillPermittedStatus || '');
        const permitBy = latestCheck?.permittedBy || exactMatch?.permittedBy || '';

        if (exactMatch || latestCheck || (invRemark && invRemark !== '-')) {
          setFormData(prev => {
            const matchAddress = exactMatch?.address || prev.address || '';
            const cleanAddr = matchAddress.trim().toLowerCase();
            const latestStayRec = allPool
              .sort((a, b) => getRecordTime(b) - getRecordTime(a))
              .find(x => (x.formC?.address?.toLowerCase() === cleanAddr || x.address?.toLowerCase() === cleanAddr) && (x.formC?.stayDescription || x.stayDescription));
            const latestMasterStay = [...(masterData || [])].reverse().find(x => x && x.type === 'Stay' && x.name && x.name.toLowerCase() === cleanAddr && x.linkedValue);
            const stayDesc = latestStayRec?.formC?.stayDescription || latestStayRec?.stayDescription || latestMasterStay?.linkedValue || exactMatch?.stayDescription || '';

            return {
              ...prev,
              fullname: exactMatch?.fullname || prev.fullname || '',
              gender: exactMatch?.gender || prev.gender || 'M',
              dob: exactMatch?.dob ? formatToDDMMYYYY(exactMatch.dob) : (prev.dob || ''),
              nationality: exactMatch?.nationality || prev.nationality || '',
              visaType: exactMatch?.visaType || prev.visaType || '',
              visaNumber: exactMatch?.visaNumber || prev.visaNumber || '',
              address: matchAddress,
              stayDescription: stayDesc || prev.stayDescription || '',
              stayFrom: exactMatch?.stayFrom || prev.stayFrom || '',
              stayTo: exactMatch?.stayTo || prev.stayTo || '',
              totalDays: exactMatch?.totalDays != null ? String(exactMatch.totalDays) : (prev.totalDays || ''),
              stillPermittedStatus: permitStatus as any,
              permittedBy: permitBy,
              remarks: prev.remarks || (invRemark !== '-' ? invRemark : ''),
              arrivedFrom: prev.arrivedFrom || (currentMode === 'OUT' ? 'MGZ' : (exactMatch?.arrivedFrom || '')),
              departedTo: prev.departedTo || (currentMode === 'IN' ? 'MGZ' : (exactMatch?.departedTo || '')),
              broughtBy: exactMatch?.broughtBy || prev.broughtBy || '',
              contactDetails: exactMatch?.contactDetails || prev.contactDetails || ''
            };
          });
        }
      } else {
        setPassportSearchResults([]);
      }

      // Conflict detection for TEMPORARY ENTRY warning
      if (pp.length >= 3) {
        const tempMatch = (tempRecords || []).find(tr => tr && tr.passport && typeof tr.passport === 'string' && tr.passport.toUpperCase() === pp);
        if (tempMatch) {
           setTempAlert(`Passport ${pp} is already assigned on FORM C (Temporary Entry: ${tempMatch.fullname || ''})`);
        } else {
           setTempAlert(null);
        }
      } else {
        setTempAlert(null);
      }
    } catch (e) {
      console.error(e);
    }
  };

  const selectPassportMatch = (r: ImmRecord) => {
    if (!r) return;
    try {
      setFormData(prev => {
        const pp = (r.passport || '').toUpperCase();
        const rAddress = r.address || '';
        const cleanAddress = rAddress.trim().toLowerCase();
        const allPool = [...(records || []), ...(tempRecords || [])];
        const latestRecStay = allPool
          .sort((a, b) => getRecordTime(b) - getRecordTime(a))
          .find(x => (x.formC?.address?.toLowerCase() === cleanAddress || x.address?.toLowerCase() === cleanAddress) && (x.formC?.stayDescription || x.stayDescription));
        const latestMasterStay = [...(masterData || [])].reverse().find(x => x && x.type === 'Stay' && x.name && x.name.toLowerCase() === cleanAddress && x.linkedValue);
        const stayDesc = latestRecStay?.formC?.stayDescription || latestRecStay?.stayDescription || latestMasterStay?.linkedValue || r.stayDescription || '';
        
        const invRemark = getLatestVerificationRemark(pp, dossierHistory, records);
        const latestCheck = (checkingHistory || []).find(c => c && c.passport && typeof c.passport === 'string' && c.passport.toUpperCase() === pp);
        const permitStatus = latestCheck?.status 
          ? (latestCheck.status === 'STAY PERMITTED' || latestCheck.status === 'STILL PERMITTED' ? 'STAY PERMITTED' : 'STAY NOT PERMITTED')
          : (r.stillPermittedStatus || '');
        const permitBy = latestCheck?.permittedBy || r.permittedBy || '';

        return {
          ...r,
          id: undefined,
          timestamp: undefined,
          mode: currentMode,
          passport: pp,
          fullname: r.fullname || prev.fullname || '',
          gender: r.gender || prev.gender || 'M',
          dob: r.dob ? formatToDDMMYYYY(r.dob) : (prev.dob || ''),
          nationality: r.nationality || prev.nationality || '',
          visaType: r.visaType || prev.visaType || '',
          address: r.address || prev.address || '',
          stayFrom: r.stayFrom ? formatToDDMMYYYY(r.stayFrom) : (prev.stayFrom || ''),
          stayTo: r.stayTo ? formatToDDMMYYYY(r.stayTo) : (prev.stayTo || ''),
          totalDays: r.totalDays != null ? String(r.totalDays) : (prev.totalDays || ''),
          broughtBy: r.broughtBy || prev.broughtBy || '',
          contactDetails: r.contactDetails || prev.contactDetails || '',
          arrivedFrom: currentMode === 'OUT' ? 'MGZ' : (r.arrivedFrom === 'MGZ' ? '' : (r.arrivedFrom || '')),
          departedTo: currentMode === 'IN' ? 'MGZ' : (r.departedTo === 'MGZ' ? '' : (r.departedTo || '')),
          stillPermittedStatus: permitStatus as any,
          permittedBy: permitBy,
          remarks: r.remarks || (invRemark !== '-' ? invRemark : prev.remarks || ''),
          officialName: prev.officialName || '',
          officialTitle: prev.officialTitle || '',
          vehicleInfo: prev.vehicleInfo || '',
          stayDescription: stayDesc || prev.stayDescription || '',
          formC: undefined
        };
      });
      setPassportSearchResults([]);
      setNameSearchResults([]);
      setTempAlert(null);
    } catch (e) {
      console.error(e);
    }
  };

  const handleNameInput = (val: string) => {
    try {
      setFormData(prev => ({ ...prev, fullname: val }));
      const q = (val || '').toLowerCase();

      if (q.length >= 3) {
        const allPool = [...(records || []), ...(tempRecords || [])];
        const matchMap = new Map<string, ImmRecord>();
        
        allPool.sort((a,b) => getRecordTime(a) - getRecordTime(b)).forEach(r => {
          if (r && r.fullname && typeof r.fullname === 'string' && r.fullname.toLowerCase().includes(q)) {
            if (r.passport && typeof r.passport === 'string') {
              matchMap.set(r.passport.toUpperCase(), r);
            }
          }
        });
        
        const results = Array.from(matchMap.values()).reverse().slice(0, 5);
        setNameSearchResults(results);
      } else {
        setNameSearchResults([]);
      }
    } catch (e) {
      console.error(e);
    }
  };

  const selectNameMatch = (r: ImmRecord) => {
    if (!r) return;
    try {
      setFormData(prev => {
        const pp = (r.passport || '').toUpperCase();
        const rAddress = r.address || '';
        const cleanAddress = rAddress.trim().toLowerCase();
        const allPool = [...(records || []), ...(tempRecords || [])];
        const latestRecStay = allPool
          .sort((a, b) => getRecordTime(b) - getRecordTime(a))
          .find(x => (x.formC?.address?.toLowerCase() === cleanAddress || x.address?.toLowerCase() === cleanAddress) && (x.formC?.stayDescription || x.stayDescription));
        const latestMasterStay = [...(masterData || [])].reverse().find(x => x && x.type === 'Stay' && x.name && x.name.toLowerCase() === cleanAddress && x.linkedValue);
        const stayDesc = latestRecStay?.formC?.stayDescription || latestRecStay?.stayDescription || latestMasterStay?.linkedValue || r.stayDescription || '';

        const invRemark = getLatestVerificationRemark(pp, dossierHistory, records);
        const latestCheck = (checkingHistory || []).find(c => c && c.passport && typeof c.passport === 'string' && c.passport.toUpperCase() === pp);
        const permitStatus = latestCheck?.status 
          ? (latestCheck.status === 'STAY PERMITTED' || latestCheck.status === 'STILL PERMITTED' ? 'STAY PERMITTED' : 'STAY NOT PERMITTED')
          : (r.stillPermittedStatus || '');
        const permitBy = latestCheck?.permittedBy || r.permittedBy || '';

        return {
          ...r,
          id: undefined,
          timestamp: undefined,
          mode: currentMode,
          passport: pp,
          fullname: r.fullname || prev.fullname || '',
          gender: r.gender || prev.gender || 'M',
          dob: r.dob ? formatToDDMMYYYY(r.dob) : (prev.dob || ''),
          nationality: r.nationality || prev.nationality || '',
          visaType: r.visaType || prev.visaType || '',
          address: r.address || prev.address || '',
          stayFrom: r.stayFrom ? formatToDDMMYYYY(r.stayFrom) : (prev.stayFrom || ''),
          stayTo: r.stayTo ? formatToDDMMYYYY(r.stayTo) : (prev.stayTo || ''),
          totalDays: r.totalDays != null ? String(r.totalDays) : (prev.totalDays || ''),
          broughtBy: r.broughtBy || prev.broughtBy || '',
          contactDetails: r.contactDetails || prev.contactDetails || '',
          arrivedFrom: currentMode === 'OUT' ? 'MGZ' : (r.arrivedFrom === 'MGZ' ? '' : (r.arrivedFrom || '')),
          departedTo: currentMode === 'IN' ? 'MGZ' : (r.departedTo === 'MGZ' ? '' : (r.departedTo || '')),
          stillPermittedStatus: permitStatus as any,
          permittedBy: permitBy,
          remarks: r.remarks || (invRemark !== '-' ? invRemark : prev.remarks || ''),
          officialName: prev.officialName || '',
          officialTitle: prev.officialTitle || '',
          vehicleInfo: prev.vehicleInfo || '',
          stayDescription: stayDesc || prev.stayDescription || '',
          formC: undefined
        };
      });
      setPassportSearchResults([]);
      setNameSearchResults([]);
      setTempAlert(null);
    } catch (e) {
      console.error(e);
    }
  };

  // Smart Detection for Person Chain / Linked Old Passports
  const linkedPersonSuggestion = useMemo(() => {
    const currentPp = (formData.passport || '').trim().toUpperCase();
    const currentName = (formData.fullname || '').trim().toLowerCase();
    if (!currentPp && (!currentName || currentName.length < 3)) return null;

    const allPool = [...(records || []), ...(tempRecords || [])];
    
    for (const r of allPool) {
      if (!r || !r.passport) continue;
      const rPp = (r.passport || '').trim().toUpperCase();
      const rName = (r.fullname || '').trim().toLowerCase();
      const rPrev = (r.previousPassport || '').trim().toUpperCase();
      
      // Case 1: Same person name, different passport
      if (currentName && rName === currentName && rPp !== currentPp) {
        return {
          oldPassport: rPp,
          fullname: r.fullname,
          nationality: r.nationality,
          address: r.address,
          visaType: r.visaType,
          totalDays: r.totalDays,
          stayFrom: r.stayFrom,
          stayTo: r.stayTo,
          broughtBy: r.broughtBy,
          contactDetails: r.contactDetails,
          reason: 'Same Person Name Match in History'
        };
      }
      
      // Case 2: Past record had previousPassport matching current or current passport matches r.previousPassport
      if (currentPp && (rPrev === currentPp || (rPp === currentPp && rPrev))) {
        const targetOld = rPrev === currentPp ? rPp : rPrev;
        if (targetOld && targetOld !== currentPp) {
          return {
            oldPassport: targetOld,
            fullname: r.fullname,
            nationality: r.nationality,
            address: r.address,
            visaType: r.visaType,
            totalDays: r.totalDays,
            stayFrom: r.stayFrom,
            stayTo: r.stayTo,
            broughtBy: r.broughtBy,
            contactDetails: r.contactDetails,
            reason: 'Linked Person Chain Record'
          };
        }
      }
    }
    return null;
  }, [formData.passport, formData.fullname, records, tempRecords]);

  const applyLinkedPerson = (suggestion: any) => {
    if (!suggestion) return;
    setFormData(prev => ({
      ...prev,
      previousPassport: suggestion.oldPassport,
      dualPassportRemarks: prev.dualPassportRemarks || 'Renewed Passport / Same Person',
      fullname: prev.fullname || suggestion.fullname || '',
      nationality: prev.nationality || suggestion.nationality || '',
      dob: prev.dob || (suggestion.dob ? formatToDDMMYYYY(suggestion.dob) : ''),
      visaType: prev.visaType || suggestion.visaType || '',
      address: prev.address || suggestion.address || '',
      stayFrom: prev.stayFrom ? formatToDDMMYYYY(prev.stayFrom) : (suggestion.stayFrom ? formatToDDMMYYYY(suggestion.stayFrom) : ''),
      stayTo: prev.stayTo ? formatToDDMMYYYY(prev.stayTo) : (suggestion.stayTo ? formatToDDMMYYYY(suggestion.stayTo) : ''),
      totalDays: prev.totalDays || (suggestion.totalDays != null ? String(suggestion.totalDays) : ''),
      broughtBy: prev.broughtBy || suggestion.broughtBy || '',
      contactDetails: prev.contactDetails || suggestion.contactDetails || ''
    }));
    showToast(`LINKED OLD PASSPORT (${suggestion.oldPassport}) TO CURRENT PROFILE`);
  };

  const handleAgentInput = (val: string) => {
    try {
      const agentVal = val || '';
      setFormData(prev => ({ ...prev, broughtBy: agentVal }));
      if (agentVal.trim()) {
        const cleanAgent = agentVal.trim().toLowerCase();
        const allPool = [...(records || []), ...(tempRecords || [])];
        const lastRec = allPool
          .sort((a, b) => getRecordTime(b) - getRecordTime(a))
          .find(r => r && r.broughtBy && typeof r.broughtBy === 'string' && r.broughtBy.toLowerCase().trim() === cleanAgent && r.contactDetails);
        const masterMatch = [...(masterData || [])].reverse().find(m => m && m.type === 'Agent' && m.name && typeof m.name === 'string' && m.name.toLowerCase().trim() === cleanAgent && m.linkedValue);
        
        const contact = lastRec?.contactDetails || masterMatch?.linkedValue || '';
        if (contact) {
          setFormData(prev => ({ ...prev, contactDetails: contact }));
        }
      }
    } catch (e) {
      console.error(e);
    }
  };

  const handleOfficialInputMain = (val: string) => {
    try {
      const offVal = val || '';
      setFormData(prev => ({ ...prev, officialName: offVal }));
      if (!offVal.trim()) return;
      const cleanVal = offVal.trim().toLowerCase().replace(/\s+/g, '');
      
      // 1. Prioritize latest from records
      const all = [...(records || []), ...(tempRecords || [])].sort((a,b) => getRecordTime(b) - getRecordTime(a));
      const latestWithTitle = all.find(r => 
        (r && r.officialName && typeof r.officialName === 'string' && r.officialName.toLowerCase().replace(/\s+/g, '') === cleanVal && r.officialTitle) ||
        (r && r.formC?.officialName && typeof r.formC.officialName === 'string' && r.formC.officialName.toLowerCase().replace(/\s+/g, '') === cleanVal && r.formC.officialTitle)
      );
      
      // 2. Fallback to Master Data
      const masterMatch = [...(masterData || [])].reverse().find(m => m && m.type === 'Official' && m.name && typeof m.name === 'string' && m.name.toLowerCase().replace(/\s+/g, '') === cleanVal && m.linkedValue);

      const title = (latestWithTitle?.formC?.officialName?.toLowerCase().replace(/\s+/g, '') === cleanVal ? latestWithTitle?.formC?.officialTitle : latestWithTitle?.officialTitle) || masterMatch?.linkedValue || '';
      if (title) {
        setFormData(prev => ({ ...prev, officialTitle: title }));
      }
    } catch (e) {
      console.error(e);
    }
  };

  useEffect(() => {
    setFormData(prev => ({
      ...prev,
      arrivedFrom: currentMode === 'OUT' ? 'MGZ' : (prev.arrivedFrom === 'MGZ' ? '' : (prev.arrivedFrom || '')),
      departedTo: currentMode === 'IN' ? 'MGZ' : (prev.departedTo === 'MGZ' ? '' : (prev.departedTo || ''))
    }));
  }, [currentMode]);

  const calculateStayDates = (field: 'from' | 'days' | 'to', value: string) => {
    try {
      let from = formData.stayFrom || '';
      let to = formData.stayTo || '';
      let days = formData.totalDays || '';

      if (field === 'from') {
        from = value;
        if (days && from) {
          const numDays = parseInt(days, 10);
          if (!isNaN(numDays) && numDays > 0) {
            const d = new Date(from);
            if (!isNaN(d.getTime())) {
              d.setDate(d.getDate() + (numDays - 1));
              to = d.toISOString().split('T')[0];
            }
          }
        }
      } else if (field === 'days') {
        days = (value || '').replace(/[^0-9]/g, '');
        if (from && days) {
          const numDays = parseInt(days, 10);
          if (!isNaN(numDays) && numDays > 0) {
            const d = new Date(from);
            if (!isNaN(d.getTime())) {
              d.setDate(d.getDate() + (numDays - 1));
              to = d.toISOString().split('T')[0];
            }
          }
        }
      } else if (field === 'to') {
        to = value;
        if (from && to) {
          const d1 = new Date(from);
          const d2 = new Date(to);
          if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
            const diff = Math.ceil((d2.getTime() - d1.getTime()) / 86400000) + 1;
            days = diff > 0 ? `${diff} D` : "";
          }
        }
      }

      setFormData(prev => ({ ...prev, stayFrom: from, stayTo: to, totalDays: days }));
    } catch (e) {
      console.error(e);
    }
  };

  const executeSaveRecord = () => {
    try {
      const finalDob = formData.dob ? formatToDDMMYYYY(formData.dob) : undefined;
      const finalStayFrom = formData.stayFrom ? formatToDDMMYYYY(formData.stayFrom) : undefined;
      const finalStayTo = formData.stayTo ? formatToDDMMYYYY(formData.stayTo) : undefined;
      const newRecord: ImmRecord = {
        ...formData as ImmRecord,
        dob: finalDob,
        stayFrom: finalStayFrom,
        stayTo: finalStayTo,
        logType: 'FFE',
        id: editTarget ? editTarget.id : Date.now(),
        timestamp: formData.timestamp || (editTarget ? editTarget.timestamp : new Date().toLocaleString('en-GB')),
        mode: formData.mode || (editTarget ? editTarget.mode : currentMode),
        gender: (formData.gender as 'M' | 'F') || 'M',
        passport: (formData.passport || '').toUpperCase().trim(),
        previousPassport: formData.previousPassport ? formData.previousPassport.toUpperCase().trim() : undefined,
        dualPassportRemarks: formData.dualPassportRemarks ? formData.dualPassportRemarks.trim() : undefined,
        linkedPassports: formData.previousPassport ? Array.from(new Set([(formData.passport || '').toUpperCase().trim(), formData.previousPassport.toUpperCase().trim()])) : formData.linkedPassports
      };

      // Auto sync master picks
      syncMaster(newRecord.nationality, 'Nationality');
      syncMaster(newRecord.visaType, 'Visa');
      syncMaster(newRecord.address, 'Stay', formData.stayDescription);
      syncMaster(newRecord.vehicleInfo, 'Vehicle');
      syncMaster(newRecord.broughtBy, 'Agent');
      if (newRecord.contactDetails) syncMaster(newRecord.contactDetails, 'Contact');
      if (newRecord.officialName) syncMaster(newRecord.officialName, 'Official');
      if (newRecord.officialTitle) syncMaster(newRecord.officialTitle, 'Title');
      if (newRecord.formC?.reporterName) syncMaster(newRecord.formC.reporterName, 'Reporter');
      if (newRecord.formC?.reporterPhone) syncMaster(newRecord.formC.reporterPhone, 'Phone');

      const activeDevName = cloudAuthUser?.deviceName || (typeof navigator !== 'undefined' && navigator.userAgent.includes('Mobile') ? 'Mobile Phone' : 'Desktop Device');
      const recordToSave: ImmRecord = {
        ...newRecord,
        syncStatus: 'pending_sync',
        createdDevice: activeDevName,
        updatedAt: new Date().toISOString()
      };

      let nextRecordsList: ImmRecord[] = [];
      if (editTarget) {
        nextRecordsList = records.map(r => r.id === editTarget.id ? recordToSave : r);
        setRecords(nextRecordsList);
        setEditTarget(null);
        logActivity({
          action: 'UPDATE',
          module: recordToSave.logType === 'FCR' ? 'FCR' : 'FFE',
          officerName: recordToSave.officialName || (recordToSave.officialTitle && recordToSave.officialName ? `${recordToSave.officialTitle} ${recordToSave.officialName}` : cloudAuthUser?.username),
          officerRole: cloudAuthUser?.role || 'Editor',
          targetId: recordToSave.passport,
          details: `Updated ${recordToSave.logType || 'FFE'} [${recordToSave.mode}] record #${recordToSave.id} for Passport ${recordToSave.passport} (${recordToSave.fullname || 'N/A'}, ${recordToSave.nationality || 'N/A'}) - Flight/Vehicle: ${recordToSave.vehicleInfo || '-'}`
        });
      } else {
        nextRecordsList = [recordToSave, ...records];
        setRecords(nextRecordsList);
        // Delete from temporary entries if passport matches
        setTempRecords(prev => prev.filter(tr => tr && tr.passport && typeof tr.passport === 'string' && tr.passport.toUpperCase() !== recordToSave.passport.toUpperCase()));
        logActivity({
          action: 'CREATE',
          module: recordToSave.logType === 'FCR' ? 'FCR' : 'FFE',
          officerName: recordToSave.officialName || (recordToSave.officialTitle && recordToSave.officialName ? `${recordToSave.officialTitle} ${recordToSave.officialName}` : cloudAuthUser?.username),
          officerRole: cloudAuthUser?.role || 'Editor',
          targetId: recordToSave.passport,
          details: `New entry created: ${recordToSave.logType || 'FFE'} [${recordToSave.mode}] Passport ${recordToSave.passport} (${recordToSave.fullname || 'N/A'}, ${recordToSave.nationality || 'N/A'}) via ${recordToSave.vehicleInfo || '-'}`
        });
      }

      // Direct instant upload to Cloud Server for zero-confusion verification
      if (isOnline) {
        showToast("📡 Cloud Server ပေါ်သို့ တိုက်ရိုက် ပို့ဆောင်နေပါသည်...");
        saveCollectionToFirestore('records', nextRecordsList, true, true).then((ok) => {
          if (ok) {
            const confirmedAt = new Date().toISOString();
            setRecords(prev => prev.map(r => (r.id === recordToSave.id || (r.passport === recordToSave.passport && r.timestamp === recordToSave.timestamp)) ? { ...r, syncStatus: 'synced', serverSyncedAt: confirmedAt } : r));
            showToast("✓ Cloud Server ပေါ်သို့ အောင်မြင်စွာ ရောက်ရှိ သိမ်းဆည်းပြီးပါပြီ (Server Verified ✓)");
          } else {
            showToast("💾 ဖုန်းထဲတွင် သိမ်းဆည်းပြီးပါပြီ (လတ်တလော ဆာဗာမရောက်သေးပါ - Auto-Sync ပို့ပါမည်)");
          }
        }).catch(() => {
          showToast("💾 ဖုန်းထဲတွင် သိမ်းဆည်းပြီးပါပြီ (လတ်တလော ဆာဗာမရောက်သေးပါ - Auto-Sync ပို့ပါမည်)");
        });
      } else {
        showToast("💾 အင်တာနက်မရှိပါ - ဖုန်းထဲတွင် လုံခြုံစွာသိမ်းထားပြီး လိုင်းရပါက Auto-Sync ပြုလုပ်ပေးပါမည်");
      }

      // Automatically sync checkpoint status to checkingHistory as checked
      if (newRecord.stillPermittedStatus) {
        const isCoAddressLocal = (addr: string) => {
          const a = (addr || '').toLowerCase();
          return a.includes('hotel') || a.includes('co ltd') || a.includes('co.,ltd') || a.includes('company') || a.includes('guesthouse') || a.includes('motel') || a.includes('inn');
        };
        const isCo = isCoAddressLocal(newRecord.address || '');
        const autoHistEntry = {
          id: Date.now(),
          passport: (newRecord.passport || '').toUpperCase(),
          fullname: newRecord.fullname || '',
          nationality: newRecord.nationality || '',
          previousPassport: newRecord.previousPassport,
          dualPassportRemarks: newRecord.dualPassportRemarks,
          linkedPassports: newRecord.linkedPassports,
          type: isCo ? 'CO_LTD' : 'OTHERS' as any,
          checkDate: (newRecord.timestamp || '').split(', ')[0] || new Date().toLocaleDateString('en-GB'),
          originalAddress: newRecord.address || '',
          confirmedAddress: newRecord.address || '',
          confirmedStayDescription: newRecord.stayDescription || '',
          status: newRecord.stillPermittedStatus,
          permittedBy: newRecord.permittedBy || '',
          officerName: newRecord.officialName || 'Checkpoint Duty',
          officerTitle: newRecord.officialTitle || 'Officer',
          timestamp: newRecord.timestamp || new Date().toLocaleString('en-GB')
        };
        setCheckingHistory(prev => {
          const filtered = (prev || []).filter(h => h && h.passport && typeof h.passport === 'string' && h.passport.toUpperCase() !== (newRecord.passport || '').toUpperCase());
          return [autoHistEntry, ...filtered];
        });
      }

      const successMsg = currentMode === 'IN' 
        ? `${newRecord.officialTitle || ''} ${newRecord.officialName || ''} စာရင်းသွင်းသော လေဆိပ် အဝင်စာရင်း စာရင်းသွင်းပြီးပါပြီ`
        : `${newRecord.officialTitle || ''} ${newRecord.officialName || ''} စာရင်းသွင်းသော လေဆိပ် အထွက်စာရင်းအား စာရင်းသွင်းပြီးပါပြီ`;
      
      setFfeSuccessMsg(successMsg);
      setTimeout(() => setFfeSuccessMsg(null), 3000);

      setFfeConfirmation(null);
      resetForm();
      showToast("RECORD SAVED");
      setActiveTab('data');
    } catch (e) {
      console.error(e);
      showToast("Error saving record");
    }
  };

  const saveRecord = (e: React.FormEvent) => {
    e.preventDefault();
    if (isViewer) {
      showToast("🔒 Viewer level အကောင့်ဖြစ်သောကြောင့် အချက်အလက်သစ် ဖြည့်သွင်း/ပြင်ဆင်ခွင့် ပိတ်ထားပါသည်");
      return;
    }
    if (!formData.passport || !formData.fullname) {
      showToast("Required: Passport & Full Name");
      return;
    }

    const mode = editTarget ? (formData.mode || editTarget.mode) : currentMode;
    const logDate = editTarget ? formatDateToDDMMYYYY(editTarget.timestamp) : formatDateToDDMMYYYY(new Date().toISOString());
    const stayToFormatted = formatDateToDDMMYYYY(formData.stayTo || "");

    let message = "";
    if (mode === 'IN') {
      message = `လေဆိပ်မှ ${logDate}ရက်နေ့တွင် ${formData.arrivedFrom || ''} မှ ${formData.departedTo || ''}သို့ ${formData.vehicleInfo || ''}ဖြင့် ဝင်ရောက်လာ သူ အား ${formData.officialTitle || ''} ${formData.officialName || ''} မှ စာရင်းသွင်း ခြင်းဖြစ်ပါသည်၊ ${stayToFormatted} မတိုင်ခင်ထိ ${formData.address || ''} တွင် ${formData.visaType || ''} ဗီဇာဖြင့် နေထိုင်မည့် ${formData.nationality || ''}နိုင်ငံသား ${formData.fullname || ''} ဖြစ်ပါသည်။ ${formData.broughtBy || ''} မှ လာရောက် ခေါ်ဆောင်သွားပါသည်။`;
    } else {
      message = `လေဆိပ်မှ ${logDate}ရက်နေ့တွင် ${formData.arrivedFrom || ''} မှ ${formData.departedTo || ''}သို့ ${formData.vehicleInfo || ''}ဖြင့် ထွက်ခွါသွား သူ အား ${formData.officialTitle || ''} ${formData.officialName || ''} မှ စာရင်းသွင်း ခြင်းဖြစ်ပါသည်၊ ${stayToFormatted} မတိုင်ခင်ထိ ${formData.address || ''} တွင် ${formData.visaType || ''} ဗီဇာဖြင့် နေထိုင်ခဲ့သည့် ${formData.nationality || ''}နိုင်ငံသား ${formData.fullname || ''} ဖြစ်ပါသည်။ ${formData.broughtBy || ''} မှ လာရောက် ပို့ဆောင်သွားပါသည်။`;
    }

    setFfeConfirmation({ show: true, message });
  };

  const resetForm = () => {
    setFormData(prev => ({
      passport: '', previousPassport: '', dualPassportRemarks: '', fullname: '', gender: 'M', dob: '', nationality: '',
      address: '', visaType: '', visaNumber: '', stayFrom: '',
      stayTo: '', totalDays: '', 
      arrivedFrom: currentMode === 'OUT' ? 'MGZ' : '', 
      departedTo: currentMode === 'IN' ? 'MGZ' : '',
      broughtBy: '', contactDetails: '',
      officialName: prev.officialName,
      officialTitle: prev.officialTitle,
      vehicleInfo: prev.vehicleInfo,
      formC: undefined, // Ensure formC is cleared
      stillPermittedStatus: '',
      permittedBy: ''
    }));
    setEditTarget(null);
  };

  // --- FILTERED DATA ---
  const filteredRecords = useMemo(() => {
    let filtered = records.filter(r => r.logType !== 'FCR');
    if (globalDateRange.from) {
      const start = new Date(globalDateRange.from).setHours(0, 0, 0, 0);
      filtered = filtered.filter(r => parseTimestamp(r.timestamp) >= start);
    }
    if (globalDateRange.to) {
      const end = new Date(globalDateRange.to).setHours(23, 59, 59, 999);
      filtered = filtered.filter(r => parseTimestamp(r.timestamp) <= end);
    }
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(r => 
        (r && r.passport && typeof r.passport === 'string' && r.passport.toLowerCase().includes(q)) || 
        (r && r.fullname && typeof r.fullname === 'string' && r.fullname.toLowerCase().includes(q))
      );
    }
    return filtered;
  }, [records, globalDateRange, searchQuery]);

  // --- MOVEMENT LOGIC ---
  const movementMap = useMemo(() => {
    const map: Record<string, MovementData> = {};
    (records || [])
      .filter(r => r && r.logType !== 'FCR' && r.passport && typeof r.passport === 'string')
      .sort((a, b) => {
        const diff = parseTimestamp(a?.timestamp || '') - parseTimestamp(b?.timestamp || '');
        if (diff !== 0) return diff;
        if (a?.mode === 'IN' && b?.mode === 'OUT') return -1;
        if (a?.mode === 'OUT' && b?.mode === 'IN') return 1;
        return 0;
      })
      .forEach(r => {
      if (!r || !r.passport || typeof r.passport !== 'string') return;
      const pp = r.passport.toUpperCase();
      if (!map[pp]) {
        map[pp] = { 
          p: pp, n: r.fullname || '', nat: r.nationality || '', loc: r.address || '', 
          visa: r.visaType || '', visaNumber: r.visaNumber || '', dob: r.dob || '', gender: r.gender || 'M', start: r.stayFrom || '', 
          end: r.stayTo || '', allowed: r.totalDays || '', vInfo: r.vehicleInfo || '', 
          agent: r.broughtBy || '',
          contact: r.contactDetails || '',
          offName: r.formC?.officialName || r.officialName || '',
          offTitle: r.formC?.officialTitle || r.officialTitle || '',
          repName: r.formC?.reporterName || '',
          repPhone: r.formC?.reporterPhone || '',
          in: '', out: '', inTime: 0, outTime: 0, 
          latestTime: 0, lastId: r.id 
        };
      }
      const recTime = parseTimestamp(r.timestamp);
      if (r.mode === 'IN') {
        Object.assign(map[pp], { 
          in: r.timestamp, inTime: recTime, out: '', 
          outTime: 0, loc: r.address, visa: r.visaType, 
          visaNumber: r.visaNumber || map[pp].visaNumber || '',
          dob: r.dob || map[pp].dob || '',
          agent: r.broughtBy || '',
          contact: r.contactDetails || map[pp].contact,
          offName: r.formC?.officialName || r.officialName || map[pp].offName,
          offTitle: r.formC?.officialTitle || r.officialTitle || map[pp].offTitle,
          repName: r.formC?.reporterName || map[pp].repName,
          repPhone: r.formC?.reporterPhone || map[pp].repPhone,
          start: r.stayFrom, end: r.stayTo, allowed: r.totalDays, 
          lastId: r.id 
        });
      } else if (r.mode === 'OUT' && recTime >= map[pp].inTime) {
        Object.assign(map[pp], {
          out: r.timestamp,
          outTime: recTime,
          lastId: r.id,
          offName: r.formC?.officialName || r.officialName || map[pp].offName,
          offTitle: r.formC?.officialTitle || r.officialTitle || map[pp].offTitle,
          repName: r.formC?.reporterName || map[pp].repName,
          repPhone: r.formC?.reporterPhone || map[pp].repPhone,
        });
      }
      if (recTime > map[pp].latestTime) {
        map[pp].latestTime = recTime;
        // Always try to get the latest meta info
        if (r.broughtBy) map[pp].agent = r.broughtBy;
        if (r.contactDetails) map[pp].contact = r.contactDetails;
      }
    });

    // Enforce real-time name & nationality synchronization across the entirety of movement logs
    Object.keys(map).forEach(pp => {
      if (passportToLatestInfo[pp]) {
        map[pp].n = passportToLatestInfo[pp].fullname;
        map[pp].nat = passportToLatestInfo[pp].nationality;
        map[pp].gender = passportToLatestInfo[pp].gender as any;
      }
    });

    return map;
  }, [records, passportToLatestInfo]);

  // --- TAB RENDERERS ---

  const handleLogin = () => {
    const raw = loginName.trim();
    if (!raw) {
      showToast("ကျေးဇူးပြု၍ စာရင်းသွင်းအရာရှိ အမည် ရိုက်ထည့်ပါ");
      return;
    }
    const val = raw.toLowerCase().replace(/\s+/g, '');
    const match = masterData.find(m => m.type === 'Official' && m.name.toLowerCase().replace(/\s+/g, '') === val);
    
    if (!match) {
      showToast("⚠️ စာရင်းသွင်းအရာရှိ အမည် မမှန်ကန်ပါ (Master DB တွင် မတွေ့ရှိပါ)");
      return;
    }

    const officialName = match.name;
    let finalTitle = match.linkedValue && match.linkedValue.trim() ? match.linkedValue.trim() : 'Officer';

    if (!match.linkedValue) {
      const all = [...records, ...tempRecords].sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
      const latestWithTitle = all.find(r => 
        (r.officialName?.toLowerCase().replace(/\s+/g, '') === val && r.officialTitle) ||
        (r.formC?.officialName?.toLowerCase().replace(/\s+/g, '') === val && r.formC.officialTitle)
      );
      if (latestWithTitle) {
        if (latestWithTitle.formC?.officialName?.toLowerCase().replace(/\s+/g, '') === val) {
          finalTitle = latestWithTitle.formC.officialTitle;
        } else {
          finalTitle = latestWithTitle.officialTitle || 'Officer';
        }
      } else {
        const titleMatch = masterData.find(m => m.type === 'Title' && m.name.trim() !== '');
        if (titleMatch) finalTitle = titleMatch.name;
      }
    }

    const user = { name: officialName, title: finalTitle };
    setCurrentUser(user);
    setFormData(prev => ({ ...prev, officialName: officialName, officialTitle: finalTitle }));

    localStorage.setItem('lastCheckedOfficerName', officialName);
    localStorage.setItem('lastCheckedOfficerTitle', finalTitle);
    showToast(`မင်္ဂလာပါ ${finalTitle} ${officialName} ဝင်ရောက်လိုက်ပါပြီ`);
  };



  const renderNav = () => (
    <>
      <nav className="sticky top-0 z-40 transition-all duration-300 shadow-lg bg-[#1B365D] text-white py-1.5 px-3 sm:px-4 no-print border-b border-white/10">
        <div className="max-w-[1700px] mx-auto flex items-center justify-between gap-2">
          {/* BRAND & TITLE SECTION */}
          <div className="flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-white/10 p-1.5 rounded-xl backdrop-blur-md border border-white/15 shrink-0">
              <ShieldCheck size={20} className="text-emerald-400" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <h1 className="text-xs sm:text-sm font-black tracking-tight uppercase text-white leading-none truncate">
                  DEPARTMENT OF IMMIGRATION
                </h1>
                <span className="hidden sm:inline-block text-[9px] bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 px-1.5 py-0.2 rounded-full font-bold">
                  v10.0
                </span>
              </div>
              <div className="flex items-center gap-2 text-[10px] text-indigo-200 font-bold tracking-wider leading-none mt-0.5 truncate">
                <span>SECOND DIVISION (MYEIK) • PWA</span>
              </div>
            </div>
          </div>

          {/* MIDDLE STATUS BADGE (COMPACT POPOVER TRIGGER) */}
          <div className="relative">
            <button
              onClick={() => {
                setIsSyncModalOpen(true);
              }}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-[10px] font-black tracking-wide border backdrop-blur-sm transition-all cursor-pointer shadow-xs"
              title="Click to view Cloud Server & All Devices Sync Hub"
            >
              {downloadProgress?.isDownloading ? (
                <div className="flex items-center gap-1.5 bg-cyan-500/25 text-cyan-200 border-cyan-400/50 px-2 py-0.5 rounded-lg animate-pulse" title={`Cloud Data Downloading (${downloadProgress.currentCol})`}>
                  <Download size={11} className="animate-bounce text-cyan-300" />
                  <span className="font-extrabold text-[10px]">Cloud Down: <span className="font-mono text-cyan-200">{downloadProgress.percent}%</span></span>
                </div>
              ) : !isOnline ? (
                <div className="flex items-center gap-1 bg-rose-500/20 text-rose-300 border-rose-400/30 px-1.5 py-0.5 rounded-lg">
                  <WifiOff size={11} className="text-rose-400" />
                  <span className="hidden xs:inline">Offline</span>
                </div>
              ) : isAutoSyncing ? (
                <div className="flex items-center gap-1 bg-blue-500/20 text-blue-200 border-blue-400/40 px-1.5 py-0.5 rounded-lg animate-pulse" title="Cloud သို့ ချက်ချင်း Sync ပြုလုပ်နေပါသည် (Instant Cloud Sync)">
                  <RefreshCw size={11} className="animate-spin text-blue-300" />
                  <span className="hidden xs:inline">Syncing...</span>
                </div>
              ) : records.some(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed') ? (
                <div className="flex items-center gap-1 bg-amber-500/25 text-amber-200 border-amber-400/50 px-2 py-0.5 rounded-lg animate-pulse" title="ဆာဗာသို့ မရောက်သေးသော စာရင်းများ ရှိပါသည် (နှိပ်၍ စစ်ဆေး/ပို့ဆောင်ပါ)">
                  <Clock size={11} className="text-amber-300" />
                  <span>{records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed').length} ခု မရောက်သေး</span>
                </div>
              ) : isCloudSynced ? (
                <div className="flex items-center gap-1 bg-emerald-500/20 text-emerald-300 border-emerald-400/30 px-1.5 py-0.5 rounded-lg" title={lastSyncTime ? `နောက်ဆုံး Cloud Sync ဖိုင်၏ အချိန်: ${lastSyncTime}` : 'Cloud Synced'}>
                  <CheckCircle2 size={11} className="text-emerald-400" />
                  <span className="hidden xs:inline">Server: Synced ({records.length})</span>
                </div>
              ) : (
                <div className="flex items-center gap-1 bg-amber-500/20 text-amber-300 border-amber-400/30 px-1.5 py-0.5 rounded-lg">
                  <HardDrive size={11} className="text-amber-400" />
                  <span className="hidden xs:inline">Saved Locally</span>
                </div>
              )}

              {isQuotaExhausted && (
                <span className="bg-rose-500/30 text-rose-200 border border-rose-400/50 px-1.5 py-0.5 rounded-md text-[9px] font-black flex items-center gap-1 animate-pulse">
                  <AlertCircle size={9} />
                  <span>Quota</span>
                </span>
              )}
              <ChevronDown size={11} className={`text-slate-300 transition-transform ${isStatusPopoverOpen ? 'rotate-180' : ''}`} />
            </button>

          {/* STATUS POPOVER PANEL */}
          <AnimatePresence>
            {isStatusPopoverOpen && (
              <motion.div
                initial={{ opacity: 0, scale: 0.95, y: 6 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 6 }}
                className="absolute left-1/2 -translate-x-1/2 mt-2 w-72 bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl overflow-hidden z-50 p-3 text-white text-xs space-y-2.5"
              >
                <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                  <span className="font-black text-[10px] uppercase text-indigo-400 tracking-wider flex items-center gap-1">
                    <ShieldCheck size={13} /> System & Cloud Status
                  </span>
                  <button onClick={() => setIsStatusPopoverOpen(false)} className="text-slate-400 hover:text-white p-0.5">
                    <X size={13} />
                  </button>
                </div>

                <div className="space-y-1.5 text-[11px]">
                  <div className="flex justify-between items-center bg-slate-800/60 p-2 rounded-xl">
                    <span className="text-slate-400 font-bold">Network Connection:</span>
                    <span className={`font-black flex items-center gap-1 ${isOnline ? 'text-emerald-400' : 'text-rose-400'}`}>
                      {isOnline ? <Wifi size={12} /> : <WifiOff size={12} />}
                      {isOnline ? 'Online' : 'Offline'}
                    </span>
                  </div>

                  <div className="flex justify-between items-center bg-slate-800/60 p-2 rounded-xl">
                    <span className="text-slate-400 font-bold">Cloud Sync State:</span>
                    <span className={`font-black ${isCloudSynced ? 'text-emerald-400' : 'text-amber-400'}`}>
                      {isCloudSynced ? '✓ Synced to Cloud' : '💾 Saved Locally (Cloud Sync Pending)'}
                    </span>
                  </div>

                  <div className="flex justify-between items-center bg-slate-800/60 p-2 rounded-xl">
                    <span className="text-slate-400 font-bold">နောက်ဆုံး Sync အချိန် (Last Sync):</span>
                    <span className="font-mono font-bold text-indigo-300 text-[10px]">
                      {lastSyncTime || 'မပြုလုပ်ရသေးပါ'}
                    </span>
                  </div>

                  <div className="flex justify-between items-center bg-slate-800/60 p-2 rounded-xl">
                    <span className="text-slate-400 font-bold">Quota Status:</span>
                    <span className={`font-black ${isQuotaExhausted ? 'text-rose-400' : 'text-emerald-400'}`}>
                      {isQuotaExhausted ? '⚠️ Limit Reached' : 'Normal Active'}
                    </span>
                  </div>
                </div>

                <button
                  onClick={() => {
                    setIsStatusPopoverOpen(false);
                    setIsSyncModalOpen(true);
                  }}
                  className="w-full bg-indigo-600 hover:bg-indigo-500 text-white font-black py-2 px-3 rounded-xl text-[11px] uppercase flex items-center justify-center gap-1.5 shadow-sm transition-all cursor-pointer"
                >
                  <Server size={13} /> Open Server & Devices Monitor
                </button>

                {isQuotaExhausted && (
                  <button
                    onClick={() => {
                      resetQuotaState();
                      handleManualSync();
                      setIsStatusPopoverOpen(false);
                    }}
                    className="w-full bg-amber-600 hover:bg-amber-500 text-white font-black py-1.5 px-2 rounded-xl text-[10px] uppercase flex items-center justify-center gap-1 shadow-sm transition-all"
                  >
                    <RefreshCw size={11} /> Reset & Retry Cloud Sync
                  </button>
                )}
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* RIGHT ACTION CONTROLS & PROFILE */}
        <div className="flex items-center gap-1.5 sm:gap-2">
          {/* SERVER & DEVICES MONITOR BUTTON */}
          <button
            onClick={() => setIsSyncModalOpen(true)}
            className="bg-white/10 hover:bg-white/20 border border-white/15 px-2.5 py-1 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer text-xs font-bold text-slate-100 shadow-xs"
            title="ဆာဗာဒေတာနှင့် ချိတ်ဆက်ထားသော စက်များ စစ်ဆေးရန် (Server & Devices Monitor)"
          >
            <Server size={13} className="text-cyan-300" />
            <span className="hidden md:inline">Server & Devices</span>
            {records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed').length > 0 && (
              <span className="bg-amber-500 text-slate-900 text-[9px] font-black px-1.5 py-0.2 rounded-full">
                {records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed').length}
              </span>
            )}
          </button>
          {/* OFFICER PROFILE AVATAR & DROPDOWN */}
          <div className="relative">
            <button
              onClick={() => {
                setIsUserDropdownOpen(!isUserDropdownOpen);
                setIsSettingsOpen(false);
                setIsStatusPopoverOpen(false);
              }}
              className="bg-white/10 hover:bg-white/20 border border-white/15 px-2 py-1 rounded-xl flex items-center gap-1.5 transition-all cursor-pointer"
              title="User Officer Profile"
            >
              <div className="w-6 h-6 rounded-lg bg-indigo-500 flex items-center justify-center text-[11px] font-black text-white shrink-0 shadow-xs border border-indigo-300/40">
                {currentUser ? currentUser.name.charAt(0).toUpperCase() : <User size={13} />}
              </div>
              <span className="hidden sm:inline text-[11px] font-bold text-slate-100 max-w-[110px] md:max-w-[150px] truncate">
                {currentUser ? `${currentUser.title} ${currentUser.name}` : 'အရာရှိ မသတ်မှတ်ရသေး'}
              </span>
              <ChevronDown size={12} className={`text-slate-300 transition-transform ${isUserDropdownOpen ? 'rotate-180' : ''}`} />
            </button>

            {/* OFFICER PROFILE DROPDOWN */}
            <AnimatePresence>
              {isUserDropdownOpen && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.95, y: 6 }}
                  animate={{ opacity: 1, scale: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95, y: 6 }}
                  className="absolute right-0 mt-2 w-64 bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl overflow-hidden z-50 p-2 text-white space-y-2"
                >
                  <div className="p-2 border-b border-slate-800 bg-slate-800/40 rounded-xl">
                    <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Current Active Officer</div>
                    <div className="text-xs font-black text-white mt-0.5 truncate">
                      {currentUser ? `${currentUser.title} ${currentUser.name}` : 'မသတ်မှတ်ရသေးပါ'}
                    </div>
                    {cloudAuthUser && (
                      <div className="text-[10px] text-emerald-400 font-bold mt-1 flex items-center gap-1">
                        <span>Role:</span> {cloudAuthUser.username} ({cloudAuthUser.role})
                      </div>
                    )}
                  </div>

                  {currentUser ? (
                    <button
                      onClick={() => {
                        setCurrentUser(null);
                        localStorage.removeItem('lastCheckedOfficerName');
                        localStorage.removeItem('lastCheckedOfficerTitle');
                        setOfficerNameInput('');
                        setOfficerTitleInput('');
                        setIsUserDropdownOpen(false);
                        showToast("စာရင်းသွင်းအရာရှိ အကောင့်မှ ထွက်လိုက်ပါပြီ");
                      }}
                      className="w-full bg-rose-600/80 hover:bg-rose-600 text-white text-xs font-black py-2 rounded-xl flex items-center justify-center gap-1.5 transition-all shadow-sm cursor-pointer"
                    >
                      <LogOut size={13} /> ထွက်မည် (Logout)
                    </button>
                  ) : (
                    <button
                      onClick={() => {
                        setIsUserDropdownOpen(false);
                        setShowOfficerLoginModal(true);
                      }}
                      className="w-full bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-black py-2 rounded-xl flex items-center justify-center gap-1.5 transition-all shadow-sm cursor-pointer"
                    >
                      <User size={13} /> စာရင်းသွင်းအရာရှိ ရွေးမည်
                    </button>
                  )}
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* SETTINGS & ZOOM ICON MENU */}
          <div className="relative">
            <button
              onClick={() => {
                setIsSettingsOpen(!isSettingsOpen);
                setIsUserDropdownOpen(false);
                setIsStatusPopoverOpen(false);
              }}
              className="p-1.5 rounded-xl bg-white/10 hover:bg-white/20 border border-white/15 text-slate-200 hover:text-white transition-all cursor-pointer"
              title="Display & Cloud Settings"
            >
              <Settings size={15} />
            </button>

            {/* SETTINGS DROPDOWN */}
            <AnimatePresence>
              {isSettingsOpen && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.95, y: 6 }}
                  animate={{ opacity: 1, scale: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95, y: 6 }}
                  className="absolute right-0 mt-2 w-64 bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl overflow-hidden z-50 p-3 text-white space-y-3"
                >
                  <div className="flex justify-between items-center border-b border-slate-800 pb-2">
                    <span className="text-[10px] font-black uppercase text-indigo-400 tracking-wider flex items-center gap-1">
                      <SlidersHorizontal size={12} /> System Settings
                    </span>
                    <button onClick={() => setIsSettingsOpen(false)} className="text-slate-400 hover:text-white p-0.5">
                      <X size={13} />
                    </button>
                  </div>

                  {/* Font Scale Selector */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">UI Zoom / Font Scale:</label>
                    <select
                      value={fontScale}
                      onChange={e => setFontScale(e.target.value)}
                      className="w-full bg-slate-800 text-white text-xs font-black px-2.5 py-1.5 rounded-xl border border-slate-700 outline-none cursor-pointer"
                    >
                      <option value="80">80% (Extra Small)</option>
                      <option value="85">85%</option>
                      <option value="90">90% (Compact)</option>
                      <option value="95">95%</option>
                      <option value="100">100% (Standard)</option>
                      <option value="105">105%</option>
                      <option value="110">110%</option>
                      <option value="115">115%</option>
                      <option value="120">120%</option>
                      <option value="125">125%</option>
                      <option value="130">130% (Large)</option>
                    </select>
                  </div>

                  {/* Cloud Auth Controls */}
                  <div className="border-t border-slate-800 pt-2 space-y-1.5">
                    <div className="text-[10px] font-bold text-slate-400 uppercase">Cloud Account:</div>
                    {cloudAuthUser ? (
                      <div className="flex items-center justify-between bg-slate-800/80 p-2 rounded-xl">
                        <span className="text-xs font-black text-emerald-300 truncate">{cloudAuthUser.username}</span>
                        <button
                          onClick={() => {
                            handleCloudLogout();
                            setIsSettingsOpen(false);
                          }}
                          className="bg-rose-600 hover:bg-rose-500 text-white text-[10px] font-black px-2 py-1 rounded-lg flex items-center gap-1 cursor-pointer"
                        >
                          <LogOut size={10} /> Logout
                        </button>
                      </div>
                    ) : (
                      <button
                        onClick={() => {
                          setLoginAuthError(null);
                          setShowCloudLoginModal(true);
                          setIsSettingsOpen(false);
                        }}
                        className="w-full bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black py-1.5 rounded-xl flex items-center justify-center gap-1.5 cursor-pointer"
                      >
                        <LogIn size={13} /> Cloud Login
                      </button>
                    )}
                  </div>

                  {/* PWA Install Button */}
                  {pwaInstallPrompt && (
                    <div className="border-t border-slate-800 pt-2">
                      <button
                        onClick={() => {
                          pwaInstallPrompt.prompt();
                          pwaInstallPrompt.userChoice.then(() => setPwaInstallPrompt(null));
                          setIsSettingsOpen(false);
                        }}
                        className="w-full bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black py-1.5 rounded-xl flex items-center justify-center gap-1.5 shadow-sm cursor-pointer"
                      >
                        <Smartphone size={13} /> Install PWA
                      </button>
                    </div>
                  )}
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* MAIN MENU HAMBURGER BUTTON */}
          <div className="relative">
            <button
              onClick={() => {
                setIsMainMenuOpen(prev => !prev);
                setIsUserDropdownOpen(false);
                setIsSettingsOpen(false);
                setIsStatusPopoverOpen(false);
              }}
              className="p-1.5 sm:px-3 sm:py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-black text-xs flex items-center gap-1.5 shadow-md border border-indigo-400/30 transition-all cursor-pointer"
              title="Main Navigation Menu"
            >
              <Menu size={16} />
              <span className="hidden sm:inline text-[11px] uppercase tracking-wider font-black">Menu</span>
              <ChevronDown size={12} className={`transition-transform duration-200 ${isMainMenuOpen ? 'rotate-180' : ''}`} />
            </button>

            {/* Dropdown Menu Panel */}
            <AnimatePresence>
              {isMainMenuOpen && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.95, y: 6 }}
                  animate={{ opacity: 1, scale: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95, y: 6 }}
                  className="absolute right-0 mt-2 w-80 bg-slate-900 border border-slate-700 rounded-2xl shadow-2xl overflow-hidden z-50 p-2 text-white"
                >
                  <div className="p-2.5 border-b border-slate-800 flex justify-between items-center">
                    <span className="text-[10px] font-black uppercase text-indigo-400 tracking-widest">Main Application Views</span>
                    <button onClick={() => setIsMainMenuOpen(false)} className="text-slate-400 hover:text-white p-1 cursor-pointer">
                      <X size={14} />
                    </button>
                  </div>

                  <div className="grid grid-cols-1 gap-1 py-1.5 max-h-[70vh] overflow-y-auto">
                    {[
                      { id: 'entry', label: 'Flight Form Entry (FFE)', short: 'FFE', icon: PlusCircle },
                      { id: 'data', label: 'Flight History (FH)', short: 'FH', icon: History },
                      { id: 'movement', label: 'Movement Logs (M)', short: 'M', icon: Truck },
                      { id: 'stillIn', label: 'Statistical Analysis (SIA)', short: 'SIA', icon: Activity },
                      { id: 'checking', label: 'Checking Console (CHK)', short: 'CHK', icon: CheckSquare },
                      { id: 'individualSearch', label: 'Investigation Dossier (INV)', short: 'INV', icon: Search },
                      { id: 'master', label: 'Master DB Management', short: 'MD', icon: Database },
                      { id: 'tableOutput', label: 'Table Output', short: 'TO', icon: Table },
                      { id: 'customReport', label: 'Custom Report Table (CRT)', short: 'CRT', icon: FileSpreadsheet },
                      { id: 'telegraph', label: 'Telegraph and Tables (T&T)', short: 'T&T', icon: ClipboardList },
                      { id: 'watchList', label: 'Watch-list (WL)', short: 'WL', icon: ShieldAlert },
                      ...(cloudAuthUser?.role === 'Superadmin' ? [
                        { id: 'auditLog', label: 'Activity & Audit Log (Superadmin)', short: 'AUDIT', icon: Activity }
                      ] : []),
                      { id: 'info', label: 'System & Quota Info', short: 'INFO', icon: ShieldCheck },
                    ].map(t => {
                      const IconComp = t.icon;
                      const isActive = activeTab === t.id;
                      return (
                        <button
                          key={t.id}
                          onClick={() => {
                            switchTab(t.id as Tab);
                            setIsMainMenuOpen(false);
                          }}
                          className={`w-full text-left px-3 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between cursor-pointer ${
                            isActive
                              ? 'bg-indigo-600 text-white shadow-md'
                              : 'text-slate-300 hover:bg-slate-800 hover:text-white'
                          }`}
                        >
                          <div className="flex items-center gap-2.5">
                            <IconComp size={15} className={isActive ? 'text-white' : 'text-indigo-400'} />
                            <span>{t.label}</span>
                          </div>
                          <span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${
                            isActive ? 'bg-white/20 text-white' : 'bg-slate-800 text-slate-400'
                          }`}>
                            {t.short}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </div>
    </nav>
  </>
);

  const switchTab = (t: Tab) => {
    setActiveTab(t);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const renderInfoTab = () => {
    const currentDeviceId = localStorage.getItem('imm_pwa_device_id') || 'local_device_terminal';

    return (
      <motion.div initial={{ opacity: 0, y: 15 }} animate={{ opacity: 1, y: 0 }} className="max-w-7xl mx-auto py-8 px-4 space-y-8">
        {/* Header Banner */}
        <div className="bg-slate-900 text-white p-8 rounded-3xl shadow-xl border border-slate-800 flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
          <div className="flex items-center gap-3">
            <span className="p-3 bg-indigo-600/30 text-indigo-400 rounded-2xl border border-indigo-500/30">
              <ShieldCheck size={32} />
            </span>
            <div>
              <h2 className="text-2xl font-black uppercase tracking-tight">System & Device Info</h2>
              <p className="text-xs text-slate-400 font-semibold mt-0.5">
                အသုံးပြုနေသော စက်အချက်အလက်များနှင့် Cloud Account ထိန်းချုပ်မှု
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="bg-slate-800 border border-slate-700 px-4 py-2 rounded-2xl text-right">
              <span className="text-[10px] text-slate-400 font-bold uppercase block">Current Active Role</span>
              <span className="text-sm font-black text-emerald-400 uppercase">{cloudAuthUser ? cloudAuthUser.role : 'Offline Viewer'}</span>
            </div>
            <button
              onClick={handleCloudLogout}
              className="bg-rose-600 hover:bg-rose-700 text-white text-xs font-black px-4 py-2.5 rounded-2xl flex items-center gap-2 shadow-lg transition-all"
            >
              <LogOut size={16} /> Cloud Logout (အကောင့်ထွက်မည်)
            </button>
          </div>
        </div>

        {/* Device & Cloud Account Info Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {/* Active Terminal Info */}
          <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-200 shadow-lg space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <h3 className="text-base font-black uppercase text-slate-900 flex items-center gap-2">
                <Smartphone className="text-indigo-600" size={20} />
                အသုံးပြုနေသော စက်အချက်အလက် (Current Device Info)
              </h3>
              <span className="px-2.5 py-1 bg-emerald-100 text-emerald-800 rounded-full text-[10px] font-black flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span>
                ACTIVE TERMINAL
              </span>
            </div>

            <div className="space-y-3.5 text-xs font-bold text-slate-700">
              <div className="flex justify-between items-center bg-slate-50 p-3 rounded-2xl border border-slate-200/80">
                <span className="text-slate-500 font-semibold">Device Name:</span>
                <span className="font-extrabold text-indigo-950">{cloudAuthUser?.deviceName || 'Local Duty Terminal'}</span>
              </div>
              <div className="flex justify-between items-center bg-slate-50 p-3 rounded-2xl border border-slate-200/80">
                <span className="text-slate-500 font-semibold">Device ID:</span>
                <span className="font-mono text-[11px] text-slate-700 bg-slate-200/60 px-2 py-0.5 rounded">{currentDeviceId}</span>
              </div>
              <div className="flex justify-between items-center bg-slate-50 p-3 rounded-2xl border border-slate-200/80">
                <span className="text-slate-500 font-semibold">App System Version:</span>
                <span className="text-emerald-700 font-black">V 10.0 PWA (Offline Fail-Safe)</span>
              </div>
              <div className="flex justify-between items-center bg-slate-50 p-3 rounded-2xl border border-slate-200/80">
                <span className="text-slate-500 font-semibold">Local DB Storage:</span>
                <span className="text-blue-700 font-black">IndexedDB Active (Fail-Safe Local Persistence)</span>
              </div>
            </div>
          </div>

          {/* Cloud Account & Quota Optimization Card */}
          <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-200 shadow-lg space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <h3 className="text-base font-black uppercase text-slate-900 flex items-center gap-2">
                <UserCheck className="text-emerald-600" size={20} />
                Cloud Account & Quota Status
              </h3>
            </div>

            <div className="space-y-4">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3 text-xs font-bold">
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Account Username:</span>
                  <span className="text-slate-900 font-extrabold">{cloudAuthUser?.username || 'Not Logged In'}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Role Privilege:</span>
                  <span className="text-emerald-700 font-black">{cloudAuthUser?.role || 'None'}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Quota Status:</span>
                  {isQuotaExhausted ? (
                    <span className="text-rose-600 bg-rose-100 border border-rose-300 px-2.5 py-0.5 rounded-full font-black animate-pulse">
                      ⚠️ Limit Reached (Local DB Active)
                    </span>
                  ) : (
                    <span className="text-emerald-700 bg-emerald-100 border border-emerald-300 px-2.5 py-0.5 rounded-full font-black">
                      ✓ Normal Operational
                    </span>
                  )}
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-slate-500">Quota Saver Mode:</span>
                  <span className="text-blue-600 font-black">Active (0 Background Listener Reads)</span>
                </div>
              </div>

              <div className={isQuotaExhausted ? "bg-rose-50 border border-rose-300 rounded-2xl p-4 text-xs space-y-2" : "bg-emerald-50 border border-emerald-200 rounded-2xl p-4 text-xs space-y-2"}>
                <div className={isQuotaExhausted ? "flex items-center gap-2 text-rose-900 font-black uppercase" : "flex items-center gap-2 text-emerald-900 font-black uppercase"}>
                  <ShieldCheck size={16} /> {isQuotaExhausted ? 'Quota Limit Detected — Local Fail-Safe Active' : 'Offline Fail-Safe & Auto Sync'}
                </div>
                <p className="text-slate-700 text-[11px] leading-relaxed font-semibold">
                  {isQuotaExhausted
                    ? 'လက်ရှိတွင် Firestore Quota ပြည့်သွားသဖြင့် ဒေတာအသစ်များကို Local DB (IndexedDB) တွင် စိတ်ချစွာ သိမ်းဆည်းပေးနေပါသည်။ Quota ပြန် reset ဖြစ်ချိန် သို့မဟုတ် Cloud ဆာဗာ အဆင်ပြေချိန်တွင် Auto-Sync ပြုလုပ်ပေးပါမည်။'
                    : 'အင်တာနက် သို့မဟုတ် Quota ကုန်နေချိန်တွင်လည်း ဒေတာအသစ်များကို Local DB တွင် လုံခြုံစွာ သိမ်းဆည်းပေးထားပါသည်။ Quota ပြန် reset ဖြစ်ချိန် သို့မဟုတ် အင်တာနက် မိချိန်တွင် Cloud သို့ Auto-Sync ပြုလုပ်ပေးပါမည်။'}
                </p>
              </div>

              <div className="flex items-center gap-2 pt-1">
                <button
                  onClick={() => {
                    resetQuotaState();
                    showToast("Quota အခြေအနေအား Reset ပြုလုပ်ပြီးပါပြီ");
                  }}
                  className="flex-1 py-2.5 bg-slate-800 hover:bg-slate-900 text-white font-extrabold text-[11px] rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <RefreshCw size={14} />
                  <span>Reset Quota Status</span>
                </button>
                <button
                  onClick={() => {
                    setQuotaExhausted(true);
                    showToast("Quota Exhausted Warning အား စမ်းသပ်ပြသနေပါသည်");
                  }}
                  className="py-2.5 px-3 bg-amber-500 hover:bg-amber-600 text-slate-950 font-black text-[11px] rounded-xl transition-all flex items-center justify-center gap-1 cursor-pointer shadow-sm"
                  title="Quota Limit Banner ကို UI တွင် စမ်းသပ်ရန်"
                >
                  <span>Test Limit UI</span>
                </button>
              </div>
            </div>
          </div>

          {/* Active Connected User Sessions & Multi-Device Control (Superadmin Access Revocation) */}
          <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-200 shadow-lg space-y-5 col-span-1 md:col-span-2">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between border-b border-slate-100 pb-4 gap-3">
              <div>
                <h3 className="text-base font-black uppercase text-slate-900 flex items-center gap-2">
                  <Users className="text-indigo-600" size={20} />
                  ချိတ်ဆက်ထားသော အကောင့်များနှင့် စက်များ ထိန်းချုပ်မှု (Active Users & Device Session Management)
                </h3>
                <p className="text-xs text-slate-500 font-medium mt-1">
                  {cloudAuthUser?.role === 'Superadmin'
                    ? 'Superadmin စီမံခန့်ခွဲခွင့် - ချိတ်ဆက်ထားသော မည်သည့်စက် (Device) ကိုမဆို Viewer (Read-Only) သို့ ပြောင်းလိုက်ပါက အဆိုပါစက်သည် Logout ထွက်စရာမလိုဘဲ ချက်ချင်း View-Only Mode သို့ အလိုအလျောက် ပြောင်းလဲသွားပါမည်။'
                    : 'လက်ရှိ အကောင့်နှင့် ချိတ်ဆက်ထားသော စက်များ စာရင်း'}
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={async () => {
                    showToast("🔄 Active Sessions စာရင်းအား Cloud မှ ရယူနေပါသည်...");
                    const fresh = await fetchDeviceSessions();
                    if (Array.isArray(fresh) && fresh.length > 0) {
                      setDeviceSessions(fresh);
                      localStorage.setItem('imm_pwa_device_sessions', JSON.stringify(fresh));
                      showToast("✅ Active Sessions စာရင်း ရယူပြီးပါပြီ");
                    } else {
                      showToast("ℹ️ Active Sessions မရှိသေးပါ (သို့မဟုတ် Local Data သာရှိသည်)");
                    }
                  }}
                  className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-bold border border-slate-300 flex items-center gap-1 cursor-pointer transition-all"
                  title="Cloud မှ Active Sessions စာရင်း အသစ်ပြန်လည် ရယူရန်"
                >
                  <RefreshCw size={12} /> Sync Active
                </button>
                <span className="px-3 py-1 bg-indigo-50 text-indigo-700 rounded-full text-xs font-black border border-indigo-200">
                  {deviceSessions.length} Connected Sessions
                </span>
              </div>
            </div>

            <div className="space-y-3">
              {deviceSessions.map((session, idx) => (
                <div 
                  key={session.deviceId || idx}
                  className={`p-4 rounded-2xl border transition-all flex flex-col md:flex-row items-start md:items-center justify-between gap-4 ${
                    session.kicked 
                      ? 'bg-rose-50/60 border-rose-200 text-rose-900 opacity-70' 
                      : 'bg-slate-50 border-slate-200/80 hover:border-indigo-200'
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-black text-sm shrink-0 ${
                      session.kicked ? 'bg-rose-200 text-rose-800' : 'bg-indigo-100 text-indigo-700'
                    }`}>
                      {session.accountRole === 'Superadmin' ? <Shield size={18} /> : <UserCheck size={18} />}
                    </div>
                    <div>
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-extrabold text-sm text-slate-900">{session.username || 'USER'}</span>
                        
                        {/* Editable Role Selector */}
                        <div className="flex items-center gap-1">
                          {cloudAuthUser?.role === 'Superadmin' ? (
                            <select
                              value={session.accountRole || 'Editor'}
                              onChange={(e) => handleUpdateDeviceRole(session.deviceId, e.target.value as CloudRole)}
                              className="text-[10px] font-black uppercase px-2 py-0.5 rounded-lg border border-slate-300 bg-white text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer shadow-2xs"
                              title="Role ကို ပြောင်းလဲပါ (Viewer သို့ အဆင့်လျှော့ချနိုင်ပါသည်)"
                            >
                              <option value="Superadmin">👑 Superadmin</option>
                              <option value="Editor">✏️ Editor</option>
                              <option value="Viewer">👁️ Viewer (Read-Only)</option>
                            </select>
                          ) : (
                            <span className={`text-[10px] font-black uppercase px-2 py-0.5 rounded-full ${
                              session.accountRole === 'Superadmin' 
                                ? 'bg-purple-100 text-purple-800 border border-purple-200' 
                                : session.accountRole === 'Viewer'
                                ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                : 'bg-blue-100 text-blue-800 border border-blue-200'
                            }`}>
                              {session.accountRole}
                            </span>
                          )}
                        </div>

                        {session.kicked ? (
                          <span className="text-[10px] font-black uppercase bg-rose-100 text-rose-700 px-2 py-0.5 rounded-full border border-rose-300">
                            REVOKED (ထုတ်ပစ်ထားသည်)
                          </span>
                        ) : (
                          <span className="text-[10px] font-black uppercase bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full border border-emerald-300 flex items-center gap-1">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping" /> ACTIVE
                          </span>
                        )}
                      </div>
                      <div className="text-xs text-slate-500 font-semibold mt-1 flex flex-wrap items-center gap-x-4 gap-y-1">
                        <span>💻 {session.deviceName}</span>
                        <span className="font-mono text-[10px] bg-slate-200/60 text-slate-700 px-1.5 py-0.5 rounded">ID: {session.deviceId}</span>
                        <span>⏰ Heartbeat: {session.lastActive}</span>
                      </div>
                    </div>
                  </div>

                  {cloudAuthUser?.role === 'Superadmin' && session.deviceId !== localStorage.getItem('imm_pwa_device_id') && (
                    <div className="flex items-center gap-2 flex-wrap self-end md:self-center">
                      {!session.kicked ? (
                        <button
                          onClick={() => handleKickDevice(session.deviceId)}
                          className="py-1.5 px-3 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs rounded-xl shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
                          title="အဆိုပါ အကောင့်/စက်အား စနစ်သုံးစွဲခွင့် ပိတ်ပစ်မည် (Kick User)"
                        >
                          <LogOut size={13} /> Kick Device
                        </button>
                      ) : (
                        <>
                          <button
                            onClick={() => handleReauthorizeDevice(session.deviceId)}
                            className="py-1.5 px-3 bg-slate-800 hover:bg-slate-900 text-white font-black text-xs rounded-xl shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
                          >
                            <RefreshCw size={13} /> Re-Authorize
                          </button>
                          <button
                            onClick={() => handleDeleteDeviceSession(session.deviceId)}
                            className="py-1.5 px-3 bg-rose-800 hover:bg-rose-950 text-white font-black text-xs rounded-xl shadow-xs transition-all flex items-center gap-1.5 cursor-pointer border border-rose-900"
                            title="Kicked စက်အား စနစ်မှ အပြီးသတ် ဖျက်ထုတ်မည်"
                          >
                            <Trash2 size={13} /> Delete
                          </button>
                        </>
                      )}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>
      </motion.div>
    );
  };

  const renderEntryForm = () => (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-7xl mx-auto py-8 px-4">
      <div className="card border-t-8 border-[#2C6CB0]">
        {isViewer && (
          <div className="mb-6 p-4 bg-amber-50 border-2 border-amber-300 text-amber-900 rounded-2xl flex items-center gap-3 font-bold text-xs shadow-xs">
            <AlertCircle size={20} className="text-amber-600 shrink-0" />
            <div>
              <div className="font-black text-sm uppercase">🔒 Viewer Level Mode Active</div>
              <div className="text-[11px] font-medium text-amber-800 mt-0.5">
                သင်သည် Viewer အဆင့်ဖြစ်သောကြောင့် အချက်အလက်များ ကြည့်ရှုခွင့်သာ ရရှိပါမည်။ အချက်အလက်သစ် ဖြည့်သွင်းခြင်းနှင့် ပြင်ဆင်ခြင်းများကို ပိတ်ပင်ထားပါသည် (Read-Only Mode)
              </div>
            </div>
          </div>
        )}
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-8 border-b pb-6 gap-4">
          <div>
            <h2 className={`text-2xl font-black uppercase flex items-center gap-3 ${currentMode === 'IN' ? 'text-emerald-700' : 'text-orange-600'}`}>
              {currentMode === 'IN' ? <LogIn className="animate-pulse" /> : <LogOut className="animate-pulse" />}
              {editTarget ? 'Modify Record' : `Flight Log Manager`}
            </h2>
            <div className="flex items-center gap-4 mt-2">
               <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-widest ${currentMode === 'IN' ? 'bg-emerald-100 text-emerald-700' : 'bg-orange-100 text-orange-700'}`}>
                 Current Status: {currentMode}
               </span>
               <p className="text-[10px] font-bold text-slate-400 border-l pl-4 uppercase tracking-widest">Entry System • Second Division</p>
            </div>
          </div>

          {!editTarget && (
            <div className="flex items-center gap-3 bg-slate-100 p-2 rounded-2xl border-2 border-slate-200">
              <button 
                type="button"
                onClick={() => { setCurrentMode('IN'); setFormData(p => ({ ...p, mode: 'IN' })); }}
                className={`flex items-center gap-3 px-8 py-3 rounded-xl font-black text-xs transition-all ${currentMode === 'IN' ? 'bg-emerald-600 text-white shadow-xl scale-110' : 'text-slate-500 hover:text-slate-700'}`}
              >
                <LogIn size={18} /> IN
              </button>
              <button 
                type="button"
                onClick={() => { setCurrentMode('OUT'); setFormData(p => ({ ...p, mode: 'OUT' })); }}
                className={`flex items-center gap-3 px-8 py-3 rounded-xl font-black text-xs transition-all ${currentMode === 'OUT' ? 'bg-orange-600 text-white shadow-xl scale-110' : 'text-slate-500 hover:text-slate-700'}`}
              >
                <LogOut size={18} /> OUT
              </button>
            </div>
          )}

          {editTarget && (
            <div className="flex items-center gap-2 bg-yellow-50 text-yellow-800 px-3 py-1.5 rounded-lg border border-yellow-200 text-xs font-bold animate-bounce">
              <AlertCircle size={14} /> EDITING MODE ENABLED
            </div>
          )}
        </div>

        <form onSubmit={saveRecord} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">
          <div className="space-y-4">
            <div>
              <label className="input-label text-slate-700 font-bold">Passport No. (ပတ်စပို့နံပါတ်)</label>
              <div className="relative group">
                <input 
                  type="text" 
                  value={formData.passport || ''}
                  onChange={(e) => handlePassportInput(e.target.value)}
                  className={`input-field pl-10 h-11 ${tempAlert ? 'border-amber-400 bg-amber-50' : 'border-indigo-200'} text-gray-900 font-black`}
                  required
                  placeholder="Enter Passport..."
                />
                <Database size={16} className={`absolute left-3 top-3.5 ${tempAlert ? 'text-amber-500' : 'text-indigo-400'}`} />
                {formData.passport && (
                  <button 
                    type="button"
                    onClick={() => handlePassportInput('')}
                    className="absolute right-3 top-3 text-slate-400 hover:text-red-500 transition-colors"
                  >
                    <X size={18} />
                  </button>
                )}
              </div>

              {passportSearchResults.length > 0 && (
                <div className="mt-1 bg-white border border-indigo-100 rounded-lg shadow-xl overflow-hidden z-50 relative">
                  {passportSearchResults.map(r => (
                    <button
                      key={r.id || r.passport}
                      type="button"
                      onClick={() => selectPassportMatch(r)}
                      className="w-full text-left p-3 hover:bg-indigo-50 border-b border-indigo-50 last:border-0 transition-colors flex items-center justify-between"
                    >
                      <div>
                        <div className="font-black text-xs text-indigo-900">{(r.passport || '').toUpperCase()}</div>
                        <div className="text-[10px] text-gray-500 font-bold">{r.fullname || ''}</div>
                      </div>
                      <div className="text-[8px] bg-indigo-100 text-indigo-600 px-1.5 py-0.5 rounded font-black uppercase">
                        {r.nationality || ''}
                      </div>
                    </button>
                  ))}
                </div>
              )}

              {tempAlert && (
                <div className="mt-2 p-2 bg-amber-100 text-amber-900 text-[10px] font-black uppercase rounded-lg border border-amber-200 flex items-center gap-2 animate-in slide-in-from-top-1">
                  <AlertCircle size={14} />
                  {tempAlert}
                </div>
              )}
            </div>
            <div className="relative group">
              <label className="input-label text-slate-700 font-bold">အမည် (Full Name)</label>
              <div className="relative">
                <input 
                  type="text" 
                  value={formData.fullname || ''}
                  onChange={(e) => handleNameInput(e.target.value)}
                  className="input-field pl-10 h-11 text-gray-900 font-bold" 
                  required 
                  placeholder="Enter Full Name..."
                />
                <Users size={16} className="absolute left-3 top-3.5 text-indigo-400" />
                {formData.fullname && (
                  <button 
                    type="button"
                    onClick={() => handleNameInput('')}
                    className="absolute right-3 top-3 text-slate-400 hover:text-red-500 transition-colors"
                  >
                    <X size={18} />
                  </button>
                )}
              </div>

              {nameSearchResults.length > 0 && (
                <div className="mt-1 bg-white border border-indigo-100 rounded-lg shadow-xl overflow-hidden z-50 relative">
                  {nameSearchResults.map(r => (
                    <button
                      key={r.id || r.passport}
                      type="button"
                      onClick={() => selectNameMatch(r)}
                      className="w-full text-left p-3 hover:bg-indigo-50 border-b border-indigo-50 last:border-0 transition-colors flex items-center justify-between"
                    >
                      <div>
                        <div className="font-black text-xs text-indigo-900">{r.fullname || ''}</div>
                        <div className="text-[10px] text-gray-500 font-bold font-mono">{(r.passport || '').toUpperCase()}</div>
                      </div>
                      <div className="text-[8px] bg-indigo-100 text-indigo-600 px-1.5 py-0.5 rounded font-black uppercase">
                        {r.nationality || ''}
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div>
              <label className="input-label">Gender (ကျား/မ)</label>
              <select 
                value={formData.gender || 'M'}
                onChange={(e) => setFormData(prev => ({ ...prev, gender: e.target.value as 'M' | 'F' }))}
                className="input-field"
              >
                <option value="M">M (ကျား)</option>
                <option value="F">F (မ)</option>
              </select>
            </div>
            <div>
              <DobNumpadInput
                value={formData.dob || ''}
                onChange={(val) => setFormData(prev => ({ ...prev, dob: val }))}
              />
            </div>

            {/* Previous Passport and Dual Passport Remarks */}
            <div className="bg-purple-50/50 border border-purple-200/80 p-3 rounded-2xl space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="input-label text-purple-900 font-black text-xs flex items-center gap-1.5 mb-0">
                  <Link2 size={14} className="text-purple-600" />
                  <span>Previous / Old PP (ယခင် ပတ်စပို့ဟောင်း)</span>
                </label>
                {formData.previousPassport && (
                  <span className="text-[8px] bg-purple-200 text-purple-900 font-mono font-black px-1.5 py-0.5 rounded">
                    LINKED
                  </span>
                )}
              </div>
              <div className="relative">
                <input
                  type="text"
                  value={formData.previousPassport || ''}
                  onChange={(e) => setFormData(prev => ({ ...prev, previousPassport: e.target.value.toUpperCase() }))}
                  className="input-field pl-9 h-10 bg-white border-purple-200 text-purple-950 font-mono font-bold text-xs uppercase"
                  placeholder="Old Passport No. (if renewed)..."
                />
                <Link2 size={15} className="absolute left-3 top-3 text-purple-400" />
                {formData.previousPassport && (
                  <button
                    type="button"
                    onClick={() => setFormData(prev => ({ ...prev, previousPassport: '' }))}
                    className="absolute right-3 top-2.5 text-purple-400 hover:text-red-500 transition-colors"
                  >
                    <X size={16} />
                  </button>
                )}
              </div>
              <div>
                <input
                  type="text"
                  value={formData.dualPassportRemarks || ''}
                  onChange={(e) => setFormData(prev => ({ ...prev, dualPassportRemarks: e.target.value }))}
                  className="input-field bg-white border-purple-200 text-slate-800 font-medium text-xs h-9"
                  placeholder="Dual Passport Remarks (မှတ်ချက်)..."
                />
              </div>
            </div>

            {/* Smart Person Chain Suggestion Alert */}
            {linkedPersonSuggestion && !formData.previousPassport && (
              <div className="p-3 bg-gradient-to-br from-indigo-50 to-purple-50 border border-indigo-200 rounded-2xl space-y-2 animate-in slide-in-from-top-1 shadow-xs">
                <div className="flex items-center justify-between">
                  <span className="font-black text-indigo-900 text-xs flex items-center gap-1.5">
                    <Link2 size={14} className="text-indigo-600" />
                    <span>🔗 စနစ်ဟောင်းမှ ယခင် Passport တွေ့ရှိသည်</span>
                  </span>
                  <span className="font-mono text-[10px] font-black bg-indigo-200 text-indigo-900 px-2 py-0.5 rounded">
                    {linkedPersonSuggestion.oldPassport}
                  </span>
                </div>
                <div className="text-[11px] text-slate-700 leading-tight">
                  <strong className="text-indigo-950">{linkedPersonSuggestion.fullname}</strong> ({linkedPersonSuggestion.nationality || 'Foreigner'}) ၏ ယခင်အချက်အလက်များနှင့် ပတ်စပို့ဟောင်း ချိတ်ဆက်ရန်
                </div>
                <button
                  type="button"
                  onClick={() => applyLinkedPerson(linkedPersonSuggestion)}
                  className="w-full text-center py-2 px-3 bg-[#1A365D] hover:bg-black text-white rounded-xl text-[10px] font-black uppercase tracking-wider transition-all shadow-xs flex items-center justify-center gap-1.5"
                >
                  <Link2 size={13} />
                  Old Passport ({linkedPersonSuggestion.oldPassport}) နှင့် Auto-Link ချိတ်ဆက်မည်
                </button>
              </div>
            )}
          </div>

          <div className="space-y-4">
            <datalist id="natList">{(masterData || []).filter(m => m && m.type === 'Nationality' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="visaList">{(masterData || []).filter(m => m && m.type === 'Visa' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="stayList">{(masterData || []).filter(m => m && m.type === 'Stay' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="vehicleList">{(masterData || []).filter(m => m && m.type === 'Vehicle' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="locationList">{(masterData || []).filter(m => m && m.type === 'Stay' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>

            <div>
              <label className="input-label">နိုင်ငံသား (Nationality)</label>
              <input list="natList" value={formData.nationality || ''} onChange={(e) => setFormData(prev => ({ ...prev, nationality: e.target.value }))} className="input-field" placeholder="Select or type..." />
            </div>
            <div>
              <label className="input-label">ဗီဇာအမျိုးအစား (Visa Type)</label>
              <input list="visaList" value={formData.visaType || ''} onChange={(e) => setFormData(prev => ({ ...prev, visaType: e.target.value }))} className="input-field" placeholder="Select or type..." />
            </div>
            <div>
              <label className="input-label text-slate-700 font-bold">Visa Number (ဗီဇာနံပါတ် / အမှတ်)</label>
              <input 
                type="text" 
                value={formData.visaNumber || ''} 
                onChange={(e) => setFormData(prev => ({ ...prev, visaNumber: e.target.value.toUpperCase() }))} 
                className="input-field uppercase text-gray-900 font-bold" 
                placeholder="Enter Visa Number (e.g. EV-123456)..." 
              />
            </div>
            <div>
              <label className="input-label">နေထိုင်မည့်နေရာ (Stay Location / Company)</label>
              <input 
                list="stayList" 
                value={formData.address || ''} 
                onChange={(e) => {
                  const val = e.target.value || '';
                  const cleanAddr = val.trim().toLowerCase();
                  const allPool = [...(records || []), ...(tempRecords || [])];
                  const latestStayRec = allPool
                    .sort((a, b) => getRecordTime(b) - getRecordTime(a))
                    .find(x => (x.formC?.address?.toLowerCase() === cleanAddr || x.address?.toLowerCase() === cleanAddr) && (x.formC?.stayDescription || x.stayDescription));
                  const latestMaster = [...(masterData || [])].reverse().find(m => m && m.type === 'Stay' && m.name && m.name.toLowerCase() === cleanAddr && m.linkedValue);
                  const desc = latestStayRec?.formC?.stayDescription || latestStayRec?.stayDescription || latestMaster?.linkedValue || '';
                  setFormData(prev => ({ 
                    ...prev, 
                    address: val,
                    stayDescription: desc || prev.stayDescription
                  }));
                }} 
                className="input-field" 
                placeholder="Select or type..." 
              />
            </div>
            <div>
              <label className="input-label text-indigo-900 font-extrabold">🏡 Detail Address / Stay Description (လိပ်စာအသေးစိတ်)</label>
              <input 
                type="text" 
                value={formData.stayDescription || ''} 
                readOnly
                className="input-field bg-slate-100 border-slate-300 text-slate-500 font-medium cursor-not-allowed select-none" 
                placeholder="Auto-populated detail stay description..." 
              />
            </div>

            <div className="p-3.5 bg-indigo-50/50 rounded-xl border border-indigo-200/60 mt-2 space-y-3 animate-in fade-in duration-300">
              <div>
                <label className="input-label font-black text-indigo-900 text-[10px] uppercase tracking-wider flex items-center gap-1">
                  ခွင့်ပြုချက် အခြေအနေ (Permit Status)
                </label>
                <select 
                  value={normalizePermitStatus(formData.stillPermittedStatus) || ''} 
                  onChange={(e) => setFormData(prev => ({ 
                    ...prev, 
                    stillPermittedStatus: (e.target.value as PermitStatus) || ''
                  }))} 
                  className="input-field bg-white border-indigo-300 text-xs font-black text-indigo-900"
                >
                  <option value="">-- အခြေအနေ ရွေးချယ်ရန် --</option>
                  <option value="နေထိုင်ခွင့်ကျထားသောသူ">🟢 နေထိုင်ခွင့်ကျထားသောသူ</option>
                  <option value="နေထိုင်ခွင့် မလျှောက်ထားသေးသူ">🟠 နေထိုင်ခွင့် မလျှောက်ထားသေးသူ</option>
                </select>
              </div>
              {(formData.stillPermittedStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' || formData.stillPermittedStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' || formData.stillPermittedStatus === 'STAY PERMITTED' || formData.stillPermittedStatus === 'STAY NOT PERMITTED') && (
                <div className="animate-in slide-in-from-top-1 duration-200">
                  <label className="input-label font-black text-indigo-900 text-[10px] uppercase tracking-wider">
                    {normalizePermitStatus(formData.stillPermittedStatus) === 'နေထိုင်ခွင့်ကျထားသောသူ' ? '📝 ခွင့်ပြုသူ / အကြောင်းအရာ (Permitted By / Description)' : '📝 အကြောင်းပြချက် / မှတ်ချက် (Reason / Remarks)'}
                  </label>
                  <input 
                    type="text" 
                    list="permitDescriptionList"
                    value={formData.permittedBy || ''} 
                    onChange={(e) => setFormData(prev => ({ ...prev, permittedBy: e.target.value }))} 
                    className="input-field bg-white border-indigo-300 font-bold text-xs" 
                    placeholder={normalizePermitStatus(formData.stillPermittedStatus) === 'နေထိုင်ခွင့်ကျထားသောသူ' ? "ခွင့်ပြုသူ အမည် သို့မဟုတ် နေထိုင်ခွင့်အကြောင်းအရာ..." : "ခွင့်မပြုသေးသည့် အကြောင်းပြချက် သို့မဟုတ် မှတ်ချက်..."} 
                  />
                </div>
              )}
            </div>
          </div>

          {/* Column 3: Stay Period, Vehicle Info & Route */}
          <div className="space-y-4">
            <StayPeriodInput
              stayFrom={formData.stayFrom || ''}
              stayTo={formData.stayTo || ''}
              totalDays={formData.totalDays || ''}
              onChange={(updates) => setFormData(prev => ({ ...prev, ...updates }))}
            />
            <div>
              <label className="input-label">Vehicle Info (ကားနံပါတ်)</label>
              <div className="relative">
                <input 
                  type="text" 
                  list="vehicleList"
                  value={formData.vehicleInfo || ''} 
                  onChange={(e) => setFormData(prev => ({ ...prev, vehicleInfo: e.target.value }))} 
                  className="input-field pl-10" 
                  placeholder="e.g. 1A/1234" 
                />
                <Truck size={16} className="absolute left-3 top-3 text-gray-400" />
              </div>
            </div>
            <div>
              <label className="input-label">Route (Arrived From - To)</label>
              <div className="flex gap-2">
                <input 
                  type="text" 
                  value={formData.arrivedFrom || ''} 
                  onChange={(e) => setFormData(prev => ({ ...prev, arrivedFrom: e.target.value }))} 
                  className={`input-field ${currentMode === 'OUT' ? 'bg-gray-100 font-black cursor-not-allowed text-emerald-700' : ''}`} 
                  placeholder="From" 
                  readOnly={currentMode === 'OUT'}
                />
                <ChevronRight size={18} className="mt-2 text-gray-300" />
                <input 
                  type="text" 
                  value={formData.departedTo || ''} 
                  onChange={(e) => setFormData(prev => ({ ...prev, departedTo: e.target.value }))} 
                  className={`input-field ${currentMode === 'IN' ? 'bg-gray-100 font-black cursor-not-allowed text-orange-700' : ''}`} 
                  placeholder="To" 
                  readOnly={currentMode === 'IN'}
                />
              </div>
            </div>
          </div>

          {/* Column 4: Investigation Note / Remarks */}
          <div className="space-y-4">
            <div className="bg-amber-50/60 border border-amber-200 p-3.5 rounded-2xl space-y-2.5 h-full flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between gap-1 mb-1.5">
                  <label className="input-label text-slate-900 font-black text-[10px] uppercase tracking-wider flex items-center gap-1.5 mb-0">
                    <FileText size={14} className="text-amber-600 shrink-0" />
                    <span>🔍 စုံစမ်းစစ်ဆေးချက် မှတ်ချက် (INV Note)</span>
                  </label>
                  {formData.passport && getLatestVerificationRemark(formData.passport, dossierHistory, records) !== '-' && (
                    <span className="text-[8px] font-black uppercase bg-amber-100 text-amber-900 border border-amber-300 px-1.5 py-0.5 rounded flex items-center gap-1 shrink-0 animate-pulse">
                      <AlertCircle size={10} /> Auto-filled
                    </span>
                  )}
                </div>
                <textarea
                  rows={4}
                  value={formData.remarks || ''}
                  onChange={(e) => setFormData(prev => ({ ...prev, remarks: e.target.value }))}
                  className="input-field bg-white border-amber-200 text-xs font-bold text-slate-800 focus:ring-amber-500 rounded-xl resize-none w-full p-2.5"
                  placeholder="INV tab မှ စုံစမ်းစစ်ဆေးချက် မှတ်ချက် သို့မဟုတ် ထပ်မံဖြည့်စွက်လိုသော အချက်အလက်များ..."
                />
              </div>
              <p className="text-[9px] text-slate-500 font-medium italic leading-tight">
                (INV tab ၏ စုံစမ်းစစ်ဆေးရေး မှတ်တမ်းမှ အလိုအလျောက် ရယူပြသပေးပါသည်)
              </p>
            </div>
          </div>

          <div className="lg:col-span-4 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8 border-t pt-8 mt-4">
            <datalist id="agentList">{(masterData || []).filter(m => m && m.type === 'Agent' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="contactList">{(masterData || []).filter(m => m && m.type === 'Contact' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="ffeOfficialList">{(masterData || []).filter(m => m && m.type === 'Official' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="ffeTitleList">{(masterData || []).filter(m => m && m.type === 'Title' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            <datalist id="permitDescriptionList">{(masterData || []).filter(m => m && m.type === 'PermitDescription' && m.name).map(m => <option key={m.id} value={m.name} />)}</datalist>
            
            <div>
              <label className="input-label text-emerald-800 font-black">Bring Up By (ခေါ်ဆောင်လာသူ)</label>
              <input 
                list="agentList" 
                value={formData.broughtBy || ''} 
                onChange={(e) => handleAgentInput(e.target.value)} 
                className="input-field bg-emerald-50/40 border-emerald-200 text-gray-900 font-bold" 
                placeholder="Agent or Guide Name" 
              />
            </div>
            <div>
              <label className="input-label text-emerald-800 font-black">Contact Details (ဆက်သွယ်ရန်)</label>
              <input list="contactList" value={formData.contactDetails || ''} onChange={(e) => setFormData(prev => ({ ...prev, contactDetails: e.target.value }))} className="input-field bg-emerald-50/40 border-emerald-200 text-gray-900 font-bold" placeholder="Phone or Messenger" />
            </div>
            <div>
              <label className="input-label text-purple-800 font-black">Official Name (စာရင်းသွင်းသူ)</label>
              <input 
                list="ffeOfficialList" 
                value={formData.officialName || ''} 
                onChange={(e) => !currentUser && handleOfficialInputMain(e.target.value)} 
                className={`input-field bg-purple-50/40 border-purple-200 text-gray-900 ${currentUser ? 'font-black' : 'font-bold'}`}
                placeholder="Officer Name" 
                readOnly={!!currentUser}
              />
            </div>
            <div>
              <label className="input-label text-purple-800 font-black">Official Title (ရာထူး)</label>
              <input 
                list="ffeTitleList" 
                value={formData.officialTitle || ''} 
                onChange={(e) => !currentUser && setFormData(prev => ({ ...prev, officialTitle: e.target.value }))} 
                className={`input-field bg-purple-50/40 border-purple-200 text-gray-900 ${currentUser ? 'font-black' : 'font-bold'}`}
                placeholder="Title" 
                readOnly={!!currentUser}
              />
            </div>
            
            {editTarget && (
              <>
                <div>
                  <label className="input-label text-red-600 font-black">System Mode (Log Type)</label>
                  <select 
                    value={formData.mode}
                    onChange={(e) => setFormData(prev => ({ ...prev, mode: e.target.value as Mode }))}
                    className="input-field border-red-100 bg-red-50/20"
                  >
                    <option value="IN">IN (Arrival)</option>
                    <option value="OUT">OUT (Departure)</option>
                  </select>
                </div>
                <div>
                  <label className="input-label text-red-600 font-black">Log Timestamp (Date/Time)</label>
                  <input 
                    type="text" 
                    value={formData.timestamp}
                    onChange={(e) => setFormData(prev => ({ ...prev, timestamp: e.target.value }))}
                    className="input-field border-red-100 bg-red-50/20"
                    placeholder="DD/MM/YYYY, HH:MM:SS"
                  />
                </div>
              </>
            )}
          </div>

          <div className="lg:col-span-3 flex justify-between items-center gap-4 mt-8 pt-6 border-t">
             <div className="flex items-center gap-2 text-[10px] uppercase font-bold text-gray-400">
               <RefreshCcw size={12} /> Auto-fill from history enabled
             </div>
             <div className="flex gap-4">
               <button type="button" onClick={resetForm} className="btn border border-gray-200 text-gray-600 hover:bg-gray-50 uppercase text-xs">Clear Form</button>
               <button 
                 type="submit" 
                 disabled={isViewer}
                 className={`btn px-10 text-white uppercase text-xs tracking-wider shadow-lg ${
                   isViewer 
                     ? 'bg-slate-400 cursor-not-allowed text-slate-200' 
                     : currentMode === 'IN' 
                     ? 'bg-emerald-600 hover:bg-emerald-700 shadow-emerald-200' 
                     : 'bg-orange-600 hover:bg-orange-700 shadow-orange-200'
                 }`}
               >
                 {isViewer ? '🔒 Viewer Mode (Read-Only)' : editTarget ? 'Update Record' : 'Confirm Entry'}
               </button>
             </div>
          </div>
        </form>
      </div>
    </motion.div>
  );

  const renderTempHistory = () => null;

  const renderTempHistory_OLD = () => {
    const toggleSelectAll = () => {
      setSelectedTempIds(selectedTempIds.length === tempRecords.length ? [] : tempRecords.map(r => r.id));
    };
    
    const filteredTempRecords = tempRecords.filter(r => 
      r.fullname.toLowerCase().includes(tehQuery.toLowerCase()) || 
      r.passport.toLowerCase().includes(tehQuery.toLowerCase())
    );

    const toggleSelect = (id: number) => {
      setSelectedTempIds(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
    };

    const deleteTempRecord = (id: number) => {
      setTempRecords(prev => prev.filter(r => r.id !== id));
      setConfirmTempDeleteId(null);
      showToast("TEMPORARY RECORD DELETED");
    };

    const confirmMoveToFH = (tr: ImmRecord) => {
      const exists = records.some(r => r.passport.toUpperCase() === tr.passport.toUpperCase());
      if (exists) {
        setRecords(prev => prev.map(r => r.passport.toUpperCase() === tr.passport.toUpperCase() ? { ...r, ...tr, id: r.id } : r));
      } else {
        setRecords(prev => [{ ...tr, id: Date.now() }, ...prev]);
      }
      setTempRecords(prev => prev.filter(r => r.id !== tr.id));
      setTehAnalysisTarget(null);
      showToast("PROMOTED TO FLIGHT HISTORY (FH)");
    };

    const editTempRecord = (r: ImmRecord) => {
      setPreFilledTempId(r.id);
      setActiveTab('formC');
    };

    const exportTempToExcel = () => {
      const data = tempRecords.map(r => ({
        "Timestamp": r.timestamp,
        "Status": r.mode,
        "Full Name": r.fullname,
        "Passport": r.passport,
        "Nationality": r.nationality,
        "Gender": r.gender,
        "Visa Type": r.visaType,
        "Address": r.address,
        "Vehicle Info": r.vehicleInfo,
        "Total Days": r.totalDays,
        "Stay From": r.stayFrom,
        "Stay To": r.stayTo,
        "Official": r.formC?.officialName || 'N/A',
        "Reporter": r.formC?.reporterName || 'N/A',
        "Phone": r.formC?.reporterPhone || 'N/A',
        "Remarks": r.remarks || ''
      }));
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "TempRecords");
      XLSX.writeFile(wb, `Temporary_Entries_${new Date().toISOString().split('T')[0]}.xlsx`);
    };

    const batchDelete = () => {
      if (selectedTempIds.length === 0) return;
      setTempRecords(prev => prev.filter(r => !selectedTempIds.includes(r.id)));
      setSelectedTempIds([]);
      showToast("TEMPORARY RECORDS REMOVED");
    };

    return (
      <div className="max-w-[1600px] mx-auto py-8 px-4 space-y-6">
        <div className="card shadow-2xl border-t-8 border-amber-500">
          <div className="flex flex-col md:flex-row justify-between items-center mb-8 gap-6 no-print">
             <div className="flex flex-col md:flex-row items-center gap-6 w-full">
                <div className="flex items-center gap-4">
                   <div className="bg-amber-100 p-3 rounded-full text-amber-700">
                     <Clock size={28} />
                   </div>
                   <div>
                     <h2 className="text-2xl font-black uppercase text-gray-800 tracking-tight">Temporary Entry History</h2>
                     <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Records created via Form C searching</p>
                   </div>
                </div>
                
                <div className="flex-1 relative group max-w-md w-full ml-auto">
                   <input 
                     type="text" 
                     placeholder="Search TEH records..." 
                     value={tehQuery}
                     onChange={(e) => setTehQuery(e.target.value)}
                     className="input-field pl-10 h-10 text-xs font-bold border-amber-200 focus:border-amber-500 transition-all text-gray-800 bg-amber-50/30" 
                   />
                   <Search size={18} className="absolute left-3 top-2.5 text-amber-400 group-focus-within:text-amber-600" />
                   {tehQuery && (
                     <button 
                       onClick={() => setTehQuery('')}
                       className="absolute right-3 top-2.5 text-amber-400 hover:text-red-500 transition-colors"
                     >
                       <X size={18} />
                     </button>
                   )}
                </div>
             </div>
          </div>

          <div className="flex flex-wrap justify-between items-center mb-6 gap-4 border-b border-amber-100 pb-4">
             <div className="flex gap-2">
                <button onClick={analyzeAllTEH} className="btn bg-indigo-600 text-white hover:bg-black text-[10px] uppercase font-black px-4 flex items-center gap-2"><RefreshCcw size={14} /> Analysis All</button>
                {selectedTempIds.length > 0 && (
                  <button onClick={batchDelete} className="btn bg-red-600 text-white hover:bg-black text-[10px] uppercase font-black px-4 flex items-center gap-2 animate-pulse"><Trash2 size={14} /> Delete Selected ({selectedTempIds.length})</button>
                )}
                <button onClick={exportTempToExcel} className="btn bg-emerald-600 text-white hover:bg-emerald-700 text-[10px] uppercase font-black px-4 flex items-center gap-2"><Download size={14} /> Export Excel</button>
             </div>
             <div className="text-[10px] font-black text-gray-400 uppercase tracking-widest bg-amber-50 px-3 py-1 rounded-full border border-amber-100">Showing {filteredTempRecords.length} of {tempRecords.length} records</div>
          </div>

          <div className="overflow-x-auto border rounded-2xl shadow-sm">
             <table className="w-full text-left min-w-[1200px]">
                <thead className="bg-gray-900 text-white uppercase text-[10px] tracking-widest font-black">
                   <tr>
                      <th className="p-4 text-center w-12 border-r border-white/10">
                        <input type="checkbox" checked={tempRecords.length > 0 && selectedTempIds.length === tempRecords.length} onChange={toggleSelectAll} className="rounded" />
                      </th>
                      <th className="p-4 border-r border-white/10">Timestamp</th>
                      <th className="p-4 border-r border-white/10">Identity</th>
                      <th className="p-4 border-r border-white/10">Stay Location</th>
                      <th className="p-4 border-r border-white/10">Authority / Reporter</th>
                      <th className="p-4 border-r border-white/10">Remarks</th>
                      <th className="p-4 text-center border-r border-white/10">Status</th>
                      <th className="p-4 text-center sticky right-0 bg-gray-900">Action</th>
                   </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 text-[11px] text-gray-700 font-bold">
                   {filteredTempRecords.map(r => (
                     <tr key={r.id} className="hover:bg-amber-50/30 transition-colors">
                        <td className="p-4 text-center border-r border-gray-100 uppercase">
                          <input type="checkbox" checked={selectedTempIds.includes(r.id)} onChange={() => toggleSelect(r.id)} className="rounded" />
                        </td>
                        <td className="p-4 border-r border-gray-100">
                           <div className="text-gray-400 mb-1 font-medium">{r.timestamp}</div>
                           <span className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full font-black text-[9px] tracking-tighter shadow-sm border border-amber-200">TEMPORARY ENTRY</span>
                        </td>
                        <td className="p-4 border-r border-gray-100">
                           <div className="font-black text-gray-900 text-sm mb-1">{r.fullname}</div>
                           <div className="font-mono text-gray-400 font-bold">{r.passport} | {r.nationality}</div>
                        </td>
                        <td className="p-4 border-r border-gray-100">
                           <div className="font-black text-amber-700 mb-1">{r.address}</div>
                           <div className="text-[10px] text-gray-400 font-bold">{r.vehicleInfo || '-'}</div>
                        </td>
                        <td className="p-4 border-r border-gray-100">
                           <div className="font-black text-gray-800 uppercase mb-1">{r.formC?.officialName}</div>
                           <div className="text-gray-400 uppercase text-[9px] font-bold">{r.formC?.reporterName} ({r.formC?.reporterPhone})</div>
                        </td>
                        <td className="p-4 border-r border-gray-100 italic text-gray-500 max-w-[200px] truncate">
                           {r.remarks || '-'}
                        </td>
                        <td className="p-4 text-center border-r border-gray-100">
                           <span className={`px-4 py-1.5 rounded-lg font-black text-[10px] border-2 ${r.mode === 'IN' ? 'bg-emerald-50 text-emerald-800 border-emerald-100' : 'bg-orange-50 text-orange-800 border-orange-100'}`}>{r.mode === 'IN' ? 'CHECKED IN' : 'CHECKED OUT'}</span>
                        </td>
                        <td className="p-2 text-center sticky right-0 bg-white shadow-[-4px_0_10px_-4px_rgba(0,0,0,0.05)] text-gray-700">
                           <div className="flex flex-col gap-1">
                              <button onClick={() => editTempRecord(r)} className="text-[9px] bg-blue-50 text-blue-700 px-2 py-1 rounded-md font-black hover:bg-blue-100 uppercase">Edit</button>
                              <button onClick={() => setTehAnalysisTarget(r)} className="text-[9px] bg-indigo-50 text-indigo-700 px-2 py-1 rounded-md font-black hover:bg-indigo-100 uppercase">Analysis</button>
                              {confirmTempDeleteId === r.id ? (
                                <div className="flex gap-1 animate-in zoom-in-95">
                                  <button onClick={() => deleteTempRecord(r.id)} className="text-[8px] bg-red-600 text-white px-2 py-1 rounded-md font-black uppercase">YES</button>
                                  <button onClick={() => setConfirmTempDeleteId(null)} className="text-[8px] bg-gray-200 text-gray-600 px-2 py-1 rounded-md font-black uppercase">NO</button>
                                </div>
                              ) : (
                                <button onClick={() => setConfirmTempDeleteId(r.id)} className="text-[9px] bg-red-50 text-red-500 px-2 py-1 rounded-md font-black hover:bg-red-100 uppercase">Delete</button>
                              )}
                           </div>
                        </td>
                     </tr>
                   ))}
                </tbody>
             </table>
             {filteredTempRecords.length === 0 && (
               <div className="p-12 text-center text-gray-300 font-black uppercase tracking-widest">No matching temporary records found</div>
             )}
          </div>
        </div>
      </div>
    );
  };

  const renderHistory = () => {
    const toggleSelectAll = () => {
      if (selectedIds.length === filteredRecords.length) {
        setSelectedIds([]);
      } else {
        setSelectedIds(filteredRecords.map(r => r.id));
      }
    };

    const toggleSelect = (id: number) => {
      setSelectedIds(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
    };

    const startEdit = (r: ImmRecord) => {
      if (isViewer) {
        showToast("🔒 Viewer level အကောင့်ဖြစ်သောကြောင့် ပြင်ဆင်ခွင့် ပိတ်ထားပါသည်");
        return;
      }
      setEditTarget(r);
      setFormData({
        ...r,
        dob: r.dob ? formatToDDMMYYYY(r.dob) : '',
        stayFrom: r.stayFrom ? formatToDDMMYYYY(r.stayFrom) : '',
        stayTo: r.stayTo ? formatToDDMMYYYY(r.stayTo) : '',
        stayDescription: r.stayDescription || '',
        totalDays: r.totalDays != null ? String(r.totalDays) : ''
      });
      setCurrentMode(r.mode);
      setActiveTab('entry');
    };

    const deleteRecord = (id: number) => {
      if (isViewer) {
        showToast("🔒 Viewer level အကောင့်ဖြစ်သောကြောင့် ဖျက်ပစ်ခွင့် ပိတ်ထားပါသည်");
        return;
      }
      const target = records.find(r => r.id === id);
      setRecords(prev => prev.filter(r => r.id !== id));
      setConfirmDeleteId(null);
      if (target) {
        logActivity({
          action: 'DELETE',
          module: target.logType === 'FCR' ? 'FCR' : 'FFE',
          officerName: cloudAuthUser?.username,
          officerRole: cloudAuthUser?.role || 'Editor',
          targetId: target.passport,
          details: `Deleted ${target.logType || 'FFE'} [${target.mode}] record #${target.id} (${target.passport}, ${target.fullname || 'N/A'})`
        });
      }
      showToast("RECORD DELETED");
    };

    const batchDelete = () => {
      if (isViewer) {
        showToast("🔒 Viewer level အကောင့်ဖြစ်သောကြောင့် ဖျက်ပစ်ခွင့် ပိတ်ထားပါသည်");
        return;
      }
      if (selectedIds.length === 0) return;
      const count = selectedIds.length;
      setRecords(prev => prev.filter(r => !selectedIds.includes(r.id)));
      setSelectedIds([]);
      logActivity({
        action: 'DELETE',
        module: 'FFE',
        officerName: cloudAuthUser?.username,
        officerRole: cloudAuthUser?.role || 'Editor',
        targetId: `${count} records`,
        details: `Batch deleted ${count} records from database`
      });
      showToast(`${count} RECORDS REMOVED`);
    };

    return (
      <div className="max-w-[1600px] mx-auto py-3 sm:py-6 px-2 sm:px-4 space-y-3 sm:space-y-4">
        {isViewer && (
          <div className="p-3 bg-amber-50 border-2 border-amber-300 text-amber-900 rounded-2xl flex items-center gap-2.5 font-bold text-xs shadow-xs no-print">
            <AlertCircle size={18} className="text-amber-600 shrink-0" />
            <div>
              <span className="font-black">🔒 Viewer Mode Active: </span>
              သင်သည် Viewer အဆင့်ဖြစ်သောကြောင့် အချက်အလက်များ ရှာဖွေ/ကြည့်ရှုခွင့်သာရှိပြီး အသစ်ထည့်ခြင်း၊ ပြင်ဆင်ခြင်းနှင့် ဖျက်ပစ်ခြင်းများအား ပိတ်ထားပါသည် (Read-Only)
            </div>
          </div>
        )}
        <div className="card p-3 sm:p-5">
            <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 w-full no-print border-b pb-3">
               <div className="flex items-center gap-2.5">
                  <div className="bg-[#2C6CB0]/10 p-2 rounded-xl text-[#2C6CB0] border border-[#2C6CB0]/20 shrink-0">
                    <History size={20} />
                  </div>
                  <div>
                    <h2 className="text-sm sm:text-base font-black uppercase text-slate-800 tracking-tight leading-tight">Main Flight Log History</h2>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5 mt-0.5">
                       <Activity size={10} className="text-emerald-500" /> Database: {records.length} total entries
                    </p>
                  </div>
               </div>
               
               <div className="relative group max-w-md w-full no-print">
                  <input 
                    type="text" 
                    placeholder="Search logs by name, passport, vehicle..." 
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-8 h-9 text-xs font-bold border border-slate-200 rounded-xl focus:border-[#2C6CB0] outline-none transition-all text-slate-800 bg-slate-50/60" 
                  />
                  <Search size={15} className="absolute left-3 top-2.5 text-slate-400 group-focus-within:text-[#2C6CB0] transition-colors" />
                  {searchQuery && (
                    <button 
                      onClick={() => setSearchQuery('')}
                      className="absolute right-2.5 top-2.5 text-slate-400 hover:text-red-500 transition-colors p-0.5"
                    >
                      <X size={15} />
                    </button>
                  )}
               </div>
            </div>

          <div className="flex flex-wrap items-center justify-between gap-2.5 my-3 no-print border-b pb-3 text-xs">
             <div className="flex items-center gap-2">
                <div className="flex items-center gap-1.5 bg-slate-100/80 p-1 rounded-xl border border-slate-200/80">
                  <div className="flex items-center gap-1 px-2 py-0.5 border-r border-slate-200">
                    <Calendar size={12} className="text-slate-500" />
                    <span className="text-[9px] font-black uppercase tracking-wider text-slate-500">Date Range</span>
                  </div>
                  <div className="flex items-center gap-1 px-1">
                    <input 
                      type="date" 
                      value={globalDateRange.from}
                      onChange={(e) => setGlobalDateRange(prev => ({ ...prev, from: e.target.value }))}
                      className="bg-white text-slate-900 border border-slate-200 text-[10px] font-bold px-1.5 py-0.5 rounded-md outline-none focus:ring-1 focus:ring-[#2C6CB0]" 
                    />
                    <span className="text-[9px] font-black opacity-40">TO</span>
                    <input 
                      type="date" 
                      value={globalDateRange.to}
                      onChange={(e) => setGlobalDateRange(prev => ({ ...prev, to: e.target.value }))}
                      className="bg-white text-slate-900 border border-slate-200 text-[10px] font-bold px-1.5 py-0.5 rounded-md outline-none focus:ring-1 focus:ring-[#2C6CB0]" 
                    />
                    {(globalDateRange.from || globalDateRange.to) && (
                      <button 
                        onClick={() => setGlobalDateRange({ from: '', to: '' })}
                        className="text-red-500 hover:bg-red-50 p-0.5 rounded transition-all ml-0.5"
                      >
                        <X size={12} />
                      </button>
                    )}
                  </div>
                </div>
             </div>
             
             <div className="flex items-center gap-2">
                {selectedIds.length > 0 && (
                  <button 
                    onClick={batchDelete}
                    className="btn bg-red-600 text-white hover:bg-black text-[10px] uppercase font-black px-3 py-1.5 rounded-xl animate-pulse flex items-center gap-1"
                  >
                    <Trash2 size={12} /> Delete ({selectedIds.length})
                  </button>
                )}
                <button 
                  onClick={exportAllToExcel} 
                  className="btn bg-emerald-600 text-white hover:bg-emerald-700 text-[10px] font-black uppercase px-3.5 py-1.5 rounded-xl shadow-xs flex items-center gap-1.5 transition-all"
                >
                  <Download size={12} /> Export Excel
                </button>
                <button 
                  onClick={() => setIsCounterCheckOpen(true)} 
                  className="btn bg-indigo-600 hover:bg-indigo-700 text-white text-[10px] font-black uppercase px-3.5 py-1.5 rounded-xl shadow-xs flex items-center gap-1.5 transition-all cursor-pointer relative font-pyidaungsu"
                  title="Passport နံပါတ်/အမည် စာရိုက်မှားယွင်းမှုများ စိစစ်ရန်"
                >
                  <CheckCircle2 size={13} className="text-indigo-200" />
                  <span>Counter Check</span>
                  {counterCheckSuspectCount > 0 && (
                    <span className="bg-rose-500 text-white text-[9px] font-extrabold px-1.5 py-0.2 rounded-full font-mono animate-pulse">
                      {counterCheckSuspectCount}
                    </span>
                  )}
                </button>
             </div>
          </div>

          {/* MOBILE VIEW (COMPACT DETAILED CARDS) */}
          <div className="block md:hidden space-y-2">
             {filteredRecords.map((r, i) => (
                <div 
                   key={r.id}
                   onClick={() => setMobileDetailRecord(r)}
                   className="bg-white border border-slate-200/90 rounded-xl p-2.5 shadow-2xs hover:shadow-sm transition-all cursor-pointer space-y-2"
                >
                   {/* Top Header Row */}
                   <div className="flex items-center justify-between gap-1.5 border-b border-slate-100 pb-1.5">
                      <div className="flex items-center gap-1.5">
                         <span className="font-black text-indigo-950 font-mono text-xs tracking-tight">{r.passport}</span>
                          {r.previousPassport && (
                            <span className="text-[8px] bg-purple-100 text-purple-900 border border-purple-300 px-1 py-0.2 rounded font-mono font-bold flex items-center gap-0.5" title={`Previous PP: ${r.previousPassport}`}>
                              <Link2 size={8} /> {r.previousPassport}
                            </span>
                          )}
                         <div className={`w-2 h-2 rounded-full ${r.gender === 'M' ? 'bg-blue-500' : 'bg-pink-500'}`} title={r.gender === 'M' ? 'Male' : 'Female'} />
                         <span className="text-[10px] font-black text-slate-500 bg-slate-100 px-1.5 py-0.2 rounded border border-slate-200">{r.nationality || 'NO NAT'}</span>
                      </div>
                      <div className="flex items-center gap-1 flex-wrap">
                         <span className={`px-2 py-0.2 rounded-md text-[9px] font-black tracking-wider uppercase ${r.mode === 'IN' ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-orange-100 text-orange-800 border border-orange-300'}`}>
                           {r.mode}
                         </span>
                         {r.syncStatus === 'upload_failed' ? (
                           <span className="px-1.5 py-0.2 rounded text-[8px] font-black uppercase tracking-wider bg-rose-100 text-rose-800 border border-rose-300 flex items-center gap-0.5 shadow-2xs" title="⚠️ Server သို့ မရောက်သေးပါ (Upload Failed)">
                             <AlertCircle size={9} className="text-rose-600 animate-pulse"/> Failed
                           </span>
                         ) : r.syncStatus === 'pending_sync' ? (
                           <span className="px-1.5 py-0.2 rounded text-[8px] font-black uppercase tracking-wider bg-amber-100 text-amber-900 border border-amber-300 flex items-center gap-0.5 shadow-2xs" title="⏳ Server သို့ မရောက်သေးပါ (ဖုန်းထဲတွင် သိမ်းထားသည်)">
                             <Clock size={9} className="text-amber-600"/> Pending Upload
                           </span>
                         ) : (
                           <span className="px-1.5 py-0.2 rounded text-[8px] font-black tracking-wider bg-emerald-100 text-emerald-900 border border-emerald-300 flex items-center gap-0.5 shadow-2xs" title={`✓ Server ပေါ်ရောက်ရှိပြီး (Cloud Synced)${r.createdDevice ? ` • by ${r.createdDevice}` : ''}${r.serverSyncedAt ? ` • ${new Date(r.serverSyncedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}` : ''}`}>
                             <CheckCircle2 size={9} className="text-emerald-600"/> Server Synced
                           </span>
                         )}
                      </div>
                   </div>

                   {/* Full Name & Visa Type */}
                   <div className="flex items-start justify-between gap-2">
                      <div>
                         <div className="font-extrabold text-slate-900 text-xs leading-tight">{r.fullname}</div>
                         <div className="text-[10px] font-bold text-indigo-700 uppercase mt-0.5 flex items-center gap-1">
                           <span>Visa: {r.visaType || '-'}</span>
                           {r.stayFrom && <span className="text-emerald-700">({r.totalDays || 1} Days)</span>}
                         </div>
                      </div>
                      <div className="text-right shrink-0">
                         <div className="text-[9px] font-semibold text-slate-400">{r.timestamp}</div>
                      </div>
                   </div>

                   {/* Compact Detailed Info Grid */}
                   <div className="grid grid-cols-2 gap-1 text-[10px] bg-slate-50/80 p-2 rounded-lg border border-slate-100">
                      <div className="truncate"><span className="text-slate-400 font-medium">Hotel/Address:</span> <span className="font-bold text-slate-800">{r.address || '-'}</span></div>
                      <div className="truncate"><span className="text-slate-400 font-medium">Vehicle:</span> <span className="font-bold text-purple-800">{r.vehicleInfo || '-'}</span></div>
                      <div className="truncate"><span className="text-slate-400 font-medium">Stay Period:</span> <span className="font-bold text-slate-700">{r.stayFrom ? `${r.stayFrom} to ${r.stayTo}` : '-'}</span></div>
                      <div className="truncate"><span className="text-slate-400 font-medium">Route:</span> <span className="font-bold text-slate-700">{r.arrivedFrom || '-'} → {r.departedTo || '-'}</span></div>
                   </div>

                   {/* Footer Link */}
                   <div className="flex items-center justify-between pt-0.5 text-[9px] font-bold text-slate-400">
                      <span>Agent/Carrier: {r.broughtBy || '-'}</span>
                      <div className="text-indigo-600 font-black flex items-center gap-0.5">
                         <span>Full Details</span>
                         <ChevronDown size={10} className="-rotate-90" />
                      </div>
                   </div>
                </div>
             ))}

             {filteredRecords.length === 0 && (
               <div className="p-6 text-center text-slate-400 font-medium bg-slate-50 rounded-xl border border-dashed border-slate-200 text-xs">
                 NO RECORDS FOUND WITHIN THIS VIEW.
               </div>
             )}
          </div>

          {/* DESKTOP & TABLET VIEW (TABLE) */}
          <div className="hidden md:block overflow-x-auto border rounded-2xl shadow-sm">
            <table className="w-full text-left min-w-[1500px]">
               <thead className="bg-gray-900 text-white uppercase text-[9px] tracking-wider font-black">
                 <tr>
                   <th className="p-2.5 text-center w-10 border-r border-white/10">
                     <input 
                       type="checkbox" 
                       className="rounded border-gray-400 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                       checked={filteredRecords.length > 0 && selectedIds.length === filteredRecords.length}
                       onChange={toggleSelectAll}
                     />
                   </th>
                   <th className="p-2.5 text-center w-10 border-r border-white/10">#</th>
                   <th className="p-2.5 border-r border-white/10">Time / Mode</th>
                   <th className="p-2.5 border-r border-white/10">Identity</th>
                   <th className="p-2.5 border-r border-white/10">Origin</th>
                   <th className="p-2.5 border-r border-white/10">Visa & Stay</th>
                   <th className="p-2.5 border-r border-white/10">Location</th>
                   <th className="p-2.5 border-r border-white/10">Logistics</th>
                   <th className="p-2.5 text-center sticky right-0 bg-gray-900 border-l border-white/10">Action</th>
                 </tr>
               </thead>
               <tbody className="divide-y divide-gray-100 text-xs">
                  {filteredRecords.map((r, i) => (
                    <tr 
                      key={r.id} 
                      className={`hover:bg-gray-50 transition-colors group ${selectedIds.includes(r.id) ? 'bg-indigo-50/50' : ''}`}
                    >
                      <td className="p-2 text-center border-r border-gray-100">
                        <input 
                          type="checkbox" 
                          className="rounded border-gray-400 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                          checked={selectedIds.includes(r.id)}
                          onChange={() => toggleSelect(r.id)}
                        />
                      </td>
                      <td className="p-2 text-center font-black text-gray-400 bg-gray-50/50 group-hover:text-indigo-600 border-r border-gray-100 text-[11px]">{i+1}</td>
                      <td className="p-2 border-r border-gray-100">
                        <div className="font-medium text-gray-500 mb-0.5 text-[11px]">{r.timestamp}</div>
                        <div className="flex items-center gap-1 flex-wrap">
                          <span className={`px-1.5 py-0.2 rounded text-[9px] font-black tracking-tighter ${r.mode === 'IN' ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-orange-100 text-orange-800 border border-orange-200'}`}>
                            {r.mode}
                          </span>
                          {r.syncStatus === 'upload_failed' ? (
                            <span className="px-1 py-0.2 rounded text-[8px] font-black uppercase tracking-wider bg-rose-100 text-rose-800 border border-rose-300 flex items-center gap-0.5" title="⚠️ Upload Failed: Saved in Local DB">
                              <AlertCircle size={8} className="text-rose-600 animate-pulse"/> Failed
                            </span>
                          ) : r.syncStatus === 'pending_sync' ? (
                            <span className="px-1 py-0.2 rounded text-[8px] font-black uppercase tracking-wider bg-amber-100 text-amber-800 border border-amber-300 flex items-center gap-0.5" title="⏱️ Pending Cloud Upload (ဆာဗာသို့ မရောက်သေးပါ)">
                              <Clock size={8} className="text-amber-600"/> Pending
                            </span>
                          ) : (
                            <span className="px-1.5 py-0.2 rounded text-[8px] font-black tracking-tighter bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-0.5" title={`✓ Server ပေါ်ရောက်ရှိပြီး (Cloud Synced)${r.createdDevice ? ` • by ${r.createdDevice}` : ''}${r.serverSyncedAt ? ` • ${new Date(r.serverSyncedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}` : ''}`}>
                              <CheckCircle2 size={8} className="text-emerald-600" /> Synced
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="p-2 border-r border-gray-100">
                        <div className="flex items-center gap-1.5 mb-0.5">
                          <span className="font-black text-indigo-950 font-mono text-xs">{r.passport}</span>
                          {r.previousPassport && (
                            <span className="text-[9px] bg-purple-100 text-purple-900 border border-purple-300 px-1.5 py-0.2 rounded font-black font-mono flex items-center gap-0.5" title={`Previous Passport: ${r.previousPassport}${r.dualPassportRemarks ? ` (${r.dualPassportRemarks})` : ''}`}>
                              <Link2 size={9} /> Prev: {r.previousPassport}
                            </span>
                          )}
                          <div className={`w-2 h-2 rounded-full ${r.gender === 'M' ? 'bg-blue-400' : 'bg-pink-400'}`} />
                        </div>
                        <div className="font-extrabold text-slate-800 line-clamp-1 text-xs">{r.fullname}</div>
                      </td>
                      <td className="p-2 border-r border-gray-100 font-extrabold text-indigo-700 uppercase text-xs">{r.nationality || '-'}</td>
                      <td className="p-2 border-r border-gray-100">
                        <div className="font-bold text-emerald-700 leading-tight mb-0.5 text-xs">{r.visaType || '-'}</div>
                        <div className="text-[10px] text-gray-500 font-medium">{r.stayFrom ? `${formatFriendlyDate(r.stayFrom)} to ${formatFriendlyDate(r.stayTo)}` : '-'}</div>
                      </td>
                      <td className="p-2 border-r border-gray-100 text-[11px]">
                        <div className="font-bold text-gray-700 line-clamp-1 mb-0.5"><MapPin size={10} className="inline mr-0.5 text-slate-400" /> {r.address || '-'}</div>
                        <div className="text-[10px] text-purple-700 font-extrabold"><Truck size={10} className="inline mr-0.5 text-purple-500" /> {r.vehicleInfo || '-'}</div>
                      </td>
                      <td className="p-2 border-r border-gray-100 text-[11px]">
                        <div className="text-gray-600 mb-0.5 font-medium">{r.arrivedFrom || '-'} → {r.departedTo || '-'}</div>
                        <div className="text-[10px] font-bold text-slate-400">Agent: {r.broughtBy || '-'}</div>
                      </td>
                      <td className="p-4 text-center sticky right-0 bg-white group-hover:bg-gray-50 border-l border-gray-100 shadow-[-4px_0_10px_-4px_rgba(0,0,0,0.05)]">
                        <div className="flex flex-col gap-1.5 min-w-[80px]">
                          {(r.syncStatus === 'upload_failed' || r.syncStatus === 'pending_sync') && (
                            <button 
                              onClick={(e) => { e.stopPropagation(); resetQuotaState(); handleManualSync(); }}
                              className="text-[9px] bg-rose-600 text-white px-2 py-1 rounded-lg font-black hover:bg-rose-700 transition-colors uppercase flex items-center justify-center gap-1 shadow-xs cursor-pointer"
                              title="Retry Upload to Cloud"
                            >
                              <RefreshCw size={9} /> Retry Upload
                            </button>
                          )}
                          {!isViewer && (
                            <button 
                              onClick={(e) => { e.stopPropagation(); startEdit(r); }}
                              className="text-[9px] bg-blue-50 text-blue-700 px-2 py-1.5 rounded-lg font-black hover:bg-blue-100 transition-colors uppercase cursor-pointer"
                            >
                              Edit
                            </button>
                          )}
                          <button 
                             onClick={(e) => { e.stopPropagation(); setAnalyzerTarget(r.passport); }}
                             className="text-[9px] bg-indigo-50 text-indigo-700 px-2 py-1.5 rounded-lg font-black hover:bg-indigo-100 transition-colors uppercase cursor-pointer"
                          >
                            Analyze
                          </button>
                          
                          {!isViewer && (
                            confirmDeleteId === r.id ? (
                              <div className="flex flex-col gap-1">
                                <button 
                                  onClick={(e) => { e.stopPropagation(); deleteRecord(r.id); }}
                                  className="text-[9px] bg-red-600 text-white px-2 py-1.5 rounded-lg font-black hover:bg-red-700 transition-all uppercase"
                                >
                                  CONFIRM
                                </button>
                                <button 
                                  onClick={(e) => { e.stopPropagation(); setConfirmDeleteId(null); }}
                                  className="text-[8px] text-gray-500 font-bold hover:underline"
                                >
                                  Cancel
                                </button>
                              </div>
                            ) : (
                              <button 
                                 onClick={(e) => { e.stopPropagation(); setConfirmDeleteId(r.id); }}
                                 className="text-[9px] bg-red-50 text-red-500 px-2 py-1.5 rounded-lg font-black hover:bg-red-100 transition-all uppercase cursor-pointer"
                              >
                                 Delete
                              </button>
                            )
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
               </tbody>
            </table>
            {filteredRecords.length === 0 && (
              <div className="p-12 text-center text-gray-400 font-medium">
                NO RECORDS FOUND WITHIN THIS VIEW.
              </div>
            )}
          </div>
        </div>

        {/* MOBILE DETAILS & ACTIONS MODAL */}
        <AnimatePresence>
          {mobileDetailRecord && (
            <div className="fixed inset-0 bg-black/60 backdrop-blur-xs z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 no-print">
               <motion.div 
                  initial={{ opacity: 0, y: 100 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: 100 }}
                  className="bg-white w-full max-w-lg rounded-t-3xl sm:rounded-3xl p-6 max-h-[90vh] overflow-y-auto space-y-6 shadow-2xl border border-slate-200"
               >
                  <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                     <div>
                        <div className="text-xs font-black uppercase text-indigo-600 tracking-wider">Flight Log Detail</div>
                        <h3 className="text-xl font-black text-slate-900 mt-0.5">{mobileDetailRecord.fullname}</h3>
                     </div>
                     <button 
                        onClick={() => setMobileDetailRecord(null)}
                        className="bg-slate-100 hover:bg-slate-200 text-slate-600 p-2 rounded-full transition-colors"
                     >
                        <X size={18} />
                     </button>
                  </div>

                  <div className="grid grid-cols-2 gap-4 text-xs">
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Passport Number</div>
                        <div className="text-sm font-black text-indigo-900 mt-1 flex items-center gap-1.5 flex-wrap">
                          <span>{mobileDetailRecord.passport}</span>
                          {mobileDetailRecord.previousPassport && (
                            <span className="text-[9px] bg-purple-100 text-purple-900 border border-purple-300 px-1.5 py-0.5 rounded font-bold font-mono">
                              🔗 Prev: {mobileDetailRecord.previousPassport}
                            </span>
                          )}
                        </div>
                        {mobileDetailRecord.dualPassportRemarks && (
                          <div className="text-[10px] text-purple-700 font-bold mt-1">
                            Note: {mobileDetailRecord.dualPassportRemarks}
                          </div>
                        )}
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Mode & Status</div>
                        <div className="flex items-center gap-1.5 mt-1">
                           <span className={`px-2 py-0.5 rounded-md text-[10px] font-black ${mobileDetailRecord.mode === 'IN' ? 'bg-emerald-100 text-emerald-800' : 'bg-orange-100 text-orange-800'}`}>
                             {mobileDetailRecord.mode}
                           </span>
                           {mobileDetailRecord.syncStatus === 'upload_failed' ? (
                             <span className="text-[9px] font-bold text-rose-600 bg-rose-50 px-1.5 py-0.5 rounded border border-rose-200">
                               ⚠️ Upload Failed
                             </span>
                           ) : mobileDetailRecord.syncStatus === 'pending_sync' ? (
                             <span className="text-[9px] font-bold text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded border border-amber-200">
                               ⏳ Pending Upload (Local Only)
                             </span>
                           ) : (
                             <span className="text-[9px] font-bold text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 flex items-center gap-1">
                               <CheckCircle2 size={10} /> Server Synced
                             </span>
                           )}
                        </div>
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Nationality</div>
                        <div className="font-black text-slate-800 mt-1">{mobileDetailRecord.nationality || '-'}</div>
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Gender</div>
                        <div className="font-black text-slate-800 mt-1">{mobileDetailRecord.gender === 'M' ? 'Male (M)' : 'Female (F)'}</div>
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100 col-span-2">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Visa Type & Stay Duration</div>
                        <div className="font-black text-emerald-700 mt-1">{mobileDetailRecord.visaType || '-'}</div>
                        <div className="text-[11px] text-slate-500 font-medium mt-0.5">
                           {mobileDetailRecord.stayFrom ? `${formatFriendlyDate(mobileDetailRecord.stayFrom)} to ${formatFriendlyDate(mobileDetailRecord.stayTo)}` : '-'}
                        </div>
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100 col-span-2">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Address / Location</div>
                        <div className="font-bold text-slate-800 mt-1">{mobileDetailRecord.address || '-'}</div>
                        {mobileDetailRecord.stayDescription && (
                           <div className="text-[11px] text-slate-500 font-medium mt-1 italic">{mobileDetailRecord.stayDescription}</div>
                        )}
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100 col-span-2">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Logistics & Route</div>
                        <div className="font-bold text-slate-800 mt-1">Vehicle: {mobileDetailRecord.vehicleInfo || '-'}</div>
                        <div className="text-[11px] text-slate-600 font-medium mt-0.5">
                           {mobileDetailRecord.arrivedFrom || '-'} → {mobileDetailRecord.departedTo || '-'}
                        </div>
                     </div>
                     <div className="bg-slate-50 p-3 rounded-2xl border border-slate-100 col-span-2">
                        <div className="text-[10px] font-bold text-slate-400 uppercase">Agent / Official</div>
                        <div className="font-bold text-slate-800 mt-1">Agent: {mobileDetailRecord.broughtBy || '-'}</div>
                        <div className="text-[11px] text-slate-500 font-medium mt-0.5">
                           Officer: {mobileDetailRecord.officerName || '-'} ({mobileDetailRecord.officerTitle || '-'})
                        </div>
                     </div>
                  </div>

                  {/* Actions Bar */}
                  <div className="pt-2 border-t border-slate-100 flex flex-col gap-2">
                     {(mobileDetailRecord.syncStatus === 'upload_failed' || mobileDetailRecord.syncStatus === 'pending_sync') && (
                       <button 
                         onClick={() => {
                           resetQuotaState();
                           handleManualSync();
                           setMobileDetailRecord(null);
                         }}
                         className={`w-full ${mobileDetailRecord.syncStatus === 'upload_failed' ? 'bg-rose-600 hover:bg-rose-700' : 'bg-indigo-600 hover:bg-indigo-700'} text-white font-black py-2.5 rounded-2xl uppercase text-xs flex items-center justify-center gap-2 shadow-sm cursor-pointer`}
                       >
                         <UploadCloud size={14} /> {mobileDetailRecord.syncStatus === 'upload_failed' ? 'Retry Upload to Cloud Server' : 'Upload to Cloud Server Now (ဆာဗာသို့ ပို့မည်)'}
                       </button>
                     )}
                     <div className={`grid ${isViewer ? 'grid-cols-1' : 'grid-cols-2'} gap-2`}>
                        {!isViewer && (
                          <button 
                             onClick={() => {
                                const target = mobileDetailRecord;
                                setMobileDetailRecord(null);
                                startEdit(target);
                             }}
                             className="bg-blue-600 hover:bg-blue-700 text-white font-black py-2.5 rounded-2xl uppercase text-xs flex items-center justify-center gap-1.5 cursor-pointer"
                          >
                             Edit Entry
                          </button>
                        )}
                        <button 
                           onClick={() => {
                              const p = mobileDetailRecord.passport;
                              setMobileDetailRecord(null);
                              setAnalyzerTarget(p);
                           }}
                           className="bg-indigo-600 hover:bg-indigo-700 text-white font-black py-2.5 rounded-2xl uppercase text-xs flex items-center justify-center gap-1.5 cursor-pointer"
                        >
                           Analyze
                        </button>
                     </div>

                     {!isViewer && (
                       confirmDeleteId === mobileDetailRecord.id ? (
                         <div className="flex gap-2 mt-1">
                           <button 
                             onClick={() => {
                               deleteRecord(mobileDetailRecord.id);
                               setMobileDetailRecord(null);
                             }}
                             className="flex-1 bg-red-600 hover:bg-red-700 text-white font-black py-2.5 rounded-2xl uppercase text-xs cursor-pointer"
                           >
                             CONFIRM DELETE
                           </button>
                           <button 
                             onClick={() => setConfirmDeleteId(null)}
                             className="bg-slate-200 text-slate-700 font-bold px-4 py-2.5 rounded-2xl text-xs cursor-pointer"
                           >
                             Cancel
                           </button>
                         </div>
                       ) : (
                         <button 
                            onClick={() => setConfirmDeleteId(mobileDetailRecord.id)}
                            className="w-full bg-red-50 hover:bg-red-100 text-red-600 font-black py-2.5 rounded-2xl uppercase text-xs cursor-pointer"
                         >
                            Delete Record
                         </button>
                       )
                     )}
                  </div>
               </motion.div>
            </div>
          )}
        </AnimatePresence>
      </div>
    );
  };

  const renderMovementLog = () => {
    let list = (Object.values(movementMap) as MovementData[]).sort((a, b) => b.latestTime - a.latestTime);
    const totalTrackedAll = list.length;
    const stillInCount = list.filter(m => !m.out).length;
    const leftCount = list.filter(m => !!m.out).length;
    
    if (movementFilter === 'IN') {
      list = list.filter(m => !m.out);
    } else if (movementFilter === 'OUT') {
      list = list.filter(m => !!m.out);
    }

    if (movementSearchQuery.trim()) {
      const q = movementSearchQuery.toLowerCase().trim();
      list = list.filter(m => 
        m.n.toLowerCase().includes(q) || 
        m.p.toLowerCase().includes(q) || 
        m.nat.toLowerCase().includes(q) || 
        (m.agent && m.agent.toLowerCase().includes(q)) ||
        (m.loc && m.loc.toLowerCase().includes(q)) ||
        (m.visa && m.visa.toLowerCase().includes(q)) ||
        (m.contact && m.contact.toLowerCase().includes(q))
      );
    }

    return (
      <div className="max-w-[1600px] mx-auto py-3 sm:py-6 px-2 sm:px-4 space-y-3 sm:space-y-4">
        <div className="card p-3 sm:p-5">
          <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 w-full no-print border-b pb-3">
            <div className="flex items-center gap-2.5">
              <div className="bg-[#2C6CB0]/10 p-2 rounded-xl text-[#2C6CB0] border border-[#2C6CB0]/20 shrink-0">
                <Activity size={20} />
              </div>
              <div>
                <h2 className="text-sm sm:text-base font-black uppercase text-slate-800 tracking-tight leading-tight">
                  Tracking Movement Status (M)
                </h2>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5 mt-0.5">
                  <Layers size={10} className="text-indigo-500" /> Database: {list.length} tracked records
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <div className="flex bg-slate-100/90 p-1 rounded-xl border border-slate-200 text-xs">
                <button 
                  onClick={() => setMovementFilter('ALL')} 
                  className={`px-3 py-1 rounded-lg text-[10px] font-black uppercase transition-all ${movementFilter === 'ALL' ? 'bg-white shadow-xs text-indigo-800 font-black' : 'text-slate-500 hover:text-slate-800'}`}
                >
                  ALL ({totalTrackedAll})
                </button>
                <button 
                  onClick={() => setMovementFilter('IN')} 
                  className={`px-3 py-1 rounded-lg text-[10px] font-black uppercase transition-all ${movementFilter === 'IN' ? 'bg-white shadow-xs text-emerald-800 font-black' : 'text-slate-500 hover:text-slate-800'}`}
                >
                  STILL IN ({stillInCount})
                </button>
                <button 
                  onClick={() => setMovementFilter('OUT')} 
                  className={`px-3 py-1 rounded-lg text-[10px] font-black uppercase transition-all ${movementFilter === 'OUT' ? 'bg-white shadow-xs text-orange-800 font-black' : 'text-slate-500 hover:text-slate-800'}`}
                >
                  LEFT ({leftCount})
                </button>
              </div>

              <button 
                onClick={exportMovementToExcel} 
                className="btn bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase px-3.5 py-1.5 rounded-xl shadow-xs flex items-center gap-1.5 transition-all"
                title="Export Movement Tracker Excel"
              >
                <Download size={12} /> Export Excel
              </button>
            </div>
          </div>

          <div className="relative group max-w-md w-full my-3 no-print">
            <input 
              type="text" 
              placeholder="Search by name, passport, nationality, location, agent..." 
              value={movementSearchQuery} 
              onChange={(e) => setMovementSearchQuery(e.target.value)}
              className="w-full pl-9 pr-8 h-9 text-xs font-bold border border-slate-200 rounded-xl focus:border-[#2C6CB0] outline-none transition-all text-slate-800 bg-slate-50/60" 
            />
            <Search size={15} className="absolute left-3 top-2.5 text-slate-400 group-focus-within:text-[#2C6CB0] transition-colors" />
            {movementSearchQuery && (
              <button 
                onClick={() => setMovementSearchQuery('')}
                className="absolute right-2.5 top-2.5 text-slate-400 hover:text-red-500 transition-colors p-0.5"
              >
                <X size={15} />
              </button>
            )}
          </div>

          {/* MOBILE VIEW (COMPACT DETAILED CARDS) */}
          <div className="block md:hidden space-y-2">
            {list.map((m) => {
              const nowMs = Date.now();
              const endTime = m.outTime ? m.outTime : nowMs;
              const stayedDays = m.inTime ? Math.max(1, Math.ceil((endTime - m.inTime) / (1000 * 60 * 60 * 24))) : 0;
              return (
                <div 
                  key={m.p}
                  className="bg-white border border-slate-200/90 rounded-xl p-2.5 shadow-2xs hover:shadow-sm transition-all cursor-pointer space-y-2"
                  onClick={() => {
                    setSearchQuery(m.p);
                    setActiveTab('data');
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                  }}
                >
                  <div className="flex items-center justify-between gap-1.5 border-b border-slate-100 pb-1.5">
                    <div className="flex items-center gap-1.5">
                      <span className="font-black text-indigo-950 font-mono text-xs tracking-tight">{m.p}</span>
                      <div className={`w-2 h-2 rounded-full ${m.gender === 'M' ? 'bg-blue-500' : 'bg-pink-500'}`} />
                      <span className="text-[10px] font-black text-slate-500 bg-slate-100 px-1.5 py-0.2 rounded border border-slate-200">{m.nat || 'NO NAT'}</span>
                    </div>
                    <span className={`px-2 py-0.2 rounded-md text-[9px] font-black tracking-wider uppercase ${!m.out ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-slate-100 text-slate-600 border border-slate-300'}`}>
                      {!m.out ? 'STILL IN' : 'LEFT'}
                    </span>
                  </div>

                  <div className="flex items-start justify-between gap-2">
                    <div>
                      <div className="font-extrabold text-slate-900 text-xs leading-tight">{m.n}</div>
                      <div className="text-[10px] font-bold text-indigo-700 uppercase mt-0.5 flex items-center gap-1">
                        <span>Visa: {m.visa || '-'}</span>
                        <span className="text-emerald-700 font-extrabold">({stayedDays} Stayed Days)</span>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-1 text-[10px] bg-slate-50/80 p-2 rounded-lg border border-slate-100">
                    <div className="truncate"><span className="text-slate-400 font-medium">Location:</span> <span className="font-bold text-slate-800">{m.loc || '-'}</span></div>
                    <div className="truncate"><span className="text-slate-400 font-medium">Vehicle:</span> <span className="font-bold text-purple-800">{m.vInfo || '-'}</span></div>
                    <div className="truncate"><span className="text-slate-400 font-medium">Arrival:</span> <span className="font-bold text-blue-700">{m.in || '-'}</span></div>
                    <div className="truncate"><span className="text-slate-400 font-medium">Departure:</span> <span className="font-bold text-orange-700">{m.out || '-'}</span></div>
                  </div>

                  <div className="flex items-center justify-between pt-0.5 text-[9px] font-bold text-slate-400">
                    <span>Agent: {m.agent || m.repName || '-'} {m.contact ? `(${m.contact})` : ''}</span>
                    <span className="text-indigo-600 font-black">View History</span>
                  </div>
                </div>
              );
            })}
            {list.length === 0 && (
              <div className="p-6 text-center text-slate-400 font-medium bg-slate-50 rounded-xl border border-dashed border-slate-200 text-xs">
                NO MOVEMENT RECORDS FOUND.
              </div>
            )}
          </div>

          {/* DESKTOP VIEW (COMPACT TABLE) */}
          <div className="hidden md:block overflow-x-auto border rounded-2xl shadow-xs">
            <table className="w-full text-left min-w-[1600px]">
              <thead className="bg-gray-900 text-white uppercase text-[9px] tracking-wider font-black">
                <tr>
                  <th className="p-2.5 text-center w-10 border-r border-white/10">#</th>
                  <th className="p-2.5 border-r border-white/10">Full Identity / Passport</th>
                  <th className="p-2.5 border-r border-white/10">Nationality</th>
                  <th className="p-2.5 border-r border-white/10">Stay Location & Visa</th>
                  <th className="p-2.5 border-r border-white/10 text-center">Stayed Days / Remain</th>
                  <th className="p-2.5 border-r border-white/10">Arrival Date (IN)</th>
                  <th className="p-2.5 border-r border-white/10">Departure Date (OUT)</th>
                  <th className="p-2.5 border-r border-white/10">Agent / Carrier Details</th>
                  <th className="p-2.5 border-r border-white/10">Authority / Officer</th>
                  <th className="p-2.5 text-center sticky right-0 bg-gray-900 border-l border-white/10">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 text-xs">
                {list.map((m, i) => {
                  const nowMs = Date.now();
                  const endTime = m.outTime ? m.outTime : nowMs;
                  const stayedDays = m.inTime ? Math.max(1, Math.ceil((endTime - m.inTime) / (1000 * 60 * 60 * 24))) : 0;

                  return (
                    <tr key={m.p} className="hover:bg-gray-50 transition-colors group">
                      <td className="p-2 text-center font-black text-gray-400 bg-gray-50/50 group-hover:text-indigo-600 border-r border-gray-100 text-[11px]">{i+1}</td>
                      <td className="p-2 border-r border-gray-100">
                        <div className="flex items-center gap-1.5 mb-0.5">
                          <span className="font-black text-indigo-950 font-mono text-xs">{m.p}</span>
                          <div className={`w-2 h-2 rounded-full ${m.gender === 'M' ? 'bg-blue-400' : 'bg-pink-400'}`} />
                        </div>
                        <button 
                          onClick={() => {
                            setSearchQuery(m.p);
                            setActiveTab('data');
                            window.scrollTo({ top: 0, behavior: 'smooth' });
                          }}
                          className="font-extrabold text-slate-800 hover:text-indigo-600 hover:underline text-xs text-left uppercase leading-tight line-clamp-1"
                        >
                          {m.n}
                        </button>
                      </td>
                      <td className="p-2 border-r border-gray-100 font-extrabold text-indigo-700 uppercase text-xs">{m.nat || '-'}</td>
                      <td className="p-2 border-r border-gray-100">
                        <div className="font-bold text-emerald-800 line-clamp-1 text-xs mb-0.5">{m.loc || '-'}</div>
                        <div className="text-[10px] text-slate-500 font-medium">{m.visa || '-'} | {m.start ? `${m.start} to ${m.end}` : '-'}</div>
                      </td>
                      <td className="p-2 border-r border-gray-100 text-center">
                        <div className="font-black text-slate-800 text-xs">{stayedDays} Days</div>
                        {!m.out && (
                          <div className={`font-black text-[9px] mt-0.5 ${parseInt(getLiveRemainingDays(m.end)) < 7 ? 'text-red-500 animate-pulse' : 'text-emerald-600'}`}>
                            {getLiveRemainingDays(m.end)}
                          </div>
                        )}
                      </td>
                      <td className="p-2 border-r border-gray-100 text-blue-700 font-extrabold text-xs">{m.in || '-'}</td>
                      <td className="p-2 border-r border-gray-100 text-orange-700 font-extrabold text-xs">{m.out || '-'}</td>
                      <td className="p-2 border-r border-gray-100 text-[11px]">
                        <div className="font-extrabold text-slate-800 truncate">{m.agent || m.repName || '-'}</div>
                        <div className="text-[10px] text-slate-500 font-medium">{m.contact || m.repPhone || '-'}</div>
                      </td>
                      <td className="p-2 border-r border-gray-100 text-[11px]">
                        <div className="font-bold text-indigo-900 truncate">{m.offName || '-'}</div>
                        <div className="text-[10px] text-slate-400 font-semibold uppercase">{m.offTitle || '-'}</div>
                      </td>
                      <td className="p-2 text-center sticky right-0 bg-white group-hover:bg-gray-50 border-l border-gray-100 shadow-[-4px_0_10px_-4px_rgba(0,0,0,0.05)]">
                        <span className={`px-2 py-0.5 rounded-md font-black text-[9px] uppercase tracking-wider ${!m.out ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-slate-100 text-slate-600 border border-slate-300'}`}>
                          {!m.out ? 'STILL IN' : 'LEFT'}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {list.length === 0 && (
              <div className="p-8 text-center text-slate-400 font-medium text-xs">
                NO MOVEMENT RECORDS FOUND WITHIN THIS FILTER.
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  // --- ACTIONS ---

  const startEdit = (r: ImmRecord) => {
    setEditTarget(r);
    setFormData({ ...r });
    setCurrentMode(r.mode);
    setActiveTab('entry');
  };

  const deleteRecord = (id: number) => {
    const target = records.find(r => r.id === id);
    setRecords(prev => prev.filter(r => r.id !== id));
    setTempRecords(prev => prev.filter(r => r.id !== id));
    setConfirmDeleteId(null);
    if (target) {
      logActivity({
        action: 'DELETE',
        module: target.logType === 'FCR' ? 'FCR' : 'FFE',
        officerName: cloudAuthUser?.username,
        officerRole: cloudAuthUser?.role || 'Editor',
        targetId: target.passport,
        details: `Deleted ${target.logType || 'FFE'} [${target.mode}] record #${target.id} (${target.passport}, ${target.fullname || 'N/A'})`
      });
    }
    showToast("RECORD REMOVED SUCCESSFULLY");
  };

  const exportAllToExcel = () => {
    const data = filteredRecords.map((r, i) => ({
      "No": i + 1, "Log Time": r.timestamp, "Mode": r.mode, "Passport": r.passport, "Name": r.fullname, "Gender": r.gender, "Nationality": r.nationality, "Visa": r.visaType, "Address": r.address, "Stay From": r.stayFrom, "Stay To": r.stayTo, "Vehicle": r.vehicleInfo, "Agent": r.broughtBy
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "History Logs");
    XLSX.writeFile(wb, `IMM_HistoryLog_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportMovementToExcel = () => {
    const data = (Object.values(movementMap) as MovementData[]).map((m, i) => {
      const nowMs = Date.now();
      const endTime = m.outTime ? m.outTime : nowMs;
      const stayedDays = m.inTime ? Math.max(1, Math.ceil((endTime - m.inTime) / (1000 * 60 * 60 * 24))) : 0;

      return {
        "No": i + 1,
        "Passport": m.p,
        "Full Name": m.n,
        "Gender": m.gender,
        "Nationality": m.nat,
        "Arrival Date (IN)": m.in,
        "Departure Date (OUT)": m.out || 'STILL IN',
        "Stay Location (Hotel/Address)": m.loc || '-',
        "Total Stayed Days (From Last Arrival)": m.inTime ? `${stayedDays} Days${!m.out ? ' (Current)' : ''}` : '-',
        "Agent Name (Brought By)": m.agent || m.repName || '-',
        "Stay Permissions (Visa Type)": m.visa || '-',
        "Agent Name / Contact Info": m.contact || m.repPhone || '-',
        "Vehicle Info": m.vInfo || '-',
        "Tracking Status": !m.out ? 'STILL IN' : 'LEFT'
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Movement Tracker");
    XLSX.writeFile(wb, `IMM_MovementTracker_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Excel Movement Tracker Exported!");
  };

  const backupCombinedJSON = () => {
    const backupObj = { 
      records, 
      tempRecords, 
      masterData, 
      vehicleSummaries,
      dailyPdfs,
      checkingHistory,
      dossierHistory,
      watchList,
      settings: {
        language
      }
    };
    const blob = new Blob([JSON.stringify(backupObj, null, 2)], { type: 'application/json' });
    const now = new Date();
    const dateStr = now.toISOString().split('T')[0];
    const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-');
    const a = document.createElement('a'); 
    a.href = URL.createObjectURL(blob); 
    a.download = `IMM_COMBINED_BACKUP_${dateStr}_T${timeStr}.json`; 
    a.click();
    logActivity({
      action: 'BACKUP',
      module: 'SYSTEM',
      officerName: cloudAuthUser?.username,
      officerRole: cloudAuthUser?.role || 'Editor',
      details: `Generated Full System JSON Backup (${records.length} records, ${masterData.length} master items, ${watchList.length} watchlist entries)`
    });
    showToast("COMBINED FULL SYSTEM BACKUP CREATED");
  };

  const restoreCombinedJSON = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const p = JSON.parse(event.target?.result as string);
        if (p && typeof p === 'object') {
          if (p.records && Array.isArray(p.records)) {
            const restoredRecords = p.records.map((r: any) => ({
              ...r,
              syncStatus: 'pending_sync',
              updatedAt: r.updatedAt || new Date().toISOString()
            }));
            setRecords(restoredRecords);
            await miniDB.set('records', restoredRecords);
          }
          if (p.tempRecords && Array.isArray(p.tempRecords)) {
            const restoredTemp = p.tempRecords.map((r: any) => ({
              ...r,
              syncStatus: 'pending_sync',
              updatedAt: r.updatedAt || new Date().toISOString()
            }));
            setTempRecords(restoredTemp);
            await miniDB.set('tempRecords', restoredTemp);
          }
          if (p.masterData && Array.isArray(p.masterData)) {
            setMasterData(p.masterData);
            await miniDB.set('masterData', p.masterData);
          }
          if (p.vehicleSummaries && Array.isArray(p.vehicleSummaries)) {
            setVehicleSummaries(p.vehicleSummaries);
            await miniDB.set('vehicleSummaries', p.vehicleSummaries);
          }
          if (p.dailyPdfs && typeof p.dailyPdfs === 'object') {
            setDailyPdfs(p.dailyPdfs);
            await miniDB.set('dailyPdfs', p.dailyPdfs);
          }
          if (p.checkingHistory && Array.isArray(p.checkingHistory)) {
            const ch = p.checkingHistory.map((h: any) => ({
              ...h,
              status: h.status === 'STILL PERMITTED' ? 'STAY PERMITTED' : (h.status === 'STILL NOT PERMITTED YET' ? 'STAY NOT PERMITTED' : h.status)
            }));
            setCheckingHistory(ch);
            await miniDB.set('checkingHistory', ch);
          }
          if (p.dossierHistory && Array.isArray(p.dossierHistory)) {
            setDossierHistory(p.dossierHistory);
            await miniDB.set('dossierHistory', p.dossierHistory);
          }
          if (p.watchList && Array.isArray(p.watchList)) {
            setWatchList(p.watchList);
            await miniDB.set('watchList', p.watchList);
            try {
              localStorage.setItem('imm_watchlist_records_v1', JSON.stringify(p.watchList));
            } catch (e) {}
          }
          if (p.settings && typeof p.settings === 'object') {
            if (p.settings.language) setLanguage(p.settings.language);
          }
          logActivity({
            action: 'RESTORE',
            module: 'SYSTEM',
            officerName: cloudAuthUser?.username,
            officerRole: cloudAuthUser?.role || 'Editor',
            details: `Restored Full System from JSON Backup file`
          });
          showToast("COMBINED SYSTEM RESTORED SUCCESSFULLY!");
        } else {
          showToast("INVALID COMBINED BACKUP STRUCTURE");
        }
      } catch (err) { 
        alert("Invalid Combined Backup File Structure"); 
      }
    };
    reader.readAsText(file);
    e.target.value = ''; // Reset input to allow re-uploading the same file
  };

  return (
    <>
    {/* Global Autocomplete Datalists specifically for the Checking tab */}
    <datalist id="checkingStayList">
      {masterData.filter(m => m.type === 'Stay').map(m => <option key={m.id} value={m.name} />)}
    </datalist>
    <datalist id="checkingStayDescList">
      {Array.from(new Set(
        (masterData || []).filter(m => m.type === 'Stay' && m.linkedValue).map(m => m.linkedValue as string)
          .concat(records.filter(r => r.stayDescription).map(r => r.stayDescription as string))
      )).map((desc, idx) => (
        <option key={desc + '_' + idx} value={desc} />
      ))}
    </datalist>
    <datalist id="checkingPermitDescriptionList">
      {masterData.filter(m => m.type === 'PermitDescription').map(m => <option key={m.id} value={m.name} />)}
    </datalist>
    <datalist id="checkingOfficialList">
      {masterData.filter(m => m.type === 'Official').map(m => <option key={m.id} value={m.name} />)}
    </datalist>
    <datalist id="checkingTitleList">
      {masterData.filter(m => m.type === 'Title').map(m => <option key={m.id} value={m.name} />)}
    </datalist>

    <div className="min-h-screen bg-slate-50 transition-colors duration-300 print:hidden">
      <Toast message={toast.message} visible={toast.visible} />
      {renderNav()}

      {/* EXPLICIT WRITE / UPLOAD ERROR ALARM BANNER */}
      <AnimatePresence>
        {uploadErrorAlarm?.show && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="bg-rose-950 text-white border-b-2 border-rose-500 p-3.5 shadow-2xl sticky top-[57px] z-[300] flex flex-col md:flex-row items-center justify-between gap-3 no-print"
          >
            <div className="flex items-center gap-3">
              <span className="p-2 bg-rose-800/80 text-rose-100 rounded-2xl animate-pulse">
                <AlertTriangle size={22} className="text-rose-200" />
              </span>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h4 className="font-black uppercase tracking-wider text-[11px] text-rose-200">
                    ⚠️ Cloud Write Failure Alarm
                  </h4>
                  {uploadErrorAlarm.collectionName && (
                    <span className="text-[9px] font-mono bg-rose-900 text-rose-300 px-2 py-0.5 rounded font-bold uppercase">
                      Target: {uploadErrorAlarm.collectionName}
                    </span>
                  )}
                  <span className="text-[9px] font-mono bg-rose-900 text-rose-300 px-2 py-0.5 rounded font-bold">
                    {uploadErrorAlarm.timestamp}
                  </span>
                </div>
                <p className="text-[11px] font-bold text-rose-100 mt-0.5 leading-snug">
                  ⚠️ Upload Failed: Your edits could not be uploaded to the server [Reason: {uploadErrorAlarm.reason}]. Changes are safely stored in Local DB on this device.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 w-full md:w-auto">
              <button
                onClick={() => {
                  resetQuotaState();
                  handleManualSync();
                }}
                className="flex-1 md:flex-none px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all flex items-center justify-center gap-1.5 border border-rose-400/40 cursor-pointer"
              >
                <RefreshCw size={13} className="animate-spin-slow" />
                Retry Upload
              </button>
              <button
                onClick={() => setUploadErrorAlarm(null)}
                className="p-2 text-rose-300 hover:text-white bg-rose-900/60 hover:bg-rose-900 rounded-xl transition-all cursor-pointer"
                title="Dismiss Alarm"
              >
                <X size={16} />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {!currentUser && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
             <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-slate-900/90 backdrop-blur-xl" />
             <motion.div 
               initial={{ scale: 0.9, opacity: 0, y: 40 }} 
               animate={{ scale: 1, opacity: 1, y: 0 }} 
               className="relative bg-white w-full max-w-md rounded-[3rem] overflow-hidden shadow-2xl border-t-8 border-[#2C6CB0]"
             >
                <div className="p-12 space-y-8 text-center">
                   <div className="w-24 h-24 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto">
                      <Users size={48} />
                   </div>
                   <div className="space-y-2">
                      <h2 className="text-3xl font-black text-slate-900 tracking-tight">Welcome, Officer</h2>
                      <p className="text-slate-500 font-medium tracking-tight">Please enter your official name to continue</p>
                   </div>
                   <div className="space-y-4">
                      <input 
                        type="text" 
                        value={loginName}
                        onChange={(e) => setLoginName(e.target.value)}
                        onKeyDown={(e) => e.key === 'Enter' && handleLogin()}
                        className="w-full py-5 px-8 rounded-2xl bg-slate-50 border-2 border-slate-100 focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all outline-none text-center font-black text-xl" 
                        placeholder="OFFICER NAME"
                      />
                      <button 
                        onClick={handleLogin}
                        className="w-full py-5 px-8 rounded-2xl bg-[#2C6CB0] text-white font-black uppercase tracking-widest hover:bg-black transition-all shadow-lg active:scale-95 flex items-center justify-center gap-3"
                      >
                        <LogIn size={20} />
                        Access System Control
                      </button>
                      </div>
                   <p className="text-[10px] text-slate-300 font-black uppercase tracking-widest pt-4 leading-relaxed">2ND DIVISION LOGISTICS UNIT • IMMIGRATION CONTROL MYEIK</p>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 30-Minute MD Tab Backup Reminder Modal */}
      <AnimatePresence>
        {showBackupReminderModal && (
          <div className="fixed inset-0 z-[250] flex items-center justify-center p-4 no-print">
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setShowBackupReminderModal(false)} className="absolute inset-0 bg-slate-900/80 backdrop-blur-md" />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 30 }} 
              animate={{ scale: 1, opacity: 1, y: 0 }} 
              exit={{ scale: 0.9, opacity: 0, y: 30 }}
              className="relative bg-white w-full max-w-lg rounded-[2.5rem] overflow-hidden shadow-2xl border-t-8 border-indigo-600 p-8 space-y-6 text-center"
            >
              <div className="w-16 h-16 bg-amber-100 text-amber-700 rounded-full flex items-center justify-center mx-auto animate-bounce">
                <HardDrive size={32} />
              </div>
              <div className="space-y-3">
                <h3 className="text-xl font-black text-slate-900 uppercase tracking-tight flex items-center justify-center gap-2">
                  <span>MD Tab</span> • Backup File Reminder
                </h3>
                <p className="text-base font-bold text-amber-950 bg-amber-50 p-4.5 rounded-2xl border border-amber-200 leading-relaxed font-pyidaungsu shadow-inner">
                  Drive ရှိ DATAIMMI တွင် ကျေးဇူးပြုပြီး backup file ကို save ပေးပါ
                </p>
                <p className="text-[11px] text-slate-400 font-bold uppercase tracking-wider">
                  (Every 30 Minutes Backup Reminder Alert)
                </p>
              </div>
              <div className="flex flex-col sm:flex-row gap-3 pt-2">
                <button 
                  onClick={() => {
                    backupCombinedJSON();
                    setShowBackupReminderModal(false);
                  }}
                  className="flex-1 py-4 px-4 rounded-2xl bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs uppercase flex items-center justify-center gap-2 shadow-lg shadow-indigo-100 transition-all active:scale-95"
                >
                  <Download size={16} /> Save Backup Now
                </button>
                <button 
                  onClick={() => {
                    setActiveTab('master');
                    setShowBackupReminderModal(false);
                  }}
                  className="flex-1 py-4 px-4 rounded-2xl bg-slate-100 hover:bg-slate-200 text-slate-800 font-black text-xs uppercase flex items-center justify-center gap-2 transition-all active:scale-95"
                >
                  <Database size={16} /> Go to MD Tab
                </button>
              </div>
              <button 
                onClick={() => setShowBackupReminderModal(false)}
                className="text-xs text-slate-400 hover:text-slate-600 font-bold uppercase tracking-wider block mx-auto underline pt-1 transition-colors"
              >
                ခေတ္တပိတ်မည် (Dismiss)
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {!currentUser && (
        <div className="fixed bottom-4 right-4 z-[300] flex flex-col items-end gap-2 no-print">
          <label className="p-2.5 rounded-full bg-slate-800/10 hover:bg-slate-800/20 text-slate-500 hover:text-slate-800 transition-all cursor-pointer flex items-center gap-2 group shadow-lg backdrop-blur-sm" title="Upload Backup">
             <div className="overflow-hidden w-0 group-hover:w-auto transition-all text-[8px] font-black uppercase tracking-tighter whitespace-nowrap">File Restore</div>
             <Upload size={12} strokeWidth={4} />
             <input type="file" className="hidden" accept=".json" onChange={restoreCombinedJSON} />
          </label>
        </div>
      )}

      <main>
        {activeTab === 'entry' && renderEntryForm()}
        {activeTab === 'data' && renderHistory()}
        {activeTab === 'movement' && renderMovementLog()}
        {activeTab === 'stillIn' && (
          <StillInAnalytics 
            movementMap={movementMap} 
            records={records} 
            setRecords={setRecords}
            setTempRecords={setTempRecords}
            showToast={showToast}
            setActivePrintPreview={setActivePrintPreview}
          />
        )}
        {activeTab === 'checking' && (
          <PermitManagementConsole 
            records={records} 
            tempRecords={tempRecords}
            setRecords={setRecords}
            setTempRecords={setTempRecords}
            showToast={showToast}
            movementMap={movementMap}
            currentUser={currentUser}
            cloudAuthUser={cloudAuthUser}
            isViewer={isViewer}
            masterData={masterData}
            setActivePrintPreview={setActivePrintPreview}
          />
        )}
        {activeTab === 'daily' && (
          <DailyReport 
            records={records} 
            showToast={showToast} 
            movementMap={movementMap} 
            setActivePrintPreview={setActivePrintPreview} 
            dailyPdfs={dailyPdfs}
            setDailyPdfs={setDailyPdfs}
          />
        )}
        {activeTab === 'individualSearch' && (
          <IndividualSearch 
            records={records} 
            setRecords={setRecords}
            tempRecords={tempRecords}
            setTempRecords={setTempRecords}
            checkingHistory={checkingHistory} 
            setCheckingHistory={setCheckingHistory}
            dailyPdfs={dailyPdfs} 
            showToast={showToast} 
            setActivePrintPreview={setActivePrintPreview}
            setAnalyzerTarget={setAnalyzerTarget}
            dossierHistory={dossierHistory}
            setDossierHistory={setDossierHistory}
            currentUser={currentUser}
            cloudAuthUser={cloudAuthUser}
            masterData={masterData}
            movementMap={movementMap}
            passportToLatestInfo={passportToLatestInfo}
            isViewer={isViewer}
          />
        )}
        {activeTab === 'master' && (
          <MasterDB 
            masterData={masterData} 
            setMasterData={setMasterData} 
            showToast={showToast} 
            backupCombinedJSON={backupCombinedJSON}
            restoreCombinedJSON={restoreCombinedJSON}
            setSearchQuery={setSearchQuery}
            setActiveTab={setActiveTab}
            setRecords={setRecords}
            setTempRecords={setTempRecords}
            handleManualSync={handleManualSync}
            handleFetchDataFromCloud={handleFetchDataFromCloud}
            records={records}
            tempRecords={tempRecords}
          />
        )}
        {activeTab === 'tableOutput' && (
          <TableOutput 
            records={records} 
            summaries={vehicleSummaries} 
            setSummaries={setVehicleSummaries} 
            masterData={masterData}
            showToast={showToast} 
            onEditRecord={startEdit}
          />
        )}
        {activeTab === 'customReport' && (
          <CustomReportTable
            records={records}
            movementMap={movementMap}
            currentUser={currentUser}
            masterData={masterData}
            showToast={showToast}
          />
        )}
        {activeTab === 'telegraph' && (
          <TelegraphTables 
            records={records} 
            summaries={vehicleSummaries}
            currentUser={currentUser} 
            setActivePrintPreview={setActivePrintPreview} 
            showToast={showToast} 
            movementMap={movementMap}
            checkingHistory={checkingHistory}
            dossierHistory={dossierHistory}
            masterData={masterData}
            tempRecords={tempRecords}
          />
        )}
        {activeTab === 'watchList' && (
          <WatchList 
            records={records} 
            tempRecords={tempRecords}
            movementMap={movementMap} 
            checkingHistory={checkingHistory}
            masterData={masterData}
            currentUser={currentUser} 
            showToast={showToast} 
            setActivePrintPreview={setActivePrintPreview} 
            isViewer={isViewer}
            watchList={watchList}
            setWatchList={setWatchList}
            handleManualSync={handleManualSync}
            handleFetchDataFromCloud={handleFetchDataFromCloud}
          />
        )}
        {activeTab === 'info' && renderInfoTab()}
        {activeTab === 'auditLog' && (
          <ActivityLogView
            cloudAuthUser={cloudAuthUser}
            showToast={showToast}
            records={records}
            watchList={watchList}
            masterData={masterData}
            vehicleSummaries={vehicleSummaries}
            checkingHistory={checkingHistory}
          />
        )}
      </main>

      {/* Analyzer Modal */}
      <AnimatePresence>
        {tehAnalysisTarget && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
             <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setTehAnalysisTarget(null)} className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
             <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-xl rounded-2xl overflow-hidden shadow-2xl">
                <div className="p-6 bg-indigo-900 text-white flex justify-between items-center">
                   <h3 className="text-xl font-black tracking-tight flex items-center gap-3 uppercase"><RefreshCcw /> TEH Analysis Profile</h3>
                   <button onClick={() => setTehAnalysisTarget(null)} className="hover:bg-white/10 p-2 rounded-full"><X /></button>
                </div>
                <div className="p-8 space-y-6">
                   <div className="bg-gray-50 p-4 rounded-xl border border-gray-100">
                      <div className="text-xs font-black text-gray-400 uppercase mb-2">Analyzing Target</div>
                      <div className="text-lg font-black text-indigo-900">{tehAnalysisTarget.fullname}</div>
                      <div className="font-mono text-sm text-gray-500 font-bold">{tehAnalysisTarget.passport} | {tehAnalysisTarget.nationality}</div>
                   </div>

                   <div className="space-y-4">
                      {records.some(r => r.passport.toUpperCase() === tehAnalysisTarget.passport.toUpperCase()) ? (
                        <div className="p-4 bg-emerald-50 border-2 border-emerald-100 rounded-xl">
                           <div className="flex items-center gap-2 text-emerald-700 font-black uppercase text-xs mb-2">
                              <CheckCircle2 size={16} /> Matching Flight History Found
                           </div>
                           <p className="text-[10px] text-emerald-600 font-bold">
                              This person has a permanent record in Flight History. You can synchronize the Form C details from this Temporary Entry into the Flight History log.
                           </p>
                        </div>
                      ) : (
                        <div className="p-4 bg-amber-50 border-2 border-amber-100 rounded-xl">
                           <div className="flex items-center gap-2 text-amber-700 font-black uppercase text-xs mb-2">
                              <AlertCircle size={16} /> No FH Match Found
                           </div>
                           <p className="text-[10px] text-amber-600 font-bold">
                              This person exists only in Temporary History. You can "Promote" them to Flight History to create a permanent record.
                           </p>
                        </div>
                      )}
                   </div>

                   <div className="flex gap-4 pt-4">
                      <button onClick={() => setTehAnalysisTarget(null)} className="btn flex-1 bg-gray-100 text-gray-600 uppercase text-xs font-black">Cancel</button>
                      <button 
                        onClick={() => {
                          const r = records.find(rec => rec.passport.toUpperCase() === tehAnalysisTarget.passport.toUpperCase());
                          if (r) {
                            // If match, go to edit it
                            startEdit(r);
                            setTehAnalysisTarget(null);
                          } else {
                            // If no match, start new entry with this data
                            setFormData(prev => ({
                              ...tehAnalysisTarget,
                              id: undefined,
                              timestamp: undefined,
                              mode: currentMode,
                              vehicleInfo: prev.vehicleInfo,
                              officialName: prev.officialName,
                              officialTitle: prev.officialTitle
                            }));
                            setTehAnalysisTarget(null);
                            setActiveTab('entry');
                          }
                        }}
                        className="btn flex-1 bg-indigo-100 text-indigo-700 hover:bg-indigo-700 hover:text-white uppercase text-xs font-black"
                      >
                        Edit FH Link
                      </button>
                      <button 
                        onClick={() => {
                          // Logic for moving to FH
                          const tr = tehAnalysisTarget;
                          const exists = records.some(r => r.passport.toUpperCase() === tr.passport.toUpperCase());
                          if (exists) {
                            setRecords(prev => prev.map(r => r.passport.toUpperCase() === tr.passport.toUpperCase() ? { ...r, ...tr, id: r.id } : r));
                          } else {
                            setRecords(prev => [{ ...tr, id: Date.now() }, ...prev]);
                          }
                          setTempRecords(prev => prev.filter(r => r.id !== tr.id));
                          setTehAnalysisTarget(null);
                          showToast("SUCCESSFULLY MOVED TO FH");
                        }}
                        className="btn flex-1 bg-indigo-700 text-white hover:bg-black uppercase text-xs font-black shadow-lg"
                      >
                        Confirm & Move
                      </button>
                   </div>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Analyzer Modal */}
      <AnimatePresence>
        {analyzerTarget && (() => {
          const matchingLogs = [...records, ...tempRecords]
            .filter(r => r.passport.toUpperCase() === analyzerTarget.toUpperCase())
            .sort((a,b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
          const latestLog = matchingLogs[0];

          const copyDossierAsText = () => {
            if (!analyzerTarget || matchingLogs.length === 0) return;
            const bio = matchingLogs[0];
            let text = `==================================================\n`;
            text += `   FOREIGNER HISTORICAL TRAVEL DOSSIER & LOGS\n`;
            text += `==================================================\n`;
            text += `Full Name: ${bio.fullname}\n`;
            text += `Passport No: ${bio.passport.toUpperCase()}\n`;
            text += `Nationality: ${bio.nationality || '-'}\n`;
            text += `Gender: ${bio.gender === 'M' ? 'Male' : 'Female'}\n`;
            text += `Latest Address: ${bio.address || '-'}\n`;
            text += `Permit Status: ${bio.stillPermittedStatus || 'N/A'}\n`;
            if (bio.stillPermittedStatus === 'STAY PERMITTED' || bio.stillPermittedStatus === 'STILL PERMITTED') {
              text += `Permitted By: ${bio.permittedBy || '-'}\n`;
            }
            text += `\n------------------ MOVEMENT HISTORY ------------------\n`;
            matchingLogs.forEach((r, idx) => {
              text += `\n[Log #${idx + 1}] | Time: ${r.timestamp} | Mode: ${r.mode}\n`;
              text += `  - Visa Type: ${r.visaType || '-'}\n`;
              text += `  - Stay Period: ${r.stayFrom ? `${r.stayFrom} to ${r.stayTo} (${r.totalDays} Days)` : '-'}\n`;
              text += `  - Destination/Stay Location: ${r.address || '-'}\n`;
              text += `  - Carrier/Vehicle: ${r.vehicleInfo || '-'}\n`;
              text += `  - Route: ${r.arrivedFrom || '-'} --> ${r.departedTo || '-'}\n`;
              text += `  - Sponsoring Agent: ${r.broughtBy || '-'} (${r.contactDetails || '-'})\n`;
              text += `  - Registered By: ${r.officialTitle || ''} ${r.officialName || ''}\n`;
              if (r.stillPermittedStatus) {
                text += `  - Registration Permit: ${r.stillPermittedStatus} ${r.permittedBy ? `(By: ${r.permittedBy})` : ''}\n`;
              }
              if (r.remarks) {
                text += `  - Remarks: ${r.remarks}\n`;
              }
            });
            text += `\n==================================================\n`;
            text += `Report Generated: ${new Date().toLocaleString()}\n`;
            
            navigator.clipboard.writeText(text);
            showToast("DOSSIER COPIED TO CLIPBOARD!");
          };

          return (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
               <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} onClick={() => setAnalyzerTarget(null)} className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
               <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-4xl rounded-3xl overflow-hidden shadow-2xl flex flex-col h-[90vh]">
                  <div className={`p-6 text-white flex justify-between items-center bg-indigo-900`}>
                     <div className="flex items-center gap-3">
                       <Activity size={24} className="animate-pulse" />
                       <div>
                         <h3 className="text-lg font-black tracking-tight">Foreigner Complete Dossier & Travel Portfolio</h3>
                         <p className="text-[10px] text-indigo-200 uppercase font-black tracking-widest mt-0.5">Dossier Code: {analyzerTarget.toUpperCase()}</p>
                       </div>
                     </div>
                     <button onClick={() => setAnalyzerTarget(null)} className="hover:bg-white/10 p-2 rounded-full transition-colors"><X /></button>
                  </div>

                  {/* Bio block */}
                  {latestLog && (
                    <div className="bg-slate-50 p-6 border-b border-gray-100 grid grid-cols-1 md:grid-cols-4 gap-4 text-xs font-bold shrink-0">
                      <div>
                        <div className="text-[10px] text-gray-400 uppercase font-black">Full Name (အမည်)</div>
                        <div className="text-slate-800 text-sm font-black mt-0.5">{latestLog.fullname}</div>
                      </div>
                      <div>
                        <div className="text-[10px] text-gray-400 uppercase font-black">Passport Number (ပတ်စပို့)</div>
                        <div className="text-indigo-900 text-sm font-mono font-black mt-0.5">{latestLog.passport.toUpperCase()}</div>
                      </div>
                      <div>
                        <div className="text-[10px] text-gray-400 uppercase font-black">Nationality (နိုင်ငံသား)</div>
                        <div className="text-slate-800 text-sm font-black uppercase mt-0.5">{latestLog.nationality || '-'}</div>
                      </div>
                      <div>
                        <div className="text-[10px] text-gray-400 uppercase font-black">Gender (ကျား/မ)</div>
                        <div className="text-slate-800 text-sm font-black mt-0.5">{latestLog.gender === 'M' ? 'M (MALE)' : 'F (FEMALE)'}</div>
                      </div>
                    </div>
                  )}

                  {/* Action row */}
                  <div className="bg-white px-6 py-4 border-b border-gray-100 flex items-center justify-between gap-4 shrink-0">
                    <span className="text-[10px] text-amber-600 font-black uppercase tracking-widest flex items-center gap-1.5 bg-amber-50 px-2.5 py-1 rounded-full border border-amber-200">
                      ⚡ Action Console
                    </span>
                    <div className="flex gap-2">
                       <button 
                         onClick={copyDossierAsText}
                         className="btn bg-indigo-50 text-indigo-700 hover:bg-indigo-100 text-[10px] uppercase font-black px-4 py-2 rounded-lg flex items-center gap-1.5"
                       >
                         <Clipboard size={14} /> Copy as Text
                       </button>
                       <button 
                         onClick={() => {
                            setActivePrintPreview({
                              title: `Movement Dossier for Passport: ${analyzerTarget.toUpperCase()}`,
                              type: 'DOSSIER',
                              data: matchingLogs
                            });
                          }}
                         className="btn bg-emerald-600 text-white hover:bg-emerald-700 text-[10px] uppercase font-black px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-md shadow-emerald-50"
                       >
                         <Printer size={14} /> Print Dossier
                       </button>
                    </div>
                  </div>

                  {/* Registry Details container */}
                  <div className="p-6 overflow-y-auto space-y-4 flex-1">
                     <h4 className="text-[10px] text-slate-400 font-black uppercase tracking-wider mb-2">Historical Records Timeline ({matchingLogs.length} total events)</h4>
                     {matchingLogs.map((r, i) => (
                       <div key={r.id || i} className="p-5 border border-slate-100 rounded-2xl hover:border-indigo-100 hover:bg-slate-50/50 transition-all flex flex-col gap-4">
                          <div className="flex justify-between items-center border-b pb-2">
                            <div className="flex items-center gap-3">
                              <span className={`px-2.5 py-1 rounded-full text-[10px] font-black tracking-widest border ${r.mode === 'IN' ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-orange-50 text-orange-850 border-orange-200'}`}>
                                REGISTRY MODE: {r.mode}
                              </span>
                              <span className="text-[10px] font-mono text-gray-400 font-bold">{r.timestamp}</span>
                            </div>
                            <span className="text-[9px] bg-slate-100 text-gray-500 px-2 py-0.5 rounded font-bold uppercase tracking-wider">{r.logType || 'FFE'}</span>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-y-4 gap-x-6 text-xs text-slate-700">
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Visa Type & Stay Duration</span>
                               <span className="font-bold text-slate-900">{r.visaType || '-'}</span>
                               {r.stayFrom && (
                                 <span className="text-[10px] text-indigo-700 font-bold block mt-0.5">Stay: {r.stayFrom} to {r.stayTo} ({r.totalDays} Days)</span>
                               )}
                             </div>
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Address Location / Sponsor Info</span>
                               <span className="font-bold text-slate-900 block leading-tight">{r.address || '-'}</span>
                             </div>
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Carrier & Transport Vehicle</span>
                               <span className="font-bold text-slate-900 block">{r.vehicleInfo || '-'}</span>
                             </div>
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Flight Route</span>
                               <span className="font-bold text-slate-900 block">{r.arrivedFrom || '-'} ➔ {r.departedTo || '-'}</span>
                             </div>
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Agent / Bringing Up Guide</span>
                               <span className="font-bold text-slate-900 block leading-none">{r.broughtBy || '-'}</span>
                               {r.contactDetails && (
                                 <span className="text-[9px] text-[#2C6CB0] font-bold block mt-0.5">Contact: {r.contactDetails}</span>
                               )}
                             </div>
                             <div>
                               <span className="text-[9px] font-black text-slate-400 block uppercase tracking-wider">Registry Officer Responsibles</span>
                               <span className="font-bold text-slate-900 block leading-none">{r.officialTitle || 'Officer'}</span>
                               <span className="text-[10px] text-purple-700 font-bold block mt-0.5">Officer Name: {r.officialName || '-'}</span>
                             </div>
                             {r.stillPermittedStatus && (
                               <div className="md:col-span-2 lg:col-span-3 bg-indigo-50/40 p-3 rounded-xl border border-indigo-100 flex items-center justify-between">
                                  <div>
                                    <span className="text-[9px] font-black text-indigo-900 block uppercase tracking-wider">🟢 Dynamic Permit Status (ခွင့်ပြုချက် အခြေအနေ)</span>
                                    <span className="text-sm font-black text-indigo-950 uppercase">{r.stillPermittedStatus}</span>
                                  </div>
                                  {(r.stillPermittedStatus === 'STAY PERMITTED' || r.stillPermittedStatus === 'STILL PERMITTED') && r.permittedBy && (
                                    <div className="text-right">
                                      <span className="text-[9px] font-black text-indigo-900 block uppercase tracking-wider">📝 Permitted By (ခွင့်ပြုသူ)</span>
                                      <span className="font-black text-slate-800 uppercase">{r.permittedBy}</span>
                                    </div>
                                  )}
                               </div>
                             )}
                          </div>

                          {r.remarks && (
                            <div className="text-[11px] bg-slate-100 p-2.5 rounded-lg text-slate-600 border border-slate-200">
                              <span className="font-black text-[9px] uppercase block mb-1 text-slate-400">System Remarks & Notes</span>
                              {r.remarks}
                            </div>
                          )}
                       </div>
                     ))}
                  </div>
               </motion.div>
            </div>
          );
        })()}
      </AnimatePresence>

      <AnimatePresence>
        {ffeConfirmation && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 no-print">
             <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/80 backdrop-blur-md" />
             <motion.div 
               initial={{ scale: 0.9, opacity: 0, y: 20 }} 
               animate={{ scale: 1, opacity: 1, y: 0 }} 
               exit={{ scale: 0.9, opacity: 0, y: 20 }} 
               className={`relative bg-white w-full max-w-lg rounded-[2.5rem] overflow-hidden shadow-2xl border-4 ${currentMode === 'IN' ? 'border-emerald-500' : 'border-orange-500'}`}
             >
                <div className="p-8 text-center space-y-6">
                  <div className={`w-20 h-20 rounded-full flex items-center justify-center mx-auto ${currentMode === 'IN' ? 'bg-emerald-100 text-emerald-600' : 'bg-orange-100 text-orange-600'}`}>
                    <ShieldCheck size={40} />
                  </div>
                  <div className="space-y-4">
                    <h3 className="text-2xl font-black text-gray-900 leading-tight">သေချာပါသလား?</h3>
                    <p className="text-gray-600 font-medium leading-relaxed px-4 text-xs">
                      {ffeConfirmation.message}
                    </p>
                  </div>
                  <div className="flex gap-4 pt-4">
                    <button 
                      onClick={() => setFfeConfirmation(null)}
                      className="flex-1 py-4 px-6 rounded-2xl bg-gray-100 text-gray-500 font-black uppercase text-sm hover:bg-gray-200 transition-colors"
                    >
                      ပြင်ဆင်ရန် (NO)
                    </button>
                    <button 
                      onClick={executeSaveRecord}
                      className={`flex-1 py-4 px-6 rounded-2xl text-white font-black uppercase text-sm shadow-xl transition-colors ${currentMode === 'IN' ? 'bg-emerald-600 hover:bg-emerald-900 shadow-emerald-200' : 'bg-orange-600 hover:bg-orange-900 shadow-orange-200'}`}
                    >
                      မှန်ကန်ပါသည်။ (YES)
                    </button>
                  </div>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {ffeSuccessMsg && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center pointer-events-none no-print">
             <motion.div 
               initial={{ scale: 0.5, opacity: 0, y: 50 }} 
               animate={{ scale: 1, opacity: 1, y: 0 }} 
               exit={{ scale: 1.5, opacity: 0 }} 
               className="bg-gray-900 text-white px-10 py-6 rounded-full shadow-2xl flex items-center gap-4 border-2 border-white/20 backdrop-blur-xl"
             >
                <div className={`w-10 h-10 rounded-full flex items-center justify-center text-white ${currentMode === 'IN' ? 'bg-emerald-500' : 'bg-orange-500'}`}>
                  <Check size={24} strokeWidth={4} />
                </div>
                <div className="flex flex-col">
                  <span className="text-[10px] font-black uppercase opacity-60 tracking-widest">Entry Recorded</span>
                  <span className="text-lg font-black">{ffeSuccessMsg}</span>
                </div>
             </motion.div>
          </div>
        )}
      </AnimatePresence>

      <footer className="py-20 text-center no-print">
        <div className="opacity-20 flex justify-center gap-2 items-center text-sm font-black grayscale">
           <Users size={16} /> 2ND DIVISION LOGISTICS UNIT
        </div>
      </footer>
    </div>


    {/* On-screen Interactive Print Preview Modal Overlay */}
    {activePrintPreview && (
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 md:p-8 z-50 no-print animate-in fade-in duration-200">
        <div className="bg-slate-900 text-white rounded-2xl w-full max-w-5xl h-full max-h-[92vh] flex flex-col shadow-2xl border border-slate-700/60 overflow-hidden">
          {/* Header Bar */}
          <div className="bg-slate-950 p-4 border-b border-slate-800 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-xl">🖨️</span>
              <div>
                <h3 className="font-black text-sm uppercase tracking-wide">Interactive Visual Print Preview</h3>
                <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">
                  Type: {activePrintPreview.type} • Entries Check Count: {activePrintPreview.data.length}
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2.5">
              <button
                onClick={() => { window.focus(); window.print(); }}
                className="btn bg-indigo-600 text-white hover:bg-indigo-700 font-extrabold text-[11px] uppercase tracking-wider px-5 py-2.5 rounded-xl shadow-md transition-all flex items-center gap-2"
              >
                <Printer size={13} /> CONFIRM & PRINT
              </button>
              <button
                onClick={() => setActivePrintPreview(null)}
                className="btn bg-slate-800 text-slate-300 hover:bg-slate-700 font-extrabold text-[11px] uppercase tracking-wider px-5 py-2.5 rounded-xl border border-slate-700 transition-all"
              >
                CLOSE PREVIEW
              </button>
            </div>
          </div>

          {/* Paper Content Space */}
          <div className="flex-1 bg-slate-800 p-6 overflow-y-auto">
            <div className="bg-white text-black p-8 md:p-12 shadow-inner rounded-xl mx-auto max-w-4xl text-xs font-sans leading-relaxed select-text min-h-[1050px]">
              
              {/* Document Header */}
              {activePrintPreview.type !== 'TELEGRAPH_REPORT' && (
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4 text-black">
                  {activePrintPreview.type.startsWith('SIA_') || activePrintPreview.type === 'STILL_IN_ANALYTICS' ? (
                    <>
                      <h1 className="text-sm font-black uppercase tracking-wider text-slate-800">Second Division, Department of Immigration (Myeik)</h1>
                      <h2 className="text-md font-extrabold uppercase tracking-tight text-indigo-950">STATISTICAL INVESTIGATION ANALYSIS (SIA)</h2>
                      <p className="text-[10px] font-bold text-gray-600 uppercase tracking-widest">{activePrintPreview.title}</p>
                    </>
                  ) : (
                    <h1 className="text-md font-black uppercase tracking-tight">{activePrintPreview.title}</h1>
                  )}
                  <p className="text-[9px] text-gray-500 font-mono uppercase tracking-wide">
                    SYSTEM REGISTER COPY • TIMESTAMP: {new Date().toLocaleString()}
                  </p>
                </div>
              )}

              {/* Document Body */}
              {activePrintPreview.type === 'DAILY_REPORT' && (
                <div>
                  <div className="text-center mb-6">
                    <h2 className="text-xs font-bold text-gray-700 tracking-wider uppercase border border-black inline-block px-4 py-1">Daily Movement Log Report</h2>
                    <div className="flex justify-center gap-6 text-[10px] font-bold pt-1.5">
                      <span>DATE: {activePrintPreview.meta?.reportDate || ''}</span>
                      <span>TOTAL LOGS: {activePrintPreview.data.length}</span>
                    </div>
                  </div>
                  
                  <table className="w-full text-left text-[9px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-gray-100 uppercase text-[8px] font-black">
                        <th className="p-1 border border-gray-400 text-center">#</th>
                        <th className="p-1 border border-gray-400">Time</th>
                        <th className="p-1 border border-gray-400">Passport & Fullname</th>
                        <th className="p-1 border border-gray-400 text-center">Dir</th>
                        <th className="p-1 border border-gray-400">Nationality</th>
                        <th className="p-1 border border-gray-400">Visa Type</th>
                        <th className="p-1 border border-gray-400">Destination/Address</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((r: any, idx: number) => (
                        <tr key={idx} className="border-b border-gray-300 align-top">
                          <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                          <td className="p-1 border border-gray-400 font-mono font-bold text-gray-700">{r.timestamp ? r.timestamp.split(', ')[1] || r.timestamp : ''}</td>
                          <td className="p-1 border border-gray-400 leading-tight">
                            <span className="font-extrabold font-mono text-[#1a365d] mr-1.5">{r.passport}</span>
                            <span className="font-bold text-gray-800 uppercase">{r.fullname}</span>
                          </td>
                          <td className="p-1 border border-gray-400 text-center">
                            <span className={`px-1 py-0.5 rounded text-[8px] font-black ${
                              r.direction === 'IN' 
                                ? 'bg-emerald-50 text-emerald-800 border border-emerald-300' 
                                : 'bg-rose-50 text-rose-800 border border-rose-300'
                            }`}>{r.direction}</span>
                          </td>
                          <td className="p-1 border border-gray-400 font-bold uppercase text-gray-700">{r.nationality}</td>
                          <td className="p-1 border border-gray-400 font-mono text-gray-600">{r.visaType || '-'}</td>
                          <td className="p-1 border border-gray-400 leading-tight">
                            <div className="font-bold text-gray-800">{r.address}</div>
                            {r.stayDescription && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 {r.stayDescription}</div>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {activePrintPreview.type === 'STILL_IN_ANALYTICS' && (
                <div>
                  <div className="text-center mb-6">
                    <h2 className="text-xs font-bold text-gray-700 tracking-wider uppercase border border-black inline-block px-4 py-1">Active Stay-In Guests Ledger</h2>
                    <div className="flex justify-center gap-6 text-[10px] font-bold pt-1.5">
                      <span>TOTAL ACTIVE GUESTS: {activePrintPreview.data.length}</span>
                    </div>
                  </div>
                  
                  <table className="w-full text-left text-[9px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-gray-100 uppercase text-[8px] font-black">
                        <th className="p-1.5 border border-gray-400 text-center">#</th>
                        <th className="p-1.5 border border-gray-400">Passport / Guest</th>
                        <th className="p-1.5 border border-gray-400">Nationality</th>
                        <th className="p-1.5 border border-gray-400">Stay Address / Description</th>
                        <th className="p-1.5 border border-gray-400">Stay Duration Schedule</th>
                        <th className="p-1.5 border border-gray-400 text-center">Approved Days Remaining</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((r: any, idx: number) => (
                        <tr key={idx} className="border-b border-gray-300 align-top">
                          <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                          <td className="p-1.5 border border-gray-455 leading-tight">
                            <div className="font-extrabold font-mono text-[#1a365d] uppercase">{r.p}</div>
                            <div className="font-bold text-gray-800 uppercase mt-0.5">{r.n}</div>
                          </td>
                          <td className="p-1.5 border border-gray-450 text-gray-700 uppercase font-bold">{r.nat}</td>
                          <td className="p-1.5 border border-gray-455 leading-tight">
                            <div className="font-bold text-gray-900">{r.loc}</div>
                            {r.stayDescription && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Depth: {r.stayDescription}</div>}
                          </td>
                          <td className="p-1.5 border border-gray-450 font-mono text-[9px] text-gray-800">
                            <div>{r.start} to {r.end}</div>
                          </td>
                          <td className="p-1.5 border border-gray-450 text-center leading-tight">
                            <span className="font-extrabold font-mono text-indigo-700 block">{getLiveRemainingDays(r.end)} Days</span>
                            <span className="text-[7px] text-slate-400 font-semibold uppercase">Allowed: {r.allowed || 'N/A'}</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {activePrintPreview.type === 'DOSSIER' && (
                <div>
                  <h2 className="text-xs font-black uppercase text-center border border-black py-1.5 tracking-wider mb-6">Foreigner Dossier File Card Index</h2>
                  <div className="grid grid-cols-2 gap-4 pb-6 mb-6 border-b border-dashed border-gray-300 text-[10px]">
                    <div>
                      <span className="block text-gray-400 text-[8px] uppercase font-black">Full Name / အမည်အပြည့်အစုံ</span>
                      <strong className="text-xs font-black text-slate-900 uppercase">{activePrintPreview.data[0]?.fullname || '-'}</strong>
                    </div>
                    <div>
                      <span className="block text-gray-400 text-[8px] uppercase font-black">Passport Number / ပတ်စ်ပို့နံပါတ်</span>
                      <strong className="text-xs font-black text-indigo-900 font-mono uppercase">{activePrintPreview.data[0]?.passport || '-'}</strong>
                    </div>
                    <div>
                      <span className="block text-gray-400 text-[8px] uppercase font-black">Nationality / နိုင်ငံသား</span>
                      <span className="font-black uppercase text-gray-700">{activePrintPreview.data[0]?.nationality || '-'}</span>
                    </div>
                    <div>
                      <span className="block text-gray-400 text-[8px] uppercase font-black">Visa Category</span>
                      <span className="font-black text-gray-700">{activePrintPreview.data[0]?.visaType || '-'}</span>
                    </div>
                  </div>

                  <h3 className="text-[10px] font-black uppercase tracking-wider mb-2 border-b pb-1">Verified Verification History Lists</h3>
                  <table className="w-full text-left text-[9px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-gray-100 uppercase text-[8px] font-black">
                        <th className="p-1.5 border border-gray-400 text-center">#</th>
                        <th className="p-1.5 border border-gray-400">Timestamp</th>
                        <th className="p-1.5 border border-gray-400">Physical Stay Direction / Address</th>
                        <th className="p-1.5 border border-gray-400">Verification Status</th>
                        <th className="p-1.5 border border-gray-400">Officer Stamp</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((r: any, idx: number) => (
                        <tr key={idx} className="border-b border-gray-300 align-top">
                          <td className="p-1.5 border border-gray-400 text-center font-bold">{idx + 1}</td>
                          <td className="p-1.5 border border-gray-400 font-mono">{r.timestamp}</td>
                          <td className="p-1.5 border border-gray-400 leading-tight">
                            <span className={`inline-block px-1 rounded text-[7.5px] font-black mr-1 text-white ${
                              r.direction === 'IN' ? 'bg-emerald-600' : 'bg-rose-600'
                            }`}>{r.direction}</span>
                            <span className="font-bold text-gray-800">{r.address}</span>
                            {r.stayDescription && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {r.stayDescription}</div>}
                          </td>
                          <td className="p-1.5 border border-gray-400">
                            <span className="font-black uppercase text-[#1a202c]">{r.stillPermittedStatus || 'REGISTERED'}</span>
                            {r.permittedBy && <div className="text-[7.5px] text-purple-700 mt-0.5 font-bold">Desc: {r.permittedBy}</div>}
                          </td>
                          <td className="p-1.5 border border-gray-400 font-mono text-[8px] text-gray-500">
                            {r.officialName || '-'} ({r.officialTitle || 'Officer'})
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {(activePrintPreview.type === 'CO_LTD_PENDING' || activePrintPreview.type === 'OTHERS_PENDING' || activePrintPreview.type === 'CHECKED_HISTORY') && (
                <div>
                  <table className="w-full text-left text-[9px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-gray-100 uppercase text-[8px] font-black">
                        <th className="p-1.5 border border-gray-400 text-center">#</th>
                        <th className="p-1.5 border border-gray-400">Passport & Guest Name</th>
                        <th className="p-1.5 border border-gray-400">Nationality</th>
                        {activePrintPreview.type === 'CHECKED_HISTORY' ? (
                          <>
                            <th className="p-1.5 border border-gray-400">Confirmed Stay Address & Detail Description</th>
                            <th className="p-1.5 border border-gray-400 text-center">Status</th>
                            <th className="p-1.5 border border-gray-400">Officer details</th>
                          </>
                        ) : (
                          <>
                            <th className="p-1.5 border border-gray-400">Expected Stay Address & Detail Description</th>
                            <th className="p-1.5 border border-gray-400">Stay Schedule</th>
                            <th className="p-1.5 border border-gray-400 min-w-[120px]">Verification Signature Box</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((h: any, idx: number) => {
                        const isHist = h.status !== undefined;
                        const pNo = h.passport || h.p || '';
                        const name = h.fullname || h.n || '';
                        const nat = h.nationality || h.nat || '';
                        const address = h.confirmedAddress || h.loc || '';
                        const addressDesc = h.confirmedStayDescription || (masterData.find(m => m.type === 'Stay' && m.name.toLowerCase() === (h.loc || '').toLowerCase())?.linkedValue || h.stayDescription);

                        return (
                          <tr key={idx} className="border-b border-gray-300 align-top">
                            <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                            <td className="p-1.5 border border-gray-450 leading-tight">
                              <div className="font-extrabold font-mono uppercase text-[#2C6CB0]">{pNo}</div>
                              <div className="font-bold text-gray-800 uppercase mt-0.5">{name}</div>
                            </td>
                            <td className="p-1.5 border border-gray-455 font-black uppercase text-gray-700">{nat}</td>
                            
                            {isHist ? (
                              <>
                                <td className="p-1.5 border border-gray-450 leading-tight">
                                  <div className="font-extrabold text-gray-900">{address}</div>
                                  {addressDesc && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                                </td>
                                <td className="p-1.5 border border-gray-450 text-center">
                                  <span className="font-black uppercase text-emerald-800 bg-emerald-50 border border-emerald-300 rounded px-1.5 py-0.5 text-[8px]">{h.status}</span>
                                </td>
                                <td className="p-1.5 border border-gray-450 font-mono text-[8px] leading-tight text-gray-600">
                                  <div>{h.officerTitle || 'Officer'} {h.officerName || 'System'}</div>
                                  <div className="text-[7px] text-slate-400 mt-0.5">{h.checkDate}</div>
                                </td>
                              </>
                            ) : (
                              <>
                                <td className="p-1.5 border border-gray-455 leading-tight">
                                  <div className="font-extrabold text-gray-900">{address}</div>
                                  {addressDesc && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                                </td>
                                <td className="p-1.5 border border-gray-455 leading-tight font-mono text-[9px]">
                                  <div>{h.start} to {h.end}</div>
                                  <div className="text-[8px] text-slate-400 mt-1 uppercase">Rem: {getLiveRemainingDays(h.end)} Days</div>
                                </td>
                                <td className="p-1.5 border border-gray-455 text-center bg-slate-50/50 min-w-[130px]">
                                  <div className="h-10 border border-dashed border-slate-300 rounded mb-1 bg-white flex items-center justify-center">
                                    <span className="text-[6.5px] text-slate-400">STAMP & SIGN</span>
                                  </div>
                                </td>
                              </>
                            )}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}

              {activePrintPreview.type === 'SIA_NATIONALITY_REPORT' && (
                <div className="space-y-4">
                  <div className="border border-slate-200 rounded-xl overflow-hidden">
                    <table className="w-full text-left text-[9px] border-collapse">
                      <thead>
                        <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-slate-250">
                          <th className="p-2 border-r border-slate-250 text-center w-12">No.</th>
                          <th className="p-2 border-r border-slate-250">Nationality</th>
                          <th className="p-2 border-r border-slate-250 text-center">Inbound (IN)</th>
                          <th className="p-2 text-center">Outbound (OUT)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activePrintPreview.data.map((r: any, idx: number) => (
                          <tr key={idx} className="border-b border-slate-200 hover:bg-slate-50/55">
                            <td className="p-2 border-r border-slate-200 text-center font-bold">{idx + 1}</td>
                            <td className="p-2 border-r border-slate-200 font-extrabold text-slate-800 uppercase">{r.country}</td>
                            <td className="p-2 border-r border-slate-200 text-center font-bold text-emerald-700">{r.in}</td>
                            <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {activePrintPreview.type === 'SIA_VISA_REPORT' && (
                <div className="space-y-4">
                  <div className="border border-slate-200 rounded-xl overflow-hidden">
                    <table className="w-full text-left text-[9px] border-collapse">
                      <thead>
                        <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-slate-250">
                          <th className="p-2 border-r border-slate-250 text-center w-12">No.</th>
                          <th className="p-2 border-r border-slate-250">Visa Type</th>
                          <th className="p-2 border-r border-slate-250 text-center">Inbound (IN)</th>
                          <th className="p-2 text-center">Outbound (OUT)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activePrintPreview.data.map((r: any, idx: number) => (
                          <tr key={idx} className="border-b border-slate-200 hover:bg-slate-50/55">
                            <td className="p-2 border-r border-slate-200 text-center font-bold">{idx + 1}</td>
                            <td className="p-2 border-r border-slate-200 font-extrabold text-slate-800 uppercase">{r.visaType}</td>
                            <td className="p-2 border-r border-slate-200 text-center font-bold text-emerald-700">{r.in}</td>
                            <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {activePrintPreview.type === 'SIA_FULL_REPORT' && (
                <div className="space-y-8">
                  <div>
                    <h3 className="text-xs font-black uppercase text-indigo-950 mb-2 border-b pb-1">1. Nationality Inflow & Outflow</h3>
                    <div className="border border-slate-200 rounded-xl overflow-hidden">
                      <table className="w-full text-left text-[9px] border-collapse">
                        <thead>
                          <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-slate-250">
                            <th className="p-2 border-r border-slate-250 text-center w-12">No.</th>
                            <th className="p-2 border-r border-slate-250">Nationality</th>
                            <th className="p-2 border-r border-slate-250 text-center">Inbound (IN)</th>
                            <th className="p-2 text-center">Outbound (OUT)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {activePrintPreview.data.nationalityList.map((r: any, idx: number) => (
                            <tr key={idx} className="border-b border-slate-200 hover:bg-slate-50/55">
                              <td className="p-2 border-r border-slate-200 text-center font-bold">{idx + 1}</td>
                              <td className="p-2 border-r border-slate-200 font-extrabold text-slate-800 uppercase">{r.country}</td>
                              <td className="p-2 border-r border-slate-200 text-center font-bold text-emerald-700">{r.in}</td>
                              <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div>
                    <h3 className="text-xs font-black uppercase text-indigo-950 mb-2 border-b pb-1">2. Visa Type Distribution</h3>
                    <div className="border border-slate-200 rounded-xl overflow-hidden">
                      <table className="w-full text-left text-[9px] border-collapse">
                        <thead>
                          <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-slate-250">
                            <th className="p-2 border-r border-slate-250 text-center w-12">No.</th>
                            <th className="p-2 border-r border-slate-250">Visa Type</th>
                            <th className="p-2 border-r border-slate-250 text-center">Inbound (IN)</th>
                            <th className="p-2 text-center">Outbound (OUT)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {activePrintPreview.data.visaList.map((r: any, idx: number) => (
                            <tr key={idx} className="border-b border-slate-200 hover:bg-slate-50/55">
                              <td className="p-2 border-r border-slate-200 text-center font-bold">{idx + 1}</td>
                              <td className="p-2 border-r border-slate-200 font-extrabold text-slate-800 uppercase">{r.visaType}</td>
                              <td className="p-2 border-r border-slate-200 text-center font-bold text-emerald-700">{r.in}</td>
                              <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {activePrintPreview.type === 'TELEGRAPH_REPORT' && (
                <div className="font-pyidaungsu text-black relative select-text p-[0.5in] text-[13px] space-y-6">
                  {/* Absolute Top Title */}
                  <div className="text-center mb-4">
                    <span className="text-[13px] font-black uppercase tracking-wider block">
                      မှတ်ပုံတင်ကြေးနန်းစာပုံစံအစား
                    </span>
                  </div>

                  {/* Letter Header */}
                  <div className="flex justify-between items-start text-[13px] font-bold pb-2">
                    <div className="space-y-1">
                      <div>(ပ) {activePrintPreview.data.chief1}</div>
                      <div>(လ) {activePrintPreview.data.chief2}</div>
                      <div>(တ) {activePrintPreview.data.chief3}</div>
                    </div>
                    <div className="text-center space-y-0.5">
                      <div>ရက်စွဲ / အချိန်</div>
                      <div className="text-slate-400 font-light leading-none -mt-1">—————-</div>
                      <div className="font-bold">
                        {toBurmeseShortDateTime(activePrintPreview.data.selectedDate, activePrintPreview.data.reportTime)}
                      </div>
                    </div>
                  </div>

                  {/* Opening Line above title */}
                  <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

                  {/* Letter Info Line with Side-by-Side Letter No and Telegraph Title with one tab space separation */}
                  <div className="flex items-center gap-12 text-[13px] font-bold py-1 text-black">
                    <div>{activePrintPreview.data.letterNo}</div>
                    <div className="font-black text-[13px] uppercase">{activePrintPreview.data.telegraphTitle}</div>
                  </div>

                  {/* Paragraph 1 */}
                  <p className="text-[13px] leading-relaxed text-black font-medium whitespace-pre-line my-4">
                    {(activePrintPreview.data.customIntroText || '')
                      .replace('[DATE]', toBurmeseSlashDate(activePrintPreview.data.selectedDate))
                      .replace('[FLIGHT_COUNT]', toBurmeseDigits((activePrintPreview.data.rows || []).length))
                    }
                  </p>

                  {/* Table */}
                  <div className="overflow-x-auto w-full">
                    <table className="w-full text-center text-[13px] border-collapse border border-black font-pyidaungsu my-4">
                      <thead>
                        <tr className="bg-slate-50 uppercase font-black border-b border-black">
                          <th rowSpan={3} className="p-1 border border-black text-center align-middle w-10">စဉ်</th>
                          <th rowSpan={3} className="p-1 border border-black text-center align-middle w-[110px] min-w-[110px] max-w-[110px]">လေကြောင်းလိုင်းအမည်</th>
                          <th rowSpan={3} className="p-1 border border-black text-center align-middle w-16">ဝင်/ထွက်ချိန်</th>
                          <th colSpan={6} className="p-1 border border-black text-center font-black">မြန်မာနိုင်ငံသား</th>
                          <th colSpan={6} className="p-1 border border-black text-center font-black">နိုင်ငံခြားသား</th>
                          <th rowSpan={3} className="p-1 border border-black text-center align-middle">မှတ်ချက်</th>
                        </tr>
                        <tr className="bg-slate-50 border-b border-black">
                          <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အဝင်</th>
                          <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အထွက်</th>
                          <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အဝင်</th>
                          <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အထွက်</th>
                        </tr>
                        <tr className="bg-slate-50 border-b border-black text-[13px] font-bold text-slate-600">
                          {/* MMR In */}
                          <th className="p-1 border border-black w-8">က</th>
                          <th className="p-1 border border-black w-8">မ</th>
                          <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                          {/* MMR Out */}
                          <th className="p-1 border border-black w-8">က</th>
                          <th className="p-1 border border-black w-8">မ</th>
                          <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                          {/* FRN In */}
                          <th className="p-1 border border-black w-8">က</th>
                          <th className="p-1 border border-black w-8">မ</th>
                          <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                          {/* FRN Out */}
                          <th className="p-1 border border-black w-8">က</th>
                          <th className="p-1 border border-black w-8">မ</th>
                          <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activePrintPreview.data.rows && activePrintPreview.data.rows.map((row: any, idx: number) => {
                          const myanmarInTotal = row.myanmarInM + row.myanmarInF;
                          const myanmarOutTotal = row.myanmarOutM + row.myanmarOutF;
                          const foreignerInTotal = row.foreignerInM + row.foreignerInF;
                          const foreignerOutTotal = row.foreignerOutM + row.foreignerOutF;

                          const fNum = (val: number | string) => activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(val) : val;

                          return (
                            <tr key={row.id} className="border-b border-black">
                              <td className="p-1.5 border border-black text-center font-bold">{fNum(idx + 1)}</td>
                              <td className="p-1.5 border border-black font-extrabold uppercase text-left w-[110px] min-w-[110px] max-w-[110px] break-words whitespace-normal text-wrap leading-tight text-[12px]">{row.flightName}</td>
                              <td className="p-1.5 border border-black text-center font-mono">{fNum(row.time)}</td>
                              
                              {/* MMR In */}
                              <td className="p-1.5 border border-black text-center">{fNum(row.myanmarInM)}</td>
                              <td className="p-1.5 border border-black text-center">{fNum(row.myanmarInF)}</td>
                              <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(myanmarInTotal)}</td>
                              
                              {/* MMR Out */}
                              <td className="p-1.5 border border-black text-center">{fNum(row.myanmarOutM)}</td>
                              <td className="p-1.5 border border-black text-center">{fNum(row.myanmarOutF)}</td>
                              <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(myanmarOutTotal)}</td>
                              
                              {/* FRN In */}
                              <td className="p-1.5 border border-black text-center">{fNum(row.foreignerInM)}</td>
                              <td className="p-1.5 border border-black text-center">{fNum(row.foreignerInF)}</td>
                              <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(foreignerInTotal)}</td>
                              
                              {/* FRN Out */}
                              <td className="p-1.5 border border-black text-center">{fNum(row.foreignerOutM)}</td>
                              <td className="p-1.5 border border-black text-center">{fNum(row.foreignerOutF)}</td>
                              <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(foreignerOutTotal)}</td>
                              
                              <td className="p-1.5 border border-black text-left text-[13px] font-bold text-slate-700">{row.remarks}</td>
                            </tr>
                          );
                        })}

                        {/* Totals Row */}
                        <tr className="bg-slate-100 font-bold border-b border-black text-[13px]">
                          <td colSpan={3} className="p-2 border border-black text-center font-black">Total</td>
                          
                          {/* MMR In Total */}
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM, 0)}</td>
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInF, 0)}</td>
                          <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM + b.myanmarInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM + b.myanmarInF, 0)}</td>
                          
                          {/* MMR Out Total */}
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM, 0)}</td>
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutF, 0)}</td>
                          <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM + b.myanmarOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM + b.myanmarOutF, 0)}</td>
                          
                          {/* FRN In Total */}
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM, 0)}</td>
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInF, 0)}</td>
                          <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM + b.foreignerInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM + b.foreignerInF, 0)}</td>
                          
                          {/* FRN Out Total */}
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM, 0)}</td>
                          <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutF, 0)}</td>
                          <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM + b.foreignerOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM + b.foreignerOutF, 0)}</td>
                          
                          <td className="p-2 border border-black bg-slate-50"></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Paragraph 2 - (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... / */}
                  <div className="text-[13px] font-normal mt-4">
                    (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... /
                  </div>

                  {/* ပိတ်မျဉ်း အပြည့်တစ်ကြောင်း below သိရှိနိုင်ပါရန်တင်ပြအပ် */}
                  <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

                  {/* Footer Text */}
                  <div className="flex justify-between items-start text-[13px] font-bold leading-relaxed">
                    <div className="space-y-1">
                      <div className="font-medium">သင့်ရာနည်းဖြင့်</div>
                    </div>
                    <div className="text-center space-y-0.5 pr-8">
                      <div>အဆင့်အတန်း</div>
                      <div className="text-slate-400 font-light leading-none -mt-1">—————</div>
                      <div className="font-black text-rose-700">{activePrintPreview.data.urgentLevel}</div>
                    </div>
                  </div>

                  {/* Signature block */}
                  <div className="flex justify-start mt-10 pl-4">
                    <div className="text-center space-y-1 font-pyidaungsu max-w-sm flex flex-col items-center">
                      <div className="text-[13px] font-extrabold text-black">
                        {activePrintPreview.data.officerTitle}
                      </div>
                      <div className="text-[13px] font-bold text-black leading-tight">
                        {activePrintPreview.data.officerName}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Document Lightweight Verification Badge instead of heavy QRs */}
              {activePrintPreview && (
                <div className="mt-8 pt-4 border-t border-slate-200 flex flex-col sm:flex-row justify-between items-center gap-4 text-[9px] text-slate-500 font-sans">
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                    <span>OFFICIAL VERIFIED SYSTEM RECORD • MYEIK IMMIGRATION</span>
                  </div>
                  <div>
                    <span>Operator: {currentUser ? `${currentUser.title} ${currentUser.name}` : `${localStorage.getItem('lastCheckedOfficerTitle') || 'Officer'} ${localStorage.getItem('lastCheckedOfficerName') || 'Myeik Duty Officer'}`}</span>
                    <span className="mx-2">•</span>
                    <span>Printed: {new Date().toLocaleString()}</span>
                  </div>
                </div>
              )}

              {/* Document Footer */}
              <div className="mt-8 pt-4 border-t border-dashed border-gray-300 text-center text-[9px] text-gray-400 leading-normal">
                <p>Myeik Immigration Enforcement Unit • Statistical Investigation Analysis</p>
                <p>CONFIDENTIAL • DEPARTMENT OF IMMIGRATION & REGISTRATION RECORD</p>
              </div>

            </div>
          </div>
        </div>
      </div>
    )}


    {/* Standard Hidden Print Layout (Used for direct printing) */}
    {activePrintPreview && (
      <div className="hidden print:block bg-white p-8 text-black min-h-screen w-full text-xs font-sans print-content leading-relaxed">
            
            {activePrintPreview.type === 'DAILY_REPORT' && (
              <div>
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-md font-black uppercase text-gray-900">Second Division, Immigration Department (Myeik)</h1>
                  <h2 className="text-sm font-bold text-gray-700 tracking-widest uppercase border border-black inline-block px-6 py-1">Daily Movement Log Report</h2>
                  <div className="flex justify-center gap-8 text-[11px] font-bold pt-2">
                    <span>DATE: {activePrintPreview.meta?.reportDate || ''}</span>
                    <span>TOTAL LOGS: {activePrintPreview.data.length}</span>
                  </div>
                </div>
                
                <table className="w-full text-left text-[10px] border border-gray-450 border-collapse table-auto">
                  <thead>
                    <tr className="bg-gray-100 uppercase text-[8px] font-black">
                      <th className="p-1.5 border border-gray-450 text-center">#</th>
                      <th className="p-1.5 border border-gray-450">Time</th>
                      <th className="p-1.5 border border-gray-450">Passport & Fullname</th>
                      <th className="p-1.5 border border-gray-450 text-center">Dir</th>
                      <th className="p-1.5 border border-gray-450">Nationality</th>
                      <th className="p-1.5 border border-gray-450">Visa Type</th>
                      <th className="p-1.5 border border-gray-450">Stay Location & Description</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activePrintPreview.data.map((r: any, idx: number) => (
                      <tr key={idx} className="border-b border-gray-450 align-top">
                        <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                        <td className="p-1.5 border border-gray-450 text-center font-mono text-[9px]">{r.timestamp?.split(', ')[1] || '00:00'}</td>
                        <td className="p-1.5 border border-gray-450 text-[10px]">
                          <div className="font-extrabold font-mono uppercase text-indigo-950">{r.passport}</div>
                          <div className="text-[9px] text-gray-600 font-bold uppercase">{r.fullname}</div>
                        </td>
                        <td className="p-1.5 border border-gray-450 text-center">
                          <span className={`px-1 rounded text-[8px] font-black ${r.mode === 'IN' ? 'bg-emerald-50 text-emerald-800' : 'bg-orange-50 text-orange-850'}`}>
                            {r.mode}
                          </span>
                        </td>
                        <td className="p-1.5 border border-gray-450 font-bold uppercase">{r.nationality}</td>
                        <td className="p-1.5 border border-gray-450 font-bold">{r.visaType || 'N/A'}</td>
                        <td className="p-1.5 border border-gray-450 text-[9px] leading-tight">
                          <div className="font-bold text-gray-800">{r.address || 'N/A'}</div>
                          {r.stayDescription && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {r.stayDescription}</div>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activePrintPreview.type === 'STILL_IN_ANALYTICS' && (
              <div>
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-md font-black uppercase text-gray-900">Second Division, Immigration Department (Myeik)</h1>
                  <h2 className="text-sm font-bold text-gray-700 tracking-widest uppercase border border-black inline-block px-6 py-1">Active Stay Inbound Visitors Analysis</h2>
                  <div className="flex justify-center gap-8 text-[11px] font-bold pt-2">
                    <span>COMPILED TIMESTAMP: {new Date().toLocaleString()}</span>
                    <span>TOTAL ACTIVE GUESTS: {activePrintPreview.data.length}</span>
                  </div>
                </div>
                
                <table className="w-full text-left text-[10px] border border-gray-450 border-collapse">
                  <thead>
                    <tr className="bg-gray-100 uppercase text-[8px] font-black">
                      <th className="p-1.5 border border-gray-450 text-center">#</th>
                      <th className="p-1.5 border border-gray-450">Passport No</th>
                      <th className="p-1.5 border border-gray-450">Full Name</th>
                      <th className="p-1.5 border border-gray-450">Nationality</th>
                      <th className="p-1.5 border border-gray-450">Stay Address & Detail Description</th>
                      <th className="p-1.5 border border-gray-450">Stay Schedule & Remaining Days</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activePrintPreview.data.map((r: any, idx: number) => (
                      <tr key={idx} className="border-b border-gray-450 align-top">
                        <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                        <td className="p-1.5 border border-gray-450 font-mono font-extrabold uppercase text-indigo-950">{r.p}</td>
                        <td className="p-1.5 border border-gray-450 uppercase font-black">{r.n}</td>
                        <td className="p-1.5 border border-gray-450 uppercase font-bold">{r.nat}</td>
                        <td className="p-1.5 border border-gray-450 leading-tight">
                          <div className="font-bold text-gray-800">{r.loc || 'N/A'}</div>
                          {(() => {
                            const addressDesc = masterData.find(m => m.type === 'Stay' && m.name.toLowerCase() === r.loc.toLowerCase())?.linkedValue || r.stayDescription;
                            return addressDesc ? <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div> : null;
                          })()}
                        </td>
                        <td className="p-1.5 border border-gray-450 leading-tight">
                          <div>{r.start} ➔ {r.end}</div>
                          <div className="text-[9px] text-[#2C6CB0] font-bold mt-1 uppercase">Rem: {getLiveRemainingDays(r.end)} Days ({r.days || 'N/A'} total days)</div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {activePrintPreview.type === 'DOSSIER' && (
              <div>
                {/* Bilingual Header */}
                <div className="text-center space-y-2 mb-8 border-b-2 border-slate-900 pb-5">
                  <h1 className="text-lg font-black uppercase text-slate-900 tracking-tight">
                    Second Division, Department of Immigration (Myeik)
                  </h1>
                  <p className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                    Foreigner Travel Dossier & Historical Movements Ledger
                  </p>
                </div>
                
                {(() => {
                  const bio = activePrintPreview.data[0];
                  if (!bio) return null;
                  return (
                    <div className="mb-6 bg-slate-50/50 p-5 rounded-2xl border-2 border-slate-200">
                      <div className="border-b pb-3 mb-4 flex justify-between items-center">
                        <span className="text-xs font-black uppercase tracking-wide text-indigo-950">
                          Personal Demographical Data & Agent Records
                        </span>
                        <span className="text-[9px] font-bold text-slate-400 font-mono uppercase">
                          Report ID: INF-{bio.passport}-{Date.now().toString().slice(-6)}
                        </span>
                      </div>

                      <div className="grid grid-cols-2 gap-y-4 gap-x-6 text-[11px] leading-relaxed">
                        {/* Row 1 */}
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Passport Number
                          </strong>
                          <span className="font-black font-mono uppercase text-indigo-950 text-sm tracking-tight">
                            {bio.passport}
                          </span>
                        </div>
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Full Name
                          </strong>
                          <span className="font-black uppercase text-slate-900">
                            {bio.fullname}
                          </span>
                        </div>

                        {/* Row 2 */}
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Nationality
                          </strong>
                          <span className="font-extrabold uppercase text-slate-800">
                            {bio.nationality || '-'}
                          </span>
                        </div>
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Gender
                          </strong>
                          <span className="font-extrabold uppercase text-slate-800">
                            {bio.gender === 'M' ? 'MALE' : bio.gender === 'F' ? 'FEMALE' : bio.gender || '-'}
                          </span>
                        </div>

                        {/* Row 3 - Agent Details */}
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Local Agent (Broker)
                          </strong>
                          <span className="font-black uppercase text-indigo-900">
                            {bio.broughtBy || '-'}
                          </span>
                        </div>
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Contact Details
                          </strong>
                          <span className="font-bold text-slate-700">
                            {bio.contactDetails || '-'}
                          </span>
                        </div>

                        {/* Row 4 - Visa and Duration */}
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Visa Classification
                          </strong>
                          <span className="font-extrabold text-slate-800">
                            {bio.visaType || '-'}
                          </span>
                        </div>
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Allowed Duration
                          </strong>
                          <span className="font-black text-slate-900 font-mono">
                            {bio.stayFrom} to {bio.stayTo} ({bio.totalDays || '-'} Days)
                          </span>
                        </div>

                        {/* Row 5 - Vehicle and Origin */}
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Transport Carrier Vehicle
                          </strong>
                          <span className="font-mono font-bold text-slate-700">
                            {bio.vehicleInfo || '-'}
                          </span>
                        </div>
                        <div className="border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Expected Entry Route Way
                          </strong>
                          <span className="font-bold text-slate-700">
                            {bio.arrivedFrom ? `From: ${bio.arrivedFrom}` : bio.departedTo ? `To: ${bio.departedTo}` : '-'}
                          </span>
                        </div>

                        {/* Row 6 - Address Details */}
                        <div className="col-span-2 border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Registered Stay Destination Address
                          </strong>
                          <span className="font-bold text-slate-900 text-xs text-indigo-950">
                            {bio.address || '-'}
                          </span>
                        </div>

                        <div className="col-span-2 border-b pb-2">
                          <strong className="text-slate-400 block text-[8px] uppercase font-black">
                            Address Detail Description
                          </strong>
                          <span className="font-semibold text-slate-700 italic">
                            {bio.stayDescription || 'No detail location description has been logged'}
                          </span>
                        </div>

                        {/* Row 7 - Intelligent Remarks / Notes (The custom-assigned remark stored!) */}
                        <div className="col-span-2 bg-amber-50/50 p-3 rounded-xl border border-amber-200">
                          <strong className="text-amber-800 block text-[8px] uppercase font-black tracking-wide">
                            Dossier Remark Details
                          </strong>
                          <span className="font-black text-amber-950 font-sans text-xs">
                            {bio.remarks || 'No special remarks registered for this individual.'}
                          </span>
                        </div>
                      </div>
                    </div>
                  );
                })()}

                <h4 className="text-[10px] font-black uppercase tracking-wider mb-2 border-b-2 border-slate-950 pb-1 mt-6">
                  Historical Log Chronicles
                </h4>
                <table className="w-full text-left text-[9px] border border-gray-450 border-collapse table-auto">
                  <thead>
                    <tr className="bg-slate-100 uppercase text-[8px] font-black">
                      <th className="p-1.5 border border-gray-450 text-center">#</th>
                      <th className="p-1.5 border border-gray-450">Timestamp</th>
                      <th className="p-1.5 border border-gray-450 text-center">Dir</th>
                      <th className="p-1.5 border border-gray-450">Visa & Duration</th>
                      <th className="p-1.5 border border-gray-450">Expected Stay Location & Description</th>
                      <th className="p-1.5 border border-gray-450">Officer Details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activePrintPreview.data.map((r: any, idx: number) => (
                      <tr key={idx} className="border-b border-gray-350 align-top">
                        <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                        <td className="p-1.5 border border-gray-450 font-mono text-center text-[8px]">{r.timestamp}</td>
                        <td className="p-1.5 border border-gray-450 text-center">
                          <span className={`px-1 py-0.5 rounded text-[8px] font-black ${r.mode === 'IN' ? 'bg-emerald-50 text-emerald-800' : 'bg-orange-50 text-orange-850'}`}>
                            {r.mode}
                          </span>
                        </td>
                        <td className="p-1.5 border border-gray-450 leading-tight">
                          <div className="font-bold">{r.visaType || 'N/A'}</div>
                          {r.stayFrom && <div className="text-[8px] text-gray-400 font-mono mt-0.5">{r.stayFrom} to {r.stayTo} ({r.totalDays} Days)</div>}
                        </td>
                        <td className="p-1.5 border border-gray-450 leading-tight">
                          <div className="font-bold text-gray-800">{r.address || 'N/A'}</div>
                          {r.stayDescription && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {r.stayDescription}</div>}
                        </td>
                        <td className="p-1.5 border border-gray-450 text-gray-500 font-bold text-[8px] leading-tight">
                          <div>{r.officialTitle || ''}</div>
                          <div className="text-gray-800">{r.officialName || 'System'}</div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Print Signatures and Confirmation Boxes */}
                <div className="grid grid-cols-3 gap-8 px-4 mt-20 no-print-inside">
                  <div className="text-center space-y-12">
                    <p className="text-[10px] font-black uppercase text-slate-500">Prepared By</p>
                    <div className="border-t border-slate-400 pt-2 text-[10px] font-bold text-slate-800">
                      Dossier Retrieval Officer
                    </div>
                    <p className="text-[8px] text-slate-400">Signature: _______________________</p>
                  </div>
                  
                  <div className="text-center space-y-12">
                    <p className="text-[10px] font-black uppercase text-slate-500">Verified By</p>
                    <div className="border-t border-slate-400 pt-2 text-[10px] font-bold text-slate-800">
                      Duty Supervisor / Commander
                    </div>
                    <p className="text-[8px] text-slate-400">Signature: _______________________</p>
                  </div>

                  <div className="text-center bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 space-y-10">
                    <p className="text-[9px] font-black uppercase text-indigo-900 tracking-wider">
                      Confirmed & Certified Box
                    </p>
                    <div className="space-y-4 text-left text-[9px] font-bold text-slate-600">
                      <div>Stamp: __________________</div>
                      <div>Signature: ________________</div>
                      <div>Date: {new Date().toLocaleDateString()}</div>
                    </div>
                  </div>
                </div>

              </div>
            )}

            {(activePrintPreview.type === 'CO_LTD_PENDING' || activePrintPreview.type === 'OTHERS_PENDING' || activePrintPreview.type === 'CHECKED_HISTORY') && (
              <div>
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-md font-black uppercase text-gray-900">{activePrintPreview.title}</h1>
                  <p className="text-[9px] text-gray-400 font-mono uppercase tracking-wide">
                    PREVIEW DATE: {new Date().toLocaleString()} • STAY REGISTER UNIT
                  </p>
                </div>
                
                <table className="w-full text-left text-[10px] border border-gray-450 border-collapse table-auto">
                  <thead>
                    <tr className="bg-gray-100 uppercase text-[8px] font-black">
                      <th className="p-1.5 border border-gray-450 text-center">#</th>
                      <th className="p-1.5 border border-gray-450">Passport & Guest Name</th>
                      <th className="p-1.5 border border-gray-450">Nationality</th>
                      {activePrintPreview.type === 'CHECKED_HISTORY' ? (
                        <>
                          <th className="p-1.5 border border-gray-450">Confirmed Stay Address & Detail Description</th>
                          <th className="p-1.5 border border-gray-450 text-center">Status</th>
                          <th className="p-1.5 border border-gray-450">Officer details</th>
                        </>
                      ) : (
                        <>
                          <th className="p-1.5 border border-gray-450">Expected Stay Address & Detail Description</th>
                          <th className="p-1.5 border border-gray-450">Stay Schedule</th>
                          <th className="p-1.5 border border-gray-450 min-w-[120px]">Verification Signature Box</th>
                        </>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {activePrintPreview.data.map((h: any, idx: number) => {
                      const isHist = h.status !== undefined;
                      const pNo = h.passport || h.p || '';
                      const name = h.fullname || h.n || '';
                      const nat = h.nationality || h.nat || '';
                      const address = h.confirmedAddress || h.loc || '';
                      const addressDesc = h.confirmedStayDescription || (masterData.find(m => m.type === 'Stay' && m.name.toLowerCase() === (h.loc || '').toLowerCase())?.linkedValue || h.stayDescription);

                      return (
                        <tr key={idx} className="border-b border-gray-350 align-top">
                          <td className="p-1.5 border border-gray-450 text-center font-bold">{idx + 1}</td>
                          <td className="p-1.5 border border-gray-450 leading-tight">
                            <div className="font-extrabold font-mono uppercase text-[#2C6CB0]">{pNo}</div>
                            <div className="font-bold text-gray-800 uppercase mt-0.5">{name}</div>
                          </td>
                          <td className="p-1.5 border border-gray-450 font-black uppercase text-gray-700">{nat}</td>
                          
                          {isHist ? (
                            <>
                              <td className="p-1.5 border border-gray-450 leading-tight">
                                <div className="font-extrabold text-gray-900">{address}</div>
                                {addressDesc && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                              </td>
                              <td className="p-1.5 border border-gray-450 text-center">
                                <span className="px-1.5 py-0.5 rounded text-[8px] bg-emerald-50 text-emerald-800 font-extrabold border border-emerald-100 uppercase">{h.status}</span>
                                <div className="text-[8px] text-slate-400 font-mono mt-1">{h.checkDate}</div>
                              </td>
                              <td className="p-1.5 border border-gray-450 leading-tight text-[8px] text-slate-500 font-bold">
                                <div>{h.officerName}</div>
                                <div className="text-gray-400 uppercase text-[7px] font-black">{h.officerTitle}</div>
                              </td>
                            </>
                          ) : (
                            <>
                              <td className="p-1.5 border border-gray-450 leading-tight">
                                <div className="font-extrabold text-gray-900">{address}</div>
                                {addressDesc && <div className="text-[8px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                              </td>
                              <td className="p-1.5 border border-gray-450 leading-tight font-mono text-[9px]">
                                <div>{h.start} to {h.end}</div>
                                <div className="text-[8px] text-slate-400 mt-1 uppercase">Rem: {getLiveRemainingDays(h.end)} Days</div>
                              </td>
                              <td className="p-1.5 border border-gray-450 text-center bg-slate-50/50 min-w-[130px]">
                                <div className="h-14 flex items-center justify-center border border-dashed border-slate-300 rounded mb-1 bg-white shadow-inner">
                                  <span className="text-[6.5px] text-slate-400/85 uppercase tracking-widest font-black font-mono">STAMP & SIGN BOX</span>
                                </div>
                                <span className="text-[7px] text-slate-500 uppercase tracking-wider font-extrabold">Verify Inspector Sign</span>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {activePrintPreview.type === 'SIA_NATIONALITY_REPORT' && (
              <div className="space-y-4">
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-sm font-black uppercase tracking-wider text-slate-800">Second Division, Department of Immigration (Myeik)</h1>
                  <h2 className="text-md font-extrabold uppercase tracking-tight text-indigo-950">STATISTICAL INVESTIGATION ANALYSIS (SIA)</h2>
                  <p className="text-[10px] font-bold text-gray-600 uppercase tracking-widest">{activePrintPreview.title}</p>
                </div>
                <div className="border border-black rounded-xl overflow-hidden">
                  <table className="w-full text-left text-[9px] border-collapse">
                    <thead>
                      <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-black">
                        <th className="p-2 border-r border-black text-center w-12">No.</th>
                        <th className="p-2 border-r border-black">Nationality</th>
                        <th className="p-2 border-r border-black text-center">Inbound (IN)</th>
                        <th className="p-2 text-center">Outbound (OUT)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((r: any, idx: number) => (
                        <tr key={idx} className="border-b border-black">
                          <td className="p-2 border-r border-black text-center font-bold">{idx + 1}</td>
                          <td className="p-2 border-r border-black font-extrabold text-slate-800 uppercase">{r.country}</td>
                          <td className="p-2 border-r border-black text-center font-bold text-emerald-700">{r.in}</td>
                          <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {activePrintPreview.type === 'SIA_VISA_REPORT' && (
              <div className="space-y-4">
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-sm font-black uppercase tracking-wider text-slate-800">Second Division, Department of Immigration (Myeik)</h1>
                  <h2 className="text-md font-extrabold uppercase tracking-tight text-indigo-950">STATISTICAL INVESTIGATION ANALYSIS (SIA)</h2>
                  <p className="text-[10px] font-bold text-gray-600 uppercase tracking-widest">{activePrintPreview.title}</p>
                </div>
                <div className="border border-black rounded-xl overflow-hidden">
                  <table className="w-full text-left text-[9px] border-collapse">
                    <thead>
                      <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-black">
                        <th className="p-2 border-r border-black text-center w-12">No.</th>
                        <th className="p-2 border-r border-black">Visa Type</th>
                        <th className="p-2 border-r border-black text-center">Inbound (IN)</th>
                        <th className="p-2 text-center">Outbound (OUT)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {activePrintPreview.data.map((r: any, idx: number) => (
                        <tr key={idx} className="border-b border-black">
                          <td className="p-2 border-r border-black text-center font-bold">{idx + 1}</td>
                          <td className="p-2 border-r border-black font-extrabold text-slate-800 uppercase">{r.visaType}</td>
                          <td className="p-2 border-r border-black text-center font-bold text-emerald-700">{r.in}</td>
                          <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {activePrintPreview.type === 'SIA_FULL_REPORT' && (
              <div className="space-y-8">
                <div className="text-center space-y-2 mb-8 border-b-2 border-black pb-4">
                  <h1 className="text-sm font-black uppercase tracking-wider text-slate-800">Second Division, Department of Immigration (Myeik)</h1>
                  <h2 className="text-md font-extrabold uppercase tracking-tight text-indigo-950">STATISTICAL INVESTIGATION ANALYSIS (SIA)</h2>
                  <p className="text-[10px] font-bold text-gray-600 uppercase tracking-widest">{activePrintPreview.title}</p>
                </div>
                
                <div>
                  <h3 className="text-xs font-black uppercase text-indigo-950 mb-2 border-b pb-1">1. Nationality Inflow & Outflow</h3>
                  <div className="border border-black rounded-xl overflow-hidden">
                    <table className="w-full text-left text-[9px] border-collapse">
                      <thead>
                        <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-black">
                          <th className="p-2 border-r border-black text-center w-12">No.</th>
                          <th className="p-2 border-r border-black">Nationality</th>
                          <th className="p-2 border-r border-black text-center">Inbound (IN)</th>
                          <th className="p-2 text-center">Outbound (OUT)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activePrintPreview.data.nationalityList.map((r: any, idx: number) => (
                          <tr key={idx} className="border-b border-black">
                            <td className="p-2 border-r border-black text-center font-bold">{idx + 1}</td>
                            <td className="p-2 border-r border-black font-extrabold text-slate-800 uppercase">{r.country}</td>
                            <td className="p-2 border-r border-black text-center font-bold text-emerald-700">{r.in}</td>
                            <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div>
                  <h3 className="text-xs font-black uppercase text-indigo-950 mb-2 border-b pb-1">2. Visa Type Distribution</h3>
                  <div className="border border-black rounded-xl overflow-hidden">
                    <table className="w-full text-left text-[9px] border-collapse">
                      <thead>
                        <tr className="bg-slate-50 uppercase text-[8px] font-black border-b border-black">
                          <th className="p-2 border-r border-black text-center w-12">No.</th>
                          <th className="p-2 border-r border-black">Visa Type</th>
                          <th className="p-2 border-r border-black text-center">Inbound (IN)</th>
                          <th className="p-2 text-center">Outbound (OUT)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {activePrintPreview.data.visaList.map((r: any, idx: number) => (
                          <tr key={idx} className="border-b border-black">
                            <td className="p-2 border-r border-black text-center font-bold">{idx + 1}</td>
                            <td className="p-2 border-r border-black font-extrabold text-slate-800 uppercase">{r.visaType}</td>
                            <td className="p-2 border-r border-black text-center font-bold text-emerald-700">{r.in}</td>
                            <td className="p-2 text-center font-bold text-orange-700">{r.out}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {activePrintPreview.type === 'TELEGRAPH_REPORT' && (
              <div className="font-pyidaungsu text-black relative print-content w-full text-[13px] space-y-6">
                {/* Letter Header */}
                <div className="flex justify-between items-start text-[13px] font-bold border-b-2 border-black pb-4 mb-4">
                  <div className="space-y-1">
                    <div>(ပ) {activePrintPreview.data.chief1}</div>
                    <div>(လ) {activePrintPreview.data.chief2}</div>
                    <div>(တ) {activePrintPreview.data.chief3}</div>
                  </div>
                  <div className="text-center space-y-0.5">
                    <div>ရက်စွဲ / အချိန်</div>
                    <div className="text-slate-400 font-light leading-none -mt-1">—————-</div>
                    <div className="font-bold">
                      {toBurmeseShortDateTime(activePrintPreview.data.selectedDate, activePrintPreview.data.reportTime)}
                    </div>
                  </div>
                </div>

                <div className="text-center my-4">
                  <span className="text-[13px] font-bold block text-center font-pyidaungsu" style={{ color: '#000000', letterSpacing: 'normal' }}>
                    မှတ်ပုံတင်ကြေးနန်းစာပုံစံအစား
                  </span>
                </div>

                {/* Opening Line above title */}
                <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

                {/* Letter Info Line */}
                <div className="flex items-center gap-12 text-[13px] font-bold py-1 text-black">
                  <div>{activePrintPreview.data.letterNo}</div>
                  <div className="font-black text-[13px] uppercase">{activePrintPreview.data.telegraphTitle}</div>
                </div>

                {/* Paragraph 1 */}
                <p className="text-[13px] leading-relaxed font-medium whitespace-pre-line my-4">
                  {(activePrintPreview.data.customIntroText || '')
                    .replace('[DATE]', activePrintPreview.data.dateBurmeseShort || toBurmeseSlashDate(activePrintPreview.data.selectedDate))
                    .replace('[FLIGHT_COUNT]', toBurmeseDigits((activePrintPreview.data.rows || []).length))
                  }
                </p>

                {/* High Fidelity Burmese Table */}
                <table className="w-full text-center text-[13px] border-collapse border border-black min-w-[650px] font-pyidaungsu">
                  <thead>
                    <tr className="bg-slate-50 uppercase font-black border-b border-black">
                      <th rowSpan={3} className="p-1 border border-black text-center align-middle w-10">စဉ်</th>
                      <th rowSpan={3} className="p-1 border border-black text-center align-middle">လေကြောင်းလိုင်းအမည်</th>
                      <th rowSpan={3} className="p-1 border border-black text-center align-middle w-16">ဝင်/ထွက်ချိန်</th>
                      <th colSpan={6} className="p-1 border border-black text-center font-black">မြန်မာနိုင်ငံသား</th>
                      <th colSpan={6} className="p-1 border border-black text-center font-black">နိုင်ငံခြားသား</th>
                      <th rowSpan={3} className="p-1 border border-black text-center align-middle">မှတ်ချက်</th>
                    </tr>
                    <tr className="bg-slate-50 border-b border-black">
                      <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အဝင်</th>
                      <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အထွက်</th>
                      <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အဝင်</th>
                      <th colSpan={3} className="p-1 border border-black text-center text-[13px] font-bold">အထွက်</th>
                    </tr>
                    <tr className="bg-slate-50 border-b border-black text-[13px] font-bold text-slate-700">
                      {/* MMR In */}
                      <th className="p-1 border border-black w-8">က</th>
                      <th className="p-1 border border-black w-8">မ</th>
                      <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                      {/* MMR Out */}
                      <th className="p-1 border border-black w-8">က</th>
                      <th className="p-1 border border-black w-8">မ</th>
                      <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                      {/* FRN In */}
                      <th className="p-1 border border-black w-8">က</th>
                      <th className="p-1 border border-black w-8">မ</th>
                      <th className="p-1 border border-black w-8 bg-slate-100/50">ပ</th>
                      {/* FRN Out */}
                      <th className="p-1 border border-black w-8">က</th>
                      <th className="p-1 border-black w-8">မ</th>
                      <th className="p-1 border-black w-8 bg-slate-100/50">ပ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activePrintPreview.data.rows && activePrintPreview.data.rows.map((row: any, idx: number) => {
                      const myanmarInTotal = row.myanmarInM + row.myanmarInF;
                      const myanmarOutTotal = row.myanmarOutM + row.myanmarOutF;
                      const foreignerInTotal = row.foreignerInM + row.foreignerInF;
                      const foreignerOutTotal = row.foreignerOutM + row.foreignerOutF;

                      const fNum = (val: number | string) => activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(val) : val;

                      return (
                        <tr key={row.id} className="border-b border-black">
                          <td className="p-1.5 border border-black text-center font-bold">{fNum(idx + 1)}</td>
                          <td className="p-1.5 border border-black font-extrabold uppercase text-left">{row.flightName}</td>
                          <td className="p-1.5 border border-black text-center font-mono">{fNum(row.time)}</td>
                          
                          {/* MMR In */}
                          <td className="p-1.5 border border-black text-center">{fNum(row.myanmarInM)}</td>
                          <td className="p-1.5 border border-black text-center">{fNum(row.myanmarInF)}</td>
                          <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(myanmarInTotal)}</td>
                          
                          {/* MMR Out */}
                          <td className="p-1.5 border border-black text-center">{fNum(row.myanmarOutM)}</td>
                          <td className="p-1.5 border border-black text-center">{fNum(row.myanmarOutF)}</td>
                          <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(myanmarOutTotal)}</td>
                          
                          {/* FRN In */}
                          <td className="p-1.5 border border-black text-center">{fNum(row.foreignerInM)}</td>
                          <td className="p-1.5 border border-black text-center">{fNum(row.foreignerInF)}</td>
                          <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(foreignerInTotal)}</td>
                          
                          {/* FRN Out */}
                          <td className="p-1.5 border border-black text-center">{fNum(row.foreignerOutM)}</td>
                          <td className="p-1.5 border border-black text-center">{fNum(row.foreignerOutF)}</td>
                          <td className="p-1.5 border border-black text-center font-bold bg-slate-50/40">{fNum(foreignerOutTotal)}</td>
                          
                          <td className="p-1.5 border border-black text-left text-[13px] font-bold">{row.remarks}</td>
                        </tr>
                      );
                    })}

                    {/* Totals Row */}
                    <tr className="bg-slate-100/70 font-bold border-b border-black text-[13px]">
                      <td colSpan={3} className="p-2 border border-black text-center font-black">Total</td>
                      
                      {/* MMR In Total */}
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM, 0)}</td>
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInF, 0)}</td>
                      <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM + b.myanmarInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarInM + b.myanmarInF, 0)}</td>
                      
                      {/* MMR Out Total */}
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM, 0)}</td>
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutF, 0)}</td>
                      <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM + b.myanmarOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.myanmarOutM + b.myanmarOutF, 0)}</td>
                      
                      {/* FRN In Total */}
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM, 0)}</td>
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInF, 0)}</td>
                      <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM + b.foreignerInF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerInM + b.foreignerInF, 0)}</td>
                      
                      {/* FRN Out Total */}
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM, 0)}</td>
                      <td className="p-2 border border-black text-center">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutF, 0)}</td>
                      <td className="p-2 border border-black text-center font-black bg-slate-200/40">{activePrintPreview.data.useBurmeseDigits ? toBurmeseDigits(activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM + b.foreignerOutF, 0)) : activePrintPreview.data.rows.reduce((a: any, b: any) => a + b.foreignerOutM + b.foreignerOutF, 0)}</td>
                      
                      <td className="p-2 border border-black"></td>
                    </tr>
                  </tbody>
                </table>

                {/* Paragraph 2 - (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... / */}
                <div className="text-[13px] font-normal mt-4">
                  (နှစ်) သိရှိနိုင်ပါရန်တင်ပြအပ်။... /
                </div>

                {/* ပိတ်မျဉ်း အပြည့်တစ်ကြောင်း below သိရှိနိုင်ပါရန်တင်ပြအပ် */}
                <hr className="border-t-2 border-black my-3 w-full block" style={{ borderTop: '1.5px solid #000000', borderBottom: 'none', borderLeft: 'none', borderRight: 'none', width: '100%', margin: '10px 0' }} />

                {/* Footer Text */}
                <div className="flex justify-between items-start text-[13px] font-bold mt-2 leading-relaxed">
                  <div className="space-y-1">
                    <div className="font-medium">သင့်ရာနည်းဖြင့်</div>
                  </div>
                  <div className="text-center space-y-0.5 pr-8">
                    <div>အဆင့်အတန်း</div>
                    <div className="text-slate-400 font-light leading-none -mt-1">—————</div>
                    <div className="font-black text-rose-700">{activePrintPreview.data.urgentLevel}</div>
                  </div>
                </div>

                {/* Signature block */}
                <div className="flex justify-end mt-12">
                  <div className="text-center space-y-1 w-96 flex flex-col items-center">
                    <div className="h-10" />
                    <div className="border-t border-black pt-1 w-full text-center flex flex-col items-center font-pyidaungsu">
                      <div className="text-[13px] font-extrabold text-black">
                        {activePrintPreview.data.officerTitle}
                      </div>
                      <div className="text-[13px] font-bold text-black leading-tight">
                        {activePrintPreview.data.officerName}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Document Lightweight Verification Badge instead of heavy QRs */}
            {activePrintPreview && (
              <div className="mt-10 p-4 border border-dashed border-black rounded-lg flex flex-col md:flex-row justify-between items-center gap-6 no-print-inside">
                <div className="flex items-start gap-4">
                  <div className="space-y-1">
                    <h4 className="text-[10px] font-black uppercase tracking-wide font-sans">Official Verification Log</h4>
                    <p className="text-[8px] text-gray-500 leading-relaxed max-w-[340px] font-sans">
                      Verified and compiled official system record from the Myeik Division Statistical Investigation Analysis database. Contains zero QR tags as requested.
                    </p>
                    <div className="text-[8px] font-bold pt-1 font-sans">
                      <div>Operator: {currentUser ? `${currentUser.title} ${currentUser.name}` : `${localStorage.getItem('lastCheckedOfficerTitle') || 'Officer'} ${localStorage.getItem('lastCheckedOfficerName') || 'Myeik Duty Officer'}`}</div>
                      <div>Timestamp: {new Date().toLocaleString()}</div>
                    </div>
                  </div>
                </div>
              </div>
            )}
            
            <div className="mt-8 pt-4 border-t border-dashed border-gray-400 text-center text-[9px] text-gray-400 leading-normal">
              <p>Myeik Immigration Enforcement Unit • Statistical Investigation Analysis</p>
              <p>CONFIDENTIAL • DEPARTMENT OF IMMIGRATION & REGISTRATION RECORD</p>
            </div>
      </div>
    )}

    {/* Dedicated Hidden Printable Container for Checking Verification Lists in Print Media */}
    {checkingPrintData && (
      <div className="hidden print:block bg-white p-8 text-black min-h-screen w-full text-xs font-sans print-content leading-relaxed">
        <div className="border-b-4 border-gray-900 pb-4 mb-6 text-center">
          <h1 className="text-lg font-black uppercase text-gray-900">{checkingPrintData.title}</h1>
          <p className="text-[9px] text-gray-400 font-mono uppercase tracking-wide">
            GENERATED COPY TIMESTAMP: {new Date().toLocaleString()} • SYSTEM REGISTER COPY
          </p>
        </div>

        <table className="w-full text-left text-xs border border-black border-collapse">
          <thead>
            <tr className="bg-gray-100 uppercase text-[8px] font-black">
              <th className="p-2 border border-black text-center">#</th>
              <th className="p-2 border border-black">Passport & Visitor Name</th>
              <th className="p-2 border border-black">Nationality</th>
              {checkingPrintData.type === 'CHECKED_HISTORY' ? (
                <>
                  <th className="p-2 border border-black">Confirmed Stay Address Location & Detail Address/Description</th>
                  <th className="p-2 border border-black text-center">Verified Status</th>
                  <th className="p-2 border border-black">Officer Details</th>
                </>
              ) : (
                <>
                  <th className="p-2 border border-black">Expected Stay Address & Detail Address/Description</th>
                  <th className="p-2 border border-black">Stay Schedule</th>
                  <th className="p-2 border border-black text-center">Verification Stamp / Sign</th>
                </>
              )}
            </tr>
          </thead>
          <tbody>
            {checkingPrintData.data.map((h: any, idx: number) => {
              const isHist = checkingPrintData.type === 'CHECKED_HISTORY';
              const pNo = isHist ? h.passport : h.p;
              const name = isHist ? h.fullname : h.n;
              const nat = isHist ? h.nationality : h.nat;
              const address = isHist ? h.confirmedAddress : h.loc;
              const addressDesc = isHist 
                ? h.confirmedStayDescription 
                : (masterData.find(m => m.type === 'Stay' && m.name.toLowerCase() === h.loc.toLowerCase())?.linkedValue || h.stayDescription);

              return (
                <tr key={idx} className="align-top border-b border-black">
                  <td className="p-2 border border-black text-center font-bold">{idx + 1}</td>
                  <td className="p-2 border border-black leading-tight">
                    <div className="font-extrabold font-mono uppercase">{pNo}</div>
                    <div className="font-bold text-gray-800 uppercase mt-0.5">{name}</div>
                  </td>
                  <td className="p-2 border border-black font-black uppercase text-gray-700">{nat}</td>
                  
                  {isHist ? (
                    <>
                      <td className="p-2 border border-black leading-tight">
                        <div className="font-extrabold text-gray-900">{address}</div>
                        {addressDesc && <div className="text-[9px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                      </td>
                      <td className="p-2 border border-black text-center font-black">
                        <div className="text-emerald-800">{h.status}</div>
                        <div className="text-[8px] text-gray-400 font-mono mt-1">{h.checkDate}</div>
                      </td>
                      <td className="p-2 border border-black font-bold leading-tight">
                        <div>{h.officerName}</div>
                        <div className="text-[8px] text-gray-400 uppercase font-black">{h.officerTitle}</div>
                      </td>
                    </>
                  ) : (
                    <>
                      <td className="p-2 border border-black leading-tight">
                        <div className="font-extrabold text-gray-950">{address}</div>
                        {addressDesc && <div className="text-[9px] text-gray-500 italic mt-0.5">🏡 Detail: {addressDesc}</div>}
                      </td>
                      <td className="p-2 border border-black leading-tight font-mono">
                        <div>{h.start} to {h.end}</div>
                        <div className="text-[9px] text-indigo-750 font-black mt-1 uppercase">Rem: {getLiveRemainingDays(h.end)} Days</div>
                      </td>
                      <td className="p-2 border border-black text-center bg-gray-50/30 min-w-[140px]">
                        <div className="h-14 flex items-center justify-center border border-dashed border-gray-400 rounded-md mb-1 bg-white">
                          <span className="text-[7px] text-gray-400 uppercase tracking-widest font-black font-mono">PHYSICAL STAMP HERE</span>
                        </div>
                        <span className="text-[7.5px] text-gray-600 uppercase font-bold">Immig. Inspector Verified</span>
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Footer for checklist */}
        <div className="mt-12 pt-8 border-t border-dashed border-black text-center text-[10px] text-gray-400 leading-normal">
          <p>Myeik Immigration Department • Second Division Stay Verification Dossier</p>
          <p className="mt-1 font-bold">CONFIDENTIAL • FOR OFFICIAL USE ONLY</p>
        </div>
      </div>
    )}

    {/* Initial Startup & Data Loading Waiting Screen (Both Offline and Online Modes) */}
    <AnimatePresence>
      {(!isDBCardLoaded || isInitialSyncing) && (
        <motion.div
          initial={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          className="fixed inset-0 z-[500] bg-slate-950 flex flex-col items-center justify-center p-6 text-white text-center no-print"
        >
          <div className="w-20 h-20 bg-[#2C6CB0]/25 text-blue-400 rounded-3xl flex items-center justify-center border border-blue-500/30 mb-5 shadow-2xl shadow-blue-500/20 animate-pulse">
            <ShieldCheck size={40} className="text-blue-400" />
          </div>
          
          <h2 className="text-2xl font-black tracking-tight mb-1 text-slate-100 uppercase">
            IMMIGRATION LOG SYSTEM V10
          </h2>
          <p className="text-xs text-slate-400 font-semibold max-w-md mb-4 leading-relaxed">
            Myeik Division • Second Division Logistics & Stay Tracking System
          </p>

          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-bold mb-5 border ${isOnline ? 'bg-blue-950/60 border-blue-500/40 text-blue-300' : 'bg-amber-950/60 border-amber-500/40 text-amber-300'}">
            <span className="w-2 h-2 rounded-full ${isOnline ? 'bg-blue-400 animate-ping' : 'bg-amber-400'}"></span>
            <span>{isOnline ? "🌐 Online Mode • Cloud Synchronizing" : "⚡ Offline Mode • Local Database Active"}</span>
          </div>

          <div className="w-72 h-2.5 bg-slate-800/80 rounded-full overflow-hidden relative mb-4 border border-slate-700/50">
            <div 
              className="absolute inset-y-0 bg-gradient-to-r from-blue-500 via-indigo-500 to-emerald-400 rounded-full transition-all duration-300 ${downloadProgress?.percent ? '' : 'animate-pulse w-full'}"
              style={{ width: downloadProgress?.percent ? `${downloadProgress.percent}%` : "100%" }}
            />
          </div>

          <div className="flex flex-col items-center gap-1.5 mb-6">
            <p className="text-xs text-blue-300 font-bold">
              {downloadProgress?.currentCol 
                ? `Downloading: ${downloadProgress.currentCol} (${downloadProgress.percent}%)`
                : !isDBCardLoaded 
                  ? "စက်တွင်း Local Database အချက်အလက်များ တင်ယူနေပါသည်..."
                  : "Cloud Server ဒေတာများနှင့် စစ်ဆေးချိတ်ဆက်နေပါသည်..."
              }
            </p>
            <p className="text-[11px] font-mono text-slate-400 bg-slate-900/80 px-3 py-1 rounded-lg border border-slate-800">
              စောင့်ဆိုင်းချိန် (Waiting Time): <span className="font-bold text-amber-400">{waitingElapsedSeconds}</span> စက္ကန့်
            </p>
          </div>

          {/* Quick Direct Launch Button if loading takes longer than 2 seconds */}
          {waitingElapsedSeconds >= 2 && (
            <motion.button
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              onClick={() => {
                setIsDBCardLoaded(true);
                setIsInitialSyncing(false);
              }}
              className="px-5 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-bold text-xs rounded-xl shadow-lg shadow-blue-500/20 border border-blue-400/30 transition-all cursor-pointer mb-4 active:scale-95"
            >
              တိုက်ရိုက်စတင်ရန် (Open App Now) →
            </motion.button>
          )}

          <p className="text-[10px] text-slate-500 font-mono font-bold uppercase tracking-widest">
            OFFLINE-FIRST LOCAL STORAGE • AUTOMATIC SYNC
          </p>
        </motion.div>
      )}
    </AnimatePresence>

    {/* Cloud Sync Role Authentication Modal Overlay */}
    {showCloudLoginModal && (
      <div className="fixed inset-0 z-[400] bg-slate-900/95 backdrop-blur-md flex items-center justify-center p-4 overflow-y-auto no-print">
        <motion.div
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="bg-white rounded-3xl shadow-2xl max-w-md w-full p-6 sm:p-8 border border-slate-200 text-slate-900 relative my-8"
        >
          {/* Close Button - Only available if already logged in with a Cloud user */}
          {cloudAuthUser && cloudAuthUser.username !== 'LOCAL_OFFLINE' && (
            <button
              onClick={dismissCloudLogin}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 p-2 rounded-full hover:bg-slate-100 transition-colors"
              title="Close"
            >
              <X size={20} />
            </button>
          )}

          <div className="text-center mb-6">
            <div className="w-16 h-16 bg-[#2C6CB0] text-white rounded-2xl mx-auto flex items-center justify-center shadow-lg shadow-blue-500/30 mb-3">
              <Shield size={32} />
            </div>
            <h2 className="text-2xl font-black text-slate-900 uppercase tracking-tight">Cloud Sync Login</h2>
            <p className="text-xs text-slate-500 mt-1 font-semibold">
              Department of Immigration Log System • Role Authentication
            </p>
          </div>

          {loginAuthError && (
            <div className="mb-6 p-3 bg-rose-50 border border-rose-200 text-rose-700 text-xs font-bold rounded-xl flex items-center gap-2">
              <AlertCircle size={16} className="shrink-0 text-rose-500" />
              <span>{loginAuthError}</span>
            </div>
          )}

          {/* Login Form */}
          <div className="space-y-4 mb-6">
            <div>
              <label className="block text-xs font-bold uppercase text-slate-700 mb-1">
                Username
              </label>
              <div className="relative">
                <input
                  type="text"
                  value={loginUsernameInput}
                  onChange={e => setLoginUsernameInput(e.target.value)}
                  placeholder="SUPERADMIN သို့မဟုတ် EDITOR"
                  className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#2C6CB0]"
                />
                <UserCheck size={18} className="absolute left-3 top-3 text-slate-400" />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase text-slate-700 mb-1">
                Password
              </label>
              <div className="relative">
                <input
                  type="password"
                  value={loginPasswordInput}
                  onChange={e => setLoginPasswordInput(e.target.value)}
                  placeholder="Password ရိုက်ထည့်ပါ"
                  className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#2C6CB0]"
                  onKeyDown={e => e.key === 'Enter' && handleCloudLogin()}
                />
                <Lock size={18} className="absolute left-3 top-3 text-slate-400" />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase text-slate-700 mb-1">
                Device Name / Identifier (Optional)
              </label>
              <div className="relative">
                <input
                  type="text"
                  value={loginDeviceInput}
                  onChange={e => setLoginDeviceInput(e.target.value)}
                  placeholder="e.g. iPad Terminal 01"
                  className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#2C6CB0]"
                />
                <Smartphone size={18} className="absolute left-3 top-3 text-slate-400" />
              </div>
            </div>

            <button
              onClick={() => handleCloudLogin()}
              className="w-full py-3.5 bg-[#2C6CB0] hover:bg-blue-800 text-white font-extrabold uppercase tracking-wide rounded-xl shadow-lg shadow-blue-600/30 transition-all flex items-center justify-center gap-2 text-sm cursor-pointer"
            >
              <LogIn size={18} /> ဝင်ရောက်မည် (Sign In)
            </button>
          </div>

          <div className="border-t border-slate-200 pt-4 text-center">
            <p className="text-xs text-slate-500 font-semibold leading-relaxed">
              🔒 Cloud Account တစ်ခုဖြင့် မဖြစ်မနေ ဝင်ရောက်ရန် လိုအပ်ပါသည်။ အကောင့်ဝင်ပြီးပါက အင်တာနက် သို့မဟုတ် Quota ပြတ်တောက်ချိန်တွင်လည်း Local Data ကို အဆင်ပြေစွာ ဆက်လက် အသုံးပြုနိုင်မည်ဖြစ်ပါသည်။
            </p>
          </div>
        </motion.div>
      </div>
    )}

    {/* Counter Check Modal */}
    <CounterCheckModal
      isOpen={isCounterCheckOpen}
      onClose={() => setIsCounterCheckOpen(false)}
      records={records}
      setRecords={setRecords}
      showToast={showToast}
      onEditRecord={startEdit}
    />

    {/* Cloud Server & Devices Sync Hub Modal */}
    <CloudSyncStatusModal
      isOpen={isSyncModalOpen}
      onClose={() => setIsSyncModalOpen(false)}
      isOnline={isOnline}
      isAutoSyncing={isAutoSyncing}
      isCloudSynced={isCloudSynced}
      lastSyncTime={lastSyncTime}
      isQuotaExhausted={isQuotaExhausted}
      cloudAuthUser={cloudAuthUser}
      deviceSessions={deviceSessions}
      records={records}
      masterDataCount={masterData.length}
      tempRecordsCount={tempRecords.length}
      checkingHistoryCount={checkingHistory.length}
      onManualSync={handleManualSync}
      onFetchFromCloud={handleFetchDataFromCloud}
      onResetQuota={resetQuotaState}
    />

    </>
  );
}
