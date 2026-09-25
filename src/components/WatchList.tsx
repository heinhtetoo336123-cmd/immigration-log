import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  ShieldAlert,
  Search,
  Plus,
  Trash2,
  Edit,
  Printer,
  FileSpreadsheet,
  Download,
  Upload,
  CheckCircle2,
  AlertTriangle,
  FileText,
  User,
  Users,
  Calendar,
  Building,
  MapPin,
  RefreshCw,
  Eye,
  X,
  Clock,
  ShieldCheck,
  Check,
  ChevronDown,
  Copy,
  Info,
  ExternalLink,
  ChevronRight,
  Filter,
  UserCheck,
  ArrowLeft,
  Plane
} from 'lucide-react';
import { ImmRecord, MasterItem, MovementData, Tab, WatchListPerson, WatchListRecord } from '../types';
import { logActivity } from '../utils/activityLogger';
import * as XLSX from 'xlsx';
import { miniDB } from '../db';
import { saveCollectionToFirestore, subscribeToFirestoreCollection, fetchCollectionFromFirestore, setSyncedHash } from '../lib/firebase';

interface WatchListProps {
  records: ImmRecord[];
  tempRecords?: ImmRecord[];
  movementMap?: Record<string, MovementData>;
  checkingHistory?: any[];
  masterData?: MasterItem[];
  currentUser?: { title?: string; name?: string } | null;
  showToast: (msg: string) => void;
  setActivePrintPreview?: (preview: any) => void;
  isViewer?: boolean;
  watchList?: WatchListRecord[];
  setWatchList?: React.Dispatch<React.SetStateAction<WatchListRecord[]>>;
  handleManualSync?: () => void;
  handleFetchDataFromCloud?: (force?: boolean) => void;
}

const STORAGE_KEY = 'imm_watchlist_records_v1';

export const toBurmeseDigits = (numStr: string | number | undefined | null) => {
  if (numStr === undefined || numStr === null) return '';
  const burmeseDigits = ['၀', '၁', '၂', '၃', '၄', '၅', '၆', '၇', '၈', '၉'];
  return String(numStr).replace(/[0-9]/g, (match) => burmeseDigits[parseInt(match)]);
};

export const toBurmeseFormattedDate = (dateStr: string | undefined | null, withParens: boolean = true) => {
  if (!dateStr) return '';
  try {
    const s = String(dateStr).trim();
    const clean = s.replace(/[()]/g, '');
    const justDate = clean.split(', ')[0].split('T')[0].trim();
    let day = '';
    let month = '';
    let year = '';

    if (justDate.includes('-')) {
      const parts = justDate.split('-');
      if (parts.length === 3) {
        if (parts[0].length === 4) {
          // YYYY-MM-DD
          year = parts[0];
          month = parts[1].padStart(2, '0');
          day = parts[2].padStart(2, '0');
        } else {
          // DD-MM-YYYY
          day = parts[0].padStart(2, '0');
          month = parts[1].padStart(2, '0');
          year = parts[2];
        }
      }
    } else if (justDate.includes('/')) {
      const parts = justDate.split('/');
      if (parts.length === 3) {
        if (parts[2].length === 4) {
          // DD/MM/YYYY
          day = parts[0].padStart(2, '0');
          month = parts[1].padStart(2, '0');
          year = parts[2];
        } else if (parts[0].length === 4) {
          // YYYY/MM/DD
          year = parts[0];
          month = parts[1].padStart(2, '0');
          day = parts[2].padStart(2, '0');
        }
      }
    }

    if (day && month && year) {
      const formatted = `${toBurmeseDigits(day)}-${toBurmeseDigits(month)}-${toBurmeseDigits(year)}`;
      return withParens ? `(${formatted})` : formatted;
    }
  } catch (e) {
    console.error('Date parsing error', e);
  }
  const fallback = toBurmeseDigits(dateStr);
  return withParens && !fallback.startsWith('(') ? `(${fallback})` : fallback;
};

const parseTimestamp = (str: any): number => {
  if (!str) return 0;
  try {
    const s = String(str).trim();
    if (!s) return 0;
    if (s.includes('T')) {
      const p = Date.parse(s);
      if (!isNaN(p)) return p;
    }
    const nums = s.match(/\d+/g)?.map(Number);
    if (!nums || nums.length < 3) {
      const parsed = Date.parse(s);
      return isNaN(parsed) ? 0 : parsed;
    }
    let y = nums[2], m = nums[1] - 1, d = nums[0];
    if (nums[0] > 1000) {
      y = nums[0];
      m = nums[1] - 1;
      d = nums[2];
    }
    return new Date(y, m, d, nums[3] || 0, nums[4] || 0, nums[5] || 0).getTime();
  } catch {
    return 0;
  }
};

const initialPersonState = (): WatchListPerson => ({
  id: `P-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
  name: '',
  passport: '',
  nationality: 'THAILAND',
  fatherName: '',
  motherName: '',
  dob: '',
  birthPlace: '',
  address: '',
  idNumber: '',
  idType: 'Passport',
  additionalInfo: ''
});

export const WatchList: React.FC<WatchListProps> = ({
  records = [],
  tempRecords = [],
  movementMap = {},
  checkingHistory = [],
  masterData = [],
  currentUser,
  showToast,
  isViewer = false,
  watchList: parentWatchList,
  setWatchList: parentSetWatchList,
  handleManualSync,
  handleFetchDataFromCloud
}) => {
  const today = new Date().toISOString().split('T')[0];

  // Local fallback state if not provided from parent
  const [localWatchList, setLocalWatchList] = useState<WatchListRecord[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error('Failed to load watchlist from storage', e);
    }
    return [];
  });

  const watchList = parentWatchList !== undefined ? parentWatchList : localWatchList;
  const setWatchList = parentSetWatchList || setLocalWatchList;

  // Secondary hydration from IndexedDB if local is empty
  useEffect(() => {
    if (watchList.length === 0) {
      miniDB.get('watchList').then((val) => {
        if (Array.isArray(val) && val.length > 0) {
          setWatchList(val);
          try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(val));
          } catch (e) {}
        }
      }).catch(() => {});
    }
  }, []);

  // Save to localStorage & IndexedDB whenever watchList updates
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(watchList));
      if (watchList.length > 0) {
        miniDB.set('watchList', watchList).catch(() => {});
      }
    } catch (e) {
      console.error('Failed to persist watchlist', e);
    }
  }, [watchList]);

  // Form State
  const [editingId, setEditingId] = useState<string | null>(null);
  const [letterDate, setLetterDate] = useState<string>(today);
  const [letterNo, setLetterNo] = useState<string>('');
  const [requestDepartment, setRequestDepartment] = useState<string>('');
  const [reasonDescription, setReasonDescription] = useState<string>('');
  const [persons, setPersons] = useState<WatchListPerson[]>([initialPersonState()]);
  const [officerName, setOfficerName] = useState<string>(() => currentUser?.name || '');
  const [officerTitle, setOfficerTitle] = useState<string>(() => currentUser?.title || '');

  // UI state
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'CONFIRMED_MATCH' | 'CLEARED_MISMATCH' | 'FLAGGED' | 'CHECKED' | 'PENDING'>('ALL');
  const [showForm, setShowForm] = useState<boolean>(false);
  const [activeCheckRecord, setActiveCheckRecord] = useState<WatchListRecord | null>(null);
  const [activeReportRecord, setActiveReportRecord] = useState<WatchListRecord | null>(null);
  const [reportOfficerName, setReportOfficerName] = useState<string>(() => currentUser?.name || 'ဦးအောင်မျိုးကျော်');
  const [reportOfficerTitle, setReportOfficerTitle] = useState<string>(() => currentUser?.title || 'ဒုတိယလဝကမှူး');
  const [reportLetterNo, setReportLetterNo] = useState<string>(`၁၀၂/လဝက(မြိတ်)/${toBurmeseDigits(new Date().getFullYear())}`);
  const [reportDate, setReportDate] = useState<string>(today);

  // Official Match Confirmation State
  const [confirmDecision, setConfirmDecision] = useState<'CONFIRMED_MATCH' | 'CLEARED_MISMATCH' | 'PENDING'>('PENDING');
  const [confirmOfficerName, setConfirmOfficerName] = useState<string>(() => currentUser?.name || 'ဦးအောင်မျိုးကျော်');
  const [confirmOfficerTitle, setConfirmOfficerTitle] = useState<string>(() => currentUser?.title || 'ဒုတိယလဝကမှူး');
  const [confirmRemarks, setConfirmRemarks] = useState<string>('');

  const reportPrintRef = useRef<HTMLDivElement>(null);

  // Sync confirmation state whenever activeCheckRecord changes
  useEffect(() => {
    if (activeCheckRecord) {
      const oc = activeCheckRecord.officialConfirmation;
      if (oc) {
        setConfirmDecision(oc.status || 'PENDING');
        setConfirmOfficerName(oc.confirmedBy || currentUser?.name || 'ဦးအောင်မျိုးကျော်');
        setConfirmOfficerTitle(oc.confirmedTitle || currentUser?.title || 'ဒုတိယလဝကမှူး');
        setConfirmRemarks(oc.remarks || '');
      } else {
        setConfirmDecision(activeCheckRecord.status === 'CONFIRMED_MATCH' ? 'CONFIRMED_MATCH' : (activeCheckRecord.status === 'CLEARED_MISMATCH' ? 'CLEARED_MISMATCH' : (activeCheckRecord.status === 'FLAGGED' ? 'CONFIRMED_MATCH' : 'PENDING')));
        setConfirmOfficerName(currentUser?.name || officerName || 'ဦးအောင်မျိုးကျော်');
        setConfirmOfficerTitle(currentUser?.title || officerTitle || 'ဒုတိယလဝကမှူး');
        setConfirmRemarks('');
      }
    }
  }, [activeCheckRecord, currentUser]);

  // Sync currentUser changes
  useEffect(() => {
    if (currentUser?.name && !officerName) {
      setOfficerName(currentUser.name);
    }
    if (currentUser?.title && !officerTitle) {
      setOfficerTitle(currentUser.title);
    }
    if (currentUser?.name && reportOfficerName === 'ဦးအောင်မျိုးကျော်') {
      setReportOfficerName(currentUser.name);
    }
    if (currentUser?.title && reportOfficerTitle === 'လဝကမှူး') {
      setReportOfficerTitle(currentUser.title);
    }
  }, [currentUser]);

  // Common departments for autocomplete
  const suggestedDepartments = [
    'အထူးစုံစမ်းစစ်ဆေးရေးဦးစီးဌာန (စရဖ)',
    'သတင်းရဲတပ်ဖွဲ့ (SB)',
    'မြန်မာနိုင်ငံရဲတပ်ဖွဲ့ (မိတ်ခရိုင်)',
    'နယ်စပ်ရေးရာဝန်ကြီးဌာန',
    'အကောက်ခွန်ဦးစီးဌာန',
    'လူကုန်ကူးမှုတားဆီးနှိမ်နင်းရေးရဲတပ်ဖွဲ့',
    'မူးယစ်ဆေးဝါးတားဆီးနှိမ်နင်းရေးရဲတပ်ဖွဲ့',
    'စစ်ဘက်ရေးရာလုံခြုံရေးတပ်ဖွဲ့',
    'တိုင်းဒေသကြီးအစိုးရအဖွဲ့ရုံး'
  ];

  // Helper: Person Management in Form
  const handleAddPerson = () => {
    setPersons(prev => [...prev, initialPersonState()]);
  };

  const handleRemovePerson = (id: string) => {
    if (persons.length <= 1) {
      showToast('အနည်းဆုံး လူပုဂ္ဂိုလ် (၁) ဦး ပါဝင်ရပါမည်');
      return;
    }
    setPersons(prev => prev.filter(p => p.id !== id));
  };

  const handlePersonChange = (id: string, field: keyof WatchListPerson, value: string) => {
    setPersons(prev => prev.map(p => {
      if (p.id === id) {
        return { ...p, [field]: value };
      }
      return p;
    }));
  };

  // Reset form
  const resetForm = () => {
    setEditingId(null);
    setLetterDate(today);
    setLetterNo('');
    setRequestDepartment('');
    setReasonDescription('');
    setPersons([initialPersonState()]);
    setOfficerName(currentUser?.name || '');
    setOfficerTitle(currentUser?.title || '');
    setShowForm(false);
  };

  // Start edit
  const startEdit = (rec: WatchListRecord) => {
    setEditingId(rec.id);
    setLetterDate(rec.letterDate || today);
    setLetterNo(rec.letterNo || '');
    setRequestDepartment(rec.requestDepartment || '');
    setReasonDescription(rec.reasonDescription || '');
    setPersons(rec.persons && rec.persons.length > 0 ? rec.persons : [initialPersonState()]);
    setOfficerName(rec.officerName || currentUser?.name || '');
    setOfficerTitle(rec.officerTitle || currentUser?.title || '');
    setShowForm(true);
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast(`ပြင်ဆင်နေသည် - စာအမှတ်: ${rec.letterNo}`);
  };

  // Save entry
  const handleSaveEntry = (e: React.FormEvent) => {
    e.preventDefault();
    if (isViewer) {
      showToast('🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် မေးမြန်းစာရင်း အသစ်သွင်း/ပြင်ဆင်ခွင့် မရှိပါ (Read-Only)');
      return;
    }
    if (!letterNo.trim()) {
      showToast('စာအမှတ် ထည့်သွင်းရန် လိုအပ်ပါသည်');
      return;
    }
    if (!requestDepartment.trim()) {
      showToast('မေးမြန်းသည့် ဌာန ထည့်သွင်းရန် လိုအပ်ပါသည်');
      return;
    }

    const validPersons = persons.filter(p => p.name.trim() || p.passport.trim());
    if (validPersons.length === 0) {
      showToast('အနည်းဆုံး အမည် သို့မဟုတ် နိုင်ငံကူးလက်မှတ်အမှတ် ထည့်သွင်းပေးပါ');
      return;
    }

    if (editingId) {
      // Update existing
      setWatchList(prev => prev.map(item => {
        if (item.id === editingId) {
          return {
            ...item,
            letterDate,
            letterNo: letterNo.trim(),
            requestDepartment: requestDepartment.trim(),
            reasonDescription: reasonDescription.trim(),
            persons: validPersons,
            totalPersons: validPersons.length,
            officerName: officerName.trim(),
            officerTitle: officerTitle.trim()
          };
        }
        return item;
      }));
      logActivity({
        action: 'UPDATE',
        module: 'WATCHLIST',
        officerName: officerName.trim() || (currentUser?.title && currentUser?.name ? `${currentUser.title} ${currentUser.name}` : undefined),
        officerRole: currentUser?.title || 'Editor',
        targetId: letterNo.trim(),
        details: `Updated WatchList record - Letter No: ${letterNo.trim()}, Dept: ${requestDepartment.trim()}, Targets: ${validPersons.map(p => p.name || p.passport).filter(Boolean).join(', ')}`
      });
      showToast('စောင့်ကြည့်စာရင်း အောင်မြင်စွာ ပြင်ဆင်ပြီးပါပြီ');
    } else {
      // Create new
      const newEntry: WatchListRecord = {
        id: `WL-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
        letterDate,
        letterNo: letterNo.trim(),
        requestDepartment: requestDepartment.trim(),
        reasonDescription: reasonDescription.trim(),
        persons: validPersons,
        totalPersons: validPersons.length,
        officerName: officerName.trim(),
        officerTitle: officerTitle.trim(),
        createdAt: new Date().toISOString(),
        status: 'PENDING'
      };
      setWatchList(prev => [newEntry, ...prev]);
      logActivity({
        action: 'CREATE',
        module: 'WATCHLIST',
        officerName: officerName.trim() || (currentUser?.title && currentUser?.name ? `${currentUser.title} ${currentUser.name}` : undefined),
        officerRole: currentUser?.title || 'Editor',
        targetId: letterNo.trim(),
        details: `New WatchList record created - Letter No: ${letterNo.trim()}, Dept: ${requestDepartment.trim()}, Targets: ${validPersons.map(p => p.name || p.passport).filter(Boolean).join(', ')}`
      });
      showToast('စောင့်ကြည့်စာရင်းအသစ် အောင်မြင်စွာ ထည့်သွင်းပြီးပါပြီ');
    }

    resetForm();
  };

  // Delete entry
  const handleDeleteEntry = (id: string) => {
    if (isViewer) {
      showToast('🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် မေးမြန်းစာရင်း ဖျက်ပစ်ခွင့် မရှိပါ (Read-Only)');
      return;
    }
    const targetWl = watchList.find(r => r.id === id);
    if (confirm('ဤစောင့်ကြည့်စာရင်း မှတ်တမ်းကို ဖျက်မည်မှာ သေချာပါသလား?')) {
      setWatchList(prev => prev.filter(r => r.id !== id));
      if (activeCheckRecord?.id === id) setActiveCheckRecord(null);
      if (activeReportRecord?.id === id) setActiveReportRecord(null);
      if (targetWl) {
        logActivity({
          action: 'DELETE',
          module: 'WATCHLIST',
          officerName: currentUser?.name,
          officerRole: currentUser?.title || 'Editor',
          targetId: targetWl.letterNo || id,
          details: `Deleted WatchList record - Letter No: ${targetWl.letterNo || id} (${targetWl.requestDepartment || '-'})`
        });
      }
      showToast('မှတ်တမ်း ဖျက်သိမ်းပြီးပါပြီ');
    }
  };

  // Clean string helper for matching
  const cleanVal = (val: any): string => {
    if (val === undefined || val === null) return '';
    return String(val).trim().toLowerCase();
  };

  // Extract normalized name tokens without honorifics
  const extractNameTokens = (name: string): string[] => {
    if (!name) return [];
    const cleaned = name
      .toLowerCase()
      .replace(/[().,/-]/g, ' ')
      .replace(/\b(u|daw|mg|ma|mr|mrs|ms|miss|dr|ဦး|ဒေါ်|မောင်|မ)\b/gi, ' ')
      .trim();
    return cleaned.split(/\s+/).filter(w => w.length >= 2);
  };

  // Deep name matching: exact, substring, or token intersection
  const isNameMatching = (query: string, target: string): boolean => {
    const q = cleanVal(query);
    const t = cleanVal(target);
    if (!q || !t) return false;
    if (q.length < 2 || t.length < 2) return false;
    if (t === q || t.includes(q) || q.includes(t)) return true;

    const qTokens = extractNameTokens(query);
    const tTokens = extractNameTokens(target);

    if (qTokens.length >= 1 && tTokens.length >= 1) {
      const matchCount = qTokens.filter(qt => tTokens.some(tt => tt === qt || tt.includes(qt) || qt.includes(tt))).length;
      if (qTokens.length >= 2 && matchCount >= Math.min(qTokens.length, 2)) return true;
      if (qTokens.length === 1 && qTokens[0].length >= 3 && matchCount === 1) return true;
    }
    return false;
  };

  // Deep Cross-reference search function (တိုက်စစ်ဆေးခြင်း)
  const performCrossReferenceCheck = (rec: WatchListRecord) => {
    const allRecords = [...records, ...tempRecords];
    let hasAnyMatch = false;
    const resultsMap: Record<string, {
      matchedRecords: (ImmRecord & { matchReasons?: string[]; isDirect?: boolean })[];
      movementData: MovementData | null;
      checkingEntries: any[];
      matchType: string[];
      summaryStatus: string;
      hasDirectMatch: boolean;
      hasRelatedMatch: boolean;
    }> = {};

    rec.persons.forEach(person => {
      const pp = (person.passport || '').trim().toUpperCase();
      const pName = (person.name || '').trim();
      const pFather = (person.fatherName || '').trim();
      const pMother = (person.motherName || '').trim();
      const pID = (person.idNumber || '').trim().toUpperCase();
      const pAddr = (person.address || '').trim();

      const matched: (ImmRecord & { matchReasons?: string[]; isDirect?: boolean })[] = [];
      const matchTypes: string[] = [];
      const allReasons: string[] = [];
      let hasDirectMatch = false;
      let hasRelatedMatch = false;

      allRecords.forEach(r => {
        const rPP = (r.passport || '').trim().toUpperCase();
        const rName = (r.fullname || '').trim();
        const rAddress = (r.address || '').trim();
        const rRemarks = (r.remarks || '').trim();
        const rContact = (r.contactDetails || '').trim();
        const rBrought = (r.broughtBy || '').trim();
        const rFormC = r.formC || {};

        const recordMatchReasons: string[] = [];
        let isRecordDirect = false;

        // 1. Direct Passport Match
        if (pp && rPP && (rPP === pp || rPP.includes(pp) || pp.includes(rPP))) {
          isRecordDirect = true;
          hasDirectMatch = true;
          recordMatchReasons.push(`တိုက်ရိုက် ပတ်စ်ပို့အမှတ် (${person.passport}) ကိုက်ညီမှု`);
          if (!matchTypes.includes('Passport No')) matchTypes.push('Passport No');
        }

        // 2. Direct Name Match
        if (pName && isNameMatching(pName, rName)) {
          isRecordDirect = true;
          hasDirectMatch = true;
          recordMatchReasons.push(`တိုက်ရိုက် ခရီးသည်အမည် (${person.name}) ကိုက်ညီမှု`);
          if (!matchTypes.includes('Name Match')) matchTypes.push('Name Match');
        }

        // 3. Father's Name Match (အဖအမည် ဆက်စပ်စစ်ဆေးမှု)
        if (pFather && pFather.length >= 2) {
          if (isNameMatching(pFather, rName)) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အဖအမည် (${person.fatherName}) ဖြင့် ခရီးသည်မှတ်တမ်း တွေ့ရှိ`);
            if (!matchTypes.includes('Father Match')) matchTypes.push('Father Match');
          } else if (isNameMatching(pFather, rBrought) || isNameMatching(pFather, rContact) || isNameMatching(pFather, rRemarks)) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အဖအမည် (${person.fatherName}) အား အဆက်အသွယ်/မှတ်ချက်တွင် တွေ့ရှိ`);
            if (!matchTypes.includes('Father in Remarks')) matchTypes.push('Father in Remarks');
          } else if (rFormC && (isNameMatching(pFather, rFormC.fatherName) || isNameMatching(pFather, rFormC.father) || isNameMatching(pFather, rFormC.guardianName))) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အဖအမည် (${person.fatherName}) အား ဆက်စပ်မှတ်တမ်းတွင် တွေ့ရှိ`);
            if (!matchTypes.includes('Father in Records')) matchTypes.push('Father in Records');
          }
        }

        // 4. Mother's Name Match (အမိအမည် ဆက်စပ်စစ်ဆေးမှု)
        if (pMother && pMother.length >= 2) {
          if (isNameMatching(pMother, rName)) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အမိအမည် (${person.motherName}) ဖြင့် ခရီးသည်မှတ်တမ်း တွေ့ရှိ`);
            if (!matchTypes.includes('Mother Match')) matchTypes.push('Mother Match');
          } else if (isNameMatching(pMother, rContact) || isNameMatching(pMother, rRemarks)) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အမိအမည် (${person.motherName}) အား အဆက်အသွယ်/မှတ်ချက်တွင် တွေ့ရှိ`);
            if (!matchTypes.includes('Mother in Remarks')) matchTypes.push('Mother in Remarks');
          } else if (rFormC && (isNameMatching(pMother, rFormC.motherName) || isNameMatching(pMother, rFormC.mother))) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`အမိအမည် (${person.motherName}) အား ဆက်စပ်မှတ်တမ်းတွင် တွေ့ရှိ`);
            if (!matchTypes.includes('Mother in Records')) matchTypes.push('Mother in Records');
          }
        }

        // 5. ID / NRC Number Match (မှတ်ပုံတင်/ID တိုက်ဆိုင်မှု)
        if (pID && pID.length >= 4) {
          const idClean = cleanVal(pID);
          if (cleanVal(rPP).includes(idClean) || cleanVal(rRemarks).includes(idClean) || cleanVal(rContact).includes(idClean) || (rFormC && (cleanVal(rFormC.idNumber).includes(idClean) || cleanVal(rFormC.nrc).includes(idClean)))) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`ID/မှတ်ပုံတင် (${person.idNumber}) တိုက်ဆိုင်တွေ့ရှိ`);
            if (!matchTypes.includes('ID Match')) matchTypes.push('ID Match');
          }
        }

        // 6. Address Match (တည်းခိုလိပ်စာ တိုက်ဆိုင်မှု)
        if (pAddr && pAddr.length >= 5) {
          const addrClean = cleanVal(pAddr);
          if (cleanVal(rAddress).includes(addrClean) || addrClean.includes(cleanVal(rAddress))) {
            hasRelatedMatch = true;
            recordMatchReasons.push(`တည်းခိုလိပ်စာ (${person.address}) တိုက်ဆိုင်တွေ့ရှိ`);
            if (!matchTypes.includes('Address Match')) matchTypes.push('Address Match');
          }
        }

        if (recordMatchReasons.length > 0) {
          allReasons.push(...recordMatchReasons);
          matched.push({
            ...r,
            matchReasons: recordMatchReasons,
            isDirect: isRecordDirect
          });
        }
      });

      // Sort matched records by timestamp desc
      matched.sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

      // MovementMap check
      const movData = pp ? movementMap[pp] || null : null;

      // Checking History entries check
      const chkEntries: any[] = [];
      (checkingHistory || []).forEach(c => {
        const cPP = (c.passport || '').trim().toUpperCase();
        const cName = (c.fullname || c.name || '').trim();
        const cFather = (c.fatherName || '').trim();
        const cMother = (c.motherName || '').trim();
        const cHost = (c.hostName || '').trim();
        const cNRC = (c.nrc || c.idNumber || '').trim();
        const cAddr = (c.confirmedAddress || c.originalAddress || '').trim();

        const chkReasons: string[] = [];

        if (pp && cPP === pp) chkReasons.push(`တိုက်ရိုက် ပတ်စ်ပို့အမှတ် (${person.passport})`);
        if (pName && isNameMatching(pName, cName)) chkReasons.push(`တိုက်ရိုက် အမည် (${person.name})`);
        if (pFather && (isNameMatching(pFather, cFather) || isNameMatching(pFather, cName) || isNameMatching(pFather, cHost))) {
          chkReasons.push(`အဖအမည် (${person.fatherName}) ဖြင့် တည်းခိုစစ်ဆေးချက် တွေ့ရှိ`);
        }
        if (pMother && (isNameMatching(pMother, cMother) || isNameMatching(pMother, cName))) {
          chkReasons.push(`အမိအမည် (${person.motherName}) ဖြင့် တည်းခိုစစ်ဆေးချက် တွေ့ရှိ`);
        }
        if (pID && (cleanVal(cNRC).includes(cleanVal(pID)) || cleanVal(cPP).includes(cleanVal(pID)))) {
          chkReasons.push(`ID/မှတ်ပုံတင် (${person.idNumber}) ဖြင့် တည်းခိုစစ်ဆေးချက် တွေ့ရှိ`);
        }
        if (pAddr && pAddr.length >= 5 && cleanVal(cAddr).includes(cleanVal(pAddr))) {
          chkReasons.push(`တည်းခိုလိပ်စာ (${person.address}) တိုက်ဆိုင်မှု`);
        }

        if (chkReasons.length > 0) {
          hasRelatedMatch = true;
          chkEntries.push({
            ...c,
            matchReason: chkReasons.join(', ')
          });
        }
      });

      if (matched.length > 0 || movData || chkEntries.length > 0) {
        hasAnyMatch = true;
      }

      let summaryStatus = 'မှတ်တမ်း မတွေ့ရှိပါ (No Record)';
      if (hasDirectMatch) {
        if (movData) {
          if (movData.outTime > movData.inTime) {
            summaryStatus = `မြိတ်မြို့မှ ထွက်ခွာသွားပြီး (${movData.out || 'Departed'}) - တိုက်ရိုက်ကိုက်ညီမှု`;
          } else if (movData.inTime > 0) {
            summaryStatus = `မြိတ်မြို့အတွင်း ရောက်ရှိ/နေထိုင်ဆဲ (${movData.loc || 'In Myeik'}) - တိုက်ရိုက်ကိုက်ညီမှု`;
          }
        } else if (matched.length > 0) {
          const latest = matched[0];
          summaryStatus = `${latest.mode === 'IN' ? 'အဝင်မှတ်တမ်းရှိ' : 'အထွက်မှတ်တမ်းရှိ'} (${toBurmeseFormattedDate(latest.timestamp.split(', ')[0])}) - တိုက်ရိုက်ကိုက်ညီမှု`;
        }
      } else if (hasRelatedMatch) {
        const uniqueReasons = Array.from(new Set(allReasons));
        if (uniqueReasons.length > 0) {
          summaryStatus = `ဆက်စပ်မှတ်တမ်း တွေ့ရှိ (${uniqueReasons.slice(0, 2).join(' / ')})`;
        } else if (chkEntries.length > 0) {
          summaryStatus = `တည်းခိုစစ်ဆေးချက် မှတ်တမ်း (${toBurmeseDigits(chkEntries.length)} ကြိမ် တွေ့ရှိ)`;
        } else {
          summaryStatus = 'ဆက်စပ်အချက်အလက် ကိုက်ညီမှု တွေ့ရှိ';
        }
      }

      resultsMap[person.id] = {
        matchedRecords: matched,
        movementData: movData,
        checkingEntries: chkEntries,
        matchType: matchTypes,
        summaryStatus,
        hasDirectMatch,
        hasRelatedMatch
      };
    });

    const isAlreadyConfirmedMatch = rec.officialConfirmation?.status === 'CONFIRMED_MATCH';
    const isAlreadyCleared = rec.officialConfirmation?.status === 'CLEARED_MISMATCH';

    const newStatus: WatchListRecord['status'] = isAlreadyConfirmedMatch
      ? 'CONFIRMED_MATCH'
      : isAlreadyCleared
      ? 'CLEARED_MISMATCH'
      : (hasAnyMatch ? 'FLAGGED' : 'CHECKED');

    const updatedRecord: WatchListRecord = {
      ...rec,
      status: newStatus,
      lastCheckedAt: new Date().toLocaleString(),
      matchResults: resultsMap,
      findingsSummary: hasAnyMatch
        ? `စနစ်တွင်း တိုက်ဆိုင်စစ်ဆေးရာတွင် မှတ်တမ်းအချက်အလက် တွေ့ရှိပါသည်`
        : 'စနစ်တွင်း စစ်ဆေးရာတွင် ဝင်ထွက်သွားလာမှု မှတ်တမ်း မတွေ့ရှိပါ'
    };

    // Update state
    setWatchList(prev => prev.map(r => r.id === rec.id ? updatedRecord : r));
    setActiveCheckRecord(updatedRecord);

    if (hasAnyMatch) {
      showToast('⚠️ စနစ်တွင်း ကိုက်ညီသည့် အချက်အလက်များ တွေ့ရှိပါသည်!');
    } else {
      showToast('✓ စနစ်တွင်း စစ်ဆေးပြီး - သွားလာမှု မှတ်တမ်း မတွေ့ရှိပါ');
    }
  };

  // Official Match Confirmation Handler
  const handleSaveOfficialConfirmation = (
    recId: string,
    decision: 'CONFIRMED_MATCH' | 'CLEARED_MISMATCH' | 'PENDING',
    officer: string,
    title: string,
    remarks: string
  ) => {
    if (isViewer) {
      showToast('🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် အရာရှိအတည်ပြုချက် ပြောင်းလဲခွင့် မရှိပါ (Read-Only)');
      return;
    }
    if (!officer.trim()) {
      showToast('စိစစ်အတည်ပြုသူ အရာရှိအမည် ထည့်သွင်းရန် လိုအပ်ပါသည်');
      return;
    }

    const confirmedAt = new Date().toISOString();
    let newStatus: WatchListRecord['status'] = 'FLAGGED';
    if (decision === 'CONFIRMED_MATCH') newStatus = 'CONFIRMED_MATCH';
    else if (decision === 'CLEARED_MISMATCH') newStatus = 'CLEARED_MISMATCH';
    else newStatus = 'FLAGGED';

    const updatedOC = {
      status: decision,
      confirmedBy: officer.trim(),
      confirmedTitle: title.trim(),
      confirmedAt,
      remarks: remarks.trim()
    };

    setWatchList(prev => {
      const updated = prev.map(item => {
        if (item.id === recId) {
          return {
            ...item,
            status: newStatus,
            officialConfirmation: updatedOC
          };
        }
        return item;
      });
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
        miniDB.set('watchList', updated).catch(() => {});
      } catch (e) {}
      return updated;
    });

    if (activeCheckRecord && activeCheckRecord.id === recId) {
      setActiveCheckRecord({
        ...activeCheckRecord,
        status: newStatus,
        officialConfirmation: updatedOC
      });
    }

    if (activeReportRecord && activeReportRecord.id === recId) {
      setActiveReportRecord({
        ...activeReportRecord,
        status: newStatus,
        officialConfirmation: updatedOC
      });
    }

    const decisionLabel = decision === 'CONFIRMED_MATCH'
      ? 'ကိုက်ညီမှု အတည်ပြုသည် (CONFIRMED MATCH)'
      : decision === 'CLEARED_MISMATCH'
      ? 'မကိုက်ညီပါ/လူမှားဖြစ်သည် (CLEARED / MISMATCH)'
      : 'စိစစ်ဆဲ (PENDING)';

    showToast(`✓ တာဝန်ခံအရာရှိ၏ အတည်ပြုချက် [${decisionLabel}] အား အောင်မြင်စွာ မှတ်တမ်းတင်ပြီးပါပြီ`);
  };

  // Open official report modal
  const openReportModal = (rec: WatchListRecord) => {
    // If not checked yet, auto run check first
    if (!rec.matchResults) {
      performCrossReferenceCheck(rec);
    }
    setActiveReportRecord(rec);
    setReportLetterNo(`၁၀၂/လဝက(မြိတ်)/${toBurmeseDigits(new Date().getFullYear())}`);
    setReportDate(today);
  };

  // Print handler
  const triggerPrintReport = () => {
    if (!reportPrintRef.current) {
      showToast('ပုံနှိပ်ရန် စာရွက် မတွေ့ရှိပါ');
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
    styleEl.innerHTML = `@media print {
      @page {
        size: A4 portrait;
        margin: 12mm 15mm 12mm 15mm;
      }
      body, html {
        background: #ffffff !important;
        color: #000000 !important;
        width: 210mm !important;
        height: auto !important;
        font-family: 'Pyidaungsu', 'Myanmar Text', sans-serif !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }
      #temp-print-area {
        display: block !important;
        width: 100% !important;
        max-width: 210mm !important;
        margin: 0 auto !important;
        padding: 0 !important;
      }
      .no-print {
        display: none !important;
      }
    }`;

    const clone = reportPrintRef.current.cloneNode(true) as HTMLElement;
    clone.style.display = 'block';
    clone.style.visibility = 'visible';

    printArea.appendChild(styleEl);
    printArea.appendChild(clone);

    window.print();
  };

  // Copy report as plain text for Signal / Telegram / Email
  const copyReportText = (rec: WatchListRecord) => {
    const p: WatchListPerson | undefined = rec.persons && rec.persons.length > 0 ? rec.persons[0] : undefined;
    const personCountSuffix = rec.persons && rec.persons.length > 1
      ? ` ပါ (${toBurmeseDigits(rec.persons.length)}) ဦး`
      : '';
    const subjectReasonPart = rec.reasonDescription && rec.reasonDescription.trim()
      ? `"${rec.reasonDescription.trim()}" ကိစ္စနှင့် စပ်လျဉ်း၍ `
      : '';
    const paraReasonPart = rec.reasonDescription && rec.reasonDescription.trim()
      ? `"${rec.reasonDescription.trim()}" ဆိုင်ရာ `
      : '';

    let text = `ပြည်ထောင်စုသမ္မတမြန်မာနိုင်ငံတော်
လူဝင်မှုကြီးကြပ်ရေးနှင့် ပြည်သူ့အင်အားဝန်ကြီးဌာန
လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်)

စာအမှတ်၊ ${reportLetterNo}
ရက်စွဲ၊ ${toBurmeseFormattedDate(reportDate)}

သို့
${rec.requestDepartment}

အကြောင်းအရာ။ ${toBurmeseFormattedDate(rec.letterDate)} ရက်စွဲပါ စာအမှတ် ${toBurmeseDigits(rec.letterNo)} ဖြင့် မေးမြန်းထားသည့် ${subjectReasonPart}(${p?.nationality || ''}) နိုင်ငံသား ${p?.name || ''} (${p?.passport || ''})${personCountSuffix} အား စနစ်တွင်း စစ်ဆေးတွေ့ရှိမှု အစီရင်ခံ တင်ပြခြင်း

၁။ အထက်ပါ ကိစ္စနှင့်ပတ်သက်၍ ${rec.requestDepartment} မှ မေးမြန်းထားသော ${paraReasonPart}စောင့်ကြည့်စာရင်းပါ အောက်ဖော်ပြပါ ပုဂ္ဂိုလ်များအား လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်) ၏ လူဝင်မှုကြီးကြပ်ရေး မှတ်တမ်းစနစ်အတွင်းရှိ ဝင်ရောက်/ထွက်ခွာမှု၊ တည်းခိုနေထိုင်မှု မှတ်တမ်းများနှင့် တိုက်ဆိုင်စစ်ဆေးခဲ့ပါသည်။

၂။ မေးမြန်းထားသည့် ပုဂ္ဂိုလ်များ၏ ကိုယ်ရေးအချက်အလက်များ -
`;

    rec.persons.forEach((person, idx) => {
      text += `\n(${toBurmeseDigits(idx + 1)}) ${person.name} (${person.passport || '-'}) - (${person.nationality || '-'})\n`;
      text += `    • အဖအမည်: ${person.fatherName || '-'}\n`;
      text += `    • အမိအမည်: ${person.motherName || '-'}\n`;
      text += `    • မွေးသက္ကရာဇ်: ${person.dob ? toBurmeseFormattedDate(person.dob) : '-'}\n`;
      text += `    • ဇာတိ: ${person.birthPlace || '-'}\n`;
      text += `    • နေရပ်လိပ်စာ: ${person.address || '-'}\n`;
      if (person.idNumber) text += `    • ID/မှတ်ပုံတင်: ${person.idNumber} (${person.idType || 'ID'})\n`;
      if (person.additionalInfo) text += `    • အခြားမှတ်ချက်: ${person.additionalInfo}\n`;
    });

    text += `\n၃။ စနစ်အတွင်း တိုက်ဆိုင်စစ်ဆေး တွေ့ရှိကိုက်ညီသည့် အချက်အလက်များ (Cross-Referenced System Findings) -\n`;

    rec.persons.forEach((person, idx) => {
      const pMatch = rec.matchResults?.[person.id];
      const pRecs: any[] = pMatch?.matchedRecords || [];
      const pChks = pMatch?.checkingEntries || [];

      text += `\n[ ${toBurmeseDigits(idx + 1)}. ${person.name} (${person.passport || 'NO PP'}) ]\n`;
      text += `  • စစ်ဆေးတွေ့ရှိမှု အခြေအနေ: ${pMatch?.summaryStatus || 'စစ်ဆေးဆဲ'}\n`;

      if (pRecs.length > 0) {
        text += `  • ဝင်ရောက်/ထွက်ခွာမှု မှတ်တမ်း (${toBurmeseDigits(pRecs.length)} ကြိမ် တွေ့ရှိ):\n`;
        pRecs.forEach((r, rIdx) => {
          const stayPeriod = r.stayFrom ? `${toBurmeseFormattedDate(r.stayFrom)} မှ ${toBurmeseFormattedDate(r.stayTo)} ထိ` : '-';
          const matchTag = r.matchReasons && r.matchReasons.length > 0 ? ` | ကိုက်ညီမှု: ${r.matchReasons.join(', ')}` : '';
          text += `    (${toBurmeseDigits(rIdx + 1)}) ရက်စွဲ: ${toBurmeseFormattedDate(r.timestamp.split(', ')[0])} | ${r.mode === 'IN' ? 'အဝင်' : 'အထွက်'} | ခရီးသည်အမည်: ${r.fullname} | ယာဉ်/လေယာဉ်: ${r.vehicleInfo || '-'} | ဗီဇာ: ${r.visaType || '-'} | နေခွင့်: ${stayPeriod} | တည်းခိုရာ: ${r.address || '-'}${matchTag} | တာဝန်ခံ: ${r.officialName || '-'}\n`;
        });
      }

      if (pChks.length > 0) {
        text += `  • တည်းခိုနေထိုင်မှု စစ်ဆေးချက်မှတ်တမ်း (${toBurmeseDigits(pChks.length)} ကြိမ်):\n`;
        pChks.forEach((c: any, cIdx: number) => {
          const matchTag = c.matchReason ? ` | ကိုက်ညီမှု: ${c.matchReason}` : '';
          text += `    (${toBurmeseDigits(cIdx + 1)}) ရက်စွဲ: ${toBurmeseFormattedDate(c.checkDate)} | အတည်ပြုလိပ်စာ: ${c.confirmedAddress || c.originalAddress || '-'}${matchTag} | ခွင့်ပြုချက်: ${c.status || '-'} | စစ်ဆေးသူ: ${c.officerName || '-'}\n`;
        });
      }

      if (pRecs.length === 0 && pChks.length === 0) {
        text += `  • သွားလာ/တည်းခိုမှု မှတ်တမ်း: စနစ်အတွင်း ဝင်ရောက်/ထွက်ခွာမှုနှင့် တည်းခိုနေထိုင်မှု မှတ်တမ်း (လုံးဝ) မတွေ့ရှိပါ (NO RECORD FOUND)\n`;
      }
    });

    if (rec.officialConfirmation) {
      text += `\n၄။ တာဝန်ခံအရာရှိ၏ စိစစ်အတည်ပြုချက် (Official Verification Finding) -\n`;
      if (rec.officialConfirmation.status === 'CONFIRMED_MATCH') {
        text += `  • ဆုံးဖြတ်ချက်: စနစ်အတွင်း တိုက်ဆိုင်စစ်ဆေးတွေ့ရှိရသော အချက်အလက်များနှင့် မေးမြန်းထားသည့် စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်မှာ အမှန်တကယ် (ကိုက်ညီမှုရှိကြောင်း) တာဝန်ခံအရာရှိမှ စိစစ်အတည်ပြုပါသည်။\n`;
      } else if (rec.officialConfirmation.status === 'CLEARED_MISMATCH') {
        text += `  • ဆုံးဖြတ်ချက်: စနစ်အတွင်း တွေ့ရှိရသော အချက်အလက်များသည် အမည်တူ/လိပ်စာတူ ဆက်စပ်ရုံမျှသာဖြစ်ပြီး မေးမြန်းထားသည့် စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်နှင့် (ကိုက်ညီမှုမရှိကြောင်း/လူမှားဖြစ်ကြောင်း) တာဝန်ခံအရာရှိမှ စိစစ်အတည်ပြုပါသည်။\n`;
      } else {
        text += `  • ဆုံးဖြတ်ချက်: စနစ်တွင်း တိုက်ဆိုင်စစ်ဆေးချက်များအရ စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်နှင့် ကိုက်ညီမှု ရှိ/မရှိအား ဆက်လက် စိစစ်ဆဲ ဖြစ်ပါသည်။\n`;
      }
      if (rec.officialConfirmation.remarks) {
        text += `  • စိစစ်ချက်မှတ်ချက်: ${rec.officialConfirmation.remarks}\n`;
      }
      if (rec.officialConfirmation.confirmedBy) {
        text += `  • စိစစ်အတည်ပြုသူ: ${rec.officialConfirmation.confirmedBy} (${rec.officialConfirmation.confirmedTitle || 'ဒုတိယလဝကမှူး'}) | ရက်စွဲ: ${toBurmeseFormattedDate(rec.officialConfirmation.confirmedAt)}\n`;
      }
      text += `\n၅။ သိရှိနိုင်ပါရန်နှင့် လိုအပ်သလို ဆက်လက်ဆောင်ရွက်နိုင်ပါရန် လေးစားစွာဖြင့် အစီရင်ခံတင်ပြအပ်ပါသည်။
`;
    } else {
      text += `\n၄။ သိရှိနိုင်ပါရန်နှင့် လိုအပ်သလို ဆက်လက်ဆောင်ရွက်နိုင်ပါရန် လေးစားစွာဖြင့် အစီရင်ခံတင်ပြအပ်ပါသည်။
`;
    }

    text += `\n(${reportOfficerName || 'တာဝန်ခံအရာရှိ'})
${reportOfficerTitle || 'ဒုတိယလဝကမှူး'}
လူဝင်မှုကြီးကြပ်ရေးနှင့် ပြည်သူ့အင်အားဝန်ကြီးဌာန
လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်)`;

    navigator.clipboard.writeText(text);
    showToast('အစီရင်ခံစာ စာသားကို Copy ကူးယူပြီးပါပြီ (Clipboard Copied)');
  };

  // Export Watch-list to Excel
  const exportToExcel = () => {
    if (watchList.length === 0) {
      showToast('Export ပြုလုပ်ရန် စောင့်ကြည့်စာရင်း မရှိပါ');
      return;
    }

    const rows: any[] = [];
    watchList.forEach((rec, recIdx) => {
      rec.persons.forEach((p, pIdx) => {
        rows.push({
          'စဉ်': recIdx + 1,
          'စာအမှတ်': rec.letterNo,
          'မေးမြန်းသည့်ရက်စွဲ': toBurmeseFormattedDate(rec.letterDate),
          'မေးမြန်းသည့်ဌာန': rec.requestDepartment,
          'အကြောင်းအရာ': rec.reasonDescription,
          'လူပုဂ္ဂိုလ်အမည်': p.name,
          'နိုင်ငံကူးလက်မှတ်': p.passport,
          'နိုင်ငံသား': p.nationality,
          'အဖအမည်': p.fatherName || '',
          'အမိအမည်': p.motherName || '',
          'မွေးသက္ကရာဇ်': p.dob ? toBurmeseFormattedDate(p.dob) : '',
          'ဇာတိ': p.birthPlace || '',
          'လိပ်စာ': p.address || '',
          'ID / မှတ်ပုံတင်': p.idNumber || '',
          'စစ်ဆေးမှုအခြေအနေ': rec.status || 'PENDING',
          'အရာရှိအတည်ပြုချက်': rec.officialConfirmation?.status || '-',
          'အတည်ပြုသူအရာရှိ': rec.officialConfirmation?.confirmedBy || '',
          'နောက်ဆုံးစစ်ဆေးချိန်': rec.lastCheckedAt || '',
          'တွေ့ရှိချက်': rec.matchResults?.[p.id]?.summaryStatus || rec.findingsSummary || ''
        });
      });
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Watch-list');
    XLSX.writeFile(wb, `Watchlist_Inquiries_${today}.xlsx`);
    showToast('Excel ဖိုင် အောင်မြင်စွာ Export ပြုလုပ်ပြီးပါပြီ');
  };

  // Export Backup JSON
  const exportBackupJSON = () => {
    const backupPayload = {
      version: '1.0',
      exportedAt: new Date().toISOString(),
      source: 'WatchList_Module',
      data: watchList
    };
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupPayload, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `watchlist_backup_${today}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showToast('Watch-list JSON Backup သိမ်းဆည်းပြီးပါပြီ');
  };

  // Import Backup JSON
  const importBackupJSON = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], "UTF-8");
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          let listToRestore: WatchListRecord[] = [];
          if (Array.isArray(parsed)) {
            listToRestore = parsed;
          } else if (parsed && Array.isArray(parsed.data)) {
            listToRestore = parsed.data;
          }

          if (listToRestore.length > 0) {
            setWatchList(listToRestore);
            try {
              localStorage.setItem(STORAGE_KEY, JSON.stringify(listToRestore));
              miniDB.set('watchList', listToRestore).catch(() => {});
            } catch (err) {}
            showToast(`Watch-list (${toBurmeseDigits(listToRestore.length)}) ခု အောင်မြင်စွာ Restore လုပ်ပြီးပါပြီ`);
          } else {
            showToast('မှားယွင်းသော ဖိုင်ပုံစံဖြစ်နေပါသည်');
          }
        } catch (err) {
          showToast('JSON ဖတ်ရှုရာတွင် အမှားဖြစ်ပေါ်ပါသည်');
        }
      };
    }
  };

  // Filtered List
  const filteredList = useMemo(() => {
    return watchList.filter(rec => {
      // Status filter
      if (statusFilter !== 'ALL') {
        if (rec.status !== statusFilter) {
          return false;
        }
      }
      // Search query filter
      if (!searchQuery.trim()) return true;
      const q = searchQuery.toLowerCase().trim();
      const inLetter = (rec.letterNo || '').toLowerCase().includes(q) ||
                       (rec.requestDepartment || '').toLowerCase().includes(q) ||
                       (rec.reasonDescription || '').toLowerCase().includes(q);
      const inPersons = rec.persons.some(p => 
        (p.name || '').toLowerCase().includes(q) ||
        (p.passport || '').toLowerCase().includes(q) ||
        (p.nationality || '').toLowerCase().includes(q) ||
        (p.fatherName || '').toLowerCase().includes(q) ||
        (p.idNumber || '').toLowerCase().includes(q) ||
        (p.address || '').toLowerCase().includes(q)
      );
      const inOfficer = (rec.officialConfirmation?.confirmedBy || '').toLowerCase().includes(q) ||
                        (rec.officialConfirmation?.remarks || '').toLowerCase().includes(q);
      return inLetter || inPersons || inOfficer;
    });
  }, [watchList, searchQuery, statusFilter]);

  // Quick stats
  const stats = useMemo(() => {
    const totalInquiries = watchList.length;
    let totalPersons = 0;
    let confirmedCount = 0;
    let clearedCount = 0;
    let flaggedCount = 0;
    let checkedCount = 0;
    let pendingCount = 0;

    watchList.forEach(w => {
      totalPersons += (w.persons?.length || 1);
      if (w.status === 'CONFIRMED_MATCH') confirmedCount++;
      else if (w.status === 'CLEARED_MISMATCH') clearedCount++;
      else if (w.status === 'FLAGGED') flaggedCount++;
      else if (w.status === 'CHECKED') checkedCount++;
      else pendingCount++;
    });

    return { totalInquiries, totalPersons, confirmedCount, clearedCount, flaggedCount, checkedCount, pendingCount };
  }, [watchList]);

  return (
    <div className="max-w-[1600px] mx-auto py-6 px-3 sm:px-6 space-y-6 select-text">
      {/* Viewer Notice Banner */}
      {isViewer && (
        <div className="p-4 bg-amber-50 border-2 border-amber-300 rounded-2xl text-amber-900 text-xs font-bold flex items-center gap-3 shadow-sm no-print">
          <div className="p-2 bg-amber-200/70 rounded-xl text-amber-800 shrink-0">
            <ShieldAlert size={22} />
          </div>
          <div>
            <div className="font-black text-sm uppercase text-amber-950 flex items-center gap-2">
              <span>🔒 Viewer Level (Read-Only Mode) သတ်မှတ်ထားပါသည်</span>
              <span className="text-[10px] bg-amber-200 text-amber-900 px-2 py-0.5 rounded-full font-mono">SEARCH / VIEW ONLY</span>
            </div>
            <div className="text-amber-800 mt-0.5">
              စောင့်ကြည့်/စုံစမ်းမေးမြန်းမှု စာရင်းများအား ရှာဖွေစစ်ဆေးခြင်း (Search & Cross-Reference Check) သာ ဆောင်ရွက်နိုင်ပြီး၊ မေးမြန်းစာရင်း အသစ်သွင်းခြင်း၊ ပြင်ဆင်ခြင်း၊ ဖျက်ပစ်ခြင်းနှင့် အတည်ပြုချက်ပြောင်းလဲခြင်းများ ပိတ်ထားပါသည်
            </div>
          </div>
        </div>
      )}

      {/* HEADER & TOP STATS BAR */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 border border-slate-700/80 rounded-3xl p-5 sm:p-7 text-white shadow-2xl relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-rose-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="space-y-2">
            <div className="flex items-center gap-3 flex-wrap">
              <div className="p-2.5 rounded-2xl bg-rose-600/30 text-rose-300 border border-rose-500/40 shadow-inner">
                <ShieldAlert size={26} className="text-rose-400 animate-pulse" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white uppercase">
                    Watch-list & Inquiry Management
                  </h1>
                  <span className="bg-rose-500/20 text-rose-300 border border-rose-400/40 text-[10px] font-black px-2 py-0.5 rounded-full">
                    စောင့်ကြည့်စာရင်း
                  </span>
                </div>
                <p className="text-xs text-slate-300 font-bold mt-0.5">
                  ဌာနဆိုင်ရာ မေးမြန်းစုံစမ်းစာရင်းများ မှတ်တမ်းတင်ခြင်း၊ စနစ်တွင်း အလိုအလျောက် တိုက်ဆိုင်စစ်ဆေးခြင်းနှင့် အစီရင်ခံစာထုတ်ယူခြင်း
                </p>
              </div>
            </div>
          </div>

          {/* Quick Actions */}
          <div className="flex items-center gap-2.5 flex-wrap">
            {!isViewer && (
              <button
                onClick={() => {
                  if (showForm) {
                    resetForm();
                  } else {
                    resetForm();
                    setShowForm(true);
                  }
                }}
                className={`px-4 py-2.5 rounded-xl font-black text-xs transition-all flex items-center gap-2 shadow-lg cursor-pointer ${
                  showForm
                    ? 'bg-slate-700 hover:bg-slate-600 text-white'
                    : 'bg-rose-600 hover:bg-rose-500 text-white ring-2 ring-rose-400/50'
                }`}
              >
                {showForm ? <X size={15} /> : <Plus size={15} />}
                <span>{showForm ? 'ဖောင် ပိတ်မည်' : 'မေးမြန်းစာရင်း အသစ်သွင်းရန်'}</span>
              </button>
            )}

            <button
              onClick={async () => {
                if (handleManualSync) {
                  handleManualSync();
                } else {
                  showToast("Cloud သို့ Watch-list Sync ပြုလုပ်နေပါသည်...");
                  const ok = await saveCollectionToFirestore('watchList', watchList, true);
                  if (ok) showToast(" Watch-list Cloud Sync အောင်မြင်ပါသည်");
                }
              }}
              className="px-3.5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-black text-xs transition-all flex items-center gap-1.5 shadow-md cursor-pointer"
              title="Cloud Sync (All Devices)"
            >
              <RefreshCw size={15} />
              <span className="hidden sm:inline">Cloud Sync</span>
            </button>

            <button
              onClick={async () => {
                if (handleFetchDataFromCloud) {
                  handleFetchDataFromCloud(true);
                } else {
                  showToast("Cloud မှ Watch-list ရယူနေပါသည်...");
                  const items = await fetchCollectionFromFirestore('watchList');
                  if (items && Array.isArray(items)) {
                    setWatchList(items);
                    showToast(` Cloud မှ Watch-list (${items.length}) ခု ရယူပြီးပါပြီ`);
                  }
                }
              }}
              className="px-3.5 py-2.5 rounded-xl bg-cyan-700 hover:bg-cyan-600 text-white font-black text-xs transition-all flex items-center gap-1.5 shadow-md cursor-pointer"
              title="Cloud Fetch"
            >
              <Download size={15} />
              <span className="hidden sm:inline">Cloud Fetch</span>
            </button>

            <button
              onClick={exportToExcel}
              className="px-3.5 py-2.5 rounded-xl bg-emerald-700 hover:bg-emerald-600 text-white font-black text-xs transition-all flex items-center gap-1.5 shadow-md cursor-pointer"
              title="Export Watch-list to Excel"
            >
              <FileSpreadsheet size={15} />
              <span className="hidden sm:inline">Excel Export</span>
            </button>

            <button
              onClick={exportBackupJSON}
              className="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 font-black text-xs transition-all flex items-center gap-1.5 border border-slate-700 cursor-pointer"
              title="Download JSON Backup"
            >
              <Download size={15} />
              <span className="hidden sm:inline">Backup</span>
            </button>

            {!isViewer && (
              <label
                className="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 font-black text-xs transition-all flex items-center gap-1.5 border border-slate-700 cursor-pointer"
                title="Restore JSON Backup"
              >
                <Upload size={15} />
                <span className="hidden sm:inline">Restore</span>
                <input type="file" accept=".json" onChange={importBackupJSON} className="hidden" />
              </label>
            )}
          </div>
        </div>

        {/* Stats Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 mt-6 pt-5 border-t border-slate-800">
          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-slate-700/60 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-slate-400 tracking-wider">မေးမြန်းစာ စုစုပေါင်း</div>
              <div className="text-lg font-black text-white mt-0.5">{stats.totalInquiries} <span className="text-[11px] font-normal text-slate-400">စောင်</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-indigo-500/20 text-indigo-300">
              <FileText size={16} />
            </div>
          </div>

          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-slate-700/60 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-slate-400 tracking-wider">စောင့်ကြည့် ပုဂ္ဂိုလ်</div>
              <div className="text-lg font-black text-white mt-0.5">{stats.totalPersons} <span className="text-[11px] font-normal text-slate-400">ဦး</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-cyan-500/20 text-cyan-300">
              <Users size={16} />
            </div>
          </div>

          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-rose-500/30 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-rose-300 tracking-wider">အရာရှိ အတည်ပြုပြီး</div>
              <div className="text-lg font-black text-rose-400 mt-0.5">{stats.confirmedCount} <span className="text-[11px] font-normal text-rose-300/80">စောင်</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-rose-500/20 text-rose-400">
              <ShieldAlert size={16} />
            </div>
          </div>

          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-blue-500/30 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-blue-300 tracking-wider">မကိုက်ညီ/လူမှားရှင်းပြီး</div>
              <div className="text-lg font-black text-blue-400 mt-0.5">{stats.clearedCount} <span className="text-[11px] font-normal text-blue-300/80">စောင်</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-blue-500/20 text-blue-400">
              <ShieldCheck size={16} />
            </div>
          </div>

          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-amber-500/30 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-amber-300 tracking-wider">တွေ့ရှိမှုရှိ/စိစစ်ဆဲ</div>
              <div className="text-lg font-black text-amber-400 mt-0.5">{stats.flaggedCount} <span className="text-[11px] font-normal text-amber-300/80">စောင်</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-amber-500/20 text-amber-300">
              <AlertTriangle size={16} />
            </div>
          </div>

          <div className="bg-slate-800/60 backdrop-blur-sm p-3 rounded-2xl border border-slate-700/60 flex items-center justify-between">
            <div>
              <div className="text-[10px] font-black uppercase text-slate-400 tracking-wider">မှတ်တမ်းမတွေ့/မစစ်ရသေး</div>
              <div className="text-lg font-black text-slate-300 mt-0.5">{stats.checkedCount + stats.pendingCount} <span className="text-[11px] font-normal text-slate-400">စောင်</span></div>
            </div>
            <div className="p-1.5 rounded-xl bg-slate-700/40 text-slate-400">
              <CheckCircle2 size={16} />
            </div>
          </div>
        </div>
      </div>

      {/* ENTRY FORM MODAL / COLLAPSIBLE SECTION */}
      {showForm && (
        <form onSubmit={handleSaveEntry} className="bg-white rounded-3xl p-6 sm:p-8 shadow-xl border-2 border-indigo-200 animate-in slide-in-from-top-4 duration-300 space-y-6">
          <div className="flex justify-between items-center border-b pb-4">
            <div className="flex items-center gap-2.5">
              <div className="p-2 bg-indigo-100 text-indigo-800 rounded-xl">
                <FileText size={20} />
              </div>
              <div>
                <h2 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  {editingId ? 'စောင့်ကြည့်စာရင်း အချက်အလက် ပြင်ဆင်ရန်' : 'စောင့်ကြည့်/စုံစမ်းမေးမြန်းမှု အချက်အလက် သွင်းယူခြင်း'}
                </h2>
                <p className="text-xs text-slate-500 font-bold">
                  (Date)ရက်စွဲပါစာအမှတ် (Letter No) ဖြင့် (မည်သည့်ဌာန) မှ (အကြောင်းအရာ) အတွက် မေးမြန်းထားသည့် အချက်အလက်များ
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={resetForm}
              className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-100"
            >
              <X size={20} />
            </button>
          </div>

          {/* SECTION 1: INQUIRY LETTER DETAILS */}
          <div className="p-5 bg-slate-50 rounded-2xl border border-slate-200 space-y-4">
            <div className="flex items-center gap-2 text-xs font-black uppercase text-indigo-900 tracking-wider">
              <Building size={16} className="text-indigo-600" />
              <span>၁။ မေးမြန်းသည့် စာအမှတ်နှင့် ဌာနဆိုင်ရာ အချက်အလက်များ</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label className="input-label font-bold text-slate-700 text-xs">
                  ရက်စွဲ (Inquiry Date) <span className="text-rose-500">*</span>
                </label>
                <input
                  type="date"
                  value={letterDate}
                  onChange={(e) => setLetterDate(e.target.value)}
                  className="input-field bg-white border-slate-300 font-bold"
                  required
                />
              </div>

              <div>
                <label className="input-label font-bold text-slate-700 text-xs">
                  စာအမှတ် (Letter Number that asked) <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={letterNo}
                  onChange={(e) => setLetterNo(e.target.value)}
                  placeholder="e.g. ၁၀၂/စရဖ(မိတ်)/၂၀၂၆"
                  className="input-field bg-white border-slate-300 font-bold"
                  required
                />
              </div>

              <div>
                <label className="input-label font-bold text-slate-700 text-xs">
                  မည်သည့်ဌာန (Requesting Department) <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  list="deptList"
                  value={requestDepartment}
                  onChange={(e) => setRequestDepartment(e.target.value)}
                  placeholder="ဌာနအမည် ရွေးချယ် သို့မဟုတ် ရိုက်ထည့်ပါ"
                  className="input-field bg-white border-slate-300 font-bold"
                  required
                />
                <datalist id="deptList">
                  {suggestedDepartments.map((d, i) => (
                    <option key={i} value={d} />
                  ))}
                </datalist>
              </div>

              <div>
                <label className="input-label font-bold text-slate-700 text-xs">
                  အကြောင်းအရာဖော်ပြချက် (Subject / Reason)
                </label>
                <input
                  type="text"
                  value={reasonDescription}
                  onChange={(e) => setReasonDescription(e.target.value)}
                  placeholder="e.g. တရားမဝင် ဝင်ထွက်သွားလာမှု စုံစမ်းစစ်ဆေးပေးရန်"
                  className="input-field bg-white border-slate-300"
                />
              </div>
            </div>
          </div>

          {/* SECTION 2: PERSON(S) DETAILS */}
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <div className="flex items-center gap-2 text-xs font-black uppercase text-indigo-900 tracking-wider">
                <Users size={16} className="text-indigo-600" />
                <span>၂။ မေးမြန်းထားသည့် လူပုဂ္ဂိုလ် ကိုယ်ရေးအချက်အလက်များ ({persons.length} ဦး)</span>
              </div>
              <button
                type="button"
                onClick={handleAddPerson}
                className="px-3 py-1.5 rounded-xl bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-xs font-black flex items-center gap-1.5 transition-all cursor-pointer"
              >
                <Plus size={14} />
                <span>+ လူပုဂ္ဂိုလ် ထပ်မံထည့်သွင်းရန်</span>
              </button>
            </div>

            <div className="space-y-4">
              {persons.map((person, index) => (
                <div
                  key={person.id}
                  className="p-5 bg-slate-50/80 rounded-2xl border-2 border-indigo-100 relative space-y-4 transition-all hover:border-indigo-300"
                >
                  <div className="flex justify-between items-center border-b border-slate-200 pb-2">
                    <div className="flex items-center gap-2">
                      <span className="w-6 h-6 rounded-full bg-indigo-600 text-white font-black text-xs flex items-center justify-center">
                        {index + 1}
                      </span>
                      <span className="font-black text-slate-800 text-xs uppercase">
                        {person.name || `ပုဂ္ဂိုလ်အမှတ် (${index + 1})`}
                      </span>
                      {person.passport && (
                        <span className="font-mono text-xs font-bold bg-white px-2 py-0.5 rounded border border-slate-200 text-indigo-700">
                          {person.passport}
                        </span>
                      )}
                    </div>
                    {persons.length > 1 && (
                      <button
                        type="button"
                        onClick={() => handleRemovePerson(person.id)}
                        className="text-rose-500 hover:text-rose-700 p-1 rounded-lg hover:bg-rose-50 transition-colors"
                        title="ဤပုဂ္ဂိုလ်ကို ပယ်ဖျက်မည်"
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>

                  {/* Primary fields: Name, Passport, Nationality */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="input-label font-bold text-slate-700 text-xs">
                        အမည် (Full Name) <span className="text-rose-500">*</span>
                      </label>
                      <input
                        type="text"
                        value={person.name}
                        onChange={(e) => handlePersonChange(person.id, 'name', e.target.value)}
                        placeholder="e.g. SOMCHAI PRASERT"
                        className="input-field bg-white border-slate-300 font-bold"
                        required
                      />
                    </div>

                    <div>
                      <label className="input-label font-bold text-slate-700 text-xs">
                        နိုင်ငံကူးလက်မှတ်အမှတ် (Passport No)
                      </label>
                      <input
                        type="text"
                        value={person.passport}
                        onChange={(e) => handlePersonChange(person.id, 'passport', e.target.value.toUpperCase())}
                        placeholder="e.g. AA1234567"
                        className="input-field bg-white border-slate-300 font-mono font-bold uppercase"
                      />
                    </div>

                    <div>
                      <label className="input-label font-bold text-slate-700 text-xs">
                        နိုင်ငံသား (Nationality)
                      </label>
                      <input
                        type="text"
                        list="natList"
                        value={person.nationality}
                        onChange={(e) => handlePersonChange(person.id, 'nationality', e.target.value.toUpperCase())}
                        placeholder="e.g. THAILAND"
                        className="input-field bg-white border-slate-300 font-bold uppercase"
                      />
                      <datalist id="natList">
                        {(masterData || []).filter(m => m && m.type === 'Nationality' && m.name).map(m => (
                          <option key={m.id} value={m.name} />
                        ))}
                      </datalist>
                    </div>
                  </div>

                  {/* Extra personal info (ကိုယ်ရေးအချက်အလက်များ - သိသမျှထည့်နိုင်) */}
                  <div className="pt-2 border-t border-slate-200/80">
                    <div className="text-[11px] font-bold text-slate-500 mb-2">
                      ကိုယ်ရေးအချက်အလက်များ (သိသမျှထည့်သွင်းနိုင်သည် - ရှာဖွေရာတွင် တိုက်ဆိုင်စစ်ဆေးရန်)
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
                      <div>
                        <label className="input-label text-[11px]">အဖအမည် (Father Name)</label>
                        <input
                          type="text"
                          value={person.fatherName || ''}
                          onChange={(e) => handlePersonChange(person.id, 'fatherName', e.target.value)}
                          placeholder="အဖအမည်"
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div>
                        <label className="input-label text-[11px]">အမိအမည် (Mother Name)</label>
                        <input
                          type="text"
                          value={person.motherName || ''}
                          onChange={(e) => handlePersonChange(person.id, 'motherName', e.target.value)}
                          placeholder="အမိအမည်"
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div>
                        <label className="input-label text-[11px]">မွေးသက္ကရာဇ် (DOB)</label>
                        <input
                          type="date"
                          value={person.dob || ''}
                          onChange={(e) => handlePersonChange(person.id, 'dob', e.target.value)}
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div>
                        <label className="input-label text-[11px]">ဇာတိနေရာ (Birth Place)</label>
                        <input
                          type="text"
                          value={person.birthPlace || ''}
                          onChange={(e) => handlePersonChange(person.id, 'birthPlace', e.target.value)}
                          placeholder="မွေးဖွားရာဒေသ/ဇာတိ"
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div>
                        <label className="input-label text-[11px]">ID / မှတ်ပုံတင်အမှတ် (ID Number)</label>
                        <input
                          type="text"
                          value={person.idNumber || ''}
                          onChange={(e) => handlePersonChange(person.id, 'idNumber', e.target.value)}
                          placeholder="NRC / National ID No"
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div>
                        <label className="input-label text-[11px]">ID အမျိုးအစား (ID Type)</label>
                        <select
                          value={person.idType || 'Passport'}
                          onChange={(e) => handlePersonChange(person.id, 'idType', e.target.value)}
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        >
                          <option value="Passport">Passport</option>
                          <option value="National ID">National ID</option>
                          <option value="NRC (နိုင်ငံသားစိစစ်ရေးကတ်)">NRC (နိုင်ငံသားစိစစ်ရေးကတ်)</option>
                          <option value="Border Pass (နယ်စပ်ဖြတ်သန်းခွင့်)">Border Pass (နယ်စပ်ဖြတ်သန်းခွင့်)</option>
                          <option value="Seaman Book">Seaman Book</option>
                          <option value="Other">အခြား (Other)</option>
                        </select>
                      </div>

                      <div className="lg:col-span-2">
                        <label className="input-label text-[11px]">နေရပ်လိပ်စာ (Address / Residence)</label>
                        <input
                          type="text"
                          value={person.address || ''}
                          onChange={(e) => handlePersonChange(person.id, 'address', e.target.value)}
                          placeholder="နေထိုင်ရာလိပ်စာ / တည်းခိုရာနေရာ"
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>

                      <div className="lg:col-span-4">
                        <label className="input-label text-[11px]">အခြား ဆက်စပ်အချက်အလက်များ (Additional Info / Notes)</label>
                        <input
                          type="text"
                          value={person.additionalInfo || ''}
                          onChange={(e) => handlePersonChange(person.id, 'additionalInfo', e.target.value)}
                          placeholder="စုံစမ်းရာတွင် အထောက်အကူပြုမည့် အခြားမှတ်ချက်များ..."
                          className="input-field bg-white border-slate-200 py-1.5 text-xs"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* SECTION 3: RESPONSIBLE OFFICER & ACTIONS */}
          <div className="p-5 bg-indigo-50/60 rounded-2xl border border-indigo-200 flex flex-col md:flex-row justify-between items-center gap-4">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 w-full md:w-auto">
              <div>
                <label className="input-label text-[11px] font-bold text-indigo-900">တာဝန်ခံအရာရှိအမည်</label>
                <input
                  type="text"
                  value={officerName}
                  onChange={(e) => setOfficerName(e.target.value)}
                  placeholder="အရာရှိအမည်"
                  className="input-field bg-white border-indigo-200 text-xs font-bold"
                />
              </div>
              <div>
                <label className="input-label text-[11px] font-bold text-indigo-900">ရာထူး</label>
                <input
                  type="text"
                  value={officerTitle}
                  onChange={(e) => setOfficerTitle(e.target.value)}
                  placeholder="ရာထူး"
                  className="input-field bg-white border-indigo-200 text-xs font-bold"
                />
              </div>
            </div>

            <div className="flex items-center gap-3 w-full md:w-auto justify-end">
              <button
                type="button"
                onClick={resetForm}
                className="px-5 py-2.5 rounded-xl border border-slate-300 text-slate-700 hover:bg-slate-100 font-bold text-xs cursor-pointer"
              >
                မလုပ်ဆောင်ပါ (Cancel)
              </button>

              <button
                type="submit"
                className="px-6 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs shadow-lg flex items-center gap-2 cursor-pointer"
              >
                <Check size={16} />
                <span>{editingId ? 'ပြင်ဆင်မှု သိမ်းဆည်းမည်' : 'စောင့်ကြည့်စာရင်း သိမ်းဆည်းမည်'}</span>
              </button>
            </div>
          </div>
        </form>
      )}

      {/* FILTER & SEARCH CONTROL BAR */}
      <div className="bg-white rounded-2xl p-4 shadow-sm border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
        <div className="relative w-full md:w-96">
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="စာအမှတ်၊ ဌာန၊ အမည်၊ ပတ်စ်ပို့၊ လိပ်စာ ရှာဖွေရန်..."
            className="input-field pl-10 pr-4 py-2 text-xs"
          />
          <Search size={16} className="absolute left-3.5 top-2.5 text-slate-400" />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery('')}
              className="absolute right-3 top-2.5 text-slate-400 hover:text-slate-600"
            >
              <X size={14} />
            </button>
          )}
        </div>

        {/* Status Filter Tabs */}
        <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl w-full md:w-auto overflow-x-auto">
          {[
            { id: 'ALL', label: 'အားလုံး (All)', count: watchList.length },
            { id: 'CONFIRMED_MATCH', label: 'အရာရှိ အတည်ပြုပြီး (Confirmed Match)', count: stats.confirmedCount, color: 'text-rose-600' },
            { id: 'CLEARED_MISMATCH', label: 'မကိုက်ညီ/လူမှားရှင်းပြီး (Cleared)', count: stats.clearedCount, color: 'text-blue-600' },
            { id: 'FLAGGED', label: 'တွေ့ရှိမှုရှိ/စိစစ်ဆဲ (Matches)', count: stats.flaggedCount, color: 'text-amber-600' },
            { id: 'CHECKED', label: 'မတွေ့ရှိပါ (No Match)', count: stats.checkedCount, color: 'text-emerald-600' },
            { id: 'PENDING', label: 'မစစ်ရသေး (Pending)', count: stats.pendingCount, color: 'text-slate-600' },
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setStatusFilter(tab.id as any)}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all flex items-center gap-1.5 whitespace-nowrap cursor-pointer ${
                statusFilter === tab.id
                  ? 'bg-white text-slate-900 shadow-sm ring-1 ring-slate-200'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <span className={tab.color}>{tab.label}</span>
              <span className="text-[10px] px-1.5 py-0.2 bg-slate-200 rounded-full font-bold">
                {tab.count}
              </span>
            </button>
          ))}
        </div>
      </div>

      {/* RECENT WATCH-LIST INQUIRIES LIST */}
      <div className="space-y-4">
        <div className="flex justify-between items-center">
          <div className="flex items-center gap-2">
            <h3 className="text-sm font-black text-slate-800 uppercase tracking-tight">
              လတ်တလော စောင့်ကြည့်/စုံစမ်းမေးမြန်းမှု မှတ်တမ်းများ (Recent Watch-list Inquiries)
            </h3>
            <span className="bg-slate-200 text-slate-700 text-xs px-2 py-0.5 rounded-full font-bold">
              {filteredList.length}
            </span>
          </div>
          <span className="text-xs text-slate-400 font-bold hidden sm:inline">
            * "တိုက်စစ်မည်" ကို နှိပ်၍ တိုက်ဆိုင်မှုများအား အရာရှိမှ အတည်ပြု (Official Confirm) ဆောင်ရွက်နိုင်ပါသည်
          </span>
        </div>

        {filteredList.length === 0 ? (
          <div className="bg-white rounded-3xl p-12 text-center border border-slate-200 space-y-3">
            <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mx-auto text-slate-400">
              <ShieldAlert size={32} />
            </div>
            <div className="text-slate-600 font-bold text-sm">
              {searchQuery ? 'ရှာဖွေထားသော စကားလုံးနှင့် ကိုက်ညီသည့် မှတ်တမ်း မတွေ့ရှိပါ' : 'စောင့်ကြည့်စာရင်း မှတ်တမ်း မရှိသေးပါ'}
            </div>
            <p className="text-xs text-slate-400 max-w-md mx-auto">
              ဌာနဆိုင်ရာ မေးမြန်းစုံစမ်းစာများကို "မေးမြန်းစာရင်း အသစ်သွင်းရန်" ခလုတ်ဖြင့် စတင်ထည့်သွင်းနိုင်ပါသည်။
            </p>
            {!showForm && (
              <button
                onClick={() => setShowForm(true)}
                className="px-4 py-2 bg-indigo-600 text-white rounded-xl font-bold text-xs hover:bg-indigo-700 inline-flex items-center gap-1.5 cursor-pointer shadow-md"
              >
                <Plus size={15} /> မေးမြန်းစာရင်း အသစ်သွင်းရန်
              </button>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-4">
            {filteredList.map((rec) => {
              const isConfirmedMatch = rec.status === 'CONFIRMED_MATCH';
              const isClearedMismatch = rec.status === 'CLEARED_MISMATCH';
              const isFlagged = rec.status === 'FLAGGED';
              const isChecked = rec.status === 'CHECKED';

              return (
                <div
                  key={rec.id}
                  className={`bg-white rounded-3xl p-5 sm:p-6 border transition-all hover:shadow-lg space-y-4 relative ${
                    isConfirmedMatch
                      ? 'border-rose-400 ring-2 ring-rose-300 bg-rose-50/20'
                      : isClearedMismatch
                      ? 'border-blue-300 ring-1 ring-blue-100 bg-blue-50/15'
                      : isFlagged
                      ? 'border-amber-400 ring-2 ring-amber-200 bg-amber-50/20'
                      : isChecked
                      ? 'border-emerald-300'
                      : 'border-slate-200'
                  }`}
                >
                  {/* Top Bar: Letter No, Dept, Date, Status */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="font-mono font-black text-sm text-indigo-900 bg-indigo-50 px-2.5 py-1 rounded-xl border border-indigo-200">
                        {rec.letterNo}
                      </span>
                      <span className="font-bold text-xs text-slate-700 flex items-center gap-1">
                        <Building size={14} className="text-slate-400" />
                        {rec.requestDepartment}
                      </span>
                      <span className="text-xs text-slate-500 flex items-center gap-1">
                        <Calendar size={14} className="text-slate-400" />
                        {rec.letterDate}
                      </span>
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      {/* Status Badge */}
                      {isConfirmedMatch ? (
                        <span className="bg-rose-600 text-white font-black text-[11px] px-3 py-1 rounded-xl flex items-center gap-1.5 shadow-xs">
                          <ShieldAlert size={13} className="text-white" />
                          <span>အရာရှိ အတည်ပြုပြီး (CONFIRMED MATCH)</span>
                        </span>
                      ) : isClearedMismatch ? (
                        <span className="bg-blue-600 text-white font-black text-[11px] px-3 py-1 rounded-xl flex items-center gap-1.5 shadow-xs">
                          <ShieldCheck size={13} className="text-white" />
                          <span>မကိုက်ညီ/လူမှား ရှင်းလင်းပြီး (CLEARED)</span>
                        </span>
                      ) : isFlagged ? (
                        <span className="bg-amber-100 text-amber-900 border border-amber-300 text-[11px] font-black px-2.5 py-1 rounded-xl flex items-center gap-1.5 animate-pulse">
                          <AlertTriangle size={13} className="text-amber-700" />
                          <span>တွေ့ရှိမှုရှိ / အရာရှိ စိစစ်ဆဲ (MATCHES - PENDING CONFIRM)</span>
                        </span>
                      ) : isChecked ? (
                        <span className="bg-emerald-100 text-emerald-800 border border-emerald-300 text-[11px] font-black px-2.5 py-1 rounded-xl flex items-center gap-1.5">
                          <CheckCircle2 size={13} className="text-emerald-600" />
                          <span>သွားလာမှု မတွေ့ရှိပါ (NO RECORD)</span>
                        </span>
                      ) : (
                        <span className="bg-slate-100 text-slate-700 border border-slate-300 text-[11px] font-black px-2.5 py-1 rounded-xl flex items-center gap-1.5">
                          <Clock size={13} className="text-slate-500" />
                          <span>မစစ်ဆေးရသေး (PENDING CHECK)</span>
                        </span>
                      )}

                      {rec.lastCheckedAt && (
                        <span className="text-[10px] text-slate-400 font-mono hidden md:inline">
                          (စစ်ဆေးချိန်: {rec.lastCheckedAt})
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Reason / Subject */}
                  {rec.reasonDescription && (
                    <div className="text-xs text-slate-600 bg-slate-50 p-2.5 rounded-xl border border-slate-200/60 font-medium">
                      <span className="font-bold text-slate-700">အကြောင်းအရာ - </span>
                      {rec.reasonDescription}
                    </div>
                  )}

                  {/* Official Confirmation Banner if Confirmed/Cleared */}
                  {rec.officialConfirmation && (
                    <div className={`p-3 rounded-2xl border text-xs flex flex-col sm:flex-row sm:items-center justify-between gap-2 ${
                      isConfirmedMatch
                        ? 'bg-rose-50 border-rose-200 text-rose-900'
                        : isClearedMismatch
                        ? 'bg-blue-50 border-blue-200 text-blue-900'
                        : 'bg-slate-50 border-slate-200 text-slate-800'
                    }`}>
                      <div className="flex items-center gap-2">
                        <UserCheck size={16} className={isConfirmedMatch ? 'text-rose-600' : isClearedMismatch ? 'text-blue-600' : 'text-slate-600'} />
                        <div>
                          <span className="font-black">တာဝန်ခံအရာရှိ အတည်ပြုချက်: </span>
                          <span className="font-medium">{rec.officialConfirmation.remarks || (isConfirmedMatch ? 'ကိုက်ညီမှု မှန်ကန်ကြောင်း အတည်ပြုသည်' : 'လူမှားဖြစ်ကြောင်း ရှင်းလင်းအတည်ပြုသည်')}</span>
                        </div>
                      </div>
                      <div className="text-[11px] font-bold text-slate-600 flex items-center gap-1 self-end sm:self-auto">
                        <span>စိစစ်သူ: {rec.officialConfirmation.confirmedBy || '-'} ({rec.officialConfirmation.confirmedTitle || '-'})</span>
                        <span className="text-slate-400">| {toBurmeseFormattedDate(rec.officialConfirmation.confirmedAt)}</span>
                      </div>
                    </div>
                  )}

                  {/* Persons Cards Grid */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {rec.persons.map((person, pIdx) => {
                      const pMatch = rec.matchResults?.[person.id];
                      const matchedCount = pMatch?.matchedRecords?.length || 0;

                      return (
                        <div
                          key={person.id}
                          className="bg-slate-50/70 p-3.5 rounded-2xl border border-slate-200/80 space-y-2 hover:border-indigo-200 transition-colors"
                        >
                          <div className="flex justify-between items-start">
                            <div>
                              <div className="font-black text-slate-900 text-xs uppercase flex items-center gap-1.5">
                                <User size={13} className="text-indigo-600" />
                                {person.name}
                              </div>
                              <div className="font-mono text-xs font-bold text-indigo-700">
                                {person.passport || '-'} | {person.nationality || '-'}
                              </div>
                            </div>
                            <span className="text-[10px] bg-slate-200 font-bold px-1.5 py-0.5 rounded text-slate-700">
                              {pIdx + 1}/{rec.persons.length}
                            </span>
                          </div>

                          <div className="text-[11px] text-slate-500 space-y-0.5 border-t border-slate-200/60 pt-1.5">
                            {person.fatherName && <div><span className="font-bold text-slate-600">အဖ:</span> {person.fatherName}</div>}
                            {person.dob && <div><span className="font-bold text-slate-600">မွေးသက္ကရာဇ်:</span> {person.dob}</div>}
                            {person.address && <div className="truncate"><span className="font-bold text-slate-600">လိပ်စာ:</span> {person.address}</div>}
                            {person.idNumber && <div><span className="font-bold text-slate-600">ID:</span> {person.idNumber}</div>}
                          </div>

                          {/* Quick Match Status for Person */}
                          {pMatch && (
                            <div className={`text-[10px] font-black p-1.5 rounded-lg border flex items-center justify-between ${
                              matchedCount > 0
                                ? 'bg-rose-50 border-rose-200 text-rose-700'
                                : 'bg-emerald-50 border-emerald-200 text-emerald-700'
                            }`}>
                              <span>{pMatch.summaryStatus}</span>
                              {matchedCount > 0 && (
                                <span className="bg-rose-600 text-white px-1.5 py-0.2 rounded font-mono">
                                  {matchedCount} logs
                                </span>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  {/* Bottom Action Controls for Entry */}
                  <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
                    <div className="text-xs text-slate-400 font-medium">
                      တာဝန်ခံ: <span className="font-bold text-slate-700">{rec.officerName || '-'}</span> ({rec.officerTitle || '-'})
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      {/* Button: တိုက်စစ်မည် / စိစစ်မည် */}
                      <button
                        onClick={() => performCrossReferenceCheck(rec)}
                        className="px-3.5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs shadow-md flex items-center gap-1.5 cursor-pointer transition-all hover:scale-[1.02]"
                        title="စနစ်တစ်ခုလုံးရှိ Flight History နှင့် Movement Log များတွင် တိုက်ဆိုင်စစ်ဆေးပြီး အရာရှိ အတည်ပြုချက် ထည့်သွင်းမည်"
                      >
                        <RefreshCw size={14} className={rec.status === 'PENDING' ? 'animate-spin' : ''} />
                        <span>တိုက်စစ်/အတည်ပြုမည်</span>
                      </button>

                      {/* Button: အစီရင်ခံစာ ထုတ်မည် / ပုံနှိပ်မည် */}
                      <button
                        onClick={() => openReportModal(rec)}
                        className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white font-black text-xs shadow-sm flex items-center gap-1.5 cursor-pointer transition-all"
                        title="မေးမြန်းမှုအတွက် ပုံသေဖောမက် အစီရင်ခံစာ ထုတ်ယူမည်"
                      >
                        <FileText size={14} className="text-amber-400" />
                        <span>အစီရင်ခံစာ ထုတ်မည်</span>
                      </button>

                      {/* Button: Edit */}
                      {!isViewer && (
                        <button
                          onClick={() => startEdit(rec)}
                          className="p-2 rounded-xl text-slate-600 hover:text-indigo-600 hover:bg-indigo-50 border border-slate-200 transition-colors cursor-pointer"
                          title="အချက်အလက် ပြင်ဆင်မည်"
                        >
                          <Edit size={14} />
                        </button>
                      )}

                      {/* Button: Delete */}
                      {!isViewer && (
                        <button
                          onClick={() => handleDeleteEntry(rec.id)}
                          className="p-2 rounded-xl text-slate-400 hover:text-rose-600 hover:bg-rose-50 border border-slate-200 transition-colors cursor-pointer"
                          title="ဖျက်သိမ်းမည်"
                        >
                          <Trash2 size={14} />
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* CROSS-REFERENCE CHECK RESULT & OFFICIAL CONFIRMATION - FULL SCREEN WORKSPACE */}
      {activeCheckRecord && (
        <div className="fixed inset-0 z-50 bg-slate-950 text-slate-100 flex flex-col overflow-hidden animate-in fade-in duration-200">
          {/* Top Full-Screen Navigation Bar */}
          <div className="bg-slate-900 border-b border-slate-800 px-4 sm:px-8 py-3.5 flex flex-wrap justify-between items-center gap-3 shrink-0 shadow-lg">
            <div className="flex items-center gap-3">
              <button
                onClick={() => setActiveCheckRecord(null)}
                className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors cursor-pointer flex items-center gap-1.5 text-xs font-bold"
                title="စောင့်ကြည့်စာရင်းသို့ ပြန်သွားမည်"
              >
                <ArrowLeft size={16} />
                <span className="hidden sm:inline">နောက်သို့</span>
              </button>

              <div className="h-6 w-px bg-slate-700 hidden sm:block" />

              <div className="flex items-center gap-2.5">
                <div className="p-2 rounded-xl bg-indigo-600/30 text-indigo-400 border border-indigo-500/30">
                  <ShieldCheck size={22} />
                </div>
                <div>
                  <h3 className="font-black text-sm sm:text-base text-white tracking-wide flex items-center gap-2">
                    <span>စနစ်တွင်း တိုက်ဆိုင်စစ်ဆေးမှုနှင့် တာဝန်ခံအရာရှိ အတည်ပြုချက်</span>
                    <span className="text-[11px] font-bold text-indigo-300 bg-indigo-950/80 border border-indigo-700/50 px-2 py-0.5 rounded-md hidden md:inline">
                      FULL SCREEN VERIFICATION
                    </span>
                  </h3>
                  <div className="flex items-center gap-2 text-xs text-slate-400 font-medium">
                    <span className="text-slate-200 font-bold">စာအမှတ်: {activeCheckRecord.letterNo}</span>
                    <span>•</span>
                    <span>{activeCheckRecord.requestDepartment}</span>
                    <span>•</span>
                    <span>မေးမြန်းရက်: {toBurmeseFormattedDate(activeCheckRecord.letterDate)}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => {
                  const target = activeCheckRecord;
                  setActiveCheckRecord(null);
                  openReportModal(target);
                }}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-xs flex items-center gap-1.5 shadow-md cursor-pointer transition-all hover:scale-[1.02]"
              >
                <FileText size={15} />
                <span>တရားဝင် အစီရင်ခံစာ ထုတ်မည်</span>
              </button>

              <button
                onClick={() => setActiveCheckRecord(null)}
                className="p-2 rounded-xl bg-slate-800 hover:bg-rose-900/40 text-slate-400 hover:text-rose-300 border border-slate-700 transition-colors cursor-pointer"
                title="ပိတ်မည်"
              >
                <X size={18} />
              </button>
            </div>
          </div>

          {/* Full Screen Scrollable Body Container */}
          <div className="flex-1 overflow-y-auto bg-slate-950 p-4 sm:p-6 lg:p-8 space-y-6 select-text">
            <div className="max-w-6xl mx-auto space-y-6">
              
              {/* Inquiry Overview Card */}
              <div className="bg-slate-900/80 backdrop-blur-md rounded-2xl p-5 border border-slate-800 shadow-md flex flex-wrap items-center justify-between gap-4">
                <div className="space-y-1">
                  <div className="text-[11px] font-black uppercase text-indigo-400 tracking-wider">
                    မေးမြန်းမှုဆိုင်ရာ အချက်အလက်
                  </div>
                  <div className="text-sm font-black text-white flex items-center gap-2">
                    <span>{activeCheckRecord.letterNo}</span>
                    <span className="text-slate-400 font-normal">({activeCheckRecord.requestDepartment})</span>
                  </div>
                  <div className="text-xs text-slate-400">
                    စောင့်ကြည့်စစ်ဆေးရမည့် ပုဂ္ဂိုလ်ဦးရေ: <b className="text-white">{activeCheckRecord.persons.length} ဦး</b> | နောက်ဆုံးစစ်ဆေးချိန်: <span className="font-mono text-slate-300">{activeCheckRecord.lastCheckedAt || 'ယခုစစ်ဆေးဆဲ'}</span>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  {activeCheckRecord.status === 'CONFIRMED_MATCH' && (
                    <span className="bg-rose-600 text-white font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 shadow-md">
                      <ShieldAlert size={15} />
                      အရာရှိ အတည်ပြုပြီး (CONFIRMED MATCH)
                    </span>
                  )}
                  {activeCheckRecord.status === 'CLEARED_MISMATCH' && (
                    <span className="bg-blue-600 text-white font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 shadow-md">
                      <ShieldCheck size={15} />
                      မကိုက်ညီ/လူမှား ရှင်းလင်းပြီး (CLEARED)
                    </span>
                  )}
                  {activeCheckRecord.status === 'FLAGGED' && (
                    <span className="bg-amber-600 text-white font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 shadow-md">
                      <AlertTriangle size={15} />
                      တွေ့ရှိမှုရှိ / အရာရှိ စိစစ်ဆဲ
                    </span>
                  )}
                  {activeCheckRecord.status === 'CHECKED' && (
                    <span className="bg-emerald-600 text-white font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 shadow-md">
                      <CheckCircle2 size={15} />
                      သွားလာမှု မတွေ့ရှိပါ
                    </span>
                  )}
                </div>
              </div>

              {/* PERSONS FINDINGS CARDS */}
              {activeCheckRecord.persons.map((person, idx) => {
                const pResult = activeCheckRecord.matchResults?.[person.id];
                const matchedRecs: (ImmRecord & { matchReasons?: string[]; isDirect?: boolean })[] = pResult?.matchedRecords || [];
                const chkEntries: any[] = pResult?.checkingEntries || [];

                return (
                  <div key={person.id} className="bg-slate-900 rounded-2xl border border-slate-800 p-5 sm:p-6 space-y-4 shadow-md">
                    {/* Person Header */}
                    <div className="flex justify-between items-center border-b border-slate-800 pb-3 flex-wrap gap-2">
                      <div className="flex items-center gap-2.5 flex-wrap">
                        <span className="w-6 h-6 rounded-full bg-indigo-600 text-white font-black text-xs flex items-center justify-center shadow-xs">
                          {idx + 1}
                        </span>
                        <span className="font-black text-base text-white uppercase tracking-wide">
                          {person.name}
                        </span>
                        <span className="font-mono font-bold text-xs text-indigo-300 bg-indigo-950/80 px-2.5 py-1 rounded-lg border border-indigo-800/60">
                          {person.passport || 'NO PP'}
                        </span>
                        <span className="font-bold text-xs text-slate-300 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">
                          {person.nationality}
                        </span>
                        {person.idNumber && (
                          <span className="font-bold text-xs text-slate-300 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">
                            ID: {person.idNumber}
                          </span>
                        )}
                        {person.fatherName && (
                          <span className="text-xs text-slate-400 bg-slate-800/80 px-2.5 py-1 rounded-lg border border-slate-700">
                            အဖ: <b className="text-slate-200">{person.fatherName}</b>
                          </span>
                        )}
                        {person.address && (
                          <span className="text-xs text-slate-400 bg-slate-800/80 px-2.5 py-1 rounded-lg border border-slate-700">
                            လိပ်စာ: <span className="text-slate-300">{person.address}</span>
                          </span>
                        )}
                      </div>

                      <div className="flex items-center gap-2">
                        {pResult?.hasDirectMatch && (
                          <span className="px-2.5 py-1 rounded-lg bg-rose-600 text-white font-black text-[11px] shadow-xs">
                            တိုက်ရိုက်ကိုက်ညီမှု
                          </span>
                        )}
                        {pResult?.hasRelatedMatch && (
                          <span className="px-2.5 py-1 rounded-lg bg-amber-600 text-white font-black text-[11px] shadow-xs">
                            မိဘ/ID/လိပ်စာ ဆက်စပ်မှု
                          </span>
                        )}
                        <span className={`px-3 py-1 rounded-xl font-black text-xs border ${
                          (matchedRecs.length > 0 || chkEntries.length > 0)
                            ? 'bg-rose-950/80 text-rose-300 border-rose-700'
                            : 'bg-emerald-950/80 text-emerald-300 border-emerald-700'
                        }`}>
                          {(matchedRecs.length > 0 || chkEntries.length > 0)
                            ? `တွေ့ရှိမှု ${matchedRecs.length + chkEntries.length} ခု`
                            : 'မှတ်တမ်း မရှိပါ'}
                        </span>
                      </div>
                    </div>

                    {/* Matched Flight/Movement Records Table */}
                    {matchedRecs.length > 0 && (
                      <div className="space-y-2">
                        <div className="font-bold text-slate-300 text-xs flex items-center gap-1.5">
                          <Plane size={14} className="text-indigo-400" />
                          <span>စနစ်တွင်း ဝင်ရောက်/ထွက်ခွာမှု မှတ်တမ်းများ ({matchedRecs.length} ခု):</span>
                        </div>
                        <div className="overflow-x-auto border border-slate-700 rounded-xl bg-slate-950 shadow-xs">
                          <table className="w-full text-center text-xs border-collapse">
                            <thead className="bg-slate-800 text-slate-200 font-bold text-[11px]">
                              <tr>
                                <th className="p-2.5 border-r border-slate-700">စဉ်</th>
                                <th className="p-2.5 border-r border-slate-700">ရက်စွဲ/အချိန်</th>
                                <th className="p-2.5 border-r border-slate-700">အဝင်/အထွက်</th>
                                <th className="p-2.5 border-r border-slate-700">စနစ်တွင်း ခရီးသည်အမည်</th>
                                <th className="p-2.5 border-r border-slate-700">တိုက်ဆိုင်တွေ့ရှိမှု အကြောင်းရင်း</th>
                                <th className="p-2.5 border-r border-slate-700">လေယာဉ်/ယာဉ်</th>
                                <th className="p-2.5 border-r border-slate-700">ဗီဇာ</th>
                                <th className="p-2.5 border-r border-slate-700">နေခွင့်ကာလ</th>
                                <th className="p-2.5 border-r border-slate-700">တည်းခိုရာလိပ်စာ</th>
                                <th className="p-2.5">တာဝန်ခံ</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-800">
                              {matchedRecs.map((r, rIdx) => (
                                <tr key={r.id} className="hover:bg-slate-900/60 transition-colors">
                                  <td className="p-2.5 border-r border-slate-800 font-bold text-slate-400">{rIdx + 1}</td>
                                  <td className="p-2.5 border-r border-slate-800 font-mono text-[11px] text-slate-300">{r.timestamp}</td>
                                  <td className="p-2.5 border-r border-slate-800 font-black">
                                    <span className={`px-2 py-0.5 rounded text-[10px] text-white ${
                                      r.mode === 'IN' ? 'bg-emerald-600' : 'bg-orange-600'
                                    }`}>
                                      {r.mode}
                                    </span>
                                  </td>
                                  <td className="p-2.5 border-r border-slate-800 font-black text-white text-left">
                                    {r.fullname}
                                    {r.passport && <span className="block font-mono text-[10px] text-indigo-300">{r.passport}</span>}
                                  </td>
                                  <td className="p-2.5 border-r border-slate-800 text-left">
                                    {(r.matchReasons && r.matchReasons.length > 0) ? (
                                      <div className="space-y-1">
                                        {r.matchReasons.map((reason, rsIdx) => (
                                          <span
                                            key={rsIdx}
                                            className="inline-block px-2 py-0.5 rounded text-[10px] font-bold mr-1 mb-0.5 bg-rose-950 text-rose-300 border border-rose-800"
                                          >
                                            {reason}
                                          </span>
                                        ))}
                                      </div>
                                    ) : (
                                      <span className="text-slate-500">-</span>
                                    )}
                                  </td>
                                  <td className="p-2.5 border-r border-slate-800 font-mono text-slate-300">{r.vehicleInfo || '-'}</td>
                                  <td className="p-2.5 border-r border-slate-800 text-slate-300">{r.visaType || '-'}</td>
                                  <td className="p-2.5 border-r border-slate-800 text-slate-300">{r.totalDays ? `${r.totalDays} ရက်` : (r.stayFrom ? `${r.stayFrom} ~ ${r.stayTo}` : '-')}</td>
                                  <td className="p-2.5 border-r border-slate-800 text-slate-300 max-w-[150px] truncate text-left" title={r.address}>
                                    {r.address || '-'}
                                  </td>
                                  <td className="p-2.5 text-slate-400 text-[11px]">{r.officialName || '-'}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {/* Matched Checking Records Table */}
                    {chkEntries.length > 0 && (
                      <div className="space-y-2">
                        <div className="font-bold text-amber-300 text-xs flex items-center gap-1.5">
                          <FileText size={14} className="text-amber-400" />
                          <span>တည်းခိုနေထိုင်မှု / ဧည့်စာရင်းစစ်ဆေးမှု မှတ်တမ်းများ ({chkEntries.length} ခု):</span>
                        </div>
                        <div className="overflow-x-auto border border-slate-700 rounded-xl bg-slate-950 shadow-xs">
                          <table className="w-full text-center text-xs border-collapse">
                            <thead className="bg-slate-800 text-slate-200 font-bold text-[11px]">
                              <tr>
                                <th className="p-2.5 border-r border-slate-700">စဉ်</th>
                                <th className="p-2.5 border-r border-slate-700">စစ်ဆေးသည့်ရက်</th>
                                <th className="p-2.5 border-r border-slate-700">အမည် / နိုင်ငံသား</th>
                                <th className="p-2.5 border-r border-slate-700">ပတ်စ်ပို့ / ID</th>
                                <th className="p-2.5 border-r border-slate-700">တည်းခိုရာ ဟိုတယ်/လိပ်စာ</th>
                                <th className="p-2.5 border-r border-slate-700">ရလဒ် / မှတ်ချက်</th>
                                <th className="p-2.5">စစ်ဆေးသူ</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-800">
                              {chkEntries.map((ce: any, ceIdx: number) => (
                                <tr key={ce.id || ceIdx} className="hover:bg-slate-900/60 transition-colors">
                                  <td className="p-2.5 border-r border-slate-800 font-bold text-slate-400">{ceIdx + 1}</td>
                                  <td className="p-2.5 border-r border-slate-800 font-mono text-[11px] text-slate-300">{toBurmeseFormattedDate(ce.checkDate || ce.date || '')}</td>
                                  <td className="p-2.5 border-r border-slate-800 font-bold text-white text-left">
                                    {ce.guestName || ce.name || '-'}
                                    <span className="block text-[10px] text-slate-400">{ce.nationality || '-'}</span>
                                  </td>
                                  <td className="p-2.5 border-r border-slate-800 font-mono text-indigo-300">{ce.passport || ce.idNumber || '-'}</td>
                                  <td className="p-2.5 border-r border-slate-800 text-slate-300 text-left">{ce.confirmedAddress || ce.originalAddress || ce.hotelName || ce.address || '-'}</td>
                                  <td className="p-2.5 border-r border-slate-800 text-slate-300">{ce.matchReason || ce.status || ce.remark || '-'}</td>
                                  <td className="p-2.5 text-slate-400 text-[11px]">{ce.officerName || ce.officer || '-'}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {/* No Matches Found Notice */}
                    {matchedRecs.length === 0 && chkEntries.length === 0 && (
                      <div className="p-4 bg-emerald-950/40 rounded-xl border border-emerald-800/60 text-emerald-300 flex items-center gap-2.5">
                        <CheckCircle2 size={18} className="text-emerald-400 shrink-0" />
                        <div className="text-xs">
                          <b>မှတ်တမ်းမတွေ့ရှိပါ -</b> စနစ်တွင်းရှိ Flight History, Movement Logs နှင့် ဧည့်စာရင်း မှတ်တမ်းများတွင် အဆိုပါပုဂ္ဂိုလ်၏ အချက်အလက်များနှင့် တိုက်ဆိုင်တွေ့ရှိမှု မရှိပါ။
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}

              {/* OFFICIAL VERIFICATION & AUDIT DECISION CARD (EXPANDED & PROMINENT) */}
              <div className="bg-slate-900 text-white p-6 sm:p-8 rounded-3xl border-2 border-indigo-500/40 shadow-2xl space-y-6">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800 pb-4">
                  <div className="flex items-center gap-3">
                    <div className="p-2.5 rounded-2xl bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                      <UserCheck size={24} />
                    </div>
                    <div>
                      <h4 className="font-black text-base sm:text-lg uppercase tracking-wide text-white flex items-center gap-2">
                        <span>တာဝန်ခံအရာရှိ၏ စိစစ်အတည်ပြုချက် (Official Verification & Decision)</span>
                      </h4>
                      <p className="text-xs text-slate-400">
                        စနစ်တွင်း တိုက်ဆိုင်တွေ့ရှိမှုများအား မြေပြင်/အထောက်အထားများနှင့် ချိန်ညှိ၍ အတည်ပြုချက် သတ်မှတ်ပါ
                      </p>
                    </div>
                  </div>

                  {activeCheckRecord.officialConfirmation && (
                    <span className="text-xs bg-slate-800 text-indigo-300 px-3.5 py-1.5 rounded-full border border-indigo-700/50 font-mono self-start sm:self-auto">
                      အတည်ပြုပြီး: {toBurmeseFormattedDate(activeCheckRecord.officialConfirmation.confirmedAt)}
                    </span>
                  )}
                </div>

                {/* 3 Decision Radio Cards */}
                <div>
                  <label className="block text-xs font-black uppercase text-indigo-300 tracking-wider mb-2.5">
                    ၁။ စိစစ်ဆုံးဖြတ်ချက် ရွေးချယ်ပါ (Select Decision) *
                  </label>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
                    {/* Option 1: CONFIRMED_MATCH */}
                    <label
                      onClick={() => setConfirmDecision('CONFIRMED_MATCH')}
                      className={`p-4 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between space-y-2.5 ${
                        confirmDecision === 'CONFIRMED_MATCH'
                          ? 'bg-rose-950/80 border-rose-500 text-rose-200 ring-2 ring-rose-500/40 shadow-lg'
                          : 'bg-slate-800/60 border-slate-700 hover:border-slate-600 text-slate-300'
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-black text-sm uppercase flex items-center gap-2 text-rose-400">
                          <ShieldAlert size={18} />
                          အမှန်တကယ် ကိုက်ညီသည်
                        </span>
                        <input
                          type="radio"
                          name="confirmDecision"
                          checked={confirmDecision === 'CONFIRMED_MATCH'}
                          onChange={() => setConfirmDecision('CONFIRMED_MATCH')}
                          className="accent-rose-500 w-4 h-4 cursor-pointer"
                        />
                      </div>
                      <p className="text-xs text-slate-300 leading-relaxed">
                        စနစ်တွင်း တွေ့ရှိမှုသည် မေးမြန်းထားသော ပုဂ္ဂိုလ်နှင့် အမှန်တကယ် ကိုက်ညီပါသည် <b>(CONFIRMED MATCH)</b>။
                      </p>
                    </label>

                    {/* Option 2: CLEARED_MISMATCH */}
                    <label
                      onClick={() => setConfirmDecision('CLEARED_MISMATCH')}
                      className={`p-4 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between space-y-2.5 ${
                        confirmDecision === 'CLEARED_MISMATCH'
                          ? 'bg-blue-950/80 border-blue-500 text-blue-200 ring-2 ring-blue-500/40 shadow-lg'
                          : 'bg-slate-800/60 border-slate-700 hover:border-slate-600 text-slate-300'
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-black text-sm uppercase flex items-center gap-2 text-blue-400">
                          <ShieldCheck size={18} />
                          မကိုက်ညီ / လူမှားဖြစ်သည်
                        </span>
                        <input
                          type="radio"
                          name="confirmDecision"
                          checked={confirmDecision === 'CLEARED_MISMATCH'}
                          onChange={() => setConfirmDecision('CLEARED_MISMATCH')}
                          className="accent-blue-500 w-4 h-4 cursor-pointer"
                        />
                      </div>
                      <p className="text-xs text-slate-300 leading-relaxed">
                        အမည်တူ/လိပ်စာတူ ဆက်စပ်ရုံသာဖြစ်ပြီး စောင့်ကြည့်ခံ ပုဂ္ဂိုလ်နှင့် မကိုက်ညီပါ <b>(CLEARED / MISMATCH)</b>။
                      </p>
                    </label>

                    {/* Option 3: PENDING */}
                    <label
                      onClick={() => setConfirmDecision('PENDING')}
                      className={`p-4 rounded-2xl border-2 cursor-pointer transition-all flex flex-col justify-between space-y-2.5 ${
                        confirmDecision === 'PENDING'
                          ? 'bg-amber-950/80 border-amber-500 text-amber-200 ring-2 ring-amber-500/40 shadow-lg'
                          : 'bg-slate-800/60 border-slate-700 hover:border-slate-600 text-slate-300'
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-black text-sm uppercase flex items-center gap-2 text-amber-400">
                          <Clock size={18} />
                          စိစစ်ဆဲ အခြေအနေ
                        </span>
                        <input
                          type="radio"
                          name="confirmDecision"
                          checked={confirmDecision === 'PENDING'}
                          onChange={() => setConfirmDecision('PENDING')}
                          className="accent-amber-500 w-4 h-4 cursor-pointer"
                        />
                      </div>
                      <p className="text-xs text-slate-300 leading-relaxed">
                        အချက်အလက် ထပ်မံစုံစမ်းရန် လိုအပ်နေသေးသဖြင့် ဆိုင်းငံ့ထားဆဲ ဖြစ်ပါသည် <b>(PENDING REVIEW)</b>။
                      </p>
                    </label>
                  </div>
                </div>

                {/* Officer Meta & Remarks Input */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2">
                  <div>
                    <label className="block text-xs font-bold text-slate-300 mb-1.5">
                      ၂။ စိစစ်အတည်ပြုသူ အရာရှိအမည် *
                    </label>
                    <input
                      type="text"
                      value={confirmOfficerName}
                      onChange={(e) => setConfirmOfficerName(e.target.value)}
                      placeholder="ဥပမာ - ဦးအောင်မျိုးကျော်"
                      className="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-300 mb-1.5">
                      ၃။ ရာထူး / တာဝန်
                    </label>
                    <input
                      type="text"
                      value={confirmOfficerTitle}
                      onChange={(e) => setConfirmOfficerTitle(e.target.value)}
                      placeholder="ဥပမာ - ဒုတိယလဝကမှူး / လဝကမှူး"
                      className="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-xs text-white focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                    />
                  </div>

                  <div className="sm:col-span-2 space-y-2.5">
                    <div className="flex justify-between items-center flex-wrap gap-2">
                      <label className="text-xs font-bold text-slate-300">
                        ၄။ စိစစ်တွေ့ရှိချက် အသေးစိတ် မှတ်ချက် (Verification Remarks)
                      </label>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <span className="text-[11px] text-slate-400 font-bold">အမြန်ရွေးရန်:</span>
                        <button
                          type="button"
                          onClick={() => setConfirmRemarks('အမည်၊ ပတ်စ်ပို့အမှတ်နှင့် ဓာတ်ပုံပါ အချက်အလက်များ ကိုက်ညီမှုရှိကြောင်း စိစစ်တွေ့ရှိရပါသည်')}
                          className="text-[11px] bg-slate-800 hover:bg-slate-700 border border-slate-700 px-2.5 py-1 rounded-lg text-slate-300 hover:text-white cursor-pointer transition-colors"
                        >
                          ✓ ကိုက်ညီမှုရှိ
                        </button>
                        <button
                          type="button"
                          onClick={() => setConfirmRemarks('အမည်တူရုံသာဖြစ်ပြီး မိဘအမည်နှင့် ဇာတိ မတူညီသဖြင့် စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်မဟုတ်ပါ')}
                          className="text-[11px] bg-slate-800 hover:bg-slate-700 border border-slate-700 px-2.5 py-1 rounded-lg text-slate-300 hover:text-white cursor-pointer transition-colors"
                        >
                          ✕ အမည်တူလူမှားဖြစ်
                        </button>
                        <button
                          type="button"
                          onClick={() => setConfirmRemarks('အချက်အလက် မပြည့်စုံသေးသဖြင့် သက်ဆိုင်ရာ ဌာနသို့ ထပ်မံမေးမြန်းရန် လိုအပ်ပါသည်')}
                          className="text-[11px] bg-slate-800 hover:bg-slate-700 border border-slate-700 px-2.5 py-1 rounded-lg text-slate-300 hover:text-white cursor-pointer transition-colors"
                        >
                          ⋯ ထပ်မံစိစစ်ရန်လိုအပ်
                        </button>
                      </div>
                    </div>
                    <textarea
                      rows={3}
                      value={confirmRemarks}
                      onChange={(e) => setConfirmRemarks(e.target.value)}
                      placeholder="အရာရှိ၏ စိစစ်တွေ့ရှိချက်နှင့် ဆုံးဖြတ်ချက် မှတ်ချက်ကို ရေးသွင်းပါ..."
                      className="w-full bg-slate-800 border border-slate-700 rounded-xl p-3 text-xs text-white focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                    />
                  </div>
                </div>

                {/* Save Confirmation Button */}
                <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-4 border-t border-slate-800">
                  <div className="text-xs text-slate-400">
                    * အတည်ပြုချက် သိမ်းဆည်းပြီးပါက စောင့်ကြည့်စာရင်းနှင့် အစီရင်ခံစာများတွင် ချက်ချင်း အကျိုးသက်ရောက်မည်ဖြစ်ပါသည်
                  </div>

                  <button
                    type="button"
                    onClick={() => handleSaveOfficialConfirmation(
                      activeCheckRecord.id,
                      confirmDecision,
                      confirmOfficerName,
                      confirmOfficerTitle,
                      confirmRemarks
                    )}
                    className="w-full sm:w-auto px-6 py-3 bg-gradient-to-r from-indigo-600 via-purple-600 to-rose-600 hover:from-indigo-500 hover:to-rose-500 text-white rounded-xl font-black text-xs sm:text-sm flex items-center justify-center gap-2 shadow-xl cursor-pointer transition-all hover:scale-[1.02] active:scale-[0.98]"
                  >
                    <CheckCircle2 size={18} />
                    <span>တာဝန်ခံအရာရှိ အတည်ပြုချက် သိမ်းဆည်းမည် (Save Confirmation)</span>
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Bottom Fixed Action Bar */}
          <div className="p-3.5 bg-slate-900 border-t border-slate-800 flex justify-between items-center px-4 sm:px-8 shrink-0">
            <button
              onClick={() => setActiveCheckRecord(null)}
              className="px-4 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-xl text-slate-300 font-bold text-xs cursor-pointer flex items-center gap-1.5"
            >
              <ArrowLeft size={14} />
              <span>စောင့်ကြည့်စာရင်းသို့ ပြန်သွားမည် (Back)</span>
            </button>

            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={() => handleSaveOfficialConfirmation(
                  activeCheckRecord.id,
                  confirmDecision,
                  confirmOfficerName,
                  confirmOfficerTitle,
                  confirmRemarks
                )}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl font-black text-xs flex items-center gap-1.5 shadow-md cursor-pointer transition-all hover:scale-[1.02]"
              >
                <CheckCircle2 size={15} />
                <span>အတည်ပြုချက် သိမ်းဆည်းမည်</span>
              </button>

              <button
                onClick={() => {
                  const target = activeCheckRecord;
                  setActiveCheckRecord(null);
                  openReportModal(target);
                }}
                className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-white rounded-xl font-bold text-xs flex items-center gap-1.5 border border-slate-700 cursor-pointer"
              >
                <FileText size={15} className="text-amber-400" />
                <span>တရားဝင် အစီရင်ခံစာ ထုတ်မည်</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* OFFICIAL FIXED FORMAT BURMESE REPORT MODAL */}
      {activeReportRecord && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-5xl max-h-[95vh] rounded-3xl shadow-2xl overflow-hidden flex flex-col">
            {/* Header controls */}
            <div className="p-4 bg-slate-900 text-white flex flex-wrap justify-between items-center gap-3">
              <div className="flex items-center gap-2">
                <FileText size={18} className="text-amber-400" />
                <span className="font-black text-sm uppercase">
                  ပုံသေဖောမက် အစီရင်ခံစာ ပုံနှိပ်ထုတ်ယူခြင်း (Official Finding Report)
                </span>
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                <button
                  onClick={() => copyReportText(activeReportRecord)}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold flex items-center gap-1 border border-slate-700 cursor-pointer"
                  title="Copy formatted text to Clipboard"
                >
                  <Copy size={13} />
                  <span>Copy စာသား</span>
                </button>

                <button
                  onClick={triggerPrintReport}
                  className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md cursor-pointer"
                >
                  <Printer size={14} />
                  <span>ပုံနှိပ်မည် (Print Report)</span>
                </button>

                <button
                  onClick={() => setActiveReportRecord(null)}
                  className="p-1.5 rounded-xl hover:bg-white/10 text-slate-400 hover:text-white"
                >
                  <X size={18} />
                </button>
              </div>
            </div>

            {/* Officer & Report Settings Bar */}
            <div className="p-3 bg-slate-100 border-b flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-3 flex-wrap">
                <div>
                  <label className="font-bold text-slate-600 mr-1.5">အကြောင်းပြန်စာအမှတ်:</label>
                  <input
                    type="text"
                    value={reportLetterNo}
                    onChange={(e) => setReportLetterNo(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 font-bold text-xs"
                  />
                </div>
                <div>
                  <label className="font-bold text-slate-600 mr-1.5">ရက်စွဲ:</label>
                  <input
                    type="date"
                    value={reportDate}
                    onChange={(e) => setReportDate(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 font-bold text-xs"
                  />
                </div>
              </div>

              <div className="flex items-center gap-3 flex-wrap">
                <div>
                  <label className="font-bold text-slate-600 mr-1.5">တာဝန်ခံအရာရှိ:</label>
                  <input
                    type="text"
                    value={reportOfficerName}
                    onChange={(e) => setReportOfficerName(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 font-bold text-xs"
                  />
                </div>
                <div>
                  <label className="font-bold text-slate-600 mr-1.5">ရာထူး:</label>
                  <input
                    type="text"
                    value={reportOfficerTitle}
                    onChange={(e) => setReportOfficerTitle(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 font-bold text-xs"
                  />
                </div>
              </div>
            </div>

            {/* A4 PRINTABLE CANVAS SHEET */}
            <div className="p-6 overflow-y-auto flex justify-center bg-slate-200/70 flex-1">
              <div
                ref={reportPrintRef}
                className="bg-white text-black p-[0.65in] shadow-2xl rounded-sm w-[210mm] min-w-[210mm] max-w-[210mm] border border-slate-300 min-h-[297mm] relative font-pyidaungsu text-[13px] leading-relaxed select-text box-border"
                style={{ backgroundColor: '#ffffff', color: '#000000' }}
              >
                {/* Official Letterhead */}
                <div className="text-center space-y-1 mb-6 border-b-2 border-black pb-4">
                  <div className="font-extrabold text-[15px] uppercase tracking-wide text-black">
                    ပြည်ထောင်စုသမ္မတမြန်မာနိုင်ငံတော်
                  </div>
                  <div className="font-bold text-[14px] text-black">
                    လူဝင်မှုကြီးကြပ်ရေးနှင့် ပြည်သူ့အင်အားဝန်ကြီးဌာန
                  </div>
                  <div className="font-bold text-[13px] text-black">
                    လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်)
                  </div>
                </div>

                {/* Letter Meta & Date */}
                <div className="flex justify-between items-start text-[13px] font-bold mb-4">
                  <div>
                    <div>စာအမှတ်၊ {reportLetterNo}</div>
                    <div className="mt-1">သို့</div>
                    <div className="pl-4 font-normal mt-0.5">{activeReportRecord.requestDepartment}</div>
                  </div>
                  <div className="text-right">
                    <div>ရက်စွဲ၊ {toBurmeseFormattedDate(reportDate)}</div>
                    <div className="mt-1 text-slate-700 font-mono text-[11px]">မြိတ်မြို့</div>
                  </div>
                </div>

                {/* Subject Line */}
                <div className="font-black text-[13px] leading-snug my-4 p-2.5 bg-slate-50 border border-black/40">
                  အကြောင်းအရာ။ <span className="underline font-bold">
                    {toBurmeseFormattedDate(activeReportRecord.letterDate)} ရက်စွဲပါ စာအမှတ် {toBurmeseDigits(activeReportRecord.letterNo)} ဖြင့် မေးမြန်းထားသည့် {activeReportRecord.reasonDescription && activeReportRecord.reasonDescription.trim() ? `"${activeReportRecord.reasonDescription.trim()}" ကိစ္စနှင့် စပ်လျဉ်း၍ ` : ''}({activeReportRecord.persons[0]?.nationality || ''}) နိုင်ငံသား {activeReportRecord.persons[0]?.name || ''} ({activeReportRecord.persons[0]?.passport || '-'}){activeReportRecord.persons.length > 1 ? ` ပါ (${toBurmeseDigits(activeReportRecord.persons.length)}) ဦး` : ''} အား စနစ်တွင်း စစ်ဆေးတွေ့ရှိမှု အစီရင်ခံ တင်ပြခြင်း
                  </span>
                </div>

                {/* Paragraph 1 */}
                <p className="text-[13px] text-justify leading-relaxed my-3">
                  ၁။ အထက်ပါ ကိစ္စနှင့်ပတ်သက်၍ {activeReportRecord.requestDepartment} မှ မေးမြန်းထားသော {activeReportRecord.reasonDescription && activeReportRecord.reasonDescription.trim() ? `"${activeReportRecord.reasonDescription.trim()}" ဆိုင်ရာ ` : ''}စောင့်ကြည့်စာရင်းပါ အောက်ဖော်ပြပါ ပုဂ္ဂိုလ်များအား လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်) ၏ လူဝင်မှုကြီးကြပ်ရေး မှတ်တမ်းစနစ်အတွင်းရှိ ဝင်ရောက်/ထွက်ခွာမှု၊ တည်းခိုနေထိုင်မှု မှတ်တမ်းများနှင့် တိုက်ဆိုင်စစ်ဆေးခဲ့ပါသည်။
                </p>

                {/* SECTION 1: Subject Person Details Box */}
                <div className="my-4 border border-black p-3 space-y-2 text-[12px] bg-slate-50/60 rounded-xs">
                  <div className="font-bold underline text-[13px] text-black">
                    ၂။ မေးမြန်းထားသည့် ပုဂ္ဂိုလ်များ၏ ကိုယ်ရေးအချက်အလက်များ -
                  </div>
                  {activeReportRecord.persons.map((p, idx) => (
                    <div key={p.id} className="grid grid-cols-2 gap-2 border-b border-black/20 pb-2.5 last:border-0 last:pb-0 pt-1">
                      <div><span className="font-bold">({toBurmeseDigits(idx + 1)}) အမည်:</span> {p.name}</div>
                      <div><span className="font-bold">နိုင်ငံကူးလက်မှတ်အမှတ်:</span> {p.passport || '-'}</div>
                      <div><span className="font-bold">နိုင်ငံသား:</span> {p.nationality || '-'}</div>
                      <div><span className="font-bold">အဖအမည်:</span> {p.fatherName || '-'}</div>
                      <div><span className="font-bold">အမိအမည်:</span> {p.motherName || '-'}</div>
                      <div><span className="font-bold">မွေးသက္ကရာဇ်:</span> {p.dob ? toBurmeseFormattedDate(p.dob) : '-'}</div>
                      <div><span className="font-bold">ဇာတိ:</span> {p.birthPlace || '-'}</div>
                      <div><span className="font-bold">နေရပ်လိပ်စာ:</span> {p.address || '-'}</div>
                      {p.idNumber && <div><span className="font-bold">ID / မှတ်ပုံတင်:</span> {p.idNumber} ({p.idType || 'ID'})</div>}
                      {p.additionalInfo && <div className="col-span-2"><span className="font-bold">အခြားမှတ်ချက်:</span> {p.additionalInfo}</div>}
                    </div>
                  ))}
                </div>

                {/* SECTION 2: DISTINCT MATCHED & CROSS-REFERENCED DATA */}
                <div className="my-5 border-2 border-black p-3.5 space-y-3 text-[12px] bg-white rounded-xs">
                  <div className="font-bold text-[13px] text-black flex items-center justify-between border-b-2 border-black pb-2">
                    <span className="underline">၃။ စနစ်အတွင်း တိုက်ဆိုင်စစ်ဆေး တွေ့ရှိကိုက်ညီသည့် အချက်အလက်များ (Matched Verification Data) -</span>
                  </div>

                  {activeReportRecord.persons.map((p, pIdx) => {
                    const matchInfo = activeReportRecord.matchResults?.[p.id];
                    const recs: (ImmRecord & { matchReasons?: string[]; isDirect?: boolean })[] = matchInfo?.matchedRecords || [];
                    const chks: any[] = matchInfo?.checkingEntries || [];

                    return (
                      <div key={p.id} className="space-y-2 border-b border-black/20 pb-3 last:border-0 last:pb-0">
                        <div className="font-bold text-[13px] flex items-center justify-between flex-wrap gap-1 bg-slate-100 p-1.5 border border-black/30">
                          <span>
                            ({toBurmeseDigits(pIdx + 1)}) {p.name} ({p.passport || 'No Passport'}) - {p.nationality || ''}
                            {p.fatherName && <span className="font-normal text-[11px] ml-2">(အဖ: {p.fatherName})</span>}
                          </span>
                          <span className="text-[11px] px-2 py-0.5 rounded font-bold border border-black">
                            {matchInfo?.summaryStatus || 'စစ်ဆေးပြီး'}
                          </span>
                        </div>

                        {/* Movement logs table */}
                        {recs.length > 0 && (
                          <div className="mt-2 space-y-1">
                            <div className="font-bold text-[11px] text-black underline">
                              • လေဆိပ်/ဂိတ် ဝင်ရောက်/ထွက်ခွာမှု မှတ်တမ်း ({toBurmeseDigits(recs.length)} ကြိမ်):
                            </div>
                            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu">
                              <thead>
                                <tr className="bg-slate-200 font-bold border-b border-black">
                                  <th className="p-1 border border-black w-7">စဉ်</th>
                                  <th className="p-1 border border-black">ရက်စွဲ</th>
                                  <th className="p-1 border border-black">အဝင်/အထွက်</th>
                                  <th className="p-1 border border-black">စနစ်တွင်းခရီးသည်</th>
                                  <th className="p-1 border border-black">တိုက်ဆိုင်ကိုက်ညီမှု အကြောင်းရင်း</th>
                                  <th className="p-1 border border-black">ယာဉ်/လေယာဉ်</th>
                                  <th className="p-1 border border-black">ဗီဇာ</th>
                                  <th className="p-1 border border-black">နေခွင့်ကာလ</th>
                                  <th className="p-1 border border-black">တည်းခိုရာလိပ်စာ</th>
                                  <th className="p-1 border border-black">တာဝန်ခံ</th>
                                </tr>
                              </thead>
                              <tbody>
                                {recs.map((r, rIdx) => (
                                  <tr key={r.id} className="border-b border-black">
                                    <td className="p-1 border border-black font-bold">{toBurmeseDigits(rIdx + 1)}</td>
                                    <td className="p-1 border border-black">{toBurmeseFormattedDate(r.timestamp.split(', ')[0])}</td>
                                    <td className="p-1 border border-black font-bold">{r.mode === 'IN' ? 'အဝင်' : 'အထွက်'}</td>
                                    <td className="p-1 border border-black text-left font-bold">{r.fullname}</td>
                                    <td className="p-1 border border-black text-left">
                                      {r.matchReasons && r.matchReasons.length > 0 ? r.matchReasons.join(', ') : '-'}
                                    </td>
                                    <td className="p-1 border border-black font-bold">{r.vehicleInfo || '-'}</td>
                                    <td className="p-1 border border-black">{r.visaType || '-'}</td>
                                    <td className="p-1 border border-black">{r.stayFrom ? `${toBurmeseFormattedDate(r.stayFrom)} မှ ${toBurmeseFormattedDate(r.stayTo)}` : '-'}</td>
                                    <td className="p-1 border border-black text-left">{r.address || '-'}</td>
                                    <td className="p-1 border border-black">{r.officialName || '-'}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}

                        {/* Stay checking logs */}
                        {chks.length > 0 && (
                          <div className="mt-2 space-y-1">
                            <div className="font-bold text-[11px] text-black underline">
                              • တည်းခိုနေထိုင်မှု မြေပြင်စစ်ဆေးချက် မှတ်တမ်း ({toBurmeseDigits(chks.length)} ကြိမ်):
                            </div>
                            <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu">
                              <thead>
                                <tr className="bg-slate-200 font-bold border-b border-black">
                                  <th className="p-1 border border-black w-7">စဉ်</th>
                                  <th className="p-1 border border-black">စစ်ဆေးသည့်ရက်</th>
                                  <th className="p-1 border border-black">စစ်ဆေးတွေ့ရှိလိပ်စာ</th>
                                  <th className="p-1 border border-black">တိုက်ဆိုင်ကိုက်ညီမှု</th>
                                  <th className="p-1 border border-black">ခွင့်ပြုချက်</th>
                                  <th className="p-1 border border-black">စစ်ဆေးသူအရာရှိ</th>
                                </tr>
                              </thead>
                              <tbody>
                                {chks.map((c: any, cIdx: number) => (
                                  <tr key={c.id || cIdx} className="border-b border-black">
                                    <td className="p-1 border border-black font-bold">{toBurmeseDigits(cIdx + 1)}</td>
                                    <td className="p-1 border border-black">{toBurmeseFormattedDate(c.checkDate)}</td>
                                    <td className="p-1 border border-black text-left">{c.confirmedAddress || c.originalAddress || '-'}</td>
                                    <td className="p-1 border border-black text-left">{c.matchReason || '-'}</td>
                                    <td className="p-1 border border-black font-bold">{c.status || '-'}</td>
                                    <td className="p-1 border border-black">{c.officerName || '-'}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}

                        {/* If no match */}
                        {recs.length === 0 && chks.length === 0 && (
                          <div className="p-2 border border-dashed border-black/50 text-[12px] text-slate-800 bg-slate-50/50">
                            • လူဝင်မှုကြီးကြပ်ရေး မှတ်တမ်းစနစ်အတွင်း အဆိုပါပုဂ္ဂိုလ်၏ ဝင်ရောက်/ထွက်ခွာမှုနှင့် တည်းခိုနေထိုင်မှု မှတ်တမ်းများ (လုံးဝ) တွေ့ရှိခြင်း မရှိပါ (No Movement/Stay Record Found)။
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>

                {/* SECTION 3: OFFICIAL CONFIRMATION & VERIFICATION OPINION */}
                {activeReportRecord.officialConfirmation && (
                  <div className="my-4 border border-black p-3 space-y-2 text-[12px] bg-slate-50 rounded-xs">
                    <div className="font-bold underline text-[13px] text-black">
                      ၄။ တာဝန်ခံအရာရှိ၏ စိစစ်အတည်ပြုချက် (Official Verification Finding & Decision) -
                    </div>
                    <div className="space-y-1.5 pl-2">
                      <div className="flex items-start gap-2">
                        <span className="font-bold min-w-[120px]">• စိစစ်ဆုံးဖြတ်ချက်:</span>
                        <span className="font-black text-[13px]">
                          {activeReportRecord.officialConfirmation.status === 'CONFIRMED_MATCH'
                            ? 'စနစ်တွင်း တိုက်ဆိုင်တွေ့ရှိချက်အရ စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်နှင့် အမှန်တကယ် (ကိုက်ညီမှုရှိပါသည် - CONFIRMED MATCH)'
                            : activeReportRecord.officialConfirmation.status === 'CLEARED_MISMATCH'
                            ? 'စနစ်တွင်း တွေ့ရှိချက်သည် အမည်တူ/လိပ်စာတူ ဆက်စပ်ရုံမျှဖြစ်ပြီး စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်နှင့် (ကိုက်ညီမှုမရှိပါ/လူမှားဖြစ်ပါသည် - CLEARED MISMATCH)'
                            : 'စောင့်ကြည့်စာရင်းပါ ပုဂ္ဂိုလ်နှင့် ပတ်သက်၍ ဆက်လက် စိစစ်ဆဲ ဖြစ်ပါသည် (PENDING REVIEW)'}
                        </span>
                      </div>
                      {activeReportRecord.officialConfirmation.remarks && (
                        <div className="flex items-start gap-2">
                          <span className="font-bold min-w-[120px]">• စိစစ်ချက်မှတ်ချက်:</span>
                          <span>{activeReportRecord.officialConfirmation.remarks}</span>
                        </div>
                      )}
                      <div className="flex items-start gap-2 text-slate-700">
                        <span className="font-bold min-w-[120px]">• စိစစ်သည့်အရာရှိ:</span>
                        <span>
                          {activeReportRecord.officialConfirmation.confirmedBy || reportOfficerName || '-'} ({activeReportRecord.officialConfirmation.confirmedTitle || reportOfficerTitle || 'ဒုတိယလဝကမှူး'}) | စိစစ်သည့်ရက်စွဲ: {toBurmeseFormattedDate(activeReportRecord.officialConfirmation.confirmedAt)}
                        </span>
                      </div>
                    </div>
                  </div>
                )}

                {/* Conclusion Paragraph */}
                <p className="text-[13px] leading-relaxed my-4 text-justify">
                  {activeReportRecord.officialConfirmation ? '၅။' : '၄။'} သို့ဖြစ်ပါ၍ စောင့်ကြည့်စာရင်းပါ အချက်အလက်များနှင့် ပတ်သက်၍ စနစ်တွင်း စစ်ဆေးတွေ့ရှိချက်များအား သိရှိနိုင်ပါရန်နှင့် လိုအပ်သလို ဆက်လက်ဆောင်ရွက်နိုင်ပါရန် လေးစားစွာဖြင့် အစီရင်ခံ တင်ပြအပ်ပါသည်။
                </p>

                {/* Signature block */}
                <div className="flex justify-end mt-12 pr-4" style={{ pageBreakInside: 'avoid', breakInside: 'avoid' }}>
                  <div className="text-center space-y-1 font-pyidaungsu min-w-[200px]">
                    <div className="h-10"></div>
                    <div className="text-[13px] font-bold">({reportOfficerName || 'တာဝန်ခံအရာရှိ'})</div>
                    <div className="text-[13px] font-bold">{reportOfficerTitle || 'ဒုတိယလဝကမှူး'}</div>
                    <div className="text-[12px] font-medium text-black">လူဝင်မှုကြီးကြပ်ရေးနှင့် ပြည်သူ့အင်အားဝန်ကြီးဌာန</div>
                    <div className="text-[12px] font-medium text-black">လူဝင်မှုကြီးကြပ်ရေးဒုတိယဌာနခွဲ(မြိတ်)</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
