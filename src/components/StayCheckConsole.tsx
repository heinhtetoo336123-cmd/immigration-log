import React, { useState, useMemo } from 'react';
import { 
  Building2, 
  MapPin, 
  Search, 
  CheckCircle2, 
  Clock, 
  FileSpreadsheet, 
  Printer, 
  CheckSquare, 
  Square, 
  RefreshCw, 
  ArrowLeft, 
  ShieldCheck, 
  Users, 
  Calendar, 
  FileText, 
  Check, 
  AlertCircle,
  Home,
  Briefcase,
  X,
  UserCheck,
  ShieldAlert,
  ArrowRight,
  Info
} from 'lucide-react';
import { ImmRecord, MasterItem, CheckingHistoryEntry, MovementData, normalizePermitStatus, PermitStatus } from '../types';
import { isCompanyAddress } from '../utils/addressUtils';
import { logActivity } from '../utils/activityLogger';
import { MasterLocationInput } from './MasterLocationInput';
import * as XLSX from 'xlsx';

interface StayCheckConsoleProps {
  records: ImmRecord[];
  setRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  tempRecords?: ImmRecord[];
  setTempRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  checkingHistory: CheckingHistoryEntry[];
  setCheckingHistory?: React.Dispatch<React.SetStateAction<CheckingHistoryEntry[]>>;
  masterData?: MasterItem[];
  movementMap?: Record<string, MovementData>;
  passportToLatestInfo?: Record<string, { fullname: string; nationality: string; [key: string]: any }>;
  currentUser?: { title?: string; name?: string } | null;
  cloudAuthUser?: { username: string; role: string } | null;
  isViewer?: boolean;
  showToast: (msg: string) => void;
  setActivePrintPreview?: (preview: any) => void;
  syncMaster?: (value: string, type: MasterItem['type'], linkedValue?: string) => void;
}

export interface LocationResidentItem {
  passport: string;
  fullname: string;
  gender: 'M' | 'F' | '';
  nationality: string;
  visaType: string;
  visaNumber?: string;
  address: string;
  stayDescription?: string;
  stayFrom?: string;
  stayTo?: string;
  totalDays?: string;
  vehicleInfo?: string;
  broughtBy?: string;
  contactDetails?: string;
  isStillIn: boolean;
  arrivedDate?: string;
  departedDate?: string;
  stillPermittedStatus: PermitStatus | '';
  permittedBy?: string;
  remarks?: string;
  latestCheck?: CheckingHistoryEntry;
  latestTimestamp?: string;
}

export interface LocationSummaryItem {
  name: string;
  category: 'COMPANY' | 'OTHER';
  isCo: boolean;
  totalRecords: number;
  stillInCount: number;
  departedCount: number;
  verifiedCount: number;
  unverifiedCount: number;
  permittedCount: number;
  notPermittedCount: number;
  residents: LocationResidentItem[];
  linkedStayDesc?: string;
}

export const StayCheckConsole: React.FC<StayCheckConsoleProps> = ({
  records = [],
  setRecords,
  tempRecords = [],
  setTempRecords,
  checkingHistory = [],
  setCheckingHistory,
  masterData = [],
  movementMap = {},
  passportToLatestInfo = {},
  currentUser,
  cloudAuthUser,
  isViewer,
  showToast,
  setActivePrintPreview,
  syncMaster
}) => {
  // Navigation & Location Selection Mode
  const [selectedLocationName, setSelectedLocationName] = useState<string | null>(null);
  const [searchLocationQuery, setSearchLocationQuery] = useState('');
  const [locationCategoryFilter, setLocationCategoryFilter] = useState<'ALL' | 'COMPANY' | 'OTHER'>('ALL');

  // In-Location Resident Filters (Strictly Still In Only)
  const [residentSearchQuery, setResidentSearchQuery] = useState('');
  const [residentStatusFilter, setResidentStatusFilter] = useState<'ALL' | 'VERIFIED' | 'PENDING' | 'PERMITTED' | 'NOT_PERMITTED'>('ALL');
  const [selectedPassports, setSelectedPassports] = useState<string[]>([]);

  // Officer Confirmation controls
  const [officerTitle, setOfficerTitle] = useState<string>(() => {
    return localStorage.getItem('staycheck_officer_title') || currentUser?.title || 'လဝကမှူး';
  });
  const [officerName, setOfficerName] = useState<string>(() => {
    return localStorage.getItem('staycheck_officer_name') || currentUser?.name || cloudAuthUser?.username || '';
  });

  const getTodayISO = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  const getCurrentTimeISO = () => {
    const d = new Date();
    return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  };

  const [checkDate, setCheckDate] = useState<string>(getTodayISO);
  const [checkTime, setCheckTime] = useState<string>(getCurrentTimeISO);

  // Per-person inline inspection remarks
  const [inlineInspectionRemarks, setInlineInspectionRemarks] = useState<Record<string, string>>({});

  // Relocation Modal State
  const [isRelocateModalOpen, setIsRelocateModalOpen] = useState(false);
  const [relocateTargetAddress, setRelocateTargetAddress] = useState('');
  const [relocateTargetDescription, setRelocateTargetDescription] = useState('');
  const [isRelocatingBatch, setIsRelocatingBatch] = useState(false);
  const [singleRelocatePassport, setSingleRelocatePassport] = useState<string | null>(null);

  const parseTimestamp = (ts?: string): number => {
    if (!ts) return 0;
    try {
      const parts = ts.trim().split(' ');
      if (parts[0] && parts[0].includes('-')) {
        const dateParts = parts[0].split('-');
        if (dateParts[0].length === 4) {
          return new Date(ts.replace(' ', 'T')).getTime() || 0;
        }
        const [d, m, y] = dateParts.map(Number);
        return new Date(y, m - 1, d).getTime();
      }
      return new Date(ts).getTime() || 0;
    } catch {
      return 0;
    }
  };

  // Group all records and master data into clean location summaries
  const locationSummaries: LocationSummaryItem[] = useMemo(() => {
    const locMap = new Map<string, LocationSummaryItem>();

    // 1. Seed locations from masterData
    masterData.filter(m => m.type === 'Stay' && m.name?.trim()).forEach(m => {
      const cleanName = m.name.trim();
      const isCo = isCompanyAddress(cleanName, masterData);
      const category: 'COMPANY' | 'OTHER' = isCo ? 'COMPANY' : 'OTHER';

      if (!locMap.has(cleanName)) {
        locMap.set(cleanName, {
          name: cleanName,
          category,
          isCo,
          totalRecords: 0,
          stillInCount: 0,
          departedCount: 0,
          verifiedCount: 0,
          unverifiedCount: 0,
          permittedCount: 0,
          notPermittedCount: 0,
          residents: [],
          linkedStayDesc: m.linkedValue || ''
        });
      }
    });

    // 2. Aggregate unique individuals by passport and their latest stay location
    const passportMap = new Map<string, ImmRecord[]>();
    [...records, ...tempRecords].forEach(r => {
      if (!r.passport) return;
      const p = r.passport.trim().toUpperCase();
      if (!passportMap.has(p)) {
        passportMap.set(p, []);
      }
      passportMap.get(p)!.push(r);
    });

    passportMap.forEach((pRecords, p) => {
      // Sort newest first
      pRecords.sort((a, b) => (b.id || 0) - (a.id || 0));
      const latest = pRecords[0];
      const addr = (latest.address || 'Unknown Location').trim();

      // Check movement status - STRICT STILL IN ONLY (exclude left/departed persons completely)
      const mov = movementMap[p];
      let isStillIn = true;
      if (mov) {
        const inTime = Number(mov.inTime || 0);
        const outTime = Number(mov.outTime || 0);
        const hasDeparted = Boolean(mov.out || mov.isStillIn === false || (outTime > 0 && inTime > 0 && outTime >= inTime));
        isStillIn = !hasDeparted && Boolean(mov.in || latest.mode === 'IN');
      } else {
        isStillIn = latest.mode === 'IN';
      }

      // Check physical latest mode
      if (latest.mode === 'OUT') {
        isStillIn = false;
      }

      // STRICT USER REQUIREMENT: Only Still In foreigners are included; left/departed persons are excluded!
      if (!isStillIn) return;

      // Check latest inspection check
      const checksForP = checkingHistory
        .filter(c => c.passport?.trim().toUpperCase() === p)
        .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
      const latestCheck = checksForP[0];

      // Normalized Permit Status
      const rawStatus = latest.stillPermittedStatus || mov?.stillPermittedStatus || '';
      const permStatus = normalizePermitStatus(rawStatus) || 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ';

      let locItem = locMap.get(addr);
      if (!locItem) {
        const isCo = isCompanyAddress(addr, masterData);
        const category: 'COMPANY' | 'OTHER' = isCo ? 'COMPANY' : 'OTHER';

        locItem = {
          name: addr,
          category,
          isCo,
          totalRecords: 0,
          stillInCount: 0,
          departedCount: 0,
          verifiedCount: 0,
          unverifiedCount: 0,
          permittedCount: 0,
          notPermittedCount: 0,
          residents: [],
          linkedStayDesc: latest.stayDescription || ''
        };
        locMap.set(addr, locItem);
      }

      const passInfo = passportToLatestInfo[p] || {};
      const fullname = latest.fullname || mov?.n || passInfo.fullname || 'Unknown';
      const nationality = latest.nationality || mov?.nat || passInfo.nationality || '-';
      const gender = (latest.gender || mov?.gender || passInfo.gender || '') as 'M' | 'F' | '';

      const residentObj: LocationResidentItem = {
        passport: p,
        fullname,
        gender,
        nationality,
        visaType: latest.visaType || mov?.visa || '-',
        visaNumber: latest.visaNumber || mov?.vInfo || '',
        address: addr,
        stayDescription: latest.stayDescription || mov?.stayDescription || locItem.linkedStayDesc,
        stayFrom: latest.stayFrom || '',
        stayTo: latest.stayTo || '',
        totalDays: latest.totalDays || '',
        vehicleInfo: latest.vehicleInfo || '',
        broughtBy: latest.broughtBy || '',
        contactDetails: latest.contactDetails || '',
        isStillIn: true,
        arrivedDate: latest.mode === 'IN' ? latest.timestamp : (mov?.in || '-'),
        departedDate: '-',
        stillPermittedStatus: permStatus,
        permittedBy: latest.permittedBy || mov?.permittedBy || '',
        remarks: latest.remarks || '',
        latestCheck,
        latestTimestamp: latest.timestamp
      };

      locItem.residents.push(residentObj);
      locItem.totalRecords += 1;
      locItem.stillInCount += 1;

      if (latestCheck) {
        locItem.verifiedCount += 1;
      } else {
        locItem.unverifiedCount += 1;
      }

      if (permStatus === 'နေထိုင်ခွင့်ကျထားသောသူ') {
        locItem.permittedCount += 1;
      } else if (permStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') {
        locItem.notPermittedCount += 1;
      }
    });

    return Array.from(locMap.values()).sort((a, b) => {
      // Sort with active still-in residents first, then by name
      if (b.stillInCount !== a.stillInCount) {
        return b.stillInCount - a.stillInCount;
      }
      return a.name.localeCompare(b.name);
    });
  }, [records, tempRecords, checkingHistory, masterData, movementMap, passportToLatestInfo]);

  // Filtered Locations List (only show locations that have still-in residents, or matches search query)
  const filteredLocations = useMemo(() => {
    return locationSummaries.filter(loc => {
      if (locationCategoryFilter === 'COMPANY' && !loc.isCo) return false;
      if (locationCategoryFilter === 'OTHER' && loc.isCo) return false;

      if (searchLocationQuery.trim()) {
        const q = searchLocationQuery.toLowerCase();
        const matchesName = loc.name.toLowerCase().includes(q);
        const matchesDesc = (loc.linkedStayDesc || '').toLowerCase().includes(q);
        const matchesResident = loc.residents.some(r => 
          r.fullname.toLowerCase().includes(q) || 
          r.passport.toLowerCase().includes(q) ||
          r.nationality.toLowerCase().includes(q)
        );
        if (!matchesName && !matchesDesc && !matchesResident) return false;
        return true;
      }

      // Default: show locations that currently host still-in foreigners
      return loc.stillInCount > 0;
    });
  }, [locationSummaries, searchLocationQuery, locationCategoryFilter]);

  // Currently Selected Location Object
  const selectedLocation = useMemo(() => {
    if (!selectedLocationName) return null;
    return locationSummaries.find(l => l.name === selectedLocationName) || null;
  }, [locationSummaries, selectedLocationName]);

  // Filtered Residents inside Selected Location (100% Still-In residents only)
  const filteredResidents = useMemo(() => {
    if (!selectedLocation) return [];
    return selectedLocation.residents.filter(r => {
      if (residentStatusFilter === 'VERIFIED' && !r.latestCheck) return false;
      if (residentStatusFilter === 'PENDING' && r.latestCheck) return false;
      if (residentStatusFilter === 'PERMITTED' && r.stillPermittedStatus !== 'နေထိုင်ခွင့်ကျထားသောသူ') return false;
      if (residentStatusFilter === 'NOT_PERMITTED' && r.stillPermittedStatus !== 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') return false;

      if (residentSearchQuery.trim()) {
        const q = residentSearchQuery.toLowerCase();
        const matchesP = r.passport.toLowerCase().includes(q);
        const matchesN = r.fullname.toLowerCase().includes(q);
        const matchesNat = r.nationality.toLowerCase().includes(q);
        const matchesDesc = (r.stayDescription || '').toLowerCase().includes(q);
        const matchesPermittedBy = (r.permittedBy || '').toLowerCase().includes(q);
        if (!matchesP && !matchesN && !matchesNat && !matchesDesc && !matchesPermittedBy) return false;
      }
      return true;
    });
  }, [selectedLocation, residentSearchQuery, residentStatusFilter]);

  // Quick Timestamp updater
  const handleSetNow = () => {
    const today = getTodayISO();
    const timeNow = getCurrentTimeISO();
    setCheckDate(today);
    setCheckTime(timeNow);
    showToast('🕒 ယနေ့ရက်စွဲနှင့် လက်ရှိအချိန် သတ်မှတ်ပြီးပါပြီ');
  };

  /**
   * Perform physical stay inspection check for one resident at this location.
   * NOTE: This DOES NOT modify or overwrite the foreigner's stillPermittedStatus!
   */
  const handleVerifyResident = (resident: LocationResidentItem) => {
    if (isViewer) {
      showToast('⚠️ View-Only mode: Cannot modify stay check records');
      return;
    }

    const checkTimestamp = `${checkDate} ${checkTime}`;
    const customRemark = inlineInspectionRemarks[resident.passport] || '';
    const defaultInspectionDesc = customRemark 
      ? customRemark 
      : `${selectedLocation?.name} တွင် နေထိုင်လျက်ရှိကြောင်း စစ်ဆေးတွေ့ရှိရပါသည် (${checkDate})`;

    const newCheckEntry: CheckingHistoryEntry = {
      id: `CHK-${Date.now()}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`,
      passport: resident.passport,
      fullname: resident.fullname,
      nationality: resident.nationality,
      type: selectedLocation?.isCo ? 'CO_LTD' : 'OTHERS',
      checkDate: checkDate,
      originalAddress: selectedLocation?.name || resident.address,
      confirmedAddress: selectedLocation?.name || resident.address,
      confirmedStayDescription: resident.stayDescription || selectedLocation?.linkedStayDesc || '',
      status: resident.stillPermittedStatus || 'နေထိုင်ခွင့်ကျထားသောသူ',
      permittedBy: defaultInspectionDesc,
      officerName: officerName.trim() || currentUser?.name || 'Officer',
      officerTitle: officerTitle.trim() || 'လဝကမှူး',
      timestamp: checkTimestamp,
      syncStatus: 'pending_sync',
      updatedAt: new Date().toISOString()
    };

    // Save to checkingHistory
    if (setCheckingHistory) {
      setCheckingHistory(prev => [newCheckEntry, ...prev.filter(c => c.passport.toUpperCase() !== resident.passport.toUpperCase())]);
    }

    logActivity({
      action: 'UPDATE',
      module: 'CHECKING',
      details: `[Stay Check Inspection] Verified physical presence for ${resident.passport} (${resident.fullname}) at "${selectedLocation?.name}"`,
      targetId: resident.passport,
      officerName: `${officerTitle} ${officerName}`
    });

    showToast(`✓ ${resident.passport} (${resident.fullname}) ၏ နေထိုင်မှုစစ်ဆေးချက်အား အောင်မြင်စွာ မှတ်တမ်းတင်ပြီးပါပြီ`);
  };

  /**
   * Perform Batch Physical Verification for selected residents.
   * NOTE: Does NOT overwrite stillPermittedStatus.
   */
  const handleBatchVerify = () => {
    if (isViewer) {
      showToast('⚠️ View-Only mode: Action not permitted');
      return;
    }
    if (selectedPassports.length === 0) {
      showToast('⚠️ ကျေးဇူးပြု၍ စစ်ဆေးအတည်ပြုလိုသူ အနည်းဆုံးတစ်ဦးကို ရွေးချယ်ပါ');
      return;
    }

    const checkTimestamp = `${checkDate} ${checkTime}`;
    const newEntries: CheckingHistoryEntry[] = [];

    const selectedResidents = (selectedLocation?.residents || []).filter(r => 
      selectedPassports.includes(r.passport)
    );

    selectedResidents.forEach(res => {
      const customRemark = inlineInspectionRemarks[res.passport] || '';
      const desc = customRemark 
        ? customRemark 
        : `${selectedLocation?.name} တွင် နေထိုင်လျက်ရှိကြောင်း စစ်ဆေးတွေ့ရှိရပါသည် (${checkDate})`;

      newEntries.push({
        id: `CHK-${Date.now()}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`,
        passport: res.passport,
        fullname: res.fullname,
        nationality: res.nationality,
        type: selectedLocation?.isCo ? 'CO_LTD' : 'OTHERS',
        checkDate: checkDate,
        originalAddress: selectedLocation?.name || res.address,
        confirmedAddress: selectedLocation?.name || res.address,
        confirmedStayDescription: res.stayDescription || selectedLocation?.linkedStayDesc || '',
        status: res.stillPermittedStatus || 'နေထိုင်ခွင့်ကျထားသောသူ',
        permittedBy: desc,
        officerName: officerName.trim() || currentUser?.name || 'Officer',
        officerTitle: officerTitle.trim() || 'လဝကမှူး',
        timestamp: checkTimestamp,
        syncStatus: 'pending_sync',
        updatedAt: new Date().toISOString()
      });
    });

    if (setCheckingHistory) {
      const updatedPassportSet = new Set(selectedPassports.map(p => p.toUpperCase()));
      setCheckingHistory(prev => [...newEntries, ...prev.filter(c => !updatedPassportSet.has(c.passport.toUpperCase()))]);
    }

    logActivity({
      action: 'UPDATE',
      module: 'CHECKING',
      details: `[Stay Check Batch Inspection] Verified physical presence for ${selectedPassports.length} individuals at "${selectedLocation?.name}"`,
      officerName: `${officerTitle} ${officerName}`
    });

    showToast(`✓ ရွေးချယ်ထားသော ${selectedPassports.length} ဦးအား "${selectedLocation?.name}" တွင် နေထိုင်ကြောင်း စစ်ဆေးအတည်ပြုပြီးပါပြီ`);
    setSelectedPassports([]);
  };

  // Relocate / Replace Stay Address for selected residents or whole location
  const handleExecuteRelocation = () => {
    if (isViewer) {
      showToast('⚠️ View-Only mode: Cannot modify locations');
      return;
    }
    if (!relocateTargetAddress.trim()) {
      showToast('⚠️ ကျေးဇူးပြု၍ ပြောင်းရွှေ့မည့် တည်းခိုနေရာအသစ်အား ဖြည့်သွင်းပါ');
      return;
    }

    const newLocName = relocateTargetAddress.trim();
    const newLocDesc = relocateTargetDescription.trim();

    if (singleRelocatePassport) {
      // Relocate single individual
      const pUpper = singleRelocatePassport.trim().toUpperCase();
      if (setRecords) {
        setRecords(prev => prev.map(r => {
          if (r.passport.trim().toUpperCase() === pUpper) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (setTempRecords) {
        setTempRecords(prev => prev.map(r => {
          if (r.passport.trim().toUpperCase() === pUpper) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (syncMaster) {
        syncMaster(newLocName, 'Stay', newLocDesc || undefined);
      }

      logActivity({
        action: 'UPDATE',
        module: 'CHECKING',
        details: `[Stay Check] Relocated passport ${pUpper} to "${newLocName}"`,
        officerName: `${officerTitle} ${officerName}`
      });

      showToast(`✓ ${pUpper} အား "${newLocName}" သို့ အောင်မြင်စွာ ပြောင်းရွှေ့ပေးလိုက်ပါပြီ`);
      setSingleRelocatePassport(null);
    } else if (isRelocatingBatch && selectedPassports.length > 0) {
      // Relocate only selected passports
      const targetPassportSet = new Set(selectedPassports.map(p => p.toUpperCase()));
      if (setRecords) {
        setRecords(prev => prev.map(r => {
          if (targetPassportSet.has(r.passport.trim().toUpperCase())) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (setTempRecords) {
        setTempRecords(prev => prev.map(r => {
          if (targetPassportSet.has(r.passport.trim().toUpperCase())) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (syncMaster) {
        syncMaster(newLocName, 'Stay', newLocDesc || undefined);
      }

      logActivity({
        action: 'UPDATE',
        module: 'CHECKING',
        details: `[Stay Check] Relocated ${selectedPassports.length} individuals from "${selectedLocation?.name}" to "${newLocName}"`,
        officerName: `${officerTitle} ${officerName}`
      });

      showToast(`✓ ရွေးချယ်ထားသော ${selectedPassports.length} ဦးအား "${newLocName}" သို့ အောင်မြင်စွာ ပြောင်းရွှေ့ပေးလိုက်ပါပြီ`);
      setSelectedPassports([]);
    } else if (selectedLocation) {
      // Relocate everyone currently at this location
      const locPassports = new Set(selectedLocation.residents.map(r => r.passport.toUpperCase()));
      if (setRecords) {
        setRecords(prev => prev.map(r => {
          if (locPassports.has(r.passport.trim().toUpperCase())) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (setTempRecords) {
        setTempRecords(prev => prev.map(r => {
          if (locPassports.has(r.passport.trim().toUpperCase())) {
            return {
              ...r,
              address: newLocName,
              stayDescription: newLocDesc || r.stayDescription,
              syncStatus: 'pending_sync',
              updatedAt: new Date().toISOString()
            };
          }
          return r;
        }));
      }

      if (syncMaster) {
        syncMaster(newLocName, 'Stay', newLocDesc || undefined);
      }

      logActivity({
        action: 'UPDATE',
        module: 'CHECKING',
        details: `[Stay Check] Relocated entire location "${selectedLocation.name}" (${locPassports.size} residents) to "${newLocName}"`,
        officerName: `${officerTitle} ${officerName}`
      });

      showToast(`✓ "${selectedLocation.name}" မှ လူဦးရေ ${locPassports.size} ဦးလုံးအား "${newLocName}" သို့ ပြောင်းရွှေ့ပြီးပါပြီ`);
      setSelectedLocationName(newLocName);
    }

    setIsRelocateModalOpen(false);
    setRelocateTargetAddress('');
    setRelocateTargetDescription('');
    setSingleRelocatePassport(null);
    setIsRelocatingBatch(false);
  };

  // Export current list to Excel
  const handleExportExcel = () => {
    if (!selectedLocation) {
      // Export all locations summary
      const excelData = locationSummaries.map(l => ({
        'တည်းခို/နေထိုင်သည့်နေရာ (Location)': l.name,
        'အမျိုးအစား (Type)': l.isCo ? 'ကုမ္ပဏီ (Co., Ltd)' : 'အခြား / ဟိုတယ် (Other/Hotel)',
        'လိပ်စာအသေးစိတ် (Stay Desc)': l.linkedStayDesc || '-',
        'လက်ရှိနေထိုင်သူ (Still In)': l.stillInCount,
        'စစ်ဆေးပြီးသူ (Verified)': l.verifiedCount,
        'စစ်ဆေးရန်ကျန် (Pending)': l.unverifiedCount,
        'နေထိုင်ခွင့်ကျထားသူ (Permit Granted)': l.permittedCount,
        'မလျှောက်ထားသေးသူ (Not Permitted)': l.notPermittedCount,
        'စုစုပေါင်း မှတ်တမ်း (Total)': l.totalRecords
      }));

      const ws = XLSX.utils.json_to_sheet(excelData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Locations_Summary');
      XLSX.writeFile(wb, `Stay_Check_Locations_Summary_${getTodayISO()}.xlsx`);
      showToast('📊 တည်းခိုနေရာများ စာရင်း အကျဉ်းချုပ်အား Excel ဒေါင်းလုဒ်လုပ်ပြီးပါပြီ');
    } else {
      // Export residents of selected location
      const excelData = filteredResidents.map((r, idx) => ({
        'စဉ် (Sr)': idx + 1,
        'နိုင်ငံကူးလက်မှတ် (Passport)': r.passport,
        'အမည် (Full Name)': r.fullname,
        'ကျား/မ (Gender)': r.gender || '-',
        'နိုင်ငံသား (Nationality)': r.nationality,
        'ဗီဇာ (Visa)': r.visaType,
        'ဆိုက်ရောက်ရက် (Arrival)': r.arrivedDate || '-',
        'နေထိုင်သည့်နေရာ (Location)': r.address,
        'အသေးစိတ် (Detail)': r.stayDescription || '-',
        'နေထိုင်ခွင့် အခြေအနေ (Permit Status)': r.stillPermittedStatus || 'မလျှောက်ထားသေးသူ',
        'နေထိုင်ခွင့် အမိန့်စာ (Permit Order)': r.permittedBy || '-',
        'နေရာ စစ်ဆေးမှု (Physical Check)': r.latestCheck ? `စစ်ဆေးပြီး (${r.latestCheck.timestamp})` : 'မစစ်ဆေးရသေး',
        'စစ်ဆေးသူ အရာရှိ (Officer)': r.latestCheck?.officerName || '-'
      }));

      const ws = XLSX.utils.json_to_sheet(excelData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Residents_List');
      XLSX.writeFile(wb, `Stay_Check_${selectedLocation.name}_${getTodayISO()}.xlsx`);
      showToast(`📊 "${selectedLocation.name}" ၏ နေထိုင်သူများ စာရင်းအား Excel ဒေါင်းလုဒ်လုပ်ပြီးပါပြီ`);
    }
  };

  // Print current view
  const handlePrint = () => {
    if (!setActivePrintPreview) {
      window.print();
      return;
    }

    if (selectedLocation) {
      setActivePrintPreview({
        title: `တည်းခိုနေထိုင်မှု စစ်ဆေးချက် မှတ်တမ်း - ${selectedLocation.name}`,
        type: 'CUSTOM_REPORT',
        columns: [
          { id: 'sr', label: 'စဉ်', width: '40px' },
          { id: 'passport', label: 'နိုင်ငံကူးလက်မှတ်', width: '90px' },
          { id: 'fullname', label: 'အမည်', width: '130px' },
          { id: 'nationality', label: 'နိုင်ငံသား', width: '80px' },
          { id: 'permitStatus', label: 'နေထိုင်ခွင့် အခြေအနေ', width: '120px' },
          { id: 'physicalCheck', label: 'နေရာ စစ်ဆေးချက်', width: '140px' },
          { id: 'officer', label: 'စစ်ဆေးသူ', width: '90px' }
        ],
        rows: filteredResidents.map((r, idx) => ({
          sr: idx + 1,
          passport: r.passport,
          fullname: r.fullname,
          nationality: r.nationality,
          permitStatus: r.stillPermittedStatus || 'မလျှောက်ထားသေးသူ',
          physicalCheck: r.latestCheck ? `စစ်ဆေးပြီး (${r.latestCheck.timestamp})` : 'မစစ်ဆေးရသေး',
          officer: r.latestCheck?.officerName || '-'
        })),
        date: new Date().toLocaleDateString('my-MM'),
        officer: currentUser ? `${currentUser.title} ${currentUser.name}` : ''
      });
    } else {
      setActivePrintPreview({
        title: 'မြိတ်ခရိုင်အတွင်း တည်းခိုနေထိုင်သည့် နေရာများ အကျဉ်းချုပ် စာရင်း',
        type: 'CUSTOM_REPORT',
        columns: [
          { id: 'sr', label: 'စဉ်', width: '40px' },
          { id: 'location', label: 'တည်းခိုနေထိုင်ရာ နေရာ', width: '180px' },
          { id: 'type', label: 'အမျိုးအစား', width: '90px' },
          { id: 'stillIn', label: 'လက်ရှိနေထိုင်သူ', width: '90px' },
          { id: 'verified', label: 'စစ်ဆေးပြီးသူ', width: '80px' },
          { id: 'pending', label: 'စစ်ဆေးရန်ကျန်', width: '80px' }
        ],
        rows: filteredLocations.map((l, idx) => ({
          sr: idx + 1,
          location: l.name,
          type: l.isCo ? 'ကုမ္ပဏီ' : 'အခြား/ဟိုတယ်',
          stillIn: l.stillInCount,
          verified: l.verifiedCount,
          pending: l.unverifiedCount
        })),
        date: new Date().toLocaleDateString('my-MM'),
        officer: currentUser ? `${currentUser.title} ${currentUser.name}` : ''
      });
    }
  };

  // Toggle selection for all filtered residents
  const handleToggleSelectAll = () => {
    if (selectedPassports.length === filteredResidents.length) {
      setSelectedPassports([]);
    } else {
      setSelectedPassports(filteredResidents.map(r => r.passport));
    }
  };

  const handleToggleSelectOne = (passport: string) => {
    setSelectedPassports(prev => 
      prev.includes(passport) ? prev.filter(p => p !== passport) : [...prev, passport]
    );
  };

  // -------------------------------------------------------------
  // RENDER STEP 1: LOCATIONS DIRECTORY VIEW
  // -------------------------------------------------------------
  if (!selectedLocation) {
    const totalLocations = locationSummaries.length;
    const totalStillIn = locationSummaries.reduce((sum, l) => sum + l.stillInCount, 0);
    const totalVerified = locationSummaries.reduce((sum, l) => sum + l.verifiedCount, 0);
    const totalUnverified = locationSummaries.reduce((sum, l) => sum + l.unverifiedCount, 0);
    const totalCompanies = locationSummaries.filter(l => l.isCo).length;
    const totalOthers = locationSummaries.filter(l => !l.isCo).length;

    return (
      <div className="space-y-6">
        {/* Banner Header */}
        <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white p-6 sm:p-8 rounded-3xl shadow-xl border border-slate-800 flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
          <div className="flex items-center gap-4">
            <div className="p-3.5 bg-indigo-600/30 text-indigo-400 rounded-2xl border border-indigo-500/30 shrink-0">
              <MapPin size={32} />
            </div>
            <div>
              <div className="flex items-center gap-2.5 flex-wrap">
                <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight">
                  နေရာအလိုက် နေထိုင်မှုစစ်ဆေးခြင်း (Stay Check Console)
                </h2>
                <span className="text-[10px] bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider">
                  FIELD INSPECTION
                </span>
              </div>
              <p className="text-xs text-slate-300 font-medium mt-1">
                သက်ဆိုင်ရာ လိပ်စာ / ကုမ္ပဏီ / ဟိုတယ်များတွင် နိုင်ငံခြားသားများ အမှန်တကယ် နေထိုင်ခြင်း ရှိ/မရှိ ကွင်းဆင်းစစ်ဆေး မှတ်တမ်းတင်ခြင်း
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5 w-full md:w-auto">
            <button
              onClick={handleExportExcel}
              className="bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer border border-emerald-500/40"
            >
              <FileSpreadsheet size={15} />
              <span>Excel ထုတ်မည်</span>
            </button>
            <button
              onClick={handlePrint}
              className="bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer border border-slate-700"
            >
              <Printer size={15} />
              <span>Print</span>
            </button>
          </div>
        </div>

        {/* Global Stats Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3.5">
          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <div className="text-[10px] font-bold uppercase text-slate-400">တည်းခိုနေရာ စုစုပေါင်း</div>
              <div className="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">{totalLocations}</div>
            </div>
            <div className="w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center text-slate-700">
              <MapPin size={20} />
            </div>
          </div>

          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <div className="text-[10px] font-bold uppercase text-indigo-600">လက်ရှိ နေထိုင်သူဦးရေ</div>
              <div className="text-xl sm:text-2xl font-black text-indigo-950 mt-0.5">{totalStillIn}</div>
            </div>
            <div className="w-10 h-10 rounded-xl bg-indigo-50 text-indigo-700 flex items-center justify-center">
              <Users size={20} />
            </div>
          </div>

          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <div className="text-[10px] font-bold uppercase text-emerald-600">စစ်ဆေးပြီးသူ (Verified)</div>
              <div className="text-xl sm:text-2xl font-black text-emerald-700 mt-0.5">{totalVerified}</div>
            </div>
            <div className="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-700 flex items-center justify-center">
              <CheckCircle2 size={20} />
            </div>
          </div>

          <div 
            onClick={() => setLocationCategoryFilter(locationCategoryFilter === 'COMPANY' ? 'ALL' : 'COMPANY')}
            className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between ${
              locationCategoryFilter === 'COMPANY' 
                ? 'bg-purple-50 border-purple-400 ring-2 ring-purple-500/20' 
                : 'bg-white border-slate-200 hover:border-purple-300'
            }`}
          >
            <div>
              <div className="text-[10px] font-bold uppercase text-purple-700">ကုမ္ပဏီများ (Co., Ltd)</div>
              <div className="text-xl sm:text-2xl font-black text-purple-700 mt-0.5">{totalCompanies}</div>
            </div>
            <div className="w-10 h-10 rounded-xl bg-purple-100 text-purple-700 flex items-center justify-center">
              <Building2 size={20} />
            </div>
          </div>

          <div 
            onClick={() => setLocationCategoryFilter(locationCategoryFilter === 'OTHER' ? 'ALL' : 'OTHER')}
            className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between col-span-2 sm:col-span-1 ${
              locationCategoryFilter === 'OTHER' 
                ? 'bg-blue-50 border-blue-400 ring-2 ring-blue-500/20' 
                : 'bg-white border-slate-200 hover:border-blue-300'
            }`}
          >
            <div>
              <div className="text-[10px] font-bold uppercase text-blue-700">အခြား / ဟိုတယ်များ</div>
              <div className="text-xl sm:text-2xl font-black text-blue-700 mt-0.5">{totalOthers}</div>
            </div>
            <div className="w-10 h-10 rounded-xl bg-blue-100 text-blue-700 flex items-center justify-center">
              <Home size={20} />
            </div>
          </div>
        </div>

        {/* Search & Location Filter Bar */}
        <div className="bg-white p-4 sm:p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col md:flex-row gap-3 items-stretch md:items-center justify-between">
          <div className="relative flex-1">
            <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={17} />
            <input
              type="text"
              value={searchLocationQuery}
              onChange={(e) => setSearchLocationQuery(e.target.value)}
              placeholder="တည်းခိုနေရာ၊ ကုမ္ပဏီအမည်၊ လိပ်စာ သို့မဟုတ် နေထိုင်သူအမည်ဖြင့် ရှာပါ..."
              className="w-full pl-10 pr-10 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all"
            />
            {searchLocationQuery && (
              <button 
                onClick={() => setSearchLocationQuery('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
              >
                <X size={14} />
              </button>
            )}
          </div>

          <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl shrink-0">
            <button
              onClick={() => setLocationCategoryFilter('ALL')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer ${
                locationCategoryFilter === 'ALL'
                  ? 'bg-white text-indigo-950 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              အားလုံး ({locationSummaries.length})
            </button>
            <button
              onClick={() => setLocationCategoryFilter('COMPANY')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                locationCategoryFilter === 'COMPANY'
                  ? 'bg-purple-700 text-white shadow-xs'
                  : 'text-purple-700 hover:bg-purple-50'
              }`}
            >
              <Building2 size={13} />
              <span>🏢 ကုမ္ပဏီများ ({totalCompanies})</span>
            </button>
            <button
              onClick={() => setLocationCategoryFilter('OTHER')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                locationCategoryFilter === 'OTHER'
                  ? 'bg-blue-700 text-white shadow-xs'
                  : 'text-blue-700 hover:bg-blue-50'
              }`}
            >
              <Home size={13} />
              <span>🏨 အခြား/ဟိုတယ်များ ({totalOthers})</span>
            </button>
          </div>
        </div>

        {/* Location Cards Grid (Step 1) */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredLocations.length === 0 ? (
            <div className="col-span-full p-16 text-center text-slate-400 font-bold uppercase tracking-wider text-xs bg-slate-50 rounded-2xl border border-dashed border-slate-200">
              ⚠️ ရှာဖွေမှုနှင့် ကိုက်ညီသော တည်းခိုနေရာ မရှိပါ
            </div>
          ) : (
            filteredLocations.map(loc => {
              const isCo = loc.isCo;
              const hasUnverified = loc.unverifiedCount > 0;

              return (
                <div
                  key={'loc_' + loc.name}
                  onClick={() => {
                    setSelectedLocationName(loc.name);
                    setSelectedPassports([]);
                    setResidentSearchQuery('');
                    setResidentStatusFilter('ALL');
                  }}
                  className="bg-white hover:bg-slate-50/80 rounded-2xl p-5 border border-slate-200 hover:border-indigo-400 shadow-sm hover:shadow-md transition-all cursor-pointer flex flex-col justify-between group"
                >
                  <div className="space-y-3">
                    <div className="flex items-start justify-between gap-2">
                      <div className="flex items-center gap-2">
                        <div className={`p-2 rounded-xl shrink-0 ${
                          isCo ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'
                        }`}>
                          {isCo ? <Building2 size={18} /> : <Home size={18} />}
                        </div>
                        <div>
                          <span className={`text-[9px] font-black px-1.5 py-0.2 rounded uppercase ${
                            isCo ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                          }`}>
                            {isCo ? '🏢 CO., LTD' : '🏨 HOTEL / OTHER'}
                          </span>
                        </div>
                      </div>

                      {loc.stillInCount > 0 && (
                        <span className="text-[10px] font-black px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-900 border border-indigo-200">
                          {loc.stillInCount} ဦး နေထိုင်ဆဲ
                        </span>
                      )}
                    </div>

                    <div>
                      <h3 className="text-sm font-black text-slate-900 group-hover:text-indigo-600 transition-colors line-clamp-1">
                        {loc.name}
                      </h3>
                      {loc.linkedStayDesc && (
                        <p className="text-[11px] text-slate-400 font-medium line-clamp-1 mt-0.5">
                          {loc.linkedStayDesc}
                        </p>
                      )}
                    </div>
                  </div>

                  <div className="pt-4 mt-3 border-t border-slate-100 space-y-2">
                    <div className="grid grid-cols-2 gap-2 text-[11px] font-bold">
                      <div className="bg-slate-50 p-2 rounded-xl flex items-center justify-between">
                        <span className="text-slate-400">စစ်ဆေးပြီး:</span>
                        <span className="text-emerald-700 font-black">{loc.verifiedCount}</span>
                      </div>
                      <div className={`p-2 rounded-xl flex items-center justify-between ${
                        hasUnverified ? 'bg-amber-50 text-amber-800' : 'bg-slate-50 text-slate-400'
                      }`}>
                        <span>စစ်ဆေးရန်:</span>
                        <span className="font-black">{loc.unverifiedCount}</span>
                      </div>
                    </div>

                    <div className="flex items-center justify-between text-xs font-black text-indigo-600 group-hover:translate-x-1 transition-transform pt-1">
                      <span>စစ်ဆေးရန် ဝင်ရောက်မည်</span>
                      <ArrowRight size={14} />
                    </div>
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  }

  // -------------------------------------------------------------
  // RENDER STEP 2: RESIDENTS AT CHOSEN LOCATION
  // -------------------------------------------------------------
  return (
    <div className="space-y-6">
      {/* Location Header Breadcrumb & Controls */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white p-6 rounded-3xl shadow-xl border border-slate-800 space-y-4">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div className="flex items-center gap-3">
            <button
              onClick={() => {
                setSelectedLocationName(null);
                setSelectedPassports([]);
              }}
              className="p-2.5 rounded-xl bg-white/10 hover:bg-white/20 text-white border border-white/10 transition-all cursor-pointer flex items-center gap-1.5 text-xs font-black"
            >
              <ArrowLeft size={16} />
              <span>နောက်သို့ (နေရာများစာရင်း)</span>
            </button>

            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h2 className="text-lg sm:text-xl font-black uppercase text-white">
                  {selectedLocation.name}
                </h2>
                <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${
                  selectedLocation.isCo ? 'bg-purple-500/30 text-purple-200 border border-purple-400/40' : 'bg-blue-500/30 text-blue-200 border border-blue-400/40'
                }`}>
                  {selectedLocation.isCo ? '🏢 COMPANY (CO., LTD)' : '🏨 OTHER / HOTEL / RESIDENCE'}
                </span>
              </div>
              {selectedLocation.linkedStayDesc && (
                <p className="text-xs text-indigo-200 font-medium mt-0.5">
                  {selectedLocation.linkedStayDesc}
                </p>
              )}
            </div>
          </div>

          {/* Header Action Buttons */}
          <div className="flex items-center gap-2 flex-wrap w-full md:w-auto">
            <button
              onClick={() => {
                setIsRelocatingBatch(false);
                setRelocateTargetAddress('');
                setRelocateTargetDescription('');
                setIsRelocateModalOpen(true);
              }}
              className="bg-amber-600 hover:bg-amber-500 text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer"
              title="ဤနေရာရှိ နေထိုင်သူများအား နေရာအသစ်သို့ ပြောင်းရွှေ့မည်"
            >
              <RefreshCw size={14} />
              <span>နေရာအားလုံး ပြောင်းမည်</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer"
            >
              <FileSpreadsheet size={15} />
              <span className="hidden sm:inline">Excel</span>
            </button>

            <button
              onClick={handlePrint}
              className="bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer"
            >
              <Printer size={15} />
              <span className="hidden sm:inline">Print</span>
            </button>
          </div>
        </div>

        {/* Location Sub-stats */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-2 border-t border-slate-800/80 text-xs font-bold">
          <div className="bg-slate-800/60 p-2.5 rounded-xl">
            <span className="text-slate-400 block text-[10px] uppercase">လက်ရှိနေထိုင်သူ</span>
            <span className="text-indigo-300 font-black text-base">{selectedLocation.stillInCount} ဦး</span>
          </div>
          <div className="bg-slate-800/60 p-2.5 rounded-xl">
            <span className="text-slate-400 block text-[10px] uppercase">စစ်ဆေးပြီး</span>
            <span className="text-emerald-400 font-black text-base">{selectedLocation.verifiedCount} ဦး</span>
          </div>
          <div className="bg-slate-800/60 p-2.5 rounded-xl">
            <span className="text-slate-400 block text-[10px] uppercase">စစ်ဆေးရန်ကျန်</span>
            <span className="text-amber-400 font-black text-base">{selectedLocation.unverifiedCount} ဦး</span>
          </div>
          <div className="bg-slate-800/60 p-2.5 rounded-xl">
            <span className="text-slate-400 block text-[10px] uppercase">နေထိုင်ခွင့်ကျထားသူ</span>
            <span className="text-cyan-400 font-black text-base">{selectedLocation.permittedCount} ဦး</span>
          </div>
        </div>
      </div>

      {/* Officer Duty Stamp Control Bar */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col lg:flex-row gap-4 items-stretch lg:items-center justify-between">
        <div className="flex items-center gap-3 flex-wrap">
          <div className="flex items-center gap-1.5">
            <Calendar size={15} className="text-slate-400" />
            <span className="text-xs font-black uppercase text-slate-500">စစ်ဆေးသည့်ရက်:</span>
            <input
              type="date"
              value={checkDate}
              onChange={(e) => setCheckDate(e.target.value)}
              className="p-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800"
            />
            <input
              type="time"
              value={checkTime}
              onChange={(e) => setCheckTime(e.target.value)}
              className="p-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800"
            />
            <button
              onClick={handleSetNow}
              className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[11px] font-black cursor-pointer"
            >
              လက်ရှိအချိန်
            </button>
          </div>

          <div className="flex items-center gap-1.5">
            <UserCheck size={15} className="text-slate-400" />
            <span className="text-xs font-black uppercase text-slate-500">စစ်ဆေးသူ:</span>
            <input
              type="text"
              value={officerName}
              onChange={(e) => setOfficerName(e.target.value)}
              placeholder="Officer Name"
              className="p-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 w-32"
            />
            <input
              type="text"
              value={officerTitle}
              onChange={(e) => setOfficerTitle(e.target.value)}
              placeholder="Title"
              className="p-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 w-24"
            />
          </div>
        </div>

        {/* Batch Actions */}
        {selectedPassports.length > 0 && (
          <div className="flex items-center gap-2 bg-indigo-50 border border-indigo-200 p-2 rounded-xl">
            <span className="text-xs font-black text-indigo-950">
              ရွေးချယ်ထားသူ: <span className="text-indigo-600">{selectedPassports.length}</span> ဦး
            </span>
            <button
              onClick={handleBatchVerify}
              className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black px-3 py-1.5 rounded-lg flex items-center gap-1 shadow-xs cursor-pointer"
            >
              <CheckCircle2 size={13} />
              <span>စစ်ဆေးပြီး အတည်ပြုမည်</span>
            </button>
            <button
              onClick={() => {
                setIsRelocatingBatch(true);
                setRelocateTargetAddress('');
                setRelocateTargetDescription('');
                setIsRelocateModalOpen(true);
              }}
              className="bg-amber-600 hover:bg-amber-500 text-white text-xs font-black px-3 py-1.5 rounded-lg flex items-center gap-1 shadow-xs cursor-pointer"
            >
              <RefreshCw size={13} />
              <span>နေရာပြောင်းမည်</span>
            </button>
          </div>
        )}
      </div>

      {/* Resident Filters & Search */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col md:flex-row gap-3 items-stretch md:items-center justify-between">
        <div className="relative flex-1">
          <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
          <input
            type="text"
            value={residentSearchQuery}
            onChange={(e) => setResidentSearchQuery(e.target.value)}
            placeholder="Passport၊ အမည် သို့မဟုတ် နိုင်ငံသားဖြင့် ရှာပါ..."
            className="w-full pl-9 pr-8 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500"
          />
          {residentSearchQuery && (
            <button 
              onClick={() => setResidentSearchQuery('')}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 p-1"
            >
              <X size={13} />
            </button>
          )}
        </div>

        <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl shrink-0 overflow-x-auto text-xs font-black">
          <button
            onClick={() => setResidentStatusFilter('ALL')}
            className={`px-2.5 py-1 rounded-lg transition-all ${residentStatusFilter === 'ALL' ? 'bg-white text-indigo-950 shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
          >
            အားလုံး (နေထိုင်ဆဲ) ({selectedLocation.residents.length})
          </button>
          <button
            onClick={() => setResidentStatusFilter('VERIFIED')}
            className={`px-2.5 py-1 rounded-lg transition-all ${residentStatusFilter === 'VERIFIED' ? 'bg-emerald-600 text-white shadow-xs' : 'text-emerald-700 hover:text-emerald-900'}`}
          >
            စစ်ဆေးပြီး ({selectedLocation.verifiedCount})
          </button>
          <button
            onClick={() => setResidentStatusFilter('PENDING')}
            className={`px-2.5 py-1 rounded-lg transition-all ${residentStatusFilter === 'PENDING' ? 'bg-amber-600 text-white shadow-xs' : 'text-amber-700 hover:text-amber-900'}`}
          >
            မစစ်ဆေးရသေး ({selectedLocation.unverifiedCount})
          </button>
          <button
            onClick={() => setResidentStatusFilter('PERMITTED')}
            className={`px-2.5 py-1 rounded-lg transition-all ${residentStatusFilter === 'PERMITTED' ? 'bg-indigo-600 text-white shadow-xs' : 'text-indigo-700 hover:text-indigo-900'}`}
          >
            နေထိုင်ခွင့်ကျထားသူ ({selectedLocation.permittedCount})
          </button>
          <button
            onClick={() => setResidentStatusFilter('NOT_PERMITTED')}
            className={`px-2.5 py-1 rounded-lg transition-all ${residentStatusFilter === 'NOT_PERMITTED' ? 'bg-slate-700 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'}`}
          >
            မလျှောက်ထားသေး ({selectedLocation.notPermittedCount})
          </button>
        </div>
      </div>

      {/* Residents Table */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="p-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <button
              onClick={handleToggleSelectAll}
              className="p-1 rounded text-slate-600 hover:text-indigo-600 cursor-pointer flex items-center gap-1.5 text-xs font-black"
            >
              {selectedPassports.length === filteredResidents.length && filteredResidents.length > 0 ? (
                <CheckSquare size={16} className="text-indigo-600" />
              ) : (
                <Square size={16} />
              )}
              <span>အားလုံးရွေးမည် ({filteredResidents.length})</span>
            </button>
          </div>
        </div>

        <div className="overflow-x-auto max-h-[650px] overflow-y-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-slate-100 sticky top-0 border-b border-slate-200 z-10 font-bold">
              <tr className="uppercase text-[10px] text-slate-600 font-black">
                <th className="p-3 w-10 text-center"></th>
                <th className="p-3 w-12 text-center">စဉ်</th>
                <th className="p-3">နိုင်ငံကူးလက်မှတ်</th>
                <th className="p-3">အမည်နှင့် နိုင်ငံသား</th>
                <th className="p-3">တရားဝင် နေထိုင်ခွင့် အခြေအနေ</th>
                <th className="p-3">နေရာ စစ်ဆေးအတည်ပြုချက်</th>
                <th className="p-3 text-right w-44">လုပ်ဆောင်ချက်</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 bg-white">
              {filteredResidents.length === 0 ? (
                <tr>
                  <td colSpan={7} className="p-16 text-center text-slate-400 font-bold uppercase tracking-wider text-xs">
                    ⚠️ ရှာဖွေမှုနှင့် ကိုက်ညီသော နေထိုင်သူ မရှိပါ
                  </td>
                </tr>
              ) : (
                filteredResidents.map((r, idx) => {
                  const isChecked = selectedPassports.includes(r.passport);
                  const isVerified = Boolean(r.latestCheck);
                  const isPermitted = r.stillPermittedStatus === 'နေထိုင်ခွင့်ကျထားသောသူ';

                  return (
                    <tr 
                      key={'res_' + r.passport + '_' + idx}
                      className={`hover:bg-slate-50 transition-colors ${
                        isVerified ? 'bg-emerald-50/20' : ''
                      }`}
                    >
                      <td className="p-3 text-center">
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => handleToggleSelectOne(r.passport)}
                          className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                        />
                      </td>

                      <td className="p-3 text-center text-slate-400 font-bold">
                        {idx + 1}
                      </td>

                      <td className="p-3">
                        <div className="flex items-center gap-1.5">
                          <span className="font-mono font-black text-indigo-950 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">
                            {r.passport}
                          </span>
                          {r.gender && (
                            <span className={`text-[9px] font-black px-1.5 py-0.2 rounded ${
                              r.gender === 'M' ? 'bg-blue-100 text-blue-800' : 'bg-pink-100 text-pink-800'
                            }`}>
                              {r.gender}
                            </span>
                          )}
                        </div>
                        <div className="text-[10px] text-slate-400 mt-1">
                          ဆိုက်ရောက်: {r.arrivedDate}
                        </div>
                      </td>

                      <td className="p-3">
                        <div className="font-black text-slate-900 text-xs">
                          {r.fullname}
                        </div>
                        <div className="text-[11px] text-slate-500 font-semibold mt-0.5">
                          {r.nationality} • <span className="text-indigo-600">{r.visaType}</span>
                        </div>
                      </td>

                      {/* Official Permit Status Badge (Reference Info) */}
                      <td className="p-3">
                        <div>
                          {isPermitted ? (
                            <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                              <CheckCircle2 size={12} />
                              <span>နေထိုင်ခွင့်ကျထားသောသူ</span>
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-amber-100 text-amber-800 border border-amber-300">
                              <Clock size={12} />
                              <span>မလျှောက်ထားသေးသူ</span>
                            </span>
                          )}

                          {r.permittedBy && (
                            <div className="text-[10px] text-slate-500 font-medium truncate max-w-xs mt-1">
                              အမိန့်စာ: {r.permittedBy}
                            </div>
                          )}
                        </div>
                      </td>

                      {/* Physical Location Inspection State */}
                      <td className="p-3">
                        {isVerified ? (
                          <div className="space-y-0.5">
                            <div className="inline-flex items-center gap-1 text-[11px] font-black text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded-md">
                              <Check size={13} />
                              <span>စစ်ဆေးပြီး ({r.latestCheck?.timestamp})</span>
                            </div>
                            <div className="text-[10px] text-slate-500 font-medium">
                              စစ်ဆေးသူ: {r.latestCheck?.officerName || 'Officer'}
                            </div>
                          </div>
                        ) : (
                          <span className="text-[11px] font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded-md border border-amber-200">
                            ⏳ မစစ်ဆေးရသေးပါ
                          </span>
                        )}
                      </td>

                      {/* Single Action */}
                      <td className="p-3 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          <button
                            onClick={() => {
                              setSingleRelocatePassport(r.passport);
                              setIsRelocatingBatch(false);
                              setRelocateTargetAddress(r.address || '');
                              setRelocateTargetDescription(r.stayDescription || '');
                              setIsRelocateModalOpen(true);
                            }}
                            className="px-2.5 py-1.5 rounded-xl text-xs font-black bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 flex items-center gap-1 transition-all cursor-pointer shadow-xs"
                            title="ဤသူအား နေရာပြောင်းရွှေ့မည်"
                          >
                            <MapPin size={13} className="text-amber-600" />
                            <span>နေရာပြောင်း</span>
                          </button>

                          <button
                            onClick={() => handleVerifyResident(r)}
                            className={`px-3 py-1.5 rounded-xl text-xs font-black flex items-center justify-end gap-1.5 transition-all cursor-pointer shadow-xs ${
                              isVerified 
                                ? 'bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200' 
                                : 'bg-emerald-600 hover:bg-emerald-500 text-white'
                            }`}
                          >
                            <CheckCircle2 size={13} />
                            <span>{isVerified ? 'ပြန်စစ်မည်' : 'စစ်ဆေးပြီး'}</span>
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Relocation Modal */}
      {isRelocateModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs">
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-lg p-6 space-y-4">
            <div className="flex items-center justify-between border-b pb-3">
              <div className="flex items-center gap-2">
                <RefreshCw size={18} className="text-amber-600" />
                <h3 className="text-base font-black uppercase text-slate-900">
                  တည်းခိုနေထိုင်သည့် နေရာ ပြောင်းရွှေ့ခြင်း
                </h3>
              </div>
              <button 
                onClick={() => {
                  setIsRelocateModalOpen(false);
                  setSingleRelocatePassport(null);
                }}
                className="p-1 text-slate-400 hover:text-slate-600"
              >
                <X size={18} />
              </button>
            </div>

            <p className="text-xs text-slate-600 font-medium leading-relaxed">
              {singleRelocatePassport
                ? `ရွေးချယ်ထားသော နိုင်ငံခြားသား (${singleRelocatePassport}) အား တည်းခိုနေရာအသစ်သို့ ပြောင်းရွှေ့ပေးပါမည်။`
                : isRelocatingBatch 
                ? `ရွေးချယ်ထားသော လူဦးရေ (${selectedPassports.length}) ဦးအား တည်းခိုနေရာအသစ်သို့ ပြောင်းရွှေ့ပေးပါမည်။` 
                : `"${selectedLocation?.name}" ရှိ နေထိုင်သူများ အားလုံး (${selectedLocation?.residents.length || 0} ဦး) အား တည်းခိုနေရာအသစ်သို့ ပြောင်းရွှေ့ပေးပါမည်။`}
            </p>

            <div className="pt-1">
              <MasterLocationInput
                locationValue={relocateTargetAddress}
                onLocationChange={setRelocateTargetAddress}
                descriptionValue={relocateTargetDescription}
                onDescriptionChange={setRelocateTargetDescription}
                masterData={masterData}
                locationLabel="ပြောင်းရွှေ့မည့် နေရာ / လိပ်စာ အသစ် *"
                descriptionLabel="လိပ်စာ အသေးစိတ် (Stay Description / Detail)"
                locationPlaceholder="ဥပမာ- Tidy Co., Ltd (သို့) ဟိုတယ်အမည် ရွေးချယ်ပါ..."
                required
                idPrefix="relocate_stay"
              />
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-3 border-t">
              <button
                type="button"
                onClick={() => {
                  setIsRelocateModalOpen(false);
                  setSingleRelocatePassport(null);
                }}
                className="px-4 py-2 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100"
              >
                ပယ်ဖျက်မည်
              </button>
              <button
                type="button"
                onClick={handleExecuteRelocation}
                className="px-5 py-2 rounded-xl bg-amber-600 hover:bg-amber-500 text-white text-xs font-black shadow-md cursor-pointer"
              >
                နေရာပြောင်းရွှေ့မည်
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
