import React, { useState, useMemo, useEffect } from 'react';
import { 
  Search, 
  FileText, 
  Download, 
  Printer, 
  User, 
  Mail, 
  Calendar, 
  MapPin, 
  CheckCircle2, 
  AlertTriangle, 
  ShieldAlert, 
  X, 
  Plus, 
  Trash2, 
  Edit, 
  AlertCircle, 
  Check, 
  Users, 
  FileSpreadsheet, 
  Activity, 
  Clipboard,
  ShieldCheck,
  ChevronRight,
  Save,
  Eye,
  Building2
} from 'lucide-react';
import { ImmRecord, MasterItem, DossierRecord, CheckingHistoryEntry, MovementData, normalizePermitStatus, PermitStatus } from '../types';
import { logActivity } from '../utils/activityLogger';
import { StayCheckConsole } from './StayCheckConsole';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';

interface IndividualSearchProps {
  records: ImmRecord[];
  setRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  tempRecords?: ImmRecord[];
  setTempRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  checkingHistory: any[];
  setCheckingHistory?: React.Dispatch<React.SetStateAction<CheckingHistoryEntry[]>>;
  dailyPdfs: Record<string, { base64: string; name: string }>;
  showToast: (msg: string) => void;
  setActivePrintPreview?: (preview: any) => void;
  setAnalyzerTarget?: (passport: string | null) => void;
  dossierHistory?: DossierRecord[];
  setDossierHistory?: React.Dispatch<React.SetStateAction<DossierRecord[]>>;
  currentUser?: { title?: string; name?: string } | null;
  cloudAuthUser?: { username: string; role: string } | null;
  masterData?: MasterItem[];
  movementMap?: Record<string, any>;
  passportToLatestInfo?: Record<string, any>;
  isViewer?: boolean;
}

const parseTimestamp = (str: string): number => {
  if (!str) return 0;
  try {
    if (str.includes('T')) {
      const p = Date.parse(str);
      if (!isNaN(p)) return p;
    }
    const nums = str.match(/\d+/g)?.map(Number);
    if (!nums || nums.length < 3) return Date.parse(str) || 0;
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

const parseLogDateToKey = (timestamp: string) => {
  if (!timestamp) return null;
  const parts = timestamp.split(', ')[0].split('/');
  if (parts.length === 3) {
    const [d, m, y] = parts;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
};

// Help calculate remaining days numerically
const getRemainingDaysNum = (dateStr: string): number => {
  if (!dateStr) return 0;
  try {
    const targetDate = new Date(dateStr);
    targetDate.setHours(23, 59, 59, 999);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const diffTime = targetDate.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  } catch {
    return 0;
  }
};

// Help calculate actual elapsed days from check-in
const getDaysStayedSoFar = (startDateStr: string): number => {
  if (!startDateStr) return 0;
  try {
    const startDate = new Date(startDateStr);
    startDate.setHours(0, 0, 0, 0);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const diffTime = today.getTime() - startDate.getTime();
    return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
  } catch {
    return 0;
  }
};

// Helper to determine if a guest is currently staying (STILL_IN) or departed (DEPARTED)
const getGuestStayStatus = (passport: string, computedMovementMap: Record<string, any>) => {
  const pp = (passport || '').trim().toUpperCase();
  const m = computedMovementMap ? computedMovementMap[pp] : null;
  if (!m) return { status: 'UNKNOWN', label: 'မှတ်တမ်းမရှိ', subLabel: 'UNKNOWN', timeText: '', isStillIn: false, isDeparted: false };

  const isStillIn = Boolean(m.in && !m.out);
  const isDeparted = Boolean((m.out && m.outTime >= m.inTime) || (!m.in && m.out));

  if (isStillIn) {
    return {
      status: 'STILL_IN',
      label: 'လက်ရှိနေထိုင်သူ',
      subLabel: 'STILL IN',
      timeText: m.in ? `ဝင်: ${m.in}` : '',
      isStillIn: true,
      isDeparted: false
    };
  } else if (isDeparted) {
    return {
      status: 'DEPARTED',
      label: 'ထွက်သွားပြီးသူ',
      subLabel: 'DEPARTED',
      timeText: m.out ? `ထွက်: ${m.out}` : '',
      isStillIn: false,
      isDeparted: true
    };
  } else {
    return {
      status: 'REGISTERED',
      label: 'မှတ်ပုံတင်ထားသူ',
      subLabel: 'REGISTERED',
      timeText: '',
      isStillIn: false,
      isDeparted: false
    };
  }
};

export const IndividualSearch = ({
  records,
  setRecords,
  tempRecords,
  setTempRecords,
  checkingHistory,
  setCheckingHistory,
  dailyPdfs,
  showToast,
  setActivePrintPreview,
  setAnalyzerTarget,
  dossierHistory = [],
  setDossierHistory,
  currentUser,
  cloudAuthUser,
  masterData = [],
  movementMap = {},
  passportToLatestInfo = {},
  isViewer
}: IndividualSearchProps) => {
  const [invSubTab, setInvSubTab] = useState<'STAY_CHECK' | 'DOSSIER'>('STAY_CHECK');
  const [searchQuery, setSearchQuery] = useState('');

  // Collect all selectable stay addresses from Master Data and existing records
  const stayAddressOptions = useMemo(() => {
    const set = new Set<string>();
    if (Array.isArray(masterData)) {
      masterData.filter(m => m.type === 'Stay').forEach(m => {
        if (m.name?.trim()) set.add(m.name.trim());
        if (m.linkedValue?.trim()) set.add(m.linkedValue.trim());
      });
    }
    if (Array.isArray(records)) {
      records.forEach(r => {
        if (r.address?.trim()) set.add(r.address.trim());
      });
    }
    return Array.from(set);
  }, [masterData, records]);
  const [selectedPassports, setSelectedPassports] = useState<string[]>([]);
  const [currentRemarks, setCurrentRemarks] = useState<Record<string, string>>({});
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [showPrintPreviewModal, setShowPrintPreviewModal] = useState(false);
  const [activeDossierID, setActiveDossierID] = useState<string>('');
  const [verificationInput, setVerificationInput] = useState<string>('');

  // Auto-timestamp note inputs per passport
  const [newNoteInputs, setNewNoteInputs] = useState<Record<string, string>>({});

  // Direct Inline Stay Editing state per passport
  const [editingStayPassport, setEditingStayPassport] = useState<string | null>(null);
  const [inlineStayEdits, setInlineStayEdits] = useState<Record<string, { loc: string; visa: string; start: string; end: string; allowed: string }>>({});

  // Helper to handle auto-timestamp note addition
  const handleAddAutoNoteForPassport = (pp: string) => {
    const targetPp = pp.toUpperCase();
    const input = (newNoteInputs[targetPp] || '').trim();
    if (!input) {
      showToast("PLEASE ENTER NOTE CONTENT BEFORE ADDING");
      return;
    }

    const matched = computedMovementMap[targetPp] || {};
    const now = new Date();
    const timestampStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
    
    const officerTitle = currentUser?.title || '';
    const officerName = currentUser?.name || 'Officer';
    const authorTag = officerTitle ? `${officerTitle} ${officerName}` : officerName;

    const formattedNote = `[${timestampStr} • ${authorTag}]: ${input}`;
    
    const currentText = currentRemarks[targetPp] !== undefined 
      ? currentRemarks[targetPp] 
      : (matched.remarks || '');

    const updatedRemarks = currentText 
      ? `${currentText}\n${formattedNote}` 
      : formattedNote;

    saveRemarkForPassport(targetPp, updatedRemarks);
    setNewNoteInputs(prev => ({ ...prev, [targetPp]: '' }));
    showToast(`NOTE ADDED WITH TIMESTAMP FOR ${targetPp}`);
  };

  // Helper to start inline editing stay details
  const startEditingStayDetails = (pp: string, matchedLog: any) => {
    const targetPp = pp.toUpperCase();
    setInlineStayEdits(prev => ({
      ...prev,
      [targetPp]: {
        loc: prev[targetPp]?.loc ?? (matchedLog.loc || ''),
        visa: prev[targetPp]?.visa ?? (matchedLog.visa || ''),
        start: prev[targetPp]?.start ?? (matchedLog.start || ''),
        end: prev[targetPp]?.end ?? (matchedLog.end || ''),
        allowed: prev[targetPp]?.allowed ?? (matchedLog.allowed || '')
      }
    }));
    setEditingStayPassport(targetPp);
  };

  // Helper to handle saving inline stay details directly in INV card
  const handleSaveInlineStayDetails = (pp: string) => {
    if (isViewer) {
      showToast("🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် အချက်အလက် ပြင်ဆင်ခွင့် ပိတ်ထားပါသည် (Read-Only)");
      return;
    }
    const targetPp = pp.toUpperCase();
    const editData = inlineStayEdits[targetPp];
    if (!editData) return;

    if (setRecords) {
      setRecords(prev => prev.map(r => {
        if (r.passport.toUpperCase() === targetPp) {
          return {
            ...r,
            address: editData.loc,
            visaType: editData.visa,
            stayFrom: editData.start,
            stayTo: editData.end,
            totalDays: editData.allowed
          };
        }
        return r;
      }));
    }

    if (setTempRecords) {
      setTempRecords(prev => prev.map(r => {
        if (r.passport.toUpperCase() === targetPp) {
          return {
            ...r,
            address: editData.loc,
            visaType: editData.visa,
            stayFrom: editData.start,
            stayTo: editData.end,
            totalDays: editData.allowed
          };
        }
        return r;
      }));
    }

    setEditingStayPassport(null);
    showToast(`UPDATED STAY DETAILS FOR ${targetPp}`);
  };

  // Auto-generate active Dossier ID whenever guests are selected
  useEffect(() => {
    if (selectedPassports.length > 0) {
      const now = new Date();
      const yyyymmdd = now.getFullYear().toString() + 
                       (now.getMonth() + 1).toString().padStart(2, '0') + 
                       now.getDate().toString().padStart(2, '0');
      const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
      let rand = '';
      for (let i = 0; i < 4; i++) {
        rand += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      setActiveDossierID(`IMD-${yyyymmdd}-${rand}`);
    } else {
      setActiveDossierID('');
    }
  }, [selectedPassports.length]);
  
  // Tab for the Alerts Panel (OVERSTAY vs LONG_STAY)
  const [activeAlertTab, setActiveAlertTab] = useState<'OVERSTAY' | 'LONG_STAY'>('OVERSTAY');
  const [longStayThresholdDays, setLongStayThresholdDays] = useState<number>(60);
  
  // Quick Edit Modal state
  const [quickEditPassport, setQuickEditPassport] = useState<string | null>(null);
  const [editFullName, setEditFullName] = useState('');
  const [editNationality, setEditNationality] = useState('');
  const [editGender, setEditGender] = useState<'M' | 'F' | ''>('');
  const [editDob, setEditDob] = useState('');
  const [editVisaNumber, setEditVisaNumber] = useState('');
  const [editAddress, setEditAddress] = useState('');
  const [editVisaType, setEditVisaType] = useState('');
  const [editStayFrom, setEditStayFrom] = useState('');
  const [editStayTo, setEditStayTo] = useState('');
  const [editTotalDays, setEditTotalDays] = useState('');

  // Re-compute active movement data map in real-time
  const computedMovementMap = useMemo(() => {
    const map: Record<string, any> = {};
    [...records]
      .sort((a, b) => {
        const diff = parseTimestamp(a.timestamp) - parseTimestamp(b.timestamp);
        if (diff !== 0) return diff;
        if (a.mode === 'IN' && b.mode === 'OUT') return -1;
        if (a.mode === 'OUT' && b.mode === 'IN') return 1;
        return 0;
      })
      .forEach(r => {
        const pp = r.passport.toUpperCase();
        if (!map[pp]) {
          map[pp] = { 
            p: pp, n: r.fullname, nat: r.nationality, loc: r.address, 
            visa: r.visaType, gender: r.gender, start: r.stayFrom, 
            end: r.stayTo, allowed: r.totalDays, vInfo: r.vehicleInfo, 
            agent: r.broughtBy || '',
            contact: r.contactDetails || '',
            in: '', out: '', inTime: 0, outTime: 0, 
            latestTime: 0, lastId: r.id,
            stayDescription: r.stayDescription || '',
            stillPermittedStatus: r.stillPermittedStatus || '',
            permittedBy: r.permittedBy || '',
            remarks: r.remarks || ''
          };
        }
        const recTime = parseTimestamp(r.timestamp);
        if (r.mode === 'IN') {
          Object.assign(map[pp], { 
            in: r.timestamp, inTime: recTime, out: '', 
            outTime: 0, loc: r.address, visa: r.visaType, 
            agent: r.broughtBy || '',
            contact: r.contactDetails || map[pp].contact,
            start: r.stayFrom, end: r.stayTo, allowed: r.totalDays, 
            lastId: r.id,
            stayDescription: r.stayDescription || map[pp].stayDescription,
            stillPermittedStatus: r.stillPermittedStatus || map[pp].stillPermittedStatus,
            permittedBy: r.permittedBy || map[pp].permittedBy,
            remarks: r.remarks || map[pp].remarks
          });
        } else if (r.mode === 'OUT' && recTime >= map[pp].inTime) {
          Object.assign(map[pp], {
            out: r.timestamp,
            outTime: recTime,
            lastId: r.id
          });
        }
        if (recTime > map[pp].latestTime) {
          map[pp].latestTime = recTime;
          if (r.fullname) map[pp].n = r.fullname;
          if (r.nationality) map[pp].nat = r.nationality;
          if (r.gender) map[pp].gender = r.gender;
          if (r.broughtBy) map[pp].agent = r.broughtBy;
          if (r.contactDetails) map[pp].contact = r.contactDetails;
          if (r.remarks) map[pp].remarks = r.remarks;
        }
      });
    return map;
  }, [records]);

  // Compute active checked-in guests
  const activeStayedIn = useMemo(() => {
    return Object.values(computedMovementMap).filter((m: any) => m.in && !m.out);
  }, [computedMovementMap]);

  // Overstayed: Active guests with remaining days < 0
  const overstayGuests = useMemo(() => {
    return activeStayedIn.filter((m: any) => getRemainingDaysNum(m.end) < 0);
  }, [activeStayedIn]);

  // Long-Term: Active guests checked in >= threshold days ago
  const longStayGuests = useMemo(() => {
    return activeStayedIn.filter((m: any) => getDaysStayedSoFar(m.start) >= longStayThresholdDays);
  }, [activeStayedIn, longStayThresholdDays]);

  // Group records by passport to get unique people for search bar suggestions
  const uniquePeople = useMemo(() => {
    const map: Record<string, { passport: string; previousPassport?: string; dualPassportRemarks?: string; fullname: string; nationality: string; gender: string; address: string; latestTime: number }> = {};
    records.forEach(r => {
      const p = r.passport.trim().toUpperCase();
      if (!p) return;
      const t = parseTimestamp(r.timestamp);
      if (!map[p] || t > map[p].latestTime) {
        map[p] = {
          passport: r.passport,
          previousPassport: r.previousPassport,
          dualPassportRemarks: r.dualPassportRemarks,
          fullname: r.fullname,
          nationality: r.nationality,
          gender: r.gender,
          address: r.address || '',
          latestTime: t
        };
      }
    });
    return Object.values(map);
  }, [records]);

  // Filter unique people by query (passport, previous passport, name, nationality, or stay address)
  const filteredPeople = useMemo(() => {
    if (!searchQuery.trim()) return [];
    const q = searchQuery.toLowerCase();
    return uniquePeople.filter(
      p => p.passport.toLowerCase().includes(q) || 
           (p.previousPassport && p.previousPassport.toLowerCase().includes(q)) ||
           p.fullname.toLowerCase().includes(q) ||
           (p.nationality && p.nationality.toLowerCase().includes(q)) ||
           (p.address && p.address.toLowerCase().includes(q))
    );
  }, [uniquePeople, searchQuery]);

  // Add passport to current investigation group
  const addPassportToGroup = (passport: string) => {
    const pp = passport.trim().toUpperCase();
    if (!pp) return;
    if (!selectedPassports.includes(pp)) {
      setSelectedPassports(prev => [...prev, pp]);
      
      // Initialize remark in currentRemarks
      const matched = computedMovementMap[pp];
      if (matched && matched.remarks) {
        setCurrentRemarks(prev => ({ ...prev, [pp]: matched.remarks }));
      }
      showToast(`SUCCESSFULLY ADDED ${pp} TO INVESTIGATION GROUP`);
    } else {
      showToast(`${pp} IS ALREADY IN THE INVESTIGATION GROUP`);
    }
    setSearchQuery('');
  };

  // Remove passport from investigation group
  const removePassportFromGroup = (passport: string) => {
    setSelectedPassports(prev => prev.filter(p => p !== passport));
    showToast(`REMOVED ${passport} FROM INVESTIGATION GROUP`);
  };

  // Open Quick Edit modal
  const openQuickEdit = (passport: string) => {
    if (isViewer) {
      showToast("🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် အချက်အလက် ပြင်ဆင်ခွင့် ပိတ်ထားပါသည် (Read-Only)");
      return;
    }
    const pp = passport.toUpperCase();
    const matched = records.find(r => r.passport.toUpperCase() === pp) || 
                    tempRecords?.find(r => r.passport.toUpperCase() === pp) || 
                    computedMovementMap[pp];
    if (matched) {
      setQuickEditPassport(pp);
      setEditFullName(matched.fullname || '');
      setEditNationality(matched.nationality || '');
      setEditGender(matched.gender || '');
      setEditDob(matched.dob || '');
      setEditVisaNumber(matched.visaNumber || (matched as any).visaNo || '');
      setEditAddress(matched.address || '');
      setEditVisaType(matched.visaType || '');
      setEditStayFrom(matched.stayFrom || '');
      setEditStayTo(matched.stayTo || '');
      setEditTotalDays(matched.totalDays || '');
    } else {
      showToast("GUEST HISTORY LOG NOT FOUND");
    }
  };

  // Save quick-edit updates
  const handleQuickSave = () => {
    if (isViewer) {
      showToast("🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် အချက်အလက် ပြင်ဆင်ခွင့် ပိတ်ထားပါသည် (Read-Only)");
      return;
    }
    if (!quickEditPassport) return;
    const pp = quickEditPassport.toUpperCase();

    if (setRecords) {
      setRecords(prev => prev.map(r => {
        if (r.passport.toUpperCase() === pp) {
          return {
            ...r,
            fullname: editFullName,
            nationality: editNationality,
            gender: editGender,
            dob: editDob,
            visaNumber: editVisaNumber,
            address: editAddress,
            visaType: editVisaType,
            stayFrom: editStayFrom,
            stayTo: editStayTo,
            totalDays: editTotalDays,
            updatedAt: new Date().toISOString()
          };
        }
        return r;
      }));
    }

    if (setTempRecords) {
      setTempRecords(prev => prev.map(r => {
        if (r.passport.toUpperCase() === pp) {
          return {
            ...r,
            fullname: editFullName,
            nationality: editNationality,
            gender: editGender,
            dob: editDob,
            visaNumber: editVisaNumber,
            address: editAddress,
            visaType: editVisaType,
            stayFrom: editStayFrom,
            stayTo: editStayTo,
            totalDays: editTotalDays,
            updatedAt: new Date().toISOString()
          };
        }
        return r;
      }));
    }

    logActivity({
      action: 'UPDATE',
      module: 'FFE',
      officerName: currentUser?.name,
      targetId: pp,
      details: `Quick Edit profile updated for Passport ${pp} (${editFullName || 'N/A'}, ${editNationality || 'N/A'}) - DOB: ${editDob || '-'}, Visa No: ${editVisaNumber || '-'}, Visa: ${editVisaType || '-'}, Stay: ${editStayFrom} to ${editStayTo}`
    });

    showToast(`UPDATED ALL HISTORICAL REGISTER RECORDS FOR ${pp}`);
    setQuickEditPassport(null);
  };

  // Save custom remark for a specific passport
  const saveRemarkForPassport = (pp: string, remarkText: string) => {
    if (isViewer) {
      showToast("🔒 Viewer Level သတ်မှတ်ထားသောကြောင့် မှတ်ချက် ပြင်ဆင်ခွင့် ပိတ်ထားပါသည် (Read-Only)");
      return;
    }
    if (setRecords) {
      setRecords(prev => prev.map(r => r.passport.toUpperCase() === pp.toUpperCase() ? { ...r, remarks: remarkText, updatedAt: new Date().toISOString() } : r));
    }
    if (setTempRecords) {
      setTempRecords(prev => prev.map(r => r.passport.toUpperCase() === pp.toUpperCase() ? { ...r, remarks: remarkText, updatedAt: new Date().toISOString() } : r));
    }
    setCurrentRemarks(prev => ({ ...prev, [pp]: remarkText }));
    logActivity({
      action: 'UPDATE',
      module: 'FFE',
      officerName: currentUser?.name,
      targetId: pp,
      details: `Dossier remarks updated for Passport ${pp}: "${remarkText}"`
    });
    showToast(`SAVED DOSSIER REMARK FOR ${pp}`);
  };

  // Export warning lists to Excel
  const handleExportWarningsExcel = () => {
    const targetData = activeAlertTab === 'OVERSTAY' ? overstayGuests : longStayGuests;
    if (targetData.length === 0) {
      showToast("No warning records available to export");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();
      const rows = targetData.map((m: any, idx) => ({
        "Sr. No.": idx + 1,
        "Passport Number": m.p,
        "Guest Full Name": m.n,
        "Nationality": m.nat,
        "Gender": m.gender === 'M' ? 'MALE' : m.gender === 'F' ? 'FEMALE' : '-',
        "Residence Stay Address": m.loc,
        "Address Detail Description": m.stayDescription || '-',
        "Visa Type": m.visa,
        "Stay Start From": m.start,
        "Stay Expired To": m.end,
        "Allowed Duration Days": m.allowed,
        "Remaining Permitted Days": getRemainingDaysNum(m.end),
        "Elapsed Days Stayed So Far": getDaysStayedSoFar(m.start),
        "Local Broker / Sponsor": m.agent || '-',
        "Broker Contact Details": m.contact || '-'
      }));

      const sheetName = activeAlertTab === 'OVERSTAY' ? "Overstay Warning List" : "Long Stayers Warning List";
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      const filename = activeAlertTab === 'OVERSTAY'
        ? `IMM_Overstay_Permits_Warnings_${new Date().toISOString().split('T')[0]}.xlsx`
        : `IMM_Long_Term_Residence_Warnings_${new Date().toISOString().split('T')[0]}.xlsx`;

      XLSX.writeFile(wb, filename);
      showToast("EXCEL DOSSIER WARNER SPREADSHEET DOWNLOADED");
    } catch (e) {
      console.error(e);
      showToast("EXCEL SPREADSHEET EXPORT FAILURE");
    }
  };

  // Export Consolidated Group Dossiers to Excel
  const handleExportGroupExcel = () => {
    if (selectedPassports.length === 0) return;

    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Group Comparison Matrix Summary
      const summaryRows = selectedPassports.map((pp, idx) => {
        const matched = computedMovementMap[pp] || { p: pp, n: '-', nat: '-', gender: '-', loc: '-', visa: '-', start: '-', end: '-', allowed: '-', agent: '', contact: '' };
        const statusInfo = getGuestStayStatus(pp, computedMovementMap);
        return {
          "Sr No": idx + 1,
          "Passport": matched.p,
          "Full Name": matched.n,
          "Stay Status (နေထိုင်မှုအခြေအနေ)": statusInfo.label + (statusInfo.subLabel ? ` (${statusInfo.subLabel})` : ''),
          "Nationality": matched.nat,
          "Gender": matched.gender === 'M' ? 'MALE' : matched.gender === 'F' ? 'FEMALE' : matched.gender || '-',
          "Registered Address": matched.loc,
          "Visa Classification": matched.visa,
          "Validity Schedule": `${matched.start} to ${matched.end}`,
          "Allowed Days": matched.allowed,
          "Remaining Days": getRemainingDaysNum(matched.end),
          "Local Sponsor / Caller": matched.agent || '-',
          "Sponsor Contact details": matched.contact || '-',
          "Custom Investigation Remarks": currentRemarks[pp] || matched.remarks || 'No special remarks.'
        };
      });

      const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
      XLSX.utils.book_append_sheet(wb, wsSummary, "Group Overview Matrix");

      // Sheet 2: Consolidated Travel History
      const consolidatedLogs: any[] = [];
      selectedPassports.forEach(pp => {
        const logs = records
          .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
          .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));
        
        logs.forEach((log, index) => {
          consolidatedLogs.push({
            "Investigated Passport": pp,
            "Log Index": index + 1,
            "Log Date & Time": log.timestamp,
            "Movement Direction": log.mode,
            "Full Name": log.fullname,
            "Visa Type": log.visaType,
            "Stay From": log.stayFrom,
            "Stay To": log.stayTo,
            "Days Allowed": log.totalDays,
            "Stay Location Address": log.address,
            "Detailed Stay Description": log.stayDescription || '-',
            "Sponsor / Agent": log.broughtBy || '-',
            "Sponsor Contact Details": log.contactDetails || '-',
            "Transport vehicle Carrier": log.vehicleInfo || '-',
            "Origin Arrived From": log.arrivedFrom || '-',
            "Destination Departed To": log.departedTo || '-',
            "Registered Officer": `${log.officialTitle || ''} ${log.officialName || ''}`,
            "Permit verification status": log.stillPermittedStatus || 'CONFIRMED'
          });
        });
      });

      const wsLogs = XLSX.utils.json_to_sheet(consolidatedLogs);
      XLSX.utils.book_append_sheet(wb, wsLogs, "All Movement Logs");

      XLSX.writeFile(wb, `IMM_Group_Investigation_Dossiers_Count_${selectedPassports.length}.xlsx`);
      showToast("CONSOLIDATED GROUP INVESTIGATION EXCEL EXPORTED");
    } catch (e) {
      console.error(e);
      showToast("GROUP EXCEL EXPORT FAILURE");
    }
  };

  // 100% Bulletproof Native HTML A4 Printing pattern
  const handlePrintGroupDossier = () => {
    if (selectedPassports.length === 0) {
      showToast("NO PASSPORTS SELECTED FOR PRINTING");
      return;
    }

    saveActiveDossierToHistory();

    const sourceEl = document.getElementById('temp-print-content-source');
    if (!sourceEl) {
      showToast("PRINT DATA SOURCE NOT FOUND");
      return;
    }

    // Reuse or create #temp-print-area directly on document.body
    let printArea = document.getElementById('temp-print-area');
    if (!printArea) {
      printArea = document.createElement('div');
      printArea.id = 'temp-print-area';
      document.body.appendChild(printArea);
    }

    // Reset inner content
    printArea.innerHTML = '';

    // Printable style overrides for high-contrast A4 print output
    const styleEl = document.createElement('style');
    styleEl.innerHTML = `
      @media print {
        @page { size: A4 portrait; margin: 0.5in; }
        body > :not(#temp-print-area), .no-print { display: none !important; }
        #temp-print-area { display: block !important; visibility: visible !important; width: 100% !important; background: #ffffff !important; color: #000000 !important; }
        .print-page-break { page-break-before: always !important; break-before: page !important; }
        table { page-break-inside: avoid !important; width: 100% !important; border-collapse: collapse !important; }
        tr, td, th { page-break-inside: avoid !important; }
      }
    `;
    printArea.appendChild(styleEl);

    // Clone sourceEl content and unhide for print preview rasterizer
    const contentClone = sourceEl.cloneNode(true) as HTMLElement;
    contentClone.style.display = 'block';
    contentClone.style.visibility = 'visible';
    contentClone.classList.remove('hidden');

    printArea.appendChild(contentClone);

    showToast("OPENING PRINT PREVIEW...");

    setTimeout(() => {
      window.focus();
      window.print();
    }, 250);
  };

  const saveActiveDossierToHistory = () => {
    if (!activeDossierID || selectedPassports.length === 0) return;
    const isSaved = dossierHistory.some(d => d.id === activeDossierID);
    if (!isSaved) {
      const newRecord: DossierRecord = {
        id: activeDossierID,
        generatedAt: new Date().toLocaleString(),
        passports: [...selectedPassports],
        fileType: 'DOCX/HTML',
        dossierTitle: selectedPassports.length === 1 
          ? `Individual Investigation: ${selectedPassports[0]}` 
          : `Group Investigation: ${selectedPassports.length} Guests`,
        remarksMap: { ...currentRemarks }
      };
      if (setDossierHistory) {
        setDossierHistory(prev => {
          const updated = [newRecord, ...prev];
          localStorage.setItem('imm_dossier_history_react', JSON.stringify(updated));
          return updated;
        });
      }
    }
  };

  // High-fidelity Microsoft Word (.docx) document exporter
  const handleExportGroupWordDoc = () => {
    if (selectedPassports.length === 0) {
      showToast("NO PASSPORTS SELECTED FOR EXPORT");
      return;
    }

    saveActiveDossierToHistory();
    showToast("GENERATING FRESH WORD DOCUMENT...");

    let htmlContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <title>Immigration Group Investigation Dossier</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            font-size: 10.5pt;
            line-height: 1.5;
            color: #334155;
            background-color: #ffffff;
            margin: 20px;
          }
          .dossier-header {
            text-align: center;
            border-bottom: 3px double #1e3a8a;
            padding-bottom: 15px;
            margin-bottom: 25px;
          }
          .mya-header {
            font-size: 15pt;
            font-weight: bold;
            color: #1e3a8a;
            margin: 0;
            letter-spacing: 0.5px;
          }
          .eng-header {
            font-size: 11pt;
            font-weight: 600;
            color: #475569;
            margin: 4px 0 0 0;
            text-transform: uppercase;
          }
          .dossier-title-badge {
            display: inline-block;
            background-color: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1e40af;
            font-weight: bold;
            font-size: 10pt;
            padding: 4px 14px;
            border-radius: 100px;
            margin-top: 8px;
            text-transform: uppercase;
          }
          .meta-info {
            font-size: 8.5pt;
            color: #1e3a8a;
            font-weight: bold;
            margin-top: 10px;
          }
          h3 {
            font-size: 11.5pt;
            font-weight: bold;
            color: #0f172a;
            background-color: #f8fafc;
            border-left: 4px solid #1e3a8a;
            padding: 6px 12px;
            margin-top: 25px;
            margin-bottom: 12px;
            text-transform: uppercase;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
          }
          th {
            background-color: #1e3a8a;
            color: #ffffff;
            border: 1px solid #1e3a8a;
            font-weight: bold;
            font-size: 9pt;
            padding: 8px 10px;
            text-align: left;
            text-transform: uppercase;
          }
          td {
            border: 1px solid #e2e8f0;
            font-size: 9pt;
            padding: 8px 10px;
            vertical-align: top;
          }
          tr:nth-child(even) td {
            background-color: #f8fafc;
          }
          .text-center {
            text-align: center;
          }
          .font-bold {
            font-weight: bold;
          }
          .font-mono {
            font-family: Consolas, Monaco, "Courier New", Courier, monospace;
          }
          .profile-container {
            border: 1px solid #e2e8f0;
            border-left: 5px solid #2563eb;
            background-color: #f8fafc;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
          }
          .profile-grid {
            width: 100%;
            border-collapse: collapse;
          }
          .profile-grid td {
            border: none;
            padding: 6px 10px;
            background-color: transparent !important;
          }
          .label {
            color: #64748b;
            font-size: 8pt;
            text-transform: uppercase;
            font-weight: bold;
            margin-bottom: 2px;
          }
          .value {
            font-size: 10pt;
            font-weight: bold;
            color: #0f172a;
          }
          .remark-box {
            background-color: #fffbeb;
            border: 1px solid #fde68a;
            border-left: 5px solid #d97706;
            padding: 12px 15px;
            margin: 15px 0;
            border-radius: 6px;
          }
          .remark-title {
            font-size: 8.5pt;
            font-weight: bold;
            color: #b45309;
            text-transform: uppercase;
            margin-bottom: 4px;
          }
          .remark-text {
            font-weight: bold;
            font-size: 10pt;
            color: #78350f;
          }
          .badge-green {
            background-color: #dcfce7;
            color: #15803d;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
          }
          .badge-red {
            background-color: #fee2e2;
            color: #b91c1c;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
          }
          .page-break {
            page-break-before: always;
          }
          .signature-table {
            margin-top: 35px;
            border: none;
          }
          .signature-table td {
            border: none;
            text-align: center;
            padding: 40px 10px 10px 10px;
            background-color: transparent !important;
          }
          .stamp-box {
            background-color: #f8fafc;
            border: 1px dashed #cbd5e1;
            padding: 15px;
            border-radius: 6px;
            text-align: left;
          }
        </style>
      </head>
      <body>
    `;

    // 1. First Page: Group Summary Matrix (only if multiple)
    if (selectedPassports.length > 1) {
      htmlContent += `
        <div class="dossier-header">
          <div class="mya-header">Department of Immigration (Myeik)</div>
          <div class="eng-header">Second Division, Department of Immigration</div>
          <div>
            <span class="dossier-title-badge">Group Stay Verification & Intelligence Report</span>
          </div>
          <div class="meta-info">
            DOSSIER IDENTIFICATION NO: ${activeDossierID} <br/>
            GENERATED: ${new Date().toLocaleString()} • CONFIDENTIAL OFFICIAL RECORDS COPY
          </div>
        </div>

        <h3>Group Comparison Overview Matrix</h3>
        <table>
          <thead>
            <tr>
              <th style="width: 4%; text-align: center;">#</th>
              <th style="width: 20%;">Guest Name / Passport</th>
              <th style="width: 14%; text-align: center;">Stay Status</th>
              <th style="width: 12%;">Nationality</th>
              <th style="width: 22%;">Residence Stay Address</th>
              <th style="width: 16%;">Visa & Validity</th>
              <th style="width: 12%;">Local Sponsor</th>
              <th style="width: 10%; text-align: center;">Remaining</th>
            </tr>
          </thead>
          <tbody>
      `;

      selectedPassports.forEach((pp, idx) => {
        const m = computedMovementMap[pp] || {};
        const rem = getRemainingDaysNum(m.end);
        const statusInfo = getGuestStayStatus(pp, computedMovementMap);
        const statusBadgeHtml = statusInfo.isStillIn
          ? `<span class="badge-green">🟢 STILL IN (လက်ရှိ)</span>`
          : statusInfo.isDeparted
          ? `<span class="badge-red">🔴 DEPARTED (ထွက်သွား)</span>`
          : `<span style="background-color:#f1f5f9; color:#475569; padding:2px 6px; border-radius:4px; font-weight:bold;">⚪ REGISTERED</span>`;

        htmlContent += `
          <tr>
            <td class="text-center font-bold">${idx + 1}</td>
            <td>
              <div class="font-bold font-mono" style="color: #1e3a8a; font-size: 9.5pt;">${pp}</div>
              <div style="font-size: 8.5pt; font-weight: bold; color: #334155;">${m.n || '-'}</div>
            </td>
            <td class="text-center">${statusBadgeHtml}</td>
            <td class="font-bold">${m.nat || '-'}</td>
            <td>${m.loc || '-'}</td>
            <td>
              <div class="font-bold">${m.visa || '-'}</div>
              <div style="font-size: 8pt; color: #64748b;">${m.start} to ${m.end}</div>
            </td>
            <td>${m.agent || '-'}</td>
            <td class="text-center font-bold font-mono">
              <span class="${rem < 0 ? 'badge-red' : 'badge-green'}">${rem} Days</span>
            </td>
          </tr>
        `;
      });

      htmlContent += `
          </tbody>
        </table>

        <!-- Signatures on Summary Page -->
        <table class="signature-table" style="width: 100%;">
          <tr>
            <td style="width: 33%;">
              <div style="border-top: 1px solid #94a3b8; padding-top: 8px; font-weight: bold; font-size: 9.5pt;">Prepared By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 3px;">Immigration Investigator</div>
            </td>
            <td style="width: 33%;">
              <div style="border-top: 1px solid #94a3b8; padding-top: 8px; font-weight: bold; font-size: 9.5pt;">Verified By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 3px;">Duty Commander</div>
            </td>
            <td style="width: 34%;">
              <div class="stamp-box">
                <div style="font-weight: bold; font-size: 8.5pt; color: #1e3a8a; margin-bottom: 6px;">Certified Stamp / Verification</div>
                <div style="font-size: 8pt; color: #64748b; font-style: italic;">Stamp Box / Verification Stamp Here</div>
              </div>
            </td>
          </tr>
        </table>

        <div class="page-break"></div>
      `;
    }

    // 2. Individual Pages
    selectedPassports.forEach((pp, idx) => {
      const bio = computedMovementMap[pp];
      if (!bio) return;

      const personLogs = records
        .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
        .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

      const personCheckingHistory = checkingHistory
        .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
        .sort((a, b) => b.id - a.id);

      htmlContent += `
        <div class="dossier-header">
          <div class="mya-header">Department of Immigration (Myeik)</div>
          <div class="eng-header">Individual Stay Verification Dossier</div>
          <div>
            <span class="dossier-title-badge">CONFIDENTIAL INVESTIGATION RECORD</span>
          </div>
          <div class="meta-info">
            DOSSIER IDENTIFICATION NO: ${activeDossierID} <br/>
            Dossier Record #${idx + 1} of ${selectedPassports.length} • Generated on ${new Date().toLocaleString()}
          </div>
        </div>

        <div class="profile-container">
          <table class="profile-grid">
            <tr>
              <td style="width: 50%;">
                <div class="label">Passport Number</div>
                <div class="value" style="font-size: 11pt; color: #1e3a8a;">🛂 ${pp}</div>
              </td>
              <td style="width: 50%;">
                <div class="label">Full Name</div>
                <div class="value">👤 ${bio.n}</div>
              </td>
            </tr>
            <tr>
              <td>
                <div class="label">Nationality & Gender</div>
                <div class="value">🌐 ${bio.nat || '-'} (${bio.gender === 'M' ? 'MALE' : 'FEMALE'})</div>
              </td>
              <td>
                <div class="label">Visa Category & Validity</div>
                <div class="value">📅 ${bio.visa || '-'} (${bio.start} to ${bio.end})</div>
              </td>
            </tr>
            <tr>
              <td colspan="2" style="padding-top: 10px;">
                <div class="label">Registered Destination Address</div>
                <div class="value" style="font-weight: normal; color: #0f172a;">📍 ${bio.loc || '-'}</div>
              </td>
            </tr>
            <tr>
              <td colspan="2" style="padding-top: 10px;">
                <div class="label">Address Description Details</div>
                <div class="value" style="font-weight: normal; font-style: italic; color: #475569;">💬 ${bio.stayDescription || 'No description logged.'}</div>
              </td>
            </tr>
            <tr>
              <td style="padding-top: 10px;">
                <div class="label">Local Broker/Sponsor</div>
                <div class="value">🏢 ${bio.agent || '-'}</div>
              </td>
              <td style="padding-top: 10px;">
                <div class="label font-mono">Sponsor Contact</div>
                <div class="value font-mono">📞 ${bio.contact || '-'}</div>
              </td>
            </tr>
          </table>
        </div>

        <div class="remark-box">
          <div class="remark-title">📝 Dossier Remark Details</div>
          <div class="remark-text">
            ${currentRemarks[pp] || bio.remarks || 'No special remarks registered for this profile.'}
          </div>
        </div>

        <!-- Flight logs -->
        <h3>✈️ Flight Movement Logs</h3>
        <table>
          <thead>
            <tr>
              <th style="width: 5%; text-align: center;">#</th>
              <th style="width: 20%;">Timestamp</th>
              <th style="width: 10%; text-align: center;">Direction</th>
              <th style="width: 25%;">Visa & Stay Schedule</th>
              <th style="width: 25%;">Stay Address</th>
              <th style="width: 15%;">Registered Officer</th>
            </tr>
          </thead>
          <tbody>
      `;

      if (personLogs.length === 0) {
        htmlContent += `
          <tr>
            <td colspan="6" class="text-center" style="color: #94a3b8; font-style: italic; padding: 15px;">No flight movement logs found</td>
          </tr>
        `;
      } else {
        personLogs.slice(0, 10).forEach((log, lIdx) => {
          htmlContent += `
            <tr>
              <td class="text-center">${lIdx + 1}</td>
              <td class="font-mono" style="font-size: 8.5pt;">${log.timestamp}</td>
              <td class="text-center font-bold">
                <span class="${log.mode === 'IN' ? 'badge-green' : 'badge-red'}">${log.mode}</span>
              </td>
              <td>
                <div class="font-bold" style="font-size: 8.5pt;">${log.visaType}</div>
                ${log.stayFrom ? `<div style="font-size: 7.5pt; color: #64748b; margin-top: 2px;">${log.stayFrom} to ${log.stayTo}</div>` : ''}
              </td>
              <td style="font-size: 8.5pt;">${log.address || '-'}</td>
              <td style="font-size: 8.5pt; font-weight: bold; color: #475569;">${log.officialTitle || ''} ${log.officialName || 'System'}</td>
            </tr>
          `;
        });
      }

      htmlContent += `
          </tbody>
        </table>

        <!-- Verification annals -->
        ${personCheckingHistory.length > 0 ? `
          <h3>🔍 Physical Verification Annals</h3>
          <table>
            <thead>
              <tr>
                <th style="width: 5%; text-align: center;">#</th>
                <th style="width: 20%;">Check Date</th>
                <th style="width: 22%;">Verified Status</th>
                <th style="width: 35%;">Confirmed Address</th>
                <th style="width: 18%;">Verifier</th>
              </tr>
            </thead>
            <tbody>
        ` : ''}
      `;

      if (personCheckingHistory.length > 0) {
        personCheckingHistory.forEach((ch, cIdx) => {
          const isNotMatch = ch.status?.includes('NOT');
          htmlContent += `
            <tr>
              <td class="text-center">${cIdx + 1}</td>
              <td class="font-mono" style="font-size: 8.5pt;">${ch.checkDate}</td>
              <td class="font-bold">
                <span class="${isNotMatch ? 'badge-red' : 'badge-green'}">${ch.status}</span>
              </td>
              <td style="font-size: 8.5pt;">${ch.confirmedAddress || '-'}</td>
              <td style="font-size: 8.5pt; font-weight: bold; color: #475569;">${ch.officerTitle || ''} ${ch.officerName || 'Officer'}</td>
            </tr>
          `;
        });

        htmlContent += `
            </tbody>
          </table>
        `;
      }

      htmlContent += `
        <!-- Signatures on Individual Page -->
        <table class="signature-table" style="width: 100%;">
          <tr>
            <td style="width: 50%;">
              <div style="border-top: 1px solid #cbd5e1; padding-top: 8px; font-weight: bold; font-size: 9pt;">Prepared & Authenticated By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 2px;">Investigation Analyst Officer</div>
            </td>
            <td style="width: 50%;">
              <div style="border-top: 1px solid #cbd5e1; padding-top: 8px; font-weight: bold; font-size: 9pt;">Reviewed & Certified By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 2px;">Duty Supervisor / Commander</div>
            </td>
          </tr>
        </table>
      `;

      // Page break between guests (except last)
      if (idx < selectedPassports.length - 1) {
        htmlContent += `<div class="page-break"></div>`;
      }
    });

    htmlContent += `
      </body>
      </html>
    `;

    // Download file as highly-compatible Word .doc format (Office 97-2003 standard)
    // Legacy .doc extension combined with 'application/msword' MIME type activates the universal HTML-in-Word fallback parser,
    // which is perfectly supported on mobile office viewers (WPS Office, Microsoft Word Mobile, Google Docs, iOS viewer etc.)
    // whereas modern .docx extension enforces strict zip archive verification and fails on raw HTML inputs.
    const blob = new Blob([htmlContent], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `IMM_Dossier_${activeDossierID}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("OFFICE WORD (DOC) EXPORTED SUCCESSFULLY!");
  };

  // Web Dossier (.html) document exporter for perfect mobile viewing, sharing, and native browser printing
  const handleExportGroupHTML = () => {
    if (selectedPassports.length === 0) {
      showToast("NO PASSPORTS SELECTED FOR EXPORT");
      return;
    }

    saveActiveDossierToHistory();
    showToast("GENERATING WEB REPORT...");

    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Immigration Group Investigation Dossier</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            font-size: 10.5pt;
            line-height: 1.5;
            color: #334155;
            background-color: #f1f5f9;
            margin: 0;
            padding: 15px;
          }
          .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
          }
          .dossier-header {
            text-align: center;
            border-bottom: 3px double #1e3a8a;
            padding-bottom: 15px;
            margin-bottom: 25px;
          }
          .mya-header {
            font-size: 14pt;
            font-weight: bold;
            color: #1e3a8a;
            margin: 0;
            line-height: 1.4;
          }
          .eng-header {
            font-size: 10.5pt;
            font-weight: 600;
            color: #475569;
            margin: 4px 0 0 0;
            text-transform: uppercase;
          }
          .dossier-title-badge {
            display: inline-block;
            background-color: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1e40af;
            font-weight: bold;
            font-size: 9pt;
            padding: 4px 14px;
            border-radius: 100px;
            margin-top: 8px;
            text-transform: uppercase;
          }
          .meta-info {
            font-size: 8.5pt;
            color: #1e3a8a;
            font-weight: bold;
            margin-top: 10px;
          }
          h3 {
            font-size: 11pt;
            font-weight: bold;
            color: #0f172a;
            background-color: #f8fafc;
            border-left: 4px solid #1e3a8a;
            padding: 6px 12px;
            margin-top: 25px;
            margin-bottom: 12px;
            text-transform: uppercase;
          }
          .table-responsive {
            width: 100%;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            margin-bottom: 20px;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
          }
          th {
            background-color: #1e3a8a;
            color: #ffffff;
            font-weight: bold;
            font-size: 8.5pt;
            padding: 8px 10px;
            text-align: left;
            text-transform: uppercase;
          }
          td {
            border-bottom: 1px solid #e2e8f0;
            font-size: 8.5pt;
            padding: 8px 10px;
            vertical-align: top;
          }
          tr:last-child td {
            border-bottom: none;
          }
          tr:nth-child(even) td {
            background-color: #f8fafc;
          }
          .text-center {
            text-align: center;
          }
          .font-bold {
            font-weight: bold;
          }
          .font-mono {
            font-family: Consolas, Monaco, "Courier New", Courier, monospace;
          }
          .profile-container {
            border: 1px solid #e2e8f0;
            border-left: 5px solid #2563eb;
            background-color: #f8fafc;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
          }
          .profile-grid {
            width: 100%;
            border-collapse: collapse;
          }
          .profile-grid td {
            border: none;
            padding: 6px 10px;
            background-color: transparent !important;
          }
          .label {
            color: #64748b;
            font-size: 8pt;
            text-transform: uppercase;
            font-weight: bold;
            margin-bottom: 2px;
          }
          .value {
            font-size: 9.5pt;
            font-weight: bold;
            color: #0f172a;
          }
          .remark-box {
            background-color: #fffbeb;
            border: 1px solid #fde68a;
            border-left: 5px solid #d97706;
            padding: 12px 15px;
            margin: 15px 0;
            border-radius: 6px;
          }
          .remark-title {
            font-size: 8pt;
            font-weight: bold;
            color: #b45309;
            text-transform: uppercase;
            margin-bottom: 4px;
          }
          .remark-text {
            font-weight: bold;
            font-size: 9.5pt;
            color: #78350f;
          }
          .badge-green {
            background-color: #dcfce7;
            color: #15803d;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
            display: inline-block;
          }
          .badge-red {
            background-color: #fee2e2;
            color: #b91c1c;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
            display: inline-block;
          }
          .page-break {
            border-top: 2px dashed #cbd5e1;
            margin: 30px 0;
            padding-top: 30px;
          }
          .signature-table {
            margin-top: 30px;
            border: none;
            width: 100%;
          }
          .signature-table td {
            border: none;
            text-align: center;
            padding: 20px 10px 10px 10px;
            background-color: transparent !important;
          }
          .stamp-box {
            background-color: #f8fafc;
            border: 1px dashed #cbd5e1;
            padding: 12px;
            border-radius: 6px;
            text-align: left;
          }
          @media print {
            body {
              background-color: #ffffff;
              padding: 0;
            }
            .container {
              box-shadow: none;
              padding: 0;
              border-radius: 0;
              max-width: 100%;
            }
            .page-break {
              page-break-before: always;
              border-top: none;
              margin: 0;
              padding-top: 0;
            }
          }
        </style>
      </head>
      <body>
        <div class="container">
    `;

    // 1. First Page: Group Summary Matrix (only if multiple)
    if (selectedPassports.length > 1) {
      htmlContent += `
        <div class="dossier-header">
          <div class="mya-header">Department of Immigration (Myeik)</div>
          <div class="eng-header">Second Division, Department of Immigration</div>
          <div>
            <span class="dossier-title-badge">Group Stay Verification & Intelligence Report</span>
          </div>
          <div class="meta-info">
            DOSSIER IDENTIFICATION NO: ${activeDossierID} <br/>
            GENERATED: ${new Date().toLocaleString()} • CONFIDENTIAL OFFICIAL RECORDS COPY
          </div>
        </div>

        <h3>Group Comparison Overview Matrix</h3>
        <div class="table-responsive">
          <table>
            <thead>
              <tr>
                <th style="width: 4%; text-align: center;">#</th>
                <th style="width: 20%;">Guest Name / Passport</th>
                <th style="width: 14%; text-align: center;">Stay Status</th>
                <th style="width: 12%;">Nationality</th>
                <th style="width: 22%;">Residence Stay Address</th>
                <th style="width: 16%;">Visa & Validity</th>
                <th style="width: 12%;">Local Sponsor</th>
                <th style="width: 10%; text-align: center;">Remaining</th>
              </tr>
            </thead>
            <tbody>
      `;

      selectedPassports.forEach((pp, idx) => {
        const m = computedMovementMap[pp] || {};
        const rem = getRemainingDaysNum(m.end);
        const statusInfo = getGuestStayStatus(pp, computedMovementMap);
        const statusBadgeHtml = statusInfo.isStillIn
          ? `<span class="badge-green">🟢 STILL IN (လက်ရှိ)</span>`
          : statusInfo.isDeparted
          ? `<span class="badge-red">🔴 DEPARTED (ထွက်သွား)</span>`
          : `<span style="background-color:#f1f5f9; color:#475569; padding:2px 6px; border-radius:4px; font-weight:bold;">⚪ REGISTERED</span>`;

        htmlContent += `
          <tr>
            <td class="text-center font-bold">${idx + 1}</td>
            <td>
              <div class="font-bold font-mono" style="color: #1e3a8a; font-size: 9pt;">${pp}</div>
              <div style="font-size: 8pt; font-weight: bold; color: #334155;">${m.n || '-'}</div>
            </td>
            <td class="text-center">${statusBadgeHtml}</td>
            <td class="font-bold">${m.nat || '-'}</td>
            <td>${m.loc || '-'}</td>
            <td>
              <div class="font-bold">${m.visa || '-'}</div>
              <div style="font-size: 7.5pt; color: #64748b;">${m.start} to ${m.end}</div>
            </td>
            <td>${m.agent || '-'}</td>
            <td class="text-center font-bold font-mono">
              <span class="${rem < 0 ? 'badge-red' : 'badge-green'}">${rem} Days</span>
            </td>
          </tr>
        `;
      });

      htmlContent += `
            </tbody>
          </table>
        </div>

        <!-- Signatures on Summary Page -->
        <table class="signature-table">
          <tr>
            <td style="width: 33%;">
              <div style="border-top: 1px solid #94a3b8; padding-top: 8px; font-weight: bold; font-size: 9pt;">Prepared By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 3px;">Immigration Investigator</div>
            </td>
            <td style="width: 33%;">
              <div style="border-top: 1px solid #94a3b8; padding-top: 8px; font-weight: bold; font-size: 9pt;">Verified By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 3px;">Duty Commander</div>
            </td>
            <td style="width: 34%;">
              <div class="stamp-box">
                <div style="font-weight: bold; font-size: 8pt; color: #1e3a8a; margin-bottom: 6px;">Certified Stamp / Verification</div>
                <div style="font-size: 7.5pt; color: #64748b; font-style: italic;">Stamp Box / Verification Stamp Here</div>
              </div>
            </td>
          </tr>
        </table>

        <div class="page-break"></div>
      `;
    }

    // 2. Individual Pages
    selectedPassports.forEach((pp, idx) => {
      const bio = computedMovementMap[pp];
      if (!bio) return;

      const personLogs = records
        .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
        .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

      const personCheckingHistory = checkingHistory
        .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
        .sort((a, b) => b.id - a.id);

      htmlContent += `
        <div class="dossier-header">
          <div class="mya-header">Department of Immigration (Myeik)</div>
          <div class="eng-header">Individual Stay Verification Dossier</div>
          <div>
            <span class="dossier-title-badge">CONFIDENTIAL INVESTIGATION RECORD</span>
          </div>
          <div class="meta-info">
            DOSSIER IDENTIFICATION NO: ${activeDossierID} <br/>
            Dossier Record #${idx + 1} of ${selectedPassports.length} • Generated on ${new Date().toLocaleString()}
          </div>
        </div>

        <div class="profile-container">
          <table class="profile-grid">
            <tr>
              <td style="width: 50%;">
                <div class="label">Passport Number</div>
                <div class="value" style="font-size: 10pt; color: #1e3a8a;">🛂 ${pp}</div>
              </td>
              <td style="width: 50%;">
                <div class="label">Full Name</div>
                <div class="value">👤 ${bio.n}</div>
              </td>
            </tr>
            <tr>
              <td>
                <div class="label">Nationality & Gender</div>
                <div class="value">🌐 ${bio.nat || '-'} (${bio.gender === 'M' ? 'MALE' : 'FEMALE'})</div>
              </td>
              <td>
                <div class="label">Visa Category & Validity</div>
                <div class="value">📅 ${bio.visa || '-'} (${bio.start} to ${bio.end})</div>
              </td>
            </tr>
            <tr>
              <td colspan="2" style="padding-top: 10px;">
                <div class="label">Registered Destination Address</div>
                <div class="value" style="font-weight: normal; color: #0f172a;">📍 ${bio.loc || '-'}</div>
              </td>
            </tr>
            <tr>
              <td colspan="2" style="padding-top: 10px;">
                <div class="label">Address Description Details</div>
                <div class="value" style="font-weight: normal; font-style: italic; color: #475569;">💬 ${bio.stayDescription || 'No description logged.'}</div>
              </td>
            </tr>
            <tr>
              <td style="padding-top: 10px;">
                <div class="label">Local Broker/Sponsor</div>
                <div class="value">🏢 ${bio.agent || '-'}</div>
              </td>
              <td style="padding-top: 10px;">
                <div class="label">Sponsor Contact</div>
                <div class="value font-mono">📞 ${bio.contact || '-'}</div>
              </td>
            </tr>
          </table>
        </div>

        <div class="remark-box">
          <div class="remark-title">📝 Dossier Remark Details</div>
          <div class="remark-text">
            ${currentRemarks[pp] || bio.remarks || 'No special remarks registered for this profile.'}
          </div>
        </div>

        <!-- Flight logs -->
        <h3>✈️ Flight Movement Logs</h3>
        <div class="table-responsive">
          <table>
            <thead>
              <tr>
                <th style="width: 5%; text-align: center;">#</th>
                <th style="width: 20%;">Timestamp</th>
                <th style="width: 10%; text-align: center;">Direction</th>
                <th style="width: 25%;">Visa & Stay Schedule</th>
                <th style="width: 25%;">Stay Address</th>
                <th style="width: 15%;">Registered Officer</th>
              </tr>
            </thead>
            <tbody>
      `;

      if (personLogs.length === 0) {
        htmlContent += `
          <tr>
            <td colspan="6" class="text-center" style="color: #94a3b8; font-style: italic; padding: 15px;">No flight movement logs found</td>
          </tr>
        `;
      } else {
        personLogs.slice(0, 10).forEach((log, lIdx) => {
          htmlContent += `
            <tr>
              <td class="text-center">${lIdx + 1}</td>
              <td class="font-mono" style="font-size: 8pt;">${log.timestamp}</td>
              <td class="text-center font-bold">
                <span class="${log.mode === 'IN' ? 'badge-green' : 'badge-red'}">${log.mode}</span>
              </td>
              <td>
                <div class="font-bold" style="font-size: 8pt;">${log.visaType}</div>
                ${log.stayFrom ? `<div style="font-size: 7.5pt; color: #64748b; margin-top: 2px;">${log.stayFrom} to ${log.stayTo}</div>` : ''}
              </td>
              <td style="font-size: 8pt;">${log.address || '-'}</td>
              <td style="font-size: 8pt; font-weight: bold; color: #475569;">${log.officialTitle || ''} ${log.officialName || 'System'}</td>
            </tr>
          `;
        });
      }

      htmlContent += `
            </tbody>
          </table>
        </div>

        <!-- Verification annals -->
        ${personCheckingHistory.length > 0 ? `
          <h3>🔍 Physical Verification Annals</h3>
          <div class="table-responsive">
            <table>
              <thead>
                <tr>
                  <th style="width: 5%; text-align: center;">#</th>
                  <th style="width: 20%;">Check Date</th>
                  <th style="width: 22%;">Verified Status</th>
                  <th style="width: 35%;">Confirmed Address</th>
                  <th style="width: 18%;">Verifier</th>
                </tr>
              </thead>
              <tbody>
        ` : ''}
      `;

      if (personCheckingHistory.length > 0) {
        personCheckingHistory.forEach((ch, cIdx) => {
          const isNotMatch = ch.status?.includes('NOT');
          htmlContent += `
            <tr>
              <td class="text-center">${cIdx + 1}</td>
              <td class="font-mono" style="font-size: 8pt;">${ch.checkDate}</td>
              <td class="font-bold">
                <span class="${isNotMatch ? 'badge-red' : 'badge-green'}">${ch.status}</span>
              </td>
              <td style="font-size: 8pt;">${ch.confirmedAddress || '-'}</td>
              <td style="font-size: 8pt; font-weight: bold; color: #475569;">${ch.officerTitle || ''} ${ch.officerName || 'Officer'}</td>
            </tr>
          `;
        });

        htmlContent += `
              </tbody>
            </table>
          </div>
        `;
      }

      htmlContent += `
        <!-- Signatures on Individual Page -->
        <table class="signature-table">
          <tr>
            <td style="width: 50%;">
              <div style="border-top: 1px solid #cbd5e1; padding-top: 8px; font-weight: bold; font-size: 8.5pt;">Prepared & Authenticated By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 2px;">Investigation Analyst Officer</div>
            </td>
            <td style="width: 50%;">
              <div style="border-top: 1px solid #cbd5e1; padding-top: 8px; font-weight: bold; font-size: 8.5pt;">Reviewed & Certified By</div>
              <div style="font-size: 8pt; color: #64748b; margin-top: 2px;">Duty Supervisor / Commander</div>
            </td>
          </tr>
        </table>
      `;

      // Page break between guests (except last)
      if (idx < selectedPassports.length - 1) {
        htmlContent += `<div class="page-break"></div>`;
      }
    });

    htmlContent += `
        </div>
      </body>
      </html>
    `;

    // Download file as native HTML web page (100% responsive and readable on any mobile device)
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `IMM_Dossier_${activeDossierID}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("WEB DOSSIER (HTML) EXPORTED FOR MOBILE SHARING!");
  };

  // Export Consolidated Group Dossiers to A4 PDF using jsPDF + html2canvas page-by-page
  const handleSaveAsPDF = async () => {
    if (selectedPassports.length === 0) {
      showToast("NO PASSPORTS SELECTED FOR EXPORT");
      return;
    }

    saveActiveDossierToHistory();
    setIsExportingPDF(true);
    showToast("PREPARING HIGH-FIDELITY A4 PDF...");

    // Store original styles to restore in finally block
    const styleTags = Array.from(document.querySelectorAll('style'));
    const originalStyles = styleTags.map(tag => ({ tag, html: tag.innerHTML }));

    // Temporarily replace oklch(...) colors inside styles to prevent html2canvas crashes!
    styleTags.forEach(tag => {
      if (tag.innerHTML.includes('oklch')) {
        tag.innerHTML = tag.innerHTML.replace(/oklch\([^)]+\)/g, 'rgb(75, 85, 99)');
      }
    });

    try {
      // Small timeout to allow styling variables to settle in DOM
      await new Promise(resolve => setTimeout(resolve, 250));

      const pdfPages = document.querySelectorAll('#pdf-capture-container .pdf-page');
      if (pdfPages.length === 0) {
        showToast("PDF PAGES CONTAINER NOT FOUND");
        setIsExportingPDF(false);
        return;
      }

      const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4',
        compress: true
      });

      for (let i = 0; i < pdfPages.length; i++) {
        const pageEl = pdfPages[i] as HTMLElement;

        // Convert the styled A4 page div to high-fidelity crisp canvas (without any oklch parsing errors!)
        const canvas = await html2canvas(pageEl, {
          scale: 2, // High resolution crispness
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff'
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        
        // Add page to PDF (except first page which is created by default)
        if (i > 0) {
          doc.addPage('a4', 'portrait');
        }

        // Standard A4 dimensions: 210mm x 297mm
        doc.addImage(imgData, 'JPEG', 0, 0, 210, 297, undefined, 'FAST');
      }

      doc.save(`IMM_Dossier_${activeDossierID}.pdf`);
      showToast("HIGH-FIDELITY PDF DOWNLOADED SUCCESSFULLY!");
    } catch (e) {
      console.error(e);
      showToast("PDF GENERATION FAILURE");
    } finally {
      // Always restore the exact original styles back to the document
      originalStyles.forEach(({ tag, html }) => {
        tag.innerHTML = html;
      });
      setIsExportingPDF(false);
    }
  };

  return (
    <div className="max-w-[1600px] mx-auto py-8 px-4 space-y-8">
      {/* Global Datalist for Selectable Stay Addresses */}
      <datalist id="invStayAddressDatalist">
        {stayAddressOptions.map((addr, idx) => (
          <option key={addr + '_' + idx} value={addr} />
        ))}
      </datalist>

      {/* Viewer Notice Banner */}
      {isViewer && (
        <div className="p-4 bg-amber-50 border-2 border-amber-300 rounded-2xl text-amber-900 text-xs font-bold flex items-center gap-3 no-print shadow-sm">
          <div className="p-2 bg-amber-200/70 rounded-xl text-amber-800">
            <ShieldAlert size={22} />
          </div>
          <div>
            <div className="font-black text-sm uppercase text-amber-950 flex items-center gap-2">
              <span>🔒 Viewer Level (Read-Only Mode) သတ်မှတ်ထားပါသည်</span>
              <span className="text-[10px] bg-amber-200 text-amber-900 px-2 py-0.5 rounded-full font-mono">SEARCH ONLY</span>
            </div>
            <div className="text-amber-800 mt-0.5">စုံစမ်းစစ်ဆေးမှု ဖိုင်တွဲများနှင့် နိုင်ငံကူးလက်မှတ် မှတ်တမ်းများအား ရှာဖွေစစ်ဆေးခြင်း (Search / Dossier View) သာ ဆောင်ရွက်နိုင်ပြီး၊ Stay Details ပြင်ဆင်ခြင်း၊ Quick-Edit ပြင်ဆင်ခြင်းနှင့် မှတ်ချက်အသစ် ရေးသားခြင်းများ ပိတ်ထားပါသည်</div>
          </div>
        </div>
      )}

      {/* INV Navigation Sub-Tabs */}
      <div className="flex items-center gap-2 p-1.5 bg-slate-100/90 rounded-2xl border border-slate-200 no-print shadow-xs">
        <button
          onClick={() => setInvSubTab('STAY_CHECK')}
          className={`flex-1 sm:flex-none px-5 py-2.5 rounded-xl text-xs sm:text-sm font-black transition-all flex items-center justify-center gap-2 cursor-pointer ${
            invSubTab === 'STAY_CHECK'
              ? 'bg-indigo-900 text-white shadow-md'
              : 'bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/70'
          }`}
        >
          <Building2 size={16} />
          <span>🏢 Stay Check (နေရာအလိုက် နေထိုင်မှုစစ်ဆေးခြင်း)</span>
          <span className="text-[10px] bg-emerald-500 text-white px-2 py-0.5 rounded-full font-bold">
            NEW
          </span>
        </button>

        <button
          onClick={() => setInvSubTab('DOSSIER')}
          className={`flex-1 sm:flex-none px-5 py-2.5 rounded-xl text-xs sm:text-sm font-black transition-all flex items-center justify-center gap-2 cursor-pointer ${
            invSubTab === 'DOSSIER'
              ? 'bg-indigo-900 text-white shadow-md'
              : 'bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/70'
          }`}
        >
          <Search size={16} />
          <span>🔍 Individual Dossier & Profiler (လူပုဂ္ဂိုလ်အလိုက် စုံစမ်းစစ်ဆေးမှု)</span>
        </button>
      </div>

      {invSubTab === 'STAY_CHECK' ? (
        <StayCheckConsole
          records={records}
          setRecords={setRecords}
          tempRecords={tempRecords}
          setTempRecords={setTempRecords}
          checkingHistory={checkingHistory}
          setCheckingHistory={setCheckingHistory}
          masterData={masterData}
          movementMap={movementMap}
          passportToLatestInfo={passportToLatestInfo}
          currentUser={currentUser}
          cloudAuthUser={cloudAuthUser}
          isViewer={isViewer}
          showToast={showToast}
          setActivePrintPreview={setActivePrintPreview}
        />
      ) : (
        <>
          {/* 1. Header & Quick Group Search Card */}
          <div className="card shadow-xl border-t-8 border-indigo-900 bg-white p-4 sm:p-6 no-print">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 sm:gap-6">
          <div className="flex items-center gap-3 sm:gap-4">
            <div className="bg-indigo-900 p-2.5 sm:p-3.5 rounded-2xl text-white shadow-lg shadow-indigo-100 shrink-0">
              <Users size={26} className="sm:w-8 sm:h-8" />
            </div>
            <div>
              <h2 className="text-lg sm:text-2xl font-black text-gray-800 uppercase tracking-tight flex items-center gap-2 flex-wrap">
                Investigation Center <span className="text-[10px] sm:text-xs bg-indigo-100 text-indigo-900 px-2 sm:px-2.5 py-0.5 sm:py-1 rounded-full font-black">စုံစမ်းစစ်ဆေးရေးနှင့် စိစစ်မှုဗဟို</span>
              </h2>
              <p className="text-gray-400 text-[10px] sm:text-xs font-bold font-mono uppercase mt-0.5 sm:mt-1">Multi-person suspect check, visa validations, and stay history logs audit</p>
            </div>
          </div>

          <div className="relative w-full lg:w-[480px]">
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Passport, Name, Nationality သို့မဟုတ် Stay Address ဖြင့် ရှာရန်..."
              className="input-field pl-10 sm:pl-11 pr-4 py-2.5 sm:py-3.5 text-xs sm:text-sm bg-slate-50 border-2 border-slate-200 rounded-xl sm:rounded-2xl focus:bg-white focus:border-indigo-600 transition-all font-bold placeholder-gray-400 text-slate-800 w-full"
            />
            <Search size={18} className="absolute left-3.5 top-3 sm:top-4 text-indigo-900" />
          </div>
        </div>

        {/* Suggestion Dropdown */}
        {searchQuery.trim() !== '' && (
          <div className="mt-4 border border-slate-100 rounded-2xl overflow-hidden shadow-2xl bg-white max-h-72 overflow-y-auto divide-y divide-slate-50 text-xs">
            {filteredPeople.length === 0 ? (
              <div className="p-4 text-center text-slate-400 font-bold uppercase tracking-wider">
                No matching stayed people or addresses found
              </div>
            ) : (
              filteredPeople.map((person) => {
                const statusInfo = getGuestStayStatus(person.passport, computedMovementMap);
                return (
                  <div
                    key={person.passport}
                    onClick={() => addPassportToGroup(person.passport)}
                    className="p-3.5 flex justify-between items-center hover:bg-indigo-50/50 cursor-pointer transition-colors border-b last:border-0"
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-3 flex-wrap">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-extrabold text-indigo-900 text-sm font-mono tracking-tight bg-indigo-50 px-2 py-1 rounded-lg border border-indigo-100">{person.passport}</span>
                        <span className="font-black text-slate-800 uppercase text-xs">{person.fullname}</span>
                        
                        {/* Status Badge */}
                        {statusInfo.isStillIn ? (
                          <span className="text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-1 shadow-sm">
                            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                            🟢 {statusInfo.label} (STILL IN)
                          </span>
                        ) : statusInfo.isDeparted ? (
                          <span className="text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-200 flex items-center gap-1">
                            <span className="w-2 h-2 rounded-full bg-rose-500"></span>
                            🔴 {statusInfo.label} (DEPARTED)
                          </span>
                        ) : (
                          <span className="text-[10px] font-black uppercase px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200">
                            ⚪ {statusInfo.label}
                          </span>
                        )}
                      </div>
                      {person.address && (
                        <span className="text-[10px] text-slate-500 font-bold flex items-center gap-1 truncate max-w-[220px]" title={person.address}>
                          <MapPin size={10} className="text-indigo-500 shrink-0" /> {person.address}
                        </span>
                      )}
                      {statusInfo.timeText && (
                        <span className="text-[9.5px] font-mono font-bold text-indigo-700 bg-indigo-50/80 px-2 py-0.5 rounded border border-indigo-100">
                          {statusInfo.timeText}
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <span className="text-[10px] uppercase font-black px-2.5 py-1 rounded-full bg-slate-100 text-slate-600">
                        {person.nationality} ({person.gender || '-'})
                      </span>
                      <span className="text-[10px] font-black text-white bg-indigo-900 hover:bg-black px-3 py-1 rounded-lg flex items-center gap-1 transition-colors">
                        <Plus size={11} strokeWidth={3} /> ADD
                      </span>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        )}
      </div>

      {/* 2. CASE A: NO FOREIGNER SEARCHED YET -> SHOW INTELLIGENCE WARNINGS DASHBOARD */}
      {selectedPassports.length === 0 ? (
        <div className="space-y-6 sm:space-y-8 animate-in fade-in duration-300">
          
          {/* Warn Summary Stats Row */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-6 no-print">
            
            <div 
              onClick={() => setActiveAlertTab('OVERSTAY')}
              className={`p-4 sm:p-6 rounded-2xl sm:rounded-3xl border-2 cursor-pointer transition-all flex items-center justify-between shadow-sm hover:shadow-md ${
                activeAlertTab === 'OVERSTAY' 
                  ? 'bg-rose-50 border-rose-200 text-rose-900 ring-2 sm:ring-4 ring-rose-100' 
                  : 'bg-white border-slate-100 text-slate-700 hover:border-slate-200'
              }`}
            >
              <div className="flex items-center gap-3 sm:gap-4">
                <div className={`p-3 sm:p-4 rounded-xl sm:rounded-2xl ${activeAlertTab === 'OVERSTAY' ? 'bg-rose-600 text-white' : 'bg-slate-100 text-slate-500'}`}>
                  <ShieldAlert size={24} className="sm:w-7 sm:h-7" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wide">နေထိုင်ခွင့်ကုန်ဆုံး သက်တမ်းလွန်နေသူများ</h3>
                  <p className="text-[9px] sm:text-[10px] font-bold opacity-75 uppercase font-mono mt-0.5">Overstayed Permits Warnings</p>
                </div>
              </div>
              <div className="text-2xl sm:text-3xl font-black ml-2">{overstayGuests.length}</div>
            </div>

            <div 
              onClick={() => setActiveAlertTab('LONG_STAY')}
              className={`p-4 sm:p-6 rounded-2xl sm:rounded-3xl border-2 cursor-pointer transition-all flex items-center justify-between shadow-sm hover:shadow-md ${
                activeAlertTab === 'LONG_STAY' 
                  ? 'bg-amber-50 border-amber-200 text-amber-900 ring-2 sm:ring-4 ring-amber-100' 
                  : 'bg-white border-slate-100 text-slate-700 hover:border-slate-200'
              }`}
            >
              <div className="flex items-center gap-3 sm:gap-4">
                <div className={`p-3 sm:p-4 rounded-xl sm:rounded-2xl ${activeAlertTab === 'LONG_STAY' ? 'bg-amber-500 text-white' : 'bg-slate-100 text-slate-500'}`}>
                  <AlertTriangle size={24} className="sm:w-7 sm:h-7" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wide">{longStayThresholdDays} ရက်နှင့်အထက် တည်းခိုနေထိုင်သူများ</h3>
                  <p className="text-[9px] sm:text-[10px] font-bold opacity-75 uppercase font-mono mt-0.5">Long-Term Stayers ({'>='} {longStayThresholdDays} Days)</p>
                </div>
              </div>
              <div className="text-2xl sm:text-3xl font-black ml-2">{longStayGuests.length}</div>
            </div>

          </div>

          {/* Warning Table Card */}
          <div className="card shadow-xl border border-slate-100 bg-white p-4 sm:p-6 no-print">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b pb-4 sm:pb-5 mb-4 sm:mb-6">
              <div className="flex-1 w-full">
                <div className="flex flex-col lg:flex-row lg:items-center gap-3 sm:gap-4">
                  <div>
                    <h3 className="text-base sm:text-lg font-black text-slate-800 uppercase tracking-tight flex items-center gap-2 flex-wrap">
                      <AlertCircle className={activeAlertTab === 'OVERSTAY' ? 'text-rose-600' : 'text-amber-500'} size={20} />
                      {activeAlertTab === 'OVERSTAY' 
                        ? 'နေထိုင်ခွင့်ကုန်ဆုံး သက်တမ်းလွန်နေသော နိုင်ငံခြားသားများ ဇယား' 
                        : `${longStayThresholdDays} ရက်နှင့်အထက် တစ်နေရာတည်းတွင် ကြာမြင့်စွာ တည်းခိုနေထိုင်သူများ ဇယား`}
                    </h3>
                    <p className="text-gray-400 text-[10px] sm:text-xs font-bold uppercase font-mono mt-0.5">
                      Showing {activeAlertTab === 'OVERSTAY' ? overstayGuests.length : longStayGuests.length} active flagged guest profiles
                    </p>
                  </div>

                  {activeAlertTab === 'LONG_STAY' && (
                    <div className="flex items-center gap-2 bg-amber-50 border border-amber-200 px-3 py-1.5 rounded-2xl lg:ml-6 shrink-0 w-full sm:w-max">
                      <span className="text-[10px] sm:text-[11px] font-black uppercase text-amber-900">သတ်မှတ်ရက် (Days Threshold):</span>
                      <input 
                        type="number"
                        value={longStayThresholdDays}
                        onChange={(e) => {
                          const val = parseInt(e.target.value);
                          setLongStayThresholdDays(isNaN(val) ? 0 : val);
                        }}
                        className="w-16 sm:w-20 px-2 py-1 text-center font-bold font-mono border-2 border-amber-200 bg-white rounded-xl focus:border-amber-500 text-amber-950 focus:outline-none text-xs"
                        min={1}
                      />
                      <span className="text-[10px] font-extrabold text-amber-800 uppercase">ရက်</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="flex gap-2 w-full md:w-auto">
                <button
                  onClick={handleExportWarningsExcel}
                  className="btn bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black uppercase tracking-wider py-2.5 px-4 rounded-xl flex items-center gap-1.5 shadow-md w-full md:w-auto justify-center"
                >
                  <FileSpreadsheet size={15} /> Excel Sheet Output
                </button>
              </div>
            </div>

            {/* Warn Table Render */}
            {(() => {
              const data = activeAlertTab === 'OVERSTAY' ? overstayGuests : longStayGuests;
              if (data.length === 0) {
                return (
                  <div className="py-12 sm:py-20 text-center border-2 border-dashed border-slate-100 rounded-2xl bg-slate-50/50 space-y-3 p-4">
                    <CheckCircle2 className="mx-auto text-emerald-500" size={40} />
                    <h4 className="text-sm sm:text-base font-black text-slate-700 uppercase tracking-tight">No Active Risk Flags Detected</h4>
                    <p className="text-xs text-slate-400 font-bold max-w-md mx-auto leading-normal">
                      All currently registered checked-in foreign guests comply safely with their stay durations and no overstay or long-term risk parameters are active.
                    </p>
                  </div>
                );
              }

              return (
                <>
                  {/* Mobile Cards View */}
                  <div className="block md:hidden space-y-3">
                    {data.map((m: any, idx) => {
                      const rem = getRemainingDaysNum(m.end);
                      const elapsed = getDaysStayedSoFar(m.start);
                      return (
                        <div key={m.p} className="p-4 bg-slate-50/80 rounded-2xl border border-slate-200/80 space-y-3 text-xs">
                          <div className="flex justify-between items-start">
                            <div>
                              <span className="text-[9px] font-bold text-slate-400 font-mono">#{idx + 1}</span>
                              <div className="font-black text-indigo-950 text-sm font-mono tracking-tight">{m.p}</div>
                              <div className="font-bold text-slate-800 uppercase text-xs mt-0.5">{m.n}</div>
                            </div>
                            <div className="text-right">
                              {rem < 0 ? (
                                <span className="text-[9px] font-black uppercase text-rose-700 bg-rose-50 px-2 py-1 border border-rose-200 rounded-lg inline-block">
                                  {rem} DAYS OVERSTAY
                                </span>
                              ) : (
                                <span className="text-[9px] font-black uppercase text-slate-700 bg-slate-100 px-2 py-0.5 rounded inline-block">
                                  {rem} Days Left
                                </span>
                              )}
                            </div>
                          </div>

                          <div className="grid grid-cols-2 gap-2 text-[11px] pt-1 border-t border-slate-200/60">
                            <div>
                              <span className="text-[9px] font-bold text-slate-400 block uppercase">Nationality</span>
                              <span className="font-bold text-slate-700 uppercase">{m.nat} ({m.gender === 'M' ? 'M' : m.gender === 'F' ? 'F' : '-'})</span>
                            </div>
                            <div>
                              <span className="text-[9px] font-bold text-slate-400 block uppercase">Visa & Dates</span>
                              <span className="font-bold text-slate-700 uppercase">{m.visa}</span>
                              <div className="text-[9px] font-mono text-indigo-700">{m.start} ➔ {m.end}</div>
                            </div>
                            <div className="col-span-2">
                              <span className="text-[9px] font-bold text-slate-400 block uppercase">Stay Address</span>
                              <span className="font-semibold text-slate-800">{m.loc}</span>
                              {m.stayDescription && (
                                <div className="text-[9px] text-slate-500 italic mt-0.5">🏡 {m.stayDescription}</div>
                              )}
                            </div>
                            <div>
                              <span className="text-[9px] font-bold text-slate-400 block uppercase">Days Stayed</span>
                              <span className="font-bold font-mono text-slate-700">{elapsed} Days ({Math.round(elapsed/30)} Mos)</span>
                            </div>
                          </div>

                          <div className="flex items-center gap-2 pt-2 border-t border-slate-200/60">
                            <button
                              onClick={() => addPassportToGroup(m.p)}
                              className="btn bg-indigo-50 hover:bg-indigo-900 hover:text-white text-indigo-900 py-2 px-3 rounded-xl border border-indigo-100 text-[10px] font-black uppercase flex items-center justify-center gap-1 flex-1"
                            >
                              <Plus size={12} strokeWidth={3} /> Add check
                            </button>
                            <button
                              onClick={() => openQuickEdit(m.p)}
                              className="btn bg-slate-100 hover:bg-slate-200 text-slate-700 py-2 px-3 rounded-xl border border-slate-200 text-[10px] font-black uppercase flex items-center justify-center gap-1 flex-1"
                            >
                              <Edit size={12} /> Edit Details
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {/* Desktop Table View */}
                  <div className="hidden md:block overflow-x-auto border border-slate-100 rounded-2xl overflow-hidden text-xs shadow-inner">
                    <table className="w-full text-left">
                      <thead className="bg-slate-50 border-b font-black text-slate-500 uppercase text-[9px] tracking-wide">
                        <tr>
                          <th className="p-3.5 text-center">Sr. No.</th>
                          <th className="p-3.5">Passport / Name</th>
                          <th className="p-3.5">Nationality / Gender</th>
                          <th className="p-3.5">Stay Address Location</th>
                          <th className="p-3.5">Visa / Duration</th>
                          <th className="p-3.5 text-center">Remaining Days</th>
                          <th className="p-3.5 text-center">Days Stayed</th>
                          <th className="p-3.5 text-center">Actions Console</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-bold text-slate-700 bg-white">
                        {data.map((m: any, idx) => {
                          const rem = getRemainingDaysNum(m.end);
                          const elapsed = getDaysStayedSoFar(m.start);
                          return (
                            <tr key={m.p} className="hover:bg-slate-50/50 transition-colors">
                              <td className="p-3.5 text-center font-mono font-bold text-slate-400">{idx + 1}</td>
                              <td className="p-3.5">
                                <div className="font-black text-indigo-950 font-mono tracking-tight">{m.p}</div>
                                <div className="text-[11px] text-slate-600 font-bold uppercase mt-0.5">{m.n}</div>
                              </td>
                              <td className="p-3.5">
                                <div className="uppercase text-slate-800">{m.nat}</div>
                                <div className="text-[10px] text-slate-400 mt-0.5">{m.gender === 'M' ? 'MALE' : m.gender === 'F' ? 'FEMALE' : '-'}</div>
                              </td>
                              <td className="p-3.5 max-w-[220px]">
                                <div className="truncate text-slate-800 leading-tight" title={m.loc}>{m.loc}</div>
                                {m.stayDescription && (
                                  <div className="text-[9px] text-slate-400 italic mt-0.5 truncate max-w-[200px]">🏡 {m.stayDescription}</div>
                                )}
                              </td>
                              <td className="p-3.5">
                                <div className="uppercase text-[11px]">{m.visa}</div>
                                <div className="text-[10px] font-mono text-indigo-700 mt-0.5">{m.start} ➔ {m.end}</div>
                              </td>
                              <td className="p-3.5 text-center font-mono">
                                {rem < 0 ? (
                                  <span className="text-[10px] font-black uppercase text-rose-700 bg-rose-50 px-2.5 py-1 border border-rose-200 rounded-lg block w-max mx-auto">
                                    {rem} DAYS OVERSTAY
                                  </span>
                                ) : (
                                  <span className="text-[10px] font-black uppercase text-slate-700 bg-slate-100 px-2 py-0.5 rounded block w-max mx-auto">
                                    {rem} Days
                                  </span>
                                )}
                              </td>
                              <td className="p-3.5 text-center font-mono">
                                {elapsed >= longStayThresholdDays ? (
                                  <span className="text-[10px] font-black uppercase text-amber-700 bg-amber-50 px-2.5 py-1 border border-amber-200 rounded-lg block w-max mx-auto">
                                    {elapsed} DAYS ({Math.round(elapsed/30)} Mos)
                                  </span>
                                ) : (
                                  <span className="text-[10px] font-bold text-slate-500 block">
                                    {elapsed} Days
                                  </span>
                                )}
                              </td>
                              <td className="p-3.5">
                                <div className="flex items-center justify-center gap-1.5">
                                  <button
                                    onClick={() => addPassportToGroup(m.p)}
                                    className="btn bg-indigo-50 hover:bg-indigo-900 hover:text-white text-indigo-900 p-2 rounded-xl border border-indigo-100 transition-all text-[10px] font-black uppercase flex items-center gap-1 shrink-0"
                                    title="Add to current group check list"
                                  >
                                    <Plus size={12} strokeWidth={3} /> Add check
                                  </button>
                                  <button
                                    onClick={() => openQuickEdit(m.p)}
                                    className="btn bg-slate-50 hover:bg-slate-200 text-slate-700 p-2 rounded-xl border border-slate-200 transition-all text-[10px] font-black uppercase flex items-center gap-1 shrink-0"
                                    title="Immediate edit personal details & stay validity"
                                  >
                                    <Edit size={12} /> Edit Details
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      ) : (
        /* 3. CASE B: ACTIVE PASSPORTS ARE SELECTED -> SHOW MULTI-PERSON DOSSIER INVESTIGATION GROUP BOARD */
        <div className="space-y-6 sm:space-y-8 animate-in fade-in duration-300">
          
          {/* Active Group List Banner */}
          <div className="bg-slate-900 text-white p-4 sm:p-5 rounded-2xl sm:rounded-3xl shadow-xl flex flex-col md:flex-row justify-between items-start md:items-center gap-4 sm:gap-6 no-print">
            <div className="flex items-center gap-2 sm:gap-3 flex-wrap">
              <span className="text-[10px] sm:text-xs font-black uppercase tracking-widest text-indigo-400">Current Investigation Group ({selectedPassports.length}):</span>
              <div className="flex flex-wrap gap-1.5 sm:gap-2">
                {selectedPassports.map(pp => {
                  const guestInfo = computedMovementMap[pp] || { n: pp };
                  const statusInfo = getGuestStayStatus(pp, computedMovementMap);
                  return (
                    <div key={pp} className="bg-slate-800 text-slate-100 pl-2.5 sm:pl-3 pr-1.5 sm:pr-2 py-1 sm:py-1.5 rounded-xl border border-slate-700/60 flex items-center gap-1.5 sm:gap-2 text-[10px] sm:text-xs font-bold font-mono">
                      <span>{pp} ({guestInfo.n})</span>
                      {statusInfo.isStillIn ? (
                        <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded-md bg-emerald-600 text-white flex items-center gap-1">
                          🟢 လက်ရှိ
                        </span>
                      ) : statusInfo.isDeparted ? (
                        <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded-md bg-rose-600 text-white flex items-center gap-1">
                          🔴 ထွက်သွား
                        </span>
                      ) : (
                        <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded-md bg-slate-700 text-slate-300">
                          ⚪ မှတ်ပုံတင်
                        </span>
                      )}
                      <button 
                        onClick={() => removePassportFromGroup(pp)}
                        className="text-rose-400 hover:text-rose-600 hover:bg-slate-700 p-0.5 rounded-md transition-colors ml-0.5"
                      >
                        <X size={12} strokeWidth={2.5} />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>

            <button 
              onClick={() => {
                setSelectedPassports([]);
                setCurrentRemarks({});
                showToast("CLEARED INVESTIGATION GROUP LIST");
              }}
              className="btn bg-slate-800 text-rose-300 hover:bg-rose-950 hover:text-white text-[10px] sm:text-xs font-extrabold uppercase py-2 px-3 sm:px-4 rounded-xl border border-slate-700/60 flex items-center gap-1 shrink-0 w-full sm:w-auto justify-center"
            >
              <Trash2 size={13} /> Clear Group List
            </button>
          </div>

          {/* Group Comparison Matrix Board */}
          <div className="card shadow-xl border border-slate-100 bg-white p-4 sm:p-6 no-print">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b pb-4 sm:pb-5 mb-4 sm:mb-6">
              <div>
                <h3 className="text-base sm:text-lg font-black text-slate-800 uppercase tracking-tight flex items-center gap-2 flex-wrap">
                  <FileSpreadsheet className="text-indigo-900" size={20} />
                  အုပ်စုလိုက် အချက်အလက် ခြုံငုံနှိုင်းယှဥ်စိစစ်ရေးဇယား (Group Overview Comparison Matrix)
                </h3>
                <p className="text-gray-400 text-[10px] sm:text-xs font-bold uppercase font-mono mt-0.5">
                  Consolidated look of all {selectedPassports.length} selected investigated profiles
                </p>
              </div>

              <div className="grid grid-cols-2 sm:flex sm:flex-wrap gap-2 sm:gap-2.5 w-full md:w-auto">
                <button
                  onClick={() => setShowPrintPreviewModal(true)}
                  className="btn bg-indigo-950 hover:bg-black text-white text-[10px] sm:text-xs font-black uppercase tracking-wider py-2.5 sm:py-3 px-3 sm:px-4 rounded-xl flex items-center gap-1.5 shadow-lg justify-center cursor-pointer transition-all"
                >
                  <Eye size={14} /> Preview Print
                </button>
                <button
                  onClick={handlePrintGroupDossier}
                  className="btn bg-indigo-900 hover:bg-black text-white text-[10px] sm:text-xs font-black uppercase tracking-wider py-2.5 sm:py-3 px-3 sm:px-5 rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-lg shadow-indigo-100 justify-center cursor-pointer transition-all"
                >
                  <Printer size={14} /> Print Group
                </button>
                <button
                  onClick={handleExportGroupWordDoc}
                  className="btn bg-blue-700 hover:bg-blue-800 text-white text-[10px] sm:text-xs font-black uppercase tracking-wider py-2.5 sm:py-3 px-3 sm:px-5 rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-lg justify-center cursor-pointer transition-all"
                >
                  <FileText size={14} /> Word (DOC)
                </button>
                <button
                  onClick={handleExportGroupHTML}
                  className="btn bg-purple-700 hover:bg-purple-800 text-white text-[10px] sm:text-xs font-black uppercase tracking-wider py-2.5 sm:py-3 px-3 sm:px-5 rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-lg justify-center cursor-pointer transition-all"
                >
                  <FileText size={14} /> Web (HTML)
                </button>
                <button
                  onClick={handleExportGroupExcel}
                  className="btn bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] sm:text-xs font-black uppercase tracking-wider py-2.5 sm:py-3 px-3 sm:px-5 rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-md justify-center cursor-pointer transition-all"
                >
                  <Download size={14} /> Export Excel
                </button>
              </div>
            </div>

            {/* Mobile Cards for Group Overview */}
            <div className="block md:hidden space-y-3">
              {selectedPassports.map((pp, index) => {
                const m = computedMovementMap[pp] || { p: pp, n: '-', nat: '-', gender: '-', loc: '-', visa: '-', start: '-', end: '-', allowed: '-', agent: '', contact: '', stillPermittedStatus: '' };
                const rem = getRemainingDaysNum(m.end);
                const statusInfo = getGuestStayStatus(pp, computedMovementMap);
                return (
                  <div key={pp} className="p-4 bg-slate-50/80 rounded-2xl border border-slate-200/80 space-y-2.5 text-xs">
                    <div className="flex justify-between items-start">
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 font-mono">#{index + 1}</span>
                        <div className="font-black text-indigo-900 text-sm font-mono tracking-tight">{pp}</div>
                        <div className="font-bold text-slate-800 uppercase text-xs mt-0.5">{m.n}</div>
                      </div>
                      <div className="text-right flex flex-col items-end gap-1">
                        {rem < 0 ? (
                          <span className="text-[9px] font-black uppercase text-rose-700 bg-rose-50 px-2 py-0.5 rounded-lg border border-rose-200 inline-block">
                            {rem} Days
                          </span>
                        ) : (
                          <span className="text-[9px] font-black text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200 inline-block">
                            {rem} Days
                          </span>
                        )}
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-2 text-[11px] pt-1 border-t border-slate-200/60">
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Stay Status (အခြေအနေ)</span>
                        {statusInfo.isStillIn ? (
                          <span className="text-[9.5px] font-black uppercase text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-lg border border-emerald-200 inline-block">
                            🟢 လက်ရှိနေထိုင်သူ
                          </span>
                        ) : statusInfo.isDeparted ? (
                          <span className="text-[9.5px] font-black uppercase text-rose-800 bg-rose-50 px-2 py-0.5 rounded-lg border border-rose-200 inline-block">
                            🔴 ထွက်သွားပြီးသူ
                          </span>
                        ) : (
                          <span className="text-[9.5px] font-black uppercase text-slate-600 bg-slate-100 px-2 py-0.5 rounded border inline-block">
                            ⚪ မှတ်ပုံတင်ထား
                          </span>
                        )}
                      </div>
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Nationality</span>
                        <span className="font-bold text-slate-700 uppercase">{m.nat} ({m.gender || '-'})</span>
                      </div>
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Visa & Validity</span>
                        <span className="font-bold text-slate-700 uppercase">{m.visa}</span>
                        <div className="text-[9px] text-gray-400 font-mono">{m.start} to {m.end}</div>
                      </div>
                      <div className="col-span-2">
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Registered Address</span>
                        <span className="font-semibold text-slate-800">{m.loc}</span>
                      </div>
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Sponsor Agent</span>
                        <span className="font-bold text-indigo-950 uppercase">{m.agent || '-'}</span>
                      </div>
                      <div>
                        <span className="text-[9px] font-bold text-slate-400 block uppercase">Permit Status</span>
                        {normalizePermitStatus(m.stillPermittedStatus) === 'နေထိုင်ခွင့်ကျထားသောသူ' ? (
                          <span className="text-[9px] font-black uppercase text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-lg border border-emerald-200 inline-block">
                            PERMITTED
                          </span>
                        ) : normalizePermitStatus(m.stillPermittedStatus) === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? (
                          <span className="text-[9px] font-black uppercase text-rose-800 bg-rose-50 px-2 py-0.5 rounded-lg border border-rose-200 inline-block">
                            NOT PERMITTED
                          </span>
                        ) : (
                          <span className="text-[9px] font-black uppercase text-slate-600 bg-slate-100 px-2 py-0.5 rounded border inline-block">
                            CONFIRMED
                          </span>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Desktop Table View */}
            <div className="hidden md:block overflow-x-auto border border-slate-100 rounded-2xl overflow-hidden text-xs">
              <table className="w-full text-left">
                <thead className="bg-slate-50 border-b font-black text-slate-500 uppercase text-[9px] tracking-wide">
                  <tr>
                    <th className="p-3.5 text-center">Sr.</th>
                    <th className="p-3.5">Passport</th>
                    <th className="p-3.5">Full Name</th>
                    <th className="p-3.5 text-center">Stay Status (အခြေအနေ)</th>
                    <th className="p-3.5">Nationality</th>
                    <th className="p-3.5">Registered Address</th>
                    <th className="p-3.5">Sponsor Agent (လာခေါ်သူ)</th>
                    <th className="p-3.5">Visa & Validity</th>
                    <th className="p-3.5 text-center">Remaining Days</th>
                    <th className="p-3.5 text-center">Permit Verification</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-bold text-slate-700 bg-white">
                  {selectedPassports.map((pp, index) => {
                    const m = computedMovementMap[pp] || { p: pp, n: '-', nat: '-', gender: '-', loc: '-', visa: '-', start: '-', end: '-', allowed: '-', agent: '', contact: '', stillPermittedStatus: '' };
                    const rem = getRemainingDaysNum(m.end);
                    const statusInfo = getGuestStayStatus(pp, computedMovementMap);
                    return (
                      <tr key={pp} className="hover:bg-slate-50/50 transition-colors">
                        <td className="p-3.5 text-center font-mono text-slate-400">{index + 1}</td>
                        <td className="p-3.5 font-black text-indigo-900 font-mono tracking-tight">{pp}</td>
                        <td className="p-3.5 uppercase text-slate-800 text-[11px]">{m.n}</td>
                        <td className="p-3.5 text-center">
                          {statusInfo.isStillIn ? (
                            <span className="text-[10px] font-black uppercase text-emerald-800 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200 inline-block shadow-xs">
                              🟢 လက်ရှိနေထိုင်သူ
                            </span>
                          ) : statusInfo.isDeparted ? (
                            <span className="text-[10px] font-black uppercase text-rose-800 bg-rose-50 px-2.5 py-1 rounded-lg border border-rose-200 inline-block">
                              🔴 ထွက်သွားပြီးသူ
                            </span>
                          ) : (
                            <span className="text-[10px] font-black uppercase text-slate-600 bg-slate-100 px-2 py-0.5 rounded border inline-block">
                              ⚪ မှတ်ပုံတင်ထား
                            </span>
                          )}
                          {statusInfo.timeText && (
                            <div className="text-[9px] font-mono font-bold text-indigo-700 mt-0.5">{statusInfo.timeText}</div>
                          )}
                        </td>
                        <td className="p-3.5 uppercase">{m.nat} ({m.gender || '-'})</td>
                        <td className="p-3.5 max-w-[180px] truncate" title={m.loc}>{m.loc}</td>
                        <td className="p-3.5 uppercase text-indigo-950 font-black">
                          <div>{m.agent || '-'}</div>
                          {m.contact && <div className="text-[9px] text-slate-400 font-mono mt-0.5">{m.contact}</div>}
                        </td>
                        <td className="p-3.5 font-mono text-[10px]">
                          <div>{m.visa}</div>
                          <div className="text-gray-400 mt-0.5">{m.start} to {m.end}</div>
                        </td>
                        <td className="p-3.5 text-center font-mono">
                          {rem < 0 ? (
                            <span className="text-[10px] font-black uppercase text-rose-700 bg-rose-50 px-2 py-0.5 rounded-lg border border-rose-200">
                              {rem} Days
                            </span>
                          ) : (
                            <span className="text-[10px] font-black text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                              {rem} Days
                            </span>
                          )}
                        </td>
                        <td className="p-3.5 text-center">
                          {normalizePermitStatus(m.stillPermittedStatus) === 'နေထိုင်ခွင့်ကျထားသောသူ' ? (
                            <span className="text-[9px] font-black uppercase text-emerald-800 bg-emerald-50 px-2 py-1 rounded-lg border border-emerald-200">
                              PERMITTED
                            </span>
                          ) : normalizePermitStatus(m.stillPermittedStatus) === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? (
                            <span className="text-[9px] font-black uppercase text-rose-800 bg-rose-50 px-2 py-1 rounded-lg border border-rose-200">
                              NOT PERMITTED
                            </span>
                          ) : (
                            <span className="text-[9px] font-black uppercase text-slate-600 bg-slate-100 px-2 py-1 rounded border">
                              CONFIRMED
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3.3. DETAILED INDIVIDUAL DOSSIER BOARDS (STACKED AS ACCORDION/EXPANDED CARDS) */}
          <div className="space-y-6 sm:space-y-8 no-print">
            <h3 className="text-xs sm:text-md font-black text-slate-500 uppercase tracking-wider block">Individual Detailed Investigation Cards</h3>
            
            {selectedPassports.map((pp) => {
              const matchedLog = computedMovementMap[pp];
              if (!matchedLog) return null;

              // Complete flight logs history for this person
              const personLogs = records
                .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
                .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

              // Checking history logs for this person
              const personCheckingHistory = checkingHistory
                .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
                .sort((a, b) => b.id - a.id);

              const latestCheck = personCheckingHistory[0] || null;

              return (
                <div key={pp} className="card shadow-lg bg-white border border-slate-100 rounded-2xl sm:rounded-3xl overflow-hidden">
                  
                  {/* Card Header Bar */}
                  <div className="bg-indigo-950 p-4 sm:p-5 text-white flex justify-between items-start sm:items-center flex-col sm:flex-row gap-3 sm:gap-4">
                    <div className="flex items-center gap-3 flex-wrap">
                      <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-full bg-white/10 flex items-center justify-center font-black text-indigo-200 shrink-0">
                        <User size={18} className="sm:w-5 sm:h-5" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <h4 className="text-sm sm:text-base font-black uppercase tracking-tight leading-tight text-white">{matchedLog.n}</h4>
                          {(() => {
                            const statusInfo = getGuestStayStatus(pp, computedMovementMap);
                            return statusInfo.isStillIn ? (
                              <span className="text-[10px] font-black uppercase bg-emerald-500 text-white px-2.5 py-0.5 rounded-full flex items-center gap-1 border border-emerald-300 shadow-sm">
                                <span className="w-1.5 h-1.5 rounded-full bg-white animate-pulse"></span>
                                🟢 {statusInfo.label} (STILL IN)
                              </span>
                            ) : statusInfo.isDeparted ? (
                              <span className="text-[10px] font-black uppercase bg-rose-600 text-white px-2.5 py-0.5 rounded-full flex items-center gap-1 border border-rose-400">
                                🔴 {statusInfo.label} (DEPARTED)
                              </span>
                            ) : (
                              <span className="text-[10px] font-black uppercase bg-slate-700 text-slate-200 px-2.5 py-0.5 rounded-full">
                                ⚪ {statusInfo.label}
                              </span>
                            );
                          })()}
                        </div>
                        <p className="text-[9px] sm:text-[10px] text-indigo-200 uppercase font-mono tracking-wider sm:tracking-widest mt-0.5">PASSPORT: {pp} • {matchedLog.nat}</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                      <button
                        onClick={handlePrintGroupDossier}
                        className="btn bg-indigo-800 hover:bg-white text-white hover:text-slate-900 text-[10px] font-black uppercase px-2.5 sm:px-3.5 py-1.5 rounded-lg border border-indigo-400/30 transition-all flex items-center gap-1 cursor-pointer"
                      >
                        <Printer size={12} /> Print
                      </button>
                      <button
                        onClick={() => openQuickEdit(pp)}
                        className="btn bg-white/10 hover:bg-white text-white hover:text-slate-900 text-[10px] font-black uppercase px-2.5 sm:px-3.5 py-1.5 rounded-lg border border-white/20 transition-all flex items-center gap-1"
                      >
                        <Edit size={12} /> Quick Edit
                      </button>
                      <button
                        onClick={() => removePassportFromGroup(pp)}
                        className="btn bg-rose-900/40 text-rose-200 hover:bg-rose-600 hover:text-white text-[10px] font-black uppercase px-2.5 sm:px-3.5 py-1.5 rounded-lg border border-rose-500/20 transition-all flex items-center gap-1"
                      >
                        <X size={12} strokeWidth={2.5} /> Remove
                      </button>
                    </div>
                  </div>

                  {/* Card Bento grid Details */}
                  <div className="p-4 sm:p-6 grid grid-cols-1 xl:grid-cols-3 gap-4 sm:gap-6">
                    
                    {/* Demographics Block */}
                    <div className="bg-slate-50/50 p-4 sm:p-5 rounded-2xl border border-slate-100 text-xs space-y-3 sm:space-y-4">
                      <h5 className="font-black uppercase text-indigo-900 tracking-wider border-b pb-2 mb-2 text-xs">Demographic Profile / ကိုယ်ရေးရာဇဝင်</h5>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 sm:gap-4">
                        <div>
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Passport Number</span>
                          <span className="font-extrabold font-mono text-indigo-950 uppercase">{pp}</span>
                        </div>
                        <div>
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Nationality / Gender</span>
                          <span className="font-extrabold uppercase text-slate-700">{matchedLog.nat} ({matchedLog.gender || 'Unknown'})</span>
                        </div>
                        <div className="sm:col-span-2">
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Full Guest Name</span>
                          <span className="font-black uppercase text-slate-800">{matchedLog.n}</span>
                        </div>
                        <div className="sm:col-span-2">
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Stay Allowed validity</span>
                          <span className="font-bold text-slate-700 block">{matchedLog.start} ➔ {matchedLog.end}</span>
                          <span className="text-[10px] text-indigo-600 font-extrabold mt-0.5">({matchedLog.allowed} Days permitted)</span>
                        </div>
                        <div>
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Sponsor Agent (လာခေါ်သူ)</span>
                          <span className="font-extrabold text-slate-800 uppercase">{matchedLog.agent || '-'}</span>
                        </div>
                        <div>
                          <span className="text-[9px] sm:text-[10px] font-black text-slate-400 block uppercase tracking-wider">Sponsor Contact Details</span>
                          <span className="font-bold font-mono text-slate-600 block">{matchedLog.contact || '-'}</span>
                        </div>
                      </div>
                    </div>

                    {/* Stay Residence status & Direct Inline Stay Details Editing */}
                    <div className="bg-slate-50/50 p-4 sm:p-5 rounded-2xl border border-slate-100 text-xs space-y-3.5">
                      <div className="flex justify-between items-center border-b pb-2 mb-2">
                        <h5 className="font-black uppercase text-indigo-900 tracking-wider text-xs">Stay Details & Location (နေထိုင်ခွင့် အချက်အလက်)</h5>
                        {editingStayPassport !== pp ? (
                          <button
                            onClick={() => startEditingStayDetails(pp, matchedLog)}
                            className="text-[10px] font-black uppercase text-indigo-700 hover:text-indigo-950 bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded-lg border border-indigo-200 transition-all flex items-center gap-1"
                          >
                            <Edit size={11} /> Edit Stay Info
                          </button>
                        ) : (
                          <button
                            onClick={() => setEditingStayPassport(null)}
                            className="text-[10px] font-black uppercase text-slate-500 hover:text-slate-700 bg-slate-100 px-2 py-1 rounded-lg border border-slate-200"
                          >
                            Cancel
                          </button>
                        )}
                      </div>

                      {editingStayPassport === pp ? (
                        <div className="space-y-3 bg-white p-3 sm:p-4 rounded-xl border border-indigo-200 shadow-sm">
                          <p className="text-[10px] font-black text-indigo-900 uppercase">Direct Stay Details Editor</p>
                          <div>
                            <label className="block text-[9px] font-black uppercase text-slate-500 mb-1">Stay Location Address (တရားဝင်နေရပ်လိပ်စာ - ရွေးချယ်နိုင်သည်)</label>
                            <input
                              type="text"
                              list="invStayAddressDatalist"
                              value={inlineStayEdits[pp]?.loc ?? (matchedLog.loc || '')}
                              onChange={(e) => setInlineStayEdits(prev => ({ ...prev, [pp]: { ...prev[pp], loc: e.target.value } }))}
                              className="w-full text-xs font-semibold p-2.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none bg-slate-50/50 hover:bg-white"
                              placeholder="Select or type stay location address (တည်းခိုလိပ်စာ ရွေးရန် သို့မဟုတ် ရိုက်ထည့်ရန်)..."
                            />
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div>
                              <label className="block text-[9px] font-black uppercase text-slate-500 mb-1">Visa Type (ဗီဇာအမျိုးအစား)</label>
                              <input
                                type="text"
                                value={inlineStayEdits[pp]?.visa ?? (matchedLog.visa || '')}
                                onChange={(e) => setInlineStayEdits(prev => ({ ...prev, [pp]: { ...prev[pp], visa: e.target.value } }))}
                                className="w-full text-xs font-semibold p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none"
                                placeholder="e.g. BUSINESS"
                              />
                            </div>
                            <div>
                              <label className="block text-[9px] font-black uppercase text-slate-500 mb-1">Allowed Days (ခွင့်ပြုရက်)</label>
                              <input
                                type="text"
                                value={inlineStayEdits[pp]?.allowed ?? (matchedLog.allowed || '')}
                                onChange={(e) => setInlineStayEdits(prev => ({ ...prev, [pp]: { ...prev[pp], allowed: e.target.value } }))}
                                className="w-full text-xs font-semibold p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none"
                                placeholder="e.g. 70"
                              />
                            </div>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div>
                              <label className="block text-[9px] font-black uppercase text-slate-500 mb-1">Stay From Date (စတင်ရက်)</label>
                              <input
                                type="text"
                                value={inlineStayEdits[pp]?.start ?? (matchedLog.start || '')}
                                onChange={(e) => setInlineStayEdits(prev => ({ ...prev, [pp]: { ...prev[pp], start: e.target.value } }))}
                                className="w-full text-xs font-semibold p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none"
                                placeholder="YYYY-MM-DD"
                              />
                            </div>
                            <div>
                              <label className="block text-[9px] font-black uppercase text-slate-500 mb-1">Stay To Date (ကုန်ဆုံးရက်)</label>
                              <input
                                type="text"
                                value={inlineStayEdits[pp]?.end ?? (matchedLog.end || '')}
                                onChange={(e) => setInlineStayEdits(prev => ({ ...prev, [pp]: { ...prev[pp], end: e.target.value } }))}
                                className="w-full text-xs font-semibold p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none"
                                placeholder="YYYY-MM-DD"
                              />
                            </div>
                          </div>
                          <button
                            onClick={() => handleSaveInlineStayDetails(pp)}
                            className="w-full bg-indigo-900 hover:bg-black text-white font-extrabold uppercase text-[10px] tracking-wider py-2 rounded-lg transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                          >
                            <Save size={13} /> Save Stay Info (သိမ်းဆည်းမည်)
                          </button>
                        </div>
                      ) : (
                        <div className="space-y-3">
                          <div>
                            <span className="text-[10px] font-black text-slate-400 block uppercase tracking-wider">Registered Stay Address</span>
                            <p className="font-bold text-slate-800 leading-tight mt-1 bg-white p-2.5 rounded-xl border border-slate-200">{matchedLog.loc || '-'}</p>
                          </div>
                          <div className="grid grid-cols-2 gap-2 text-[11px]">
                            <div>
                              <span className="text-[9px] font-black text-slate-400 block uppercase">Visa Category</span>
                              <span className="font-extrabold text-indigo-950 uppercase">{matchedLog.visa || '-'}</span>
                            </div>
                            <div>
                              <span className="text-[9px] font-black text-slate-400 block uppercase">Validity Window</span>
                              <span className="font-bold text-slate-700 block text-[10px]">{matchedLog.start || '-'} ➔ {matchedLog.end || '-'}</span>
                            </div>
                          </div>
                          <div>
                            <span className="text-[10px] font-black text-slate-400 block uppercase tracking-wider font-mono">Permit Validation Status</span>
                            <div className="mt-1">
                              {normalizePermitStatus(matchedLog.stillPermittedStatus) === 'နေထိုင်ခွင့်ကျထားသောသူ' ? (
                                <div className="flex items-center gap-1.5 bg-emerald-50 text-emerald-800 border-2 border-emerald-100 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-tight block w-max">
                                  <CheckCircle2 size={13} /> နေထိုင်ခွင့်ကျထားသောသူ
                                </div>
                              ) : normalizePermitStatus(matchedLog.stillPermittedStatus) === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? (
                                <div className="flex items-center gap-1.5 bg-rose-50 text-rose-800 border-2 border-rose-100 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-tight block w-max">
                                  <ShieldAlert size={13} /> နေထိုင်ခွင့် မလျှောက်ထားသေးသူ
                                </div>
                              ) : (
                                <div className="flex items-center gap-1.5 bg-blue-50 text-blue-900 border-2 border-blue-100 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-tight block w-max">
                                  <AlertTriangle size={13} /> CONFIRMED / STILL IN
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Investigation Notes (စုံစမ်းစစ်ဆေးချက်) with Auto-Timestamp & Author Tracking */}
                    <div className="bg-slate-50/50 p-4 sm:p-5 rounded-2xl border border-slate-100 text-xs space-y-3.5">
                      <div className="flex justify-between items-center border-b pb-2 mb-2">
                        <h5 className="font-black uppercase text-indigo-900 tracking-wider text-xs">Investigation Notes (စုံစမ်းစစ်ဆေးချက်)</h5>
                        <span className="text-[9px] font-mono font-bold text-slate-500 bg-slate-200 px-2 py-0.5 rounded">
                          {currentUser?.title ? `${currentUser.title} ${currentUser.name || ''}` : (currentUser?.name || 'Duty Officer')}
                        </span>
                      </div>

                      {/* Auto-Timestamp Note Add Form */}
                      <div className="bg-white p-3 rounded-xl border border-indigo-100 shadow-sm space-y-2">
                        <label className="block text-[9.5px] font-black uppercase text-slate-600 tracking-wider flex items-center justify-between">
                          <span>Add New Remark (မှတ်ချက်အသစ် ရေးသားပါ):</span>
                          <span className="text-[8px] font-normal text-slate-400 font-mono">Auto Timestamp & Author</span>
                        </label>
                        <textarea
                          rows={2}
                          value={newNoteInputs[pp] || ''}
                          onChange={(e) => {
                            const val = e.target.value;
                            setNewNoteInputs(prev => ({ ...prev, [pp]: val }));
                          }}
                          placeholder="စုံစမ်းစစ်ဆေးတွေ့ရှိချက် မှတ်ချက်အသစ် ရေးသားပါ..."
                          className="w-full text-xs font-semibold p-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-600 outline-none text-slate-800"
                        />
                        <button
                          onClick={() => handleAddAutoNoteForPassport(pp)}
                          className="w-full bg-indigo-900 hover:bg-black text-white font-extrabold uppercase text-[10px] tracking-wider py-2 rounded-lg transition-colors flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                        >
                          <Plus size={13} /> Add Note with Auto-Timestamp
                        </button>
                      </div>

                      {/* Full Investigation Notes Log Box */}
                      <div className="space-y-1.5">
                        <label className="block text-[9px] font-black uppercase text-slate-400 tracking-wider">All Recorded Notes (မှတ်ချက်မှတ်တမ်းများ):</label>
                        <textarea
                          value={currentRemarks[pp] !== undefined ? currentRemarks[pp] : (matchedLog.remarks || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setCurrentRemarks(prev => ({ ...prev, [pp]: val }));
                          }}
                          onBlur={(e) => {
                            saveRemarkForPassport(pp, e.target.value);
                          }}
                          placeholder="All investigation notes and intelligence logs will appear here..."
                          className="w-full text-xs font-semibold p-2.5 border border-slate-200 rounded-xl focus:border-indigo-600 focus:outline-none transition-all h-28 bg-white resize-y text-slate-800 leading-relaxed font-mono"
                        />
                        <button
                          onClick={() => {
                            const val = currentRemarks[pp] || matchedLog.remarks || '';
                            saveRemarkForPassport(pp, val);
                          }}
                          className="w-full bg-slate-800 text-white font-extrabold uppercase text-[9.5px] tracking-wider py-1.5 rounded-lg hover:bg-black transition-colors"
                        >
                          Save All Notes (မှတ်ချက်အားလုံး သိမ်းဆည်းမည်)
                        </button>
                      </div>
                    </div>

                    {/* Historical movement logs list table */}
                    <div className="col-span-1 xl:col-span-3">
                      <div className="border border-slate-100 rounded-2xl overflow-hidden bg-slate-50/10">
                        <div className="bg-slate-50 p-3.5 border-b flex justify-between items-center">
                          <h6 className="text-[10px] font-black uppercase text-slate-600 tracking-wider">Historical Travel Movements log ({personLogs.length} Events)</h6>
                        </div>
                        {personLogs.length === 0 ? (
                          <div className="p-4 text-center text-slate-400 font-bold">No travel history logged</div>
                        ) : (
                          <div className="overflow-x-auto text-[11px] max-h-48 overflow-y-auto">
                            <table className="w-full text-left">
                              <thead className="bg-white/90 sticky top-0 border-b uppercase text-[8px] font-black text-slate-400">
                                <tr>
                                  <th className="p-2 pl-4">Timestamp</th>
                                  <th className="p-2 text-center">Dir</th>
                                  <th className="p-2">Visa Type</th>
                                  <th className="p-2">Residence Stay Address</th>
                                  <th className="p-2 font-mono">Carrier Vehicle</th>
                                  <th className="p-2">Assigned Officer</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-50 bg-white text-slate-700 font-bold">
                                {personLogs.map((log) => (
                                  <tr key={log.id} className="hover:bg-slate-50/30 transition-colors">
                                    <td className="p-2 pl-4 text-slate-900 font-semibold">{log.timestamp}</td>
                                    <td className="p-2 text-center">
                                      <span className={`px-1.5 py-0.5 rounded text-[8px] font-black ${
                                        log.mode === 'IN' ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-orange-50 text-orange-850 border border-orange-200'
                                      }`}>{log.mode}</span>
                                    </td>
                                    <td className="p-2 uppercase text-[10px]">{log.visaType}</td>
                                    <td className="p-2 font-semibold text-slate-600">{log.address || '-'}</td>
                                    <td className="p-2 font-mono text-indigo-900">{log.vehicleInfo || '-'}</td>
                                    <td className="p-2 text-[9px] text-slate-500">{log.officialTitle} {log.officialName}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* On site physical check records table */}
                    {personCheckingHistory.length > 0 && (
                      <div className="col-span-1 xl:col-span-3">
                        <div className="border border-slate-100 rounded-2xl overflow-hidden bg-slate-50/10">
                          <div className="bg-slate-50 p-3.5 border-b">
                            <h6 className="text-[10px] font-black uppercase text-slate-600 tracking-wider">Physical Verification Records ({personCheckingHistory.length} Checks)</h6>
                          </div>
                          <div className="overflow-x-auto text-[11px] max-h-48 overflow-y-auto">
                            <table className="w-full text-left">
                              <thead className="bg-white/90 sticky top-0 border-b uppercase text-[8px] font-black text-slate-400">
                                <tr>
                                  <th className="p-2 pl-4">Check Date</th>
                                  <th className="p-2">Status</th>
                                  <th className="p-2">Residence Confirmed Address</th>
                                  <th className="p-2">Description</th>
                                  <th className="p-2">Verification Officer</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-50 bg-white text-slate-700 font-bold">
                                {personCheckingHistory.map((ch, idx) => (
                                  <tr key={idx} className="hover:bg-slate-50/30 transition-colors">
                                    <td className="p-2 pl-4 text-slate-900 font-semibold">{ch.checkDate}</td>
                                    <td className="p-2">
                                      <span className={`px-1.5 py-0.5 rounded text-[8px] font-black ${
                                        normalizePermitStatus(ch.status) === 'နေထိုင်ခွင့်ကျထားသောသူ' 
                                          ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' 
                                          : 'bg-rose-50 text-rose-800 border border-rose-200'
                                      }`}>{normalizePermitStatus(ch.status) || 'နေထိုင်ခွင့်ကျထားသောသူ'}</span>
                                    </td>
                                    <td className="p-2 font-semibold text-slate-600">{ch.confirmedAddress || '-'}</td>
                                    <td className="p-2 italic text-slate-500 font-semibold">{ch.confirmedStayDescription || '-'}</td>
                                    <td className="p-2 text-[9px] text-slate-500">{ch.officerTitle} {ch.officerName}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    )}

                  </div>

                </div>
              );
            })}
          </div>

        </div>
      )}

      {/* 4. QUICK EDIT DIALOG MODAL PANEL */}
      {quickEditPassport && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm no-print">
          <div className="relative bg-white w-full max-w-xl rounded-[2rem] overflow-hidden shadow-2xl border-4 border-indigo-900 flex flex-col max-h-[90vh] animate-in zoom-in duration-200">
            
            <div className="bg-indigo-900 p-5 text-white flex justify-between items-center">
              <div className="flex items-center gap-2">
                <Edit size={20} />
                <div>
                  <h3 className="font-black text-sm uppercase tracking-wider">Quick Profile & Stay Editor</h3>
                  <p className="text-[10px] text-indigo-200 uppercase font-bold mt-0.5">Editing Target Passport: {quickEditPassport}</p>
                </div>
              </div>
              <button 
                onClick={() => setQuickEditPassport(null)} 
                className="hover:bg-white/10 p-1.5 rounded-full text-white transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-4 text-xs">
              <div className="grid grid-cols-2 gap-4">
                
                <div className="col-span-2 space-y-1.5">
                  <label className="input-label">Full Name / အမည်</label>
                  <input
                    type="text"
                    value={editFullName}
                    onChange={(e) => setEditFullName(e.target.value.toUpperCase())}
                    className="input-field py-2.5 font-bold uppercase"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Nationality / နိုင်ငံသား</label>
                  <input
                    type="text"
                    value={editNationality}
                    onChange={(e) => setEditNationality(e.target.value.toUpperCase())}
                    className="input-field py-2.5 font-bold uppercase"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Gender / ကျား-မ</label>
                  <select
                    value={editGender}
                    onChange={(e: any) => setEditGender(e.target.value)}
                    className="input-field py-2.5 font-bold"
                  >
                    <option value="M">MALE / ကျား</option>
                    <option value="F">FEMALE / မ</option>
                    <option value="">OTHER / အခြား</option>
                  </select>
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Date of Birth (DOB) / မွေးသက္ကရာဇ်</label>
                  <input
                    type="date"
                    value={editDob}
                    onChange={(e) => setEditDob(e.target.value)}
                    className="input-field py-2.5 font-bold font-mono"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Visa Number / ဗီဇာအမှတ်</label>
                  <input
                    type="text"
                    value={editVisaNumber}
                    onChange={(e) => setEditVisaNumber(e.target.value.toUpperCase())}
                    className="input-field py-2.5 font-bold uppercase font-mono"
                    placeholder="e.g. V-123456"
                  />
                </div>

                <div className="col-span-2 space-y-1.5">
                  <label className="input-label">Residence Stay Address / တရားဝင်နေရပ်လိပ်စာ (ရွေးချယ်နိုင်သည်)</label>
                  <input
                    type="text"
                    list="invStayAddressDatalist"
                    value={editAddress}
                    onChange={(e) => setEditAddress(e.target.value)}
                    className="input-field py-2.5 font-bold"
                    placeholder="Select or type stay location..."
                  />
                </div>

                <div className="col-span-2 space-y-1.5">
                  <label className="input-label">Visa Classification / ဗီဇာအမျိုးအစား</label>
                  <input
                    type="text"
                    value={editVisaType}
                    onChange={(e) => setEditVisaType(e.target.value.toUpperCase())}
                    className="input-field py-2.5 font-bold uppercase"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Stay From Date / နေထိုင်ခွင့်စတင်ရက်</label>
                  <input
                    type="date"
                    value={editStayFrom}
                    onChange={(e) => setEditStayFrom(e.target.value)}
                    className="input-field py-2.5 font-bold font-mono"
                  />
                </div>

                <div className="space-y-1.5">
                  <label className="input-label">Stay To Date / နေထိုင်ခွင့်ကုန်ဆုံးရက်</label>
                  <input
                    type="date"
                    value={editStayTo}
                    onChange={(e) => setEditStayTo(e.target.value)}
                    className="input-field py-2.5 font-bold font-mono"
                  />
                </div>

                <div className="col-span-2 space-y-1.5">
                  <label className="input-label">Allowed Days / ခွင့်ပြုရက်အရေအတွက်</label>
                  <input
                    type="number"
                    value={editTotalDays}
                    onChange={(e) => setEditTotalDays(e.target.value)}
                    className="input-field py-2.5 font-bold font-mono"
                    placeholder="e.g. 70"
                  />
                </div>

              </div>
            </div>

            <div className="p-6 bg-slate-50 border-t flex gap-3 shrink-0">
              <button 
                onClick={() => setQuickEditPassport(null)} 
                className="btn bg-white hover:bg-slate-200 text-slate-700 flex-1 uppercase py-3 font-black rounded-xl border border-slate-300"
              >
                Cancel / ပယ်ဖျက်မည်
              </button>
              <button 
                onClick={handleQuickSave} 
                className="btn bg-indigo-900 hover:bg-black text-white flex-1 uppercase py-3 font-black rounded-xl shadow-lg"
              >
                ✓ Save Updates / အတည်ပြုမည်
              </button>
            </div>

          </div>
        </div>
      )}

      {/* 5. GORGEOUS INTERACTIVE VISUAL PRINT PREVIEW OVERLAY (CUSTOM BUILT FOR MULTI-SEARCH / GROUPS!) */}
      {showPrintPreviewModal && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-md flex items-center justify-center p-4 md:p-8 z-50 no-print animate-in fade-in duration-200">
          <div className="bg-slate-950 text-white rounded-2xl w-full max-w-5xl h-full max-h-[95vh] flex flex-col shadow-2xl border border-slate-800 overflow-hidden">
            
            {/* Header Control Panel */}
            <div className="bg-slate-900 p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="text-xl">🖨️</span>
                <div>
                  <h3 className="font-black text-sm uppercase tracking-wide">Investigation Group Visual Print Preview</h3>
                  <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">
                    Target Count: {selectedPassports.length} Investigated Subjects
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2.5">
                <button
                  onClick={handlePrintGroupDossier}
                  className="btn bg-indigo-600 text-white hover:bg-indigo-700 font-black text-xs uppercase tracking-wider px-5 py-2.5 rounded-xl shadow-md flex items-center gap-1.5 transition-all cursor-pointer"
                >
                  <Printer size={14} /> CONFIRM & PRINT NOW
                </button>
                <button
                  onClick={() => setShowPrintPreviewModal(false)}
                  className="btn bg-slate-800 text-slate-300 hover:bg-slate-700 font-black text-xs uppercase tracking-wider px-5 py-2.5 rounded-xl border border-slate-700 transition-all cursor-pointer"
                >
                  CLOSE PREVIEW
                </button>
              </div>
            </div>

            {/* Virtual Paper Preview Body */}
            <div className="flex-1 bg-slate-900 p-6 overflow-y-auto">
              <div id="visual-print-page" className="bg-white text-black p-8 md:p-12 shadow-2xl rounded-xl mx-auto max-w-4xl text-xs font-sans leading-relaxed select-text min-h-[1100px]">
                
                {/* 1. Office Header */}
                <div className="text-center space-y-1.5 mb-8 border-b-2 border-black pb-5">
                  <h1 className="text-md font-black uppercase text-slate-900 tracking-tight leading-none">
                    Second Division, Department of Immigration (Myeik)
                  </h1>
                  <h2 className="text-xs font-black text-indigo-900 uppercase tracking-widest mt-1.5">
                    {selectedPassports.length === 1 
                      ? 'Individual Investigation & Stay Verification Report'
                      : 'Group Investigation & Stay Verification Report'}
                  </h2>
                  <p className="text-[10px] font-bold text-indigo-950 font-mono pt-1">
                    DOSSIER IDENTIFICATION NO: {activeDossierID}
                  </p>
                  <p className="text-[8px] font-mono font-bold text-slate-400 uppercase tracking-wide">
                    SYSTEM SECURITY REGISTER COPY • TIMESTAMP: {new Date().toLocaleString()}
                  </p>
                </div>

                {/* 2. Group Summary Comparison Matrix (Only if multiple selected) */}
                {selectedPassports.length > 1 && (
                  <div className="mb-8">
                    <h3 className="text-[10px] font-black uppercase tracking-wider mb-2 border-b pb-1">Group Comparison Overview Matrix</h3>
                    <table className="w-full text-left text-[9px] border border-gray-400 border-collapse">
                      <thead>
                        <tr className="bg-slate-100 uppercase text-[8px] font-black">
                          <th className="p-1 border border-gray-400 text-center">#</th>
                          <th className="p-1 border border-gray-400">Guest Name / Passport</th>
                          <th className="p-1 border border-gray-400 text-center">Stay Status</th>
                          <th className="p-1 border border-gray-400">Nationality</th>
                          <th className="p-1 border border-gray-400">Residence Stay Address</th>
                          <th className="p-1 border border-gray-400">Visa & Validity</th>
                          <th className="p-1 border border-gray-400">Local Sponsor</th>
                          <th className="p-1 border border-gray-400 text-center">Remaining Days</th>
                        </tr>
                      </thead>
                      <tbody>
                        {selectedPassports.map((pp, idx) => {
                          const m = computedMovementMap[pp] || {};
                          const rem = getRemainingDaysNum(m.end);
                          const statusInfo = getGuestStayStatus(pp, computedMovementMap);
                          return (
                            <tr key={pp} className="align-top border-b border-gray-300">
                              <td className="p-1.5 border border-gray-400 text-center font-bold">{idx + 1}</td>
                              <td className="p-1.5 border border-gray-400">
                                <div className="font-extrabold font-mono uppercase text-indigo-900 leading-none">{pp}</div>
                                <div className="font-bold text-gray-800 uppercase text-[8.5px] mt-0.5">{m.n || '-'}</div>
                              </td>
                              <td className="p-1.5 border border-gray-400 text-center font-bold text-[8px]">
                                {statusInfo.isStillIn ? (
                                  <span className="text-emerald-800 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-300 font-black">
                                    🟢 STILL IN
                                  </span>
                                ) : statusInfo.isDeparted ? (
                                  <span className="text-rose-800 bg-rose-50 px-1.5 py-0.5 rounded border border-rose-300 font-black">
                                    🔴 DEPARTED
                                  </span>
                                ) : (
                                  <span className="text-slate-600 bg-slate-100 px-1 py-0.5 rounded">
                                    ⚪ REGISTERED
                                  </span>
                                )}
                              </td>
                              <td className="p-1.5 border border-gray-400 font-bold uppercase text-gray-700">{m.nat || '-'}</td>
                              <td className="p-1.5 border border-gray-400 leading-tight max-w-[150px] truncate">{m.loc || '-'}</td>
                              <td className="p-1.5 border border-gray-400 leading-none">
                                <div className="font-bold">{m.visa || '-'}</div>
                                <div className="text-[7.5px] font-mono text-gray-500 mt-0.5">{m.start} to {m.end}</div>
                              </td>
                              <td className="p-1.5 border border-gray-400 uppercase font-black text-gray-800 leading-tight">
                                <div>{m.agent || '-'}</div>
                                {m.contact && <div className="text-[7.5px] font-mono text-gray-500 font-normal mt-0.5">{m.contact}</div>}
                              </td>
                              <td className="p-1.5 border border-gray-400 text-center font-bold font-mono">
                                <span className={rem < 0 ? 'text-rose-600' : 'text-slate-800'}>{rem} Days</span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                {/* 3. Detailed individual profiles (Page breaks between profiles during print) */}
                <div className="space-y-12">
                  {selectedPassports.map((pp, idx) => {
                    const bio = computedMovementMap[pp];
                    if (!bio) return null;

                    const personLogs = records
                      .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
                      .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

                    const personCheckingHistory = checkingHistory
                      .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
                      .sort((a, b) => b.id - a.id);

                    return (
                      <div key={pp} className={`space-y-6 ${idx > 0 ? 'print-page-break pt-6 border-t border-dashed border-gray-300' : ''}`}>
                        
                        {/* Profile Header Tag */}
                        <div className="bg-slate-100 p-2 border-b-2 border-slate-900 flex justify-between items-center">
                          <strong className="text-xs uppercase tracking-tight text-indigo-950 font-black">
                            Dossier File #{idx + 1}: {bio.n} ({pp})
                          </strong>
                          <span className="text-[8px] font-black text-slate-500 uppercase tracking-widest font-mono">Immigration Intel record</span>
                        </div>

                        {/* Demographical table list info */}
                        <div className="grid grid-cols-2 gap-y-3.5 gap-x-6 text-[10.5px] leading-relaxed">
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Passport Number</strong>
                            <span className="font-black font-mono uppercase text-indigo-950 text-xs tracking-tight">{pp}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Full Name</strong>
                            <span className="font-black uppercase text-slate-900 text-xs">{bio.n}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Nationality</strong>
                            <span className="font-extrabold uppercase text-slate-800">{bio.nat || '-'}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Gender</strong>
                            <span className="font-extrabold uppercase text-slate-800">{bio.gender === 'M' ? 'MALE' : bio.gender === 'F' ? 'FEMALE' : '-'}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Local Agent (Broker)</strong>
                    <span className="font-black uppercase text-indigo-900">{bio.agent || '-'}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Contact Details</strong>
                            <span className="font-bold text-slate-700 font-mono">{bio.contact || '-'}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Visa Category</strong>
                            <span className="font-extrabold text-slate-800">{bio.visa || '-'}</span>
                          </div>
                          <div>
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Stay Schedule Duration</strong>
                            <span className="font-black text-slate-900 font-mono">{bio.start} to {bio.end} ({bio.allowed || '-'} Days)</span>
                          </div>
                          <div className="col-span-2">
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Registered Destination Address</strong>
                            <span className="font-bold text-slate-900 text-xs">{bio.loc || '-'}</span>
                          </div>
                          <div className="col-span-2">
                            <strong className="text-slate-400 block text-[8px] uppercase font-black">Address Description Details</strong>
                            <span className="font-semibold text-slate-700 italic block">{bio.stayDescription || 'No description logged.'}</span>
                          </div>

                          {/* Custom Investigation remarks */}
                          <div className="col-span-2 bg-amber-50/50 p-3 rounded-xl border border-amber-250">
                            <strong className="text-amber-800 block text-[8px] uppercase font-black tracking-wide mb-0.5">Dossier Remark Details</strong>
                            <p className="font-black text-amber-950 font-sans text-[11px] leading-tight">
                              {currentRemarks[pp] || bio.remarks || 'No special remarks registered.'}
                            </p>
                          </div>
                        </div>

                        {/* Logs tables history */}
                        <div className="space-y-4">
                          <h4 className="text-[8px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Flight Movement logs</h4>
                          <table className="w-full text-left text-[8.5px] border border-gray-400 border-collapse table-auto">
                            <thead>
                              <tr className="bg-slate-150 uppercase text-[7.5px] font-black">
                                <th className="p-1 border border-gray-400 text-center">#</th>
                                <th className="p-1 border border-gray-400">Timestamp</th>
                                <th className="p-1 border border-gray-400 text-center">Dir</th>
                                <th className="p-1 border border-gray-400">Visa & Stay Duration</th>
                                <th className="p-1 border border-gray-400">Stay Address / Location</th>
                                <th className="p-1 border border-gray-400">Registered Officer</th>
                              </tr>
                            </thead>
                            <tbody>
                              {personLogs.map((log, idx) => (
                                <tr key={log.id} className="border-b border-gray-300">
                                  <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                                  <td className="p-1 border border-gray-400 font-mono text-[8px]">{log.timestamp}</td>
                                  <td className="p-1 border border-gray-400 text-center">
                                    <span className={`px-1 py-0.5 rounded text-[7.5px] font-black ${log.mode === 'IN' ? 'bg-emerald-50 text-emerald-800' : 'bg-orange-50 text-orange-850'}`}>
                                      {log.mode}
                                    </span>
                                  </td>
                                  <td className="p-1 border border-gray-400 leading-tight">
                                    <div className="font-bold">{log.visaType}</div>
                                    {log.stayFrom && <div className="text-[7.5px] text-gray-500 font-mono mt-0.5">{log.stayFrom} to {log.stayTo} ({log.totalDays} Days)</div>}
                                  </td>
                                  <td className="p-1 border border-gray-400 leading-tight">
                                    <div className="font-bold text-gray-800">{log.address || '-'}</div>
                                  </td>
                                  <td className="p-1 border border-gray-400 text-gray-500 font-bold leading-none text-[8px]">
                                    <div>{log.officialTitle || ''}</div>
                                    <div className="text-gray-800 mt-0.5">{log.officialName || 'System'}</div>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Physical checking verification history logs */}
                        {personCheckingHistory.length > 0 && (
                          <div className="space-y-4">
                            <h4 className="text-[8px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Physical Verification annals</h4>
                            <table className="w-full text-left text-[8.5px] border border-gray-400 border-collapse table-auto">
                              <thead>
                                <tr className="bg-slate-150 uppercase text-[7.5px] font-black">
                                  <th className="p-1 border border-gray-400 text-center">#</th>
                                  <th className="p-1 border border-gray-400">Check Date</th>
                                  <th className="p-1 border border-gray-400">Verified Status</th>
                                  <th className="p-1 border border-gray-400">Residence Confirmed Address</th>
                                  <th className="p-1 border border-gray-400">Stay Description Detail</th>
                                  <th className="p-1 border border-gray-400">Verifier Stamp</th>
                                </tr>
                              </thead>
                              <tbody>
                                {personCheckingHistory.map((ch, idx) => (
                                  <tr key={idx} className="border-b border-gray-300">
                                    <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                                    <td className="p-1 border border-gray-400 font-mono text-[8px]">{ch.checkDate}</td>
                                    <td className="p-1 border border-gray-400">
                                      <span className={`px-1 py-0.5 rounded text-[7.5px] font-black ${
                                        normalizePermitStatus(ch.status) === 'နေထိုင်ခွင့်ကျထားသောသူ' 
                                          ? 'bg-emerald-50 text-emerald-800' 
                                          : 'bg-rose-50 text-rose-800'
                                      }`}>{normalizePermitStatus(ch.status) || 'နေထိုင်ခွင့်ကျထားသောသူ'}</span>
                                    </td>
                                    <td className="p-1 border border-gray-400 font-bold text-gray-800">{ch.confirmedAddress || '-'}</td>
                                    <td className="p-1 border border-gray-400 text-gray-500 font-semibold italic">{ch.confirmedStayDescription || '-'}</td>
                                    <td className="p-1 border border-gray-400 font-mono text-[7.5px] text-gray-500 leading-none">
                                      <div>{ch.officerTitle || 'Officer'}</div>
                                      <div className="text-gray-800 font-bold mt-0.5">{ch.officerName || '-'}</div>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}

                      </div>
                    );
                  })}
                </div>

                {/* 4. Print Signatures and Confirmation Boxes */}
                <div className="grid grid-cols-3 gap-8 px-4 mt-20 no-print-inside">
                  <div className="text-center space-y-12">
                    <p className="text-[10px] font-black uppercase text-slate-500">Prepared By</p>
                    <div className="border-t border-slate-400 pt-2 text-[10px] font-bold text-slate-800">
                      Investigation Analyst Officer
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
                      Confirmed & Certified Box / Stamp
                    </p>
                    <div className="space-y-4 text-left text-[9px] font-bold text-slate-600">
                      <div>Office Stamp: __________________</div>
                      <div>Supervisor Signature: ________________</div>
                      <div>Certified Date: {new Date().toLocaleDateString()}</div>
                    </div>
                  </div>
                </div>

                {/* Document Footer */}
                <div className="mt-16 pt-6 border-t border-dashed border-gray-300 text-center text-[9px] text-gray-400 leading-normal">
                  <p>Myeik Immigration Enforcement Unit • Intelligence Group Investigation dossier</p>
                  <p>CONFIDENTIAL • DEPT OF REGISTRATION RECORD • STRICTLY SECRET</p>
                </div>

              </div>
            </div>
          </div>
        </div>
      )}

      {/* 6. Standard Hidden Print Layout (Used for direct printing) */}
      <div id="temp-print-content-source" className="hidden print:block bg-white p-8 text-black min-h-screen w-full text-xs font-sans print-content leading-relaxed">
        {/* Render same printable space inside hidden print container */}
        <div className="text-center space-y-1.5 mb-8 border-b-2 border-black pb-5">
          <h1 className="text-md font-black uppercase text-slate-900 tracking-tight leading-none">
            Second Division, Department of Immigration (Myeik)
          </h1>
          <h2 className="text-xs font-black text-indigo-900 uppercase tracking-widest mt-1.5">
            {selectedPassports.length === 1 
              ? 'Individual Investigation & Stay Verification Report'
              : 'Group Investigation & Stay Verification Report'}
          </h2>
          <p className="text-[10px] font-bold text-indigo-950 font-mono pt-1">
            DOSSIER IDENTIFICATION NO: {activeDossierID}
          </p>
          <p className="text-[8px] font-mono font-bold text-slate-400 uppercase tracking-wide">
            SYSTEM SECURITY REGISTER COPY • TIMESTAMP: {new Date().toLocaleString()}
          </p>
        </div>

        {selectedPassports.length > 1 && (
          <div className="mb-8">
            <h3 className="text-[10px] font-black uppercase tracking-wider mb-2 border-b pb-1">Group Comparison Overview Matrix</h3>
            <table className="w-full text-left text-[9px] border border-gray-400 border-collapse">
              <thead>
                <tr className="bg-slate-100 uppercase text-[8px] font-black">
                  <th className="p-1 border border-gray-400 text-center">#</th>
                  <th className="p-1 border border-gray-400">Guest Name / Passport</th>
                  <th className="p-1 border border-gray-400">Nationality</th>
                  <th className="p-1 border border-gray-400">Residence Stay Address</th>
                  <th className="p-1 border border-gray-400">Visa & Validity</th>
                  <th className="p-1 border border-gray-400">Local Sponsor</th>
                  <th className="p-1 border border-gray-400 text-center">Remaining Days</th>
                </tr>
              </thead>
              <tbody>
                {selectedPassports.map((pp, idx) => {
                  const m = computedMovementMap[pp] || {};
                  const rem = getRemainingDaysNum(m.end);
                  return (
                    <tr key={pp} className="align-top border-b border-gray-300">
                      <td className="p-1.5 border border-gray-400 text-center font-bold">{idx + 1}</td>
                      <td className="p-1.5 border border-gray-400">
                        <div className="font-extrabold font-mono uppercase text-indigo-900 leading-none">{pp}</div>
                        <div className="font-bold text-gray-800 uppercase text-[8.5px] mt-0.5">{m.n || '-'}</div>
                      </td>
                      <td className="p-1.5 border border-gray-400 font-bold uppercase text-gray-700">{m.nat || '-'}</td>
                      <td className="p-1.5 border border-gray-400 leading-tight max-w-[150px] truncate">{m.loc || '-'}</td>
                      <td className="p-1.5 border border-gray-400 leading-none">
                        <div className="font-bold">{m.visa || '-'}</div>
                        <div className="text-[7.5px] font-mono text-gray-500 mt-0.5">{m.start} to {m.end}</div>
                      </td>
                      <td className="p-1.5 border border-gray-400 uppercase font-black text-gray-800 leading-tight">
                        <div>{m.agent || '-'}</div>
                        {m.contact && <div className="text-[7.5px] font-mono text-gray-500 font-normal mt-0.5">{m.contact}</div>}
                      </td>
                      <td className="p-1.5 border border-gray-400 text-center font-bold font-mono">
                        <span className={rem < 0 ? 'text-rose-600' : 'text-slate-800'}>{rem} Days</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <div className="space-y-12">
          {selectedPassports.map((pp, idx) => {
            const bio = computedMovementMap[pp];
            if (!bio) return null;

            const personLogs = records
              .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
              .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

            const personCheckingHistory = checkingHistory
              .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
              .sort((a, b) => b.id - a.id);

            return (
              <div key={pp} className={`space-y-6 ${idx > 0 ? 'print-page-break pt-6 border-t border-dashed border-gray-300' : ''}`}>
                <div className="bg-slate-100 p-2 border-b-2 border-slate-900 flex justify-between items-center">
                  <strong className="text-xs uppercase tracking-tight text-indigo-950 font-black">
                    Dossier File #{idx + 1}: {bio.n} ({pp})
                  </strong>
                  <span className="text-[8px] font-black text-slate-500 uppercase tracking-widest font-mono">Immigration Intel record</span>
                </div>

                <div className="grid grid-cols-2 gap-y-3.5 gap-x-6 text-[10.5px] leading-relaxed">
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Passport Number</strong>
                    <span className="font-black font-mono uppercase text-indigo-950 text-xs tracking-tight">{pp}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Full Name</strong>
                    <span className="font-black uppercase text-slate-900 text-xs">{bio.n}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Nationality</strong>
                    <span className="font-extrabold uppercase text-slate-800">{bio.nat || '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Gender</strong>
                    <span className="font-extrabold uppercase text-slate-800">{bio.gender === 'M' ? 'MALE' : bio.gender === 'F' ? 'FEMALE' : '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Local Agent (Broker)</strong>
                    <span className="font-black uppercase text-indigo-900">{bio.agent || '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Contact Details</strong>
                    <span className="font-bold text-slate-700 font-mono">{bio.contact || '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Visa Category</strong>
                    <span className="font-extrabold text-slate-800">{bio.visa || '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Stay Schedule Duration</strong>
                    <span className="font-black text-slate-900 font-mono">{bio.start} to {bio.end} ({bio.allowed || '-'} Days)</span>
                  </div>
                  <div className="col-span-2">
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Registered Destination Address</strong>
                    <span className="font-bold text-slate-900 text-xs">{bio.loc || '-'}</span>
                  </div>
                  <div className="col-span-2">
                    <strong className="text-slate-400 block text-[8px] uppercase font-black">Address Description Details</strong>
                    <span className="font-semibold text-slate-700 italic block">{bio.stayDescription || 'No description logged.'}</span>
                  </div>

                  <div className="col-span-2 bg-amber-50/50 p-3 rounded-xl border border-amber-250">
                    <strong className="text-amber-800 block text-[8px] uppercase font-black tracking-wide mb-0.5">Dossier Remark Details</strong>
                    <p className="font-black text-amber-950 font-sans text-[11px] leading-tight">
                      {currentRemarks[pp] || bio.remarks || 'No special remarks registered.'}
                    </p>
                  </div>
                </div>

                <div className="space-y-4">
                  <h4 className="text-[8px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Flight Movement logs</h4>
                  <table className="w-full text-left text-[8.5px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-slate-150 uppercase text-[7.5px] font-black">
                        <th className="p-1 border border-gray-400 text-center">#</th>
                        <th className="p-1 border border-gray-400">Timestamp</th>
                        <th className="p-1 border border-gray-400 text-center">Dir</th>
                        <th className="p-1 border border-gray-400">Visa & Stay Duration</th>
                        <th className="p-1 border border-gray-400">Stay Address / Location</th>
                        <th className="p-1 border border-gray-400">Registered Officer</th>
                      </tr>
                    </thead>
                    <tbody>
                      {personLogs.map((log, idx) => (
                        <tr key={log.id} className="border-b border-gray-300">
                          <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                          <td className="p-1 border border-gray-400 font-mono text-[8px]">{log.timestamp}</td>
                          <td className="p-1 border border-gray-400 text-center">
                            <span className={`px-1 py-0.5 rounded text-[7.5px] font-black ${log.mode === 'IN' ? 'bg-emerald-50 text-emerald-800' : 'bg-orange-50 text-orange-850'}`}>
                              {log.mode}
                            </span>
                          </td>
                          <td className="p-1 border border-gray-400 leading-tight">
                            <div className="font-bold">{log.visaType}</div>
                            {log.stayFrom && <div className="text-[7.5px] text-gray-500 font-mono mt-0.5">{log.stayFrom} to {log.stayTo} ({log.totalDays} Days)</div>}
                          </td>
                          <td className="p-1 border border-gray-400 leading-tight">
                            <div className="font-bold text-gray-800">{log.address || '-'}</div>
                          </td>
                          <td className="p-1 border border-gray-400 text-gray-500 font-bold leading-none text-[8px]">
                            <div>{log.officialTitle || ''}</div>
                            <div className="text-gray-800 mt-0.5">{log.officialName || 'System'}</div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {personCheckingHistory.length > 0 && (
                  <div className="space-y-4">
                    <h4 className="text-[8px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Physical Verification annals</h4>
                    <table className="w-full text-left text-[8.5px] border border-gray-400 border-collapse table-auto">
                      <thead>
                        <tr className="bg-slate-150 uppercase text-[7.5px] font-black">
                          <th className="p-1 border border-gray-400 text-center">#</th>
                          <th className="p-1 border border-gray-400">Check Date</th>
                          <th className="p-1 border border-gray-400">Verified Status</th>
                          <th className="p-1 border border-gray-400">Residence Confirmed Address</th>
                          <th className="p-1 border border-gray-400">Stay Description Detail</th>
                          <th className="p-1 border border-gray-400">Verifier Stamp</th>
                        </tr>
                      </thead>
                      <tbody>
                        {personCheckingHistory.map((ch, idx) => (
                          <tr key={idx} className="border-b border-gray-300">
                            <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                            <td className="p-1 border border-gray-400 font-mono text-[8px]">{ch.checkDate}</td>
                            <td className="p-1 border border-gray-400">
                              <span className={`px-1 py-0.5 rounded text-[7.5px] font-black ${
                                normalizePermitStatus(ch.status) === 'နေထိုင်ခွင့်ကျထားသောသူ' 
                                  ? 'bg-emerald-50 text-emerald-800' 
                                  : 'bg-rose-50 text-rose-800'
                              }`}>{normalizePermitStatus(ch.status) || 'နေထိုင်ခွင့်ကျထားသောသူ'}</span>
                            </td>
                            <td className="p-1 border border-gray-400 font-bold text-gray-800">{ch.confirmedAddress || '-'}</td>
                            <td className="p-1 border border-gray-400 text-gray-500 font-semibold italic">{ch.confirmedStayDescription || '-'}</td>
                            <td className="p-1 border border-gray-400 font-mono text-[7.5px] text-gray-500 leading-none">
                              <div>{ch.officerTitle || 'Officer'}</div>
                              <div className="text-gray-800 font-bold mt-0.5">{ch.officerName || '-'}</div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        <div className="grid grid-cols-3 gap-8 px-4 mt-20 no-print-inside">
          <div className="text-center space-y-12">
            <p className="text-[10px] font-black uppercase text-slate-500">Prepared By</p>
            <div className="border-t border-slate-400 pt-2 text-[10px] font-bold text-slate-800">
              Immigration Investigator
            </div>
            <p className="text-[8px] text-slate-400">Signature: _______________________</p>
          </div>
          
          <div className="text-center space-y-12">
            <p className="text-[10px] font-black uppercase text-slate-500">Verified By</p>
            <div className="border-t border-slate-400 pt-2 text-[10px] font-bold text-slate-800">
              Duty Commander
            </div>
            <p className="text-[8px] text-slate-400">Signature: _______________________</p>
          </div>

          <div className="text-center bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 space-y-10">
            <p className="text-[9px] font-black uppercase text-indigo-900 tracking-wider">
              Confirmed & Certified Box / Stamp
            </p>
            <div className="space-y-4 text-left text-[9px] font-bold text-slate-600">
              <div>Office Stamp: __________________</div>
              <div>Supervisor Signature: ________________</div>
              <div>Certified Date: {new Date().toLocaleDateString()}</div>
            </div>
          </div>
        </div>

        <div className="mt-16 pt-6 border-t border-dashed border-gray-300 text-center text-[9px] text-gray-400 leading-normal">
          <p>Myeik Immigration Enforcement Unit • Intelligence Group Investigation dossier</p>
          <p>CONFIDENTIAL • DEPT OF REGISTRATION RECORD • STRICTLY SECRET</p>
        </div>
      </div>

      {/* 7. Dedicated off-screen container for high-fidelity A4 PDF generation */}
      <div 
        id="pdf-capture-container" 
        className="bg-white text-black"
        style={{
          position: 'absolute',
          left: '-9999px',
          top: '-9999px',
          width: '800px', // standard 96 dpi width for A4
          boxSizing: 'border-box'
        }}
      >
        {/* PAGE 1: Overview & Matrix (only if multiple selected) */}
        {selectedPassports.length > 1 && (
          <div className="pdf-page bg-white p-8 border border-slate-200" style={{ width: '800px', minHeight: '1130px', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
            <div>
              {/* English-Only Header */}
              <div className="text-center space-y-1.5 mb-6 border-b-2 border-black pb-4">
                <h1 className="text-md font-black uppercase text-slate-900 tracking-tight leading-none">
                  Second Division, Department of Immigration (Myeik)
                </h1>
                <h2 className="text-xs font-black text-indigo-900 uppercase tracking-widest mt-1">
                  Group Stay Verification & Intelligence Report
                </h2>
                <p className="text-[10px] font-bold text-indigo-950 font-mono mt-1">
                  DOSSIER IDENTIFICATION NO: {activeDossierID}
                </p>
                <p className="text-[7px] font-mono font-bold text-slate-400 uppercase tracking-wide pt-1">
                  OFFICIAL INTEL COPY • PRINTED DATE: {new Date().toLocaleString()}
                </p>
              </div>

              {/* Group Summary Matrix */}
              <div className="mb-6">
                <h3 className="text-[10px] font-black uppercase tracking-wider mb-2 border-b pb-1">Group Comparison Overview Matrix</h3>
                <table className="w-full text-left text-[9px] border border-gray-400 border-collapse">
                  <thead>
                    <tr className="bg-slate-100 uppercase text-[8px] font-black">
                      <th className="p-1 border border-gray-400 text-center">#</th>
                      <th className="p-1 border border-gray-400">Guest Name / Passport</th>
                      <th className="p-1 border border-gray-400">Nationality</th>
                      <th className="p-1 border border-gray-400">Residence Stay Address</th>
                      <th className="p-1 border border-gray-400">Visa & Validity</th>
                      <th className="p-1 border border-gray-400">Local Sponsor</th>
                      <th className="p-1 border border-gray-400 text-center">Rem. Days</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedPassports.map((pp, idx) => {
                      const m = computedMovementMap[pp] || {};
                      const rem = getRemainingDaysNum(m.end);
                      return (
                        <tr key={pp} className="align-top border-b border-gray-300">
                          <td className="p-1 border border-gray-400 text-center font-bold">{idx + 1}</td>
                          <td className="p-1 border border-gray-400">
                            <div className="font-extrabold font-mono uppercase text-indigo-900 leading-none">{pp}</div>
                            <div className="font-bold text-gray-800 uppercase text-[8px] mt-0.5">{m.n || '-'}</div>
                          </td>
                          <td className="p-1 border border-gray-400 font-bold uppercase text-gray-700">{m.nat || '-'}</td>
                          <td className="p-1 border border-gray-400 leading-tight max-w-[140px] truncate">{m.loc || '-'}</td>
                          <td className="p-1 border border-gray-400 leading-none">
                            <div className="font-bold">{m.visa || '-'}</div>
                            <div className="text-[7px] font-mono text-gray-500 mt-0.5">{m.start} to {m.end}</div>
                          </td>
                          <td className="p-1 border border-gray-400 uppercase font-black text-gray-850 leading-tight">
                            <div>{m.agent || '-'}</div>
                          </td>
                          <td className="p-1 border border-gray-400 text-center font-bold font-mono">
                            <span className={rem < 0 ? 'text-rose-600' : 'text-slate-800'}>{rem} Days</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div>
              {/* Signatures block at bottom of summary page */}
              <div className="grid grid-cols-3 gap-6 px-2 mt-8">
                <div className="text-center space-y-10">
                  <p className="text-[9px] font-black uppercase text-slate-500">Prepared By</p>
                  <div className="border-t border-slate-400 pt-1 text-[9px] font-bold text-slate-800">
                    Investigation Analyst Officer
                  </div>
                  <p className="text-[7px] text-slate-400">Signature: _________________</p>
                </div>
                
                <div className="text-center space-y-10">
                  <p className="text-[9px] font-black uppercase text-slate-500">Verified By</p>
                  <div className="border-t border-slate-400 pt-1 text-[9px] font-bold text-slate-800">
                    Duty Supervisor / Commander
                  </div>
                  <p className="text-[7px] text-slate-400">Signature: _________________</p>
                </div>

                <div className="text-center bg-slate-50 border border-dashed border-slate-300 rounded-lg p-2.5 space-y-6">
                  <p className="text-[8px] font-black uppercase text-indigo-900 tracking-wider">
                    Confirmed & Certified Stamp / Box
                  </p>
                  <div className="space-y-2 text-left text-[8px] font-bold text-slate-600">
                    <div>Office Stamp: _______________</div>
                    <div>Signature: ____________</div>
                  </div>
                </div>
              </div>

              {/* Page Footer */}
              <div className="mt-8 pt-4 border-t border-dashed border-gray-300 text-center text-[8px] text-gray-400">
                Myeik Immigration Enforcement Unit • Intelligence Group Investigation dossier • Page 1
              </div>
            </div>
          </div>
        )}

        {/* PAGES 2+: Detailed Profiles */}
        {selectedPassports.map((pp, idx) => {
          const bio = computedMovementMap[pp];
          if (!bio) return null;

          const personLogs = records
            .filter(r => r.passport.trim().toUpperCase() === pp.toUpperCase())
            .sort((a, b) => parseTimestamp(b.timestamp) - parseTimestamp(a.timestamp));

          const personCheckingHistory = checkingHistory
            .filter(h => h.passport.trim().toUpperCase() === pp.toUpperCase())
            .sort((a, b) => b.id - a.id);

          const pageNum = selectedPassports.length > 1 ? idx + 2 : idx + 1;

          return (
            <div key={pp} className="pdf-page bg-white p-8 border border-slate-200" style={{ width: '800px', minHeight: '1130px', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
              <div>
                {/* Header for individual page */}
                <div className="text-center space-y-1 mb-4 border-b pb-2">
                  <h1 className="text-xs font-black uppercase text-slate-900 tracking-tight leading-none">
                    Department of Immigration (Myeik)
                  </h1>
                  <h2 className="text-[9px] font-black text-indigo-900 uppercase tracking-widest">
                    Individual Investigation Profile Dossier
                  </h2>
                  <p className="text-[9.5px] font-bold text-indigo-950 font-mono mt-1">
                    DOSSIER IDENTIFICATION NO: {activeDossierID}
                  </p>
                </div>

                {/* Profile Header Tag */}
                <div className="bg-slate-100 p-1.5 border-b-2 border-slate-900 flex justify-between items-center mb-4">
                  <strong className="text-[10px] uppercase tracking-tight text-indigo-950 font-black">
                    Dossier Record #{idx + 1}: {bio.n} ({pp})
                  </strong>
                  <span className="text-[7px] font-black text-slate-500 uppercase tracking-widest font-mono">CONFIDENTIAL RECORD</span>
                </div>

                {/* Demographics details */}
                <div className="grid grid-cols-2 gap-y-2 gap-x-4 text-[9.5px] leading-snug mb-4">
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Passport Number</strong>
                    <span className="font-black font-mono uppercase text-indigo-950">{pp}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Full Name</strong>
                    <span className="font-black uppercase text-slate-900">{bio.n}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Nationality / Gender</strong>
                    <span className="font-extrabold uppercase text-slate-800">{bio.nat || '-'} ({bio.gender === 'M' ? 'MALE' : bio.gender === 'F' ? 'FEMALE' : '-'})</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Visa Category & Period</strong>
                    <span className="font-black text-slate-900 font-mono">{bio.visa || '-'} ({bio.start} to {bio.end})</span>
                  </div>
                  <div className="col-span-2">
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Registered Destination Address</strong>
                    <span className="font-bold text-slate-900">{bio.loc || '-'}</span>
                  </div>
                  <div className="col-span-2">
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Address Description Details</strong>
                    <span className="font-semibold text-slate-700 italic block">{bio.stayDescription || 'No description logged.'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Local Broker / Sponsor</strong>
                    <span className="font-black uppercase text-indigo-900">{bio.agent || '-'}</span>
                  </div>
                  <div>
                    <strong className="text-slate-400 block text-[7.5px] uppercase font-black">Sponsor Contact</strong>
                    <span className="font-bold text-slate-700 font-mono">{bio.contact || '-'}</span>
                  </div>

                  {/* Remarks */}
                  <div className="col-span-2 bg-amber-50/50 p-2 rounded border border-amber-200 mt-1">
                    <strong className="text-amber-800 block text-[7px] uppercase font-black mb-0.5">Dossier Remark Details</strong>
                    <p className="font-black text-amber-950 font-sans text-[9.5px] leading-tight">
                      {currentRemarks[pp] || bio.remarks || 'No special remarks registered.'}
                    </p>
                  </div>
                </div>

                {/* Flight Movement logs */}
                <div className="space-y-1.5 mb-4">
                  <h4 className="text-[7.5px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Flight Movement logs</h4>
                  <table className="w-full text-left text-[8px] border border-gray-400 border-collapse table-auto">
                    <thead>
                      <tr className="bg-slate-100 uppercase text-[7px] font-black">
                        <th className="p-0.5 border border-gray-400 text-center w-6">#</th>
                        <th className="p-0.5 border border-gray-400">Timestamp</th>
                        <th className="p-0.5 border border-gray-400 text-center w-10">Dir</th>
                        <th className="p-0.5 border border-gray-400">Visa & Stay Schedule</th>
                        <th className="p-0.5 border border-gray-400">Stay Address</th>
                      </tr>
                    </thead>
                    <tbody>
                      {personLogs.slice(0, 4).map((log, lIdx) => (
                        <tr key={log.id} className="border-b border-gray-300">
                          <td className="p-0.5 border border-gray-400 text-center font-bold">{lIdx + 1}</td>
                          <td className="p-0.5 border border-gray-400 font-mono text-[7px]">{log.timestamp}</td>
                          <td className="p-0.5 border border-gray-400 text-center font-black text-[7px]">
                            <span className={log.mode === 'IN' ? 'text-emerald-700' : 'text-orange-700'}>{log.mode}</span>
                          </td>
                          <td className="p-0.5 border border-gray-400 leading-none">
                            <div className="font-bold">{log.visaType}</div>
                            {log.stayFrom && <div className="text-[6.5px] text-gray-500 font-mono mt-0.5">{log.stayFrom} to {log.stayTo}</div>}
                          </td>
                          <td className="p-0.5 border border-gray-400 font-bold truncate max-w-[180px]">{log.address || '-'}</td>
                        </tr>
                      ))}
                      {personLogs.length === 0 && (
                        <tr>
                          <td colSpan={5} className="p-1 border border-gray-400 text-center text-slate-400 italic">No logs found</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Physical check logs */}
                {personCheckingHistory.length > 0 && (
                  <div className="space-y-1.5 mb-4">
                    <h4 className="text-[7.5px] font-black uppercase text-slate-500 tracking-wider border-b pb-0.5">Physical Verification annals</h4>
                    <table className="w-full text-left text-[8px] border border-gray-400 border-collapse table-auto">
                      <thead>
                        <tr className="bg-slate-100 uppercase text-[7px] font-black">
                          <th className="p-0.5 border border-gray-400 text-center w-6">#</th>
                          <th className="p-0.5 border border-gray-400">Check Date</th>
                          <th className="p-0.5 border border-gray-400">Status</th>
                          <th className="p-0.5 border border-gray-400">Residence Address</th>
                          <th className="p-0.5 border border-gray-400">Verifier</th>
                        </tr>
                      </thead>
                      <tbody>
                        {personCheckingHistory.slice(0, 3).map((ch, cIdx) => (
                          <tr key={cIdx} className="border-b border-gray-300">
                            <td className="p-0.5 border border-gray-400 text-center font-bold">{cIdx + 1}</td>
                            <td className="p-0.5 border border-gray-400 font-mono text-[7px]">{ch.checkDate}</td>
                            <td className="p-0.5 border border-gray-400 font-black text-[7px]">
                              <span className={ch.status?.includes('NOT') ? 'text-rose-700' : 'text-emerald-700'}>{ch.status}</span>
                            </td>
                            <td className="p-0.5 border border-gray-400 font-bold truncate max-w-[180px]">{ch.confirmedAddress || '-'}</td>
                            <td className="p-0.5 border border-gray-400 font-bold text-gray-600 text-[7px]">{ch.officerName || 'Officer'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              <div>
                {/* Signature box for individual profile dossier */}
                <div className="grid grid-cols-2 gap-4 px-2 mt-4 border-t pt-3">
                  <div className="text-center space-y-6">
                    <p className="text-[8px] font-black uppercase text-slate-500">Prepared & Authenticated By</p>
                    <div className="text-[8px] font-bold text-slate-800 border-t pt-1 border-slate-300">
                      Investigation Analyst Officer
                    </div>
                  </div>
                  
                  <div className="text-center space-y-6">
                    <p className="text-[8px] font-black uppercase text-slate-500">Reviewed & Stamp Certified By</p>
                    <div className="text-[8px] font-bold text-slate-800 border-t pt-1 border-slate-300">
                      Duty Supervisor / Commander
                    </div>
                  </div>
                </div>

                {/* Page Footer */}
                <div className="mt-6 pt-3 border-t border-dashed border-gray-300 text-center text-[8px] text-gray-400">
                  Myeik Immigration Enforcement Unit • Intelligence Group Investigation dossier • Page {pageNum} of {selectedPassports.length > 1 ? selectedPassports.length + 1 : 1}
                </div>
              </div>
            </div>
          );
        })}
      </div>
        </>
      )}

    </div>
  );
};
