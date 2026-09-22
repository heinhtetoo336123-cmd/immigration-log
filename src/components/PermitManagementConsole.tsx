import React, { useState, useMemo } from 'react';
import { 
  CheckSquare, 
  Search, 
  PlusCircle, 
  FileSpreadsheet, 
  Printer, 
  CheckCircle2, 
  Clock, 
  Building2, 
  Hotel, 
  Edit3, 
  UserCheck, 
  AlertCircle, 
  X, 
  Save, 
  RefreshCw, 
  ShieldCheck,
  Calendar,
  Sparkles,
  ArrowRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { ImmRecord, MovementData, MasterItem, PermitStatus, normalizePermitStatus } from '../types';
import { isCompanyAddress } from '../utils/addressUtils';
import { MasterLocationInput } from './MasterLocationInput';

interface PermitManagementConsoleProps {
  records?: ImmRecord[];
  tempRecords?: ImmRecord[];
  setRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  setTempRecords?: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  showToast: (msg: string) => void;
  movementMap?: Record<string, MovementData>;
  currentUser?: { name?: string; title?: string } | null;
  cloudAuthUser?: any;
  isViewer?: boolean;
  passportToLatestInfo?: Record<string, { fullname: string; nationality: string; gender?: string; [key: string]: any }>;
  masterData?: MasterItem[];
  syncMaster?: (value: string, type: MasterItem['type'], linkedValue?: string) => void;
  setActivePrintPreview?: (preview: any) => void;
  logActivity?: (act: any) => void;
}

interface StayerPermitItem {
  passport: string;
  fullname: string;
  nationality: string;
  gender: 'M' | 'F' | '';
  visaType: string;
  arrivalDate: string;
  currentAddress: string;
  currentStayDesc?: string;
  isCompany: boolean;
  status: PermitStatus;
  permittedBy: string;
  permittedAddress?: string;
  permittedStayDesc?: string;
  permitDate?: string;
  stayTo?: string;
  totalDays?: string;
  remarks?: string;
  latestTimestamp: string;
  rawRecord?: ImmRecord;
}

export const PermitManagementConsole: React.FC<PermitManagementConsoleProps> = ({
  records = [],
  tempRecords = [],
  setRecords,
  setTempRecords,
  showToast,
  movementMap = {},
  currentUser,
  isViewer,
  passportToLatestInfo = {},
  masterData = [],
  syncMaster,
  setActivePrintPreview,
  logActivity
}) => {
  // Search & Filter States
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'PERMITTED' | 'NOT_PERMITTED'>('ALL');
  const [categoryFilter, setCategoryFilter] = useState<'ALL' | 'COMPANY' | 'OTHER'>('ALL');

  // Modal States
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [isRegisterNewModalOpen, setIsRegisterNewModalOpen] = useState(false);
  const [activeEditingItem, setActiveEditingItem] = useState<StayerPermitItem | null>(null);

  // Edit/Register Form Fields
  const [formPassport, setFormPassport] = useState('');
  const [formFullname, setFormFullname] = useState('');
  const [formNationality, setFormNationality] = useState('');
  const [formGender, setFormGender] = useState<'M' | 'F' | ''>('M');
  const [formStatus, setFormStatus] = useState<PermitStatus>('နေထိုင်ခွင့်ကျထားသောသူ');
  const [formPermittedBy, setFormPermittedBy] = useState('');
  const [formPermittedAddress, setFormPermittedAddress] = useState('');
  const [formPermittedStayDesc, setFormPermittedStayDesc] = useState('');
  const [formPermitDate, setFormPermitDate] = useState('');
  const [formRemarks, setFormRemarks] = useState('');
  const [formOfficerName, setFormOfficerName] = useState(currentUser?.name || '');
  const [formOfficerTitle, setFormOfficerTitle] = useState(currentUser?.title || 'လဝကမှူး');

  // Search input inside Register New Modal
  const [registerSearchQuery, setRegisterSearchQuery] = useState('');

  // 1. Compile list of all foreigners currently active in Myeik (STRICT STILL IN ONLY - left/departed persons excluded)
  const stayerList = useMemo(() => {
    const list: StayerPermitItem[] = [];
    const seenPassports = new Set<string>();
    const allRecords = [...(records || []), ...(tempRecords || [])];

    // Prioritize active movement map (still in)
    Object.entries(movementMap).forEach(([pass, movVal]) => {
      const mov = movVal as MovementData;
      const pUpper = pass.trim().toUpperCase();
      if (!pUpper || seenPassports.has(pUpper)) return;

      const inTime = Number(mov.inTime || 0);
      const outTime = Number(mov.outTime || 0);
      const hasDeparted = Boolean(mov.out || mov.isStillIn === false || (outTime > 0 && inTime > 0 && outTime >= inTime));
      if (hasDeparted || !mov.in) return; // STRICTLY STILL IN ONLY: Exclude left/departed

      // Find latest record for this passport
      const matchRecords = allRecords
        .filter(r => r.passport.trim().toUpperCase() === pUpper)
        .sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || '') || (b.id || 0) - (a.id || 0));
      const latestRec = matchRecords[0];

      // If their latest physical record is OUT, they have departed!
      if (latestRec && latestRec.mode === 'OUT') return;

      seenPassports.add(pUpper);

      const passInfo = passportToLatestInfo[pUpper] || {};
      const fullname = latestRec?.fullname || mov.n || passInfo.fullname || '-';
      const nationality = latestRec?.nationality || mov.nat || passInfo.nationality || '-';
      const gender = (latestRec?.gender || mov.gender || passInfo.gender || '') as 'M' | 'F' | '';
      const currentAddress = latestRec?.address || mov.loc || '-';
      const isCo = isCompanyAddress(currentAddress, masterData);

      const rawStatus = latestRec?.stillPermittedStatus || mov.stillPermittedStatus;
      const status: PermitStatus = normalizePermitStatus(rawStatus) || 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ';
      const permittedBy = latestRec?.permittedBy || mov.permittedBy || '';

      list.push({
        passport: pUpper,
        fullname,
        nationality,
        gender,
        visaType: latestRec?.visaType || mov.visa || '-',
        arrivalDate: mov.in || latestRec?.stayFrom || latestRec?.timestamp?.split(' ')[0] || '-',
        currentAddress,
        currentStayDesc: latestRec?.stayDescription || mov.stayDescription || '',
        isCompany: isCo,
        status,
        permittedBy,
        permittedAddress: currentAddress,
        permittedStayDesc: latestRec?.stayDescription || '',
        stayTo: latestRec?.stayTo || mov.end || '',
        totalDays: latestRec?.totalDays || mov.allowed || '',
        remarks: latestRec?.remarks || mov.remarks || '',
        latestTimestamp: latestRec?.timestamp || mov.in || '',
        rawRecord: latestRec
      });
    });

    // Fallback: Check records that might not be in movementMap
    const unmappedPassports = new Map<string, ImmRecord[]>();
    allRecords.forEach(r => {
      if (!r.passport) return;
      const pUpper = r.passport.trim().toUpperCase();
      if (seenPassports.has(pUpper)) return;
      if (!unmappedPassports.has(pUpper)) {
        unmappedPassports.set(pUpper, []);
      }
      unmappedPassports.get(pUpper)!.push(r);
    });

    unmappedPassports.forEach((recs, pUpper) => {
      recs.sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || '') || (b.id || 0) - (a.id || 0));
      const latestRec = recs[0];
      if (!latestRec || latestRec.mode !== 'IN') return; // STRICT STILL IN ONLY

      const mov = movementMap[pUpper];
      if (mov && (mov.out || mov.isStillIn === false)) return; // Exclude left

      seenPassports.add(pUpper);
      const isCo = isCompanyAddress(latestRec.address, masterData);
      const status: PermitStatus = normalizePermitStatus(latestRec.stillPermittedStatus) || 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ';

      list.push({
        passport: pUpper,
        fullname: latestRec.fullname || '-',
        nationality: latestRec.nationality || '-',
        gender: latestRec.gender || '',
        visaType: latestRec.visaType || '-',
        arrivalDate: latestRec.stayFrom || latestRec.timestamp?.split(' ')[0] || '-',
        currentAddress: latestRec.address || '-',
        currentStayDesc: latestRec.stayDescription || '',
        isCompany: isCo,
        status,
        permittedBy: latestRec.permittedBy || '',
        permittedAddress: latestRec.address,
        permittedStayDesc: latestRec.stayDescription || '',
        stayTo: latestRec.stayTo || '',
        totalDays: latestRec.totalDays || '',
        remarks: latestRec.remarks || '',
        latestTimestamp: latestRec.timestamp || '',
        rawRecord: latestRec
      });
    });

    return list.sort((a, b) => {
      // Sort: Permitted first, then by passport
      if (a.status === 'နေထိုင်ခွင့်ကျထားသောသူ' && b.status !== 'နေထိုင်ခွင့်ကျထားသောသူ') return -1;
      if (a.status !== 'နေထိုင်ခွင့်ကျထားသောသူ' && b.status === 'နေထိုင်ခွင့်ကျထားသောသူ') return 1;
      return a.passport.localeCompare(b.passport);
    });
  }, [movementMap, records, tempRecords, passportToLatestInfo, masterData]);

  // 2. Filtered list for display
  const filteredList = useMemo(() => {
    return stayerList.filter(item => {
      // Status Filter
      if (statusFilter === 'PERMITTED' && item.status !== 'နေထိုင်ခွင့်ကျထားသောသူ') return false;
      if (statusFilter === 'NOT_PERMITTED' && item.status !== 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ') return false;

      // Category Filter
      if (categoryFilter === 'COMPANY' && !item.isCompany) return false;
      if (categoryFilter === 'OTHER' && item.isCompany) return false;

      // Search Filter
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchesP = item.passport.toLowerCase().includes(q);
        const matchesN = item.fullname.toLowerCase().includes(q);
        const matchesNat = item.nationality.toLowerCase().includes(q);
        const matchesAddr = item.currentAddress.toLowerCase().includes(q);
        const matchesDesc = (item.currentStayDesc || '').toLowerCase().includes(q);
        const matchesPermitBy = item.permittedBy.toLowerCase().includes(q);
        const matchesRemarks = (item.remarks || '').toLowerCase().includes(q);
        if (!matchesP && !matchesN && !matchesNat && !matchesAddr && !matchesDesc && !matchesPermitBy && !matchesRemarks) {
          return false;
        }
      }
      return true;
    });
  }, [stayerList, statusFilter, categoryFilter, searchQuery]);

  // Statistics Counts
  const stats = useMemo(() => {
    let permitted = 0;
    let notPermitted = 0;
    let companyCount = 0;
    let otherCount = 0;

    stayerList.forEach(item => {
      if (item.status === 'နေထိုင်ခွင့်ကျထားသောသူ') {
        permitted += 1;
      } else {
        notPermitted += 1;
      }

      if (item.isCompany) {
        companyCount += 1;
      } else {
        otherCount += 1;
      }
    });

    return {
      total: stayerList.length,
      permitted,
      notPermitted,
      companyCount,
      otherCount
    };
  }, [stayerList]);

  // 3. Quick 1-Click Permit Toggle
  const handleQuickTogglePermit = (item: StayerPermitItem) => {
    if (isViewer) {
      showToast('⚠️ View-Only mode: Cannot modify records');
      return;
    }

    const newStatus: PermitStatus = item.status === 'နေထိုင်ခွင့်ကျထားသောသူ' 
      ? 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' 
      : 'နေထိုင်ခွင့်ကျထားသောသူ';

    const defaultPermitDesc = newStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'
      ? (item.permittedBy.trim() || `နေထိုင်ခွင့်ပြုထားသည် (${new Date().toISOString().split('T')[0]})`)
      : '';

    const pUpper = item.passport.toUpperCase();

    // Update records state
    setRecords(prev => prev.map(r => {
      if (r.passport.trim().toUpperCase() === pUpper) {
        return {
          ...r,
          stillPermittedStatus: newStatus,
          permittedBy: defaultPermitDesc,
          officialName: currentUser?.name || r.officialName,
          officialTitle: currentUser?.title || r.officialTitle,
          syncStatus: 'pending_sync',
          updatedAt: new Date().toISOString()
        };
      }
      return r;
    }));

    // Update tempRecords state
    setTempRecords(prev => prev.map(r => {
      if (r.passport.trim().toUpperCase() === pUpper) {
        return {
          ...r,
          stillPermittedStatus: newStatus,
          permittedBy: defaultPermitDesc,
          officialName: currentUser?.name || r.officialName,
          officialTitle: currentUser?.title || r.officialTitle,
          syncStatus: 'pending_sync',
          updatedAt: new Date().toISOString()
        };
      }
      return r;
    }));

    if (logActivity) {
      logActivity({
        action: 'UPDATE',
        module: 'CHECKING',
        details: `Toggled Stay Permit Status for ${item.passport} (${item.fullname}) to "${newStatus}"`,
        targetId: item.passport,
        officerName: currentUser ? `${currentUser.title} ${currentUser.name}` : 'Officer'
      });
    }

    showToast(`✓ ${item.passport} (${item.fullname}) အား "${newStatus}" သို့ အောင်မြင်စွာ ပြောင်းလဲပြီးပါပြီ`);
  };

  // 4. Open Edit Modal
  const handleOpenEditModal = (item: StayerPermitItem) => {
    setActiveEditingItem(item);
    setFormPassport(item.passport);
    setFormFullname(item.fullname);
    setFormNationality(item.nationality);
    setFormGender(item.gender);
    setFormStatus(item.status);
    setFormPermittedBy(item.permittedBy || '');
    setFormPermittedAddress(item.permittedAddress || item.currentAddress);
    setFormPermittedStayDesc(item.permittedStayDesc || item.currentStayDesc || '');
    setFormPermitDate(new Date().toISOString().split('T')[0]);
    setFormRemarks(item.remarks || '');
    setFormOfficerName(currentUser?.name || '');
    setFormOfficerTitle(currentUser?.title || 'လဝကမှူး');
    setIsEditModalOpen(true);
  };

  // 5. Open Register New Permit Modal
  const handleOpenRegisterNewModal = () => {
    setRegisterSearchQuery('');
    setFormPassport('');
    setFormFullname('');
    setFormNationality('');
    setFormGender('M');
    setFormStatus('နေထိုင်ခွင့်ကျထားသောသူ');
    setFormPermittedBy('');
    setFormPermittedAddress('');
    setFormPermittedStayDesc('');
    setFormPermitDate(new Date().toISOString().split('T')[0]);
    setFormRemarks('');
    setFormOfficerName(currentUser?.name || '');
    setFormOfficerTitle(currentUser?.title || 'လဝကမှူး');
    setIsRegisterNewModalOpen(true);
  };

  // Select a foreigner inside Register New Modal
  const handleSelectForeignerForNewPermit = (item: StayerPermitItem) => {
    setFormPassport(item.passport);
    setFormFullname(item.fullname);
    setFormNationality(item.nationality);
    setFormGender(item.gender || 'M');
    setFormStatus('နေထိုင်ခွင့်ကျထားသောသူ');
    setFormPermittedBy(item.permittedBy || `နေထိုင်ခွင့် အမိန့်စာအမှတ် (လဝက/မြိတ်/${new Date().getFullYear()})`);
    setFormPermittedAddress(item.currentAddress);
    setFormPermittedStayDesc(item.currentStayDesc || '');
    setFormRemarks(item.remarks || '');
  };

  // 6. Save Permit Changes (from Edit or Register Modal)
  const handleSavePermitDetails = (e: React.FormEvent) => {
    e.preventDefault();
    if (isViewer) {
      showToast('⚠️ View-Only mode: Cannot modify records');
      return;
    }

    const pUpper = formPassport.trim().toUpperCase();
    if (!pUpper) {
      showToast('⚠️ ကျေးဇူးပြု၍ Passport နံပါတ် ထည့်သွင်းပါ');
      return;
    }

    // Auto sync Master Data for permit description if entered
    if (formPermittedBy.trim() && syncMaster) {
      syncMaster(formPermittedBy.trim(), 'PermitDescription');
    }
    if (formPermittedAddress.trim() && syncMaster) {
      syncMaster(formPermittedAddress.trim(), 'Stay', formPermittedStayDesc.trim() || undefined);
    }

    // Update records state
    let matched = false;
    setRecords(prev => {
      const updated = prev.map(r => {
        if (r.passport.trim().toUpperCase() === pUpper) {
          matched = true;
          return {
            ...r,
            fullname: formFullname.trim() || r.fullname,
            nationality: formNationality.trim() || r.nationality,
            gender: formGender || r.gender,
            stillPermittedStatus: formStatus,
            permittedBy: formPermittedBy.trim(),
            address: formPermittedAddress.trim() || r.address,
            stayDescription: formPermittedStayDesc.trim() || r.stayDescription,
            remarks: formRemarks.trim() || r.remarks,
            officialName: formOfficerName.trim() || r.officialName,
            officialTitle: formOfficerTitle.trim() || r.officialTitle,
            syncStatus: 'pending_sync' as const,
            updatedAt: new Date().toISOString()
          };
        }
        return r;
      });
      return updated;
    });

    // Update tempRecords state
    setTempRecords(prev => {
      return prev.map(r => {
        if (r.passport.trim().toUpperCase() === pUpper) {
          return {
            ...r,
            fullname: formFullname.trim() || r.fullname,
            nationality: formNationality.trim() || r.nationality,
            gender: formGender || r.gender,
            stillPermittedStatus: formStatus,
            permittedBy: formPermittedBy.trim(),
            address: formPermittedAddress.trim() || r.address,
            stayDescription: formPermittedStayDesc.trim() || r.stayDescription,
            remarks: formRemarks.trim() || r.remarks,
            officialName: formOfficerName.trim() || r.officialName,
            officialTitle: formOfficerTitle.trim() || r.officialTitle,
            syncStatus: 'pending_sync' as const,
            updatedAt: new Date().toISOString()
          };
        }
        return r;
      });
    });

    if (logActivity) {
      logActivity({
        action: 'UPDATE',
        module: 'CHECKING',
        details: `Saved Stay Permit details for ${pUpper} (${formFullname}): Status: "${formStatus}", Order/Ref: "${formPermittedBy}"`,
        targetId: pUpper,
        officerName: `${formOfficerTitle} ${formOfficerName}`
      });
    }

    showToast(`✓ ${pUpper} (${formFullname}) ၏ နေထိုင်ခွင့် အချက်အလက်များကို သိမ်းဆည်းပြီးပါပြီ`);
    setIsEditModalOpen(false);
    setIsRegisterNewModalOpen(false);
  };

  // 7. Export to Excel
  const handleExportExcel = () => {
    if (filteredList.length === 0) {
      showToast('⚠️ ထုတ်ယူရန် စာရင်းဒေတာ မရှိပါ');
      return;
    }

    const excelData = filteredList.map((item, index) => ({
      'စဉ် (No)': index + 1,
      'နိုင်ငံကူးလက်မှတ် (Passport)': item.passport,
      'အမည် (Full Name)': item.fullname,
      'ကျား/မ (Gender)': item.gender || '-',
      'နိုင်ငံသား (Nationality)': item.nationality,
      'ဗီဇာအမျိုးအစား (Visa)': item.visaType,
      'ဆိုက်ရောက်ရက် (Arrival)': item.arrivalDate,
      'နေထိုင်ရာလိပ်စာ (Address)': item.currentAddress,
      'လိပ်စာအသေးစိတ် (Stay Detail)': item.currentStayDesc || '-',
      'အမျိုးအစား (Classification)': item.isCompany ? 'ကုမ္ပဏီ (Co., Ltd)' : 'အခြား/ဟိုတယ် (Other/Hotel)',
      'နေထိုင်ခွင့် အခြေအနေ (Permit Status)': item.status,
      'နေထိုင်ခွင့် အမိန့်စာအမှတ် / ခွင့်ပြုသူ (Permit Order / Authority)': item.permittedBy || '-',
      'သက်တမ်းကုန်ဆုံးရက် (Stay To)': item.stayTo || '-',
      'ခွင့်ပြုရက် (Total Days)': item.totalDays || '-',
      'မှတ်ချက် (Remarks)': item.remarks || '-'
    }));

    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Stay_Permit_List');
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Stay_Permit_List_Myeik_${dateStr}.xlsx`);
    showToast('📊 နေထိုင်ခွင့် စာရင်းအား Excel ဖိုင်အဖြစ် ဒေါင်းလုဒ်လုပ်ပြီးပါပြီ');
  };

  // 8. Print Official Permit Registry
  const handlePrintPermitList = () => {
    if (!setActivePrintPreview) {
      window.print();
      return;
    }

    setActivePrintPreview({
      title: 'မြိတ်ခရိုင်အတွင်း နိုင်ငံခြားသားများ၏ နေထိုင်ခွင့် (Stay Permit) စာရင်း',
      type: 'CUSTOM_REPORT',
      columns: [
        { id: 'sr', label: 'စဉ်', width: '40px' },
        { id: 'passport', label: 'နိုင်ငံကူးလက်မှတ်', width: '90px' },
        { id: 'fullname', label: 'အမည်', width: '130px' },
        { id: 'nationality', label: 'နိုင်ငံသား', width: '80px' },
        { id: 'address', label: 'နေထိုင်ရာလိပ်စာ', width: '140px' },
        { id: 'permitStatus', label: 'နေထိုင်ခွင့် အခြေအနေ', width: '120px' },
        { id: 'permittedBy', label: 'အမိန့်စာအမှတ် / ခွင့်ပြုသူ', width: '140px' },
        { id: 'remarks', label: 'မှတ်ချက်', width: '100px' }
      ],
      rows: filteredList.map((item, idx) => ({
        sr: idx + 1,
        passport: item.passport,
        fullname: item.fullname,
        nationality: item.nationality,
        address: item.currentAddress + (item.currentStayDesc ? ` (${item.currentStayDesc})` : ''),
        permitStatus: item.status,
        permittedBy: item.permittedBy || '-',
        remarks: item.remarks || '-'
      })),
      date: new Date().toLocaleDateString('my-MM'),
      officer: currentUser ? `${currentUser.title} ${currentUser.name}` : ''
    });
  };

  return (
    <div className="max-w-[1700px] mx-auto p-3 sm:p-6 space-y-6">
      {/* 1. Header Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white p-6 sm:p-8 rounded-3xl shadow-xl border border-slate-800 flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
        <div className="flex items-center gap-4">
          <div className="p-3.5 bg-indigo-600/30 text-indigo-400 rounded-2xl border border-indigo-500/30 shrink-0">
            <CheckSquare size={32} />
          </div>
          <div>
            <div className="flex items-center gap-2.5 flex-wrap">
              <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight">
                နေထိုင်ခွင့်ကျထားသူများ စာရင်းနှင့် စီမံခန့်ခွဲမှု
              </h2>
              <span className="text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 px-2.5 py-0.5 rounded-full font-black uppercase tracking-wider">
                STAY PERMIT CONSOLE
              </span>
            </div>
            <p className="text-xs text-slate-300 font-medium mt-1">
              မြိတ်ခရိုင်အတွင်း နိုင်ငံခြားသားများ၏ နေထိုင်ခွင့် (Stay Permit) ကျထားသူများနှင့် မလျှောက်ထားသေးသူများ စာရင်းသွင်းခြင်း၊ ပြင်ဆင်ခြင်းနှင့် ထိန်းချုပ်မှု
            </p>
          </div>
        </div>

        {/* Primary Header Action Buttons */}
        <div className="flex items-center gap-2.5 flex-wrap w-full md:w-auto">
          <button
            onClick={handleOpenRegisterNewModal}
            className="flex-1 md:flex-initial bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-black px-4 py-2.5 rounded-xl flex items-center justify-center gap-2 shadow-lg transition-all cursor-pointer border border-indigo-400/40"
          >
            <PlusCircle size={16} />
            <span>နေထိုင်ခွင့် စာရင်းသွင်းရန်</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer border border-emerald-500/40"
            title="Excel ဖိုင်ထုတ်မည်"
          >
            <FileSpreadsheet size={15} />
            <span className="hidden sm:inline">Excel ထုတ်မည်</span>
          </button>

          <button
            onClick={handlePrintPermitList}
            className="bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white text-xs font-black px-3.5 py-2.5 rounded-xl flex items-center justify-center gap-1.5 shadow-md transition-all cursor-pointer border border-slate-700"
            title="Print ထုတ်မည်"
          >
            <Printer size={15} />
            <span className="hidden sm:inline">Print</span>
          </button>
        </div>
      </div>

      {/* 2. Top Metric Statistics Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3.5">
        {/* Total Active */}
        <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
          <div>
            <div className="text-[10px] font-bold uppercase text-slate-400">စုစုပေါင်း နေထိုင်သူ</div>
            <div className="text-xl sm:text-2xl font-black text-slate-900 mt-0.5">{stats.total}</div>
          </div>
          <div className="w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center text-slate-700">
            <UserCheck size={20} />
          </div>
        </div>

        {/* Permit Granted */}
        <div 
          onClick={() => setStatusFilter(statusFilter === 'PERMITTED' ? 'ALL' : 'PERMITTED')}
          className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between ${
            statusFilter === 'PERMITTED' 
              ? 'bg-emerald-50 border-emerald-400 ring-2 ring-emerald-500/20' 
              : 'bg-white border-slate-200 hover:border-emerald-300'
          }`}
        >
          <div>
            <div className="text-[10px] font-bold uppercase text-emerald-700">နေထိုင်ခွင့်ကျထားသူ</div>
            <div className="text-xl sm:text-2xl font-black text-emerald-700 mt-0.5">{stats.permitted}</div>
          </div>
          <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
            <CheckCircle2 size={20} />
          </div>
        </div>

        {/* Not Permitted Yet */}
        <div 
          onClick={() => setStatusFilter(statusFilter === 'NOT_PERMITTED' ? 'ALL' : 'NOT_PERMITTED')}
          className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between ${
            statusFilter === 'NOT_PERMITTED' 
              ? 'bg-amber-50 border-amber-400 ring-2 ring-amber-500/20' 
              : 'bg-white border-slate-200 hover:border-amber-300'
          }`}
        >
          <div>
            <div className="text-[10px] font-bold uppercase text-amber-700">မလျှောက်ထားသေးသူ</div>
            <div className="text-xl sm:text-2xl font-black text-amber-700 mt-0.5">{stats.notPermitted}</div>
          </div>
          <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center">
            <Clock size={20} />
          </div>
        </div>

        {/* Corporate Count */}
        <div 
          onClick={() => setCategoryFilter(categoryFilter === 'COMPANY' ? 'ALL' : 'COMPANY')}
          className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between ${
            categoryFilter === 'COMPANY' 
              ? 'bg-purple-50 border-purple-400 ring-2 ring-purple-500/20' 
              : 'bg-white border-slate-200 hover:border-purple-300'
          }`}
        >
          <div>
            <div className="text-[10px] font-bold uppercase text-purple-700">ကုမ္ပဏီလိပ်စာ (Co.,Ltd)</div>
            <div className="text-xl sm:text-2xl font-black text-purple-700 mt-0.5">{stats.companyCount}</div>
          </div>
          <div className="w-10 h-10 rounded-xl bg-purple-100 text-purple-700 flex items-center justify-center">
            <Building2 size={20} />
          </div>
        </div>

        {/* Other / Hotel Count */}
        <div 
          onClick={() => setCategoryFilter(categoryFilter === 'OTHER' ? 'ALL' : 'OTHER')}
          className={`p-4 rounded-2xl border transition-all cursor-pointer shadow-sm flex items-center justify-between col-span-2 sm:col-span-1 ${
            categoryFilter === 'OTHER' 
              ? 'bg-blue-50 border-blue-400 ring-2 ring-blue-500/20' 
              : 'bg-white border-slate-200 hover:border-blue-300'
          }`}
        >
          <div>
            <div className="text-[10px] font-bold uppercase text-blue-700">အခြား / ဟိုတယ် နေထိုင်သူ</div>
            <div className="text-xl sm:text-2xl font-black text-blue-700 mt-0.5">{stats.otherCount}</div>
          </div>
          <div className="w-10 h-10 rounded-xl bg-blue-100 text-blue-700 flex items-center justify-center">
            <Hotel size={20} />
          </div>
        </div>
      </div>

      {/* 3. Search & Quick Filters Toolbar */}
      <div className="bg-white p-4 sm:p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3.5">
        <div className="flex flex-col lg:flex-row gap-3 items-stretch lg:items-center justify-between">
          {/* Search Input */}
          <div className="relative flex-1">
            <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={17} />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="နိုင်ငံကူးလက်မှတ်၊ အမည်၊ နိုင်ငံသား၊ လိပ်စာ၊ အမိန့်စာအမှတ် ဖြင့် ရှာဖွေပါ..."
              className="w-full pl-10 pr-10 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all"
            />
            {searchQuery && (
              <button 
                onClick={() => setSearchQuery('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
              >
                <X size={14} />
              </button>
            )}
          </div>

          {/* Status Filter Tabs */}
          <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl shrink-0 overflow-x-auto">
            <button
              onClick={() => setStatusFilter('ALL')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer whitespace-nowrap ${
                statusFilter === 'ALL'
                  ? 'bg-white text-indigo-950 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              အားလုံး ({stayerList.length})
            </button>
            <button
              onClick={() => setStatusFilter('PERMITTED')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 whitespace-nowrap ${
                statusFilter === 'PERMITTED'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-emerald-700 hover:bg-emerald-50'
              }`}
            >
              <CheckCircle2 size={13} />
              <span>နေထိုင်ခွင့်ကျထားသောသူ ({stats.permitted})</span>
            </button>
            <button
              onClick={() => setStatusFilter('NOT_PERMITTED')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 whitespace-nowrap ${
                statusFilter === 'NOT_PERMITTED'
                  ? 'bg-amber-600 text-white shadow-xs'
                  : 'text-amber-700 hover:bg-amber-50'
              }`}
            >
              <Clock size={13} />
              <span>မလျှောက်ထားသေးသူ ({stats.notPermitted})</span>
            </button>
          </div>
        </div>

        {/* Secondary Category Filters */}
        <div className="flex items-center gap-2 pt-1 border-t border-slate-100 text-xs font-bold text-slate-600">
          <span className="text-[11px] font-black uppercase text-slate-400">လိပ်စာ အမျိုးအစား:</span>
          <button
            onClick={() => setCategoryFilter('ALL')}
            className={`px-2.5 py-1 rounded-lg text-[11px] font-black transition-all cursor-pointer ${
              categoryFilter === 'ALL' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
            }`}
          >
            အားလုံး
          </button>
          <button
            onClick={() => setCategoryFilter('COMPANY')}
            className={`px-2.5 py-1 rounded-lg text-[11px] font-black transition-all cursor-pointer flex items-center gap-1 ${
              categoryFilter === 'COMPANY' ? 'bg-purple-700 text-white' : 'bg-purple-50 text-purple-700 hover:bg-purple-100'
            }`}
          >
            <Building2 size={12} />
            <span>🏢 ကုမ္ပဏီများ ({stats.companyCount})</span>
          </button>
          <button
            onClick={() => setCategoryFilter('OTHER')}
            className={`px-2.5 py-1 rounded-lg text-[11px] font-black transition-all cursor-pointer flex items-center gap-1 ${
              categoryFilter === 'OTHER' ? 'bg-blue-700 text-white' : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
            }`}
          >
            <Hotel size={12} />
            <span>🏨 အခြား / ဟိုတယ်များ ({stats.otherCount})</span>
          </button>
        </div>
      </div>

      {/* 4. Main Data Table & List */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        {/* Table Header / Summary */}
        <div className="p-4 bg-slate-50/80 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <span className="text-xs font-black uppercase text-slate-700 tracking-wider">
              ရှာဖွေတွေ့ရှိသော စာရင်း: <span className="text-indigo-600 font-black">{filteredList.length}</span> ဦး
            </span>
          </div>
          {(statusFilter !== 'ALL' || categoryFilter !== 'ALL' || searchQuery) && (
            <button
              onClick={() => {
                setStatusFilter('ALL');
                setCategoryFilter('ALL');
                setSearchQuery('');
              }}
              className="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 cursor-pointer"
            >
              <RefreshCw size={12} /> Reset Filter
            </button>
          )}
        </div>

        {/* Desktop Table View */}
        <div className="hidden md:block overflow-x-auto max-h-[700px] overflow-y-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-slate-100 sticky top-0 border-b border-slate-200 z-10">
              <tr className="uppercase text-[10px] text-slate-600 font-black tracking-wider">
                <th className="p-3.5 text-center w-12 border-r border-slate-200">စဉ်</th>
                <th className="p-3.5 border-r border-slate-200">နိုင်ငံကူးလက်မှတ်</th>
                <th className="p-3.5 border-r border-slate-200">အမည်နှင့် နိုင်ငံသား</th>
                <th className="p-3.5 border-r border-slate-200">လက်ရှိနေထိုင်ရာလိပ်စာ</th>
                <th className="p-3.5 border-r border-slate-200 text-center">နေထိုင်ခွင့် အခြေအနေ</th>
                <th className="p-3.5 border-r border-slate-200">နေထိုင်ခွင့် အမိန့်စာအမှတ် / အကြောင်းအရာ</th>
                <th className="p-3.5 text-right w-44">လုပ်ဆောင်ချက်</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 bg-white">
              {filteredList.length === 0 ? (
                <tr>
                  <td colSpan={7} className="p-16 text-center text-slate-400 font-bold uppercase tracking-wider text-xs">
                    ⚠️ ရှာဖွေမှုနှင့် ကိုက်ညီသော နေထိုင်ခွင့် စာရင်းမှတ်တမ်း မရှိပါ
                  </td>
                </tr>
              ) : (
                filteredList.map((item, idx) => {
                  const isPermitted = item.status === 'နေထိုင်ခွင့်ကျထားသောသူ';

                  return (
                    <tr 
                      key={item.passport + '_' + idx}
                      className={`hover:bg-slate-50/80 transition-colors ${
                        isPermitted ? 'bg-emerald-50/20' : ''
                      }`}
                    >
                      {/* 1. Sr */}
                      <td className="p-3.5 text-center font-bold text-slate-400 border-r border-slate-100">
                        {idx + 1}
                      </td>

                      {/* 2. Passport */}
                      <td className="p-3.5 border-r border-slate-100">
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-black text-xs text-indigo-950 bg-indigo-50 border border-indigo-200/80 px-2 py-0.5 rounded-md">
                            {item.passport}
                          </span>
                          {item.gender && (
                            <span className={`text-[10px] font-black px-1.5 py-0.2 rounded ${
                              item.gender === 'M' ? 'bg-blue-100 text-blue-800' : 'bg-pink-100 text-pink-800'
                            }`}>
                              {item.gender}
                            </span>
                          )}
                        </div>
                        <div className="text-[10px] text-slate-400 mt-1 font-semibold">
                          ဆိုက်ရောက်: {item.arrivalDate}
                        </div>
                      </td>

                      {/* 3. Name & Nationality */}
                      <td className="p-3.5 border-r border-slate-100">
                        <div className="font-black text-slate-900 text-xs">
                          {item.fullname}
                        </div>
                        <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-bold mt-0.5">
                          <span>{item.nationality}</span>
                          <span className="text-slate-300">•</span>
                          <span className="text-indigo-600 font-semibold">{item.visaType}</span>
                        </div>
                      </td>

                      {/* 4. Address */}
                      <td className="p-3.5 border-r border-slate-100 max-w-xs">
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <span className="font-bold text-slate-900 text-xs truncate">
                            {item.currentAddress}
                          </span>
                          {item.isCompany ? (
                            <span className="text-[9px] font-black px-1.5 py-0.2 rounded bg-purple-100 text-purple-800 shrink-0">
                              🏢 CO., LTD
                            </span>
                          ) : (
                            <span className="text-[9px] font-black px-1.5 py-0.2 rounded bg-blue-100 text-blue-800 shrink-0">
                              🏨 HOTEL/RESIDENCE
                            </span>
                          )}
                          <button
                            onClick={() => handleOpenEditModal(item)}
                            className="text-[10px] text-indigo-600 hover:text-indigo-800 font-bold underline cursor-pointer ml-auto flex items-center gap-0.5 bg-indigo-50 hover:bg-indigo-100 px-1.5 py-0.5 rounded border border-indigo-200"
                            title="လိပ်စာ ပြင်ဆင်/ပြောင်းလဲမည်"
                          >
                            <span>📍 ပြင်မည်</span>
                          </button>
                        </div>
                        {item.currentStayDesc && (
                          <div className="text-[11px] text-slate-500 font-medium truncate mt-0.5">
                            {item.currentStayDesc}
                          </div>
                        )}
                      </td>

                      {/* 5. Status Badge */}
                      <td className="p-3.5 text-center border-r border-slate-100">
                        {isPermitted ? (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                            <CheckCircle2 size={13} className="text-emerald-600" />
                            <span>နေထိုင်ခွင့်ကျထားသောသူ</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-black bg-amber-100 text-amber-800 border border-amber-300">
                            <Clock size={13} className="text-amber-600" />
                            <span>မလျှောက်ထားသေးသူ</span>
                          </span>
                        )}
                      </td>

                      {/* 6. Permitted By / Description */}
                      <td className="p-3.5 border-r border-slate-100 max-w-xs">
                        {isPermitted ? (
                          <div>
                            <div className="font-bold text-slate-800 text-xs truncate">
                              {item.permittedBy || 'နေထိုင်ခွင့်ပြုထားသည်'}
                            </div>
                            {item.remarks && (
                              <div className="text-[10px] text-slate-400 font-medium truncate mt-0.5">
                                မှတ်ချက်: {item.remarks}
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-400 text-[11px] italic font-medium">
                            - မရှိသေးပါ -
                          </span>
                        )}
                      </td>

                      {/* 7. Action Controls */}
                      <td className="p-3.5 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          {/* Quick 1-Click Toggle Button */}
                          <button
                            onClick={() => handleQuickTogglePermit(item)}
                            className={`px-2.5 py-1.5 rounded-xl text-[11px] font-black flex items-center gap-1 transition-all cursor-pointer shadow-xs ${
                              isPermitted
                                ? 'bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300'
                                : 'bg-emerald-600 hover:bg-emerald-500 text-white border border-emerald-500'
                            }`}
                            title={isPermitted ? 'မလျှောက်ထားသေးသူ အဖြစ် ပြောင်းလဲမည်' : 'နေထိုင်ခွင့်ကျထားသူ အဖြစ် သတ်မှတ်မည်'}
                          >
                            {isPermitted ? (
                              <>
                                <RefreshCw size={12} />
                                <span>ပြန်ဖျက်</span>
                              </>
                            ) : (
                              <>
                                <CheckCircle2 size={12} />
                                <span>ခွင့်ပြုမည်</span>
                              </>
                            )}
                          </button>

                          {/* Edit Details Button */}
                          <button
                            onClick={() => handleOpenEditModal(item)}
                            className="p-1.5 rounded-xl bg-slate-100 hover:bg-indigo-50 hover:text-indigo-700 text-slate-700 border border-slate-200 transition-all cursor-pointer"
                            title="အသေးစိတ် ပြင်ဆင်ရန်"
                          >
                            <Edit3 size={14} />
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

        {/* Mobile Card Layout (< md screens) */}
        <div className="block md:hidden divide-y divide-slate-100 p-2 space-y-3">
          {filteredList.length === 0 ? (
            <div className="p-12 text-center text-slate-400 font-bold uppercase tracking-wider text-xs">
              ⚠️ ရှာဖွေမှုနှင့် ကိုက်ညီသော စာရင်း မရှိပါ
            </div>
          ) : (
            filteredList.map((item, idx) => {
              const isPermitted = item.status === 'နေထိုင်ခွင့်ကျထားသောသူ';

              return (
                <div 
                  key={'m_card_' + item.passport + '_' + idx}
                  className={`p-4 rounded-2xl border space-y-3 ${
                    isPermitted ? 'bg-emerald-50/30 border-emerald-200' : 'bg-white border-slate-200'
                  }`}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div>
                      <div className="flex items-center gap-1.5">
                        <span className="font-mono font-black text-xs text-indigo-950 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">
                          {item.passport}
                        </span>
                        {item.gender && (
                          <span className={`text-[10px] font-black px-1.5 py-0.2 rounded ${
                            item.gender === 'M' ? 'bg-blue-100 text-blue-800' : 'bg-pink-100 text-pink-800'
                          }`}>
                            {item.gender}
                          </span>
                        )}
                      </div>
                      <div className="font-black text-slate-900 text-sm mt-1">
                        {item.fullname}
                      </div>
                      <div className="text-xs text-slate-500 font-semibold">
                        {item.nationality} • {item.visaType}
                      </div>
                    </div>

                    <div>
                      {isPermitted ? (
                        <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                          <CheckCircle2 size={12} />
                          <span>ကျထားသူ</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-black bg-amber-100 text-amber-800 border border-amber-300">
                          <Clock size={12} />
                          <span>မလျှောက်သေး</span>
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200/80 text-xs space-y-1">
                    <div className="flex items-center gap-1.5">
                      <span className="text-slate-400 font-semibold">လိပ်စာ:</span>
                      <span className="font-bold text-slate-800 truncate">{item.currentAddress}</span>
                      {item.isCompany && (
                        <span className="text-[9px] font-black px-1 py-0.2 rounded bg-purple-100 text-purple-800">
                          CO
                        </span>
                      )}
                    </div>
                    {isPermitted && item.permittedBy && (
                      <div className="flex items-center gap-1.5">
                        <span className="text-slate-400 font-semibold">အမိန့်စာ:</span>
                        <span className="font-bold text-emerald-800 truncate">{item.permittedBy}</span>
                      </div>
                    )}
                  </div>

                  <div className="flex items-center justify-end gap-2 pt-1">
                    <button
                      onClick={() => handleQuickTogglePermit(item)}
                      className={`flex-1 py-2 rounded-xl text-xs font-black flex items-center justify-center gap-1.5 transition-all cursor-pointer ${
                        isPermitted
                          ? 'bg-amber-100 hover:bg-amber-200 text-amber-900 border border-amber-300'
                          : 'bg-emerald-600 hover:bg-emerald-500 text-white'
                      }`}
                    >
                      {isPermitted ? <RefreshCw size={13} /> : <CheckCircle2 size={13} />}
                      <span>{isPermitted ? 'မလျှောက်ထားသေးသို့ ပြောင်းမည်' : 'နေထိုင်ခွင့်ပြုမည်'}</span>
                    </button>
                    <button
                      onClick={() => handleOpenEditModal(item)}
                      className="p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 cursor-pointer"
                    >
                      <Edit3 size={15} />
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>

      {/* 5. EDIT PERMIT MODAL */}
      <AnimatePresence>
        {isEditModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-xl overflow-hidden"
            >
              {/* Modal Header */}
              <div className="p-5 bg-gradient-to-r from-slate-900 to-indigo-950 text-white flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-indigo-600/40 rounded-xl">
                    <Edit3 size={18} />
                  </div>
                  <div>
                    <h3 className="text-base font-black uppercase tracking-tight">
                      နေထိုင်ခွင့် အချက်အလက် ပြင်ဆင်ရန်
                    </h3>
                    <p className="text-[11px] text-slate-300 font-medium">
                      {formPassport} ({formFullname})
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsEditModalOpen(false)}
                  className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-white/10 transition-all cursor-pointer"
                >
                  <X size={18} />
                </button>
              </div>

              {/* Modal Form */}
              <form onSubmit={handleSavePermitDetails} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
                {/* Foreigner Details Box */}
                <div className="bg-slate-50 p-3.5 rounded-2xl border border-slate-200/80 flex items-center justify-between text-xs font-bold">
                  <div>
                    <span className="text-slate-400 block text-[10px] uppercase">Passport & Name</span>
                    <span className="font-mono text-indigo-950 font-black">{formPassport}</span> • <span>{formFullname}</span>
                  </div>
                  <div className="text-right">
                    <span className="text-slate-400 block text-[10px] uppercase">Nationality</span>
                    <span className="text-slate-800 font-black">{formNationality}</span>
                  </div>
                </div>

                {/* Status Selector (2-option clean switch) */}
                <div>
                  <label className="block text-[11px] font-black uppercase text-slate-700 mb-2">
                    နေထိုင်ခွင့် အခြေအနေ (Permit Status)
                  </label>
                  <div className="grid grid-cols-2 gap-3">
                    <label className={`p-3 rounded-2xl border-2 cursor-pointer flex items-center justify-center gap-2 text-xs font-black transition-all ${
                      formStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'
                        ? 'border-emerald-500 bg-emerald-50 text-emerald-900 shadow-xs'
                        : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'
                    }`}>
                      <input
                        type="radio"
                        name="editPermitStatus"
                        value="နေထိုင်ခွင့်ကျထားသောသူ"
                        checked={formStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'}
                        onChange={() => setFormStatus('နေထိုင်ခွင့်ကျထားသောသူ')}
                        className="sr-only"
                      />
                      <CheckCircle2 size={16} className={formStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' ? 'text-emerald-600' : 'text-slate-400'} />
                      <span>နေထိုင်ခွင့်ကျထားသောသူ</span>
                    </label>

                    <label className={`p-3 rounded-2xl border-2 cursor-pointer flex items-center justify-center gap-2 text-xs font-black transition-all ${
                      formStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ'
                        ? 'border-amber-500 bg-amber-50 text-amber-900 shadow-xs'
                        : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'
                    }`}>
                      <input
                        type="radio"
                        name="editPermitStatus"
                        value="နေထိုင်ခွင့် မလျှောက်ထားသေးသူ"
                        checked={formStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ'}
                        onChange={() => setFormStatus('နေထိုင်ခွင့် မလျှောက်ထားသေးသူ')}
                        className="sr-only"
                      />
                      <Clock size={16} className={formStatus === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ? 'text-amber-600' : 'text-slate-400'} />
                      <span>မလျှောက်ထားသေးသူ</span>
                    </label>
                  </div>
                </div>

                {/* Permitted By / Reference Order */}
                {formStatus === 'နေထိုင်ခွင့်ကျထားသောသူ' && (
                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1.5">
                      နေထိုင်ခွင့် အမိန့်စာအမှတ် / ခွင့်ပြုချက် (Permit Order / Reference)
                    </label>
                    <input
                      type="text"
                      value={formPermittedBy}
                      onChange={(e) => setFormPermittedBy(e.target.value)}
                      placeholder="ဥပမာ- လဝက/မြိတ်/၂၀၂၆/၁၂၃ (နေထိုင်ခွင့်ပြုထားဆဲ)"
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white"
                      required={formStatus === 'နေထိုင်ခွင့်ကျထားသောသူ'}
                    />
                  </div>
                )}

                {/* Permitted Stay Address & Description with Master Data Autofill */}
                <MasterLocationInput
                  locationValue={formPermittedAddress}
                  onLocationChange={setFormPermittedAddress}
                  descriptionValue={formPermittedStayDesc}
                  onDescriptionChange={setFormPermittedStayDesc}
                  masterData={masterData}
                  locationLabel="နေထိုင်ခွင့်ပြုသည့် လိပ်စာ / ကုမ္ပဏီ (Permitted Stay Address)"
                  descriptionLabel="လိပ်စာအသေးစိတ် (Stay Description / Detail)"
                  locationPlaceholder="ဥပမာ- Tidy Co., Ltd (သို့) ဟိုတယ်အမည် ရွေးချယ်ပါ..."
                  idPrefix="permit_edit_loc"
                />

                {/* Remarks */}
                <div>
                  <label className="block text-[11px] font-black uppercase text-slate-700 mb-1.5">
                    မှတ်ချက် (Remarks)
                  </label>
                  <input
                    type="text"
                    value={formRemarks}
                    onChange={(e) => setFormRemarks(e.target.value)}
                    placeholder="အပိုဆောင်း မှတ်ချက်များ..."
                    className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white"
                  />
                </div>

                {/* Officer Signature Info */}
                <div className="grid grid-cols-2 gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200">
                  <div>
                    <label className="block text-[10px] font-black uppercase text-slate-500 mb-1">
                      စာရင်းသွင်းအရာရှိ (Officer)
                    </label>
                    <input
                      type="text"
                      value={formOfficerName}
                      onChange={(e) => setFormOfficerName(e.target.value)}
                      className="w-full p-2 bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-800"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black uppercase text-slate-500 mb-1">
                      ရာထူး (Title)
                    </label>
                    <input
                      type="text"
                      value={formOfficerTitle}
                      onChange={(e) => setFormOfficerTitle(e.target.value)}
                      className="w-full p-2 bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-800"
                      required
                    />
                  </div>
                </div>

                {/* Modal Footer Actions */}
                <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-200">
                  <button
                    type="button"
                    onClick={() => setIsEditModalOpen(false)}
                    className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 cursor-pointer"
                  >
                    ပယ်ဖျက်မည်
                  </button>
                  <button
                    type="submit"
                    className="px-6 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-black flex items-center gap-1.5 shadow-md transition-all cursor-pointer"
                  >
                    <Save size={14} />
                    <span>အပြောင်းအလဲ သိမ်းဆည်းမည်</span>
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 6. REGISTER NEW PERMIT MODAL */}
      <AnimatePresence>
        {isRegisterNewModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-2xl overflow-hidden"
            >
              {/* Modal Header */}
              <div className="p-5 bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-indigo-600/40 rounded-xl">
                    <PlusCircle size={20} />
                  </div>
                  <div>
                    <h3 className="text-base font-black uppercase tracking-tight">
                      နေထိုင်ခွင့်ကျထားသူ စာရင်းသွင်းရန်
                    </h3>
                    <p className="text-[11px] text-slate-300 font-medium">
                      မြိတ်ခရိုင်အတွင်း နေထိုင်ခွင့်ရရှိသူအား အမိန့်စာအမှတ်နှင့်အတူ စာရင်းသွင်းမည်
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsRegisterNewModalOpen(false)}
                  className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-white/10 transition-all cursor-pointer"
                >
                  <X size={18} />
                </button>
              </div>

              {/* Modal Content */}
              <form onSubmit={handleSavePermitDetails} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
                {/* Select from Active Foreigners Quick Search */}
                <div>
                  <label className="block text-[11px] font-black uppercase text-slate-700 mb-1.5">
                    🔎 လက်ရှိနေထိုင်သူများမှ ရွေးချယ်ပါ (Select Active Foreigner)
                  </label>
                  <div className="relative mb-2">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={15} />
                    <input
                      type="text"
                      value={registerSearchQuery}
                      onChange={(e) => setRegisterSearchQuery(e.target.value)}
                      placeholder="Passport သို့မဟုတ် အမည်ဖြင့် အမြန်ရှာပြီး ရွေးချယ်ပါ..."
                      className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                    />
                  </div>

                  {/* Quick Select Foreigner Scroll Container */}
                  <div className="max-h-36 overflow-y-auto border border-slate-200 rounded-xl divide-y divide-slate-100 bg-slate-50/50">
                    {stayerList
                      .filter(s => {
                        if (!registerSearchQuery.trim()) return true;
                        const q = registerSearchQuery.toLowerCase();
                        return (
                          s.passport.toLowerCase().includes(q) ||
                          s.fullname.toLowerCase().includes(q) ||
                          s.nationality.toLowerCase().includes(q) ||
                          s.currentAddress.toLowerCase().includes(q)
                        );
                      })
                      .slice(0, 8)
                      .map(s => (
                        <div
                          key={'sel_' + s.passport}
                          onClick={() => handleSelectForeignerForNewPermit(s)}
                          className={`p-2.5 flex items-center justify-between text-xs cursor-pointer transition-colors ${
                            formPassport === s.passport ? 'bg-indigo-50 text-indigo-900 font-black' : 'hover:bg-slate-100 text-slate-700'
                          }`}
                        >
                          <div className="flex items-center gap-2">
                            <span className="font-mono font-bold bg-white px-1.5 py-0.5 rounded border border-slate-200 text-indigo-900">
                              {s.passport}
                            </span>
                            <span className="font-black">{s.fullname}</span>
                            <span className="text-slate-400">({s.nationality})</span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] text-slate-500 font-semibold">{s.currentAddress}</span>
                            <span className={`text-[9px] font-black px-1.5 py-0.2 rounded ${
                              s.status === 'နေထိုင်ခွင့်ကျထားသောသူ' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                            }`}>
                              {s.status === 'နေထိုင်ခွင့်ကျထားသောသူ' ? 'ကျထားသူ' : 'မလျှောက်သေး'}
                            </span>
                          </div>
                        </div>
                      ))}
                  </div>
                </div>

                {/* Foreigner Inputs */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      နိုင်ငံကူးလက်မှတ် (Passport) *
                    </label>
                    <input
                      type="text"
                      value={formPassport}
                      onChange={(e) => setFormPassport(e.target.value.toUpperCase())}
                      placeholder="PASSPORT NUMBER"
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-black text-indigo-950 uppercase"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      အမည် (Full Name) *
                    </label>
                    <input
                      type="text"
                      value={formFullname}
                      onChange={(e) => setFormFullname(e.target.value)}
                      placeholder="FULL NAME"
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-black text-slate-900"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      နိုင်ငံသား (Nationality)
                    </label>
                    <input
                      type="text"
                      value={formNationality}
                      onChange={(e) => setFormNationality(e.target.value)}
                      placeholder="NATIONALITY"
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      ကျား/မ (Gender)
                    </label>
                    <select
                      value={formGender}
                      onChange={(e) => setFormGender(e.target.value as any)}
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                    >
                      <option value="M">ကျား (Male)</option>
                      <option value="F">မ (Female)</option>
                    </select>
                  </div>
                </div>

                {/* Permit Details Section */}
                <div className="p-4 bg-emerald-50/50 rounded-2xl border border-emerald-200 space-y-3">
                  <div className="text-xs font-black uppercase text-emerald-950 flex items-center gap-1.5">
                    <CheckCircle2 size={16} className="text-emerald-700" />
                    <span>နေထိုင်ခွင့် ခွင့်ပြုချက် အချက်အလက်များ</span>
                  </div>

                  <div>
                    <label className="block text-[11px] font-black uppercase text-emerald-900 mb-1">
                      နေထိုင်ခွင့် အမိန့်စာအမှတ် (Permit Reference / Order No.) *
                    </label>
                    <input
                      type="text"
                      value={formPermittedBy}
                      onChange={(e) => setFormPermittedBy(e.target.value)}
                      placeholder="ဥပမာ- လဝက/မြိတ်/၂၀၂၆/၁၂၃ (နေထိုင်ခွင့်ကျထားသည်)"
                      className="w-full p-2.5 bg-white border border-emerald-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                      required
                    />
                  </div>

                  <div className="pt-1">
                    <MasterLocationInput
                      locationValue={formPermittedAddress}
                      onLocationChange={setFormPermittedAddress}
                      descriptionValue={formPermittedStayDesc}
                      onDescriptionChange={setFormPermittedStayDesc}
                      masterData={masterData}
                      locationLabel="နေထိုင်ခွင့်ပြုသည့် လိပ်စာ/ကုမ္ပဏီ (Address/Co.,Ltd)"
                      descriptionLabel="လိပ်စာအသေးစိတ် (Stay Detail)"
                      locationPlaceholder="နေထိုင်ခွင့်ပြုသည့် လိပ်စာ ရွေးချယ်ပါ..."
                      idPrefix="permit_reg_loc"
                    />
                  </div>
                </div>

                {/* Officer & Remarks */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      စာရင်းသွင်းအရာရှိ (Officer)
                    </label>
                    <input
                      type="text"
                      value={formOfficerName}
                      onChange={(e) => setFormOfficerName(e.target.value)}
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                      ရာထူး (Title)
                    </label>
                    <input
                      type="text"
                      value={formOfficerTitle}
                      onChange={(e) => setFormOfficerTitle(e.target.value)}
                      className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                      required
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-black uppercase text-slate-700 mb-1">
                    မှတ်ချက် (Remarks)
                  </label>
                  <input
                    type="text"
                    value={formRemarks}
                    onChange={(e) => setFormRemarks(e.target.value)}
                    placeholder="အပိုဆောင်း မှတ်ချက်..."
                    className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800"
                  />
                </div>

                {/* Actions */}
                <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-200">
                  <button
                    type="button"
                    onClick={() => setIsRegisterNewModalOpen(false)}
                    className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 cursor-pointer"
                  >
                    ပယ်ဖျက်မည်
                  </button>
                  <button
                    type="submit"
                    className="px-6 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-black flex items-center gap-1.5 shadow-md transition-all cursor-pointer"
                  >
                    <Save size={14} />
                    <span>နေထိုင်ခွင့် အသစ် စာရင်းသွင်းမည်</span>
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};
